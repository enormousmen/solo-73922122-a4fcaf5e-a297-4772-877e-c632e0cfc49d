"""XLSX 沉降表解析与校验。

业务规则（见需求）：
* 工作簿至多 3 张工作表，每表至多 200 个数据行；
* 固定四列表头：桥梁 / 观测日 / 测点 / 沉降毫米；
* 桥梁、观测日允许**纵向合并**，合并区域覆盖的行沿用左上角值；
* 普通空白单元格不沿用：桥梁/观测日/测点空白按错误拒收，沉降空白跳过不入仓；
* 沉降为 0 是有效读数；
* 观测日必须落在上传指定月份内，否则拒收整份文件；
* 跨表（含同表重复行）同一(桥梁, 观测日, 测点)出现不同读数时，
  给出全部来源单元格并整份拒收；同值重复只保留一条记录但保留全部来源。
"""
from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from io import BytesIO
from typing import Any

from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

# 固定四列（1 基列号 -> 表头名）
EXPECTED_HEADERS: dict[int, str] = {
    1: "桥梁",
    2: "观测日",
    3: "测点",
    4: "沉降毫米",
}
MAX_SHEETS = 3
MAX_DATA_ROWS_PER_SHEET = 200
# 仅这两列允许纵向合并
MERGE_ALLOWED_COLUMNS = (1, 2)

# 字符串日期兜底解析格式（Excel 日期单元格通常直接是 datetime/date）
_DATE_FORMATS = (
    "%Y-%m-%d",
    "%Y/%m/%d",
    "%Y.%m.%d",
    "%Y 年 %m 月 %d 日",
    "%Y年%m月%d日",
)


class WorkbookError(ValueError):
    """工作簿结构或内容校验失败，对应 HTTP 400。

    details 用于承载跨表冲突明细（读数 + 来源单元格）。
    """

    def __init__(self, message: str, details: Any | None = None) -> None:
        super().__init__(message)
        self.message = message
        self.details = details


@dataclass(frozen=True)
class Reading:
    """一条规范化后的测点读数。"""

    bridge: str
    observed_on: date
    point: str
    settlement_mm: Decimal
    # 全部来源单元格，形如 "Sheet1!D7"；同值重复时可能有多个
    sources: tuple[str, ...]


@dataclass(frozen=True)
class ParsedWorkbook:
    month: str
    readings: tuple[Reading, ...]
    sheet_count: int
    input_row_count: int
    # 整行四列皆空、直接忽略的行数
    blank_rows: int
    # 行内有主信息但沉降毫米为空白、因此不入仓的行数
    blank_settlement_rows: int


@dataclass
class _Aggregate:
    bridge: str
    observed_on: date
    point: str
    settlement_mm: Decimal
    sources: list[str] = field(default_factory=list)


def parse_month(text: str) -> tuple[int, int]:
    """解析 'YYYY-MM'，非法格式抛 WorkbookError。"""
    parts = text.split("-")
    if len(parts) != 2 or len(parts[0]) != 4 or len(parts[1]) != 2 or not parts[0].isdigit() or not parts[1].isdigit():
        raise WorkbookError(f"月份格式应为 YYYY-MM，实际收到：{text!r}")
    year, month = int(parts[0]), int(parts[1])
    if not 1 <= month <= 12:
        raise WorkbookError(f"月份必须在 01-12 之间，实际收到：{text!r}")
    return year, month


def _build_merge_map(ws: Worksheet, sheet_name: str) -> dict[tuple[int, int], Any]:
    """构建 (row, col) -> 合并区左上角值 的映射，并约束合并形态。

    仅允许桥梁列、观测日列上的纯纵向合并（跨多行、不跨列）。
    """
    merge_map: dict[tuple[int, int], Any] = {}
    for rng in ws.merged_cells.ranges:
        min_col, min_row, max_col, max_row = rng.bounds
        if min_col != max_col:  # 横向 / 跨列合并一律不接受
            raise WorkbookError(
                f"工作表 {sheet_name!r} 存在跨列合并 {str(rng)}，"
                f"仅允许「桥梁」「观测日」两列纵向合并"
            )
        if min_col not in MERGE_ALLOWED_COLUMNS:
            raise WorkbookError(
                f"工作表 {sheet_name!r} 的 {get_column_letter(min_col)} 列存在合并单元格 {str(rng)}，"
                f"仅「桥梁」「观测日」两列允许纵向合并"
            )
        top_value = ws.cell(row=min_row, column=min_col).value
        for row in range(min_row, max_row + 1):
            merge_map[(row, min_col)] = top_value
    return merge_map


def _check_headers(ws: Worksheet, sheet_name: str) -> None:
    extra: list[str] = []
    for col in range(1, ws.max_column + 1):
        value = ws.cell(row=1, column=col).value
        if col in EXPECTED_HEADERS:
            header = str(value).strip() if value is not None else ""
            if header != EXPECTED_HEADERS[col]:
                raise WorkbookError(
                    f"工作表 {sheet_name!r} 第 {col} 列表头应为 "
                    f"{EXPECTED_HEADERS[col]!r}，实际为 {value!r}"
                )
        elif value is not None and str(value).strip():
            extra.append(f"{get_column_letter(col)}1={value!r}")
    if extra:
        raise WorkbookError(
            f"工作表 {sheet_name!r} 存在固定四列之外的表头：{', '.join(extra)}"
        )
    # 四列必须都在（max_column 不足时补齐检查）
    for col, expected in EXPECTED_HEADERS.items():
        if col > ws.max_column:
            raise WorkbookError(
                f"工作表 {sheet_name!r} 缺少第 {col} 列表头 {expected!r}"
            )


def _effective_value(ws: Worksheet, merge_map: dict[tuple[int, int], Any], row: int, col: int) -> Any:
    """取单元格有效值：在合并区域内则沿用合并区左上角值，否则取本格值。"""
    if (row, col) in merge_map:
        return merge_map[(row, col)]
    return ws.cell(row=row, column=col).value


def _is_blank(value: Any) -> bool:
    return value is None or (isinstance(value, str) and not value.strip())


def _clean_text(value: Any, sheet_name: str, cell: str, field_name: str) -> str:
    if _is_blank(value):
        raise WorkbookError(f"工作表 {sheet_name!r} 单元格 {cell} 的{field_name}为空白，普通空白不可沿用")
    if not isinstance(value, str):
        raise WorkbookError(f"工作表 {sheet_name!r} 单元格 {cell} 的{field_name}应为文本，实际为 {value!r}")
    text = value.strip()
    if not text:
        raise WorkbookError(f"工作表 {sheet_name!r} 单元格 {cell} 的{field_name}为空白，普通空白不可沿用")
    return text


def _parse_date(value: Any, year: int, month: int, sheet_name: str, cell: str) -> date:
    if isinstance(value, datetime):
        parsed = value.date()
    elif isinstance(value, date):
        parsed = value
    elif isinstance(value, str):
        text = value.strip()
        parsed = None
        for fmt in _DATE_FORMATS:
            try:
                parsed = datetime.strptime(text, fmt).date()
                break
            except ValueError:
                continue
        if parsed is None:
            raise WorkbookError(f"工作表 {sheet_name!r} 单元格 {cell} 的观测日无法解析：{value!r}")
    else:
        raise WorkbookError(f"工作表 {sheet_name!r} 单元格 {cell} 的观测日不是合法日期：{value!r}")

    if parsed.year != year or parsed.month != month:
        raise WorkbookError(
            f"工作表 {sheet_name!r} 单元格 {cell} 的观测日 {parsed.isoformat()} "
            f"不属于上传月份 {year:04d}-{month:02d}"
        )
    return parsed


def _parse_settlement(value: Any, sheet_name: str, cell: str) -> Decimal:
    if isinstance(value, bool):  # bool 是 int 的子类，先排除
        raise WorkbookError(f"工作表 {sheet_name!r} 单元格 {cell} 的沉降毫米不是数值：{value!r}")
    if isinstance(value, (int, float, Decimal)):
        try:
            return Decimal(str(value))
        except InvalidOperation:
            raise WorkbookError(f"工作表 {sheet_name!r} 单元格 {cell} 的沉降毫米不是数值：{value!r}") from None
    if isinstance(value, str):
        text = value.strip()
        try:
            return Decimal(text)
        except InvalidOperation:
            raise WorkbookError(f"工作表 {sheet_name!r} 单元格 {cell} 的沉降毫米不是数值：{value!r}") from None
    raise WorkbookError(f"工作表 {sheet_name!r} 单元格 {cell} 的沉降毫米不是数值：{value!r}")


def parse_workbook(file_bytes: bytes, month: str) -> ParsedWorkbook:
    """解析并校验整份上传文件，返回规范化读数；任何规则不满足均抛 WorkbookError。"""
    year, month_num = parse_month(month)

    try:
        from openpyxl import load_workbook

        wb = load_workbook(filename=BytesIO(file_bytes), data_only=True, read_only=False)
    except Exception as exc:  # openpyxl 对坏文件抛多种异常，统一成 400
        raise WorkbookError(f"无法按 XLSX 打开上传文件：{exc}") from None

    if len(wb.worksheets) > MAX_SHEETS:
        raise WorkbookError(f"工作簿至多包含 {MAX_SHEETS} 张工作表，实际为 {len(wb.worksheets)} 张")

    aggregates: dict[tuple[str, date, str], _Aggregate] = {}
    input_row_count = 0
    blank_rows = 0
    blank_settlement_rows = 0

    for ws in wb.worksheets:
        sheet_name = ws.title
        _check_headers(ws, sheet_name)
        merge_map = _build_merge_map(ws, sheet_name)

        data_rows = 0
        for row in range(2, ws.max_row + 1):
            raw_values = [ws.cell(row=row, column=col).value for col in (1, 2, 3, 4)]
            if all(_is_blank(v) for v in raw_values):
                # 完全空白行：不计入 200 行限额，直接忽略
                blank_rows += 1
                continue
            data_rows += 1
            if data_rows > MAX_DATA_ROWS_PER_SHEET:
                raise WorkbookError(
                    f"工作表 {sheet_name!r} 数据行超过 {MAX_DATA_ROWS_PER_SHEET} 行上限"
                )
            input_row_count += 1

            bridge_cell = f"{sheet_name}!A{row}"
            date_cell = f"{sheet_name}!B{row}"
            point_cell = f"{sheet_name}!C{row}"
            value_cell = f"{sheet_name}!D{row}"

            bridge = _clean_text(_effective_value(ws, merge_map, row, 1), sheet_name, bridge_cell, "桥梁")
            observed_on = _parse_date(
                _effective_value(ws, merge_map, row, 2), year, month_num, sheet_name, date_cell
            )
            point = _clean_text(_effective_value(ws, merge_map, row, 3), sheet_name, point_cell, "测点")

            raw_settlement = _effective_value(ws, merge_map, row, 4)
            if _is_blank(raw_settlement):
                # 数值空白：不入仓（零不是空白，是有效读数）
                blank_settlement_rows += 1
                continue
            settlement = _parse_settlement(raw_settlement, sheet_name, value_cell)

            key = (bridge, observed_on, point)
            existing = aggregates.get(key)
            if existing is None:
                aggregates[key] = _Aggregate(
                    bridge=bridge,
                    observed_on=observed_on,
                    point=point,
                    settlement_mm=settlement,
                    sources=[value_cell],
                )
            else:
                if existing.settlement_mm != settlement:
                    conflict_sources = [*existing.sources, value_cell]
                    raise WorkbookError(
                        "同一桥梁、观测日、测点存在不同读数，整份文件拒收",
                        details=[
                            {
                                "桥梁": bridge,
                                "观测日": observed_on.isoformat(),
                                "测点": point,
                                "冲突读数": [str(existing.settlement_mm), str(settlement)],
                                "来源单元格": conflict_sources,
                            }
                        ],
                    )
                # 同值重复：只存一条，但保留全部来源
                existing.sources.append(value_cell)

    if not aggregates:
        raise WorkbookError(f"上传文件中没有任何有效沉降读数（月份 {month}），拒绝发布空数据集")

    readings = tuple(
        Reading(
            bridge=agg.bridge,
            observed_on=agg.observed_on,
            point=agg.point,
            settlement_mm=agg.settlement_mm,
            sources=tuple(agg.sources),
        )
        for agg in sorted(aggregates.values(), key=lambda a: (a.bridge, a.observed_on, a.point))
    )
    return ParsedWorkbook(
        month=month,
        readings=readings,
        sheet_count=len(wb.worksheets),
        input_row_count=input_row_count,
        blank_rows=blank_rows,
        blank_settlement_rows=blank_settlement_rows,
    )
