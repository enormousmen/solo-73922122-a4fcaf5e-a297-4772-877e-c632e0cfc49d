"""excel.parse_workbook 的规则级单元测试。"""
from __future__ import annotations

from datetime import date
from decimal import Decimal

import pytest

from app.excel import WorkbookError, parse_month, parse_workbook

from .workbook_factory import build_workbook, standard_rows

MONTH = "2026-05"


def test_basic_parse_zero_and_blank():
    parsed = parse_workbook(build_workbook({"S1": standard_rows()}), MONTH)
    assert parsed.sheet_count == 1
    assert parsed.input_row_count == 4
    assert parsed.blank_settlement_rows == 1
    # P3 沉降空白不入仓，零读数保留，共 3 条
    assert len(parsed.readings) == 3
    values = {(r.bridge, r.point): r.settlement_mm for r in parsed.readings}
    assert values[("长江大桥", "P1")] == Decimal("1.2")
    assert values[("长江大桥", "P2")] == Decimal("0")
    assert values[("黄河大桥", "P1")] == Decimal("-2.5")
    p2 = next(r for r in parsed.readings if r.bridge == "长江大桥" and r.point == "P2")
    assert p2.sources == ("S1!D3",)


def test_vertical_merge_is_inherited():
    # A、B 列纵向合并，左上角写值，合并覆盖行留空
    rows = [
        ["长江大桥", date(2026, 5, 10), "P1", 1.0],
        [None, None, "P2", 2.0],
        [None, None, "P3", 3.0],
    ]
    parsed = parse_workbook(
        build_workbook({"S1": rows}, merges={"S1": ["A2:A4", "B2:B4"]}), MONTH
    )
    assert all(r.bridge == "长江大桥" for r in parsed.readings)
    assert all(r.observed_on == date(2026, 5, 10) for r in parsed.readings)


def test_plain_blank_bridge_is_rejected():
    rows = [
        ["长江大桥", date(2026, 5, 10), "P1", 1.0],
        [None, date(2026, 5, 10), "P2", 2.0],  # 普通空白，不在合并区
    ]
    with pytest.raises(WorkbookError, match="普通空白不可沿用"):
        parse_workbook(build_workbook({"S1": rows}), MONTH)


def test_horizontal_merge_rejected():
    rows = [["长江大桥", date(2026, 5, 10), "P1", 1.0]]
    with pytest.raises(WorkbookError, match="跨列合并"):
        parse_workbook(build_workbook({"S1": rows}, merges={"S1": ["A2:B2"]}), MONTH)


def test_merge_on_point_column_rejected():
    rows = [
        ["长江大桥", date(2026, 5, 10), "P1", 1.0],
        ["长江大桥", date(2026, 5, 10), None, 2.0],
    ]
    with pytest.raises(WorkbookError, match="仅「桥梁」「观测日」两列允许纵向合并"):
        parse_workbook(build_workbook({"S1": rows}, merges={"S1": ["C2:C3"]}), MONTH)


def test_date_outside_month_rejected():
    rows = [["长江大桥", date(2026, 4, 30), "P1", 1.0]]
    with pytest.raises(WorkbookError, match="不属于上传月份"):
        parse_workbook(build_workbook({"S1": rows}), MONTH)


def test_invalid_date_string_rejected():
    rows = [["长江大桥", "五月十日", "P1", 1.0]]
    with pytest.raises(WorkbookError, match="观测日无法解析"):
        parse_workbook(build_workbook({"S1": rows}), MONTH)


def test_non_numeric_settlement_rejected():
    rows = [["长江大桥", date(2026, 5, 10), "P1", "abc"]]
    with pytest.raises(WorkbookError, match="沉降毫米不是数值"):
        parse_workbook(build_workbook({"S1": rows}), MONTH)


def test_too_many_sheets_rejected():
    sheets = {f"S{i}": standard_rows() for i in range(4)}
    with pytest.raises(WorkbookError, match="至多包含 3 张工作表"):
        parse_workbook(build_workbook(sheets), MONTH)


def test_too_many_rows_rejected():
    rows = [["B", date(2026, 5, 1), f"P{i}", 1.0] for i in range(201)]
    with pytest.raises(WorkbookError, match="数据行超过 200 行上限"):
        parse_workbook(build_workbook({"S1": rows}), MONTH)


def test_blank_rows_do_not_count_toward_limit():
    rows = [["B", date(2026, 5, 1), f"P{i}", 1.0] for i in range(198)]
    # 两条整行空白插在数据中间（openpyxl 落盘会丢弃尾部全空行）
    rows += [[None, None, None, None], [None, None, None, None]]
    rows += [
        ["B", date(2026, 5, 1), "P198", 1.0],
        ["B", date(2026, 5, 1), "P199", 1.0],
    ]
    parsed = parse_workbook(build_workbook({"S1": rows}), MONTH)
    assert parsed.input_row_count == 200
    assert parsed.blank_rows == 2


def test_row_limit_includes_blanks_between_data():
    rows = [["B", date(2026, 5, 1), f"P{i}", 1.0] for i in range(200)]
    rows.append([None, None, None, None])
    rows.append(["B", date(2026, 5, 1), "P200", 1.0])  # 第 201 个数据行
    with pytest.raises(WorkbookError, match="数据行超过 200 行上限"):
        parse_workbook(build_workbook({"S1": rows}), MONTH)


def test_bad_header_rejected():
    from openpyxl import Workbook
    from io import BytesIO

    wb = Workbook()
    ws = wb.active
    ws.title = "S1"
    ws.append(["桥名", "观测日", "测点", "沉降毫米"])  # 首列表头不符合约定
    ws.append(["长江大桥", date(2026, 5, 1), "P1", 1.0])
    buf = BytesIO()
    wb.save(buf)
    with pytest.raises(WorkbookError, match="表头应为"):
        parse_workbook(buf.getvalue(), MONTH)


def test_extra_header_column_rejected():
    from openpyxl import Workbook
    from io import BytesIO

    wb = Workbook()
    ws = wb.active
    ws.title = "S1"
    ws.append(["桥梁", "观测日", "测点", "沉降毫米", "备注"])
    buf = BytesIO()
    wb.save(buf)
    with pytest.raises(WorkbookError, match="固定四列之外的表头"):
        parse_workbook(buf.getvalue(), MONTH)


def test_same_value_duplicate_kept_once_with_all_sources():
    rows = [
        ["长江大桥", date(2026, 5, 10), "P1", 1.5],
        ["长江大桥", date(2026, 5, 10), "P1", 1.5],  # 同值重复
    ]
    parsed = parse_workbook(build_workbook({"S1": rows}), MONTH)
    assert len(parsed.readings) == 1
    assert parsed.readings[0].sources == ("S1!D2", "S1!D3")


def test_cross_sheet_conflict_reports_cells_and_rejects():
    sheet1 = [["长江大桥", date(2026, 5, 10), "P1", 1.5]]
    sheet2 = [["长江大桥", date(2026, 5, 10), "P1", 1.6]]
    with pytest.raises(WorkbookError) as exc_info:
        parse_workbook(build_workbook({"S1": sheet1, "S2": sheet2}), MONTH)
    err = exc_info.value
    assert err.details is not None
    detail = err.details[0]
    assert detail["来源单元格"] == ["S1!D2", "S2!D2"]
    assert detail["冲突读数"] == ["1.5", "1.6"]


def test_cross_sheet_same_value_merged_with_sources():
    sheet1 = [["长江大桥", date(2026, 5, 10), "P1", 1.5]]
    sheet2 = [["长江大桥", date(2026, 5, 10), "P1", 1.5]]
    parsed = parse_workbook(build_workbook({"S1": sheet1, "S2": sheet2}), MONTH)
    assert len(parsed.readings) == 1
    assert parsed.readings[0].sources == ("S1!D2", "S2!D2")


def test_empty_workbook_rejected():
    with pytest.raises(WorkbookError, match="没有任何有效沉降读数"):
        parse_workbook(build_workbook({"S1": [[None, None, None, None]]}), MONTH)


def test_broken_file_rejected():
    with pytest.raises(WorkbookError, match="无法按 XLSX 打开"):
        parse_workbook(b"not an xlsx at all", MONTH)


def test_parse_month_validation():
    assert parse_month("2026-05") == (2026, 5)
    for bad in ("2026-13", "2026/05", "26-05", "2026-5", ""):
        with pytest.raises(WorkbookError):
            parse_month(bad)


def test_string_date_iso_format_accepted():
    rows = [["长江大桥", "2026-05-10", "P1", 1.0]]
    parsed = parse_workbook(build_workbook({"S1": rows}), MONTH)
    assert parsed.readings[0].observed_on == date(2026, 5, 10)
