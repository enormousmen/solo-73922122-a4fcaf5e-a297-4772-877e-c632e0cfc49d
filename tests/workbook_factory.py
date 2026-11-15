"""测试辅助：用 openpyxl 构造沉降表工作簿。"""
from __future__ import annotations

from datetime import date
from io import BytesIO
from typing import Any

from openpyxl import Workbook

HEADERS = ["桥梁", "观测日", "测点", "沉降毫米"]


def build_workbook(sheets: dict[str, list[list[Any]]], merges: dict[str, list[str]] | None = None) -> bytes:
    """构造 xlsx 字节串。

    sheets: {工作表名: 行数据}，行数据**不含表头**，每个单元格给原始值；
    merges: {工作表名: ["A2:A4", ...]} 需要合并的区域（合并前请先把值放在左上角）。
    """
    wb = Workbook()
    wb.remove(wb.active)
    merges = merges or {}
    for name, rows in sheets.items():
        ws = wb.create_sheet(title=name)
        ws.append(HEADERS)
        for row in rows:
            ws.append(row)
        for ref in merges.get(name, []):
            ws.merge_cells(ref)
    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()


def standard_rows(start_day: int = 5) -> list[list[Any]]:
    """一份常规合法数据（2026-05），含零读数与空白沉降。"""
    return [
        ["长江大桥", date(2026, 5, start_day), "P1", 1.2],
        ["长江大桥", date(2026, 5, start_day), "P2", 0],          # 零是有效读数
        ["长江大桥", date(2026, 5, start_day), "P3", None],       # 数值空白不入仓
        ["黄河大桥", date(2026, 5, start_day + 1), "P1", -2.5],
    ]
