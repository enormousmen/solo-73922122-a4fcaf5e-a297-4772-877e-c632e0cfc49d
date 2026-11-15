"""生成一份演示用的合法沉降表（含纵向合并、零读数、空白沉降、跨表同值重复）。

用法：
    python scripts/make_sample_xlsx.py /tmp/sample_2026-05.xlsx
"""
from __future__ import annotations

import sys
from datetime import date
from pathlib import Path

from openpyxl import Workbook

HEADERS = ["桥梁", "观测日", "测点", "沉降毫米"]


def build(path: Path) -> None:
    wb = Workbook()
    ws1 = wb.active
    ws1.title = "甲班"
    ws1.append(HEADERS)
    ws1.append(["长江大桥", date(2026, 5, 10), "P1", 1.2])
    ws1.append([None, None, "P2", 0])          # 合并覆盖：沿用桥梁/观测日；零为有效读数
    ws1.append([None, None, "P3", None])       # 合并覆盖主信息；沉降空白 -> 不入仓
    ws1.append([None, None, "P4", -0.8])
    ws1.merge_cells("A2:A5")
    ws1.merge_cells("B2:B5")
    ws1.append(["黄河大桥", date(2026, 5, 11), "P1", 2.5])

    ws2 = wb.create_sheet("乙班")
    ws2.append(HEADERS)
    ws2.append(["长江大桥", date(2026, 5, 10), "P1", 1.2])  # 与甲班同值重复 -> 只存一条，来源全保留
    ws2.append(["黄河大桥", date(2026, 5, 11), "P2", -1.1])

    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    print(f"sample written: {path}")


if __name__ == "__main__":
    target = Path(sys.argv[1]) if len(sys.argv) > 1 else Path("sample_2026-05.xlsx")
    build(target)
