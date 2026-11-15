"""并发上传与事务回滚测试。"""
from __future__ import annotations

import asyncio
from datetime import date
from decimal import Decimal

import httpx
import pytest

from app import db
from app.excel import ParsedWorkbook, Reading
from app.main import app
from tests.workbook_factory import build_workbook

MONTH = "2026-05"


def test_publish_rollback_keeps_old_data(client):
    """模拟发布事务中途失败：旧版本必须原样保留（含来源单元格）。"""
    import sqlite3

    content = build_workbook({"S1": [["长江大桥", date(2026, 5, 10), "P1", 1.5]]})
    resp = client.post(
        "/uploads",
        data={"month": MONTH},
        files={"file": ("a.xlsx", content, "application/octet-stream")},
    )
    assert resp.status_code == 200

    good_parsed = ParsedWorkbook(
        month=MONTH,
        readings=(Reading("B", date(2026, 5, 1), "P9", Decimal("1"), ("X!D1",)),),
        sheet_count=1,
        input_row_count=1,
        blank_rows=0,
        blank_settlement_rows=0,
    )

    # 用触发器让发布在写入 reading_sources 时失败，验证事务回滚
    setup = db._connect()
    setup.execute(
        "CREATE TRIGGER fail_source_insert BEFORE INSERT ON reading_sources "
        "BEGIN SELECT RAISE(ABORT, 'simulated disk failure'); END"
    )
    setup.close()
    try:
        with pytest.raises(sqlite3.Error, match="simulated disk failure"):
            db.publish_month(good_parsed)
    finally:
        cleanup = db._connect()
        cleanup.execute("DROP TRIGGER IF EXISTS fail_source_insert")
        cleanup.close()

    data = db.get_month(MONTH)
    assert data is not None
    assert data["version"] == 1
    assert len(data["readings"]) == 1
    assert data["readings"][0]["bridge"] == "长江大桥"
    assert data["readings"][0]["sources"] == ["S1!D2"]


async def _post_async(month: str, content: bytes, filename: str) -> httpx.Response:
    transport = httpx.ASGITransport(app=app)
    async with httpx.AsyncClient(transport=transport, base_url="http://test") as ac:
        return await ac.post(
            "/uploads",
            data={"month": month},
            files={"file": (filename, content, "application/octet-stream")},
        )


def test_concurrent_uploads_both_complete_and_single_version_visible(client):
    """两个并发上传：均成功且版本连续，期间任何查询只能看到完整版本。"""
    content_a = build_workbook({"S1": [["长江大桥", date(2026, 5, 10), "P1", 1.0]]})
    content_b = build_workbook({
        "S1": [
            ["黄河大桥", date(2026, 5, 11), "P1", 2.0],
            ["黄河大桥", date(2026, 5, 11), "P2", 3.0],
        ]
    })

    async def scenario():
        results = await asyncio.gather(
            _post_async(MONTH, content_a, "a.xlsx"),
            _post_async(MONTH, content_b, "b.xlsx"),
        )
        return results

    results = asyncio.run(scenario())
    assert {r.status_code for r in results} == {200}
    versions = sorted(r.json()["version"] for r in results)
    assert versions == [1, 2]

    # 最终状态是某一个完整版本，而不是混合数据
    final = client.get("/readings", params={"month": MONTH}).json()
    bridges = {r["bridge"] for r in final["readings"]}
    assert bridges in [{"长江大桥"}, {"黄河大桥"}]
    if bridges == {"黄河大桥"}:
        assert len(final["readings"]) == 2


def test_concurrent_valid_and_invalid_upload_keeps_consistent_version(client):
    good = build_workbook({"S1": [["长江大桥", date(2026, 5, 10), "P1", 1.0]]})
    bad = build_workbook({
        "S1": [["长江大桥", date(2026, 4, 30), "P1", 1.0]],  # 日期不属于 5 月
    })

    async def scenario():
        return await asyncio.gather(
            _post_async(MONTH, good, "good.xlsx"),
            _post_async(MONTH, bad, "bad.xlsx"),
        )

    results = asyncio.run(scenario())
    statuses = sorted(r.status_code for r in results)
    assert statuses == [200, 400]

    final = client.get("/readings", params={"month": MONTH})
    assert final.status_code == 200
    assert final.json()["version"] == 1
