"""端到端 API 测试：上传、替换、查询、错误语义与版本原子性。"""
from __future__ import annotations

from datetime import date

from .conftest import upload
from .workbook_factory import build_workbook, standard_rows

MONTH = "2026-05"


def test_health(client):
    resp = client.get("/health")
    assert resp.status_code == 200
    assert resp.json() == {"status": "ok"}


def test_upload_then_query(client):
    resp = upload(client, MONTH, build_workbook({"S1": standard_rows()}))
    assert resp.status_code == 200, resp.text
    body = resp.json()
    assert body["month"] == MONTH
    assert body["version"] == 1
    assert body["replaced"] is False
    assert body["reading_count"] == 3
    assert body["blank_settlement_rows"] == 1

    q = client.get("/readings", params={"month": MONTH})
    assert q.status_code == 200
    data = q.json()
    assert data["version"] == 1
    assert len(data["readings"]) == 3
    p2 = next(r for r in data["readings"] if r["point"] == "P2" and r["bridge"] == "长江大桥")
    assert p2["settlement_mm"] == "0"
    assert p2["sources"] == ["S1!D3"]


def test_query_unknown_month_404(client):
    resp = client.get("/readings", params={"month": "2030-01"})
    assert resp.status_code == 404


def test_invalid_month_format_400(client):
    resp = upload(client, "2026/05", build_workbook({"S1": standard_rows()}))
    assert resp.status_code == 400


def test_validation_failure_keeps_old_version(client):
    # 先发一个合法版本
    first = upload(client, MONTH, build_workbook({"S1": standard_rows()}))
    assert first.status_code == 200

    # 再传含跨表冲突的文件 -> 拒收
    bad = build_workbook({
        "S1": [["长江大桥", date(2026, 5, 10), "P1", 1.5]],
        "S2": [["长江大桥", date(2026, 5, 10), "P1", 1.6]],
    })
    bad_resp = upload(client, MONTH, bad)
    assert bad_resp.status_code == 400
    detail = bad_resp.json()["details"][0]
    assert detail["来源单元格"] == ["S1!D2", "S2!D2"]

    # 旧版本完好
    q = client.get("/readings", params={"month": MONTH})
    assert q.status_code == 200
    assert q.json()["version"] == 1
    assert len(q.json()["readings"]) == 3


def test_successful_upload_replaces_whole_month(client):
    upload(client, MONTH, build_workbook({"S1": standard_rows()}))
    replacement = build_workbook({
        "S2": [
            ["珠江大桥", date(2026, 5, 2), "X1", 9.9],
            ["珠江大桥", date(2026, 5, 2), "X2", 0],
        ]
    })
    resp = upload(client, MONTH, replacement)
    assert resp.status_code == 200
    body = resp.json()
    assert body["replaced"] is True
    assert body["version"] == 2

    q = client.get("/readings", params={"month": MONTH}).json()
    assert q["version"] == 2
    assert len(q["readings"]) == 2
    assert {r["bridge"] for r in q["readings"]} == {"珠江大桥"}


def test_different_months_are_independent(client):
    upload(client, "2026-05", build_workbook({"S1": standard_rows()}))
    june = build_workbook({"S1": [["长江大桥", date(2026, 6, 1), "P1", 3.3]]})
    upload(client, "2026-06", june)

    may = client.get("/readings", params={"month": "2026-05"}).json()
    jun = client.get("/readings", params={"month": "2026-06"}).json()
    assert may["version"] == 1 and jun["version"] == 1
    assert len(may["readings"]) == 3 and len(jun["readings"]) == 1

    months = client.get("/months").json()
    assert [m["month"] for m in months] == ["2026-05", "2026-06"]


def test_non_xlsx_extension_rejected(client):
    resp = client.post(
        "/uploads",
        data={"month": MONTH},
        files={"file": ("data.txt", b"hello", "text/plain")},
    )
    assert resp.status_code == 400


def test_merged_cells_end_to_end(client):
    rows = [
        ["长江大桥", date(2026, 5, 10), "P1", 1.0],
        [None, None, "P2", 2.0],
        [None, None, "P3", 3.0],
    ]
    content = build_workbook({"S1": rows}, merges={"S1": ["A2:A4", "B2:B4"]})
    resp = upload(client, MONTH, content)
    assert resp.status_code == 200, resp.text
    data = client.get("/readings", params={"month": MONTH}).json()
    assert all(r["bridge"] == "长江大桥" for r in data["readings"])
    assert all(r["observed_on"] == "2026-05-10" for r in data["readings"])


def test_sources_preserved_for_duplicate_and_conflict_message(client):
    content = build_workbook({
        "S1": [["长江大桥", date(2026, 5, 10), "P1", 1.5]],
        "S2": [["长江大桥", date(2026, 5, 10), "P1", 1.5]],
    })
    resp = upload(client, MONTH, content)
    assert resp.status_code == 200
    reading = client.get("/readings", params={"month": MONTH}).json()["readings"][0]
    assert reading["sources"] == ["S1!D2", "S2!D2"]
