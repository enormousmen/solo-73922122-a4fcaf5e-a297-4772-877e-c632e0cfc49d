"""SQLite 存储层。

并发与版本一致性设计：
* WAL 模式下读写不互相阻塞；
* 每月发布在**单个事务**内 DELETE + INSERT，任何一步失败整体回滚，
  查询要么看到旧版本、要么看到新版本，不会看到半成品；
* 应用层另用 asyncio.Lock 串行化上传（见 main.py），SQLite busy_timeout
  负责多进程/外部连接场景下的写锁等待；
* sources 以 JSON 文本保存全部来源单元格，读数使用 TEXT 存 Decimal 精确值。
"""
from __future__ import annotations

import json
import sqlite3
from contextlib import contextmanager
from datetime import date, datetime, timezone
from decimal import Decimal
from typing import Any, Iterator

from .config import settings
from .excel import Reading

SCHEMA_SQL = """
PRAGMA journal_mode = WAL;
PRAGMA foreign_keys = ON;

CREATE TABLE IF NOT EXISTS month_versions (
    month                     TEXT PRIMARY KEY,
    version                   INTEGER NOT NULL,
    published_at              TEXT NOT NULL,
    input_row_count           INTEGER NOT NULL,
    reading_count             INTEGER NOT NULL,
    sheet_count               INTEGER NOT NULL,
    blank_rows                INTEGER NOT NULL,
    blank_settlement_rows     INTEGER NOT NULL
);

CREATE TABLE IF NOT EXISTS readings (
    id             INTEGER PRIMARY KEY AUTOINCREMENT,
    month          TEXT NOT NULL,
    bridge         TEXT NOT NULL,
    observed_on    TEXT NOT NULL,
    point          TEXT NOT NULL,
    settlement_mm  TEXT NOT NULL,
    version        INTEGER NOT NULL,
    created_at     TEXT NOT NULL,
    UNIQUE (month, bridge, observed_on, point)
);

CREATE TABLE IF NOT EXISTS reading_sources (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    reading_id  INTEGER NOT NULL REFERENCES readings(id) ON DELETE CASCADE,
    source_ref  TEXT NOT NULL
);

CREATE INDEX IF NOT EXISTS idx_readings_month
    ON readings (month, bridge, observed_on, point);
"""


def _connect() -> sqlite3.Connection:
    conn = sqlite3.connect(
        settings.db_path,
        timeout=settings.sqlite_timeout_seconds,
        isolation_level=None,  # 显式管理事务
    )
    conn.row_factory = sqlite3.Row
    conn.execute(f"PRAGMA busy_timeout = {int(settings.sqlite_timeout_seconds * 1000)}")
    conn.execute("PRAGMA journal_mode = WAL")
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


@contextmanager
def _connection() -> Iterator[sqlite3.Connection]:
    conn = _connect()
    try:
        yield conn
    finally:
        conn.close()


def init_db() -> None:
    """创建表结构并启用 WAL，可重复执行。"""
    with _connection() as conn:
        conn.executescript(SCHEMA_SQL)


def publish_month(parsed: Any) -> dict[str, Any]:
    """在单事务内整月替换发布。

    返回发布元信息；失败时抛异常且事务回滚，旧版本保持不变。
    """
    month = parsed.month
    now = datetime.now(timezone.utc).isoformat()

    with _connection() as conn:
        try:
            conn.execute("BEGIN IMMEDIATE")

            row = conn.execute(
                "SELECT version FROM month_versions WHERE month = ?", (month,)
            ).fetchone()
            replaced = row is not None
            new_version = (row["version"] + 1) if replaced else 1

            # 整月旧数据整体删除（子表通过外键级联）
            conn.execute("DELETE FROM reading_sources WHERE reading_id IN "
                         "(SELECT id FROM readings WHERE month = ?)", (month,))
            conn.execute("DELETE FROM readings WHERE month = ?", (month,))
            conn.execute("DELETE FROM month_versions WHERE month = ?", (month,))

            conn.execute(
                "INSERT INTO month_versions "
                "(month, version, published_at, input_row_count, reading_count, "
                " sheet_count, blank_rows, blank_settlement_rows) "
                "VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
                (
                    month,
                    new_version,
                    now,
                    parsed.input_row_count,
                    len(parsed.readings),
                    parsed.sheet_count,
                    parsed.blank_rows,
                    parsed.blank_settlement_rows,
                ),
            )

            for reading in parsed.readings:
                cur = conn.execute(
                    "INSERT INTO readings "
                    "(month, bridge, observed_on, point, settlement_mm, version, created_at) "
                    "VALUES (?, ?, ?, ?, ?, ?, ?)",
                    (
                        month,
                        reading.bridge,
                        reading.observed_on.isoformat(),
                        reading.point,
                        str(reading.settlement_mm),
                        new_version,
                        now,
                    ),
                )
                reading_id = cur.lastrowid
                conn.executemany(
                    "INSERT INTO reading_sources (reading_id, source_ref) VALUES (?, ?)",
                    [(reading_id, ref) for ref in reading.sources],
                )

            conn.execute("COMMIT")
        except Exception:
            conn.execute("ROLLBACK")
            raise

    return {
        "month": month,
        "version": new_version,
        "replaced": replaced,
        "published_at": now,
        "sheet_count": parsed.sheet_count,
        "input_row_count": parsed.input_row_count,
        "reading_count": len(parsed.readings),
        "blank_rows": parsed.blank_rows,
        "blank_settlement_rows": parsed.blank_settlement_rows,
    }


def get_month(month: str) -> dict[str, Any] | None:
    """读取某个已发布月份的完整版本；未发布返回 None。"""
    with _connection() as conn:
        version_row = conn.execute(
            "SELECT * FROM month_versions WHERE month = ?", (month,)
        ).fetchone()
        if version_row is None:
            return None

        rows = conn.execute(
            "SELECT * FROM readings WHERE month = ? "
            "ORDER BY bridge, observed_on, point",
            (month,),
        ).fetchall()

        sources_map: dict[int, list[str]] = {}
        if rows:
            placeholders = ",".join("?" for _ in rows)
            source_rows = conn.execute(
                f"SELECT reading_id, source_ref FROM reading_sources "
                f"WHERE reading_id IN ({placeholders}) ORDER BY id",
                [r["id"] for r in rows],
            ).fetchall()
            for sr in source_rows:
                sources_map.setdefault(sr["reading_id"], []).append(sr["source_ref"])

        readings = [
            {
                "bridge": r["bridge"],
                "observed_on": date.fromisoformat(r["observed_on"]),
                "point": r["point"],
                "settlement_mm": Decimal(r["settlement_mm"]),
                "sources": sources_map.get(r["id"], []),
            }
            for r in rows
        ]
        return {
            "month": month,
            "version": version_row["version"],
            "published_at": version_row["published_at"],
            "input_row_count": version_row["input_row_count"],
            "reading_count": version_row["reading_count"],
            "sheet_count": version_row["sheet_count"],
            "blank_rows": version_row["blank_rows"],
            "blank_settlement_rows": version_row["blank_settlement_rows"],
            "readings": readings,
        }


def list_months() -> list[dict[str, Any]]:
    """列出所有已发布月份（不含明细）。"""
    with _connection() as conn:
        rows = conn.execute(
            "SELECT month, version, published_at, reading_count "
            "FROM month_versions ORDER BY month"
        ).fetchall()
        return [dict(r) for r in rows]
