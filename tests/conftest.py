"""pytest 夹具：隔离的临时 SQLite 与 FastAPI TestClient。"""
from __future__ import annotations

import os
import sys
import tempfile
from pathlib import Path

import pytest
from fastapi.testclient import TestClient

# 保证可以 import app（pytest 从仓库根目录运行时通常已在 sys.path，这里再兜底一次）
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))


@pytest.fixture()
def client(monkeypatch):
    tmp_dir = tempfile.mkdtemp(prefix="bridge_test_")
    db_path = os.path.join(tmp_dir, "test.db")
    monkeypatch.setenv("DB_PATH", db_path)

    # config 在 import 时读取环境变量，因此先设好再导入
    from app import db as db_module
    from app.config import settings
    from app.main import app

    settings.db_path = db_path
    db_module.init_db()

    with TestClient(app) as c:
        yield c

    for suffix in ("", "-wal", "-shm"):
        p = Path(db_path + suffix)
        if p.exists():
            p.unlink()
    Path(tmp_dir).rmdir()


def upload(client, month: str, content: bytes, filename: str = "settlement.xlsx"):
    return client.post(
        "/uploads",
        data={"month": month},
        files={"file": (filename, content, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
    )
