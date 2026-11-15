"""运行期配置，全部可通过环境变量覆盖。"""
from __future__ import annotations

import os
from dataclasses import dataclass


def _env_int(name: str, default: int) -> int:
    raw = os.getenv(name)
    return int(raw) if raw else default


def _env_float(name: str, default: float) -> float:
    raw = os.getenv(name)
    return float(raw) if raw else default


@dataclass
class Settings:
    # SQLite 数据库文件路径，容器内默认挂载在 /data 卷上
    db_path: str = os.getenv("DB_PATH", "/data/bridge_monitor.db")
    # 上传文件体积上限，默认 10 MiB（表本身很小，主要用于防御）
    max_upload_bytes: int = _env_int("MAX_UPLOAD_BYTES", 10 * 1024 * 1024)
    # SQLite 写锁等待时间，配合 WAL 支撑并发上传
    sqlite_timeout_seconds: float = _env_float("SQLITE_TIMEOUT_SECONDS", 15.0)


settings = Settings()
