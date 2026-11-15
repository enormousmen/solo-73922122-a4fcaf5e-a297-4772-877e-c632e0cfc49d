"""请求/响应 Pydantic 模型。"""
from __future__ import annotations

from datetime import date
from decimal import Decimal

from pydantic import BaseModel, Field


class SourceReading(BaseModel):
    bridge: str = Field(..., description="桥梁名称")
    observed_on: date = Field(..., description="观测日")
    point: str = Field(..., description="测点编号/名称")
    settlement_mm: Decimal = Field(..., description="沉降量（毫米），0 为有效读数")
    sources: list[str] = Field(..., description="全部来源单元格，如 Sheet1!D7")


class UploadResponse(BaseModel):
    month: str
    version: int = Field(..., description="该月发布版本号，首次为 1，每次成功替换 +1")
    replaced: bool = Field(..., description="是否替换了该月的旧版本")
    published_at: str
    sheet_count: int
    input_row_count: int = Field(..., description="解析到的非空数据行数（不含整行空白）")
    reading_count: int = Field(..., description="去重后实际入仓的读数条数")
    blank_rows: int = Field(..., description="整行四列皆空、被忽略的行数（不计入 200 行限额）")
    blank_settlement_rows: int = Field(..., description="沉降毫米为空白、未入仓的数据行数")


class MonthResponse(BaseModel):
    month: str
    version: int
    published_at: str
    input_row_count: int
    reading_count: int
    sheet_count: int
    blank_rows: int
    blank_settlement_rows: int
    readings: list[SourceReading]


class MonthSummary(BaseModel):
    month: str
    version: int
    published_at: str
    reading_count: int


class ErrorResponse(BaseModel):
    detail: str
    details: object | None = None
