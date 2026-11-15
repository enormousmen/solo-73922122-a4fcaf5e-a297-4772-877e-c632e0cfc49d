"""FastAPI 应用入口：上传与按月查询接口。"""
from __future__ import annotations

import asyncio
from contextlib import asynccontextmanager

from fastapi import FastAPI, File, Form, HTTPException, Request, UploadFile, status
from fastapi.responses import JSONResponse

from . import db
from .config import settings
from .excel import WorkbookError, parse_month, parse_workbook
from .schemas import MonthResponse, MonthSummary, UploadResponse


def _get_upload_lock(app: FastAPI) -> asyncio.Lock:
    """返回绑定当前事件循环的上传串行锁（按循环惰性创建）。

    生产环境只有一个循环，锁全程唯一；这样写同时兼容测试中反复创建新循环。
    """
    loop = asyncio.get_running_loop()
    registry: dict[asyncio.AbstractEventLoop, asyncio.Lock] = app.state.upload_locks
    lock = registry.get(loop)
    if lock is None:
        lock = asyncio.Lock()
        registry[loop] = lock
    return lock


@asynccontextmanager
async def lifespan(app: FastAPI):
    app.state.upload_locks = {}
    db.init_db()
    yield


app = FastAPI(
    title="桥梁沉降监测数据服务",
    description=(
        "接收巡检班组的 XLSX 沉降表，校验后整月替换入仓，并提供按月查询。"
        "上传在单事务内完成发布，查询始终只能看到某一完整版本。"
    ),
    version="1.0.0",
    lifespan=lifespan,
)
# 不经 lifespan 的使用方式（如测试直连 ASGI）下也能拿到注册表
app.state.upload_locks = {}


@app.exception_handler(WorkbookError)
def _workbook_error_handler(_request, exc: WorkbookError) -> JSONResponse:
    return JSONResponse(
        status_code=status.HTTP_400_BAD_REQUEST,
        content={"detail": exc.message, "details": exc.details},
    )


@app.get("/health", tags=["运维"], summary="健康检查")
def health() -> dict[str, str]:
    return {"status": "ok"}


@app.post(
    "/uploads",
    response_model=UploadResponse,
    tags=["沉降数据"],
    summary="上传某月沉降表（成功即整月替换发布）",
    responses={400: {"description": "文件结构或内容校验失败（旧版本保留）"}},
)
async def upload_monthly_workbook(
    request: Request,
    month: str = Form(..., description="数据所属月份，格式 YYYY-MM"),
    file: UploadFile = File(..., description="XLSX 沉降表，至多 3 张工作表、每表 200 行"),
) -> UploadResponse:
    parse_month(month)  # 提前校验月份格式

    if not (file.filename or "").lower().endswith(".xlsx"):
        raise HTTPException(status_code=400, detail="仅支持 .xlsx 文件")

    content = await file.read()
    if not content:
        raise HTTPException(status_code=400, detail="上传文件为空")
    if len(content) > settings.max_upload_bytes:
        raise HTTPException(
            status_code=400,
            detail=f"文件体积 {len(content)} 字节超过上限 {settings.max_upload_bytes} 字节",
        )

    # 解析为 CPU 密集型操作，放到线程池；锁保护整个“解析 + 发布”过程
    async with _get_upload_lock(request.app):
        def _parse_and_publish():
            parsed = parse_workbook(content, month)  # 校验失败抛 WorkbookError -> 400
            return db.publish_month(parsed)  # 写入失败抛异常 -> 500，事务已回滚

        try:
            result = await asyncio.to_thread(_parse_and_publish)
        except WorkbookError:
            raise
        except HTTPException:
            raise
        except Exception as exc:  # 数据库写入失败：旧版本完好
            raise HTTPException(status_code=500, detail=f"数据写入失败，已保留旧版本：{exc}") from exc

    return UploadResponse(**result)


@app.get(
    "/readings",
    response_model=MonthResponse,
    tags=["沉降数据"],
    summary="按月查询已发布沉降读数",
    responses={404: {"description": "该月份尚无已发布数据"}},
)
def query_readings(month: str) -> MonthResponse:
    parse_month(month)
    data = db.get_month(month)
    if data is None:
        raise HTTPException(status_code=404, detail=f"月份 {month} 尚无已发布数据")
    return MonthResponse(**data)


@app.get(
    "/months",
    response_model=list[MonthSummary],
    tags=["沉降数据"],
    summary="列出所有已发布月份",
)
def list_months() -> list[dict]:
    return db.list_months()
