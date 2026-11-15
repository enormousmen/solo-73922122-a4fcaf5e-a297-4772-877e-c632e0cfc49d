FROM python:3.11-slim

ENV PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1 \
    PIP_NO_CACHE_DIR=1

WORKDIR /app

# 先装依赖，利用镜像层缓存
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY app ./app

# 数据卷：SQLite 数据库与 WAL 文件都落在这里
RUN mkdir -p /data
VOLUME ["/data"]

ENV DB_PATH=/data/bridge_monitor.db \
    MAX_UPLOAD_BYTES=10485760

EXPOSE 8000

# 容器内健康检查，compose 的 healthcheck 依赖此接口
HEALTHCHECK --interval=10s --timeout=3s --start-period=5s --retries=5 \
    CMD python -c "import urllib.request,sys; sys.exit(0 if urllib.request.urlopen('http://127.0.0.1:8000/health').status==200 else 1)"

CMD ["uvicorn", "app.main:app", "--host", "0.0.0.0", "--port", "8000"]
