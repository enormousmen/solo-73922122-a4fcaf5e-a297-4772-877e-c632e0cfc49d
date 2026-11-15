# 桥梁沉降监测数据服务

接收巡检班组每月提交的 XLSX 沉降表，完成结构与业务校验后整月替换入仓
（SQLite），并提供按月查询 API。技术栈：Python 3.11 / FastAPI / openpyxl /
SQLite，Docker Compose 一键启动。

## 快速启动

```bash
docker compose up --build
```

启动后：

- 交互式接口文档（Swagger UI，可直接上传试调）：<http://localhost:8000/docs>
- ReDoc 文档：<http://localhost:8000/redoc>
- 健康检查：<http://localhost:8000/health>

SQLite 数据文件持久化在命名卷 `bridge_data`（容器内 `/data/bridge_monitor.db`）。

本地无 Docker 时的开发方式：

```bash
python -m venv .venv && . .venv/bin/activate
pip install -r requirements.txt
DB_PATH=./bridge_monitor.db uvicorn app.main:app --reload
```

## 接口

| 方法 | 路径 | 说明 |
| --- | --- | --- |
| POST | `/uploads` | 表单字段 `month=YYYY-MM` + 文件 `file`（.xlsx），校验通过即整月替换发布 |
| GET | `/readings?month=YYYY-MM` | 查询某月已发布版本的全部读数及来源单元格 |
| GET | `/months` | 列出所有已发布月份 |
| GET | `/health` | 健康检查 |

生成演示文件并调用：

```bash
python scripts/make_sample_xlsx.py /tmp/sample_2026-05.xlsx
curl -F "month=2026-05" -F "file=@/tmp/sample_2026-05.xlsx" http://localhost:8000/uploads
curl "http://localhost:8000/readings?month=2026-05"
```

## 表格约定

- 工作簿至多 **3 张工作表**，每表至多 **200 个数据行**（整行空白不计入限额）。
- 首行固定四列表头：`桥梁 | 观测日 | 测点 | 沉降毫米`，出现额外表头列直接拒收。
- **桥梁、观测日**允许纵向合并（不得跨列；测点、沉降列不允许合并）；
  合并区域覆盖的行沿用左上角值。
- 普通空白单元格**不沿用**：桥梁/观测日/测点空白即整份拒收。
- `沉降毫米` 为空白的行不入仓；`0` 是有效读数，正常入仓。
- `观测日` 必须落在上传指定月份内，否则整份拒收（支持 Excel 日期单元格与
  `YYYY-MM-DD` 等常见文本格式）。
- 同一 `(桥梁, 观测日, 测点)`：
  - 读数相同（含跨表重复）：只存一条，但响应中保留**全部来源单元格**
    （如 `甲班!D2`、`乙班!D2`）；
  - 读数不同：返回 400，错误体列出冲突读数与全部来源单元格，**整份拒收**。

## 一致性与并发保证

1. **原子发布**：每月的 DELETE + INSERT 在单个 SQLite 事务
   （`BEGIN IMMEDIATE` ... `COMMIT`）内完成；校验失败或写入失败均回滚，
   已发布的旧版本保持不变。
2. **版本快照语义**：查询在独立连接中执行，配合 WAL 模式，只能看到某个完整
   已提交版本，绝不会读到半成品数据。
3. **并发上传**：进程内 `asyncio.Lock` 串行化发布；SQLite `busy_timeout` +
   WAL 负责多连接/多进程层面的写锁等待与读写并发。每次成功发布版本号 +1。

## 数据模型

- `month_versions`：每月一行，记录版本号、发布时间、行数统计。
- `readings`：去重后的读数，`(month, bridge, observed_on, point)` 唯一，
  沉降值以 TEXT 存储 Decimal 精确值。
- `reading_sources`：读数对应的全部来源单元格（外键级联删除）。

## 测试

```bash
pip install -r requirements.txt httpx==0.28.1 pytest==8.3.4
pytest -q
```

覆盖：合并单元格继承/普通空白拒收、跨月日期拒收、零与空白沉降、
跨表同值/冲突、表/行上限、整月替换、失败保留旧版、并发上传与事务回滚。
