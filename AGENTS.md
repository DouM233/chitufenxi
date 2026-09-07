# AGENTS.md

## 项目概览

「赤兔客服聊天分析工作台」：把客服聊天记录转化为正式需求分析报告（Excel）的本地网页应用。

真实链路：

```text
浏览器上传 (web/)
  -> Node.js 异步任务服务 (local-server.mjs)
  -> Python 解析 + OpenAI 兼容模型逐条分析 (scripts/)
  -> 售前/售后校准、去重、风险二审、基准对比
  -> 固定 Excel 母版精确写回 (templates/excel/)
  -> Excel + Markdown + result_manifest.json
```

系统不依赖 Coze；`coze/` 目录仅为早期设计资料。

## 技术栈

- 后端：Node.js（原生 `http`、`fetch`、`FormData`、Web Streams），无 npm 依赖
- 分析：Python 3.11+，依赖 `openpyxl`
- 前端：原生 HTML/CSS/JS（`web/`），无构建步骤
- 模型：任意 OpenAI 兼容 `/v1/chat/completions` 服务

## 目录结构

```text
.
├─ local-server.mjs            # Node HTTP 服务、任务队列、Python 子进程编排
├─ .env.local                  # 服务器端密钥（不入库，.gitignore 已屏蔽）
├─ web/                        # 前端静态资源（index.html / app.js / styles.css / config.js）
├─ scripts/                    # Python 分析脚本（generate_report.py 为通用主脚本）
├─ templates/excel/            # Excel 报告母版（发布资产）
├─ rules/                      # 人工可读分析规则
├─ docs/                       # 设计、接口、部署、分析规范
├─ storage/jobs/               # 任务状态快照（运行时）
├─ storage/cache/llm/          # 模型响应与分类块缓存（运行时）
└─ .coze                       # 沙箱/生产构建与运行配置
```

输出归档写入同级目录 `赤兔历史分析结果/`（可通过 `CHITU_HISTORY_ROOT` 覆盖）。

## 构建与运行

```bash
# 安装 Python 依赖
python3 -m pip install openpyxl

# 启动（端口取自 DEPLOY_RUN_PORT，默认 8787）
node local-server.mjs
```

健康检查：`curl http://localhost:${DEPLOY_RUN_PORT}/api/health`

自动化测试：`python3 -m unittest discover -s scripts -p "test_*.py" -v`

## 环境变量（关键）

| 变量 | 默认 | 说明 |
|---|---:|---|
| `DEPLOY_RUN_PORT` | — | 沙箱注入的监听端口（优先于 `CHITU_PORT`） |
| `CHITU_PORT` | `8787` | 本地开发端口 |
| `CHITU_HOST` | `0.0.0.0` | 监听地址（内网反代可改 `127.0.0.1`） |
| `CHITU_CORS_ORIGIN` | `*` | CORS 允许来源（无鉴权无 Cookie 时默认 `*`，可指定单一域名收紧跨域） |
| `CHITU_STATE_ROOT` | DEV=项目 `storage/`；PROD=`/tmp/chitu-state` | 运行时状态目录（任务快照 jobs + LLM 缓存 cache/llm） |
| `CHITU_HISTORY_ROOT` | DEV=项目同级 `赤兔历史分析结果`；PROD=`/tmp/chitu-history` | 历史归档根目录（生产需指向持久化存储，`/tmp` 会被清理） |
| `CHITU_HISTORY_LOG_FILE` | 自动查找项目内 `storage/` 及项目同级 `long_text_*.txt` | 已确认 GPT 分析口径文件（相对路径基于项目根目录；文件随部署产物自带，生产可不配） |
| `CHITU_LLM_API_BASE` | — | OpenAI 兼容 API 根地址（**生产部署必配**：`.env.local` 不进部署产物，需配在平台环境变量/密钥里） |
| `CHITU_LLM_API_KEY` | — | 服务器端 Bearer Token（**生产部署必配**，绝不放前端） |
| `CHITU_ANALYSIS_MODEL` | `gpt-5.6-sol` | 模型标识 |
| `CHITU_LLM_REASONING_EFFORT` | `low` | 推理模型思考档位（`none/low/medium/high/xhigh/max`，留空则不传该参数）。结构化分类任务用 `low` 即可，实测单批耗时约为默认档的 1/2~1/4；追求更快可试 `none`，追求更细分析用 `medium`。缓存键不含此参数，改档位后旧缓存仍可复用 |
| `CHITU_MAX_ACTIVE_JOBS` | `12` | 并发分析任务数；超过并发的任务才进入队列排队（可调大以减少排队；注意每个任务是独立 Python 子进程，内存峰值 ≈ 任务数 × 约 100MB） |
| `CHITU_LLM_WORKERS` | `3` | 单个任务内分类/风险复核并发线程数（LLM 总并发 ≈ 活跃任务数 × 此值）。上游限流（429）由 Python 内置指数退避重试兜底 |
| `COZE_BUCKET_ENDPOINT_URL` / `COZE_BUCKET_NAME` | 沙箱预置 | S3 兼容对象存储端点与桶名（**生产部署必配**，需在平台环境变量配置；两者缺一则对象存储禁用，自动降级本地行为） |
| `CHITU_S3_ACCESS_KEY_ID` / `CHITU_S3_SECRET_ACCESS_KEY` / `CHITU_S3_REGION` / `CHITU_S3_ENDPOINT` / `CHITU_S3_BUCKET` | — | 自建 S3 兼容存储的凭证与端点（配合 SDK 使用；region 默认 `cn-beijing`）。沙箱走 `COZE_BUCKET_*` 预置集成，无需配置 |

## 对象存储（object-storage.mjs）

- 产物（Excel/Markdown/manifest）与任务终态快照会同步上传对象存储；key 分别为 `chitu/reports/{task_id}/...`、`chitu/jobs/{task_id}.json`（SDK 会在文件名上追加 UUID 前缀，中文自动转 ASCII）。
- 下载三通道：data URL 内嵌（随任务结果返回，浏览器内存下载）→ `GET /api/tasks/{id}/file/{excel|markdown|manifest}` 动态生成签名 URL（前端 fetch+blob）→ 本地路径 `/api/download?path=...`（同实例兜底）。
- 任务查询内存 miss 时自动从对象存储按需恢复快照；服务启动时也会全量合并远端快照（本地优先，远端补缺）。
- 未配置对象存储时全部自动降级：不上传、不生成签名 URL，现有本地/内嵌下载行为完全不变。

## 数据请求接口

`POST /api/tasks`（multipart：`payload` JSON + `files` 二进制）——创建任务
`GET /api/tasks/{task_id}` —— 轮询状态（内存 miss 自动从对象存储恢复）
`GET /api/tasks/{task_id}/file/{kind}` —— 产物签名下载地址（kind: excel/markdown/manifest）
`POST /api/tasks/{task_id}/retry` —— 进程内重试（仅终态 failed/cancelled 可重试）
`POST /api/tasks/{task_id}/cancel` —— 取消分析（queued 直接移出队列；running 会 SIGTERM 终止 Python 子进程并释放并发槽位；终态返回 409）
`GET /api/tasks/{task_id}/file/{kind}` —— 按需生成产物签名下载地址（kind: excel/markdown/manifest）
`GET /api/download?path=...` —— 本地产物下载（同实例兜底）
`GET /api/health` —— 健康检查（含 storage 配置状态）

## 常见坑与注意事项

- 服务是单进程内存队列 + `storage/jobs` 快照；**进行中任务仍不能跨实例**（Python 子进程与内存队列是实例本地的），但已完成/失败/已取消任务的快照与产物已持久化到对象存储，实例回收/重启后任务状态与下载均可恢复；未完成任务重启后标记 `SERVICE_RESTARTED`。
- 任务状态机：`queued → running → completed | failed | cancelled`。cancelled 是终态并会同步对象存储快照；用户可对 failed/cancelled 任务点"重新分析"（内部走 retry，复用原始上传文件）。
- **生产部署引入 SDK 依赖后必须执行 `pnpm install`**（.coze 的 build 已配置）：若部署平台跳过依赖安装，服务会因 `ERR_MODULE_NOT_FOUND` 启动失败，表现为整站不可用/上传报错——排查生产"网站打不开/上传失败"时先看部署日志里 pnpm install 是否成功。
- 多任务并行是安全的：每个任务独立 Python 子进程，上传/过程/结果目录均按 task_id 隔离；LLM 缓存按内容哈希共享。上游限流（429）由 Python 内置指数退避重试兜底。
- 一个聊天文件 = 一个商品/SKU；多商品必须分文件上传，不能拼接成单文件。
- 完整性门禁：`expected_messages` 必须等于 `analyzed_messages`，否则拒绝发布 Excel。
- Excel 样式只来自 `templates/excel/` 母版；基准文件只提供数据，不是样式来源。
- API Key 严禁写入 `web/`、`window.CHITU_CONFIG`、日志或接口响应。
- 生产环境（`COZE_PROJECT_ENV=PROD` 或 `NODE_ENV=production`）下，运行时状态（`CHITU_STATE_ROOT`）与历史归档（`CHITU_HISTORY_ROOT`）默认都落在 `/tmp`（`/tmp/chitu-state`、`/tmp/chitu-history`）；`/tmp` 是临时目录且可能被清理，长期归档需接对象存储或持久卷。
- 产物下载双通道：任务完成时 Excel/Markdown/manifest 会内嵌为 data URL 随任务结果返回并写入任务快照（`result.files.*_data_url`），前端优先用 data URL 下载（浏览器内存完成，不受服务实例回收/`/tmp` 清理影响）；签名 URL 下载走 `GET /api/tasks/{id}/file/{kind}` 按需生成；`*_download_url` 服务端路径下载仅作同实例回退。
- **启动懒恢复**：`loadPersistedJobs` 只加载本地快照，不再启动时全量拉取对象存储（避免冷启动慢导致网关 502）；远端快照由 `getJob`/`restoreJobFromRemote` 在内存 miss 时按需恢复（带内存缓存）。
- **上传大小限制**：平台网关对请求体有限制（约 32MB，超限返回 413，请求不会到达服务）；前端在 `validate()` 做 25MB 预检（`MAX_UPLOAD_BYTES`）直接拦截并提示；413/502/503 在前端有针对性文案（502 通常是实例冷启动，等几秒重试）。

## 修复定位

- 上传/路由/队列/编排逻辑：`local-server.mjs`
- 前端交互：`web/app.js`
- 通用分析与报告生成：`scripts/generate_report.py`（主入口）、`scripts/llm_analysis.py`、`scripts/exact_excel_writer.py`
- 专用分析器（三款卷发棒/2449/909）：`scripts/analyze_*.py`
- Excel 母版：`templates/excel/*.xlsx`

## 安全红线

- `.env.local` 含真实 API Key，已被 `.gitignore` 屏蔽；任何曾暴露的 Key 上线前必须吊销重签。
- 前端与服务端同源，CORS 默认不对跨域开放。
- 当前版本无登录鉴权、下载接口使用绝对路径参数，公网部署前需完成 P0 改造（见 `docs/05_部署与运维交接文档.md`）。