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
- 分析：Python 3.11+，依赖 `openpyxl`（xlsx）、`xlrd`（xls 读取）
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
| `CHITU_CLASSIFY_MODEL` | `gpt-5.4-mini` | 级联初分类模型：先便宜模型粗筛，低置信/含风险/语义不全/疑问信号的消息升级主模型复核（实测升级率 ~45%，等效成本 -40%）。留空则关闭级联、全部用主模型 |
| `CHITU_CASCADE_CONFIDENCE` | `0.7` | 初分类结果低于此置信度升级主模型复核 |
| `CHITU_SCREENER_REASONING_EFFORT` | `none` | 初分类模型的思考档位 |
| `CHITU_SCREENER_CHUNK_MESSAGES` | `60` | 初分类批次大小（mini 在大批次下 JSON 完整性差，须小于主模型批次） |
| `CHITU_MESSAGE_CACHE` | `1` | 消息级缓存：按（模型+词条表哈希+商品+文本）缓存单条分类结果，改口径重跑/重复测试时几乎零调用。置 `0` 关闭 |
| `CHITU_CHUNK_MESSAGES` | `110` | 主模型分类批次大小上限（原 65；去重后大批次可摊薄每批固定的词条表+规则开销） |
| `CHITU_CHUNK_CHARS` | `24000` | 主模型分类批次字符数上限 |
| `CHITU_MAX_ACTIVE_JOBS` | `12` | 并发分析任务数；超过并发的任务才进入队列排队（可调大以减少排队；注意每个任务是独立 Python 子进程，内存峰值 ≈ 任务数 × 约 100MB） |
| `CHITU_LLM_WORKERS` | `12`（Node 服务注入；Python 直跑默认 `8`） | 单个任务内分类/风险复核并发线程数（LLM 总并发 ≈ 活跃任务数 × 此值）。实测上游 12 并发不限流；多人同时分析时可适当下调，单人使用可上调到 16~24 |

## 数据请求接口

`POST /api/tasks`（multipart：`payload` JSON + `files` 二进制）——创建任务
`GET /api/tasks/{task_id}` —— 轮询状态
`POST /api/tasks/{task_id}/retry` —— 进程内重试
`GET /api/download?path=...` —— 下载产物
`GET /api/health` —— 健康检查

## 常见坑与注意事项

- 分类流水线（`scripts/llm_analysis.py: classify_chunks`）：同文本去重（商品+文本，实测省 43~54%）→ 消息级缓存查询 → 级联初分类（`gpt-5.4-mini`，60 条/批）→ 确定性升级筛选（风险/低置信/语义不全/疑问信号）→ 主模型精分类 → 扇出回重复消息 → 逐消息售前/售后证据校准。完整性门禁始终按原始消息总数验收。
- 消息级缓存键包含词条表哈希：**改词条口径后旧缓存自动失效**，这是有意设计（口径变了结果不能复用）。
- 级联质量定位：与主模型单次跑的逐消息一致率约 62~68%，但主模型自身两次跑也只有 ~82% 一致（LLM 固有噪声），且多数分歧是初分类多检且合理。对 TOP 榜统计影响有限；保守场景把 `CHITU_CLASSIFY_MODEL` 留空即可回到纯主模型（去重/缓存/大批次仍生效）。
- 服务是单进程内存队列 + `storage/jobs` 快照；**不能多实例部署**，重启后未完成任务不可原地重试（`SERVICE_RESTARTED`）。
- 多任务并行是安全的：每个任务独立 Python 子进程，上传/过程/结果目录均按 task_id 隔离；LLM 缓存按内容哈希共享。上游限流（429）由 Python 内置指数退避重试兜底。
- 一个聊天文件 = 一个商品/SKU；多商品必须分文件上传，不能拼接成单文件。
- 聊天文件支持 `.txt/.log/.csv/.xlsx/.xls`。Excel 解析规则（`parse_excel`）：自动识别表头行（前 5 行内），按关键词映射 发送者/时间/内容 三列；时间单元格支持 datetime、Excel 序列号、中英文字符串格式，统一归一化为 `YYYY-MM-DD` + `HH:MM:SS`；无表头兜底：含换行的单元格按 log 头行格式（`昵称 日期 时间`）整块解析。前端上传时在 payload 中显式标注 `role`（chat_record/baseline），后端 `roleFromPayload` 优先采用，`inferFileRole` 仅作兜底（xlsx/xls 文件名不含聊天关键词会被判为基准文件）。
- 完整性门禁：`expected_messages` 必须等于 `analyzed_messages`，否则拒绝发布 Excel。
- Excel 样式只来自 `templates/excel/` 母版；基准文件只提供数据，不是样式来源。
- API Key 严禁写入 `web/`、`window.CHITU_CONFIG`、日志或接口响应。
- 生产环境（`COZE_PROJECT_ENV=PROD` 或 `NODE_ENV=production`）下，运行时状态（`CHITU_STATE_ROOT`）与历史归档（`CHITU_HISTORY_ROOT`）默认都落在 `/tmp`（`/tmp/chitu-state`、`/tmp/chitu-history`）；`/tmp` 是临时目录且可能被清理，长期归档需接对象存储或持久卷。

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