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
| `CHITU_HISTORY_ROOT` | 项目同级 `赤兔历史分析结果` | 历史归档根目录（生产需指向可写目录，如 `/tmp`） |
| `CHITU_LLM_API_BASE` | — | OpenAI 兼容 API 根地址（必填） |
| `CHITU_LLM_API_KEY` | — | 服务器端 Bearer Token（必填，绝不放前端） |
| `CHITU_ANALYSIS_MODEL` | `gpt-5.6-sol` | 模型标识 |
| `CHITU_MAX_ACTIVE_JOBS` | `1` | 并发分析任务数 |
| `CHITU_LLM_WORKERS` | `6` | 分类/风险复核并发线程数 |

## 数据请求接口

`POST /api/tasks`（multipart：`payload` JSON + `files` 二进制）——创建任务
`GET /api/tasks/{task_id}` —— 轮询状态
`POST /api/tasks/{task_id}/retry` —— 进程内重试
`GET /api/download?path=...` —— 下载产物
`GET /api/health` —— 健康检查

## 常见坑与注意事项

- 服务是单进程内存队列 + `storage/jobs` 快照；**不能多实例部署**，重启后未完成任务不可原地重试（`SERVICE_RESTARTED`）。
- 一个聊天文件 = 一个商品/SKU；多商品必须分文件上传，不能拼接成单文件。
- 完整性门禁：`expected_messages` 必须等于 `analyzed_messages`，否则拒绝发布 Excel。
- Excel 样式只来自 `templates/excel/` 母版；基准文件只提供数据，不是样式来源。
- API Key 严禁写入 `web/`、`window.CHITU_CONFIG`、日志或接口响应。
- 生产环境 `CHITU_HISTORY_ROOT` 必须指向可写目录（沙箱 PROD 仅 `/tmp` 可写）。

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