# 赤兔客服聊天分析工作台

这是一个把客服聊天记录转换为正式需求分析报告的本地网页工作台。当前真实链路为：

```text
网页上传
  -> Node.js 异步任务服务
  -> Python 解析与全量 GPT/OpenAI 兼容模型分析
  -> 售前/售后校准、去重、风险二审、基准比较
  -> 固定 Excel 母版精确写回
  -> Excel + Markdown + result_manifest.json
```

当前系统不依赖 Coze 才能运行。`coze/` 目录保留为早期设计资料，不代表当前生产链路。

## 核心原则

- 分析内容由本次聊天记录重新产生，不套用历史报告的现成结论。
- 有基准时继承词条和统计口径，但仍逐条重新分析本期消息。
- Excel 只套用固定格式母版，基准文件不作为输出样式来源。
- 多个聊天文件代表多个商品/SKU，必须分别分析并生成商品组汇总。
- 全部有效买家消息必须分析完成；不得因模型超时跳过消息后继续发布正式报告。
- 当前默认模型标识为 `gpt-5.6-sol`，实际可用性取决于配置的 OpenAI 兼容 API。

## 必读交接文档

建议按以下顺序阅读：

1. [部署与运维交接文档](docs/05_部署与运维交接文档.md)  
   架构、依赖、配置、启动、网页部署、安全、队列、监控、备份、故障处理和上线验收。

2. [HTTP 接口文档](docs/06_HTTP接口文档.md)  
   创建任务、状态轮询、重试、下载、错误码、JSON 字段、前端示例和生产 V2 建议。

3. [分析方法与输出规范](docs/07_分析方法与输出规范.md)  
   文件解析、消息过滤、历史 GPT 逻辑、词条、正式需求、售前售后、去重、风险、基准对比、管理结论和全部 Excel Sheet。

专项资料：

- [正式文件结构](docs/01_正式文件结构.md)
- [网页界面设计](docs/02_网页界面设计.md)
- [历史结果存储规范](docs/03_赤兔历史分析结果_存储规范.md)
- [Excel 输出模板规范](docs/04_Excel输出模板规范.md)
- `rules/`：专项分析规则附件。
- `coze/`：早期 Coze 方案和提示词资料，仅供参考。

## 项目结构

```text
赤兔聊天分析工作台/
├─ 00_README.md
├─ local-server.mjs
├─ web/
├─ scripts/
├─ templates/excel/
├─ rules/
├─ docs/
├─ storage/jobs/
└─ storage/cache/llm/

同级目录：
赤兔历史分析结果/
├─ 00_索引/
├─ 01_用户上传/
├─ 02_过程文件/
└─ 03_最终报告/
```

## 快速启动

### 1. 配置服务器端模型

在项目根目录创建或维护 `.env.local`：

```dotenv
CHITU_LLM_API_BASE=https://your-openai-compatible-host
CHITU_LLM_API_KEY=replace-with-server-side-secret
CHITU_ANALYSIS_MODEL=gpt-5.6-sol
# 可选：推理模型思考档位，none/low/medium/high/xhigh/max，默认 low（留空则不传）
CHITU_LLM_REASONING_EFFORT=low
# 可选：级联初分类模型（默认 gpt-5.4-mini，低置信/风险消息升级主模型复核，等效成本约 -70%）；
# 留空关闭级联，全部分类用主模型
CHITU_CLASSIFY_MODEL=gpt-5.4-mini
```

成本优化内建于分类流水线：同文本消息去重（省 26~34%）、消息级缓存（重跑同文件几乎零调用）、大批次（110 条/批）。详见 `AGENTS.md`。

API Key 不能写进 `web/`。任何曾出现在聊天、截图或日志里的 Key，在公开部署前都应吊销并重新生成。

### 2. 安装 Python 依赖

```powershell
python -m pip install openpyxl
```

### 3. 启动

```powershell
cd E:\新2\赤兔聊天分析工作台
node local-server.mjs
```

打开：

```text
http://127.0.0.1:8787/
```

健康检查：

```powershell
Invoke-RestMethod http://127.0.0.1:8787/api/health
```

## 自动化测试

```powershell
python -m unittest discover -s scripts -p "test_*.py" -v
```

## 当前实现状态

已实现：

- 网页多文件上传和基准模式选择。
- 异步任务、进度轮询和同进程失败重试。
- CSV/LOG/TXT 解析和买家/客服过滤。
- 真实模型 API 分析、分块、并发、缓存和无限总尝试模式。
- 历史 GPT 分析逻辑加载。
- 基准词条继承和新增词条发现。
- 严格正式需求门禁和售前/售后校准。
- 同人同词条去重、风险二审、基准变化和管理结论。
- 单商品及多商品 Excel。
- 固定母版样式和图表保留。
- 全量消息完整性门禁。

当前限制：

- 无登录、用户权限和公网安全下载。
- 单进程内存队列，不能直接多实例部署。
- 服务重启不能恢复未完成任务，需重新上传。
- 状态和下载响应含服务器绝对路径，不适合公网。
- `2449` 和 `909` 专用脚本尚未接入网页主路由；除三款卷发棒外目前走通用正式分析器。
- 前端 `config.example.js` 仍保留旧 Coze 示例；当前网页真正使用 `service.tasksUrl`、`service.healthUrl` 和 `service.pollIntervalMs`。

公开部署前，请完成《部署与运维交接文档》中的 P0 安全和持久化改造。

