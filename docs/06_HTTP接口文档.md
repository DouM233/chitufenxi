# 赤兔客服聊天分析工作台：HTTP 接口文档

> 接口版本：V1（当前实现）  
> 基础地址（本地）：`http://127.0.0.1:8787`  
> 数据格式：除文件下载外均为 UTF-8 JSON  
> 任务模式：异步创建、轮询状态、完成后下载

## 1. 接口总览

| 方法 | 路径 | 用途 | 当前认证 |
|---|---|---|---|
| `GET` | `/api/health` | 服务健康和队列状态 | 无 |
| `POST` | `/api/tasks` | 创建分析任务 | 无 |
| `POST` | `/api/chitu-analyze` | 创建任务的兼容别名 | 无 |
| `GET` | `/api/tasks/{task_id}` | 查询任务状态与结果 | 无 |
| `POST` | `/api/tasks/{task_id}/retry` | 重试当前进程内的失败任务 | 无 |
| `GET` | `/api/download?path=...` | 下载生成文件 | 无 |
| `OPTIONS` | 任意路径 | CORS 预检 | 无 |

生产环境必须增加认证、授权和任务归属校验。本文先描述当前代码已经实现的 V1 契约，再给出生产版建议。

## 2. 通用约定

### 2.1 时间

接口中的 `created_at`、`updated_at` 使用 ISO 8601 时间字符串。结果清单中的 `created_at` 为服务器本地格式文本，调用方不应依赖它排序，应以任务接口时间为准。

### 2.2 任务 ID

`task_id` 由服务器创建，当前为 UUID，例如：

```text
5e810c18-07f0-49ef-8aab-39eaa945b92d
```

客户端必须把它视为不透明字符串，不能解析其中含义。

### 2.3 Content-Type

- 创建任务：`multipart/form-data; boundary=...`
- 其他 JSON 接口：`application/json; charset=utf-8`
- 下载：根据扩展名返回 MIME，未知扩展名为 `application/octet-stream`

### 2.4 请求限制

| 项目 | 当前限制 |
|---|---:|
| 单次 HTTP 请求总大小 | 30 MB |
| 单个文件大小 | 20 MB |
| 单次文件数量 | 10 个 |
| 本次聊天记录 | 至少 1 个 |
| 基准文件 | 基准对比模式建议且前端要求 1 个 |

超过限制返回 HTTP `413`。

### 2.5 文件角色

`payload.files` 中的角色优先于文件名推断。

| role | 含义 | 常见格式 |
|---|---|---|
| `chat_record` | 本次聊天记录，每个文件代表一个商品/SKU | `.csv`、`.log`、`.txt` |
| `baseline` | 上一次正式基准报告 | `.xlsx`、`.xls` |
| `reference` | 其他参考资料；当前通用分析流程只归档，不作为主聊天输入 | 任意受支持文件 |

如果角色缺失，服务器根据文件名和扩展名推断：包含“基准、baseline、上期、上一期、历史”等倾向基准；包含“聊天、chat、客服、raw、fulltext、记录”或为 CSV/LOG/TXT 倾向聊天；否则为参考文件。

## 3. 健康检查

### `GET /api/health`

用于负载均衡存活检查和前端连接检查。

#### 请求示例

```http
GET /api/health HTTP/1.1
Host: 127.0.0.1:8787
```

#### 成功响应

HTTP `200 OK`

```json
{
  "status": "ok",
  "active_jobs": 1,
  "queued_jobs": 2
}
```

#### 字段说明

| 字段 | 类型 | 说明 |
|---|---|---|
| `status` | string | 当前固定为 `ok`。只说明 Node 进程可响应，不代表模型 API 可用。 |
| `active_jobs` | integer | 当前正在执行的任务数。 |
| `queued_jobs` | integer | 当前等待执行的任务数。 |

健康接口目前不会主动测试 Python、Excel 母版、磁盘权限或模型 API。生产版建议拆分 `/health/live` 与 `/health/ready`。

## 4. 创建分析任务

### `POST /api/tasks`

兼容别名：`POST /api/chitu-analyze`。

接口接收一个 JSON 字符串字段 `payload` 和一个或多个重复的二进制字段 `files`。

### 4.1 Multipart 字段

| 字段 | 类型 | 必填 | 说明 |
|---|---|---:|---|
| `payload` | string(JSON) | 是 | 任务说明、分析类型和文件描述。 |
| `files` | binary，可重复 | 是 | 实际上传文件。至少一个聊天文件。 |

### 4.2 payload 结构

```json
{
  "user_message": "产品：2449小气泡；分析周期：2026-08-26 至 2026-09-01；与上次基准对比。",
  "analysis_type": "baseline_compare",
  "input_mode": "baseline_and_chat",
  "model": "gpt-5.6-sol",
  "files": [
    {
      "name": "2449小气泡聊天记录.log",
      "size": 123456,
      "type": "text/plain",
      "role": "chat_record"
    },
    {
      "name": "2449小气泡_上期基准.xlsx",
      "size": 654321,
      "type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      "role": "baseline"
    }
  ]
}
```

### 4.3 payload 字段

| 字段 | 类型 | 必填 | 允许值/默认 | 说明 |
|---|---|---:|---|---|
| `user_message` | string | 建议 | 空字符串 | 产品名、周期和分析意图。也用于任务类型和产品识别。 |
| `analysis_type` | string | 建议 | 自动推断 | `first_baseline` 或 `baseline_compare`。 |
| `input_mode` | string | 否 | 前端辅助 | `chat_only` 或 `baseline_and_chat`；服务端主要依据 `analysis_type` 和文件角色。 |
| `model` | string | 否 | 环境变量模型 | 写入任务资料并传递当前配置。生产环境不建议允许普通用户任意覆盖。 |
| `files` | array | 强烈建议 | `[]` | 文件名、大小、MIME 和角色描述。服务端按“文件名 + 大小”匹配实际 multipart 文件。 |

`analysis_type`：

- `first_baseline`：没有历史基准，基于本期聊天建立词条并生成首次基准。
- `baseline_compare`：上传历史正式报告，继承词条口径并生成本期对比。

### 4.4 多商品上传规则

- 1 个聊天文件 = 1 个商品/SKU。
- 3 个聊天文件 = 3 个商品/SKU。
- 4 个聊天文件 = 4 个商品/SKU。
- 文件不得先拼接成一个大文件，否则系统无法可靠区分商品。
- 文件名应包含商品/SKU 名，例如 `S-ZY-F4-5WH.log`。
- 有多商品基准时，本期聊天文件数量必须与基准商品数量一致，并且名称能够匹配。

### 4.5 浏览器 JavaScript 示例

```javascript
async function createAnalysisTask({ message, chatFiles, baselineFile }) {
  const descriptors = [];
  const form = new FormData();

  for (const file of chatFiles) {
    form.append("files", file, file.name);
    descriptors.push({
      name: file.name,
      size: file.size,
      type: file.type,
      role: "chat_record"
    });
  }

  if (baselineFile) {
    form.append("files", baselineFile, baselineFile.name);
    descriptors.push({
      name: baselineFile.name,
      size: baselineFile.size,
      type: baselineFile.type,
      role: "baseline"
    });
  }

  form.append("payload", JSON.stringify({
    user_message: message,
    analysis_type: baselineFile ? "baseline_compare" : "first_baseline",
    input_mode: baselineFile ? "baseline_and_chat" : "chat_only",
    files: descriptors
  }));

  const response = await fetch("/api/tasks", {
    method: "POST",
    body: form
  });

  const body = await response.json();
  if (!response.ok) throw new Error(body.message || `HTTP ${response.status}`);
  return body;
}
```

浏览器会自动生成 multipart boundary，不要手动设置 `Content-Type`。

### 4.6 curl 示例（Bash）

```bash
payload='{
  "user_message":"产品：化妆镜；分析周期：2026-08-02 至 2026-08-31；没有历史基准，建立本期分析。",
  "analysis_type":"first_baseline",
  "input_mode":"chat_only",
  "files":[
    {"name":"化妆镜.log","size":123456,"type":"text/plain","role":"chat_record"}
  ]
}'

curl -X POST 'http://127.0.0.1:8787/api/tasks' \
  -F "payload=$payload" \
  -F 'files=@/data/化妆镜.log;type=text/plain'
```

描述中的 `size` 必须与实际文件大小一致，否则角色描述可能匹配失败并退回文件名推断。

### 4.7 成功响应

HTTP `202 Accepted`

```json
{
  "task_id": "5e810c18-07f0-49ef-8aab-39eaa945b92d",
  "status": "queued",
  "status_url": "/api/tasks/5e810c18-07f0-49ef-8aab-39eaa945b92d"
}
```

`202` 只表示任务已进入队列，不表示分析完成。

### 4.8 同步错误

#### 请求过大

HTTP `413 Payload Too Large`

```json
{
  "status": "failed",
  "message": "单次上传不能超过 30 MB。"
}
```

或：

```json
{
  "status": "failed",
  "message": "单次最多上传 10 个文件，单个文件不能超过 20 MB。"
}
```

#### payload 非法

当前实现会由统一异常处理返回 HTTP `500`。生产版应改为 HTTP `400` 并返回稳定错误码 `INVALID_PAYLOAD`。

#### 没有聊天文件

当前任务可能先成功创建，再在异步执行阶段失败。失败信息：

```text
请至少上传本次聊天记录文件。若要做基准对照，请同时上传上次分析基准。
```

生产版建议在创建任务时直接返回 HTTP `422`。

## 5. 查询任务状态

### `GET /api/tasks/{task_id}`

前端默认每 1500 毫秒查询一次。生产环境建议使用 2 至 5 秒轮询，并在长任务中逐步退避。

### 5.1 响应结构

HTTP `200 OK`

```json
{
  "task_id": "5e810c18-07f0-49ef-8aab-39eaa945b92d",
  "status": "running",
  "stage": "classifying",
  "progress": 52,
  "completed_chunks": 18,
  "total_chunks": 40,
  "message": "正在逐批分析有效买家消息",
  "created_at": "2026-09-03T02:12:00.000Z",
  "updated_at": "2026-09-03T02:15:42.000Z",
  "error": null,
  "result": null
}
```

### 5.2 字段说明

| 字段 | 类型 | 说明 |
|---|---|---|
| `task_id` | string | 任务 ID。 |
| `status` | enum | `queued`、`running`、`completed`、`failed`。 |
| `stage` | string | 当前处理阶段。 |
| `progress` | number | 0 至 100 的整体进度估计。只用于展示，不作为结算依据。 |
| `completed_chunks` | integer/null | 已完成模型分块数。非分类阶段可能为 null。 |
| `total_chunks` | integer/null | 总模型分块数。 |
| `message` | string | 面向用户的简短状态。 |
| `created_at` | string | 创建时间。 |
| `updated_at` | string | 最后更新时间。 |
| `error` | object/null | 失败详情。 |
| `result` | object/null | 完成结果。 |

### 5.3 状态机

```text
queued
  -> running
       -> completed
       -> failed

failed --retry--> queued
```

服务重启时，原 `queued/running` 任务会被恢复为 `failed`，错误码为 `SERVICE_RESTARTED`。

### 5.4 常见 stage

| stage | 典型进度 | 含义 |
|---|---:|---|
| `queued` | 0 | 等待 Worker。 |
| `starting` | 1 | 任务开始。 |
| `saving_files` | 5 | 上传文件归档。 |
| `analyzing` | 10 | 解析输入并准备词条。 |
| `taxonomy` | 11-14 | 建立或继承词条体系。 |
| `classifying` | 15-85 | 分块逐条分析。 |
| `risk_review` | 86 | 风险候选二次证据复核。 |
| `summarizing` | 88 | 生成管理结论。 |
| `reporting` | 93 | 生成 Excel。 |
| `publishing` | 96 | 写入最终目录和清单。 |
| `completed` | 100 | 正式报告完成。 |

stage 是展示字段，客户端应允许未来出现新值，不要写成封闭枚举导致页面崩溃。

### 5.5 完成响应

```json
{
  "task_id": "5e810c18-07f0-49ef-8aab-39eaa945b92d",
  "status": "completed",
  "stage": "completed",
  "progress": 100,
  "completed_chunks": 40,
  "total_chunks": 40,
  "message": "Excel 报告已生成",
  "created_at": "2026-09-03T02:12:00.000Z",
  "updated_at": "2026-09-03T02:34:10.000Z",
  "error": null,
  "result": {
    "status": "completed",
    "report_level": "formal",
    "message": "AI分析已完成",
    "task_id": "5e810c18-07f0-49ef-8aab-39eaa945b92d",
    "summary": [
      "本期共识别……",
      "售前需求以……为主。"
    ],
    "files": {
      "excel": "E:\\...\\报告.xlsx",
      "markdown": "E:\\...\\管理结论.md",
      "manifest": "E:\\...\\result_manifest.json",
      "excel_download_url": "/api/download?path=...",
      "markdown_download_url": "/api/download?path=...",
      "manifest_download_url": "/api/download?path=..."
    }
  }
}
```

生产版必须移除 `excel/markdown/manifest` 的服务器绝对路径，只保留受权限保护的下载 URL 或 artifact ID。

### 5.6 失败响应

任务失败仍通过状态接口返回 HTTP `200`，失败由 `status=failed` 表示：

```json
{
  "task_id": "5e810c18-07f0-49ef-8aab-39eaa945b92d",
  "status": "failed",
  "stage": "classifying",
  "progress": 61,
  "completed_chunks": 22,
  "total_chunks": 40,
  "message": "本次分析没有完成",
  "created_at": "2026-09-03T02:12:00.000Z",
  "updated_at": "2026-09-03T02:25:00.000Z",
  "error": {
    "code": "UPSTREAM_TIMEOUT",
    "message": "模型分析服务响应超时，已完成的缓存会在重试时继续使用。",
    "retryable": true,
    "detail": "已脱敏的错误详情"
  },
  "result": null
}
```

### 5.7 找不到任务

HTTP `404 Not Found`

```json
{
  "status": "failed",
  "message": "找不到该分析任务。"
}
```

## 6. 错误码

| code | 含义 | retryable | 客户端处理 |
|---|---|---:|---|
| `UPSTREAM_TIMEOUT` | 模型服务读取超时 | true | 显示可重试；同进程重试可复用缓存。 |
| `UPSTREAM_RATE_LIMIT` | 上游 429 或限流 | true | 延迟后重试，避免并发重复提交。 |
| `ANALYSIS_FAILED` | 解析、基准、母版、模型 JSON 或报告生成失败 | 视错误而定 | 展示最后一条用户可读信息，详细错误交给管理员。 |
| `SERVICE_RESTARTED` | 服务重启中断了任务 | false | 当前版本要求重新上传并创建新任务。 |

当前错误归类较粗。生产版建议补充：

- `INVALID_PAYLOAD`
- `NO_CHAT_FILE`
- `UNSUPPORTED_FILE`
- `BASELINE_PRODUCT_MISMATCH`
- `BASELINE_TAXONOMY_MISSING`
- `MODEL_AUTH_FAILED`
- `MODEL_RESPONSE_INVALID`
- `TEMPLATE_MISSING`
- `TEMPLATE_INCOMPATIBLE`
- `STORAGE_WRITE_FAILED`
- `ANALYSIS_INCOMPLETE`

## 7. 重试任务

### `POST /api/tasks/{task_id}/retry`

只有以下条件同时满足才可重试：

- 任务存在。
- 当前状态为 `failed`。
- Node 进程仍持有原上传文件对象。

因此，服务重启后不能使用该接口重试历史失败任务。

### 7.1 成功响应

HTTP `202 Accepted`

响应为完整任务状态，核心字段示例：

```json
{
  "task_id": "5e810c18-07f0-49ef-8aab-39eaa945b92d",
  "status": "queued",
  "stage": "queued",
  "progress": 0,
  "message": "任务已重新排队",
  "error": null,
  "result": null
}
```

### 7.2 不可重试

HTTP `409 Conflict`

```json
{
  "status": "failed",
  "message": "该任务当前不能重试。"
}
```

### 7.3 客户端重试原则

1. 只有 `error.retryable=true` 时显示“重试”按钮。
2. 点击一次后立即禁用按钮，避免重复入队。
3. 收到 `409` 时提示用户重新上传。
4. 不能同时自动重试和用户手动重试。
5. 不要新建多个相同任务规避等待；这会重复扣费并抢占并发。

## 8. 下载文件

### `GET /api/download?path={absolute_path}`

`path` 必须进行 URL 编码，且规范化后位于《赤兔历史分析结果》根目录内。

### 8.1 请求示例

```text
GET /api/download?path=E%3A%5C%E6%96%B02%5C%E8%B5%A4%E5%85%94%E5%8E%86%E5%8F%B2%E5%88%86%E6%9E%90%E7%BB%93%E6%9E%9C%5C...
```

### 8.2 成功响应

HTTP `200 OK`，响应头包含：

```http
Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet
Content-Disposition: attachment; filename*=UTF-8''...
```

### 8.3 错误

- 缺少 `path`：HTTP `400`，文本 `Missing path`。
- 文件不存在、路径越界或传入根目录本身：HTTP `404`，文本 `File not found`。

### 8.4 生产替代接口

当前绝对路径下载接口只适用于本机联调。推荐改为：

```text
GET /api/tasks/{task_id}/artifacts/excel
GET /api/tasks/{task_id}/artifacts/markdown
GET /api/tasks/{task_id}/artifacts/manifest
```

服务端根据当前登录用户和任务归属查找真实路径，禁止客户端提交服务器路径。

## 9. 前端轮询参考实现

```javascript
const wait = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

async function waitForTask(taskId, onProgress) {
  let interval = 2000;

  while (true) {
    const response = await fetch(`/api/tasks/${encodeURIComponent(taskId)}`);
    const task = await response.json();

    if (!response.ok) {
      throw new Error(task.message || `HTTP ${response.status}`);
    }

    onProgress?.(task);

    if (task.status === "completed") return task.result;

    if (task.status === "failed") {
      const error = new Error(task.error?.message || task.message);
      error.code = task.error?.code;
      error.retryable = task.error?.retryable;
      throw error;
    }

    await wait(interval);
    interval = Math.min(5000, Math.round(interval * 1.1));
  }
}
```

页面刷新恢复：当前前端将活动任务 ID 保存在：

```text
localStorage["chitu_active_task"]
```

恢复时应先查询任务；完成或失败后清理该值。

## 10. result_manifest.json

该文件是每份正式报告的机器可读审计清单。

### 10.1 示例结构

```json
{
  "task_id": "5e810c18-07f0-49ef-8aab-39eaa945b92d",
  "product_name": "2449小气泡",
  "analysis_type": "baseline_compare",
  "analyzer": "scripts/generate_report.py",
  "analysis_engine": "llm_api",
  "report_level": "formal",
  "model": "gpt-5.6-sol",
  "llm_api_base": "https://example.invalid",
  "llm_api_key_configured": true,
  "llm_usage": {
    "api_calls": 42,
    "api_attempts": 45,
    "failed_attempts": 3,
    "cache_hits": 8,
    "degraded_messages": 0,
    "expected_messages": 2800,
    "analyzed_messages": 2800,
    "analysis_complete": true,
    "max_api_attempts": null,
    "model": "gpt-5.6-sol",
    "prompt_tokens": 120000,
    "completion_tokens": 18000,
    "total_tokens": 138000,
    "history_logic_source": "...",
    "history_logic_chars": 30000
  },
  "taxonomy": [],
  "products": ["2449小气泡"],
  "multi_product_sheets": [],
  "supplemental_sheets": ["口径说明", "有效买家消息"],
  "status": "completed",
  "input_files": [],
  "baseline_files": [],
  "chat_files": [],
  "reference_files": [],
  "output_files": {
    "excel": "...",
    "markdown": "...",
    "manifest": "..."
  },
  "summary": [],
  "created_at": "2026-09-03 10:34:10"
}
```

### 10.2 发布门禁

正式报告应同时满足：

```text
status == "completed"
report_level == "formal"
llm_usage.analysis_complete == true
llm_usage.expected_messages == llm_usage.analyzed_messages
llm_usage.degraded_messages == 0
```

任何一项不满足，都不能对外标注为正式报告。

## 11. CORS 和同源

当前 API 响应允许来源：

```text
http://127.0.0.1:{CHITU_PORT}
```

静态资源允许 `*`。推荐生产网页和 API 使用同一域名，由反向代理转发，这样无需开放跨域。若确需跨域，应使用来源白名单并补充认证凭证策略。

## 12. 当前接口缺陷和 V2 建议

### 12.1 必须修复

- 增加认证和任务所有权。
- 绝对路径改为 artifact ID。
- 服务器路径不再进入客户端 JSON。
- 创建任务前完成 payload、角色、扩展名和基准完整性校验。
- 错误响应统一为 JSON 和稳定错误码。
- 文件持久化后再返回 202，使服务重启后可恢复。
- 增加幂等键，防止重复提交和重复扣费。

### 12.2 推荐 V2 创建接口

请求头：

```http
Authorization: Bearer <user-token>
Idempotency-Key: <uuid>
```

成功响应：

```json
{
  "data": {
    "task_id": "...",
    "status": "queued",
    "links": {
      "self": "/api/v2/tasks/..."
    }
  },
  "request_id": "..."
}
```

错误响应：

```json
{
  "error": {
    "code": "BASELINE_PRODUCT_MISMATCH",
    "message": "本次上传 4 个聊天文件，但基准包含 3 个商品。",
    "retryable": false,
    "details": {
      "chat_product_count": 4,
      "baseline_product_count": 3
    }
  },
  "request_id": "..."
}
```

## 13. OpenAPI 3.1 摘要

下面的定义用于快速接入，不包含静态网页和生产认证扩展：

```yaml
openapi: 3.1.0
info:
  title: Chitu Customer Chat Analysis API
  version: 1.0.0
servers:
  - url: http://127.0.0.1:8787
paths:
  /api/health:
    get:
      summary: Health and queue status
      responses:
        '200':
          description: OK
  /api/tasks:
    post:
      summary: Create an asynchronous analysis task
      requestBody:
        required: true
        content:
          multipart/form-data:
            schema:
              type: object
              required: [payload, files]
              properties:
                payload:
                  type: string
                  description: JSON-encoded TaskPayload
                files:
                  type: array
                  items:
                    type: string
                    format: binary
      responses:
        '202':
          description: Task queued
        '413':
          description: Upload too large
  /api/tasks/{task_id}:
    get:
      summary: Get task state
      parameters:
        - in: path
          name: task_id
          required: true
          schema: { type: string }
      responses:
        '200':
          description: Task state
        '404':
          description: Task not found
  /api/tasks/{task_id}/retry:
    post:
      summary: Retry a failed in-memory task
      parameters:
        - in: path
          name: task_id
          required: true
          schema: { type: string }
      responses:
        '202':
          description: Requeued
        '409':
          description: Cannot retry
  /api/download:
    get:
      summary: Download a generated artifact (local V1 only)
      parameters:
        - in: query
          name: path
          required: true
          schema: { type: string }
      responses:
        '200':
          description: Binary artifact
        '404':
          description: Not found or outside allowed root
```

## 14. 接口联调验收清单

- 健康接口返回 200。
- 创建任务返回 202 和 UUID。
- payload 文件描述与 multipart 文件一一匹配。
- 多个聊天文件都使用相同字段名 `files` 重复上传。
- 状态轮询能展示 queued、running 和最终状态。
- 页面刷新后能通过 localStorage 恢复任务。
- 超时/限流失败时显示正确错误码和可重试状态。
- 完成任务的三个文件均可下载。
- 下载地址不能跨出历史结果目录。
- 未完成任务不返回正式 Excel。
- 上线版本已移除服务器绝对路径并增加权限验证。

