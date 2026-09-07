import json
import os
import hashlib
import http.client
import re
import socket
import time
from collections import Counter, defaultdict
from concurrent.futures import ThreadPoolExecutor, as_completed
from functools import lru_cache
from pathlib import Path
from threading import Lock
from urllib import error, request


class LLMAnalysisError(RuntimeError):
    def __init__(self, message, retryable=False):
        super().__init__(message)
        self.retryable = retryable


def emit_progress(stage, progress, message, **extra):
    payload = {"stage": stage, "progress": progress, "message": message, **extra}
    print("CHITU_PROGRESS " + json.dumps(payload, ensure_ascii=False), flush=True)


@lru_cache(maxsize=1)
def load_historical_logic():
    """Load the user-confirmed GPT analysis conversation as the logic source."""
    configured = os.environ.get("CHITU_HISTORY_LOG_FILE")
    if configured:
        path_candidate = Path(configured)
        if not path_candidate.is_absolute():
            path_candidate = Path(__file__).resolve().parents[1] / path_candidate
        candidates = [path_candidate]
    else:
        # 优先在项目内置 storage 目录查找（git 跟踪、随部署产物自带），
        # 兜底沿用旧约定：项目同级工作区根目录（历史习惯存放位置）。
        project_root = Path(__file__).resolve().parents[1]
        workspace_root = project_root.parent
        candidates = []
        seen = set()
        for search_root in (project_root / "storage", workspace_root):
            if not search_root.is_dir():
                continue
            for item in search_root.glob("long_text_*.txt"):
                if item in seen:
                    continue
                seen.add(item)
                candidates.append(item)
        candidates.sort(key=lambda item: item.stat().st_mtime, reverse=True)
    for candidate in candidates:
        if not candidate.exists():
            continue
        text = candidate.read_text(encoding="utf-8", errors="replace").strip()
        if text:
            return text[:30000], str(candidate)
    raise LLMAnalysisError(
        "未找到历史 GPT 分析记录。请配置 CHITU_HISTORY_LOG_FILE，网页不能在缺少已确认分析口径时冒充正式分析。"
    )


class OpenAICompatibleClient:
    def __init__(self, model=None):
        self.base_url = (os.environ.get("CHITU_LLM_API_BASE") or os.environ.get("OPENAI_BASE_URL") or "").rstrip("/")
        self.api_key = os.environ.get("CHITU_LLM_API_KEY") or os.environ.get("OPENAI_API_KEY") or ""
        self.model = model or os.environ.get("CHITU_ANALYSIS_MODEL") or "gpt-5.6-sol"
        # 推理模型默认档位思考 token 多、延迟高；分类/二审等结构化任务用低档即可，
        # 可通过 CHITU_LLM_REASONING_EFFORT 覆盖（none/low/medium/high/xhigh/max，留空则不传）。
        self.reasoning_effort = (os.environ.get("CHITU_LLM_REASONING_EFFORT") or "low").strip().lower() or None
        self.timeout = int(os.environ.get("CHITU_LLM_TIMEOUT", "75"))
        self.retries = max(1, int(os.environ.get("CHITU_LLM_RETRIES", "3")))
        self.calls = 0
        self.attempts = 0
        self.failed_attempts = 0
        self.cache_hits = 0
        self.degraded_messages = 0
        self.expected_messages = 0
        self.analyzed_messages = 0
        # 级联初分类（便宜模型）的去重/升级/用量统计，用于成本可视化。
        self.unique_messages = 0
        self.deduped_messages = 0
        self.escalated_messages = 0
        self.screener_model = ""
        self.screener_usage = Counter()
        configured_max_attempts = int(os.environ.get("CHITU_LLM_MAX_ATTEMPTS", "0"))
        self.max_attempts = configured_max_attempts if configured_max_attempts > 0 else None
        self.usage = Counter()
        self._usage_lock = Lock()
        default_cache = Path(__file__).resolve().parents[1] / "storage" / "cache" / "llm"
        self.cache_dir = Path(os.environ.get("CHITU_LLM_CACHE_DIR") or default_cache)
        self.cache_dir.mkdir(parents=True, exist_ok=True)
        if not self.base_url or not self.api_key:
            raise LLMAnalysisError("未配置 CHITU_LLM_API_BASE / CHITU_LLM_API_KEY，无法执行真实 AI 分析。")

    @property
    def endpoint(self):
        suffix = "/chat/completions" if self.base_url.endswith("/v1") else "/v1/chat/completions"
        return self.base_url + suffix

    def cache_path_for(self, system_prompt, user_prompt, max_tokens):
        payload = {
            "model": self.model,
            "messages": [
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt},
            ],
            "response_format": {"type": "json_object"},
            "max_tokens": max_tokens,
        }
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        # 缓存键不包含 reasoning_effort：同一提示词的历史缓存（含更高思考档位的结果）仍然有效。
        cache_key = hashlib.sha256(self.endpoint.encode("utf-8") + b"\0" + body).hexdigest()
        request_body = body
        if self.reasoning_effort:
            enriched = dict(payload)
            enriched["reasoning_effort"] = self.reasoning_effort
            request_body = json.dumps(enriched, ensure_ascii=False).encode("utf-8")
        return self.cache_dir / f"{cache_key}.json", payload, request_body

    def has_cached(self, system_prompt, user_prompt, max_tokens):
        cache_path, _, _ = self.cache_path_for(system_prompt, user_prompt, max_tokens)
        return cache_path.exists()

    def invalidate_cache(self, system_prompt, user_prompt, max_tokens):
        cache_path, _, _ = self.cache_path_for(system_prompt, user_prompt, max_tokens)
        cache_path.unlink(missing_ok=True)

    def complete_json(self, system_prompt, user_prompt, max_tokens=10000):
        cache_path, payload, body = self.cache_path_for(system_prompt, user_prompt, max_tokens)
        if cache_path.exists():
            try:
                cached = json.loads(cache_path.read_text(encoding="utf-8"))
                with self._usage_lock:
                    self.cache_hits += 1
                return cached
            except (OSError, json.JSONDecodeError):
                cache_path.unlink(missing_ok=True)
        last_error = None
        retryable = False
        for attempt in range(self.retries):
            req = request.Request(
                self.endpoint,
                data=body,
                method="POST",
                headers={
                    "Authorization": f"Bearer {self.api_key}",
                    "Content-Type": "application/json",
                    "Accept": "application/json",
                },
            )
            try:
                with self._usage_lock:
                    self.attempts += 1
                with request.urlopen(req, timeout=self.timeout) as response:
                    result = json.loads(response.read().decode("utf-8"))
                usage = result.get("usage") or {}
                with self._usage_lock:
                    self.calls += 1
                    for key in ("prompt_tokens", "completion_tokens", "total_tokens"):
                        self.usage[key] += int(usage.get(key) or 0)
                content = result["choices"][0]["message"]["content"]
                parsed = parse_json_content(content)
                if not isinstance(parsed, dict):
                    raise LLMAnalysisError("模型返回的 JSON 不是对象。", retryable=True)
                temporary = cache_path.with_suffix(f".{os.getpid()}.{time.time_ns()}.tmp")
                temporary.write_text(json.dumps(parsed, ensure_ascii=False), encoding="utf-8")
                os.replace(temporary, cache_path)
                return parsed
            except (
                error.HTTPError,
                error.URLError,
                TimeoutError,
                socket.timeout,
                OSError,
                http.client.HTTPException,
                KeyError,
                ValueError,
                json.JSONDecodeError,
            ) as exc:
                last_error = exc
                status = getattr(exc, "code", None)
                timeout_error = isinstance(exc, (TimeoutError, socket.timeout)) or (
                    isinstance(exc, error.URLError)
                    and isinstance(getattr(exc, "reason", None), (TimeoutError, socket.timeout))
                )
                connection_error = (
                    isinstance(exc, http.client.HTTPException)
                    or (isinstance(exc, OSError) and not isinstance(exc, error.HTTPError))
                )
                retryable = timeout_error or connection_error or status in (408, 409, 429, 500, 502, 503, 504)
                with self._usage_lock:
                    self.failed_attempts += 1
                if not retryable or attempt == self.retries - 1:
                    break
                time.sleep(1.5 * (attempt + 1))
        raise LLMAnalysisError(f"模型 API 调用失败：{last_error}", retryable=retryable)

    def usage_summary(self):
        return {
            "api_calls": self.calls,
            "api_attempts": self.attempts,
            "failed_attempts": self.failed_attempts,
            "cache_hits": self.cache_hits,
            "degraded_messages": self.degraded_messages,
            "expected_messages": self.expected_messages,
            "analyzed_messages": self.analyzed_messages,
            "analysis_complete": self.expected_messages == self.analyzed_messages,
            "max_api_attempts": self.max_attempts,
            "model": self.model,
            "prompt_tokens": self.usage["prompt_tokens"],
            "completion_tokens": self.usage["completion_tokens"],
            "total_tokens": self.usage["total_tokens"],
            "unique_messages": self.unique_messages,
            "deduped_messages": self.deduped_messages,
            "escalated_messages": self.escalated_messages,
            "screener_model": self.screener_model or None,
            "screener_api_calls": self.screener_usage["api_calls"],
            "screener_prompt_tokens": self.screener_usage["prompt_tokens"],
            "screener_completion_tokens": self.screener_usage["completion_tokens"],
        }


def ensure_attempt_budget(client):
    with client._usage_lock:
        exhausted = client.max_attempts is not None and client.attempts >= client.max_attempts
    if exhausted:
        raise LLMAnalysisError(
            "模型 API 尝试次数达到配置上限，完整性门禁已阻止生成不完整报告。",
            retryable=False,
        )


def required_retry_delay(retry_number):
    return min(60, 3 * (2 ** min(max(0, retry_number - 1), 4)))


def complete_json_required(client, system_prompt, user_prompt, max_tokens, operation):
    retry_number = 0
    while True:
        ensure_attempt_budget(client)
        try:
            return client.complete_json(system_prompt, user_prompt, max_tokens=max_tokens)
        except LLMAnalysisError as exc:
            if not exc.retryable:
                raise
            client.invalidate_cache(system_prompt, user_prompt, max_tokens)
            retry_number += 1
            delay = required_retry_delay(retry_number)
            emit_progress(
                "retrying",
                84,
                f"{operation}遇到接口超时，{delay} 秒后继续重试（已重试 {retry_number} 次）",
                retry_count=retry_number,
            )
            time.sleep(delay)


def validate_complete_batch(result, expected_ids, coverage_field="analyzed_ids"):
    if not isinstance(result, dict):
        raise LLMAnalysisError("模型返回的批次结果不是 JSON 对象。", retryable=True)
    expected = [str(item) for item in expected_ids]
    returned = [str(item) for item in (result.get(coverage_field) or [])]
    if len(returned) != len(set(returned)) or set(returned) != set(expected):
        missing = len(set(expected) - set(returned))
        extra = len(set(returned) - set(expected))
        raise LLMAnalysisError(
            f"模型未完整确认本批消息：缺少 {missing} 条，多出 {extra} 条。",
            retryable=True,
        )
    expected_set = set(expected)
    for field in ("demands", "risks", "reviews"):
        for row in result.get(field) or []:
            if isinstance(row, dict):
                row_id = str(row.get("id") or row.get("message_id") or "")
            elif isinstance(row, list) and row:
                row_id = str(row[0])
            else:
                continue
            if row_id and row_id not in expected_set:
                raise LLMAnalysisError(
                    f"模型在 {field} 中返回了不属于本批的消息 ID：{row_id}",
                    retryable=True,
                )
    return result


def validate_decision_coverage(result, expected_ids):
    if not isinstance(result, dict):
        raise LLMAnalysisError("模型返回的风险二审结果不是 JSON 对象。", retryable=True)
    returned = []
    for row in result.get("decisions") or []:
        if isinstance(row, dict):
            returned.append(str(row.get("id") or row.get("group_id") or ""))
        elif isinstance(row, list) and row:
            returned.append(str(row[0]))
    expected = [str(item) for item in expected_ids]
    if len(returned) != len(set(returned)) or set(returned) != set(expected):
        raise LLMAnalysisError("模型未对每个风险候选分组给出唯一决定。", retryable=True)
    return result


def parse_json_content(content):
    if isinstance(content, dict):
        return content
    text = str(content or "").strip()
    if text.startswith("```"):
        lines = text.splitlines()
        text = "\n".join(lines[1:-1]).strip()
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        start = text.find("{")
        end = text.rfind("}")
        if start >= 0 and end > start:
            return json.loads(text[start : end + 1])
        raise


def evenly_sample(messages, max_messages=900, max_chars=70000):
    if len(messages) <= max_messages:
        candidates = messages
    else:
        step = (len(messages) - 1) / (max_messages - 1)
        candidates = [messages[round(index * step)] for index in range(max_messages)]
    sample = []
    used = 0
    for item in candidates:
        compact = {
            "id": item["_llm_id"],
            "product": item["product"],
            "text": item["text"][:800],
        }
        size = len(json.dumps(compact, ensure_ascii=False))
        if sample and used + size > max_chars:
            break
        sample.append(compact)
        used += size
    return sample


def configured_chunk_limits():
    """Batch caps are env-tunable: bigger batches amortize the per-call fixed prompt."""
    try:
        max_messages = max(20, int(os.environ.get("CHITU_CHUNK_MESSAGES", "110")))
    except ValueError:
        max_messages = 110
    try:
        max_chars = max(4000, int(os.environ.get("CHITU_CHUNK_CHARS", "24000")))
    except ValueError:
        max_chars = 24000
    return max_messages, max_chars


def build_chunks(messages, max_messages=None, max_chars=None):
    if max_messages is None or max_chars is None:
        env_messages, env_chars = configured_chunk_limits()
        max_messages = max_messages or env_messages
        max_chars = max_chars or env_chars
    chunks = []
    current = []
    used = 0
    for item in messages:
        compact = {
            "id": item["_llm_id"],
            "product": item["product"],
            "buyer": item["sender"],
            "conversation_id": item.get("conversation_id", 0),
            "date": item.get("date", ""),
            "time": item.get("time", ""),
            "text": re.sub(r"\s+", " ", item["text"]).strip()[:1200],
            "context": item.get("context", "")[:1800],
        }
        size = len(json.dumps(compact, ensure_ascii=False))
        if current and (len(current) >= max_messages or used + size > max_chars):
            chunks.append(current)
            current = []
            used = 0
        current.append(compact)
        used += size
    if current:
        chunks.append(current)
    return chunks


def discover_taxonomy(client, messages, task_context="", baseline_context="", historical_logic=""):
    # A tiny sample misses long-tail售前/售后 needs and permanently caps coverage,
    # because the classifier may only use labels discovered here.
    system_prompt = """你是电商客服聊天需求分析师。严格执行用户确认的历史 GPT 分析协议，根据本次真实买家消息建立本产品 V1 词条。
先理解产品，再建词条；禁止套用其他商品词条或复制历史报告数字、Excel旧内容。
售前是购买决策咨询；售后是已购后的操作、教程、安装、退款、实际效果、故障和服务处理。“怎么用/安装教程/怎么操作”默认售后。
一句话允许拆多个独立需求。词条既不能粗到失去动作价值，也不能把同一业务问题无限拆细。
只返回合法 JSON，不要解释。"""
    history_excerpt = historical_logic[:12000]
    def build_taxonomy_prompt(sample):
        return f"""任务说明：{task_context or '分析本次客服聊天'}
历史基准说明（只参考词条连续性，不复制统计）：{baseline_context[:3000] or '无'}

历史 GPT 分析规则片段（只提炼规则，不把示例当成本期数据）：
<historical_analysis_logic>
{history_excerpt}
</historical_analysis_logic>

请完整扫描本次聊天样本，建立足以覆盖高频和重要长尾真实需求的 V1 需求词条，只返回以下 JSON：
{{
  "labels": [
    {{"label":"简短中文词条","stage":"售前或售后","theme":"一级主题","definition":"严格定义","priority":"P0/P1/P2","action":"建议动作"}}
  ]
}}
要求：
1. 词条互斥、稳定、可复用；严格定义必须说明计入和不计入边界。
2. 必须分别覆盖售前参数/规格/材质/配置/价格/政策/物流，以及售后教程/安装/实际故障/退换货/缺件破损/履约服务；没有证据的类别不要硬建。
3. 相似但经营动作不同的问题应拆开，例如“充电参数”与“实际充不上电”、“安装教程”与“卡扣装不上”。
4. 同一详情页参数表或同一客服动作即可解决的细项必须合并，禁止把接口、容量、充满时长、续航等同类供电参数无限拆词；退款、退货、换货、回寄、取件也不要按流程节点碎片化，除非经营动作确实不同。
5. 售前“下单前预计发货/到货时效”必须与售后“已下单后的物流查询、未收到、签收异常、订单备注”分开，禁止把已购物流问题放进售前榜。
6. 单产品通常 12-24 个词条，多 SKU 必须保留各 SKU 的专属词条，不要用少量粗词条强行覆盖全部产品。

本次聊天样本：
{json.dumps(sample, ensure_ascii=False)}"""

    response = None
    selected_prompt = None
    selected_tokens = None
    sample_specs = ((90, 12000, 2500), (60, 9000, 2400), (40, 6000, 2200))
    for max_messages, max_chars, max_tokens in sample_specs:
        sample = evenly_sample(messages, max_messages=max_messages, max_chars=max_chars)
        prompt = build_taxonomy_prompt(sample)
        try:
            response = client.complete_json(
                system_prompt,
                prompt,
                max_tokens=max_tokens,
            )
            selected_prompt = prompt
            selected_tokens = max_tokens
            break
        except LLMAnalysisError as exc:
            if not exc.retryable:
                raise
    if response is None:
        sample = evenly_sample(messages, max_messages=40, max_chars=6000)
        selected_prompt = build_taxonomy_prompt(sample)
        selected_tokens = 2200
        response = complete_json_required(
            client,
            system_prompt,
            selected_prompt,
            selected_tokens,
            "V1 词条建立",
        )

    while True:
        labels = []
        seen = set()
        for item in response.get("labels") or []:
            label = str(item.get("label") or "").strip()[:40]
            stage = str(item.get("stage") or "").strip()
            if not label or stage not in ("售前", "售后") or label in seen:
                continue
            seen.add(label)
            priority = str(item.get("priority") or "P2").upper()
            labels.append(
                {
                    "label": label,
                    "stage": stage,
                    "stages": [stage],
                    "product_prefixes": [],
                    "theme": str(item.get("theme") or "其他").strip()[:40],
                    "definition": str(item.get("definition") or f"{label}相关的明确买家需求。").strip(),
                    "priority": priority if priority in ("P0", "P1", "P2") else "P2",
                    "action": str(item.get("action") or "持续观察并完善对应客服承接。").strip(),
                }
            )
        if len(labels) >= 6:
            break
        client.invalidate_cache(system_prompt, selected_prompt, selected_tokens)
        emit_progress("retrying", 12, "V1 词条返回不完整，正在重新建立")
        time.sleep(3)
        response = complete_json_required(
            client,
            system_prompt,
            selected_prompt,
            selected_tokens,
            "V1 词条建立",
        )
    return labels, [str(item) for item in (response.get("analysis_notes") or [])[:8]]


def discover_taxonomy_additions(client, messages, reference_taxonomy, task_context="", historical_logic=""):
    sample = evenly_sample(messages, max_messages=120, max_chars=18000)
    system_prompt = """你是客服需求 V1 词条治理员。本次已有正式上期词条，必须优先沿用旧词条；只识别任何旧词条都无法准确覆盖、且需要不同经营动作的真实新需求。禁止给旧词条改名、拆同义词或把低信息闲聊建成新词条。只返回合法 JSON。"""
    user_prompt = f"""任务：{task_context or '客服聊天上期基准对比'}

上期正式 V1 词条：
{json.dumps(reference_taxonomy, ensure_ascii=False)}

历史 GPT 规则片段：
{historical_logic[:6000]}

本期聊天样本：
{json.dumps(sample, ensure_ascii=False)}

只返回确实需要新增的词条；没有新增时 labels 返回空数组：
{{"labels":[{{"label":"新词条","stage":"售前或售后","stages":["售前"],"product_prefixes":[],"theme":"一级主题","definition":"计入与不计入边界","priority":"P0/P1/P2","action":"具体动作","reason":"为何旧词条无法覆盖"}}]}}"""
    response = complete_json_required(client, system_prompt, user_prompt, 1800, "新增词条检查")
    existing = {item["label"] for item in reference_taxonomy}
    additions = []
    for item in response.get("labels") or []:
        label = str(item.get("label") or "").strip()[:40]
        stage = str(item.get("stage") or "").strip()
        stages = [value for value in (item.get("stages") or [stage]) if value in ("售前", "售后")]
        if not label or label in existing or not stages:
            continue
        prefixes = item.get("product_prefixes") or []
        if isinstance(prefixes, str):
            prefixes = [prefixes]
        priority = str(item.get("priority") or "P2").upper()
        additions.append(
            {
                "label": label,
                "stage": stages[0],
                "stages": list(dict.fromkeys(stages)),
                "product_prefixes": [str(value).strip() for value in prefixes if str(value).strip()],
                "theme": str(item.get("theme") or "其他").strip()[:40],
                "definition": str(item.get("definition") or f"{label}相关的明确买家需求。").strip(),
                "priority": priority if priority in ("P0", "P1", "P2") else "P2",
                "action": str(item.get("action") or "为该新增需求建立独立承接方案。").strip(),
            }
        )
        existing.add(label)
    return additions


CLASSIFICATION_SYSTEM_PROMPT = """你是客服聊天逐条语义分析器。严格执行用户确认的历史 GPT 分析协议，必须依据买家当前原话和同会话上下文分类，不做简单关键词匹配。
只使用给定 V1 词条；一句话可拆成多个彼此独立的需求，也可以没有需求。不得因为句子短就漏掉当前原话中已经明确的诉求。
正式需求门禁：当前买家原话必须包含可独立统计的购买决策、明确提问、具体请求、操作困难、本人实际异常或售后处理诉求。纯寒暄、确认/否定、答复客服、状态同步、情绪表达、残缺片段、仅重复客服话术或只在相邻上下文出现但当前原话没有提出的主题，不输出 demands。
上下文只用于恢复代词、省略对象、已购/未购状态和连续问答，禁止把上下文中的其他问题复制成当前消息的新需求。一个意图优先使用最具体的一个词条，禁止同一意图同时命中宽泛父词条和具体子词条；只有当前原话明确包含两个不同问题或请求时才拆成多个需求。
售前是购买决策咨询；售后是已有明确购买、下单、收货、拆封、安装、实操、实际效果、故障、退款或服务处理证据的问题。“怎么用/教程/换头”等独立操作教程需求默认按售后承接压力判断。
功能、配置、部件、耗材、适用性、效果原理等同时允许售前/售后的词条，如果没有明确已购、收货、本人实操或实际异常证据，必须归售前；不能仅凭“怎么、如何、需要、放哪里、有什么作用”判成售后。询问吸头区别/用途、加什么水/是否加水、导出液或收缩液怎么搭配、是否需要热敷，默认都是售前购买与使用预期咨询。
物流阶段必须看上下文：下单前问预计发货/到货属于售前；已下单后的订单号、快递备注、没收到、签收和取件处理属于售后。
阶段必须从词条 allowed_stages 中选择；单阶段词条必须使用唯一阶段。双阶段词条先依据明确的购买/收货/实操证据判断，没有售后证据时使用售前；只有 allowed_stages 不含售前时才使用 default_stage。有产品限制的词条只能用于对应 SKU。
三款卷发产品必须遵守以下口径：
1. 856“是否需要加水/加什么水/可否不加水/水能用多久”归“加水/水质”售前；明确询问水箱如何取装、从哪里加水、怎样加水才归“水箱拆装/加水方法”售后。
2. 856询问冷雾作用、原理、何时出雾、正常雾量归“冷雾功能/原理”，没有实操证据时使用售前；本人实际不出雾或雾小归“出雾故障/雾小”售后。
3. 五合一询问有哪些头、各头用途、配置是否齐全归“五合一配置/替换头用途”售前；明确询问如何拆卸安装归“换头/拆装”售后。请求具体操作或视频时，可同时命中“使用方法/教程”，不要用具体操作词条替代教程需求。
4. 一条消息同时包含购买决策、操作教程和实际故障时应拆成多个需求；不要因为已命中更具体的词条就漏掉独立的教程需求。
risks 在这一轮只产生候选，随后还会做独立证据二审。购买前担忧、客服话术、教程咨询不得作为风险候选。
只返回合法 JSON，禁止解释。"""


AFTERSALE_EVIDENCE_RE = re.compile(
    r"已(?:经)?(?:买|下单|收到|收货|用|试|安装|拆封|充)|"
    r"买(?:了|过|回来)|下单(?:了|后)|收到|收货|到货|拿到|"
    r"刚(?:用|收到|拆|装|充)|用(?:了|过|完|后|的?时候)|使用时|试(?:了|过|用)|"
    r"打开(?:了|过)|拆封|安装(?:了|不上|不了)|充(?:了|过|一晚)|"
    r"我(?:这个|这台|的这台|的机器|的仪器|的订单|的快递)|"
    r"之前.*(?:正常|可以|能用)|现在.*(?:不|没)|"
    r"不出|不吸|不热|没热|不加热|漏水|坏了|破损|没反应|里面(?:还|居然)?有水|"
    r"退货|退款|换货|补发|少发|错发|未收到|签收|订单号|快递单|"
    r"拦截|改(?:一下)?地址|修改(?:一下)?地址|催发货|查物流"
)


def buyer_context_text(message):
    """Keep only buyer-authored context when applying deterministic stage gates."""
    parts = [str(message.get("text") or "")]
    for line in str(message.get("context") or "").splitlines():
        if line.startswith(("当前买家:", "同买家:")):
            parts.append(line.split(":", 1)[1])
    return "\n".join(parts)


def calibrate_demand_stages(messages, taxonomy, demand_map):
    """Do not let ambiguous dual-stage product questions drift into aftersale."""
    messages_by_id = {item["_llm_id"]: item for item in messages}
    taxonomy_by_label = {item["label"]: item for item in taxonomy}
    calibrated = defaultdict(list)
    for msg_id, demands in demand_map.items():
        message = messages_by_id.get(msg_id, {})
        has_aftersale_evidence = bool(AFTERSALE_EVIDENCE_RE.search(buyer_context_text(message)))
        seen = set()
        for demand in demands:
            item = dict(demand)
            meta = taxonomy_by_label.get(item.get("label"), {})
            allowed_stages = list(dict.fromkeys(meta.get("stages") or [meta.get("stage")]))
            if (
                item.get("stage") == "售后"
                and "售前" in allowed_stages
                and "售后" in allowed_stages
                and not has_aftersale_evidence
            ):
                item["stage"] = "售前"
            key = (item.get("stage"), item.get("label"))
            if key not in seen:
                calibrated[msg_id].append(item)
                seen.add(key)
    return calibrated


DEMAND_SIGNAL_RE = re.compile(
    r"吗|呢|？|\?|怎么|如何|为什么|啥|哪个|哪些|多少|能不能|可以不|是不是|有没有|"
    r"不出|不热|不亮|不灵|不吸|不好用|没用|坏|漏|退|换|伤|过敏|二手|区别"
)


# 初筛提示词：宽进严出——mini 拿不准时倾向标记需求（靠置信度和信号词升级 sol 严判），
# 只有确定为纯寒暄/确认/系统话术才输出无需求。置信度要求真实自评，低置信由主模型复核。
SCREENER_SYSTEM_PROMPT = CLASSIFICATION_SYSTEM_PROMPT + (
    "\n初筛模式补充：你是第一遍粗筛，拿不准一条消息是否包含可统计需求时，"
    "倾向于输出该需求，置信度按你的真实判断给出（确定无误 0.95+，有明确依据 0.85 左右，"
    "倾向但不确定 0.65 左右，仅疑似 0.4 以下）；"
    "只有确定为纯寒暄、确认答复、情绪表达或系统话术时才不输出 demands。"
    "你的输出随后会被主模型复核，漏报比多报更严重。"
)


def build_screener_client(model):
    """Factory indirection so tests can stub the cascade screener client."""
    return OpenAICompatibleClient(model=model)


def dedupe_messages_for_classification(messages):
    """Group identical (product, text) messages so the model sees one representative.

    Only the label decision is shared: stage calibration and risk evidence still
    run per message on its own context after the fan-out.
    """
    representatives = []
    members = {}
    seen = {}
    for message in messages:
        key = (message["product"], message["text"])
        rep_id = seen.get(key)
        if rep_id is None:
            representatives.append(message)
            rep_id = message["_llm_id"]
            seen[key] = rep_id
            members[rep_id] = [rep_id]
        else:
            members[rep_id].append(message["_llm_id"])
    return representatives, members


def message_cache_enabled():
    return os.environ.get("CHITU_MESSAGE_CACHE", "1").strip().lower() not in ("0", "false", "off")


def message_memo_path(classifier, system_prompt, taxonomy_for_prompt, message):
    payload = json.dumps(
        {
            "endpoint": getattr(classifier, "endpoint", ""),
            "model": classifier.model,
            "classifier_contract": system_prompt,
            "taxonomy": taxonomy_for_prompt,
            "product": message["product"],
            "text": message["text"],
        },
        ensure_ascii=False,
        sort_keys=True,
    )
    digest = hashlib.sha256(payload.encode("utf-8")).hexdigest()
    directory = classifier.cache_dir / "message_cache" / digest[:2]
    return digest, directory / f"{digest}.json"


def load_message_memo(classifier, system_prompt, taxonomy_for_prompt, message):
    if not message_cache_enabled():
        return None
    _, target = message_memo_path(classifier, system_prompt, taxonomy_for_prompt, message)
    if not target.exists():
        return None
    try:
        parsed = json.loads(target.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return None
    if not isinstance(parsed, dict):
        return None
    for key in ("demands", "risks", "reviews"):
        if not isinstance(parsed.get(key), list):
            return None
    return parsed


def save_message_memo(classifier, system_prompt, taxonomy_for_prompt, message, parsed):
    if not message_cache_enabled() or not isinstance(parsed, dict):
        return
    _, target = message_memo_path(classifier, system_prompt, taxonomy_for_prompt, message)
    try:
        target.parent.mkdir(parents=True, exist_ok=True)
        temporary = target.with_suffix(f".{os.getpid()}.{time.time_ns()}.tmp")
        temporary.write_text(json.dumps(parsed, ensure_ascii=False), encoding="utf-8")
        os.replace(temporary, target)
    except OSError:
        pass


def parse_chunk_rows(result):
    """Split a raw chunk result into per-message rows; ids with no rows still count."""
    per_message = {}

    def bucket(msg_id):
        return per_message.setdefault(str(msg_id), {"demands": [], "risks": [], "reviews": []})

    for msg_id in result.get("analyzed_ids") or []:
        bucket(str(msg_id))
    for row in result.get("demands") or []:
        if isinstance(row, dict):
            msg_id = str(row.get("id") or row.get("message_id") or "")
            if not msg_id:
                continue
            label = str(row.get("label") or "")
            stage = str(row.get("stage") or "")
            confidence = row.get("confidence", 0.9)
        elif isinstance(row, list) and len(row) >= 2:
            msg_id, label = str(row[0]), str(row[1])
            if len(row) >= 4:
                stage = str(row[2])
                confidence = row[3]
            elif len(row) == 3 and str(row[2]) in ("售前", "售后"):
                stage = str(row[2])
                confidence = 0.9
            else:
                stage = ""
                confidence = row[2] if len(row) > 2 else 0.9
        else:
            continue
        if label:
            bucket(msg_id)["demands"].append({"label": label, "stage": stage, "confidence": confidence})
    for row in result.get("risks") or []:
        if isinstance(row, dict):
            msg_id = str(row.get("id") or row.get("message_id") or "")
            if not msg_id:
                continue
            risk_type = str(row.get("risk_type") or row.get("type") or "安全/质量异常")
            reason = str(row.get("reason") or "模型识别为实际发生风险")
            priority = str(row.get("priority") or "P0")
        elif isinstance(row, list) and len(row) >= 2:
            msg_id, risk_type = str(row[0]), str(row[1])
            reason = str(row[2]) if len(row) > 2 else "模型识别为实际发生风险"
            priority = str(row[3]) if len(row) > 3 else "P0"
        else:
            continue
        bucket(msg_id)["risks"].append({"risk_type": risk_type[:50], "reason": reason, "priority": priority if priority in ("P0", "P1") else "P0"})
    for row in result.get("reviews") or []:
        if isinstance(row, dict):
            msg_id = str(row.get("id") or row.get("message_id") or "")
            if not msg_id:
                continue
            reason = str(row.get("reason") or "语义不完整")
        elif isinstance(row, list) and len(row) >= 2:
            msg_id, reason = str(row[0]), str(row[1])
        else:
            continue
        bucket(msg_id)["reviews"].append(reason)
    return per_message


def apply_taxonomy_gate(per_message, messages_by_id, taxonomy_by_label):
    """Validate parsed rows against the taxonomy; mirrors the historical fan-in rules."""
    demand_map = defaultdict(list)
    risks = defaultdict(list)
    reviews = defaultdict(list)
    for msg_id, parsed in per_message.items():
        message = messages_by_id.get(msg_id)
        for demand in parsed.get("demands") or []:
            label = str(demand.get("label") or "")
            stage = str(demand.get("stage") or "")
            meta = taxonomy_by_label.get(label)
            allowed_stages = (meta.get("stages") or [meta.get("stage", "售后")]) if meta else []
            product_prefixes = (meta.get("product_prefixes") or []) if meta else []
            product_allowed = not product_prefixes or (
                message is not None and any(prefix in str(message["product"]) for prefix in product_prefixes)
            )
            if meta and msg_id and message and product_allowed:
                if stage not in allowed_stages:
                    stage = allowed_stages[0] if len(allowed_stages) == 1 else ""
                if stage not in ("售前", "售后"):
                    reviews[msg_id].append(f"词条“{label}”的售前/售后阶段无法稳定判断")
                    continue
                try:
                    confidence = max(0.0, min(1.0, float(demand.get("confidence", 0.9))))
                except (TypeError, ValueError):
                    confidence = 0.9
                if (stage, label) not in {(item["stage"], item["label"]) for item in demand_map[msg_id]}:
                    demand_map[msg_id].append({"label": label, "stage": stage, "confidence": confidence})
        for risk in parsed.get("risks") or []:
            risks[msg_id].append(
                {
                    "risk_type": str(risk.get("risk_type") or "")[:50],
                    "reason": str(risk.get("reason") or ""),
                    "priority": risk.get("priority") if risk.get("priority") in ("P0", "P1") else "P0",
                }
            )
        for reason in parsed.get("reviews") or []:
            reviews[msg_id].append(str(reason))
    return demand_map, risks, reviews


def needs_screener_review(parsed, message, confidence_threshold):
    """Deterministic triage: which screener results must the main model re-check."""
    if parsed.get("risks") or parsed.get("reviews"):
        return True
    for demand in parsed.get("demands") or []:
        try:
            confidence = float(demand.get("confidence", 0.9))
        except (TypeError, ValueError):
            confidence = 0.9
        if confidence < confidence_threshold:
            return True
    if not parsed.get("demands") and DEMAND_SIGNAL_RE.search(message["text"]):
        return True
    return False


def classify_chunks(client, messages, taxonomy):
    taxonomy_for_prompt = [
        {
            "label": item["label"],
            "default_stage": item.get("stage", "售后"),
            "allowed_stages": item.get("stages") or [item.get("stage", "售后")],
            "product_prefixes": item.get("product_prefixes") or [],
            "theme": item.get("theme", "其他"),
            "definition": item["definition"],
        }
        for item in taxonomy
    ]
    taxonomy_by_label = {item["label"]: item for item in taxonomy}
    messages_by_id = {item["_llm_id"]: item for item in messages}
    system_prompt = CLASSIFICATION_SYSTEM_PROMPT
    classification_cache = client.cache_dir / "classification"
    classification_cache.mkdir(parents=True, exist_ok=True)

    screener_model = (os.environ.get("CHITU_CLASSIFY_MODEL") or "").strip()
    cascade_enabled = bool(screener_model) and screener_model != client.model
    try:
        confidence_threshold = float(os.environ.get("CHITU_CASCADE_CONFIDENCE", "0.7"))
    except ValueError:
        confidence_threshold = 0.7

    def taxonomy_for_chunk(chunk):
        products = {str(item.get("product") or "") for item in chunk}
        return [
            item
            for item in taxonomy_for_prompt
            if not item.get("product_prefixes")
            or any(
                prefix in product
                for product in products
                for prefix in item.get("product_prefixes") or []
            )
        ]

    def contract_for(classifier):
        return SCREENER_SYSTEM_PROMPT if cascade_enabled and classifier is screener else system_prompt

    def build_user_prompt(chunk, contract=None):
        return f"""V1词条：
{json.dumps(taxonomy_for_chunk(chunk), ensure_ascii=False)}

逐条分析下面消息，返回紧凑 JSON：
{{
  "analyzed_ids": ["本批每一条消息id，必须全部返回且不重复"],
  "demands": [["消息id","V1词条","售前或售后",0.0到1.0置信度]],
  "risks": [["消息id","风险类型","为什么属于本人实际发生","P0或P1"]],
  "reviews": [["消息id","无法稳定分类的原因"]]
}}
规则：
1. demands 只能使用给定 V1 词条的精确名称；当前原话没有可独立统计的明确需求时不要输出。
2. 同一消息命中同一词条最多一次。
3. 只有语义确实不完整但可能是需求时才进入 reviews，不要把所有闲聊放入复核。
4. 请完整重扫本批消息，避免漏掉一条消息里的第二个独立需求。
5. text 是需求证据主体；context 只用于消歧和判断购买阶段，禁止从 context 复制当前 text 没有提出的其他需求。
6. analyzed_ids 必须逐一列出本批全部消息 id；即使某条没有需求，也必须列入。
7. 同一意图只保留最具体的一个词条；只有当前 text 明确包含多个独立意图时才允许多词条。

消息：
{json.dumps(chunk, ensure_ascii=False)}"""

    def result_cache_path_for(classifier, chunk):
        digest = hashlib.sha256(
            json.dumps(
                {
                    "model": classifier.model,
                    "classifier_contract": contract_for(classifier),
                    "taxonomy": taxonomy_for_prompt,
                    "messages": chunk,
                },
                ensure_ascii=False,
                sort_keys=True,
            ).encode("utf-8")
        ).hexdigest()
        return classification_cache / f"{digest}.json"

    def save_result(classifier, chunk, result):
        validate_complete_batch(result, [item["id"] for item in chunk])
        target = result_cache_path_for(classifier, chunk)
        temporary = target.with_suffix(f".{os.getpid()}.{time.time_ns()}.tmp")
        temporary.write_text(json.dumps(result, ensure_ascii=False), encoding="utf-8")
        os.replace(temporary, target)
        return result

    min_retry_chunk_size = 8

    def combine_split(classifier, chunk, split_depth):
        midpoint = len(chunk) // 2
        left = analyze_chunk_with(classifier, chunk[:midpoint], split_depth + 1)
        right = analyze_chunk_with(classifier, chunk[midpoint:], split_depth + 1)
        return save_result(
            classifier,
            chunk,
            {
                "demands": (left.get("demands") or []) + (right.get("demands") or []),
                "risks": (left.get("risks") or []) + (right.get("risks") or []),
                "reviews": (left.get("reviews") or []) + (right.get("reviews") or []),
                "analyzed_ids": (left.get("analyzed_ids") or []) + (right.get("analyzed_ids") or []),
            },
        )

    def analyze_chunk_with(classifier, chunk, split_depth=0):
        contract = contract_for(classifier)
        cached_result = result_cache_path_for(classifier, chunk)
        if cached_result.exists():
            try:
                result = json.loads(cached_result.read_text(encoding="utf-8"))
                validate_complete_batch(result, [item["id"] for item in chunk])
                with client._usage_lock:
                    client.cache_hits += 1
                return result
            except (OSError, json.JSONDecodeError, LLMAnalysisError):
                cached_result.unlink(missing_ok=True)

        user_prompt = build_user_prompt(chunk)
        token_limit = min(4000, max(900, len(chunk) * 40))
        if len(chunk) > min_retry_chunk_size:
            midpoint = len(chunk) // 2
            left_chunk = chunk[:midpoint]
            right_chunk = chunk[midpoint:]
            left_limit = min(4000, max(900, len(left_chunk) * 40))
            right_limit = min(4000, max(900, len(right_chunk) * 40))
            if classifier.has_cached(contract, build_user_prompt(left_chunk), left_limit) or classifier.has_cached(
                contract, build_user_prompt(right_chunk), right_limit
            ):
                return combine_split(classifier, chunk, split_depth)
        try:
            ensure_attempt_budget(classifier)
            result = classifier.complete_json(contract, user_prompt, max_tokens=token_limit)
            return save_result(classifier, chunk, result)
        except LLMAnalysisError as exc:
            if not exc.retryable:
                raise
            classifier.invalidate_cache(contract, user_prompt, token_limit)
            if len(chunk) > min_retry_chunk_size:
                return combine_split(classifier, chunk, split_depth)
            retry_number = 0
            while True:
                retry_number += 1
                delay = required_retry_delay(retry_number)
                emit_progress(
                    "retrying",
                    84,
                    f"小批消息分析超时，{delay} 秒后继续重试（本批 {len(chunk)} 条）",
                    retry_count=retry_number,
                )
                time.sleep(delay)
                ensure_attempt_budget(classifier)
                try:
                    result = classifier.complete_json(contract, user_prompt, max_tokens=token_limit)
                    return save_result(classifier, chunk, result)
                except LLMAnalysisError as retry_exc:
                    if not retry_exc.retryable:
                        raise
                    classifier.invalidate_cache(contract, user_prompt, token_limit)

    def run_chunks(classifier, chunks, label, base_progress, span):
        per_message = {}
        worker_limit = max(1, int(os.environ.get("CHITU_LLM_WORKERS", "8")))
        worker_count = min(worker_limit, len(chunks))
        with ThreadPoolExecutor(max_workers=worker_count) as executor:
            future_chunks = {executor.submit(analyze_chunk_with, classifier, chunk): chunk for chunk in chunks}
            for completed, future in enumerate(as_completed(future_chunks), 1):
                response = future.result()
                for msg_id, parsed in parse_chunk_rows(response).items():
                    per_message[msg_id] = parsed
                emit_progress(
                    "classifying",
                    base_progress + int(span * completed / max(1, len(chunks))),
                    f"{label}第 {completed}/{len(chunks)} 批聊天",
                    completed_chunks=completed,
                    total_chunks=len(chunks),
                )
        return per_message

    representatives, members = dedupe_messages_for_classification(messages)
    deduped_messages = len(messages) - len(representatives)
    per_message = {}
    result_source = {}
    pending = []
    # 级联开启时首查初分类模型的备忘录：升级筛选必须基于 pass1 结果才能确定性重放。
    screener = build_screener_client(screener_model) if cascade_enabled else None
    first_lookup = screener if cascade_enabled else client
    for rep in representatives:
        cached = load_message_memo(first_lookup, contract_for(first_lookup), taxonomy_for_prompt, rep)
        if cached is not None:
            per_message[rep["_llm_id"]] = cached
            result_source[rep["_llm_id"]] = first_lookup
            with client._usage_lock:
                client.cache_hits += 1
        else:
            pending.append(rep)

    escalated_ids = set()
    if cascade_enabled and pending:
        # 初分类用便宜模型，思考档位独立可控（默认 none：探针实测质量足够且最快最省）；
        # mini 在大批次下 JSON 完整性差，初分类固定用较小批次。
        screener_effort = (os.environ.get("CHITU_SCREENER_REASONING_EFFORT") or "none").strip().lower() or None
        screener.reasoning_effort = screener_effort
        try:
            screener_chunk_cap = max(20, int(os.environ.get("CHITU_SCREENER_CHUNK_MESSAGES", "60")))
        except ValueError:
            screener_chunk_cap = 60
        _, screener_chars_cap = configured_chunk_limits()
        chunks = build_chunks(pending, max_messages=screener_chunk_cap, max_chars=screener_chars_cap)
        emit_progress(
            "classifying",
            18,
            f"小模型初分类 {len(pending)} 条唯一消息，共 {len(chunks)} 批",
            completed_chunks=0,
            total_chunks=len(chunks),
        )
        pass_one = run_chunks(screener, chunks, "初分类", 18, 40)
        with client._usage_lock:
            client.screener_model = screener.model
            client.screener_usage["api_calls"] += screener.calls
            client.screener_usage["prompt_tokens"] += screener.usage["prompt_tokens"]
            client.screener_usage["completion_tokens"] += screener.usage["completion_tokens"]
        for rep in pending:
            msg_id = rep["_llm_id"]
            parsed = pass_one.get(msg_id)
            if parsed is None:
                # 初分类漏报该消息 id：宁可升级大模型，也不能丢消息。
                escalated_ids.add(msg_id)
                continue
            save_message_memo(screener, contract_for(screener), taxonomy_for_prompt, rep, parsed)
            if needs_screener_review(parsed, rep, confidence_threshold):
                escalated_ids.add(msg_id)
            else:
                per_message[msg_id] = parsed
                result_source[msg_id] = screener
        escalated = [rep for rep in pending if rep["_llm_id"] in escalated_ids]
        with client._usage_lock:
            client.escalated_messages = len(escalated)
        emit_progress(
            "classifying",
            60,
            f"初分类完成，升级大模型复核 {len(escalated)}/{len(pending)} 条（{len(escalated) / max(1, len(pending)):.0%}）",
        )
    else:
        escalated = pending
        escalated_ids = {rep["_llm_id"] for rep in pending}

    sol_pending = []
    for rep in escalated:
        cached = load_message_memo(client, contract_for(client), taxonomy_for_prompt, rep)
        if cached is not None:
            per_message[rep["_llm_id"]] = cached
            result_source[rep["_llm_id"]] = client
            with client._usage_lock:
                client.cache_hits += 1
        else:
            sol_pending.append(rep)

    if sol_pending:
        chunks = build_chunks(sol_pending)
        emit_progress(
            "classifying",
            62,
            f"大模型精分类 {len(sol_pending)} 条，共 {len(chunks)} 批",
            completed_chunks=0,
            total_chunks=len(chunks),
        )
        pass_two = run_chunks(client, chunks, "精分类", 62, 22)
        per_message.update(pass_two)
        for rep in sol_pending:
            result_source[rep["_llm_id"]] = client

    for rep in representatives:
        parsed = per_message.get(rep["_llm_id"])
        if parsed is not None:
            owner = result_source.get(rep["_llm_id"], client)
            save_message_memo(owner, contract_for(owner), taxonomy_for_prompt, rep, parsed)

    expanded = {}
    for rep in representatives:
        parsed = per_message.get(rep["_llm_id"])
        if parsed is None:
            continue
        for member_id in members[rep["_llm_id"]]:
            expanded[member_id] = parsed
    demand_map, risks, reviews = apply_taxonomy_gate(expanded, messages_by_id, taxonomy_by_label)

    expected_message_ids = set(messages_by_id)
    if set(expanded) != expected_message_ids:
        raise LLMAnalysisError(
            f"完整性门禁失败：应分析 {len(expected_message_ids)} 条，实际确认 {len(expanded)} 条。"
        )
    with client._usage_lock:
        client.expected_messages = len(expected_message_ids)
        client.analyzed_messages = len(expanded)
        client.unique_messages = len(representatives)
        client.deduped_messages = deduped_messages
    return calibrate_demand_stages(messages, taxonomy, demand_map), risks, reviews


RISK_BUCKETS = (
    "出雾故障/雾小",
    "加热异常/不热",
    "水箱/漏水异常",
    "实际伤发/卡发/烫伤",
    "实际漏电/冒烟/电气安全",
    "按键/控制异常",
    "疑似二手/外观异常",
    "结构/配件异常",
    "其他质量异常",
    "普通使用或效果咨询",
)

RISK_DEMAND_LABELS = {
    "出雾故障/雾小",
    "加热异常/不热",
    "水箱/漏水异常",
    "实际伤发/卡发/烫伤",
    "实际漏电/冒烟/电气安全",
    "按键/控制异常",
    "疑似二手/外观异常",
}


def seed_risk_candidates_from_demands(demand_map, risk_map):
    """Ensure every classified P0 fault reaches the independent evidence gate."""
    for msg_id, demands in demand_map.items():
        existing = {str(item.get("risk_type") or "") for item in risk_map.get(msg_id, [])}
        for demand in demands:
            label = str(demand.get("label") or "")
            if label not in RISK_DEMAND_LABELS or label in existing:
                continue
            risk_map[msg_id].append(
                {
                    "risk_type": label,
                    "reason": f"需求分类命中强风险词条“{label}”，进入独立证据二审。",
                    "priority": "P0",
                }
            )
            existing.add(label)
    return risk_map


def normalize_risk_bucket(candidate_type, text):
    """Collapse free-form model labels into the stable risk vocabulary."""
    value = f"{candidate_type} {text}"
    if re.search(r"漏电|触电|电流|麻手|冒烟|电线|外皮|烧焦", value):
        return "实际漏电/冒烟/电气安全"
    if re.search(r"烫伤|烫到|烫手|烫脸|烫头皮|头皮疼|流血|伤发|头发焦|烧头发|卡发|掉发", value):
        return "实际伤发/卡发/烫伤"
    if re.search(r"漏水|水箱.*漏|漏.*水箱", value):
        return "水箱/漏水异常"
    if re.search(r"不出雾|没雾|无雾|雾小|不喷水|不出水|喷雾.*异常|冷雾.*异常", value):
        return "出雾故障/雾小"
    if re.search(r"不加热|不热|没温度|温度低|加热异常|加热慢|温度上不去|预热.*异常", value):
        return "加热异常/不热"
    if re.search(r"按键|按不上|按不动|指示灯|无反应|没反应|不灵|失灵|报警|控制异常", value):
        return "按键/控制异常"
    if re.search(r"划痕|磕碰|残胶|二手|用过|旧的|外观.*异常|包装.*破|表面.*损", value):
        return "疑似二手/外观异常"
    if re.search(r"松动|扣不紧|脱落|掉了|掉下来|配件.*坏|配件.*缺|卡扣", value):
        return "结构/配件异常"
    if re.search(r"质量|坏了|故障|异常|损坏|不能工作|无法工作", value):
        return "其他质量异常"
    if re.search(r"效果|卷不|卷不起|夹不|不持久|不好用|怎么用|如何|教程|操作|预热|加水|水箱", value):
        return "普通使用或效果咨询"
    return "普通使用或效果咨询"


def risk_evidence_score(item):
    text = f"{item.get('text', '')} {item.get('context', '')}"
    score = 0
    score += 3 * len(re.findall(r"我|我的|用了|收到|刚刚|已经|还是|又", text))
    score += 4 * len(re.findall(r"烫伤|流血|冒烟|漏电|触电|头发焦|电线坏|加水.*还是|擦了.*不出", text))
    score -= 5 * len(re.findall(r"会不会|是否会|能不能|可以不|担心|怕不|听说|别人|正常吗", text))
    return score


def passes_deterministic_risk_gate(group):
    """Apply the non-negotiable evidence rules before asking the model to confirm."""
    bucket = group["bucket"]
    texts = [str(item.get("text") or "") for item in group.get("messages") or []]
    contexts = [str(item.get("context") or "") for item in group.get("messages") or []]
    direct = "\n".join(texts)
    combined = direct + "\n" + "\n".join(contexts)
    if bucket in {"结构/配件异常", "其他质量异常", "普通使用或效果咨询"}:
        return False
    if bucket == "出雾故障/雾小":
        context_only = "\n".join(contexts)
        repeated = len(texts) >= 2 or len(re.findall(r"不出雾|没雾|无雾|不喷水|雾(?:很|太)?小", context_only)) >= 2
        troubleshot = bool(
            re.search(
                r"(?:加(?:了|过)?水|装(?:好|过)?水箱|擦(?:了|过)|清理(?:了|过)|按(?:照)?教程|试(?:了|过)|重新装).{0,35}(?:还是|仍然|依旧|也)?(?:不出雾|没雾|无雾|不喷水)",
                combined,
            )
        )
        return repeated or troubleshot
    if bucket == "疑似二手/外观异常":
        return bool(re.search(r"二手|别人退|用过的|使用痕迹|残胶|划痕|磨损|旧的", direct))
    if bucket == "按键/控制异常":
        return bool(re.search(r"按不上|按不动|失灵|不灵|没反应|无反应|坏了|显示.{0,8}(?:异常|不亮|乱)|指示灯.{0,8}(?:不亮|异常)", direct))
    if bucket == "实际伤发/卡发/烫伤":
        return bool(re.search(r"烫伤|烫到|烫了|头皮疼|流血|头发(?:焦|断|掉)|烧(?:焦)?头发|卡(?:住)?头发|夹头发|扯头发", direct)) and not bool(
            re.search(r"会不会|是否会|能不能|担心|怕不怕", direct)
        )
    if bucket == "实际漏电/冒烟/电气安全":
        return bool(re.search(r"漏电|触电|电到|麻手|冒烟|炸响|电线.{0,8}(?:破|坏|露)", direct)) and not bool(
            re.search(r"会不会|是否会|能不能|担心", direct)
        )
    if bucket == "水箱/漏水异常":
        return bool(re.search(r"漏水|渗水|水箱.{0,10}(?:坏|裂|漏|不出水)", direct))
    if bucket == "加热异常/不热":
        return bool(re.search(r"不加热|不热|没温度|没有温度|温度上不去|加热异常|一直不热", direct))
    return False


def review_risk_candidates(client, messages, risk_map, review_map):
    """Group candidates first, then run a bounded parallel evidence gate."""
    messages_by_id = {item["_llm_id"]: item for item in messages}
    groups = {}
    seen = set()
    for msg_id, items in risk_map.items():
        message = messages_by_id.get(msg_id)
        if message is None:
            continue
        for item in items:
            bucket = normalize_risk_bucket(item["risk_type"], message["text"])
            if bucket == "普通使用或效果咨询":
                continue
            key = (message["product"], message["sender"], bucket)
            if (msg_id, bucket) in seen:
                continue
            seen.add((msg_id, bucket))
            group = groups.setdefault(
                key,
                {
                    "id": f"rg{len(groups) + 1:06d}",
                    "product": message["product"],
                    "buyer": message["sender"],
                    "bucket": bucket,
                    "messages": [],
                },
            )
            group["messages"].append(
                {
                    "id": msg_id,
                    "text": message["text"][:1200],
                    "context": message.get("context", "")[:2200],
                    "candidate_reason": item["reason"][:500],
                }
            )
    if not groups:
        return defaultdict(list)

    candidates = [group for group in groups.values() if passes_deterministic_risk_gate(group)]
    if not candidates:
        return defaultdict(list)
    for group in candidates:
        group["messages"].sort(key=risk_evidence_score, reverse=True)
        group["messages"] = group["messages"][:6]
        group["representative_id"] = group["messages"][0]["id"]

    system_prompt = f"""你是客服质量与安全风险二审员。只有证据明确通过门禁的记录才能进入正式风险表。
确认条件：买家本人已经实际发生较强安全、质量或履约异常，并且原话或同会话上下文有直接证据。
必须排除：购买前“会不会”式担忧、客服主动介绍、评论转述、普通教程咨询、不会操作、首次模糊提问、正常预热现象、单纯不喜欢或造型效果不满意。
856“不出雾”只有在同一会话重复出现，或明确已经加水、安装/擦拭/按教程排查后仍不出雾时确认；第一次问为什么不出雾只算普通售后。
同一产品、同一买家、同类风险已经合并；每组最多确认一条代表记录。风险类别只能从以下固定值选择：{json.dumps(RISK_BUCKETS, ensure_ascii=False)}。
只返回合法 JSON。"""

    # Risk evidence carries long conversation context. Smaller requests are more
    # reliable on OpenAI-compatible gateways and avoid repeated 75-second timeouts.
    chunks = build_chunks_for_objects(candidates, max_items=8, max_chars=6000)

    def build_review_prompt(chunk):
        return f"""逐组二审风险候选，返回：
{{"decisions":[["分组id",true或false,"固定风险类别","证据理由","P0或P1"]]}}
每组都必须返回决定；证据不足一律 false；不要把普通使用/效果咨询确认成风险。

候选分组：
{json.dumps(chunk, ensure_ascii=False)}"""

    def complete_review_leaf(chunk, prompt, token_limit):
        retry_number = 0
        while True:
            retry_number += 1
            delay = required_retry_delay(retry_number)
            emit_progress(
                "retrying",
                87,
                f"风险二审接口超时，{delay} 秒后继续重试",
                retry_count=retry_number,
            )
            time.sleep(delay)
            response = complete_json_required(client, system_prompt, prompt, token_limit, "风险二审")
            try:
                return validate_decision_coverage(response, [item["id"] for item in chunk])
            except LLMAnalysisError:
                client.invalidate_cache(system_prompt, prompt, token_limit)

    def review_chunk(chunk):
        prompt = build_review_prompt(chunk)
        token_limit = min(3000, max(900, len(chunk) * 70))
        try:
            ensure_attempt_budget(client)
            response = client.complete_json(system_prompt, prompt, max_tokens=token_limit)
            return validate_decision_coverage(response, [item["id"] for item in chunk])
        except LLMAnalysisError as exc:
            if not exc.retryable:
                raise
            client.invalidate_cache(system_prompt, prompt, token_limit)
            if len(chunk) > 5:
                midpoint = len(chunk) // 2
                left = review_chunk(chunk[:midpoint])
                right = review_chunk(chunk[midpoint:])
                return {"decisions": (left.get("decisions") or []) + (right.get("decisions") or [])}
            return complete_review_leaf(chunk, prompt, token_limit)

    confirmed = defaultdict(list)
    configured_workers = max(1, int(os.environ.get("CHITU_LLM_WORKERS", "8")))
    with ThreadPoolExecutor(max_workers=min(configured_workers, len(chunks))) as executor:
        futures = {executor.submit(review_chunk, chunk): chunk for chunk in chunks}
        for future in as_completed(futures):
            chunk = futures[future]
            response = future.result()
            groups_by_id = {group["id"]: group for group in chunk}
            for row in response.get("decisions") or []:
                if isinstance(row, dict):
                    group_id = str(row.get("id") or row.get("group_id") or "")
                    raw_confirmed = row.get("confirmed")
                    risk_type = str(row.get("risk_bucket") or row.get("risk_type") or "")
                    reason = str(row.get("reason") or "风险二审确认")
                    priority = str(row.get("priority") or "P1")
                elif isinstance(row, list) and len(row) >= 2:
                    group_id = str(row[0])
                    raw_confirmed = row[1]
                    risk_type = str(row[2]) if len(row) > 2 else ""
                    reason = str(row[3]) if len(row) > 3 else "风险二审确认"
                    priority = str(row[4]) if len(row) > 4 else "P1"
                else:
                    continue
                accepted = raw_confirmed is True or str(raw_confirmed).strip().lower() == "true"
                group = groups_by_id.get(group_id)
                if not accepted or group is None:
                    continue
                stable_type = risk_type if risk_type in RISK_BUCKETS else group["bucket"]
                if stable_type == "普通使用或效果咨询":
                    continue
                confirmed[group["representative_id"]].append(
                    {
                        "risk_type": stable_type,
                        "reason": reason[:500],
                        "priority": priority if priority in ("P0", "P1") else "P1",
                        "all_text": "｜".join(message["text"] for message in group["messages"])[:30000],
                        "raw_hit": len(group["messages"]),
                    }
                )
    return confirmed


def build_chunks_for_objects(items, max_items=40, max_chars=14000):
    chunks = []
    current = []
    used = 0
    for item in items:
        size = len(json.dumps(item, ensure_ascii=False))
        if current and (len(current) >= max_items or used + size > max_chars):
            chunks.append(current)
            current = []
            used = 0
        current.append(item)
        used += size
    if current:
        chunks.append(current)
    return chunks


def aggregate_results(entries, taxonomy, demand_map, risk_map, review_map):
    taxonomy_map = {item["label"]: item for item in taxonomy}
    results = []
    for entry in entries:
        product = entry["product"]
        raw = entry["raw"]
        valid = entry["valid"]
        detail = []
        duplicates = []
        risks = []
        reviews = []
        retained = {}
        retained_detail = {}
        retained_risks = set()
        raw_hits = Counter()
        buyers_by_label = defaultdict(set)
        evidence_by_label = defaultdict(list)
        for msg in valid:
            demands = demand_map.get(msg["_llm_id"], [])
            msg["ai_stage"] = demands[0]["stage"] if demands else "未归类"
            for demand in demands:
                meta = taxonomy_map[demand["label"]]
                stage = demand["stage"]
                label = meta["label"]
                raw_hits[(stage, label)] += 1
                key = (msg["sender"], stage, label)
                if key in retained:
                    kept_detail = retained_detail[key]
                    kept_detail["all_text"] = f"{kept_detail['all_text']}｜{msg['text']}"[:30000]
                    kept_detail["raw_hit"] = raw_hits[(stage, label)]
                    duplicates.append(
                        {
                            **msg,
                            "buyer": msg["sender"],
                            "stage": stage,
                            "label": label,
                            "kept_text": retained[key]["text"],
                        }
                    )
                    continue
                retained[key] = msg
                buyers_by_label[(stage, meta["theme"], label)].add(msg["sender"])
                evidence_by_label[(stage, meta["theme"], label)].append(msg)
                detail_item = {
                    **msg,
                    "buyer": msg["sender"],
                    "stage": stage,
                    "theme": meta["theme"],
                    "label": label,
                    "all_text": msg["text"],
                    "raw_hit": raw_hits[(stage, label)],
                    "confidence": demand["confidence"],
                }
                detail.append(detail_item)
                retained_detail[key] = detail_item
            for risk in risk_map.get(msg["_llm_id"], []):
                risk_key = (msg["sender"], risk["risk_type"])
                if risk_key in retained_risks:
                    continue
                retained_risks.add(risk_key)
                risks.append(
                    {
                        **msg,
                        "buyer": msg["sender"],
                        "risk_label": risk["risk_type"],
                        "risk_reason": risk["reason"],
                        "all_text": risk.get("all_text") or msg["text"],
                        "raw_hit": risk.get("raw_hit", 1),
                        "priority": risk["priority"],
                        "demand_label": demands[0]["label"] if demands else risk["risk_type"],
                        "action": "立即人工复核实际发生情况，记录使用方式、批次、证据和最终处理结果。",
                    }
                )
            if not demands and review_map.get(msg["_llm_id"]):
                reviews.append({**msg, "buyer": msg["sender"], "reason": "；".join(review_map[msg["_llm_id"]])[:500]})
        stats = []
        for (stage, theme, label), buyers in buyers_by_label.items():
            meta = taxonomy_map[label]
            evidence = evidence_by_label[(stage, theme, label)]
            stats.append(
                {
                    "product": product,
                    "stage": stage,
                    "theme": theme,
                    "label": label,
                    "count": len(buyers),
                    "raw_hits": raw_hits[(stage, label)],
                    "duplicates": sum(1 for item in duplicates if item["stage"] == stage and item["label"] == label),
                    "priority": meta["priority"],
                    "definition": meta["definition"],
                    "buyers": "、".join(sorted(buyers)[:10]),
                    "quote": "｜".join(item["text"] for item in evidence[:5]),
                    "action": meta["action"],
                }
            )
        dates = [msg.get("date", "") for msg in valid if msg.get("date")]
        results.append(
            {
                "summary": {
                    "product": product,
                    "raw_messages": max([msg.get("source_record_count", 0) for msg in raw] or [0]) or len(raw),
                    "valid_messages": len(valid),
                    "conversations": len({msg.get("conversation_id", 0) for msg in raw}),
                    "buyers": len({msg["sender"] for msg in valid}),
                    "start_date": min(dates) if dates else "",
                    "end_date": max(dates) if dates else "",
                    "active_days": len(set(dates)),
                    "demand_count": len(detail),
                    "presale_count": sum(1 for item in detail if item["stage"] == "售前"),
                    "aftersale_count": sum(1 for item in detail if item["stage"] == "售后"),
                    "duplicate_count": len(duplicates),
                    "risk_count": len(risks),
                    "review_count": len(reviews),
                },
                "stats": sorted(stats, key=lambda item: (-item["count"], item["label"])),
                "detail": detail,
                "duplicates": duplicates,
                "risks": risks,
                "reviews": reviews,
                "valid_messages": valid,
            }
        )
    return results


def build_management_summary(client, results, taxonomy, notes, task_context="", baseline_metrics=None):
    stage_rankings = []
    for result in results:
        by_stage = {}
        active_days = result["summary"].get("active_days") or 1
        for stage in ("售前", "售后"):
            stage_rows = sorted(
                (item for item in result["stats"] if item["stage"] == stage),
                key=lambda item: (-item["count"], item["label"]),
            )
            stage_total = sum(item["count"] for item in stage_rows)
            by_stage[stage] = {
                "total": stage_total,
                "ranking": [
                    {
                        "rank": index,
                        "label": item["label"],
                        "count": item["count"],
                        "share": f"{item['count'] / stage_total:.1%}" if stage_total else "0.0%",
                        "daily": f"{item['count'] / active_days:.2f}",
                        "priority": item["priority"],
                        "action": item["action"],
                    }
                    for index, item in enumerate(stage_rows[:10], 1)
                ],
            }
        stage_rankings.append({"product": result["summary"]["product"], "stages": by_stage})
    baseline_display = {}
    if baseline_metrics:
        baseline_display["period"] = baseline_metrics.get("period", {})
        for stage in ("售前", "售后"):
            source = baseline_metrics.get(stage, {})
            baseline_display[stage] = {
                "total": source.get("total", 0),
                "days": source.get("days", 0),
                "ranking": [
                    {
                        "label": item.get("label", ""),
                        "count": item.get("count", 0),
                        "share": f"{float(item.get('share', 0)):.1%}",
                        "daily": f"{float(item.get('daily', 0)):.2f}",
                    }
                    for item in source.get("rows", [])[:20]
                ],
            }
    compact = {
        "task": task_context,
        "products": [result["summary"] for result in results],
        "stage_rankings": stage_rankings,
        "risk_examples": [
            {
                "product": item["product"],
                "risk": item["risk_label"],
                "text": item["text"][:300],
            }
            for result in results
            for item in result["risks"][:10]
        ][:30],
        "taxonomy_notes": notes,
        "baseline": baseline_display,
    }
    system_prompt = "你是电商业务分析负责人。严格执行用户确认的历史 GPT 输出规范，根据本次统计和证据输出管理者能直接使用的结论，不得引用旧报告数字。只返回JSON。"
    user_prompt = f"""请输出 5-6 条中文管理结论，每条包含“数据/发现、业务判断、具体动作”，返回：{{"summary":["结论1",... ]}}。
要求：
1. 先说明数据周期、买家、正式需求及售前售后结构，再分别解读售前榜首、售后榜首和重要风险。
2. TOP词条与排名只能使用 stage_rankings 中各阶段独立降序榜单，人数、占比必须原样引用，禁止自行相加或改名。
3. products 中的 active_days 是日均统一分母；自然日期跨度只能描述覆盖周期，禁止用它重算日均。日均必须原样使用 stage_rankings 的 daily，baseline 日均必须原样使用 baseline 各阶段 rows 的 daily。
4. 有 baseline 时，必须指出最重要的占比变化和日均变化；起止日期不同时不得仅按总人数判断改善或恶化。
5. 结论要回答为什么重要、意味着什么、应该改什么；动作必须落到详情页、说明书/视频、客服SOP、产品或质量治理。
6. 不要把客服/机器人/系统消息被过滤造成的“有效消息率低”当成经营问题，也不要把分析方法、噪声率、模型过程写成首要结论。
7. 使用大白话，避免空泛的“持续关注/加强管理”；不要机械添加“数据/发现、业务判断、具体动作”等小标题前缀，写成自然完整的中文结论。
本次AI分析结果：
{json.dumps(compact, ensure_ascii=False)}"""
    while True:
        response = complete_json_required(client, system_prompt, user_prompt, 2200, "管理结论生成")
        summary = [str(item).strip() for item in (response.get("summary") or []) if str(item).strip()]
        if len(summary) >= 3:
            return summary[:6]
        client.invalidate_cache(system_prompt, user_prompt, 2200)
        emit_progress("retrying", 90, "管理结论返回不完整，正在重新生成")
        time.sleep(3)


def add_conversation_context(entries):
    def message_key(message):
        text = re.sub(r"https?://\S+", "", str(message.get("text") or ""))
        text = re.sub(r"\s+", " ", text).strip()
        return (
            message.get("conversation_id", 0),
            message.get("source_line"),
            str(message.get("sender") or ""),
            str(message.get("date") or ""),
            str(message.get("time") or ""),
            text,
        )

    for entry in entries:
        conversations = defaultdict(list)
        positions = defaultdict(list)
        for message in entry.get("raw") or []:
            conversation_id = message.get("conversation_id", 0)
            positions[message_key(message)].append(len(conversations[conversation_id]))
            conversations[conversation_id].append(message)
        used_positions = Counter()
        for message in entry.get("valid") or []:
            conversation_id = message.get("conversation_id", 0)
            conversation = conversations.get(conversation_id, [])
            key = message_key(message)
            occurrence = used_positions[key]
            position = positions[key][occurrence] if occurrence < len(positions[key]) else None
            used_positions[key] += 1
            if position is None:
                message["context"] = message["text"]
                continue
            window_start = max(0, position - 4)
            window = conversation[window_start : min(len(conversation), position + 3)]
            lines = []
            for neighbor_position, neighbor in enumerate(window, start=window_start):
                if neighbor_position == position:
                    marker = "当前买家"
                elif neighbor.get("role") == "buyer":
                    marker = "同买家"
                else:
                    marker = "客服"
                neighbor_text = str(neighbor.get("text") or "").strip().replace("\n", " ")
                if neighbor_text:
                    lines.append(f"{marker}: {neighbor_text[:500]}")
            message["context"] = "\n".join(lines)[:2400]


def analyze_entries_with_llm(entries, task_context="", baseline_context="", reference_taxonomy=None, baseline_metrics=None):
    add_conversation_context(entries)
    all_messages = []
    counter = 0
    for entry in entries:
        for original in entry["valid"]:
            counter += 1
            original["_llm_id"] = f"m{counter:07d}"
            original["product"] = entry["product"]
            all_messages.append(original)
    if not all_messages:
        raise LLMAnalysisError("没有可供模型分析的有效买家消息。")
    client = OpenAICompatibleClient()
    historical_logic, history_source = load_historical_logic()
    if reference_taxonomy:
        additions = discover_taxonomy_additions(
            client,
            all_messages,
            reference_taxonomy,
            task_context,
            historical_logic,
        )
        taxonomy = [*reference_taxonomy, *additions]
        notes = [
            f"沿用正式 V1 词条名称、定义、阶段和适用产品；本期严格新增 {len(additions)} 个旧词条无法覆盖的词条。"
        ]
        emit_progress("taxonomy", 11, "正在载入正式 V1 词条口径")
    else:
        emit_progress("taxonomy", 11, "正在根据真实聊天建立需求词条")
        taxonomy, notes = discover_taxonomy(
            client,
            all_messages,
            task_context,
            baseline_context,
            historical_logic,
        )
    emit_progress("taxonomy", 14, f"需求词条已建立，共 {len(taxonomy)} 个")
    demand_map, risks, reviews = classify_chunks(client, all_messages, taxonomy)
    emit_progress("risk_review", 86, "正在执行质量与安全风险证据二审")
    risks = seed_risk_candidates_from_demands(demand_map, risks)
    risks = review_risk_candidates(client, all_messages, risks, reviews)
    emit_progress("summarizing", 88, "正在汇总统计和管理结论")
    results = aggregate_results(entries, taxonomy, demand_map, risks, reviews)
    summary = build_management_summary(client, results, taxonomy, notes, task_context, baseline_metrics)
    emit_progress("reporting", 93, "分析完成，正在写入标准 Excel 母版")
    usage = client.usage_summary()
    usage["history_logic_source"] = history_source
    usage["history_logic_chars"] = len(historical_logic)
    if not usage["analysis_complete"] or usage["degraded_messages"]:
        raise LLMAnalysisError(
            f"完整性门禁失败：应分析 {usage['expected_messages']} 条，实际完成 {usage['analyzed_messages']} 条。"
        )
    return results, taxonomy, summary, usage
