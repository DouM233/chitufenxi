"""Probe optimization levers: smaller chunks, higher concurrency, reasoning_effort."""
import json
import os
import sys
import time
import urllib.request
import concurrent.futures
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from generate_report import parse_messages, valid_customer_messages  # noqa: E402
from llm_analysis import build_chunks, CLASSIFICATION_SYSTEM_PROMPT  # noqa: E402

BASE = (os.environ.get("CHITU_LLM_API_BASE") or "").rstrip("/")
KEY = os.environ.get("CHITU_LLM_API_KEY") or ""
MODEL = os.environ.get("CHITU_ANALYSIS_MODEL") or "gpt-5.6-sol"
ENDPOINT = BASE + "/v1/chat/completions"

TEST_FILE = Path(r"C:\Users\admin\Desktop\赤兔历史分析结果\chat_fulltext_997158824546_20260819_20260825.raw.csv")

taxonomy = [
    {"label": "使用方法/教程", "default_stage": "售后", "allowed_stages": ["售后"], "product_prefixes": [], "theme": "使用学习", "definition": "已购后咨询操作、教程、视频。"},
    {"label": "加水/水质", "default_stage": "售前", "allowed_stages": ["售前"], "product_prefixes": [], "theme": "使用预期", "definition": "购买前询问是否加水、加什么水。"},
    {"label": "水箱拆装/加水方法", "default_stage": "售后", "allowed_stages": ["售后"], "product_prefixes": [], "theme": "使用学习", "definition": "已购后询问水箱如何取装、从哪里加水。"},
    {"label": "吸力/档位", "default_stage": "售前", "allowed_stages": ["售前"], "product_prefixes": [], "theme": "参数规格", "definition": "购买前询问吸力档位、力度。"},
    {"label": "退款退货", "default_stage": "售后", "allowed_stages": ["售后"], "product_prefixes": [], "theme": "售后服务", "definition": "实际提出退款退货换货。"},
    {"label": "款式/型号区别", "default_stage": "售前", "allowed_stages": ["售前"], "product_prefixes": [], "theme": "购买决策", "definition": "购买前比较各版本差异。"},
    {"label": "价格优惠/活动", "default_stage": "售前", "allowed_stages": ["售前"], "product_prefixes": [], "theme": "交易政策", "definition": "购买前问价格、优惠。"},
    {"label": "退换/试用政策", "default_stage": "售前", "allowed_stages": ["售前"], "product_prefixes": [], "theme": "交易政策", "definition": "购买前问试用、七天无理由、运费险。"},
    {"label": "发货时效/物流", "default_stage": "售前", "allowed_stages": ["售前"], "product_prefixes": [], "theme": "交易政策", "definition": "下单前问发货到货时间。"},
    {"label": "效果/好用程度", "default_stage": "售前", "allowed_stages": ["售前"], "product_prefixes": [], "theme": "使用预期", "definition": "购买前问有没有效果、好不好用。"},
    {"label": "疑似二手/外观异常", "default_stage": "售后", "allowed_stages": ["售后"], "product_prefixes": [], "theme": "质量安全", "definition": "反馈收到疑似使用过的商品。"},
    {"label": "不出雾/不出水故障", "default_stage": "售后", "allowed_stages": ["售后"], "product_prefixes": [], "theme": "质量安全", "definition": "实际使用后反馈不出水不出雾。"},
]


def build_prompt(chunk):
    return f"""V1词条：
{json.dumps(taxonomy, ensure_ascii=False)}

逐条分析下面消息，返回紧凑 JSON：
{{
  "analyzed_ids": ["本批每一条消息id，必须全部返回且不重复"],
  "demands": [["消息id","V1词条","售前或售后",0.0到1.0置信度]],
  "risks": [["消息id","风险类型","为什么属于本人实际发生","P0或P1"]],
  "reviews": [["消息id","无法稳定分类的原因"]]
}}

消息：
{json.dumps(chunk, ensure_ascii=False)}"""


def call(system, prompt, max_tokens, extra_body=None, timeout=600):
    payload = {
        "model": MODEL,
        "messages": [
            {"role": "system", "content": system},
            {"role": "user", "content": prompt},
        ],
        "response_format": {"type": "json_object"},
        "max_tokens": max_tokens,
    }
    if extra_body:
        payload.update(extra_body)
    body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    req = urllib.request.Request(
        ENDPOINT, data=body, method="POST",
        headers={"Authorization": f"Bearer {KEY}", "Content-Type": "application/json"},
    )
    start = time.time()
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            result = json.loads(resp.read().decode("utf-8"))
        usage = result.get("usage") or {}
        return {
            "ok": True, "elapsed": round(time.time() - start, 1),
            "completion_tokens": usage.get("completion_tokens"),
        }
    except Exception as exc:
        detail = ""
        if hasattr(exc, "read"):
            try:
                detail = exc.read().decode("utf-8", "replace")[:300]
            except Exception:
                pass
        return {"ok": False, "elapsed": round(time.time() - start, 1),
                "error": f"{type(exc).__name__}: {exc} {detail}"}


raw = parse_messages("小气泡", TEST_FILE)
valid = valid_customer_messages(raw)
for index, message in enumerate(valid, 1):
    message["_llm_id"] = f"m{index:07d}"
    message["product"] = "小气泡"
    message.setdefault("context", "")
chunks = build_chunks(valid)
big = chunks[0]          # 65 messages
small = chunks[0][:32]   # 32 messages

print("[A] 32条/批（当前65条的一半）")
print("   ", call(CLASSIFICATION_SYSTEM_PROMPT, build_prompt(small), 1800))

print("\n[B] 65条/批 + reasoning_effort=low（若网关支持）")
print("   ", call(CLASSIFICATION_SYSTEM_PROMPT, build_prompt(big), 3500, extra_body={"reasoning_effort": "low"}))

print("\n[C] 65条/批 + reasoning_effort=minimal")
print("   ", call(CLASSIFICATION_SYSTEM_PROMPT, build_prompt(big), 3500, extra_body={"reasoning_effort": "minimal"}))

print("\n[D] 32条/批 x 12 并发（测吞吐/限流）")
t0 = time.time()
with concurrent.futures.ThreadPoolExecutor(max_workers=12) as ex:
    futures = [ex.submit(call, CLASSIFICATION_SYSTEM_PROMPT, build_prompt(chunks[i][:32]), 1800) for i in range(1, 13)]
    for f in concurrent.futures.as_completed(futures):
        print(f"    {round(time.time()-t0,1)}s ->", f.result())
print(f"    总耗时 {round(time.time()-t0,1)}s")
