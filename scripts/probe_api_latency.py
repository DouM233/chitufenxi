"""Quick probe of the upstream LLM API latency and concurrency behavior."""
import json
import os
import time
import urllib.request
import concurrent.futures
from pathlib import Path

BASE = (os.environ.get("CHITU_LLM_API_BASE") or "").rstrip("/")
KEY = os.environ.get("CHITU_LLM_API_KEY") or ""
MODEL = os.environ.get("CHITU_ANALYSIS_MODEL") or "gpt-5.6-sol"
ENDPOINT = BASE + ("/chat/completions" if BASE.endswith("/v1") else "/v1/chat/completions")


def call(prompt, max_tokens=500, timeout=300):
    payload = {
        "model": MODEL,
        "messages": [
            {"role": "system", "content": "只返回合法 JSON。"},
            {"role": "user", "content": prompt},
        ],
        "response_format": {"type": "json_object"},
        "max_tokens": max_tokens,
    }
    body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    req = urllib.request.Request(
        ENDPOINT,
        data=body,
        method="POST",
        headers={"Authorization": f"Bearer {KEY}", "Content-Type": "application/json"},
    )
    start = time.time()
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            result = json.loads(resp.read().decode("utf-8"))
        elapsed = time.time() - start
        usage = result.get("usage") or {}
        content = result.get("choices", [{}])[0].get("message", {}).get("content", "")
        # some proxies expose reasoning effort/timing hints
        extra_keys = [k for k in result.keys() if k not in ("id", "object", "created", "model", "choices", "usage", "system_fingerprint")]
        return {
            "ok": True,
            "elapsed": round(elapsed, 1),
            "prompt_tokens": usage.get("prompt_tokens"),
            "completion_tokens": usage.get("completion_tokens"),
            "content_len": len(content or ""),
            "extra_keys": extra_keys,
        }
    except Exception as exc:
        return {"ok": False, "elapsed": round(time.time() - start, 1), "error": f"{type(exc).__name__}: {exc}"}


small_prompt = '返回 {"ok": true, "note": "延迟测试"}'
medium_prompt = (
    "分析下面5条客服消息，每条返回需求词条或null，返回JSON {\"rows\":[[\"id\",\"词条或null\"]]}：\n"
    + json.dumps([
        {"id": "m1", "text": "这个怎么用啊有没有教程"},
        {"id": "m2", "text": "漏水了怎么办"},
        {"id": "m3", "text": "多少钱有优惠吗"},
        {"id": "m4", "text": "能不能卷刘海"},
        {"id": "m5", "text": "退货地址发我一下"},
    ], ensure_ascii=False)
)

print(f"endpoint: {ENDPOINT}")
print(f"model: {MODEL}")

print("\n[1] 小请求 x2（串行，测基础延迟）")
for i in range(2):
    print(" ", call(small_prompt, max_tokens=100))

print("\n[2] 中等请求 x1（分类样式，测典型单调用耗时）")
print(" ", call(medium_prompt, max_tokens=800))

print("\n[3] 中等请求 x6（并发，测并发吞吐与限流）")
with concurrent.futures.ThreadPoolExecutor(max_workers=6) as ex:
    futures = [ex.submit(call, medium_prompt, max_tokens=800) for _ in range(6)]
    t0 = time.time()
    for i, f in enumerate(concurrent.futures.as_completed(futures)):
        r = f.result()
        print(f"  完成 {round(time.time()-t0,1)}s ->", r)
