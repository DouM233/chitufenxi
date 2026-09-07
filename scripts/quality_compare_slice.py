"""Quality check: sol-only vs cascade classification on a real message slice.

Runs the same pipeline twice with separate cold caches on the first N valid
buyer messages of a real file. The cascade run reuses the sol run's taxonomy so
only the classification step differs, then diffs per-message demand decisions.
"""
import json
import os
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from generate_report import parse_messages, valid_customer_messages  # noqa: E402
from llm_analysis import analyze_entries_with_llm  # noqa: E402

SLICE = int(os.environ.get("SLICE", "300"))
TEST_FILE = Path(r"C:\Users\admin\Desktop\赤兔历史分析结果\chat_fulltext_997158824546_20260819_20260825.raw.csv")
WORK_ROOT = Path(__file__).resolve().parents[1] / "storage" / "quality_check"


def extract_per_message(results):
    per_msg = {}
    for result in results:
        for row in result["detail"]:
            key = (row["sender"], row.get("date", ""), row.get("time", ""), row["text"])
            per_msg.setdefault(key, set()).add((row["label"], row["stage"]))
    return per_msg


def run(mode, valid, reference_taxonomy=None):
    cache_dir = WORK_ROOT / mode / "llm_cache"
    cache_dir.mkdir(parents=True, exist_ok=True)
    os.environ["CHITU_LLM_CACHE_DIR"] = str(cache_dir)
    os.environ["CHITU_MESSAGE_CACHE"] = "1"
    if mode == "cascade":
        os.environ["CHITU_CLASSIFY_MODEL"] = "gpt-5.4-mini"
    else:
        os.environ["CHITU_CLASSIFY_MODEL"] = ""
    start = time.time()
    results, taxonomy, _, usage = analyze_entries_with_llm(
        [{"product": "小气泡", "raw": [], "valid": [dict(m) for m in valid]}],
        task_context="质量对比：小气泡切片",
        reference_taxonomy=reference_taxonomy,
    )
    return {
        "usage": usage,
        "seconds": round(time.time() - start, 1),
        "taxonomy": taxonomy,
        "per_msg": extract_per_message(results),
    }


valid = valid_customer_messages(parse_messages("小气泡", TEST_FILE))[:SLICE]
both_sol = os.environ.get("BOTH_SOL", "") == "1"
label_b = "sol复跑(噪声基线)" if both_sol else "级联"
print(f"切片 {len(valid)} 条消息：先跑纯 sol 基线，再跑{label_b}（沿用基线词条表）……", flush=True)

sol = run("sol_only", valid)
print(f"纯 sol: {sol['seconds']}s, 调用={sol['usage']['api_calls']} tokens={sol['usage']['total_tokens']:,}", flush=True)
if both_sol:
    cascade = run("sol_again", valid, reference_taxonomy=sol["taxonomy"])
else:
    cascade = run("cascade", valid, reference_taxonomy=sol["taxonomy"])
u = cascade["usage"]
print(
    f"级联:   {cascade['seconds']}s, 主模型调用={u['api_calls']} 主模型tokens={u['total_tokens']:,} "
    f"初分类调用={u['screener_api_calls']} 初分类tokens={u['screener_prompt_tokens'] + u['screener_completion_tokens']:,} "
    f"去重={u['deduped_messages']} 升级={u['escalated_messages']}",
    flush=True,
)

keys = set(sol["per_msg"]) | set(cascade["per_msg"])
exact = sum(1 for key in keys if sol["per_msg"].get(key, set()) == cascade["per_msg"].get(key, set()))
labels_only = sum(
    1
    for key in keys
    if {label for label, _ in sol["per_msg"].get(key, set())} == {label for label, _ in cascade["per_msg"].get(key, set())}
)
print(f"\n逐消息对比：唯一消息 {len(keys)} 条，完全一致 {exact}（{exact / len(keys):.1%}），词条一致(阶段可异) {labels_only}（{labels_only / len(keys):.1%}）")
mismatches = [key for key in keys if sol["per_msg"].get(key, set()) != cascade["per_msg"].get(key, set())][:8]
for key in mismatches:
    print(f"  差异: {key[0]}「{key[3][:30]}」 sol={sorted(sol['per_msg'].get(key, set()))} cascade={sorted(cascade['per_msg'].get(key, set()))}")

WORK_ROOT.mkdir(parents=True, exist_ok=True)
(WORK_ROOT / "summary.json").write_text(
    json.dumps(
        {
            "slice": len(valid),
            "sol": {k: v for k, v in sol.items() if k != "per_msg"},
            "cascade": {k: v for k, v in cascade.items() if k != "per_msg"},
            "agreement": {"keys": len(keys), "exact": exact, "label_match": labels_only},
        },
        ensure_ascii=False,
        indent=2,
        default=str,
    ),
    encoding="utf-8",
)
print("saved -> storage/quality_check/summary.json")
