"""End-to-end harness: build report_data.json and run generate_report.py with timestamped progress."""
import json
import subprocess
import sys
import time
from pathlib import Path

PROJECT = Path(__file__).resolve().parents[1]
TEST_FILE = Path(r"C:\Users\admin\Desktop\赤兔历史分析结果\chat_fulltext_997158824546_20260819_20260825.raw.csv")

mode = sys.argv[1] if len(sys.argv) > 1 else "optimized"
out_dir = PROJECT / "storage" / "e2e_test" / mode
out_dir.mkdir(parents=True, exist_ok=True)

report_data = {
    "task_id": f"e2e_{mode}_{int(time.time())}",
    "product_name": "小气泡E2E测试",
    "analysis_type": "first_baseline",
    "user_message": "分析本期小气泡客服聊天记录",
    "chat_files": [
        {"name": TEST_FILE.name, "size": TEST_FILE.stat().st_size, "role": "chat_record", "path": str(TEST_FILE)}
    ],
    "baseline_files": [],
}
data_path = out_dir / "report_data.json"
data_path.write_text(json.dumps(report_data, ensure_ascii=False, indent=2), encoding="utf-8")
excel_path = out_dir / "报告.xlsx"

env_extra = {
    "PYTHONUTF8": "1",
    "CHITU_LLM_TIMEOUT": "240",
    "CHITU_LLM_RETRIES": "3",
    "CHITU_LLM_MAX_ATTEMPTS": "0",
    "CHITU_LLM_CACHE_DIR": str(out_dir / "llm_cache"),
}
if mode == "baseline":
    env_extra["CHITU_LLM_WORKERS"] = "6"
    env_extra["CHITU_LLM_REASONING_EFFORT"] = " "  # 留空 -> 不传参数，复现线上默认思考档
elif mode == "cascade":
    env_extra["CHITU_LLM_WORKERS"] = "12"
    env_extra["CHITU_LLM_REASONING_EFFORT"] = "low"
    env_extra["CHITU_CLASSIFY_MODEL"] = "gpt-5.4-mini"
    env_extra["CHITU_CASCADE_CONFIDENCE"] = "0.7"
else:
    env_extra["CHITU_LLM_WORKERS"] = "12"
    env_extra["CHITU_LLM_REASONING_EFFORT"] = "low"

import os

start = time.time()
proc = subprocess.Popen(
    [sys.executable, str(PROJECT / "scripts" / "generate_report.py"), str(data_path), str(excel_path)],
    cwd=str(PROJECT),
    env={**os.environ, **env_extra},
    stdout=subprocess.PIPE,
    stderr=subprocess.STDOUT,
    text=True,
    encoding="utf-8",
    errors="replace",
    bufsize=1,
)
stage_marks = []
for line in proc.stdout:
    line = line.rstrip()
    elapsed = time.time() - start
    if line.startswith("CHITU_PROGRESS "):
        try:
            payload = json.loads(line[len("CHITU_PROGRESS "):])
        except json.JSONDecodeError:
            continue
        key = (payload.get("stage"), payload.get("message"))
        if not stage_marks or stage_marks[-1][0] != key:
            stage_marks.append((key, elapsed))
            print(f"[{elapsed:7.1f}s] {payload.get('stage'):<14} {payload.get('message')}", flush=True)
    elif line.strip():
        print(f"[{elapsed:7.1f}s] (py) {line[:200]}", flush=True)
proc.wait()
total = time.time() - start
print(f"\nexit={proc.returncode} 总耗时 {total:.1f}s ({total/60:.1f} 分钟)")
if excel_path.exists():
    print(f"Excel 已生成: {excel_path} ({excel_path.stat().st_size} bytes)")
final = json.loads(data_path.read_text(encoding="utf-8"))
print("llm_usage:", json.dumps(final.get("llm_usage") or {}, ensure_ascii=False, indent=2))
print("taxonomy:", [t["label"] for t in final.get("taxonomy") or []])
