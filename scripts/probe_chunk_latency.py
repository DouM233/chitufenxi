"""Count valid buyer messages per test file and time one realistic classification chunk."""
import json
import os
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
os.environ.setdefault("PYTHONUTF8", "1")

from generate_report import parse_messages, valid_customer_messages  # noqa: E402
from llm_analysis import OpenAICompatibleClient, build_chunks, CLASSIFICATION_SYSTEM_PROMPT  # noqa: E402

TEST_DIR = Path(r"C:\Users\admin\Desktop\赤兔历史分析结果")

candidates = [
    "366聊天记录.log",
    "五合一聊天记录.log",
    "856聊天记录.log",
    "chat_fulltext_997158824546_20260819_20260825.raw.csv",
    "chat_fulltext_1019049895823_20260819_20260825.raw.csv",
    "chat_fulltext_892227536130_20260801_20260831_1976.raw.csv",
    "chat_fulltext_1053521946066_20260801_20260825.raw.csv",
    "100384272218聊天记录.log",
]

print(f"{'文件':<58}{'总消息':>8}{'有效买家':>10}")
for name in candidates:
    path = TEST_DIR / name
    if not path.exists():
        continue
    raw = parse_messages("p", path)
    valid = valid_customer_messages(raw)
    for index, message in enumerate(valid, 1):
        message["_llm_id"] = f"m{index:07d}"
        message["product"] = "p"
    chunks = build_chunks(valid)
    est_tokens = sum(len(json.dumps(c, ensure_ascii=False)) for c in [chunks[0]]) if chunks else 0
    print(f"{name:<58}{len(raw):>8}{len(valid):>10}  -> {len(chunks)} 批, 首批字符数 {est_tokens}")
