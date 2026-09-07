"""Time realistic classification chunk calls (65 messages, ~10k chars) against the real API."""
import json
import os
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from generate_report import parse_messages, valid_customer_messages  # noqa: E402
from llm_analysis import (  # noqa: E402
    OpenAICompatibleClient,
    build_chunks,
    CLASSIFICATION_SYSTEM_PROMPT,
)

TEST_FILE = Path(r"C:\Users\admin\Desktop\赤兔历史分析结果\chat_fulltext_997158824546_20260819_20260825.raw.csv")

raw = parse_messages("小气泡", TEST_FILE)
valid = valid_customer_messages(raw)
for index, message in enumerate(valid, 1):
    message["_llm_id"] = f"m{index:07d}"
    message["product"] = "小气泡"
    message.setdefault("context", "")
chunks = build_chunks(valid)
chunk = chunks[0]

taxonomy = [
    {"label": item[0], "default_stage": item[1], "allowed_stages": [item[1]], "product_prefixes": [], "theme": item[2], "definition": item[3]}
    for item in [
        ("使用方法/教程", "售后", "使用学习", "已购后咨询操作、教程、视频。"),
        ("加水/水质", "售前", "使用预期", "购买前询问是否加水、加什么水。"),
        ("水箱拆装/加水方法", "售后", "使用学习", "已购后询问水箱如何取装、从哪里加水。"),
        ("吸力/档位", "售前", "参数规格", "购买前询问吸力档位、力度。"),
        ("退款退货", "售后", "售后服务", "实际提出退款退货换货。"),
        ("款式/型号区别", "售前", "购买决策", "购买前比较各版本差异。"),
        ("价格优惠/活动", "售前", "交易政策", "购买前问价格、优惠。"),
        ("退换/试用政策", "售前", "交易政策", "购买前问试用、七天无理由、运费险。"),
        ("发货时效/物流", "售前", "交易政策", "下单前问发货到货时间。"),
        ("效果/好用程度", "售前", "使用预期", "购买前问有没有效果、好不好用。"),
        ("疑似二手/外观异常", "售后", "质量安全", "反馈收到疑似使用过的商品。"),
        ("不出雾/不出水故障", "售后", "质量安全", "实际使用后反馈不出水不出雾。"),
    ]
]

user_prompt = f"""V1词条：
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

client = OpenAICompatibleClient()
print(f"chunk 消息数: {len(chunk)}, prompt 字符数: {len(CLASSIFICATION_SYSTEM_PROMPT) + len(user_prompt)}")

t0 = time.time()
result = client.complete_json(CLASSIFICATION_SYSTEM_PROMPT, user_prompt, max_tokens=3500)
elapsed = time.time() - t0
n_demands = len(result.get("demands") or [])
print(f"单批耗时: {elapsed:.1f}s, 返回 demands {n_demands} 条, usage: {json.dumps(client.usage_summary(), ensure_ascii=False)}")
print(f"内容长度: {len(json.dumps(result, ensure_ascii=False))} 字符")

# 再测第二批（无缓存），确认稳定性
t0 = time.time()
result2 = client.complete_json(
    CLASSIFICATION_SYSTEM_PROMPT,
    user_prompt.replace(json.dumps(chunk[0], ensure_ascii=False), json.dumps(chunks[1][0], ensure_ascii=False)),
    max_tokens=3500,
)
print(f"第二批耗时: {time.time()-t0:.1f}s, demands {len(result2.get('demands') or [])}")
