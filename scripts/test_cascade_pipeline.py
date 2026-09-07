import os
import re
import sys
import tempfile
import unittest
from collections import Counter
from pathlib import Path
from threading import Lock
from unittest.mock import patch

SCRIPT_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(SCRIPT_DIR))

import llm_analysis


class FakeClassifier:
    """Duck-typed client covering everything classify_chunks touches."""

    def __init__(self, cache_dir, model="fake-main", responder=None):
        self.cache_dir = Path(cache_dir)
        self.model = model
        self.endpoint = "https://fake.invalid/v1/chat/completions"
        self.max_attempts = None
        self.attempts = 0
        self.calls = 0
        self.cache_hits = 0
        self.usage = Counter()
        self.expected_messages = 0
        self.analyzed_messages = 0
        self.unique_messages = 0
        self.deduped_messages = 0
        self.escalated_messages = 0
        self.screener_model = ""
        self.screener_usage = Counter()
        self._usage_lock = Lock()
        self.prompts_seen = []
        self.responder = responder or (lambda ids: ([], [], []))

    def has_cached(self, *_args):
        return False

    def invalidate_cache(self, *_args):
        return None

    def complete_json(self, _system_prompt, user_prompt, max_tokens=10000):
        del max_tokens
        self.calls += 1
        self.prompts_seen.append(user_prompt)
        ids = re.findall(r'"id":\s*"(m\d+)"', user_prompt)
        demands, risks, reviews = self.responder(ids)
        return {"analyzed_ids": ids, "demands": demands, "risks": risks, "reviews": reviews}


def make_messages(specs, product="测试产品"):
    return [
        {
            "_llm_id": msg_id,
            "product": product,
            "sender": f"买家{msg_id}",
            "text": text,
            "context": f"当前买家: {text}",
        }
        for msg_id, text in specs
    ]


TAXONOMY = [
    {"label": "使用方法", "stage": "售后", "stages": ["售后"], "product_prefixes": [], "theme": "使用", "definition": "询问使用方法。"},
    {"label": "退款退货", "stage": "售后", "stages": ["售后"], "product_prefixes": [], "theme": "售后", "definition": "提出退换。"},
]


class DedupeTests(unittest.TestCase):
    def test_groups_by_product_and_text(self):
        messages = make_messages(
            [("m1", "怎么使用"), ("m2", "怎么使用"), ("m3", "怎么使用", ), ("m4", "退款")]
        ) + make_messages([("m5", "怎么使用")], product="另一商品")
        representatives, members = llm_analysis.dedupe_messages_for_classification(messages)
        self.assertEqual(3, len(representatives))
        self.assertEqual(["m1", "m2", "m3"], members["m1"])
        self.assertEqual(["m4"], members["m4"])
        self.assertEqual(["m5"], members["m5"])
        self.assertEqual("m1", representatives[0]["_llm_id"])

    def test_preserves_first_occurrence_order(self):
        messages = make_messages([("m1", "乙"), ("m2", "甲"), ("m3", "乙")])
        representatives, _ = llm_analysis.dedupe_messages_for_classification(messages)
        self.assertEqual(["m1", "m2"], [item["_llm_id"] for item in representatives])


class TriageTests(unittest.TestCase):
    def test_risk_or_review_always_escalates(self):
        message = make_messages([("m1", "怎么使用")])[0]
        base = {"demands": [], "risks": [], "reviews": []}
        with_risk = {**base, "risks": [{"risk_type": "漏水", "reason": "实际", "priority": "P0"}]}
        with_review = {**base, "reviews": ["语义不完整"]}
        self.assertTrue(llm_analysis.needs_screener_review(with_risk, message, 0.75))
        self.assertTrue(llm_analysis.needs_screener_review(with_review, message, 0.75))

    def test_low_confidence_escalates(self):
        message = make_messages([("m1", "怎么使用")])[0]
        parsed = {"demands": [{"label": "使用方法", "stage": "售后", "confidence": 0.5}], "risks": [], "reviews": []}
        self.assertTrue(llm_analysis.needs_screener_review(parsed, message, 0.75))
        parsed["demands"][0]["confidence"] = 0.9
        self.assertFalse(llm_analysis.needs_screener_review(parsed, message, 0.75))

    def test_signal_words_escalate_empty_results_but_clean_chitchat_does_not(self):
        question = make_messages([("m1", "这个怎么使用")])[0]
        chitchat = make_messages([("m2", "好的谢谢")])[0]
        empty = {"demands": [], "risks": [], "reviews": []}
        self.assertTrue(llm_analysis.needs_screener_review(empty, question, 0.75))
        self.assertFalse(llm_analysis.needs_screener_review(dict(empty), chitchat, 0.75))


class ParseChunkRowsTests(unittest.TestCase):
    def test_seeds_empty_bucket_from_analyzed_ids(self):
        parsed = llm_analysis.parse_chunk_rows({"analyzed_ids": ["m1", "m2"], "demands": [], "risks": [], "reviews": []})
        self.assertEqual(["m1", "m2"], sorted(parsed))
        self.assertEqual({"demands": [], "risks": [], "reviews": []}, parsed["m1"])

    def test_list_and_dict_row_forms(self):
        result = {
            "analyzed_ids": ["m1", "m2", "m3"],
            "demands": [["m1", "使用方法", "售后", 0.9], {"id": "m2", "label": "退款退货", "stage": "售后", "confidence": 0.8}],
            "risks": [["m3", "漏水", "实际漏水", "P0"]],
            "reviews": [["m1", "语义不完整"]],
        }
        parsed = llm_analysis.parse_chunk_rows(result)
        self.assertEqual("使用方法", parsed["m1"]["demands"][0]["label"])
        self.assertEqual(0.8, parsed["m2"]["demands"][0]["confidence"])
        self.assertEqual("漏水", parsed["m3"]["risks"][0]["risk_type"])
        self.assertEqual("语义不完整", parsed["m1"]["reviews"][0])


class TaxonomyGateTests(unittest.TestCase):
    def test_unknown_label_dropped_and_stage_enforced(self):
        messages = {item["_llm_id"]: item for item in make_messages([("m1", "怎么用"), ("m2", "怎么用")])}
        per_message = {
            "m1": {"demands": [{"label": "不存在词条", "stage": "售后", "confidence": 0.9}], "risks": [], "reviews": []},
            "m2": {"demands": [{"label": "使用方法", "stage": "售前", "confidence": 0.9}], "risks": [], "reviews": []},
        }
        taxonomy_by_label = {item["label"]: item for item in TAXONOMY}
        demand_map, risks, reviews = llm_analysis.apply_taxonomy_gate(per_message, messages, taxonomy_by_label)
        self.assertNotIn("m1", demand_map)
        self.assertEqual("售后", demand_map["m2"][0]["stage"])
        self.assertFalse(risks)
        self.assertFalse(reviews)

    def test_product_prefix_blocks_cross_sku_label(self):
        messages = {item["_llm_id"]: item for item in make_messages([("m1", "怎么换头")], product="856卷发棒")}
        per_message = {"m1": {"demands": [{"label": "换头教程", "stage": "售后", "confidence": 0.9}], "risks": [], "reviews": []}}
        taxonomy_by_label = {
            "换头教程": {"label": "换头教程", "stage": "售后", "stages": ["售后"], "product_prefixes": ["五合一"]}
        }
        demand_map, _, _ = llm_analysis.apply_taxonomy_gate(per_message, messages, taxonomy_by_label)
        self.assertNotIn("m1", demand_map)


class MessageMemoTests(unittest.TestCase):
    def test_roundtrip_and_taxonomy_invalidation(self):
        parsed = {"demands": [{"label": "使用方法", "stage": "售后", "confidence": 0.9}], "risks": [], "reviews": []}
        message = make_messages([("m1", "怎么使用")])[0]
        with tempfile.TemporaryDirectory() as directory:
            client = FakeClassifier(directory)
            taxonomy = [{"label": "使用方法", "definition": "d"}]
            llm_analysis.save_message_memo(client, "sys", taxonomy, message, parsed)
            hit = llm_analysis.load_message_memo(client, "sys", taxonomy, message)
            self.assertEqual(parsed, hit)
            miss = llm_analysis.load_message_memo(client, "sys", [{"label": "其它", "definition": "x"}], message)
            self.assertIsNone(miss)

    def test_disabled_by_env(self):
        message = make_messages([("m1", "怎么使用")])[0]
        with tempfile.TemporaryDirectory() as directory, patch.dict(os.environ, {"CHITU_MESSAGE_CACHE": "0"}, clear=False):
            client = FakeClassifier(directory)
            llm_analysis.save_message_memo(client, "sys", [], message, {"demands": [], "risks": [], "reviews": []})
            self.assertIsNone(llm_analysis.load_message_memo(client, "sys", [], message))


class BuildChunksTests(unittest.TestCase):
    def test_env_caps_take_effect(self):
        messages = make_messages([(f"m{i:03d}", f"消息{i}") for i in range(1, 121)])
        with patch.dict(os.environ, {"CHITU_CHUNK_MESSAGES": "110", "CHITU_CHUNK_CHARS": "24000"}, clear=False):
            chunks = llm_analysis.build_chunks(messages)
        self.assertEqual(2, len(chunks))
        self.assertEqual(110, len(chunks[0]))

    def test_legacy_call_signature_still_works(self):
        messages = make_messages([(f"m{i:03d}", f"消息{i}") for i in range(1, 71)])
        chunks = llm_analysis.build_chunks(messages, max_messages=65, max_chars=14000)
        self.assertEqual(2, len(chunks))


class DedupeFanoutIntegrationTests(unittest.TestCase):
    def test_duplicate_texts_share_label_but_all_ids_covered(self):
        specs = [("m1", "怎么使用"), ("m2", "怎么使用"), ("m3", "怎么使用"), ("m4", "要退货"), ("m5", "要退货")]
        messages = make_messages(specs)

        def responder(ids):
            rows = []
            for msg_id in ids:
                label = "使用方法" if msg_id == "m1" else "退款退货"
                rows.append([msg_id, label, "售后", 0.9])
            return rows, [], []

        with tempfile.TemporaryDirectory() as directory, patch.dict(
            os.environ,
            {"CHITU_CLASSIFY_MODEL": "", "CHITU_MESSAGE_CACHE": "1", "CHITU_CHUNK_MESSAGES": "110"},
            clear=False,
        ):
            client = FakeClassifier(directory, responder=responder)
            demand_map, risks, reviews = llm_analysis.classify_chunks(client, messages, TAXONOMY)

        self.assertEqual(1, client.calls)
        self.assertEqual(5, len(demand_map))
        for msg_id in ("m1", "m2", "m3"):
            self.assertEqual("使用方法", demand_map[msg_id][0]["label"])
        for msg_id in ("m4", "m5"):
            self.assertEqual("退款退货", demand_map[msg_id][0]["label"])
        self.assertEqual(5, client.expected_messages)
        self.assertEqual(5, client.analyzed_messages)
        self.assertEqual(2, client.unique_messages)
        self.assertEqual(3, client.deduped_messages)
        self.assertFalse(risks)
        self.assertFalse(reviews)

    def test_hot_memo_rerun_makes_zero_calls(self):
        specs = [("m1", "怎么使用"), ("m2", "怎么使用")]
        messages = make_messages(specs)

        def responder(ids):
            return [[msg_id, "使用方法", "售后", 0.9] for msg_id in ids], [], []

        with tempfile.TemporaryDirectory() as directory, patch.dict(
            os.environ, {"CHITU_CLASSIFY_MODEL": "", "CHITU_MESSAGE_CACHE": "1"}, clear=False
        ):
            client = FakeClassifier(directory, responder=responder)
            llm_analysis.classify_chunks(client, messages, TAXONOMY)
            self.assertEqual(1, client.calls)
            demand_map, _, _ = llm_analysis.classify_chunks(client, messages, TAXONOMY)
            self.assertEqual(1, client.calls)  # 第二次全部命中消息级缓存
            self.assertEqual(2, len(demand_map))


class CascadeIntegrationTests(unittest.TestCase):
    def test_low_confidence_and_signals_escalate_to_main_model(self):
        specs = [
            ("m1", "怎么使用"),   # mini 无需求 + 信号词 -> 升级
            ("m2", "能用吗"),     # mini 低置信度 -> 升级
            ("m3", "好的谢谢"),   # mini 无需求 + 无信号 -> 保留
            ("m4", "怎么退货"),   # mini 高置信度 -> 保留
        ]
        messages = make_messages(specs)

        def screener_responder(ids):
            rows = []
            for msg_id in ids:
                if msg_id == "m2":
                    rows.append([msg_id, "退款退货", "售后", 0.4])
                elif msg_id == "m4":
                    rows.append([msg_id, "退款退货", "售后", 0.95])
            return rows, [], []

        def main_responder(ids):
            return [[msg_id, "使用方法", "售后", 0.95] for msg_id in ids], [], []

        with tempfile.TemporaryDirectory() as directory, patch.dict(
            os.environ,
            {
                "CHITU_CLASSIFY_MODEL": "fake-mini",
                "CHITU_CASCADE_CONFIDENCE": "0.75",
                "CHITU_MESSAGE_CACHE": "1",
            },
            clear=False,
        ):
            main_client = FakeClassifier(directory, model="fake-main", responder=main_responder)
            screener_client = FakeClassifier(directory, model="fake-mini", responder=screener_responder)
            with patch.object(llm_analysis, "build_screener_client", return_value=screener_client):
                demand_map, _, _ = llm_analysis.classify_chunks(main_client, messages, TAXONOMY)

        self.assertEqual(1, main_client.calls)  # 只有 2 条升级，1 批
        self.assertEqual(2, main_client.escalated_messages)
        self.assertEqual("使用方法", demand_map["m1"][0]["label"])   # 升级后以主模型为准
        self.assertEqual("使用方法", demand_map["m2"][0]["label"])   # 低置信度被主模型覆盖
        self.assertNotIn("m3", demand_map)                            # 纯寒暄保留无需求
        self.assertEqual("退款退货", demand_map["m4"][0]["label"])   # mini 高置信度直接采用
        self.assertEqual(1, screener_client.calls)

    def test_hot_cache_replays_triage_without_any_calls(self):
        specs = [("m1", "怎么使用"), ("m2", "好的谢谢")]
        messages = make_messages(specs)

        def screener_responder(ids):
            return [], [], []

        def main_responder(ids):
            return [[msg_id, "使用方法", "售后", 0.95] for msg_id in ids], [], []

        with tempfile.TemporaryDirectory() as directory, patch.dict(
            os.environ, {"CHITU_CLASSIFY_MODEL": "fake-mini", "CHITU_MESSAGE_CACHE": "1"}, clear=False
        ):
            main_client = FakeClassifier(directory, model="fake-main", responder=main_responder)
            screener_client = FakeClassifier(directory, model="fake-mini", responder=screener_responder)
            with patch.object(llm_analysis, "build_screener_client", return_value=screener_client):
                llm_analysis.classify_chunks(main_client, messages, TAXONOMY)
            calls_after_first = main_client.calls
            with patch.object(llm_analysis, "build_screener_client", return_value=screener_client):
                llm_analysis.classify_chunks(main_client, messages, TAXONOMY)
            self.assertEqual(calls_after_first, main_client.calls)
            self.assertEqual(calls_after_first, screener_client.calls)


if __name__ == "__main__":
    unittest.main()
