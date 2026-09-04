import os
import re
import sys
import tempfile
import unittest
from pathlib import Path
from threading import Lock
from unittest.mock import patch


SCRIPT_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(SCRIPT_DIR))

import llm_analysis


class CompletenessContractTests(unittest.TestCase):
    def test_formal_demand_prompt_uses_strict_evidence_gate(self):
        prompt = llm_analysis.CLASSIFICATION_SYSTEM_PROMPT
        self.assertIn("当前买家原话必须包含可独立统计", prompt)
        self.assertIn("上下文只用于恢复代词", prompt)
        self.assertIn("没有明确已购、收货、本人实操或实际异常证据，必须归售前", prompt)

    def test_dual_stage_feature_question_defaults_to_presale_without_aftersale_evidence(self):
        messages = [
            {
                "_llm_id": "m1",
                "text": "下面是倒纯净水吗",
                "context": "当前买家: 下面是倒纯净水吗\n客服: 可以加纯净水",
            },
            {
                "_llm_id": "m2",
                "text": "我收到后按说明加水还是不出水",
                "context": "当前买家: 我收到后按说明加水还是不出水",
            },
        ]
        taxonomy = [
            {
                "label": "加水/水箱",
                "stage": "售前",
                "stages": ["售前", "售后"],
            }
        ]
        demands = {
            "m1": [{"label": "加水/水箱", "stage": "售后", "confidence": 0.98}],
            "m2": [{"label": "加水/水箱", "stage": "售后", "confidence": 0.98}],
        }

        calibrated = llm_analysis.calibrate_demand_stages(messages, taxonomy, demands)

        self.assertEqual("售前", calibrated["m1"][0]["stage"])
        self.assertEqual("售后", calibrated["m2"][0]["stage"])

    def test_specific_tutorial_stays_aftersale_while_ambiguous_feature_moves_presale(self):
        messages = [
            {
                "_llm_id": "m1",
                "text": "蒸脸器怎么用",
                "context": "当前买家: 蒸脸器怎么用",
            }
        ]
        taxonomy = [
            {"label": "热敷", "stage": "售前", "stages": ["售前", "售后"]},
            {"label": "使用方法", "stage": "售后", "stages": ["售后"]},
        ]
        demands = {
            "m1": [
                {"label": "热敷", "stage": "售后", "confidence": 0.95},
                {"label": "使用方法", "stage": "售后", "confidence": 0.99},
            ]
        }

        calibrated = llm_analysis.calibrate_demand_stages(messages, taxonomy, demands)

        self.assertEqual(
            [("热敷", "售前"), ("使用方法", "售后")],
            [(item["label"], item["stage"]) for item in calibrated["m1"]],
        )

    def test_context_uses_message_identity_when_csv_messages_share_source_line(self):
        first = {
            "conversation_id": 1,
            "source_line": 8,
            "role": "buyer",
            "sender": "买家A",
            "text": "吸头有什么区别",
        }
        service = {
            "conversation_id": 1,
            "source_line": 8,
            "role": "service",
            "sender": "客服A",
            "text": "您收到后可以看说明书",
        }
        second = {
            "conversation_id": 1,
            "source_line": 8,
            "role": "buyer",
            "sender": "买家A",
            "text": "需要加什么水",
        }
        valid_first = dict(first)
        valid_second = dict(second)
        entries = [{"raw": [first, service, second], "valid": [valid_first, valid_second]}]

        llm_analysis.add_conversation_context(entries)

        self.assertIn("当前买家: 吸头有什么区别", valid_first["context"])
        self.assertIn("客服: 您收到后可以看说明书", valid_first["context"])
        self.assertIn("同买家: 需要加什么水", valid_first["context"])
        self.assertNotIn("当前买家: 您收到后可以看说明书", valid_first["context"])
        self.assertIn("当前买家: 需要加什么水", valid_second["context"])

    def test_batch_requires_every_message_id(self):
        result = {
            "analyzed_ids": ["m1", "m2"],
            "demands": [],
            "risks": [],
            "reviews": [],
        }
        self.assertIs(result, llm_analysis.validate_complete_batch(result, ["m1", "m2"]))

        with self.assertRaises(llm_analysis.LLMAnalysisError) as context:
            llm_analysis.validate_complete_batch(result, ["m1", "m2", "m3"])
        self.assertTrue(context.exception.retryable)

    def test_unlimited_attempts_are_the_default_complete_mode(self):
        with tempfile.TemporaryDirectory() as directory, patch.dict(
            os.environ,
            {
                "CHITU_LLM_API_BASE": "https://example.invalid",
                "CHITU_LLM_API_KEY": "test-key",
                "CHITU_LLM_CACHE_DIR": directory,
                "CHITU_LLM_MAX_ATTEMPTS": "0",
            },
            clear=False,
        ):
            client = llm_analysis.OpenAICompatibleClient()
        self.assertIsNone(client.max_attempts)

    def test_retryable_large_batch_is_split_without_degradation(self):
        class SplitOnceClient:
            def __init__(self, cache_dir):
                self.cache_dir = Path(cache_dir)
                self.model = "test-model"
                self.max_attempts = None
                self.attempts = 0
                self.calls = 0
                self.cache_hits = 0
                self.degraded_messages = 0
                self.expected_messages = 0
                self.analyzed_messages = 0
                self._usage_lock = Lock()

            def has_cached(self, *_args):
                return False

            def invalidate_cache(self, *_args):
                return None

            def complete_json(self, _system_prompt, user_prompt, max_tokens=10000):
                del max_tokens
                self.attempts += 1
                self.calls += 1
                if self.calls == 1:
                    raise llm_analysis.LLMAnalysisError("simulated timeout", retryable=True)
                ids = re.findall(r'"id":\s*"(m\d+)"', user_prompt)
                return {
                    "analyzed_ids": list(dict.fromkeys(ids)),
                    "demands": [],
                    "risks": [],
                    "reviews": [],
                }

        messages = [
            {
                "_llm_id": f"m{index:03d}",
                "product": "测试产品",
                "sender": f"买家{index}",
                "text": "收到",
                "context": "当前买家: 收到",
            }
            for index in range(1, 21)
        ]
        taxonomy = [
            {
                "label": "使用方法",
                "stage": "售后",
                "stages": ["售后"],
                "product_prefixes": [],
                "theme": "使用",
                "definition": "询问产品使用方法。",
            }
        ]
        with tempfile.TemporaryDirectory() as directory:
            client = SplitOnceClient(directory)
            demands, risks, reviews = llm_analysis.classify_chunks(client, messages, taxonomy)

        self.assertEqual(20, client.expected_messages)
        self.assertEqual(20, client.analyzed_messages)
        self.assertEqual(0, client.degraded_messages)
        self.assertEqual(3, client.calls)
        self.assertFalse(demands)
        self.assertFalse(risks)
        self.assertFalse(reviews)


if __name__ == "__main__":
    unittest.main()
