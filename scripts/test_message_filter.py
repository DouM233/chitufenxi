import csv
import sys
import tempfile
import unittest
from pathlib import Path


SCRIPT_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(SCRIPT_DIR))

import generate_report as report


class CustomerMessageFilterTests(unittest.TestCase):
    @staticmethod
    def message(sender, text):
        return {"sender": sender, "text": text}

    def test_buyer_text_does_not_determine_sender_identity(self):
        self.assertTrue(
            report.is_valid_customer_message(
                self.message("真实买家", "我这边换快递不方便")
            )
        )
        self.assertTrue(
            report.is_valid_customer_message(
                self.message("真实买家", "客服说可以换货")
            )
        )

    def test_known_service_senders_are_excluded(self):
        for sender in (
            "jimi_vender_182640958071",
            "现代健康自营沭阳8",
            "联想医疗自营沭阳5",
            "丹橘个护健康沭阳12",
            "苏泊尔个人护理沭阳16",
            "官方客服",
        ):
            with self.subTest(sender=sender):
                self.assertFalse(
                    report.is_valid_customer_message(
                        self.message(sender, "请问有什么可以帮您")
                    )
                )

    def test_system_events_and_rich_cards_are_excluded(self):
        for text in (
            "用户发起转人工",
            "转人工",
            "【此消息为欢迎卡片或富文本模板答案，聊天记录中暂不支持展示】",
            "订单号：123456",
            "您选择了一个订单进行咨询（订单号：123；商品编号：456）",
        ):
            with self.subTest(text=text):
                self.assertFalse(
                    report.is_valid_customer_message(self.message("真实买家", text))
                )

    def test_short_buyer_message_matches_formal_report_policy(self):
        self.assertTrue(
            report.is_valid_customer_message(self.message("真实买家", "行"))
        )

    def test_fulltext_csv_uses_sender_after_message(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "chat_fulltext_123_20260901_20260901.raw.csv"
            with path.open("w", encoding="utf-8-sig", newline="") as handle:
                writer = csv.DictWriter(
                    handle,
                    fieldnames=["buyer", "staff", "chat_date", "filter_range", "chat_text"],
                )
                writer.writeheader()
                writer.writerow(
                    {
                        "buyer": "真实买家",
                        "staff": "品牌旗舰:客服甲",
                        "chat_date": "2026-09-01",
                        "filter_range": "2026-09-01 to 2026-09-01",
                        "chat_text": "\n".join(
                            [
                                "真实买家",
                                "2026-09-01",
                                "09:00:00",
                                "这个怎么使用",
                                "真实买家",
                                "09:00:05",
                                "亲亲，请看使用视频",
                                "品牌旗舰:客服甲",
                                "没有更多内容了",
                            ]
                        ),
                    }
                )

            messages = report.parse_csv("测试商品", path)
            self.assertEqual(2, len(messages))
            self.assertEqual("真实买家", messages[0]["sender"])
            self.assertEqual("buyer", messages[0]["role"])
            self.assertEqual("品牌旗舰:客服甲", messages[1]["sender"])
            self.assertEqual("service", messages[1]["role"])
            self.assertEqual([messages[0]], report.valid_customer_messages(messages))
            self.assertEqual(1, messages[0]["source_record_count"])

    def test_fulltext_csv_excludes_delayed_messages_outside_filter_range(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "chat_fulltext_123_20260901_20260901.raw.csv"
            with path.open("w", encoding="utf-8-sig", newline="") as handle:
                writer = csv.DictWriter(
                    handle,
                    fieldnames=["buyer", "chat_date", "filter_range", "chat_text"],
                )
                writer.writeheader()
                writer.writerow(
                    {
                        "buyer": "真实买家",
                        "chat_date": "2026-09-01",
                        "filter_range": "2026-09-01 to 2026-09-01",
                        "chat_text": "\n".join(
                            [
                                "真实买家",
                                "2026-09-02",
                                "09:00:00",
                                "超出周期的消息",
                                "真实买家",
                                "没有更多内容了",
                            ]
                        ),
                    }
                )

            messages = report.parse_csv("测试商品", path)
            self.assertEqual(1, len(messages))
            self.assertFalse(messages[0]["in_filter_range"])
            self.assertEqual([], report.valid_customer_messages(messages))

    def test_baseline_filename_supplies_product_name(self):
        self.assertEqual(
            "剃毛器",
            report.product_from_baseline_file(
                "剃毛器_20260723-0728_客服聊天需求分析报告(1).xlsx"
            ),
        )


if __name__ == "__main__":
    unittest.main()
