import unittest

from generate_report import parse_manual_baseline


class ParseManualBaselineTest(unittest.TestCase):
    def test_tsv_rows_with_rank_and_full_columns(self):
        text = "\n".join(
            [
                "售前",
                "1\t版本与套装配置\t56\t26.0%\t8.0\tP1\t购买前比较各版本差异",
                "2\t产品功能与效果\t55\t25.6%\t7.9",
                "售后",
                "1\t使用方法/教程\t79\t34.1%\t11.3\tP0\t已购后咨询操作、教程、视频",
                "2\t退款退货与换货\t32\t13.8%",
            ]
        )
        data = parse_manual_baseline(text)
        self.assertEqual(data["售前"]["rows"]["版本与套装配置"]["count"], 56)
        self.assertAlmostEqual(data["售前"]["rows"]["版本与套装配置"]["share"], 0.26)
        self.assertEqual(data["售前"]["rows"]["版本与套装配置"]["priority"], "P1")
        self.assertIn("版本差异", data["售前"]["rows"]["版本与套装配置"]["definition"])
        self.assertEqual(data["售后"]["rows"]["使用方法/教程"]["count"], 79)
        self.assertEqual(data["售后"]["total"], 111)
        labels = {item["label"]: item for item in data["taxonomy"]}
        self.assertEqual(labels["使用方法/教程"]["stage"], "售后")
        self.assertEqual(labels["使用方法/教程"]["priority"], "P0")

    def test_single_space_rows_without_rank_or_share(self):
        text = "\n".join(
            [
                "售前",
                "热敷 26",
                "导出液/收缩液 24",
                "售后",
                "退款退货 23",
            ]
        )
        data = parse_manual_baseline(text)
        self.assertEqual(data["售前"]["rows"]["热敷"]["count"], 26)
        self.assertEqual(data["售前"]["rows"]["热敷"]["share"], 0.0)
        self.assertEqual(data["售后"]["rows"]["退款退货"]["count"], 23)

    def test_mixed_spacing_with_rank(self):
        # 排名后双空格、列间单空格的混合格式（手工输入常见）
        text = "\n".join(
            [
                "售前",
                "1  热敷 26 10.7%",
                "2  导出液/收缩液 24 9.9%",
                "售后",
                "1  使用方法 59 27.2% P0",
            ]
        )
        data = parse_manual_baseline(text)
        self.assertEqual(data["售前"]["rows"]["热敷"]["count"], 26)
        self.assertAlmostEqual(data["售前"]["rows"]["热敷"]["share"], 0.107)
        self.assertEqual(data["售后"]["rows"]["使用方法"]["count"], 59)
        self.assertEqual(data["售后"]["rows"]["使用方法"]["priority"], "P0")

    def test_meta_lines_days_and_period(self):
        text = "\n".join(
            [
                "统计时间 2026-08-12 至 2026-08-18",
                "售前",
                "热敷 26 10.7%",
                "售后",
                "退款退货 23 10.6%",
            ]
        )
        data = parse_manual_baseline(text)
        self.assertEqual(data["start_date"], "2026-08-12")
        self.assertEqual(data["end_date"], "2026-08-18")
        self.assertEqual(data["售前"]["days"], 7)
        self.assertAlmostEqual(data["售前"]["rows"]["热敷"]["daily"], 26 / 7)

    def test_explicit_days_param_wins(self):
        text = "\n".join(["售前", "热敷 14", "售后", "退款退货 7"])
        data = parse_manual_baseline(text, days=7)
        self.assertEqual(data["售前"]["days"], 7)
        self.assertAlmostEqual(data["售前"]["rows"]["热敷"]["daily"], 2)

    def test_stage_total_lines_are_ignored(self):
        text = "\n".join(
            [
                "售前需求统计",
                "正式需求总数 243 统计天数 7",
                "排名 需求词条 同人去重人数 占比",
                "热敷 26 10.7%",
                "售后需求统计",
                "正式需求总数 217",
                "退款退货 23",
            ]
        )
        data = parse_manual_baseline(text, days=None)
        self.assertEqual(list(data["售前"]["rows"]), ["热敷"])
        self.assertEqual(list(data["售后"]["rows"]), ["退款退货"])
        self.assertEqual(data["售前"]["days"], 7)

    def test_dual_stage_label_merges_in_taxonomy(self):
        text = "\n".join(["售前", "吸头 23", "售后", "吸头 13"])
        data = parse_manual_baseline(text)
        taxonomy = {item["label"]: item for item in data["taxonomy"]}
        self.assertEqual(sorted(taxonomy["吸头"]["stages"]), ["售前", "售后"])

    def test_duplicate_label_raises(self):
        text = "\n".join(["售前", "热敷 26", "热敷 24"])
        with self.assertRaises(RuntimeError) as ctx:
            parse_manual_baseline(text)
        self.assertIn("热敷", str(ctx.exception))

    def test_orphan_row_without_stage_raises(self):
        with self.assertRaises(RuntimeError) as ctx:
            parse_manual_baseline("热敷 26 10.7%")
        self.assertIn("售前/售后", str(ctx.exception))

    def test_no_rows_raises(self):
        with self.assertRaises(RuntimeError):
            parse_manual_baseline("这次没有数据")

    def test_share_sum_over_130_percent_raises(self):
        text = "\n".join(["售前", "热敷 26 80%", "导出液 24 70%"])
        with self.assertRaises(RuntimeError) as ctx:
            parse_manual_baseline(text)
        self.assertIn("130%", str(ctx.exception))

    def test_empty_text_raises(self):
        with self.assertRaises(RuntimeError):
            parse_manual_baseline("   ")

    def test_structure_matches_read_baseline_data_shape(self):
        text = "\n".join(["售前", "热敷 26", "售后", "退款退货 23"])
        data = parse_manual_baseline(text)
        for key in ("售前", "售后", "context", "taxonomy", "products", "start_date", "end_date"):
            self.assertIn(key, data)
        for stage in ("售前", "售后"):
            self.assertEqual(set(data[stage]), {"rows", "total", "days"})


if __name__ == "__main__":
    unittest.main()
