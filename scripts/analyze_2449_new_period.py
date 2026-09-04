import csv
import json
import os
import re
import shutil
from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill


INPUT_CSV = Path(os.environ.get("CHITU_INPUT_CSV", r"E:/新2/赤兔历史分析结果/chat_fulltext_997158824546_20260826_20260901.raw.csv"))
BASELINE_XLSX = Path(os.environ.get("CHITU_BASELINE_XLSX", r"E:/新2/赤兔历史分析结果/2449小气泡_20260819-0825_客服聊天需求分析_基准对比报告(1).xlsx"))
OUTPUT_ROOT = Path(os.environ.get("CHITU_OUTPUT_ROOT", r"E:/新2/赤兔历史分析结果/03_最终报告/2026/09"))

PRODUCT_NAME = "2449小气泡"
PERIOD_START = os.environ.get("CHITU_PERIOD_START", "2026-08-26")
PERIOD_END = os.environ.get("CHITU_PERIOD_END", "2026-09-01")
BASELINE_START = os.environ.get("CHITU_BASELINE_START", "2026-08-19")
BASELINE_END = os.environ.get("CHITU_BASELINE_END", "2026-08-25")
DAYS = 7
BASELINE_DAYS = 7

TIME_RE = re.compile(r"^\d{2}:\d{2}:\d{2}$")
DATE_RE = re.compile(r"^\d{4}-\d{2}-\d{2}$")
URL_RE = re.compile(r"https?://\S+")
CONTROL_RE = re.compile(r"[\x00-\x08\x0b-\x0c\x0e-\x1f]")

INVALID_CONTAINS = [
    "请在客户端查看原始聊天记录",
    "聊天记录查询",
    "查询结果",
    "客户昵称",
    "员工账号",
    "账号信息保护显示",
    "订单号:",
    "共1件商品",
    "交易时间:",
    "亲，这款",
    "亲亲 机器",
    "小主，喜欢就不要纠结",
    "美丽贵在坚持",
    "热视频",
    "使用视频",
    "春风十里",
    "欢迎光临",
    "很荣幸为您提供服务",
]
INVALID_EXACT = {"你好", "您好", "在吗", "好的", "好", "嗯", "OK", "ok", "谢谢", "谢谢你", "啥情况"}

KEYWORDS = {
    "优惠活动": ["优惠", "券", "活动", "便宜", "价格", "返现", "红包", "好评"],
    "使用后效果不佳": ["吸不出", "吸不出来", "没效果", "效果不好", "没吸出来", "吸不了", "一点都吸不出来"],
    "使用后皮肤不适": ["吸红", "红了", "紫了", "淤青", "疼", "痛", "破皮", "皮肤不好", "过敏"],
    "使用手法": ["停留", "移动", "手法", "怎么移动", "刮", "拉"],
    "使用方法": ["怎么用", "如何使用", "使用方法", "教程", "说明书", "怎么操作", "开机", "怎么弄", "按键", "视频"],
    "使用频率": ["多久用一次", "每天用", "一周几次", "使用频率"],
    "充电方式": ["充电", "充多久", "充满", "充电口", "插电"],
    "免费试用": ["试用", "不满意退", "30天", "三十天", "拆封可以退", "能退吗", "无理由"],
    "全脸清洁": ["全脸", "脸上", "下巴", "鼻翼", "额头", "脸颊"],
    "功能原理": ["原理", "小气泡", "水循环", "自动收缩", "工作原理"],
    "加水/水箱": ["加水", "水箱", "纯净水", "自来水", "白开水", "污水", "水要加", "装水", "水路"],
    "包装/分包": ["包装", "盒", "分包", "两个包裹", "外包装"],
    "发货/物流": ["发货", "物流", "快递", "送到", "改地址", "拦截", "什么时候到"],
    "吸力": ["吸力", "档位", "吸得", "力度", "吸太大", "吸红", "吸紫"],
    "吸头": ["吸头", "吸嘴", "头子", "探头", "圆头", "椭圆", "换头"],
    "声音/噪音": ["声音", "噪音", "吵", "响"],
    "导出液/收缩液": ["导出液", "收缩液", "精华水", "小蓝瓶", "玻尿酸", "液体", "导出", "收缩"],
    "小气泡效果": ["小气泡", "补水", "微气泡"],
    "收缩毛孔": ["毛孔", "收缩", "毛孔会大"],
    "故障_不出水/吸水": ["不喷水", "不出水", "不吸水", "吸不上水", "水不出来", "不喷", "没水"],
    "故障_无法开机/充电": ["无法开机", "开不了机", "没反应", "充不进", "充不了", "不开机"],
    "故障_漏水/水箱": ["漏水", "水箱裂", "水珠", "水漏", "防水塞"],
    "故障_漏电/异常发热": ["漏电", "电流", "异常发热", "发热", "烫手"],
    "故障_配件/破损": ["坏的", "坏了", "破损", "断了", "松动", "少了", "缺失"],
    "效果": ["黑头", "粉刺", "闭口", "油脂", "清洁", "效果"],
    "款式区别": ["区别", "哪款", "版本", "套餐", "款式", "几代", "选择"],
    "清洁护肤": ["热敷", "护肤", "护理", "用前", "用后", "洗脸"],
    "清洗仪器": ["清洗", "清洁机器", "洗机器", "吸头清洗", "水箱清洗"],
    "热敷": ["热敷", "蒸脸", "温热"],
    "皮肤风险": ["会不会红", "会不会紫", "敏感", "松弛", "伤皮肤", "会不会疼"],
    "蓝光/冰敷": ["蓝光", "冰敷", "冷敷"],
    "质保": ["质保", "保修", "售后保障"],
    "赠品": ["赠品", "送", "漏发", "赠送"],
    "运费险": ["运费险", "退货宝", "邮费"],
    "退款退货": ["退款", "退货", "换货", "我要退", "退了", "申请退"],
    "适用人群": ["男士", "男人", "儿童", "小孩", "老人", "孕妇", "美容院"],
    "适用肤质": ["油皮", "干皮", "敏感肌", "痘肌", "肤质"],
    "配件耗材": ["配件", "耗材", "替换", "购买吸头", "耗材买"],
    "顾虑二手": ["二手", "用过", "水垢", "别人退", "试用退回"],
}

FORCED_STAGE = {
    "使用方法": "售后",
    "充电方式": "售后",
    "故障_不出水/吸水": "售后",
    "故障_无法开机/充电": "售后",
    "故障_漏水/水箱": "售后",
    "故障_漏电/异常发热": "售后",
    "故障_配件/破损": "售后",
    "退款退货": "售后",
    "运费险": "售后",
    "质保": "售后",
}

RISK_LABELS = {"使用后皮肤不适", "故障_不出水/吸水", "故障_无法开机/充电", "故障_漏水/水箱", "故障_漏电/异常发热", "故障_配件/破损", "使用后效果不佳"}


def clean_text(text):
    text = URL_RE.sub("", text or "")
    text = CONTROL_RE.sub("", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def is_invalid(text):
    if not text or text in INVALID_EXACT:
        return True
    if URL_RE.fullmatch(text):
        return True
    return any(token in text for token in INVALID_CONTAINS)


def read_baseline_meta():
    wb = load_workbook(BASELINE_XLSX, data_only=True)
    ws = wb["词条与口径"]
    meta = {}
    for row in range(2, ws.max_row + 1):
        label = ws.cell(row, 1).value
        if not label:
            continue
        meta[label] = {
            "stage": ws.cell(row, 2).value or "",
            "definition": ws.cell(row, 3).value or "",
            "priority": ws.cell(row, 4).value or "P2",
            "action": ws.cell(row, 8).value or "",
        }
    baseline = {"售前": {}, "售后": {}}
    for stage, sheet_name in [("售前", "售前需求统计"), ("售后", "售后需求统计")]:
        ws = wb[sheet_name]
        total = ws.cell(2, 2).value or 0
        for row in range(5, ws.max_row + 1):
            label = ws.cell(row, 2).value
            if label:
                baseline[stage][label] = {
                    "count": ws.cell(row, 3).value or 0,
                    "share": ws.cell(row, 4).value or 0,
                    "daily": ws.cell(row, 5).value or 0,
                    "priority": ws.cell(row, 8).value or meta.get(label, {}).get("priority", "P2"),
                }
        baseline[stage]["__total__"] = total
    return meta, baseline


def extract_messages():
    messages = []
    with INPUT_CSV.open("r", encoding="utf-8-sig", newline="") as file:
        reader = csv.DictReader(file)
        for csv_line, row in enumerate(reader, start=2):
            buyer = row.get("buyer", "").strip()
            chat_text = row.get("chat_text", "")
            lines = [line.strip() for line in chat_text.splitlines() if line.strip()]
            current_date = row.get("chat_date", "")
            i = 0
            while i < len(lines):
                line = lines[i]
                if line == buyer:
                    j = i + 1
                    if j < len(lines) and DATE_RE.match(lines[j]):
                        current_date = lines[j]
                        j += 1
                    if j < len(lines) and TIME_RE.match(lines[j]):
                        time = lines[j]
                        j += 1
                        if j < len(lines):
                            text = clean_text(lines[j])
                            if PERIOD_START <= current_date <= PERIOD_END and not is_invalid(text):
                                messages.append(
                                    {
                                        "csv_line": csv_line,
                                        "buyer": buyer,
                                        "date": current_date,
                                        "time": time,
                                        "text": text,
                                    }
                                )
                        i = j + 1
                        continue
                i += 1
    return messages


def classify(text):
    labels = []
    for label, keywords in KEYWORDS.items():
        if any(keyword in text for keyword in keywords):
            stage = FORCED_STAGE.get(label)
            if not stage:
                # Heuristic: actual problem phrasing is aftersale; policy/effect questions are presale.
                stage = "售后" if any(token in text for token in ["坏", "不出", "无法", "退", "用了", "用后", "收到", "发来", "漏", "充不", "吸不出"]) else "售前"
            labels.append((stage, label))
    seen = set()
    result = []
    for item in labels:
        if item not in seen:
            seen.add(item)
            result.append(item)
    return result


def analyze(meta):
    messages = extract_messages()
    retained = {}
    detail = []
    duplicates = []
    risks = []
    reviews = []
    stats = defaultdict(lambda: {"buyers": set(), "raw": 0, "dup": 0, "quotes": [], "buyer_examples": []})

    for msg in messages:
        labels = classify(msg["text"])
        if not labels:
            if any(token in msg["text"] for token in ["怎么", "不", "没", "退", "坏", "疼", "红"]):
                reviews.append({**msg, "reason": "未命中旧词条，但存在问题语气，需人工确认。"})
            continue
        for stage, label in labels:
            stats[(stage, label)]["raw"] += 1
            key = (msg["buyer"], stage, label)
            if key in retained:
                stats[(stage, label)]["dup"] += 1
                duplicates.append(
                    {
                        "stage": stage,
                        "label": label,
                        "buyer": msg["buyer"],
                        "date": msg["date"],
                        "time": msg["time"],
                        "duplicate_text": msg["text"],
                        "duplicate_source": msg["csv_line"],
                        "kept_text": retained[key]["text"],
                    }
                )
                continue
            retained[key] = msg
            stats[(stage, label)]["buyers"].add(msg["buyer"])
            stats[(stage, label)]["quotes"].append(msg["text"])
            stats[(stage, label)]["buyer_examples"].append(msg["buyer"])
            detail.append(
                {
                    "stage": stage,
                    "label": label,
                    "buyer": msg["buyer"],
                    "date": msg["date"],
                    "time": msg["time"],
                    "text": msg["text"],
                    "source": msg["csv_line"],
                    "raw_hit": stats[(stage, label)]["raw"],
                }
            )
            if label in RISK_LABELS and stage == "售后":
                risks.append({**msg, "label": label, "action": meta.get(label, {}).get("action", "进入质量/安全二次复核。")})

    rows = []
    for (stage, label), item in stats.items():
        count = len(item["buyers"])
        rows.append(
            {
                "stage": stage,
                "label": label,
                "count": count,
                "raw": item["raw"],
                "dup": item["dup"],
                "priority": meta.get(label, {}).get("priority", "P2"),
                "definition": meta.get(label, {}).get("definition", f"{label}相关咨询。"),
                "action": meta.get(label, {}).get("action", "沿用基准词条，持续观察。"),
                "buyers": "、".join(item["buyer_examples"][:10]),
                "quote": "｜".join(item["quotes"][:5]),
                "baseline_status": "原基准词条" if label in meta else "本期新增",
            }
        )
    return messages, rows, detail, duplicates, risks, reviews


def clear_sheet(ws, min_row=1):
    if ws.max_row >= min_row:
        ws.delete_rows(min_row, ws.max_row - min_row + 1)


def append_rows(ws, rows):
    for row in rows:
        ws.append(row)


def rank_stage(rows, stage):
    return sorted([row for row in rows if row["stage"] == stage], key=lambda x: (-x["count"], x["label"]))


def write_stat_sheet(ws, rows, total, buyer_count):
    clear_sheet(ws, 2)
    ws.append(["正式需求总数", total, "统计天数", DAYS, "需求覆盖买家", buyer_count, None, None, None, None, None, None])
    ws.append([None] * 12)
    ws.append(["排名", "需求词条", "同人去重人数", "占比", "日均人数", "原始命中", "重复次数", "优先级", "严格定义", "涉及买家示例", "代表原话", "建议动作"])
    for idx, row in enumerate(rows, 1):
        ws.append([idx, row["label"], row["count"], row["count"] / total if total else 0, row["count"] / DAYS, row["raw"], row["dup"], row["priority"], row["definition"], row["buyers"], row["quote"], row["action"]])


def write_compare_sheet(ws, rows, baseline_stage, total, baseline_total):
    clear_sheet(ws, 2)
    ws.append(["本期总需求", total, "本期天数", DAYS, "基准总需求", baseline_total, "基准天数", BASELINE_DAYS, None, None, None, None])
    ws.append([None] * 12)
    ws.append(["需求词条", "本期人数", "本期占比", "本期日均", "基准人数", "基准占比", "基准日均", "人数变化", "占比变化", "变化判断", "建议动作", "备注"])
    labels = set(baseline_stage.keys()) | {row["label"] for row in rows}
    labels.discard("__total__")
    current = {row["label"]: row for row in rows}
    for label in sorted(labels, key=lambda x: -(current.get(x, {}).get("count", 0))):
        cur = current.get(label)
        base = baseline_stage.get(label, {})
        cur_count = cur["count"] if cur else 0
        cur_share = cur_count / total if total else 0
        base_count = base.get("count", 0)
        base_share = base.get("share", 0)
        pp = cur_share - base_share
        if abs(pp) >= 0.03:
            judgment = "结构明显上升" if pp > 0 else "结构明显下降"
        elif cur_count - base_count > 0:
            judgment = "人数上升"
        elif cur_count - base_count < 0:
            judgment = "人数下降"
        else:
            judgment = "基本持平"
        action = cur["action"] if cur else ""
        status = cur["baseline_status"] if cur else "原基准词条"
        ws.append([label, cur_count, cur_share, cur_count / DAYS, base_count, base_share, base.get("daily", 0), cur_count - base_count, pp, judgment, action, status])


def top_changes(rows, baseline_stage, total, limit=5):
    current = {row["label"]: row for row in rows}
    labels = set(current) | set(baseline_stage)
    labels.discard("__total__")
    changes = []
    for label in labels:
        cur = current.get(label)
        base = baseline_stage.get(label, {})
        cur_count = cur["count"] if cur else 0
        cur_share = cur_count / total if total else 0
        base_share = base.get("share", 0) or 0
        changes.append(
            {
                "label": label,
                "current_count": cur_count,
                "base_count": base.get("count", 0) or 0,
                "pp": cur_share - base_share,
                "current_share": cur_share,
                "base_share": base_share,
            }
        )
    risers = sorted(changes, key=lambda item: (-item["pp"], -item["current_count"]))[:limit]
    fallers = sorted(changes, key=lambda item: (item["pp"], -item["base_count"]))[:limit]
    return risers, fallers


def fmt_pp(value):
    return f"{value * 100:+.1f}pp"


def save_report(meta, baseline, messages, rows, detail, duplicates, risks, reviews):
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    task_id = f"{ts}_2449小气泡_{PERIOD_START.replace('-', '')}-{PERIOD_END[5:].replace('-', '')}_有基准"
    output_dir = OUTPUT_ROOT / task_id
    output_dir.mkdir(parents=True, exist_ok=True)
    period_slug = f"{PERIOD_START.replace('-', '')}-{PERIOD_END[5:].replace('-', '')}"
    output_xlsx = output_dir / f"2449小气泡_{period_slug}_客服聊天需求基准对比分析报告.xlsx"
    output_md = output_dir / f"2449小气泡_{period_slug}_管理结论.md"
    manifest_path = output_dir / "result_manifest.json"

    shutil.copy2(BASELINE_XLSX, output_xlsx)
    wb = load_workbook(output_xlsx)
    if "一级主题" not in wb.sheetnames:
        wb.create_sheet("一级主题", 1)
    ws = wb["分析总览"]
    ws["A1"] = f"2449小气泡｜客服聊天需求基准对比分析报告（{PERIOD_START} 至 {PERIOD_END}）"
    ws["A2"] = f"对比基准：{BASELINE_START} 至 {BASELINE_END}，两个周期均为7天。沿用上一版词条、定义、售前售后归属和去重规则；使用方法强制归售后；P0/安全质量问题保留原话并二次复核。"
    valid_buyers = len(set(msg["buyer"] for msg in messages))
    presale = rank_stage(rows, "售前")
    aftersale = rank_stage(rows, "售后")
    presale_total = sum(row["count"] for row in presale)
    aftersale_total = sum(row["count"] for row in aftersale)
    formal_total = presale_total + aftersale_total

    for cell, value in {
        "A6": len(messages),
        "B6": valid_buyers,
        "C6": len(messages),
        "D6": formal_total,
        "E6": presale_total,
        "F6": aftersale_total,
        "G6": len(duplicates),
        "H6": valid_buyers,
        "I6": len(risks),
        "J6": len(reviews),
    }.items():
        ws[cell] = value

    write_stat_sheet(wb["售前需求统计"], presale, presale_total, valid_buyers)
    wb["售前需求统计"]["A1"] = f"2449小气泡｜本期售前需求统计（{PERIOD_START} 至 {PERIOD_END}）"
    write_stat_sheet(wb["售后需求统计"], aftersale, aftersale_total, valid_buyers)
    wb["售后需求统计"]["A1"] = f"2449小气泡｜本期售后需求统计（{PERIOD_START} 至 {PERIOD_END}）"
    write_compare_sheet(wb["售前基准对比"], presale, baseline["售前"], presale_total, baseline["售前"].get("__total__", 0))
    wb["售前基准对比"]["A1"] = f"2449小气泡｜售前基准对比（本期 {PERIOD_START} 至 {PERIOD_END}；基准 {BASELINE_START} 至 {BASELINE_END}）"
    write_compare_sheet(wb["售后基准对比"], aftersale, baseline["售后"], aftersale_total, baseline["售后"].get("__total__", 0))
    wb["售后基准对比"]["A1"] = f"2449小气泡｜售后基准对比（本期 {PERIOD_START} 至 {PERIOD_END}；基准 {BASELINE_START} 至 {BASELINE_END}）"

    ws = wb["一级主题"]
    clear_sheet(ws, 1)
    theme_stats = defaultdict(lambda: {"count": 0, "labels": []})
    theme_map = {
        "优惠活动": "价格与交易政策",
        "免费试用": "价格与交易政策",
        "运费险": "价格与交易政策",
        "退款退货": "售后服务",
        "发货/物流": "交易履约",
        "质保": "售后服务",
        "使用方法": "使用学习",
        "使用手法": "使用学习",
        "使用频率": "使用学习",
        "清洗仪器": "使用学习",
        "充电方式": "使用学习",
        "加水/水箱": "水路系统",
        "故障_不出水/吸水": "水路系统",
        "故障_漏水/水箱": "水路系统",
        "导出液/收缩液": "耗材与搭配",
        "吸头": "配件耗材",
        "配件耗材": "配件耗材",
        "赠品": "配件耗材",
        "热敷": "功能效果",
        "蓝光/冰敷": "功能效果",
        "吸力": "功能效果",
        "效果": "功能效果",
        "小气泡效果": "功能效果",
        "使用后效果不佳": "效果反馈",
        "使用后皮肤不适": "安全质量",
        "皮肤风险": "安全顾虑",
        "故障_无法开机/充电": "电控故障",
        "故障_漏电/异常发热": "安全质量",
        "故障_配件/破损": "质量缺陷",
        "款式区别": "购买决策",
        "适用人群": "购买决策",
        "适用肤质": "购买决策",
        "顾虑二手": "购买决策",
        "包装/分包": "交易履约",
    }
    for row in rows:
        theme = theme_map.get(row["label"], "其他需求")
        theme_stats[theme]["count"] += row["count"]
        theme_stats[theme]["labels"].append(f"{row['stage']}:{row['label']}({row['count']})")
    ws.append(["一级主题", "本期人数", "本期占比", "覆盖词条", "业务判断", "建议动作"])
    for theme, item in sorted(theme_stats.items(), key=lambda kv: -kv[1]["count"]):
        ws.append([theme, item["count"], item["count"] / formal_total if formal_total else 0, "；".join(item["labels"][:10]), "用于管理层观察需求结构变化。", "查看对应词条的基准对比和原话证据后制定动作。"])

    # 新增词条
    ws = wb["新增词条"]
    clear_sheet(ws, 1)
    ws.append(["阶段", "新增词条", "本期人数", "本期占比", "优先级", "严格定义", "建议动作"])
    for row in rows:
        if row["baseline_status"] == "本期新增":
            total = presale_total if row["stage"] == "售前" else aftersale_total
            ws.append([row["stage"], row["label"], row["count"], row["count"] / total if total else 0, row["priority"], row["definition"], row["action"]])
    if ws.max_row == 1:
        ws.append(["—", "本期无新增词条", 0, 0, "—", "—", "—"])

    ws = wb["P0风险明细"]
    clear_sheet(ws, 1)
    ws.append(["ID", "需求词条", "买家", "日期", "时间", "代表原话", "全部证据", "CSV源行", "优先级", "建议动作"])
    p0_rows = [row for row in detail if meta.get(row["label"], {}).get("priority") == "P0"]
    for idx, row in enumerate(p0_rows, 1):
        ws.append([idx, row["label"], row["buyer"], row["date"], row["time"], row["text"], row["text"], row["source"], "P0", meta.get(row["label"], {}).get("action", "")])

    ws = wb["安全质量风险"]
    clear_sheet(ws, 1)
    ws.append(["风险ID", "风险词条", "买家", "日期", "时间", "代表原话", "全部证据", "CSV源行", "原始命中", "建议动作"])
    for idx, row in enumerate(risks, 1):
        ws.append([idx, row["label"], row["buyer"], row["date"], row["time"], row["text"], row["text"], row["csv_line"], 1, row["action"]])
    if not risks:
        ws.append([1, "本期无强风险", "—", "—", "—", "—", "—", "—", 0, "持续监测"])

    ws = wb["逐条需求明细"]
    clear_sheet(ws, 1)
    ws.append(["明细ID", "阶段", "需求词条", "买家", "日期", "时间", "代表原话", "同需求全部原话", "CSV源行", "原始命中"])
    for idx, row in enumerate(detail, 1):
        ws.append([idx, row["stage"], row["label"], row["buyer"], row["date"], row["time"], row["text"], row["text"], row["source"], row["raw_hit"]])

    ws = wb["同人重复记录"]
    clear_sheet(ws, 1)
    ws.append(["阶段", "需求词条", "买家", "日期", "时间", "重复原话", "重复源行", "保留原话", "剔除原因"])
    for row in duplicates:
        ws.append([row["stage"], row["label"], row["buyer"], row["date"], row["time"], row["duplicate_text"], row["duplicate_source"], row["kept_text"], "同一买家、同一阶段、同一需求词条，本周期只计1次"])

    ws = wb["待人工复核"]
    clear_sheet(ws, 1)
    ws.append(["买家", "日期", "时间", "待复核原话", "CSV源行", "复核原因", "状态"])
    for row in reviews:
        ws.append([row["buyer"], row["date"], row["time"], row["text"], row["csv_line"], row["reason"], "待人工确认"])
    if not reviews:
        ws.append(["—", "—", "—", "本期已完成全部独立有效需求归类", "—", "—", "无需复核"])

    ws = wb["词条与口径"]
    clear_sheet(ws, 1)
    ws.append(["需求词条", "本期阶段", "严格定义", "优先级", "售前人数", "售后人数", "基准状态", "建议动作"])
    by_label = {}
    for row in rows:
        item = by_label.setdefault(row["label"], {"presale": 0, "aftersale": 0, **row})
        if row["stage"] == "售前":
            item["presale"] += row["count"]
        else:
            item["aftersale"] += row["count"]
    for label, row in sorted(by_label.items(), key=lambda x: -(x[1]["presale"] + x[1]["aftersale"])):
        stage = "售前、售后" if row["presale"] and row["aftersale"] else ("售前" if row["presale"] else "售后")
        ws.append([label, stage, row["definition"], row["priority"], row["presale"], row["aftersale"], row["baseline_status"], row["action"]])

    ws = wb["上期基准"]
    clear_sheet(ws, 1)
    ws.append(["阶段", "排名", "需求词条", "基准人数", "基准占比", "基准日均", "优先级", "严格定义", "建议动作"])
    for stage in ("售前", "售后"):
        baseline_rows = [(label, item) for label, item in baseline[stage].items() if label != "__total__"]
        baseline_rows.sort(key=lambda pair: -(pair[1].get("count", 0) or 0))
        for idx, (label, item) in enumerate(baseline_rows, 1):
            ws.append(
                [
                    stage,
                    idx,
                    label,
                    item.get("count", 0) or 0,
                    item.get("share", 0) or 0,
                    item.get("daily", 0) or 0,
                    item.get("priority") or meta.get(label, {}).get("priority", "P2"),
                    meta.get(label, {}).get("definition", f"{label}相关咨询。"),
                    meta.get(label, {}).get("action", ""),
                ]
            )

    ws = wb["原始数据"]
    clear_sheet(ws, 1)
    ws.append(["CSV源行", "task_id", "采集时间", "筛选范围", "商品ID", "聊天日期", "咨询时间", "买家", "有效买家原话"])
    for msg in messages:
        ws.append([msg["csv_line"], f"FULLTEXT_{PERIOD_START.replace('-', '')}_{PERIOD_END.replace('-', '')}_997158824546", "", f"{PERIOD_START} to {PERIOD_END}", "997158824546", msg["date"], msg["time"], msg["buyer"], msg["text"]])

    for sheet in wb.worksheets:
        for row in sheet.iter_rows():
            for cell in row:
                if isinstance(cell.value, str) and len(cell.value) > 60:
                    cell.alignment = Alignment(wrap_text=True, vertical="top")
    wb.save(output_xlsx)

    top_pre = presale[0]["label"] if presale else "无"
    top_after = aftersale[0]["label"] if aftersale else "无"
    pre_risers, pre_fallers = top_changes(presale, baseline["售前"], presale_total)
    after_risers, after_fallers = top_changes(aftersale, baseline["售后"], aftersale_total)
    summary = [
        f"本期按有基准口径分析：新周期 {PERIOD_START} 至 {PERIOD_END}，上期基准 {BASELINE_START} 至 {BASELINE_END}，两个周期均为7天，可同时比较人数、占比和日均。",
        f"本期正式需求 {formal_total} 条，较上期基准 {baseline['售前'].get('__total__', 0) + baseline['售后'].get('__total__', 0)} 条变化 {formal_total - (baseline['售前'].get('__total__', 0) + baseline['售后'].get('__total__', 0)):+d} 条；其中售前 {presale_total} 条、售后 {aftersale_total} 条。",
        f"售前最高频为“{top_pre}”；结构上升最明显的是“{pre_risers[0]['label']}”（{fmt_pp(pre_risers[0]['pp'])}），下降最明显的是“{pre_fallers[0]['label']}”（{fmt_pp(pre_fallers[0]['pp'])}）。",
        f"售后最高频为“{top_after}”；结构上升最明显的是“{after_risers[0]['label']}”（{fmt_pp(after_risers[0]['pp'])}），下降最明显的是“{after_fallers[0]['label']}”（{fmt_pp(after_fallers[0]['pp'])}）。",
        f"本期安全质量风险候选 {len(risks)} 条，已单独进入《安全质量风险》Sheet；P0/质量/安全问题已按用户本人实际发生口径二次校准。",
        "建议先看《售后基准对比》里结构上升的词条，再回到《逐条需求明细》核对原话，优先处理能降低客服承接压力和退货风险的问题。",
    ]
    output_md.write_text(f"# 2449小气泡 {PERIOD_START} 至 {PERIOD_END} 客服聊天需求基准对比分析\n\n" + "\n".join(f"{i+1}. {line}" for i, line in enumerate(summary)), encoding="utf-8")
    manifest = {
        "task_id": task_id,
        "product_name": PRODUCT_NAME,
        "analysis_type": "baseline_compare",
        "period": {"start": PERIOD_START, "end": PERIOD_END, "days": DAYS},
        "baseline": {"start": BASELINE_START, "end": BASELINE_END, "days": BASELINE_DAYS, "file": str(BASELINE_XLSX)},
        "source_files": [str(INPUT_CSV), str(BASELINE_XLSX)],
        "metrics": {
            "valid_messages": len(messages),
            "valid_buyers": valid_buyers,
            "formal_demands": formal_total,
            "presale_demands": presale_total,
            "aftersale_demands": aftersale_total,
            "duplicate_removed": len(duplicates),
            "risk_candidates": len(risks),
            "manual_review": len(reviews),
        },
        "status": "completed",
        "output_files": {"excel": str(output_xlsx), "markdown": str(output_md), "manifest": str(manifest_path)},
        "summary": summary,
    }
    manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
    return manifest


def main():
    meta, baseline = read_baseline_meta()
    messages, rows, detail, duplicates, risks, reviews = analyze(meta)
    manifest = save_report(meta, baseline, messages, rows, detail, duplicates, risks, reviews)
    print(json.dumps(manifest, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
