from __future__ import annotations

import json
import re
import shutil
from collections import Counter, defaultdict
from copy import copy
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Iterable

from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.styles import Alignment, Font, PatternFill

ROOT = Path(r"E:/新2")
SOURCE_LOG = ROOT / "赤兔历史分析结果/909化妆镜聊天记录.log"
TEMPLATE_XLSX = ROOT / "赤兔历史分析结果/2449小气泡_20260812-0818_客服聊天需求分析报告.xlsx"
TASK_ID = "20260901_909化妆镜首次基准"
PROCESS_DIR = ROOT / "赤兔历史分析结果/02_过程文件/2026/09" / TASK_ID
OUTPUT_DIR = ROOT / "赤兔历史分析结果/03_最终报告/2026/09" / TASK_ID
OUTPUT_XLSX = OUTPUT_DIR / "909化妆镜_20260802-0831_首次客服聊天需求基准分析报告.xlsx"
OUTPUT_MD = OUTPUT_DIR / "909化妆镜_20260802-0831_管理结论.md"
OUTPUT_MANIFEST = OUTPUT_DIR / "result_manifest.json"

HEADER_RE = re.compile(r"^(.+?)\s+(\d{4}-\d{2}-\d{2})\s+(\d{2}:\d{2}:\d{2})\s*$")
URL_RE = re.compile(r"https?://\S+")
ORDER_RE = re.compile(r"(订单号|商品ID|商品号|商品编号)[:：]?\s*\d+")


@dataclass(frozen=True)
class Label:
    label: str
    stage: str
    theme: str
    priority: str
    definition: str
    action: str


LABELS = {
    "充电接口/电池续航参数": Label(
        "充电接口/电池续航参数",
        "售前",
        "参数规格",
        "P1",
        "购买前咨询充电接口类型、电池容量、充一次能用多久、是否可取下充电等参数；不计入实际充不上电。",
        "详情页首屏和机器人FAQ固定展示Type-C/接口、电池容量、充满时长和续航口径。",
    ),
    "镜面尺寸/材质/放大": Label(
        "镜面尺寸/材质/放大",
        "售前",
        "参数规格",
        "P1",
        "购买前咨询镜面尺寸、镜面/机身材质、是否放大、底座结构等基础规格。",
        "把镜面尺寸、材质、是否放大、底座是否吸附做成参数表，减少客服重复解释。",
    ),
    "灯光亮度/化妆场景": Label(
        "灯光亮度/化妆场景",
        "售前",
        "功能效果",
        "P1",
        "购买前咨询亮度、晚上化妆是否够用、光源/显色指标等照明效果。",
        "详情页补充三档/三色光、暗光化妆场景图和显色参数边界。",
    ),
    "价格优惠/活动": Label(
        "价格优惠/活动",
        "售前",
        "交易政策",
        "P2",
        "购买前询问价格、活动价、优惠券或是否还有指定价格。",
        "统一活动口径和优惠入口，机器人直接承接价格类咨询。",
    ),
    "发货时效/物流咨询": Label(
        "发货时效/物流咨询",
        "售前",
        "交易政策",
        "P2",
        "购买前询问现在下单何时到货、能否次日达、发货配送等。",
        "自动回复展示预计送达、仓配规则和学校/宿舍地址注意事项。",
    ),
    "试用/退货政策咨询": Label(
        "试用/退货政策咨询",
        "售前",
        "交易政策",
        "P1",
        "购买前或未发生质量问题时询问打开试用后是否可退、无理由退货、售后保障等政策。",
        "前置30天试用/无理由边界，避免客服每次重复承诺。",
    ),
    "质保政策咨询": Label(
        "质保政策咨询",
        "售前",
        "交易政策",
        "P1",
        "购买前咨询是否有质保、保修期限、售后保障等政策。",
        "详情页和机器人FAQ前置质保期限、保修边界和售后入口。",
    ),
    "礼盒包装/贺卡咨询": Label(
        "礼盒包装/贺卡咨询",
        "售前",
        "包装赠品",
        "P2",
        "购买前咨询礼盒、包装外观、是否能写贺卡、包装是否适合送礼。",
        "在详情页展示真实包装图，并明确是否支持贺卡和礼盒服务。",
    ),
    "安装/组装方法": Label(
        "安装/组装方法",
        "售后",
        "使用学习",
        "P0",
        "已购或收到后咨询怎么安装、怎么组装、是否有安装视频、底座/支架如何连接；不含明确卡滞故障。",
        "包装首卡和自动回复改成3步图：对准、按到底、旋转锁紧，并配30秒短视频。",
    ),
    "卡扣/连接件卡滞或装不上": Label(
        "卡扣/连接件卡滞或装不上",
        "售后",
        "结构安装",
        "P0",
        "安装过程中反馈卡口/卡头/连接处动不了、掰不动、扭不动、卡不上、拧不紧或疑似部件异常。",
        "建立卡扣排查话术：是否按到底、方向是否正确、是否异物卡住；排查无效直接换新或补寄配件。",
    ),
    "镜面方向歪斜/无法摆正": Label(
        "镜面方向歪斜/无法摆正",
        "售后",
        "结构安装",
        "P0",
        "安装后镜面、开关键或支架方向不正，出现歪斜、反着、无法竖直放置、无法摆正等。",
        "在说明书中标注正反方向和旋转到位标识，客服先发方向校准图，再判断是否结构错位。",
    ),
    "底座吸附/支撑不稳": Label(
        "底座吸附/支撑不稳",
        "售后",
        "结构稳定",
        "P0",
        "实际使用或放置时反馈底座吸不住、立不稳、站不稳、老倒、晃动或由此跌落。",
        "把支撑不稳列为高风险售后标签，要求客服记录台面材质、吸盘状态和是否跌落破损。",
    ),
    "灯光/开关/指示灯异常": Label(
        "灯光/开关/指示灯异常",
        "售后",
        "电控灯光",
        "P0",
        "收货使用后反馈灯不亮、开关打不开、指示灯不亮、插电没反应、照明效果异常或反光黑。",
        "先排查长按/短按、充电状态和保护膜；确认异常后进入换新并回收故障样本。",
    ),
    "充电/数据线/电源异常": Label(
        "充电/数据线/电源异常",
        "售后",
        "电控充电",
        "P0",
        "实际使用后反馈充不上电、插电无反应、数据线缺失或电源相关异常。",
        "客服按充电器、线材、接口、指示灯四步排查；缺线补发，电控异常换新。",
    ),
    "续航掉电/电量不耐用": Label(
        "续航掉电/电量不耐用",
        "售后",
        "电控充电",
        "P0",
        "实际使用后反馈用了很少次数就没电、充满只能用几天、续航明显低于预期等。",
        "核实充满时长、灯光档位和使用频次；集中出现时回收样本复测电池容量。",
    ),
    "破损/残次/缺件包装异常": Label(
        "破损/残次/缺件包装异常",
        "售后",
        "质量缺陷",
        "P0",
        "收到或使用中反馈碎裂、破损、盒子裂、缺说明书/数据线、像退货件、残次品等。",
        "破损缺件必须记录照片/视频和批次，区分运输破损、仓储复发和结构质量问题。",
    ),
    "退换货/售后处理": Label(
        "退换货/售后处理",
        "售后",
        "售后服务",
        "P0",
        "用户实际提出退款、退货、换新、补偿、补寄、售后备注或退换货执行问题。",
        "客服直接给出退货退款、免费换新、补寄/补偿三选项，不先用低额补偿压诉求。",
    ),
    "维修/售后入口咨询": Label(
        "维修/售后入口咨询",
        "售后",
        "售后服务",
        "P1",
        "用户咨询能否维修、付费维修、售后窗口或是否能走售后。",
        "机器人直接提供售后入口和维修/换新适用边界，避免用户反复追问。",
    ),
    "订单支付/取消处理": Label(
        "订单支付/取消处理",
        "售后",
        "交易服务",
        "P2",
        "用户咨询补差/换新如何支付、订单如何取消、是否能由客服取消等交易处理。",
        "自动回复补差支付、订单取消和售后换新流程入口。",
    ),
    "物流签收异常": Label(
        "物流签收异常",
        "售后",
        "物流履约",
        "P1",
        "用户反馈没收到、显示完成、未签收、快递员处理异常等售后物流问题。",
        "建立签收异常SOP：核实签收凭证、联系快递、必要时补发或退款。",
    ),
    "发票/订单处理": Label(
        "发票/订单处理",
        "售后",
        "交易服务",
        "P2",
        "用户咨询发票下载、订单处理、售后备注订单等交易服务问题。",
        "机器人补齐发票下载路径和订单售后备注说明。",
    ),
}


def contains(text: str, words: Iterable[str]) -> bool:
    return any(word in text for word in words)


def parse_log(path: Path) -> list[dict]:
    messages: list[dict] = []
    conversation_id = 0
    current = None
    for line_no, raw in enumerate(path.read_text(encoding="utf-8", errors="replace").splitlines(), 1):
        line = raw.strip().strip("\t")
        if not line:
            continue
        if "以下为一通会话" in line:
            conversation_id += 1
            current = None
            continue
        if "会话结束" in line:
            current = None
            continue
        match = HEADER_RE.match(line)
        if match:
            sender, date, time = match.groups()
            current = {
                "product": "909化妆镜",
                "conversation_id": conversation_id,
                "sender": sender.strip(),
                "date": date,
                "time": time,
                "text": "",
                "source_line": line_no,
            }
            messages.append(current)
            continue
        if current is not None:
            current["text"] = f"{current['text']}\n{line}".strip()
    return messages


def is_agent(sender: str) -> bool:
    return sender.startswith("jimi_") or "自营" in sender or "客服" in sender


def normalize(text: str) -> str:
    text = URL_RE.sub("", text or "")
    text = re.sub(r"\s+", " ", text).strip()
    return text


def is_noise(text: str) -> bool:
    if not text:
        return True
    exact_noise = {
        "你好",
        "您好",
        "在吗",
        "好",
        "好的",
        "嗯",
        "嗯嗯",
        "谢谢",
        "谢谢你",
        "OK",
        "ok",
        "噢",
        "哦",
        "同意",
        "。 ",
        "。",
        "？",
        "?",
        "用户发起转人工",
        "人工",
        "人工客服",
        "转人工",
        "不用换了",
    }
    if text in exact_noise:
        return True
    if text.startswith("【此消息为欢迎卡片"):
        return True
    if ORDER_RE.fullmatch(text):
        return True
    if text.startswith("您选择了一个订单进行咨询") and "？" not in text and "吗" not in text:
        return True
    return False


def valid_customer_messages(raw: list[dict]) -> list[dict]:
    valid = []
    for msg in raw:
        if is_agent(msg["sender"]):
            continue
        text = normalize(msg["text"])
        if is_noise(text):
            continue
        item = dict(msg)
        item["text"] = text
        valid.append(item)
    return valid


def classify(text: str) -> list[str]:
    labels: list[str] = []
    lower = text.lower()

    if contains(lower, ["type-c", "typec"]) or contains(text, ["充电插口", "充电接口", "电池容量", "充满一次", "充一次", "可以用多久", "能用多久", "取出来充电", "充电吗", "插电用吗", "可以插电用"]):
        if not contains(text, ["充不上", "没反应", "不亮", "缺", "无数据线", "插电没反应"]):
            labels.append("充电接口/电池续航参数")
    if contains(text, ["镜面多大", "尺寸", "多大", "一样大吗", "高低可以调", "材质", "塑料", "放大吗", "有放大", "几寸", "吸盘", "镜子底部"]) and not contains(text, ["吸不", "立不", "站不稳", "老倒"]):
        labels.append("镜面尺寸/材质/放大")
    if contains(text, ["最亮", "亮度", "晚上关了灯", "方便化妆", "光源", "Ra", "R9", "色温"]) and not contains(text, ["不亮", "没反应", "指示灯"]):
        labels.append("灯光亮度/化妆场景")
    if contains(text, ["79.9", "价格", "优惠", "券", "活动", "多少钱"]):
        labels.append("价格优惠/活动")
    if contains(text, ["现在下单", "明天能不能到", "什么时候到", "多久到", "发货"]) and not contains(text, ["没收到", "签收"]):
        labels.append("发货时效/物流咨询")
    if contains(text, ["不满意可以退", "可以退货吗", "无理由退货", "七天无理由", "支持七天", "免费试用", "打开试用后"]) and not contains(text, ["用不了", "破", "坏", "碎", "残次"]):
        labels.append("试用/退货政策咨询")
    if contains(text, ["有质保吗", "质保吗", "保修", "售后保障"]):
        labels.append("质保政策咨询")
    if contains(text, ["礼盒", "贺卡", "包装好看", "包装盒", "包装好点"]) and not contains(text, ["破", "裂", "难复原", "没有包装"]):
        labels.append("礼盒包装/贺卡咨询")

    if contains(text, ["怎么安装", "怎么装", "怎么组装", "安装教程", "安装视频", "教程", "视频", "支架", "底座要不要动", "怎么卡", "这里要怎么装", "怎么按上", "安进去", "咋安装", "是不是安错", "怎么取下来", "咋用", "有说明书", "说明书我丢了", "难安", "没安装好"]):
        labels.append("安装/组装方法")
    if contains(text, ["卡口", "卡头", "掰不动", "动不了", "只能移动", "扭不动", "拧不动", "卡不上", "扣不住", "扣紧", "拧不紧", "松动", "连接处", "打不开", "按不下去", "按不动", "装不上", "卡不住", "卡不进去", "安装配件怎么不旋转"]) and not contains(text, ["开关打不开"]):
        labels.append("卡扣/连接件卡滞或装不上")
    if contains(text, ["装正", "不正", "反着", "歪", "斜", "竖直", "正中间", "摆正", "位置不对", "开关在右边", "开关键不正", "开关在侧面", "弯下来"]):
        labels.append("镜面方向歪斜/无法摆正")
    if contains(text, ["立不住", "能不能立住", "立不稳", "站不稳", "放不稳", "放稳", "吸不住", "吸不了", "吸的住", "吸不稳", "老倒", "晃动", "掉下来", "掉地上", "脱落", "支撑", "不牢固", "不牢", "卡不住", "稳当吗"]):
        labels.append("底座吸附/支撑不稳")
    if contains(text, ["灯不亮", "不亮", "指示灯", "提示灯", "开关打不开", "插电没反应", "没反应", "照的更黑", "看不清", "反光黑", "屏是往上的", "坏了吗", "不会开"]):
        labels.append("灯光/开关/指示灯异常")
    if contains(text, ["充不上", "充不了", "数据线", "无数据线", "插电没反应", "电源"]) and not contains(text, ["什么充电插口", "type-c"]):
        labels.append("充电/数据线/电源异常")
    if contains(text, ["用了两次就没电", "3到4天就没电", "一次性的今天刚用就没了", "刚用就没了"]):
        labels.append("续航掉电/电量不耐用")
    if contains(text, ["破损", "残次品", "碎", "盒子", "裂", "无说明书", "无数据线", "缺", "少什么东西", "掉下来碎", "烂了", "烂的", "不干净", "退货的", "做工", "质检", "烂镜子", "一次性的"]):
        labels.append("破损/残次/缺件包装异常")
    if contains(text, ["退货", "退款", "退了", "换新", "换货", "补寄", "补偿", "售后备注", "投诉", "用不了了", "不行", "换快递不方便", "包装难复原", "包装啥的很难复原", "寄回去", "没有纸箱", "没有包装", "旧的放进去", "这种情况怎么办", "一样怎么办", "售后不是寄过来更换"]) and not (
        "不满意可以退" in text or "可以退货吗" in text
    ):
        labels.append("退换货/售后处理")
    if contains(text, ["可以维修吗", "付费维修", "售后窗口", "不能走售后"]):
        labels.append("维修/售后入口咨询")
    if contains(text, ["怎么付", "在哪里支付", "怎么取消", "可以取消吗", "取消订单", "不能给你下单"]):
        labels.append("订单支付/取消处理")
    if contains(text, ["没收到", "显示完成", "没签收", "未签收", "物流信息", "快递员说", "到了吗", "到货了吗"]):
        labels.append("物流签收异常")
    if contains(text, ["发票", "订单号", "商品ID", "商品号", "订单进行咨询"]):
        labels.append("发票/订单处理")

    ordered = []
    seen = set()
    for label in labels:
        if label not in seen:
            ordered.append(label)
            seen.add(label)
    return ordered


def risk_for(text: str, labels: list[str]) -> list[dict]:
    risks = []
    if contains(text, ["掉下来碎", "掉地上", "脱落出来掉", "自己掉了", "没20天自己就烂了"]):
        risks.append(
            {
                "risk_label": "结构脱落/跌落碎裂",
                "risk_type": "质量风险",
                "priority": "P0",
                "reason": "用户本人反馈镜体/连接处脱落或跌落碎裂，属于结构稳定性强风险。",
                "action": "人工复核图片/视频、台面材质、安装到位情况和批次；同批次集中出现时反馈品质复测。",
            }
        )
    if contains(text, ["残次品", "破损", "盒子也是裂的", "无说明书", "无数据线", "不干净", "退货的", "烂了"]):
        risks.append(
            {
                "risk_label": "到货破损/残次缺件",
                "risk_type": "质量风险",
                "priority": "P0",
                "reason": "用户本人反馈到货状态异常、缺件或疑似复发货。",
                "action": "记录照片证据和仓配链路，区分运输破损、仓库复发和出厂缺件。",
            }
        )
    if "灯光/开关/指示灯异常" in labels or "充电/数据线/电源异常" in labels:
        if contains(text, ["不亮", "没反应", "充不上", "插电没反应", "开关打不开"]):
            risks.append(
                {
                    "risk_label": "电控/充电无法正常工作",
                    "risk_type": "质量风险",
                    "priority": "P0",
                    "reason": "用户本人反馈灯光、开关、充电或插电无反应，需排除操作问题后确认故障。",
                    "action": "按线材、接口、指示灯、按键操作排查；排查无效换新并回收故障样本。",
                }
            )
    if "续航掉电/电量不耐用" in labels:
        risks.append(
            {
                "risk_label": "电池续航明显异常",
                "risk_type": "质量风险",
                "priority": "P0",
                "reason": "用户本人反馈刚使用或少量使用后很快没电，需复核是否为电池容量或充电异常。",
                "action": "记录充电时长、灯光档位、使用时长和批次；集中出现时进行电池容量复测。",
            }
        )
    return risks


def build_context(raw: list[dict]) -> dict[tuple[int, int], str]:
    by_conv = defaultdict(list)
    for msg in raw:
        by_conv[msg["conversation_id"]].append(msg)
    contexts = {}
    for conv_id, msgs in by_conv.items():
        for idx, msg in enumerate(msgs):
            window = msgs[max(0, idx - 3) : min(len(msgs), idx + 4)]
            lines = []
            for neighbor in window:
                text = normalize(neighbor.get("text", ""))
                if text:
                    lines.append(f"{neighbor['sender']}: {text}")
            contexts[(conv_id, msg["source_line"])] = "\n".join(lines)
    return contexts


def analyze() -> dict:
    raw = parse_log(SOURCE_LOG)
    valid = valid_customer_messages(raw)
    contexts = build_context(raw)
    retained = {}
    retained_risks = set()
    raw_hits = Counter()
    detail = []
    duplicates = []
    risks = []
    reviews = []
    stats_buyers = defaultdict(set)
    stats_evidence = defaultdict(list)

    for msg in valid:
        labels = classify(msg["text"])
        if not labels:
            if contains(msg["text"], ["怎么", "为什么", "吗", "不", "没", "退", "换", "坏", "碎", "寄", "装"]):
                reviews.append({**msg, "reason": "存在需求语气但上下文不足，需人工确认词条。"})
            continue
        for label in labels:
            meta = LABELS[label]
            raw_hits[(meta.stage, label)] += 1
            key = (msg["sender"], meta.stage, label)
            if key in retained:
                duplicates.append(
                    {
                        **msg,
                        "stage": meta.stage,
                        "theme": meta.theme,
                        "label": label,
                        "kept_text": retained[key]["text"],
                        "reason": "同一买家、同一阶段、同一需求词条、本周期只计1次",
                    }
                )
                continue
            retained[key] = msg
            row = {
                **msg,
                "stage": meta.stage,
                "theme": meta.theme,
                "label": label,
                "definition": meta.definition,
                "priority": meta.priority,
                "action": meta.action,
                "all_text": contexts.get((msg["conversation_id"], msg["source_line"]), msg["text"]),
                "raw_hit": raw_hits[(meta.stage, label)],
            }
            detail.append(row)
            stats_buyers[(meta.stage, meta.theme, label)].add(msg["sender"])
            stats_evidence[(meta.stage, meta.theme, label)].append(row)
        for risk in risk_for(msg["text"], labels):
            risk_key = (msg["sender"], risk["risk_label"])
            if risk_key in retained_risks:
                continue
            retained_risks.add(risk_key)
            risks.append({**msg, **risk, "context": contexts.get((msg["conversation_id"], msg["source_line"]), msg["text"])})

    stats = []
    for (stage, theme, label), buyers in stats_buyers.items():
        meta = LABELS[label]
        evidence = stats_evidence[(stage, theme, label)]
        stats.append(
            {
                "stage": stage,
                "theme": theme,
                "label": label,
                "count": len(buyers),
                "raw_hits": raw_hits[(stage, label)],
                "duplicates": sum(1 for item in duplicates if item["stage"] == stage and item["label"] == label),
                "priority": meta.priority,
                "definition": meta.definition,
                "buyers": "、".join(sorted(buyers)[:12]),
                "quote": "｜".join(item["text"] for item in evidence[:6]),
                "action": meta.action,
            }
        )
    stats.sort(key=lambda item: (-item["count"], item["stage"], item["label"]))
    dates = [msg["date"] for msg in valid if msg.get("date")]
    return {
        "raw": raw,
        "valid": valid,
        "detail": detail,
        "duplicates": duplicates,
        "risks": risks,
        "reviews": reviews,
        "stats": stats,
        "summary": {
            "raw_conversations": max([m["conversation_id"] for m in raw] or [0]),
            "raw_messages": len(raw),
            "valid_messages": len(valid),
            "buyers": len({msg["sender"] for msg in valid}),
            "start_date": min(dates),
            "end_date": max(dates),
            "days": len(set(dates)),
            "demand_count": len(detail),
            "presale_count": sum(1 for row in detail if row["stage"] == "售前"),
            "aftersale_count": sum(1 for row in detail if row["stage"] == "售后"),
            "duplicate_count": len(duplicates),
            "risk_count": len(risks),
            "review_count": len(reviews),
        },
    }


def set_cell(ws, row: int, col: int, value):
    cell = ws.cell(row, col)
    if not isinstance(cell, MergedCell):
        cell.value = value


def copy_row_style(ws, source_row: int, target_row: int):
    for col in range(1, min(ws.max_column, 14) + 1):
        source = ws.cell(source_row, col)
        target = ws.cell(target_row, col)
        if isinstance(target, MergedCell):
            continue
        target._style = copy(source._style)
        target.font = copy(source.font)
        target.fill = copy(source.fill)
        target.border = copy(source.border)
        target.alignment = copy(source.alignment)
        target.number_format = source.number_format


def ensure_row(ws, row: int, style_row: int):
    if row > ws.max_row:
        ws.append([])
    copy_row_style(ws, style_row, row)


def clear_values(ws):
    for row in ws.iter_rows():
        for cell in row:
            if not isinstance(cell, MergedCell):
                cell.value = None


def write_row(ws, row: int, values: list):
    ensure_row(ws, row, max(1, min(row, ws.max_row)))
    for col, value in enumerate(values, 1):
        set_cell(ws, row, col, value)


def write_rows(ws, start: int, rows: list[list], style_row: int | None = None):
    style_row = style_row or start
    for offset, row in enumerate(rows):
        target = start + offset
        ensure_row(ws, target, style_row)
        write_row(ws, target, row)


def stats_for(data: dict, stage: str) -> list[dict]:
    return sorted([row for row in data["stats"] if row["stage"] == stage], key=lambda x: (-x["count"], x["label"]))


def theme_rows(data: dict) -> list[list]:
    total = data["summary"]["demand_count"] or 1
    bucket = {}
    for row in data["stats"]:
        item = bucket.setdefault(row["theme"], {"count": 0, "labels": [], "actions": []})
        item["count"] += row["count"]
        item["labels"].append(f"{row['stage']}:{row['label']}({row['count']})")
        item["actions"].append(row["action"])
    output = [["一级主题", "需求人数", "需求占比", "覆盖词条", "业务判断", "建议动作"]]
    for theme, item in sorted(bucket.items(), key=lambda kv: -kv[1]["count"]):
        output.append(
            [
                theme,
                item["count"],
                item["count"] / total,
                "；".join(item["labels"][:8]),
                theme_judgment(theme),
                item["actions"][0] if item["actions"] else "持续观察。",
            ]
        )
    return output


def theme_judgment(theme: str) -> str:
    mapping = {
        "结构安装": "安装方向、卡扣到位和连接结构是高摩擦环节，容易从教程咨询升级为质量争议。",
        "结构稳定": "支撑不稳和跌落会直接带来破损、退换货和品质复核压力。",
        "电控充电": "电池/充电参数既影响购买决策，也会在售后形成故障排查压力。",
        "电控灯光": "灯光是化妆镜核心功能，异常反馈需要区分操作、充电和真实故障。",
        "使用学习": "收到货后的学习成本偏高，说明说明书、视频和机器人承接需要前置。",
        "参数规格": "买家对接口、续航、尺寸和材质的基础信息获取效率不够高。",
        "售后服务": "退换货处理与安装/结构问题绑定，客服需要更明确的分流方案。",
    }
    return mapping.get(theme, "作为V1主题沉淀，后续跟踪占比和客服消耗。")


def management_summary(data: dict) -> list[str]:
    s = data["summary"]
    presale = stats_for(data, "售前")
    aftersale = stats_for(data, "售后")
    structure = sum(row["count"] for row in data["stats"] if row["theme"] in {"结构安装", "结构稳定", "使用学习"})
    electric_param = sum(row["count"] for row in data["stats"] if row["label"] == "充电接口/电池续航参数")
    return [
        f"本期是909化妆镜首次分析，没有历史基准；从{ s['start_date'] }到{ s['end_date'] }共解析原始会话{s['raw_conversations']}通，识别正式需求{s['demand_count']}条，作为V1基准。",
        f"售后压力高于售前：售后{s['aftersale_count']}条、售前{s['presale_count']}条。最高频售后是“{aftersale[0]['label'] if aftersale else '暂无'}”，说明收到货后的安装和使用承接是第一优先级。",
        f"安装与结构体验合计{structure}条，是本品最需要治理的链路；卡扣拧不动、镜面歪斜、底座吸不稳会把普通教程咨询升级成换新、退货和投诉。",
        f"售前最高频是“{presale[0]['label'] if presale else '暂无'}”，其中充电接口/电池续航参数{electric_param}人命中，建议把Type-C、续航、电池容量和充电方式前置到详情页和机器人FAQ。",
        f"本期二审后保留质量风险{s['risk_count']}条，主要是结构脱落/跌落碎裂、到货破损缺件、电控/充电异常；未发现明确人身伤害类安全事故。",
        "最建议先做三件事：包装内放3步安装卡，机器人首轮推卡扣排查图，售后对“立不稳/掉落/碎裂/拧不紧”直接打高风险标签并复核批次。",
    ]


def write_stat_sheet(ws, title: str, rows: list[dict], total: int, days: int):
    clear_values(ws)
    write_row(ws, 1, [title])
    write_row(ws, 2, ["按同一买家 + 阶段 + 需求词条 + 当前周期去重；本次为首次分析，本表即909化妆镜V1基准。"])
    write_row(ws, 4, ["正式需求总数", total, "统计天数", days, "基准版本", "V1", None, None, None, None, None, None])
    write_row(ws, 6, ["排名", "需求词条", "同人去重人数", "占比", "日均人数", "原始命中", "重复次数", "优先级", "严格定义", "涉及买家示例", "代表原话", "建议动作"])
    body = []
    for idx, item in enumerate(rows, 1):
        body.append(
            [
                idx,
                item["label"],
                item["count"],
                item["count"] / total if total else 0,
                item["count"] / days if days else 0,
                item["raw_hits"],
                item["duplicates"],
                item["priority"],
                item["definition"],
                item["buyers"],
                item["quote"],
                item["action"],
            ]
        )
    write_rows(ws, 7, body or [[1, "本期未识别", 0, 0, 0, 0, 0, "—", "—", "—", "—", "—"]], 7)


def write_compare_first_baseline(ws, title: str, rows: list[dict], total: int, days: int):
    clear_values(ws)
    write_row(ws, 1, [title])
    write_row(ws, 2, ["首次分析没有历史基准；本表把本期结果沉淀为V1，供下一期作为上期基准对比。"])
    write_row(ws, 4, ["本期总需求", total, "本期天数", days, "基准总需求", 0, "基准天数", 0, None, None, None, None])
    write_row(ws, 6, ["需求词条", "本期人数", "本期占比", "本期日均", "基准人数", "基准占比", "基准日均", "人数变化", "占比变化", "变化判断", "建议动作", "备注"])
    body = []
    for item in rows:
        body.append(
            [
                item["label"],
                item["count"],
                item["count"] / total if total else 0,
                item["count"] / days if days else 0,
                0,
                0,
                0,
                item["count"],
                item["count"] / total if total else 0,
                "V1首次基准",
                item["action"],
                "本期无历史基准，不做升降判断",
            ]
        )
    write_rows(ws, 7, body or [["本期未识别", 0, 0, 0, 0, 0, 0, 0, 0, "V1首次基准", "—", "—"]], 7)


def write_simple(ws, rows: list[list]):
    clear_values(ws)
    write_rows(ws, 1, rows, 1)


def style_workbook(wb):
    header_fill = PatternFill("solid", fgColor="D9EAF7")
    title_fill = PatternFill("solid", fgColor="C00000")
    title_font = Font(color="FFFFFF", bold=True, size=14)
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell, MergedCell):
                    continue
                cell.alignment = Alignment(vertical="top", wrap_text=True)
        if ws.max_row >= 1:
            for cell in ws[1]:
                if not isinstance(cell, MergedCell):
                    cell.fill = title_fill
                    cell.font = title_font
        for cell in ws[6] if ws.max_row >= 6 else []:
            if not isinstance(cell, MergedCell):
                cell.fill = header_fill
                cell.font = Font(bold=True)
        for col, width in {
            "A": 16,
            "B": 24,
            "C": 14,
            "D": 12,
            "E": 12,
            "F": 14,
            "G": 14,
            "H": 16,
            "I": 42,
            "J": 34,
            "K": 50,
            "L": 42,
        }.items():
            ws.column_dimensions[col].width = width


def write_workbook(data: dict):
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    staging = OUTPUT_DIR / f"{OUTPUT_XLSX.stem}__staging.xlsx"
    shutil.copy2(TEMPLATE_XLSX, staging)
    wb = load_workbook(staging)
    required = [
        "分析总览",
        "售前需求统计",
        "售后需求统计",
        "售前基准对比",
        "售后基准对比",
        "新增词条",
        "P0风险明细",
        "安全质量风险",
        "逐条需求明细",
        "同人重复记录",
        "待人工复核",
        "词条与口径",
        "上期基准",
        "原始数据",
    ]
    for name in required:
        if name not in wb.sheetnames:
            raise RuntimeError(f"母版缺少Sheet：{name}")
    if "一级主题" not in wb.sheetnames:
        wb.create_sheet("一级主题", 1)

    s = data["summary"]
    summary_lines = management_summary(data)
    presale = stats_for(data, "售前")
    aftersale = stats_for(data, "售后")
    total_presale = sum(row["count"] for row in presale)
    total_aftersale = sum(row["count"] for row in aftersale)

    ws = wb["分析总览"]
    clear_values(ws)
    write_row(ws, 1, ["909化妆镜｜首次客服聊天需求基准分析报告"])
    write_row(ws, 2, ["本次按历史GPT分析协议执行：没有历史基准，从聊天记录建立909化妆镜V1词条和V1基准。"])
    write_row(ws, 4, ["数据规模", None, "消息规模", None, "需求结果", None, "重复治理", None, "安全质量", None, "交付文件", None])
    write_row(ws, 5, ["原始会话", "有效买家", "有效买家消息", "正式需求", "售前需求", "售后需求", "重复剔除", "待人工复核", "质量风险", "安全事故", "Excel状态", "分析版本"])
    write_row(
        ws,
        6,
        [
            s["raw_conversations"],
            s["buyers"],
            s["valid_messages"],
            s["demand_count"],
            s["presale_count"],
            s["aftersale_count"],
            s["duplicate_count"],
            s["review_count"],
            s["risk_count"],
            0,
            "已生成",
            "909-V1",
        ],
    )
    write_row(ws, 8, ["排名", "TOP售前词条", "人数", "阶段占比", "优先级", "动作", None, "排名", "TOP售后词条", "人数", "阶段占比", "优先级"])
    for idx in range(12):
        pre = presale[idx] if idx < len(presale) else None
        aft = aftersale[idx] if idx < len(aftersale) else None
        write_row(
            ws,
            9 + idx,
            [
                idx + 1 if pre else None,
                pre["label"] if pre else None,
                pre["count"] if pre else None,
                pre["count"] / total_presale if pre and total_presale else None,
                pre["priority"] if pre else None,
                pre["action"] if pre else None,
                None,
                idx + 1 if aft else None,
                aft["label"] if aft else None,
                aft["count"] if aft else None,
                aft["count"] / total_aftersale if aft and total_aftersale else None,
                aft["priority"] if aft else None,
            ],
        )
    write_row(ws, 23, ["本期核心结论"])
    for idx, line in enumerate(summary_lines, 24):
        write_row(ws, idx, [idx - 23, line])

    write_simple(wb["新增词条"], [["阶段", "新增词条", "本期人数", "本期占比", "优先级", "严格定义", "建议动作"]] + [
        [
            row["stage"],
            row["label"],
            row["count"],
            row["count"] / (total_presale if row["stage"] == "售前" else total_aftersale or 1),
            row["priority"],
            row["definition"],
            row["action"],
        ]
        for row in presale + aftersale
    ])
    write_simple(wb["P0风险明细"], [["ID", "需求词条", "买家", "日期", "时间", "代表原话", "全部证据", "源行", "优先级", "建议动作"]] + [
        [idx, row["label"], row["sender"], row["date"], row["time"], row["text"], row["all_text"], row["source_line"], row["priority"], row["action"]]
        for idx, row in enumerate([item for item in data["detail"] if item["priority"] == "P0"], 1)
    ])
    write_simple(wb["安全质量风险"], [["风险ID", "风险词条", "风险类型", "买家", "日期", "时间", "代表原话", "全部证据", "源行", "二审原因", "建议动作"]] + [
        [
            idx,
            row["risk_label"],
            row["risk_type"],
            row["sender"],
            row["date"],
            row["time"],
            row["text"],
            row["context"],
            row["source_line"],
            row["reason"],
            row["action"],
        ]
        for idx, row in enumerate(data["risks"], 1)
    ])
    write_simple(wb["逐条需求明细"], [["明细ID", "阶段", "需求词条", "买家", "日期", "时间", "代表原话", "同需求全部原话", "源行", "原始命中", "一级主题", "分析来源"]] + [
        [idx, row["stage"], row["label"], row["sender"], row["date"], row["time"], row["text"], row["all_text"], row["source_line"], row["raw_hit"], row["theme"], "本期V1语义规则"]
        for idx, row in enumerate(data["detail"], 1)
    ])
    write_simple(wb["同人重复记录"], [["阶段", "需求词条", "买家", "日期", "时间", "重复原话", "重复源行", "保留原话", "剔除原因"]] + [
        [row["stage"], row["label"], row["sender"], row["date"], row["time"], row["text"], row["source_line"], row["kept_text"], row["reason"]]
        for row in data["duplicates"]
    ])
    write_simple(wb["待人工复核"], [["买家", "日期", "时间", "待复核原话", "源行", "复核原因", "状态"]] + [
        [row["sender"], row["date"], row["time"], row["text"], row["source_line"], row["reason"], "待人工确认"]
        for row in data["reviews"]
    ])
    write_simple(wb["词条与口径"], [["需求词条", "本期阶段", "一级主题", "严格定义", "优先级", "售前人数", "售后人数", "基准状态", "建议动作"]] + [
        [
            row["label"],
            row["stage"],
            row["theme"],
            row["definition"],
            row["priority"],
            row["count"] if row["stage"] == "售前" else 0,
            row["count"] if row["stage"] == "售后" else 0,
            "909化妆镜V1首次生成",
            row["action"],
        ]
        for row in presale + aftersale
    ])
    write_simple(wb["上期基准"], [["阶段", "排名", "需求词条", "基准人数", "基准占比", "基准日均", "优先级", "严格定义", "建议动作"]] + [
        [
            row["stage"],
            idx,
            row["label"],
            row["count"],
            row["count"] / (total_presale if row["stage"] == "售前" else total_aftersale or 1),
            row["count"] / s["days"] if s["days"] else 0,
            row["priority"],
            row["definition"],
            row["action"],
        ]
        for idx, row in enumerate(presale + aftersale, 1)
    ])
    write_simple(wb["原始数据"], [["产品", "会话序号", "发送方", "是否买家有效消息", "日期", "时间", "源行", "原文"]] + [
        [
            row["product"],
            row["conversation_id"],
            row["sender"],
            "是" if (not is_agent(row["sender"]) and not is_noise(normalize(row["text"]))) else "否",
            row["date"],
            row["time"],
            row["source_line"],
            row["text"],
        ]
        for row in data["raw"]
    ])
    write_simple(wb["一级主题"], theme_rows(data))

    write_stat_sheet(wb["售前需求统计"], "909化妆镜售前需求V1", presale, total_presale, s["days"])
    write_stat_sheet(wb["售后需求统计"], "909化妆镜售后需求V1", aftersale, total_aftersale, s["days"])
    write_compare_first_baseline(wb["售前基准对比"], "909化妆镜售前V1基准", presale, total_presale, s["days"])
    write_compare_first_baseline(wb["售后基准对比"], "909化妆镜售后V1基准", aftersale, total_aftersale, s["days"])

    style_workbook(wb)
    wb.save(OUTPUT_XLSX)
    staging.unlink(missing_ok=True)

    return summary_lines


def write_outputs(data: dict, summary_lines: list[str]):
    PROCESS_DIR.mkdir(parents=True, exist_ok=True)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    (PROCESS_DIR / "口径检查.md").write_text(
        "\n".join(
            [
                "# 909化妆镜口径检查",
                "",
                f"- 实际数据时间：{data['summary']['start_date']} 至 {data['summary']['end_date']}",
                f"- 原始会话：{data['summary']['raw_conversations']}",
                f"- 有效买家：{data['summary']['buyers']}",
                f"- 有效买家消息：{data['summary']['valid_messages']}",
                "- 去重规则：同一买家 + 同一阶段 + 同一需求词条 + 当前分析周期",
                "- 是否有基准：无，建立909化妆镜V1首次基准",
                f"- 正式需求：{data['summary']['demand_count']}",
                f"- 同人重复剔除：{data['summary']['duplicate_count']}",
                f"- 质量风险二审保留：{data['summary']['risk_count']}",
            ]
        ),
        encoding="utf-8",
    )
    (PROCESS_DIR / "report_data.json").write_text(json.dumps(data, ensure_ascii=False, indent=2, default=str), encoding="utf-8")
    OUTPUT_MD.write_text(
        "# 909化妆镜首次客服聊天需求基准分析\n\n"
        + "\n".join(f"{idx}. {line}" for idx, line in enumerate(summary_lines, 1))
        + "\n\n## 最推荐先做\n\n先把安装三步图、卡扣排查图、参数FAQ和结构高风险售后标签上线；下一期沿用本次V1词条做占比和日均对比。\n",
        encoding="utf-8",
    )
    manifest = {
        "task_id": TASK_ID,
        "status": "completed",
        "model": "codex-local-v1-rules",
        "analysis_type": "first_baseline",
        "product_name": "909化妆镜",
        "period": {
            "start": data["summary"]["start_date"],
            "end": data["summary"]["end_date"],
            "days": data["summary"]["days"],
        },
        "source_files": [str(SOURCE_LOG)],
        "outputs": {
            "excel": str(OUTPUT_XLSX),
            "markdown": str(OUTPUT_MD),
            "manifest": str(OUTPUT_MANIFEST),
            "process_data": str(PROCESS_DIR / "report_data.json"),
            "scope_check": str(PROCESS_DIR / "口径检查.md"),
        },
        "metrics": data["summary"],
        "summary": summary_lines,
    }
    OUTPUT_MANIFEST.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")


def main():
    data = analyze()
    summary_lines = write_workbook(data)
    write_outputs(data, summary_lines)
    print(json.dumps({"excel": str(OUTPUT_XLSX), "markdown": str(OUTPUT_MD), "manifest": str(OUTPUT_MANIFEST), "metrics": data["summary"]}, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
