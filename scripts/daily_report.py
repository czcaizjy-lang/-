#!/usr/bin/env python3
"""
每日业绩日报生成 + 飞书推送
读取 dashboard_data.json，计算关键指标，推送到飞书群机器人

用法:
  FEISHU_WEBHOOK="https://open.feishu.cn/..." python3 daily_report.py

定时 (crontab):
  0 9 * * * FEISHU_WEBHOOK="https://..." python3 /path/to/daily_report.py
"""

import json
import os
import sys
import locale
from datetime import datetime, timedelta
from collections import defaultdict

import requests

# ---- 配置 ----
DATA_PATH = os.path.join(
    os.path.dirname(os.path.abspath(__file__)), "..", "data", "dashboard_data.json"
)
FEISHU_WEBHOOK = os.environ.get("FEISHU_WEBHOOK", "")
TODAY = datetime.now()
YESTERDAY = TODAY - timedelta(days=1)
DAY_BEFORE = TODAY - timedelta(days=2)
DAYS_IN_MONTH = 30  # 6月
WEEKDAY_CN = ["周一", "周二", "周三", "周四", "周五", "周六", "周日"]


def fmt_wan(val: float) -> str:
    """格式化为万元，保留2位小数"""
    return f"{val:,.2f}万"


def fmt_pct(val: float) -> str:
    """格式化为百分比"""
    return f"{val * 100:.2f}%"


def fmt_change(new: float, old: float) -> str:
    """格式环比变化"""
    if old == 0:
        return "新增" if new > 0 else "—"
    change = (new - old) / old
    arrow = "↑" if change > 0 else "↓" if change < 0 else "→"
    return f"{arrow}{abs(change) * 100:.1f}%"


def load_data():
    with open(DATA_PATH, "r", encoding="utf-8") as f:
        return json.load(f)


def get_daily_value(trend: dict, key: str, date_str: str) -> float:
    """从 trend 数组中取某一天的值"""
    dates = trend["dates"]
    if date_str not in dates:
        return 0.0
    idx = dates.index(date_str)
    arr = trend[key]
    if idx < len(arr):
        return float(arr[idx])
    return 0.0


def build_report(data: dict) -> list:
    """构建飞书富文本消息的 content 数组"""
    summary = data["summary"]
    trend = data["trend"]
    anchors = data["anchors"]

    yesterday_str = YESTERDAY.strftime("%m/%d")  # 如 06/17
    day_before_str = DAY_BEFORE.strftime("%m/%d")  # 如 06/16
    today_str = TODAY.strftime("%m/%d")
    month_str = TODAY.strftime("%m月")

    # ---- 基础数据 ----
    # 月度累计 (从 summary，单位：元 → 万)
    total_gmv = summary["直播GMV"] / 10000
    total_refund = summary["直播退款GMV"] / 10000
    total_settled = summary["直播结算GMV"] / 10000
    settle_rate = summary["目前结算率"]
    target_income = summary["预估6月收入"] / 10000
    achieve_rate = summary["当前收入达成率"]
    profit = summary["预估利润"] / 10000
    ad_cost = summary["投放消耗金额"] / 10000
    commission = summary["佣金支出"] / 10000

    # 昨日数据 (从 trend，单位已是万)
    yesterday_gmv = get_daily_value(trend, "daily_total_gmv", yesterday_str)
    yesterday_paid = get_daily_value(trend, "daily_total_paid", yesterday_str)
    yesterday_refund = get_daily_value(trend, "daily_total_refund", yesterday_str)

    # 前日数据
    before_gmv = get_daily_value(trend, "daily_total_gmv", day_before_str)
    before_paid = get_daily_value(trend, "daily_total_paid", day_before_str)
    before_refund = get_daily_value(trend, "daily_total_refund", day_before_str)

    # ---- 时间进度对标 ----
    days_passed = TODAY.day  # 当月已过天数
    time_progress = days_passed / DAYS_IN_MONTH
    expected_gmv = target_income * time_progress
    gmv_gap = total_gmv - expected_gmv

    # ---- 机构昨日排名 ----
    daily_by_agency = trend.get("daily_by_agency", {})
    agency_yesterday = []
    for ag, vals in daily_by_agency.items():
        dates = trend["dates"]
        if yesterday_str in dates:
            idx = dates.index(yesterday_str)
            if idx < len(vals):
                agency_yesterday.append((ag, float(vals[idx])))
    agency_yesterday.sort(key=lambda x: x[1], reverse=True)
    top5_agencies = agency_yesterday[:5]

    # ---- 达人掉量预警 ----
    anchor_daily_paid = trend.get("anchor_daily_paid", {})
    anchor_declines = []
    id_to_info = {str(a["主播抖音号"]): a for a in anchors}

    for douyin_id, daily in anchor_daily_paid.items():
        y_val = float(daily.get(yesterday_str, 0))
        b_val = float(daily.get(day_before_str, 0))
        if b_val <= 0:  # 前日没量，不算"掉量"
            continue
        decline = b_val - y_val  # 正值=下跌
        decline_pct = decline / b_val
        if decline > 0.1:  # 掉量 > 0.1万（1000元）才算
            info = id_to_info.get(douyin_id, {})
            anchor_declines.append(
                {
                    "name": info.get("主播昵称", douyin_id),
                    "agency": info.get("机构", "未知"),
                    "yesterday": y_val,
                    "before": b_val,
                    "decline": decline,
                    "decline_pct": decline_pct,
                    "douyin_id": douyin_id,
                    "gmv": float(info.get("直播GMV", 0)) / 10000,
                    "broadcast_days": info.get("开播天数", 0),
                    "settle_rate": float(info.get("目前结算率", 0)),
                    "ad_cost": float(info.get("投放消耗金额", 0)) / 10000,
                }
            )

    anchor_declines.sort(key=lambda x: x["decline"], reverse=True)
    top5_declines = anchor_declines[:5]

    # ---- 组装消息 ----
    content = []

    # 标题行
    content.append(
        [{"tag": "text", "text": f"📊 蕉下{month_str}业绩日报 | {today_str}\n\n"}]
    )

    # === 一、月度总览 ===
    content.append([{"tag": "text", "text": "━ 月度总览 ━\n"}])
    content.append(
        [
            {
                "tag": "text",
                "text": (
                    f"💰 累计GMV：{fmt_wan(total_gmv)}　"
                    f"达成率：{fmt_pct(achieve_rate)}\n"
                    f"💸 累计退款：{fmt_wan(total_refund)}　"
                    f"结算率：{fmt_pct(settle_rate)}\n"
                    f"📊 投放消耗：{fmt_wan(ad_cost)}　"
                    f"佣金支出：{fmt_wan(commission)}\n"
                    f"🏦 预估利润：{fmt_wan(profit)}　"
                    f"月度目标：{fmt_wan(target_income)}\n\n"
                ),
            }
        ]
    )

    # === 二、昨日战报 ===
    content.append([{"tag": "text", "text": "━ 昨日战报 ━\n"}])
    content.append(
        [
            {
                "tag": "text",
                "text": (
                    f"📅 {yesterday_str}（{WEEKDAY_CN[YESTERDAY.weekday()]}）\n"
                    f"GMV：{fmt_wan(yesterday_gmv)}　"
                    f"环比{fmt_change(yesterday_gmv, before_gmv)}\n"
                    f"支付：{fmt_wan(yesterday_paid)}　"
                    f"环比{fmt_change(yesterday_paid, before_paid)}\n"
                    f"退款：{fmt_wan(yesterday_refund)}　"
                    f"环比{fmt_change(yesterday_refund, before_refund)}\n\n"
                ),
            }
        ]
    )

    # === 三、时间进度对标 ===
    content.append([{"tag": "text", "text": "━ 进度对标 ━\n"}])
    gap_text = "领先 ✅" if gmv_gap >= 0 else "落后 ⚠️"
    content.append(
        [
            {
                "tag": "text",
                "text": (
                    f"⏳ 时间进度：{fmt_pct(time_progress)}（{days_passed}/{DAYS_IN_MONTH}天）\n"
                    f"🎯 实际达成：{fmt_pct(achieve_rate)}\n"
                    f"📐 应达GMV：{fmt_wan(expected_gmv)}　"
                    f"实际：{fmt_wan(total_gmv)}\n"
                    f"📌 差额：{fmt_wan(abs(gmv_gap))}　{gap_text}\n\n"
                ),
            }
        ]
    )

    # === 四、达人掉量预警 ===
    content.append([{"tag": "text", "text": "━ 达人掉量预警 Top5 ━\n"}])
    if top5_declines:
        for i, a in enumerate(top5_declines, 1):
            reasons = []
            if a["ad_cost"] <= 0:
                reasons.append("无投放")
            if a["settle_rate"] < 0.3:
                reasons.append(f"结算率低({fmt_pct(a['settle_rate'])})")
            if a["broadcast_days"] <= 2:
                reasons.append(f"仅开播{a['broadcast_days']}天")
            reason_str = "；".join(reasons) if reasons else "原因待确认"

            content.append(
                [
                    {
                        "tag": "text",
                        "text": (
                            f"{i}. {a['name']}（{a['agency']}）\n"
                            f"   支付：{fmt_wan(a['before'])} → {fmt_wan(a['yesterday'])}　"
                            f"↓{fmt_pct(a['decline_pct'])}（-{fmt_wan(a['decline'])}）\n"
                            f"   可能原因：{reason_str}\n"
                        ),
                    }
                ]
            )
    else:
        content.append([{"tag": "text", "text": "无明显掉量达人 ✅\n"}])
    content.append([{"tag": "text", "text": "\n"}])

    # === 五、机构昨日排名 ===
    content.append([{"tag": "text", "text": "━ 机构昨日GMV Top5 ━\n"}])
    medals = ["🥇", "🥈", "🥉", "④", "⑤"]
    for i, (ag, val) in enumerate(top5_agencies):
        content.append(
            [
                {
                    "tag": "text",
                    "text": f"{medals[i]} {ag}：{fmt_wan(val)}\n",
                }
            ]
    )

    return content


def send_to_feishu(content: list):
    """发送富文本消息到飞书"""
    if not FEISHU_WEBHOOK:
        print("❌ 未设置 FEISHU_WEBHOOK 环境变量，跳过推送")
        print("   用法: FEISHU_WEBHOOK='https://...' python3 daily_report.py")
        sys.exit(1)

    today_str = TODAY.strftime("%m月%d日")
    payload = {
        "msg_type": "post",
        "content": {
            "post": {
                "zh_cn": {
                    "title": f"📊 蕉下6月业绩日报 | {today_str}",
                    "content": content,
                }
            }
        },
    }

    resp = requests.post(FEISHU_WEBHOOK, json=payload, timeout=10)
    if resp.status_code == 200:
        result = resp.json()
        if result.get("code") == 0:
            print(f"✅ 日报已推送到飞书 ({TODAY.strftime('%Y-%m-%d %H:%M')})")
        else:
            print(f"❌ 飞书返回错误: {result}")
    else:
        print(f"❌ 请求失败: HTTP {resp.status_code} {resp.text}")


def main():
    print(f"📊 生成日报中... ({TODAY.strftime('%Y-%m-%d %H:%M')})")
    data = load_data()
    content = build_report(data)
    send_to_feishu(content)


if __name__ == "__main__":
    main()
