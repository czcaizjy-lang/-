#!/bin/bash
# 每日飞书日报推送脚本（供 launchd 调用）
export FEISHU_WEBHOOK="https://open.feishu.cn/open-apis/bot/v2/hook/c44a79bd-5386-4145-ac27-50dda5a6a7fa"
exec /usr/bin/python3 /Users/xiaocao/CC/每日业绩自动统计/scripts/daily_report.py >> /tmp/daily_report.log 2>&1
