#!/usr/bin/env python3
"""
业绩看板数据提取脚本
直接读取上游 Excel 文件，生成 dashboard_data.json
"""

import json
import os
import subprocess as sp
from collections import defaultdict
from datetime import datetime
from openpyxl import load_workbook

# 直接读取上游文件
XLSX_PATH = '/Users/xiaocao/Desktop/蕉下文件/业绩追击/by月业绩/6月业绩/6月业绩追击（纯直播）.xlsx'
OUTPUT_PATH = '/Users/xiaocao/CC/每日业绩自动统计/data/dashboard_data.json'
MAIN_AGENCIES = ['自达号', '集米文化', '紫语', '花开满路', '太古', '亦初', '直属']

def run():
    wb = load_workbook(XLSX_PATH, data_only=True)

    # === 1. 提取达人汇总（星辞业绩 Sheet）===
    ws_anchor = wb['星辞业绩']
    anchors = []
    for r in range(3, ws_anchor.max_row + 1):
        douyin_id = ws_anchor.cell(r, 2).value
        if not douyin_id:
            continue
        anchors.append({
            '主播昵称': ws_anchor.cell(r, 1).value,
            '主播抖音号': douyin_id,
            '达人备注': ws_anchor.cell(r, 3).value,
            '佣金率': ws_anchor.cell(r, 4).value or 0,
            '团长费率': ws_anchor.cell(r, 5).value or 0,
            '预估固定费用': ws_anchor.cell(r, 6).value or 0,
            '预估毛利': ws_anchor.cell(r, 7).value or 0,
            '预测利润率': ws_anchor.cell(r, 8).value or 0,
            '机构': ws_anchor.cell(r, 9).value,
            '开播天数': ws_anchor.cell(r, 10).value or 0,
            '日均开播时长（小时）': ws_anchor.cell(r, 11).value or 0,
            '直播GMV': ws_anchor.cell(r, 12).value or 0,
            '直播退款GMV': ws_anchor.cell(r, 13).value or 0,
            '直播结算GMV': ws_anchor.cell(r, 14).value or 0,
            '预测直播结算GMV': ws_anchor.cell(r, 15).value or 0,
            '预测税后收入': ws_anchor.cell(r, 16).value or 0,
            '目前结算率': ws_anchor.cell(r, 17).value or 0,
            '4月结算率': ws_anchor.cell(r, 18).value,
            '佣金支出': ws_anchor.cell(r, 19).value or 0,
            '团长服务费': ws_anchor.cell(r, 20).value or 0,
            '投放消耗金额': ws_anchor.cell(r, 21).value or 0,
            '预估利润': ws_anchor.cell(r, 22).value or 0,
            '直播GMV占比': ws_anchor.cell(r, 23).value,
            '预估6月收入': ws_anchor.cell(r, 24).value,
            '当前收入达成率': ws_anchor.cell(r, 25).value,
        })

    # === 2. 提取汇总行 ===
    summary = {
        '主播昵称': '汇总',
        '直播GMV': ws_anchor.cell(2, 12).value or 0,
        '直播退款GMV': ws_anchor.cell(2, 13).value or 0,
        '直播结算GMV': ws_anchor.cell(2, 14).value or 0,
        '预测直播结算GMV': ws_anchor.cell(2, 15).value or 0,
        '预测税后收入': ws_anchor.cell(2, 16).value or 0,
        '目前结算率': ws_anchor.cell(2, 17).value or 0,
        '佣金支出': ws_anchor.cell(2, 19).value or 0,
        '团长服务费': ws_anchor.cell(2, 20).value or 0,
        '投放消耗金额': ws_anchor.cell(2, 21).value or 0,
        '预估利润': ws_anchor.cell(2, 22).value or 0,
        '预估6月收入': ws_anchor.cell(2, 24).value or 0,
        '当前收入达成率': ws_anchor.cell(2, 25).value or 0,
    }

    # === 3. 提取机构汇总（col 27-39）===
    agencies = []
    for r in range(3, ws_anchor.max_row + 1):
        agency_name = ws_anchor.cell(r, 27).value
        if not agency_name:
            continue
        agencies.append({
            '机构': agency_name,
            '机构达人数': ws_anchor.cell(r, 28).value or 0,
            '直播GMV': ws_anchor.cell(r, 29).value or 0,
            '人均直播GMV': ws_anchor.cell(r, 30).value or 0,
            '直播退款GMV': ws_anchor.cell(r, 31).value or 0,
            '直播结算GMV': ws_anchor.cell(r, 32).value or 0,
            '预测直播结算GMV': ws_anchor.cell(r, 33).value or 0,
            '预估结算率': ws_anchor.cell(r, 34).value or 0,
            '佣金支出': ws_anchor.cell(r, 35).value or 0,
            '团长服务费': ws_anchor.cell(r, 36).value or 0,
            '投放消耗金额': ws_anchor.cell(r, 37).value or 0,
            '预估净利润': ws_anchor.cell(r, 38).value or 0,
            '直播GMV占比': ws_anchor.cell(r, 39).value,
        })

    # === 3b. 提取自达号数据（星辞自达号业绩 Sheet）===
    ws_zdh = wb['星辞自达号业绩']
    ZDH_SUB_AGENCIES = ['花开自达号', '集米自达号', '太古自达号', '九三自达号', '直属自达号']

    # 3b-a. 自达号达人列表 + sub-org 映射
    zdh_id_to_sub = {}   # douyin_id -> sub_org_name
    zdh_anchors = []
    for r in range(3, ws_zdh.max_row + 1):
        douyin_id = ws_zdh.cell(r, 2).value
        if not douyin_id:
            continue
        douyin_id_str = str(douyin_id)
        sub_org = ws_zdh.cell(r, 10).value   # col 10 = 机构（子机构）
        zdh_id_to_sub[douyin_id_str] = sub_org
        # 结算率
        settle_rate = ws_zdh.cell(r, 21).value
        try:
            settle_rate = float(settle_rate) if settle_rate is not None else 0
        except (ValueError, TypeError):
            settle_rate = 0
        zdh_anchors.append({
            '主播昵称': ws_zdh.cell(r, 1).value,
            '主播抖音号': douyin_id_str,
            '机构': str(sub_org) if sub_org else '其他',  # 子机构名
            '开播天数': ws_zdh.cell(r, 11).value or 0,
            '日均开播时长（小时）': ws_zdh.cell(r, 12).value or 0,
            '直播GMV': ws_zdh.cell(r, 13).value or 0,
            '直播退款GMV': ws_zdh.cell(r, 27).value or 0,
            '直播结算GMV': ws_zdh.cell(r, 20).value or 0,
            '佣金支出': ws_zdh.cell(r, 22).value or 0,
            '团长服务费': ws_zdh.cell(r, 23).value or 0,
            '投放消耗金额': ws_zdh.cell(r, 25).value or 0,
            '目前结算率': settle_rate,
            'ROI': ws_zdh.cell(r, 26).value or 0,
            '直播GMV占比': ws_zdh.cell(r, 28).value,
        })

    # 3b-b. 自达号整体汇总（row 2）
    zdh_summary = {
        '直播GMV': ws_zdh.cell(2, 13).value or 0,
        '直播结算GMV': ws_zdh.cell(2, 20).value or 0,
        '消耗金额': ws_zdh.cell(2, 25).value or 0,
        'ROI': ws_zdh.cell(2, 26).value or 0,
        '直播退款GMV': ws_zdh.cell(2, 27).value or 0,
    }

    # 3b-c. 子机构汇总（cols 31-47, rows 3-7）
    zdh_sub_agencies = []
    for r in range(3, ws_zdh.max_row + 1):
        agency_name = ws_zdh.cell(r, 31).value  # col 31 = 机构
        if not agency_name or str(agency_name).strip() not in ZDH_SUB_AGENCIES:
            continue
        zdh_sub_agencies.append({
            '机构': str(agency_name).strip(),
            '机构达人数': ws_zdh.cell(r, 32).value or 0,
            '直播GMV': ws_zdh.cell(r, 33).value or 0,
            '人均直播GMV': ws_zdh.cell(r, 34).value or 0,
            '直播退款GMV': ws_zdh.cell(r, 35).value or 0,
            '直播结算GMV': ws_zdh.cell(r, 36).value or 0,
            '投放消耗金额': ws_zdh.cell(r, 40).value or 0,
            'ROI': ws_zdh.cell(r, 42).value or 0,
            '直播GMV占比': ws_zdh.cell(r, 47).value,
        })

    # === 4. 从 6月直播数据 提取分天趋势 ===
    ws_live = wb['6月直播数据']
    id_to_info = {str(a['主播抖音号']): a for a in anchors}

    daily_total_gmv = defaultdict(float)
    daily_total_paid = defaultdict(float)
    daily_total_refund = defaultdict(float)
    daily_by_agency = defaultdict(lambda: defaultdict(float))
    anchor_daily_paid = defaultdict(lambda: defaultdict(float))
    anchor_daily_gmv = defaultdict(lambda: defaultdict(float))
    anchor_daily_refund = defaultdict(lambda: defaultdict(float))
    anchor_daily_ad_cost = defaultdict(lambda: defaultdict(float))

    # 自达号专属每日累加器
    zdh_daily_gmv = defaultdict(float)
    zdh_daily_paid = defaultdict(float)
    zdh_daily_refund = defaultdict(float)
    zdh_daily_by_sub = defaultdict(lambda: defaultdict(float))
    zdh_anchor_daily_gmv = defaultdict(lambda: defaultdict(float))
    zdh_anchor_daily_paid = defaultdict(lambda: defaultdict(float))
    zdh_anchor_daily_ad_cost = defaultdict(lambda: defaultdict(float))

    for r in range(2, ws_live.max_row + 1):
        douyin_id_raw = ws_live.cell(r, 3).value
        dt_val = ws_live.cell(r, 4).value
        gmv = float(ws_live.cell(r, 26).value or 0)
        paid = float(ws_live.cell(r, 27).value or 0)
        refund = float(ws_live.cell(r, 32).value or 0)
        ad_cost_bind = float(ws_live.cell(r, 44).value or 0)    # 投放消耗(店铺绑定)
        ad_cost_beitou = float(ws_live.cell(r, 45).value or 0)  # 投放消耗(店铺被投)
        ad_cost = ad_cost_bind if ad_cost_bind >= ad_cost_beitou else ad_cost_beitou
        if not douyin_id_raw or not dt_val:
            continue
        douyin_id = str(douyin_id_raw)
        date_key = str(dt_val)[:10].replace('/', '-')
        daily_total_gmv[date_key] += gmv
        daily_total_paid[date_key] += paid
        daily_total_refund[date_key] += refund
        info = id_to_info.get(douyin_id, {})
        agency = info.get('机构', '其他')
        daily_by_agency[agency][date_key] += gmv
        anchor_daily_paid[douyin_id][date_key] += paid
        anchor_daily_gmv[douyin_id][date_key] += gmv
        anchor_daily_refund[douyin_id][date_key] += refund
        anchor_daily_ad_cost[douyin_id][date_key] += ad_cost

        # 自达号专属累加
        if douyin_id in zdh_id_to_sub:
            zdh_daily_gmv[date_key] += gmv
            zdh_daily_paid[date_key] += paid
            zdh_daily_refund[date_key] += refund
            sub_org = zdh_id_to_sub[douyin_id]
            if sub_org:
                zdh_daily_by_sub[str(sub_org)][date_key] += gmv
            zdh_anchor_daily_gmv[douyin_id][date_key] += gmv
            zdh_anchor_daily_paid[douyin_id][date_key] += paid
            zdh_anchor_daily_ad_cost[douyin_id][date_key] += ad_cost

    # === 4b. 从日数据汇总月度消耗（用 max(绑定,被投) 口径）并回写达人/机构/汇总 ===
    # 汇总每个达人的月度消耗
    monthly_ad_by_anchor = {}
    for douyin_id, daily in anchor_daily_ad_cost.items():
        monthly_ad_by_anchor[douyin_id] = round(sum(daily.values()), 2)

    # 回写主达人列表的投放消耗金额 + 补充 ROI
    for a in anchors:
        douyin = str(a['主播抖音号'])
        new_ad = monthly_ad_by_anchor.get(douyin, 0)
        a['投放消耗金额'] = new_ad
        gmv = float(a.get('直播GMV', 0) or 0)
        a['ROI'] = round(gmv / new_ad, 2) if new_ad > 0 else 0

    # 回写自达号达人列表
    zdh_monthly_ad = {}
    for douyin_id, daily in zdh_anchor_daily_ad_cost.items():
        zdh_monthly_ad[douyin_id] = round(sum(daily.values()), 2)
    for a in zdh_anchors:
        douyin = str(a['主播抖音号'])
        new_ad = zdh_monthly_ad.get(douyin, 0)
        a['投放消耗金额'] = new_ad
        gmv = float(a.get('直播GMV', 0) or 0)
        a['ROI'] = round(gmv / new_ad, 2) if new_ad > 0 else 0

    # 回写汇总行（所有达人月度消耗之和）
    total_monthly_ad = sum(monthly_ad_by_anchor.values())
    summary['投放消耗金额'] = round(total_monthly_ad, 2)

    # 回写机构汇总
    agency_ad_sum = defaultdict(float)
    for a in anchors:
        agency = str(a.get('机构', '') or '')
        agency_ad_sum[agency] += float(a.get('投放消耗金额', 0) or 0)
    for ag in agencies:
        ag_name = str(ag.get('机构', '') or '')
        new_ad = round(agency_ad_sum.get(ag_name, 0), 2)
        ag['投放消耗金额'] = new_ad

    # 回写自达号汇总
    zdh_total_ad = sum(zdh_monthly_ad.values())
    zdh_summary['消耗金额'] = round(zdh_total_ad, 2)
    zdh_gmv = float(zdh_summary.get('直播GMV', 0) or 0)
    zdh_summary['ROI'] = round(zdh_gmv / zdh_total_ad, 2) if zdh_total_ad > 0 else 0

    # 回写自达号子机构汇总
    zdh_sub_ad_sum = defaultdict(float)
    for a in zdh_anchors:
        sub = str(a.get('机构', '') or '')
        zdh_sub_ad_sum[sub] += float(a.get('投放消耗金额', 0) or 0)
    for sa in zdh_sub_agencies:
        sub_name = str(sa.get('机构', '') or '').strip()
        new_ad = round(zdh_sub_ad_sum.get(sub_name, 0), 2)
        sa['投放消耗金额'] = new_ad
        sa_gmv = float(sa.get('直播GMV', 0) or 0)
        sa['ROI'] = round(sa_gmv / new_ad, 2) if new_ad > 0 else 0

    # === 5. 整理日期 ===
    all_dates = sorted(daily_total_gmv.keys())
    trend_dates = [d[5:].replace('-', '/') for d in all_dates]
    date_map = {all_dates[i]: trend_dates[i] for i in range(len(all_dates))}

    def to_wan(d):
        return {date_map[k]: round(v / 10000, 2) for k, v in d.items() if k in date_map}

    # === 6. 提取 agency_top5_anchors ===
    agency_top5_anchors = {}
    for agency in MAIN_AGENCIES:
        a_list = [(did, info['直播GMV']) for did, info in id_to_info.items() if info.get('机构') == agency]
        a_list.sort(key=lambda x: x[1], reverse=True)
        top5 = []
        for douyin_id, _ in a_list[:5]:
            name = id_to_info[douyin_id]['主播昵称']
            daily = anchor_daily_paid.get(douyin_id, {})
            values_wan = [round(daily.get(full, 0) / 10000, 2) for full, _ in date_map.items()]
            top5.append({'name': name, 'douyin_id': douyin_id, 'daily_paid': values_wan})
        agency_top5_anchors[agency] = top5

    # === 6b. 自达号子机构 Top5 达人 ===
    zdh_id_to_info = {str(a['主播抖音号']): a for a in zdh_anchors}
    zdh_top5_by_sub = {}
    for sub_org in ZDH_SUB_AGENCIES:
        a_list = [(did, info['直播GMV']) for did, info in zdh_id_to_info.items()
                  if str(info.get('机构', '')).strip() == sub_org]
        a_list.sort(key=lambda x: x[1], reverse=True)
        top5 = []
        for douyin_id, _ in a_list[:5]:
            name = zdh_id_to_info[douyin_id]['主播昵称']
            daily_paid = zdh_anchor_daily_paid.get(douyin_id, {})
            values_wan = [round(daily_paid.get(full, 0) / 10000, 2) for full, _ in date_map.items()]
            top5.append({'name': str(name), 'douyin_id': douyin_id, 'daily_paid': values_wan})
        zdh_top5_by_sub[sub_org] = top5

    # === 7. 提取 person_daily（久酒/雅宁/星辞）===
    person_sheet_names = ['久酒业绩', '雅宁业绩', '星辞业绩']
    person_daily = {'dates': trend_dates}
    person_daily_gmv = {'dates': trend_dates}
    person_daily_refund = {'dates': trend_dates}
    for sheet_name in person_sheet_names:
        if sheet_name not in wb.sheetnames:
            continue
        ws = wb[sheet_name]
        person_douyin_ids = set()
        for r in range(2, ws.max_row + 1):
            did = ws.cell(r, 2).value
            if did:
                person_douyin_ids.add(str(did))
        if not person_douyin_ids:
            continue
        person_name = sheet_name.replace('业绩', '')
        paid_values = []
        gmv_values = []
        refund_values = []
        for full in all_dates:
            day_paid = sum(anchor_daily_paid.get(did, {}).get(full, 0) for did in person_douyin_ids)
            day_gmv = sum(anchor_daily_gmv.get(did, {}).get(full, 0) for did in person_douyin_ids)
            day_refund = sum(anchor_daily_refund.get(did, {}).get(full, 0) for did in person_douyin_ids)
            paid_values.append(round(day_paid / 10000, 2))
            gmv_values.append(round(day_gmv / 10000, 2))
            refund_values.append(round(day_refund / 10000, 2))
        person_daily[person_name] = paid_values
        person_daily_gmv[person_name] = gmv_values
        person_daily_refund[person_name] = refund_values

    # === 7b. 自达号子机构每日 ROI（加权：ΣGMV / Σ消耗）===
    zdh_daily_roi_by_sub = {}
    zdh_anchor_ids_by_sub = defaultdict(list)
    for douyin_id, sub in zdh_id_to_sub.items():
        zdh_anchor_ids_by_sub[sub].append(douyin_id)

    for sub in ZDH_SUB_AGENCIES:
        anchors_in_sub = zdh_anchor_ids_by_sub.get(sub, [])
        roi_values = []
        for full in all_dates:
            total_gmv = sum(zdh_anchor_daily_gmv.get(aid, {}).get(full, 0) for aid in anchors_in_sub)
            total_ad = sum(zdh_anchor_daily_ad_cost.get(aid, {}).get(full, 0) for aid in anchors_in_sub)
            roi = round(total_gmv / total_ad, 2) if total_ad > 0 else None
            roi_values.append(roi)
        zdh_daily_roi_by_sub[sub] = roi_values

    # === 8. 构建最终 JSON ===
    dashboard_data = {
        'summary': summary,
        'anchors': anchors,
        'agencies': agencies,
        'trend': {
            'dates': trend_dates,
            'daily_total_gmv': [round(daily_total_gmv[d] / 10000, 2) for d in all_dates],
            'daily_total_paid': [round(daily_total_paid[d] / 10000, 2) for d in all_dates],
            'daily_total_refund': [round(daily_total_refund[d] / 10000, 2) for d in all_dates],
            'agencies': list(daily_by_agency.keys()),
            'daily_by_agency': {
                ag: [round(daily_by_agency[ag][d] / 10000, 2) for d in all_dates]
                for ag in daily_by_agency
            },
            'agency_totals': sum(daily_by_agency[ag][d] for ag in daily_by_agency for d in all_dates),
            'pie_data': [
                {'name': ag, 'value': round(sum(daily_by_agency[ag].values()) / 10000, 2)}
                for ag in daily_by_agency
            ],
            'person_daily': person_daily,
            'person_daily_gmv': person_daily_gmv,
            'person_daily_refund': person_daily_refund,
            'anchor_daily_paid': {
                douyin_id: to_wan(by_date)
                for douyin_id, by_date in anchor_daily_paid.items()
            },
            'anchor_daily_roi': {
                douyin_id: {
                    date_map[full]: round(
                        anchor_daily_gmv[douyin_id][full] / anchor_daily_ad_cost[douyin_id][full], 2
                    ) if anchor_daily_ad_cost[douyin_id][full] > 0 else 0
                    for full in anchor_daily_gmv[douyin_id]
                    if full in date_map
                }
                for douyin_id in anchor_daily_gmv
            },
            'agency_top5_anchors': agency_top5_anchors,
        },
        'zidahao': {
            'summary': zdh_summary,
            'sub_agencies': zdh_sub_agencies,
            'anchors': zdh_anchors,
            'daily_gmv': [round(zdh_daily_gmv[d] / 10000, 2) for d in all_dates],
            'daily_paid': [round(zdh_daily_paid[d] / 10000, 2) for d in all_dates],
            'daily_refund': [round(zdh_daily_refund[d] / 10000, 2) for d in all_dates],
            'daily_by_sub': {
                sub: [round(zdh_daily_by_sub[sub][d] / 10000, 2) for d in all_dates]
                for sub in ZDH_SUB_AGENCIES
            },
            'pie_data': [
                {
                    'name': sub,
                    'value': round(sum(zdh_daily_by_sub[sub].values()) / 10000, 2)
                }
                for sub in ZDH_SUB_AGENCIES
            ],
            'top5_by_sub': zdh_top5_by_sub,
            'daily_roi_by_sub': zdh_daily_roi_by_sub,
        },
    }

    with open(OUTPUT_PATH, 'w', encoding='utf-8') as f:
        json.dump(dashboard_data, f, ensure_ascii=False, indent=2)

    print(f"✓ 数据已更新到 {OUTPUT_PATH}")
    print(f"  - 汇总 GMV: ¥{summary['直播GMV']:,.2f}")
    print(f"  - 达人数量: {len(anchors)}")
    print(f"  - 机构数量: {len(agencies)}")
    print(f"  - 日期范围: {trend_dates[0]} ~ {trend_dates[-1]}")
    print(f"  - 自达号 GMV: ¥{zdh_summary['直播GMV']:,.2f} ({len(zdh_anchors)} 达人, {len(zdh_sub_agencies)} 子机构)")

    # 自动构建独立看板页面（如果构建脚本存在）
    script_dir = os.path.dirname(os.path.abspath(__file__))
    build_script = os.path.join(script_dir, 'build_standalone.py')
    if os.path.exists(build_script):
        print('')
        sp.run(['python3', build_script], cwd=script_dir)

if __name__ == '__main__':
    run()
