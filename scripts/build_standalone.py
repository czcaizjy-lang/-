#!/usr/bin/env python3
"""
构建独立看板 HTML 文件
读取 dashboard.html（模板）和 dashboard_data.json（数据），
生成内嵌数据的 public/index.html（GitHub Pages 入口）。
"""

import json
import os
import sys
from datetime import datetime

# 路径配置
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TEMPLATE_PATH = os.path.join(BASE_DIR, 'dashboard.html')
DATA_PATH = os.path.join(BASE_DIR, 'data', 'dashboard_data.json')
OUTPUT_PATHS = [
    os.path.join(BASE_DIR, 'index.html'),
    os.path.join(BASE_DIR, 'public', 'index.html'),
]

# ---- fetch 替换文本 ----
OLD_FETCH = """fetch('data/dashboard_data.json?t=' + Date.now()).then(r => r.json()).then(d => {
  DATA = d;
  render();
}).catch(e => { console.error('Failed to load dashboard data:', e); });"""


def build():
    # 读取模板
    with open(TEMPLATE_PATH, 'r', encoding='utf-8') as f:
        template = f.read()

    if OLD_FETCH not in template:
        print('✗ 模板中未找到 fetch 语句，可能已经变更', file=sys.stderr)
        return False

    # 读取数据
    with open(DATA_PATH, 'r', encoding='utf-8') as f:
        data = json.load(f)

    data_str = json.dumps(data, ensure_ascii=False, separators=(',', ':'))

    # 拆分为三个脚本块：变量声明 → 内嵌数据 → 业务逻辑
    parts = template.split(OLD_FETCH, 1)
    before = parts[0]   # 含 <script> 开头 + 变量/主题/switchDataSource + IIFE
    after = parts[1]    # fmt / render / 图表 / 表格 / 机构 / 下钻等所有业务逻辑

    inline_block = (
        '</script>\n'
        '<script id="inline-data" type="application/json">' + data_str + '</script>\n'
        '<script>\n'
        'DATA = JSON.parse(document.getElementById("inline-data").textContent);\n'
        'setTimeout(render, 0);'
    )

    output = before + inline_block + after

    # 写入所有输出路径
    for path in OUTPUT_PATHS:
        with open(path, 'w', encoding='utf-8') as f:
            f.write(output)

    # 打印统计
    summary = data.get('summary', {})
    gmv = summary.get('直播GMV', 0)
    anchors = len(data.get('anchors', []))
    zdh_anchors = len(data.get('zidahao', {}).get('anchors', []))
    size_kb = len(output) / 1024

    print(f'✓ 已生成 {len(OUTPUT_PATHS)} 个文件')
    for p in OUTPUT_PATHS:
        print(f'  → {os.path.basename(p)}')
    print(f'  模板: {os.path.basename(TEMPLATE_PATH)} ({len(template)} 字符)')
    print(f'  数据: {os.path.basename(DATA_PATH)} ({len(data_str)} 字符)')
    print(f'  输出: {size_kb:.0f} KB')
    print(f'  总GMV: ¥{gmv:,.2f} | 总达人: {anchors} | 自达号: {zdh_anchors}')
    return True


if __name__ == '__main__':
    ok = build()
    sys.exit(0 if ok else 1)
