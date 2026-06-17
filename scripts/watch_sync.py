#!/usr/bin/env python3
"""
Excel 文件监控脚本
检测上游 Excel 文件变化 → 自动运行 sync_dashboard.py 刷新数据
"""

import os
import sys
import time
import subprocess
from datetime import datetime

# 配置
XLSX_PATH = '/Users/xiaocao/Desktop/蕉下文件/业绩追击/by月业绩/6月业绩/6月业绩追击（纯直播）.xlsx'
SYNC_SCRIPT = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'sync_dashboard.py')
CHECK_INTERVAL = 5   # 检查间隔（秒）
DEBOUNCE_SEC = 3     # 消抖等待（秒）

def log(msg):
    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    line = f'[{timestamp}] {msg}'
    print(line, flush=True)


def get_mtime(path):
    """获取文件修改时间，文件不存在返回 0"""
    try:
        return os.path.getmtime(path)
    except OSError:
        return 0


def run_sync():
    """执行数据同步"""
    log('Excel 已更新，开始同步数据...')
    result = subprocess.run(
        ['python3', SYNC_SCRIPT],
        capture_output=True, text=True, timeout=120
    )
    if result.returncode == 0:
        log('✓ 同步成功')
        for line in result.stdout.strip().split('\n'):
            log(f'  {line}')
    else:
        log(f'✗ 同步失败 (exit={result.returncode})')
        if result.stderr:
            for line in result.stderr.strip().split('\n'):
                log(f'  [ERR] {line}')


def main():
    log('看板数据监控已启动')
    log(f'  监控文件: {XLSX_PATH}')
    log(f'  同步脚本: {SYNC_SCRIPT}')
    log(f'  检查间隔: {CHECK_INTERVAL}s / 消抖: {DEBOUNCE_SEC}s')

    # 启动时先同步一次
    if os.path.exists(XLSX_PATH):
        run_sync()

    last_mtime = get_mtime(XLSX_PATH)

    while True:
        try:
            time.sleep(CHECK_INTERVAL)
            current_mtime = get_mtime(XLSX_PATH)

            if current_mtime > last_mtime:
                # 消抖：等待 Excel 完全写入
                time.sleep(DEBOUNCE_SEC)
                # 再次确认 mtime（防止是临时写入）
                final_mtime = get_mtime(XLSX_PATH)
                if final_mtime >= current_mtime:
                    run_sync()
                    last_mtime = final_mtime
                else:
                    last_mtime = current_mtime

        except KeyboardInterrupt:
            log('监控已停止')
            sys.exit(0)
        except Exception as e:
            log(f'监控异常: {e}')
            time.sleep(10)  # 出错后等久一点再重试


if __name__ == '__main__':
    main()
