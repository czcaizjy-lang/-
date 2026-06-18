#!/usr/bin/env python3
"""
Excel 文件监控脚本
检测上游 Excel 文件变化 → 自动运行 sync_dashboard.py 刷新数据
                                              → 运行 build_standalone.py 生成独立页面
                                              → git push 推送到 GitHub Pages
"""

import os
import sys
import time
import subprocess
from datetime import datetime

# 配置
XLSX_PATH = '/Users/xiaocao/Desktop/蕉下文件/业绩追击/by月业绩/6月业绩/6月业绩追击（纯直播）.xlsx'
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_DIR = os.path.dirname(SCRIPT_DIR)
SYNC_SCRIPT = os.path.join(SCRIPT_DIR, 'sync_dashboard.py')
BUILD_SCRIPT = os.path.join(SCRIPT_DIR, 'build_standalone.py')
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


def run_script(script_path):
    """执行一个 Python 脚本，返回 (success: bool, output_lines: list)"""
    result = subprocess.run(
        ['python3', script_path],
        capture_output=True, text=True, timeout=120,
        cwd=PROJECT_DIR
    )
    output_lines = result.stdout.strip().split('\n') if result.stdout.strip() else []
    if result.returncode == 0:
        for line in output_lines:
            log(f'  {line}')
        return True
    else:
        if result.stderr:
            for line in result.stderr.strip().split('\n'):
                log(f'  [ERR] {line}')
        return False


def git_push():
    """提交并推送到 GitHub，失败不抛异常"""
    try:
        # 只添加数据相关文件
        files = [
            'data/dashboard_data.json',
            'standalone_dashboard.html',
            'index.html',
        ]
        for f in files:
            path = os.path.join(PROJECT_DIR, f)
            if os.path.exists(path):
                subprocess.run(['git', 'add', f], cwd=PROJECT_DIR, capture_output=True)

        # 检查是否有改动
        status = subprocess.run(
            ['git', 'diff', '--cached', '--quiet'],
            cwd=PROJECT_DIR, capture_output=True
        )
        if status.returncode == 0:
            log('  无数据变更，跳过 git push')
            return

        # 提交
        now = datetime.now().strftime('%Y-%m-%d %H:%M')
        subprocess.run(
            ['git', 'commit', '-m', f'📊 数据自动更新 {now}'],
            cwd=PROJECT_DIR, capture_output=True
        )

        # 推送
        result = subprocess.run(
            ['git', 'push', 'origin', 'main'],
            cwd=PROJECT_DIR, capture_output=True, text=True, timeout=30
        )
        if result.returncode == 0:
            log('✓ 已推送到 GitHub Pages')
        else:
            log(f'✗ git push 失败: {result.stderr.strip()}')
    except Exception as e:
        log(f'✗ git 操作异常: {e}')


def run_sync():
    """执行数据同步 → 构建独立页面 → 推送到 GitHub"""
    log('Excel 已更新，开始同步...')

    # 第 1 步：数据同步
    log('--- 同步数据 ---')
    if not run_script(SYNC_SCRIPT):
        log('✗ 数据同步失败，跳过后续步骤')
        return

    # 第 2 步：构建独立页面
    log('--- 构建页面 ---')
    if not run_script(BUILD_SCRIPT):
        log('✗ 页面构建失败，跳过推送')
        return

    # 第 3 步：推送到 GitHub
    log('--- 推送部署 ---')
    git_push()

    log('✓ 一次完整同步完成')


def main():
    log('看板数据监控已启动（含自动部署）')
    log(f'  监控文件: {XLSX_PATH}')
    log(f'  同步脚本: {SYNC_SCRIPT}')
    log(f'  构建脚本: {BUILD_SCRIPT}')
    log(f'  检查间隔: {CHECK_INTERVAL}s / 消抖: {DEBOUNCE_SEC}s')
    log(f'  Git 远程: 检测中...')

    # 检查 git 状态
    git_remote = subprocess.run(
        ['git', 'remote', 'get-url', 'origin'],
        cwd=PROJECT_DIR, capture_output=True, text=True
    )
    if git_remote.returncode == 0:
        log(f'  Git 远程: {git_remote.stdout.strip()}')
    else:
        log('  ⚠ 未配置 Git 远程仓库，自动推送不可用')

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
