"""
清除 Chrome 占用 - 自动化脚本
=====================================
无额外依赖（仅使用标准库）。

清理残留的 Chrome/Chromium 进程，释放 CDP 调试端口和 user-data-dir 目录锁，
供营业统计等自动化任务在启动前/异常退出后手动清场使用。

用法：
    python cleanup_stale_browsers.py                        # 清理默认端口与 ChromeDebug 进程
    python cleanup_stale_browsers.py --ports 9001 9004      # 只清理指定端口
    python cleanup_stale_browsers.py --keyword ChromeDebug_BT  # 只清理指定 profile 关键字
    python cleanup_stale_browsers.py --no-wait              # 清理后不等待资源释放
"""

import argparse
import json
import logging
import logging.handlers
import subprocess
import sys
import time
from pathlib import Path

# ─── 日志配置 ───────────────────────────────────────────────────────────────────
LOG_DIR = Path(__file__).resolve().parent / "log"
LOG_DIR.mkdir(exist_ok=True)

_file_handler = logging.handlers.TimedRotatingFileHandler(
    LOG_DIR / "cleanup_stale_browsers.log",
    when="midnight",
    backupCount=30,
    encoding="utf-8",
)
_file_handler.suffix = "%Y-%m-%d"

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[_file_handler, logging.StreamHandler()],
)
logger = logging.getLogger(__name__)

# ─── 配置区（与 mainyan_turnover_statistics.py 保持一致）───────────────────────
# 9001 宝泰店 / 9002 龙江店 / 9003 杏坛店 / 9004 共享 Chrome（经营宝·招行·抖音·口碑）
CDP_PORTS = {9001, 9002, 9003, 9004}

# user-data-dir 关键字：C:\ChromeDebug_BT、_LJ、_XT、_MTJYB
PROFILE_KEYWORD = "ChromeDebug"

# 释放资源的等待秒数
RELEASE_WAIT = 3


def cleanup_stale_browsers(ports=None, keyword=PROFILE_KEYWORD, wait=True):
    """清理残留的 Chrome/Chromium 进程，避免端口冲突和 user-data-dir 锁。

    返回被终止的进程 PID 集合。
    """
    ports = CDP_PORTS if ports is None else set(ports)
    killed_pids = set()

    # 1. 通过 netstat 查找占用 CDP 端口的进程
    try:
        output = subprocess.check_output(
            ["netstat", "-ano"], text=True, errors="replace"
        )
        for line in output.splitlines():
            for port in ports:
                if f":{port} " in line and "LISTENING" in line:
                    parts = line.split()
                    try:
                        pid = int(parts[-1])
                    except (ValueError, IndexError):
                        continue
                    if pid > 0 and pid not in killed_pids:
                        logger.info(f"  端口 {port} 被 PID {pid} 占用，终止进程树...")
                        subprocess.run(
                            ["taskkill", "/F", "/T", "/PID", str(pid)],
                            capture_output=True, timeout=10,
                        )
                        killed_pids.add(pid)
    except Exception as e:
        logger.warning(f"  检查端口占用失败: {e}")

    # 2. 查找使用 ChromeDebug 目录的 chrome.exe（可能端口未绑定但锁住了 profile）
    try:
        # 使用 PowerShell 而不是 wmic（wmic 在某些系统上不可用）
        ps_script = f"""
        Get-WmiObject Win32_Process -Filter "name='chrome.exe'" |
        Where-Object {{ $_.CommandLine -like '*{keyword}*' }} |
        Select-Object ProcessId, CommandLine |
        ConvertTo-Json
        """
        output = subprocess.check_output(
            ["powershell", "-NoProfile", "-Command", ps_script],
            text=True, errors="replace", timeout=15,
        )
        if output.strip() and output.strip() != "null":
            processes = json.loads(output)
            # 确保 processes 是列表
            if isinstance(processes, dict):
                processes = [processes]
            for proc in processes:
                pid = proc.get("ProcessId")
                if pid and pid > 0 and pid not in killed_pids:
                    logger.info(f"  发现 {keyword} chrome.exe PID {pid}，终止进程树...")
                    subprocess.run(
                        ["taskkill", "/F", "/T", "/PID", str(pid)],
                        capture_output=True, timeout=10,
                    )
                    killed_pids.add(pid)
    except Exception as e:
        logger.warning(f"  检查 {keyword} 进程失败: {e}")

    # 3. 清理 Playwright 残留的 chromium.exe
    try:
        output = subprocess.check_output(
            ["tasklist", "/FI", "IMAGENAME eq chromium.exe"],
            text=True, errors="replace",
        )
        for line in output.splitlines():
            if "chromium.exe" in line.lower():
                parts = line.split()
                try:
                    pid = int(parts[1])
                except (ValueError, IndexError):
                    continue
                if pid not in killed_pids:
                    logger.info(f"  发现残留 chromium.exe PID {pid}，终止进程树...")
                    subprocess.run(
                        ["taskkill", "/F", "/T", "/PID", str(pid)],
                        capture_output=True, timeout=10,
                    )
                    killed_pids.add(pid)
    except Exception as e:
        logger.warning(f"  检查 chromium.exe 进程失败: {e}")

    if killed_pids:
        if wait:
            logger.info(f"  已清理 {len(killed_pids)} 个残留进程，等待 {RELEASE_WAIT} 秒释放资源...")
            time.sleep(RELEASE_WAIT)
        else:
            logger.info(f"  已清理 {len(killed_pids)} 个残留进程")
    else:
        logger.info("  未发现残留浏览器进程")

    return killed_pids


def main():
    parser = argparse.ArgumentParser(description="清除自动化任务占用的 Chrome 进程")
    parser.add_argument(
        "--ports", type=int, nargs="*", default=None,
        help=f"要清理的 CDP 端口，默认 {sorted(CDP_PORTS)}；传空表示跳过端口检查",
    )
    parser.add_argument(
        "--keyword", default=PROFILE_KEYWORD,
        help=f"chrome.exe 命令行中的 user-data-dir 关键字，默认 {PROFILE_KEYWORD}",
    )
    parser.add_argument(
        "--no-wait", action="store_true",
        help="清理后不等待资源释放",
    )
    args = parser.parse_args()

    logger.info(f"{'=' * 55}")
    logger.info("  清除 Chrome 占用")
    logger.info(f"  目标端口：{sorted(args.ports) if args.ports is not None else sorted(CDP_PORTS)}")
    logger.info(f"  profile 关键字：{args.keyword}")
    logger.info(f"{'=' * 55}")

    try:
        killed = cleanup_stale_browsers(
            ports=args.ports, keyword=args.keyword, wait=not args.no_wait
        )
    except Exception as e:
        logger.error(f"清理失败: {e}")
        sys.exit(1)

    logger.info(f"{'=' * 55}")
    logger.info(f"  清理完成，共终止 {len(killed)} 个进程")
    logger.info(f"{'=' * 55}")


if __name__ == "__main__":
    main()
