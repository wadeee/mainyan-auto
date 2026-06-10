"""
麦安研营业统计 - 自动化脚本
=====================================
依赖：pip install playwright && playwright install chromium

自动登录 Pospal 后台，导出麦安研各门店的营业概况日度统计数据。

用法：
    python mainyan_turnover_statistics.py                    # 导出今天数据
    python mainyan_turnover_statistics.py --days -1          # 导出昨天数据
    python mainyan_turnover_statistics.py --date 2026.06.03  # 导出指定日期
    python mainyan_turnover_statistics.py --headless          # 无头模式（不显示浏览器）
"""

import argparse
import calendar
import copy
import json
import logging
import logging.handlers
import re
import sys
import time
from datetime import datetime, timedelta
from pathlib import Path

# ─── 日志配置 ───────────────────────────────────────────────────────────────────
LOG_DIR = Path(__file__).resolve().parent / "log"
LOG_DIR.mkdir(exist_ok=True)

_file_handler = logging.handlers.TimedRotatingFileHandler(
    LOG_DIR / "mainyan_turnover_statistics.log",
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

# ─── 配置区（按需修改）─────────────────────────────────────────────────────────
ACCOUNT = "huomimayzb"
WORKER_ID = "M006"
PASSWORD = "tusijia88"

LOGIN_URL = "https://beta69.pospal.cn/"
BUSINESS_SUMMARY_URL = "https://beta69.pospal.cn/Report/BusinessSummaryV2"

OUTPUT_DIR = Path(__file__).resolve().parent / "麦安研营业统计"
TEMPLATE_FILE = Path(__file__).resolve().parent / "麦安研营业统计_格式化模板.xlsx"

STORES = [
    {"full": "3 - 麦安研（东站宝泰店）", "short": "宝泰店", "template_name": "东站宝泰店"},
    {"full": "5 - 麦安研（顺德龙江店）", "short": "龙江店", "template_name": "顺德龙江店"},
    {"full": "2 - 麦安研（顺德杏坛店）", "short": "杏坛店", "template_name": "顺德杏坛店"},
]

TEMPLATE_STORE_NAME = "东方宝泰店"

WEEKDAY_NAMES = ["星期一", "星期二", "星期三", "星期四", "星期五", "星期六", "星期日"]


# ─── 工具函数 ──────────────────────────────────────────────────────────────────


def get_target_date(days: int = 0, date_str: str = None):
    if date_str:
        return datetime.strptime(date_str, "%Y.%m.%d")
    return datetime.now() + timedelta(days=days)


# ─── 格式化合并函数 ──────────────────────────────────────────────────────────


def copy_cell_style(src_cell, dst_cell):
    if src_cell.has_style:
        dst_cell.font = copy.copy(src_cell.font)
        dst_cell.fill = copy.copy(src_cell.fill)
        dst_cell.border = copy.copy(src_cell.border)
        dst_cell.alignment = copy.copy(src_cell.alignment)
        dst_cell.number_format = src_cell.number_format


def read_business_summary_c16(data_file: Path):
    from openpyxl import load_workbook as _lwb
    wb = _lwb(data_file, data_only=True)
    ws = wb.active
    val = ws.cell(row=16, column=3).value
    wb.close()
    if val is None:
        return None
    try:
        return float(str(val).strip())
    except (ValueError, TypeError):
        return val


def create_or_open_monthly_file(store, target: datetime, template_file: Path, output_dir: Path):
    year = target.year
    month = target.month
    month_str = f"{year}-{month:02d}"
    month_dir = output_dir / month_str
    month_dir.mkdir(parents=True, exist_ok=True)

    store_name = store["template_name"]
    output_file = month_dir / f"麦安研营业统计_{store_name}_{month_str}.xlsx"

    if output_file.exists():
        logger.info(f"  月度文件已存在，直接打开: {output_file.name}")
        return output_file, False

    logger.info(f"  月度文件不存在，从模板创建: {output_file.name}")
    _create_monthly_from_template(template_file, output_file, store_name, year, month)
    return output_file, True


def _create_monthly_from_template(template_file: Path, output_file: Path, store_name: str, year: int, month: int):
    from openpyxl import load_workbook as _lwb

    wb = _lwb(template_file)
    days_in_month = calendar.monthrange(year, month)[1]

    ws1 = wb.worksheets[0]
    ws2 = wb.worksheets[1]

    old_title_1 = ws1.title
    old_title_2 = ws2.title
    new_title_1 = old_title_1.replace(TEMPLATE_STORE_NAME, store_name).replace("2026年6月", f"{year}年{month}月")
    new_title_2 = old_title_2.replace(TEMPLATE_STORE_NAME, store_name).replace("2026年6月", f"{year}年{month}月")
    ws1.title = new_title_1
    ws2.title = new_title_2

    a1_val_1 = ws1.cell(row=1, column=1).value
    if a1_val_1 and isinstance(a1_val_1, str):
        ws1.cell(row=1, column=1).value = a1_val_1.replace(TEMPLATE_STORE_NAME, store_name).replace("2026年6月", f"{year}年{month}月")

    a1_val_2 = ws2.cell(row=1, column=1).value
    if a1_val_2 and isinstance(a1_val_2, str):
        ws2.cell(row=1, column=1).value = a1_val_2.replace(TEMPLATE_STORE_NAME, store_name).replace("2026年6月", f"{year}年{month}月")

    for day in range(1, 32):
        row_s1 = day + 4
        if day <= days_in_month:
            dt = datetime(year, month, day)
            ws1.cell(row=row_s1, column=2).value = f"{year}.{month}.{day}"
            ws1.cell(row=row_s1, column=3).value = WEEKDAY_NAMES[dt.weekday()]
        else:
            ws1.cell(row=row_s1, column=2).value = None
            ws1.cell(row=row_s1, column=3).value = None

    for day in range(1, 32):
        row_s2 = day + 2
        if day <= days_in_month:
            dt = datetime(year, month, day)
            ws2.cell(row=row_s2, column=1).value = f"{year}.{month}.{day}"
            ws2.cell(row=row_s2, column=2).value = WEEKDAY_NAMES[dt.weekday()]
        else:
            ws2.cell(row=row_s2, column=1).value = None
            ws2.cell(row=row_s2, column=2).value = None

    output_file.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_file)
    wb.close()
    logger.info(f"  已创建月度文件: {output_file}")


def fill_daily_data(monthly_file: Path, target: datetime, store, daily_download_dir: Path):
    from openpyxl import load_workbook as _lwb

    day = target.day
    date_label = target.strftime("%Y-%m-%d")

    summary_file = daily_download_dir / f"营业概况_{store['short']}_{date_label}.xlsx"
    if not summary_file.exists():
        logger.warning(f"  营业概况文件不存在，跳过 G 列填写: {summary_file.name}")
        c16_value = None
    else:
        c16_value = read_business_summary_c16(summary_file)
        logger.info(f"  读取 C16 (现金支付) = {c16_value}")

    wb = _lwb(monthly_file)
    ws1 = wb.worksheets[0]

    row_s1 = day + 4
    if c16_value is not None:
        ws1.cell(row=row_s1, column=7).value = c16_value

    wb.save(monthly_file)
    wb.close()
    logger.info(f"  已将第 {day} 天的数据写入: {monthly_file.name}")


# ─── 浏览器自动化函数 ────────────────────────────────────────────────────────


def set_date(page, placeholder: str, value: str):
    page.evaluate(f"""
        (function() {{
            var inp = document.querySelector('input.timeInput.hasDatepicker[placeholder="{placeholder}"]');
            if (!inp) return 'not found';
            inp.value = "{value}";
            inp.dispatchEvent(new Event('input',  {{bubbles: true}}));
            inp.dispatchEvent(new Event('change', {{bubbles: true}}));
            inp.dispatchEvent(new Event('blur',   {{bubbles: true}}));
            return inp.value;
        }})()
    """)


def click_by_text(page, text: str, desc: str = ""):
    result = page.evaluate(f"""
        (function() {{
            var els = document.querySelectorAll('*');
            for (var i = 0; i < els.length; i++) {{
                if (els[i].textContent.trim() === "{text}" && els[i].children.length === 0) {{
                    els[i].click();
                    return 'clicked tag=' + els[i].tagName + ' class=' + els[i].className;
                }}
            }}
            return 'not found';
        }})()
    """)
    tag = f"[{desc}] " if desc else ""
    logger.info(f"  {tag}→ {result}")
    return "clicked" in result


def login(page):
    logger.info("[1/4] 打开登录页面...")
    page.goto(LOGIN_URL)
    page.wait_for_load_state("networkidle")

    logger.info("[2/4] 切换到工号登录模式...")
    page.evaluate("""
        (function() {
            var els = document.querySelectorAll('div,span,a,button');
            for (var i = 0; i < els.length; i++) {
                if (els[i].textContent.trim() === '工号登录' && els[i].children.length === 0) {
                    els[i].click();
                    return 'clicked';
                }
            }
            return 'not found';
        })()
    """)
    time.sleep(1.5)

    placeholder = page.evaluate("""
        (function() {
            var inputs = document.querySelectorAll('input');
            var r = [];
            for (var i = 0; i < inputs.length; i++) {
                if (inputs[i].placeholder) r.push(inputs[i].placeholder);
            }
            return r.join(' | ');
        })()
    """)
    if "员工工号" not in placeholder:
        raise RuntimeError(f"工号登录切换失败，当前输入框：{placeholder}")
    logger.info(f"  表单已切换 → {placeholder}")

    logger.info("[3/4] 填入账号/工号/密码...")
    page.evaluate(f"""
        (function() {{
            var inputs = document.querySelectorAll('input');
            for (var i = 0; i < inputs.length; i++) {{
                if (inputs[i].placeholder === '请输入账号')     inputs[i].value = "{ACCOUNT}";
                if (inputs[i].placeholder === '请输入员工工号') inputs[i].value = "{WORKER_ID}";
                if (inputs[i].placeholder === '请输入工号密码') inputs[i].value = "{PASSWORD}";
            }}
            inputs = document.querySelectorAll('input');
            for (var i = 0; i < inputs.length; i++) {{
                if (inputs[i].value) {{
                    inputs[i].dispatchEvent(new Event('input', {{bubbles: true}}));
                }}
            }}
        }})()
    """)

    time.sleep(1.5)

    logger.info("[4/4] 点击登录按钮...")
    click_by_text(page, "登 录", "登录")
    page.wait_for_load_state("networkidle", timeout=120_000)

    time.sleep(2)


def select_store(page, store_full_name: str):
    """从自定义 div 下拉 (#ddl_subUsers) 中选择指定门店。"""
    logger.info(f"  → 选择门店: {store_full_name}")

    dropdown = page.locator("#ddl_subUsers")
    dropdown.click()
    time.sleep(0.5)

    page.evaluate("""
        (function() {
            var lis = document.querySelectorAll('#ddl_subUsers .selectBox li');
            lis.forEach(function(li) {
                if (li.classList.contains('on')) li.click();
            });
        })()
    """)
    time.sleep(0.3)

    target_li = page.locator(f"#ddl_subUsers .selectBox li[title='{store_full_name}']")
    if target_li.count() > 0:
        target_li.click()
        logger.info(f"    门店已选中: {store_full_name}")
    else:
        click_result = page.evaluate(f"""
            (function() {{
                var lis = document.querySelectorAll('#ddl_subUsers .selectBox li');
                for (var i = 0; i < lis.length; i++) {{
                    var t = lis[i].textContent.replace(/\\u00a0/g, ' ').trim();
                    if (t === '{store_full_name}') {{
                        lis[i].click();
                        return 'clicked: ' + t;
                    }}
                }}
                var names = [];
                for (var i = 0; i < lis.length; i++) names.push(lis[i].textContent.replace(/\\u00a0/g, ' ').trim());
                return 'not found in: ' + names.join(', ');
            }})()
        """)
        logger.info(f"    模糊匹配: {click_result}")

    time.sleep(0.3)

    close_btn = page.locator("#ddl_subUsers .bottomBar .btnGrey14")
    if close_btn.count() > 0:
        close_btn.click()
        logger.info("    下拉框已关闭")
    time.sleep(0.5)


def click_export(page):
    """点击导出按钮。"""
    result = page.evaluate("""
        (function() {
            var els = document.querySelectorAll('*');
            for (var i = 0; i < els.length; i++) {
                if (els[i].textContent.trim() === '导出' && els[i].children.length === 0) {
                    els[i].click();
                    return 'ok';
                }
            }
            return 'not found';
        })()
    """)
    logger.info(f"  点击导出: {result}")
    if result == "not found":
        raise RuntimeError("未找到「导出」按钮")


def main():
    parser = argparse.ArgumentParser(description="Pospal 麦安研营业统计自动化脚本")
    parser.add_argument("--days", type=int, default=0, help="日期偏移量：0=今天，-1=昨天（默认0）")
    parser.add_argument("--date", type=str, help="指定目标日期，格式 YYYY.MM.DD")
    parser.add_argument("--headless", action="store_true", help="无头模式（不显示浏览器窗口）")
    args = parser.parse_args()

    target = get_target_date(days=args.days, date_str=args.date)
    target_str = target.strftime("%Y.%m.%d")
    date_label = target.strftime("%Y-%m-%d")

    logger.info(f"{'=' * 55}")
    logger.info(f"  Pospal 麦安研营业统计")
    logger.info(f"  目标日期：{target_str}")
    logger.info(f"  输出根目录：{OUTPUT_DIR}")
    logger.info(f"{'=' * 55}")

    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        logger.error("未安装 playwright，请先运行: pip install playwright && playwright install chromium")
        sys.exit(1)

    with sync_playwright() as pw:
        browser = pw.chromium.launch(
            headless=args.headless,
            args=["--start-maximized"],
        )
        context = browser.new_context(
            viewport={"width": 1600, "height": 900},
            accept_downloads=True,
        )
        page = context.new_page()
        page.set_default_timeout(120000)
        page.set_default_navigation_timeout(120000)

        try:
            login(page)

            output_dir = OUTPUT_DIR / date_label
            output_dir.mkdir(parents=True, exist_ok=True)

            logger.info(f"{'─' * 55}")
            logger.info(f"  下载营业概况日度统计")
            logger.info(f"{'─' * 55}")

            for i, store in enumerate(STORES):
                logger.info(f"{'─' * 40}")
                logger.info(f"  门店 {i + 1}/{len(STORES)}: {store['short']} ({store['full']})")
                logger.info(f"{'─' * 40}")

                logger.info("  [导航] 前往营业概况页面...")
                page.goto(BUSINESS_SUMMARY_URL)
                page.wait_for_load_state("networkidle", timeout=120_000)
                logger.info(f"  已到达 → {page.url}")

                select_store(page, store["full"])

                logger.info(f"  → 设置日期: {target_str}...")
                set_date(page, "开始日期", f"{target_str} 00:00")
                set_date(page, "结束日期", f"{target_str} 23:59")

                logger.info("  [查询] 执行查询...")
                click_by_text(page, "查询", "查询")
                page.wait_for_load_state("networkidle", timeout=150_000)
                time.sleep(3)

                logger.info("  [导出] 导出文件...")
                with page.expect_download(timeout=180_000) as dl_info:
                    click_export(page)

                download = dl_info.value
                logger.info(f"  下载文件名: {download.suggested_filename}")

                dest = output_dir / f"营业概况_{store['short']}_{date_label}.xlsx"
                download.save_as(dest)
                logger.info(f"  已保存到: {dest}")

            logger.info(f"{'=' * 55}")
            logger.info(f"  麦安研营业统计下载全部完成！")
            logger.info(f"  输出目录: {output_dir}")
            logger.info(f"{'=' * 55}\n")

            # ── Part 2：格式化数据并写入月度统计表 ──────────────────────────
            logger.info(f"{'─' * 55}")
            logger.info(f"  Part 2：格式化数据并写入月度统计表")
            logger.info(f"{'─' * 55}")

            for i, store in enumerate(STORES):
                logger.info(f"{'─' * 40}")
                logger.info(f"  门店 {i + 1}/{len(STORES)}: {store['template_name']}")
                logger.info(f"{'─' * 40}")

                monthly_file, created = create_or_open_monthly_file(
                    store, target, TEMPLATE_FILE, OUTPUT_DIR
                )
                fill_daily_data(monthly_file, target, store, output_dir)

            logger.info(f"{'=' * 55}")
            logger.info(f"  麦安研营业统计格式化全部完成！")
            logger.info(f"{'=' * 55}\n")

        except Exception as e:
            logger.error(f"任务失败: {e}")
            try:
                screenshot = OUTPUT_DIR / "error_screenshot.png"
                page.screenshot(path=str(screenshot))
                logger.info(f"  错误截图已保存: {screenshot}")
            except Exception:
                pass
            raise
        finally:
            context.close()
            browser.close()


if __name__ == "__main__":
    main()
