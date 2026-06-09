"""
麦安研产品销售周度排行 - 自动化脚本
=====================================
依赖：pip install playwright && playwright install chromium

自动登录 Pospal 后台，导出麦安研各门店的商品销售和商品报损周度统计数据。

用法：
    python mainyan_prod_sales_rank_weekly.py                    # 导出本周数据（周一到周日）
    python mainyan_prod_sales_rank_weekly.py --weeks -1         # 导出上周数据
    python mainyan_prod_sales_rank_weekly.py --date 2026.06.03  # 根据日期推断所在周
    python mainyan_prod_sales_rank_weekly.py --headless          # 无头模式（不显示浏览器）
"""

import argparse
import json
import logging
import logging.handlers
import sys
import time
from datetime import datetime, timedelta
from pathlib import Path

# ─── 日志配置 ───────────────────────────────────────────────────────────────────
LOG_DIR = Path(__file__).resolve().parent / "log"
LOG_DIR.mkdir(exist_ok=True)

_file_handler = logging.handlers.TimedRotatingFileHandler(
    LOG_DIR / "mainyan_prod_sales_rank_weekly.log",
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
# 账号信息同 factory_delivery_tsj_weekly.py
ACCOUNT = "huomimayzb"
WORKER_ID = "M006"
PASSWORD = "tusijia88"

LOGIN_URL = "https://beta69.pospal.cn/"
PRODUCT_SALE_URL = "https://beta69.pospal.cn/ReportV2/ProductSale"
DISCARD_PRODUCT_URL = "https://beta69.pospal.cn/Inventory/DiscardProductCount"

OUTPUT_DIR = Path(__file__).resolve().parent / "麦安研产品销售周度排行"

STORES = [
    {"full": "3 - 麦安研（东站宝泰店）", "short": "宝泰店"},
    {"full": "5 - 麦安研（顺德龙江店）", "short": "龙江店"},
    {"full": "2 - 麦安研（顺德杏坛店）", "short": "杏坛店"},
]

SALE_CATEGORIES = ["热销酥类", "现烤面包", "包装面包", "裱花自制"]


# ─── 工具函数 ──────────────────────────────────────────────────────────────────


def get_week_range(weeks: int = 0, date_str: str = None):
    if date_str:
        base = datetime.strptime(date_str, "%Y.%m.%d")
    else:
        base = datetime.now()

    monday = base - timedelta(days=base.weekday())
    monday = monday + timedelta(weeks=weeks)
    sunday = monday + timedelta(days=6)

    return monday, sunday


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
    """从门店下拉中选择指定门店（支持 <select> 和自定义下拉）。"""
    logger.info(f"  → 选择门店: {store_full_name}")

    result = page.evaluate(f"""
        (function() {{
            var selects = document.querySelectorAll('select');
            for (var i = 0; i < selects.length; i++) {{
                var options = selects[i].querySelectorAll('option');
                for (var j = 0; j < options.length; j++) {{
                    if (options[j].textContent.trim().indexOf('全部门店') >= 0) {{
                        for (var k = 0; k < options.length; k++) {{
                            if (options[k].textContent.trim() === '{store_full_name}') {{
                                selects[i].value = options[k].value;
                                selects[i].dispatchEvent(new Event('change', {{bubbles: true}}));
                                return 'selected via select';
                            }}
                        }}
                        return 'store option not found in select';
                    }}
                }}
            }}
            return 'select_not_found';
        }})()
    """)
    logger.info(f"    门店选择: {result}")

    if result == "select_not_found":
        logger.info("    尝试自定义下拉...")
        click_by_text(page, "全部门店", "全部门店下拉")
        time.sleep(0.5)
        result2 = page.evaluate(f"""
            (function() {{
                var els = document.querySelectorAll('*');
                for (var i = 0; i < els.length; i++) {{
                    if (els[i].textContent.trim() === '{store_full_name}' && els[i].children.length === 0) {{
                        els[i].click();
                        return 'selected';
                    }}
                }}
                return 'not found';
            }})()
        """)
        logger.info(f"    自定义下拉: {result2}")

    time.sleep(1)


def click_advanced_search(page):
    """点击高级搜索/更多搜索，展开搜索选项。"""
    result = page.evaluate("""
        (function() {
            var labels = ['高级搜索', '更多搜索', '更多搜索▼', '高级搜索▼'];
            var els = document.querySelectorAll('*');
            for (var i = 0; i < els.length; i++) {
                var t = els[i].textContent.trim();
                for (var j = 0; j < labels.length; j++) {
                    if (t === labels[j] && els[i].children.length === 0) {
                        els[i].click();
                        return 'clicked: ' + t;
                    }
                }
            }
            var btn = document.getElementById('advancedBtn');
            if (btn) { btn.click(); return 'clicked advancedBtn'; }
            return 'not found';
        })()
    """)
    logger.info(f"  [高级搜索] → {result}")
    time.sleep(0.5)


def select_sale_categories(page, categories: list):
    """打开分类选择弹框，勾选指定分类，关闭弹框。"""
    logger.info(f"  → 选择分类: {categories}")

    select_result = page.evaluate("""
        (function() {
            var btn = document.getElementById('selectCategory');
            if (btn) { btn.click(); return 'clicked selectCategory'; }
            var els = document.querySelectorAll('*');
            for (var i = 0; i < els.length; i++) {
                if (els[i].textContent.trim() === '选择分类' && els[i].children.length === 0) {
                    els[i].click();
                    return 'clicked tag=' + els[i].tagName;
                }
            }
            return 'not found';
        })()
    """)
    logger.info(f"    [选择分类] → {select_result}")
    time.sleep(1.0)

    cats_json = json.dumps(categories, ensure_ascii=False)
    page.evaluate(f"""
        (function() {{
            var target = {cats_json};
            var divs = document.querySelectorAll('.checkBoxDiv');
            divs.forEach(function(d) {{
                var span = d.querySelector('span');
                if (!span) return;
                var name = span.textContent.trim();
                var shouldBeOn = target.indexOf(name) >= 0;
                var isOn = d.classList.contains('on');
                if (shouldBeOn && !isOn) d.click();
                if (!shouldBeOn && isOn) d.click();
            }});
        }})()
    """)

    check_result = page.evaluate(f"""
        (function() {{
            var target = {cats_json};
            var divs = document.querySelectorAll('.checkBoxDiv');
            var r = [];
            divs.forEach(function(d) {{
                var s = d.querySelector('span');
                if (s) {{
                    var name = s.textContent.trim();
                    if (target.indexOf(name) >= 0 || d.classList.contains('on')) {{
                        r.push(name + ':' + d.classList.contains('on'));
                    }}
                }}
            }});
            return r.join(' | ');
        }})()
    """)
    logger.info(f"    勾选验证: {check_result}")

    logger.info("  → 点击「确定」关闭弹框...")
    click_by_text(page, "确定", "确定弹框")
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


def click_popup_export(page):
    """点击弹窗中的导出按钮（报损页面导出有二次确认弹窗）。"""
    result = page.evaluate("""
        (function() {
            var layers = document.querySelectorAll('.layui-layer');
            for (var m = layers.length - 1; m >= 0; m--) {
                var btns = layers[m].querySelectorAll('*');
                for (var i = 0; i < btns.length; i++) {
                    if (btns[i].textContent.trim() === '导出' && btns[i].children.length === 0) {
                        btns[i].click();
                        return 'clicked layui popup';
                    }
                }
            }
            var modals = document.querySelectorAll('[class*="modal"], [class*="dialog"], [class*="popup"]');
            for (var m = modals.length - 1; m >= 0; m--) {
                var btns = modals[m].querySelectorAll('*');
                for (var i = 0; i < btns.length; i++) {
                    if (btns[i].textContent.trim() === '导出' && btns[i].children.length === 0) {
                        btns[i].click();
                        return 'clicked modal popup';
                    }
                }
            }
            var matches = [];
            var els = document.querySelectorAll('*');
            for (var i = 0; i < els.length; i++) {
                if (els[i].textContent.trim() === '导出' && els[i].children.length === 0) {
                    matches.push(els[i]);
                }
            }
            if (matches.length > 1) {
                matches[matches.length - 1].click();
                return 'clicked last of ' + matches.length;
            }
            return 'not found';
        })()
    """)
    logger.info(f"  [弹窗导出] → {result}")
    if result == "not found":
        raise RuntimeError("未找到弹窗中的「导出」按钮")


def main():
    parser = argparse.ArgumentParser(description="Pospal 麦安研产品销售周度排行自动化脚本")
    parser.add_argument("--weeks", type=int, default=0, help="周偏移量：0=本周，-1=上周（默认0）")
    parser.add_argument("--date", type=str, help="指定日期推断所在周，格式 YYYY.MM.DD")
    parser.add_argument("--headless", action="store_true", help="无头模式（不显示浏览器窗口）")
    args = parser.parse_args()

    if args.date:
        monday, sunday = get_week_range(date_str=args.date)
    else:
        monday, sunday = get_week_range(weeks=args.weeks)

    monday_str = monday.strftime("%Y.%m.%d")
    sunday_str = sunday.strftime("%Y.%m.%d")
    date_range_str = f"{monday.strftime('%Y-%m-%d')}~{sunday.strftime('%Y-%m-%d')}"

    logger.info(f"{'=' * 55}")
    logger.info(f"  Pospal 麦安研产品销售周度排行")
    logger.info(f"  目标周：{monday_str} ~ {sunday_str}")
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

            output_dir = OUTPUT_DIR / date_range_str / "原始下载"
            output_dir.mkdir(parents=True, exist_ok=True)

            # ── Part 1：商品销售周度统计（3个门店）──────────────────────────
            logger.info(f"{'─' * 55}")
            logger.info(f"  Part 1：下载商品销售周度统计")
            logger.info(f"{'─' * 55}")

            for i, store in enumerate(STORES):
                logger.info(f"\n{'─' * 40}")
                logger.info(f"  门店 {i + 1}/{len(STORES)}: {store['short']} ({store['full']})")
                logger.info(f"{'─' * 40}")

                logger.info("  [导航] 前往商品销售统计页面...")
                page.goto(PRODUCT_SALE_URL)
                page.wait_for_load_state("networkidle", timeout=120_000)
                logger.info(f"  已到达 → {page.url}")

                select_store(page, store["full"])

                logger.info(f"  → 设置日期: {monday_str} ~ {sunday_str}...")
                set_date(page, "开始日期", f"{monday_str} 00:00")
                set_date(page, "结束日期", f"{sunday_str} 23:59")

                click_advanced_search(page)
                select_sale_categories(page, SALE_CATEGORIES)

                logger.info("  [查询] 执行查询...")
                click_by_text(page, "查询", "查询")
                page.wait_for_load_state("networkidle", timeout=150_000)
                time.sleep(3)

                logger.info("  [导出] 导出文件...")
                with page.expect_download(timeout=180_000) as dl_info:
                    click_export(page)

                download = dl_info.value
                logger.info(f"  下载文件名: {download.suggested_filename}")

                dest = output_dir / f"商品销售周度统计_{store['short']}_{date_range_str}.xlsx"
                download.save_as(dest)
                logger.info(f"  已保存到: {dest}")

            # ── Part 2：商品报损周度统计（3个门店）──────────────────────────
            logger.info(f"\n{'─' * 55}")
            logger.info(f"  Part 2：下载商品报损周度统计")
            logger.info(f"{'─' * 55}")

            for i, store in enumerate(STORES):
                logger.info(f"\n{'─' * 40}")
                logger.info(f"  门店 {i + 1}/{len(STORES)}: {store['short']} ({store['full']})")
                logger.info(f"{'─' * 40}")

                logger.info("  [导航] 前往商品报损统计页面...")
                page.goto(DISCARD_PRODUCT_URL)
                page.wait_for_load_state("networkidle", timeout=120_000)
                logger.info(f"  已到达 → {page.url}")

                select_store(page, store["full"])

                logger.info(f"  → 设置日期: {monday_str} ~ {sunday_str}...")
                set_date(page, "开始日期", f"{monday_str} 00:00")
                set_date(page, "结束日期", f"{sunday_str} 23:59")

                logger.info("  [查询] 执行查询...")
                click_by_text(page, "查询", "查询")
                page.wait_for_load_state("networkidle", timeout=150_000)
                time.sleep(3)

                logger.info("  [导出] 点击导出...")
                click_export(page)
                time.sleep(1.5)

                logger.info("  [弹窗导出] 点击弹窗中的导出...")
                with page.expect_download(timeout=180_000) as dl_info:
                    click_popup_export(page)

                download = dl_info.value
                logger.info(f"  下载文件名: {download.suggested_filename}")

                dest = output_dir / f"商品报损周度统计_{store['short']}_{date_range_str}.xlsx"
                download.save_as(dest)
                logger.info(f"  已保存到: {dest}")

            logger.info(f"\n{'=' * 55}")
            logger.info(f"  麦安研产品销售周度排行下载全部完成！")
            logger.info(f"  输出目录: {output_dir}")
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
