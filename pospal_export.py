"""
Pospal 订货看板 - 自动导出脚本
=====================================
依赖：playwright (Python)
安装：pip install playwright && playwright install chromium

自动登录后依次导出：
  1. 订货商品汇总看板 (ProductRequestSummaryBoard)
  2. 订货商品明细看板 (ProductRequestItemBoard)

用法：
    python pospal_export.py                  # 导出后天的数据
    python pospal_export.py --date 2026.05.30  # 指定日期
    python pospal_export.py --days 2            # N天后（默认2=后天）
    python pospal_export.py --headless          # 无头模式（不显示浏览器）
"""

import argparse
import shutil
import sys
import time
from datetime import datetime, timedelta
from pathlib import Path
from time import sleep

# ─── 配置区（按需修改）─────────────────────────────────────────────────────────
ACCOUNT = "huomimayzb"
WORKER_ID = "M006"
PASSWORD = "tusijia88"

LOGIN_URL = "https://beta69.pospal.cn/"
SUMMARY_BOARD_URL = "https://css69.pospal.cn/ChainStoreSupplySeller/ProductRequestSummaryBoard"
ITEM_BOARD_URL = "https://css69.pospal.cn/ChainStoreSupplySeller/ProductRequestItemBoard"

OUTPUT_DIR = Path.home() / "Desktop" / "订货商品汇总看板"

# 要勾选的商品分类（与 Pospal DOM 中的 span 文本完全一致）
TARGET_CATEGORIES = [
    "配送费",
    "包材耗材",
    "工衣模具",
    "慕斯+饼干+饮品+其他",
    "原料铺料",
    "冷冻面团",
    "蛋糕及面包成品及饼干类",
]


def js(page, script: str, desc: str = ""):
    """执行 JS 并打印结果（便于调试）"""
    result = page.evaluate(script)
    tag = f"[{desc}] " if desc else ""
    print(f"  {tag}→ {result}")
    return result


def set_date(page, placeholder: str, value: str):
    """设置日期输入框的值并触发必要事件"""
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


def verify_dates(page, expected_date: str) -> bool:
    """验证两个日期输入框的值是否都包含期望日期"""
    result = page.evaluate("""
        (function() {
            var r = [];
            document.querySelectorAll('input.timeInput.hasDatepicker').forEach(function(inp) {
                r.push(inp.placeholder + '=' + inp.value);
            });
            return r.join(' | ');
        })()
    """)
    print(f"  [日期验证] {result}")
    return expected_date in result and result.count(expected_date) == 2


def click_by_text(page, text: str, desc: str = ""):
    """通过完整文本内容点击叶子节点元素"""
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
    print(f"  {tag}→ {result}")
    return "clicked" in result


def login(page):
    """登录流程"""
    print("\n[1/8] 打开登录页面...")
    page.goto(LOGIN_URL)
    page.wait_for_load_state("networkidle")

    print("[2/8] 切换到工号登录模式...")
    # 工号登录按钮实际是 <span>，不是 <a>，用通用文本匹配
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

    # 确认切换成功
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
    print(f"  表单已切换 → {placeholder}")

    print("[3/8] 填入账号/工号/密码...")
    page.evaluate(f"""
        (function() {{
            var inputs = document.querySelectorAll('input');
            for (var i = 0; i < inputs.length; i++) {{
                if (inputs[i].placeholder === '请输入账号')     inputs[i].value = "{ACCOUNT}";
                if (inputs[i].placeholder === '请输入员工工号') inputs[i].value = "{WORKER_ID}";
                if (inputs[i].placeholder === '请输入工号密码') inputs[i].value = "{PASSWORD}";
            }}
            // 触发 input 事件让框架感知到变化
            inputs = document.querySelectorAll('input');
            for (var i = 0; i < inputs.length; i++) {{
                if (inputs[i].value) {{
                    inputs[i].dispatchEvent(new Event('input', {{bubbles: true}}));
                }}
            }}
        }})()
    """)

    time.sleep(1.5)

    print("[4/8] 点击登录按钮...")
    click_by_text(page, "登 录", "登录")
    page.wait_for_load_state("networkidle", timeout=30_000)

    time.sleep(2)


def navigate_to_board(page, board_url: str, board_name: str):
    """导航到指定看板"""
    print(f"\n  [导航] 前往{board_name}...")
    page.goto(board_url)
    page.wait_for_load_state("networkidle", timeout=30_000)
    print(f"  已到达 → {page.url}")


def setup_filters(page, target_date: str):
    """
    设置筛选条件：
      1. 切换到「期望到货时间」标签
      2. 设置日期
      3. 选择商品分类
    """
    print("\n  [筛选] 设置筛选条件...")

    # ── 1. 切换日期类型标签 ──
    print("  → 切换到「期望到货时间」标签...")
    result = page.evaluate("""
        (function() {
            var lis = document.querySelectorAll('li');
            for (var i = 0; i < lis.length; i++) {
                if (lis[i].textContent.trim() === '期望到货时间') {
                    lis[i].click();
                    return 'clicked li[' + i + ']';
                }
            }
            return 'not found';
        })()
    """)
    print(f"    {result}")
    if result == "not found":
        raise RuntimeError("未找到「期望到货时间」标签")

    # 验证标签切换
    tab_status = page.evaluate("""
        (function() {
            var r = [];
            document.querySelectorAll('li').forEach(function(li) {
                var t = li.textContent.trim();
                if (t === '订货时间' || t === '期望到货时间') {
                    r.push(t + ':' + li.className);
                }
            });
            return r.join(' | ');
        })()
    """)
    print(f"    标签状态: {tab_status}")

    # ── 2. 展开高级搜索，选择分类 ──
    print("  → 展开高级搜索面板...")
    page.evaluate("document.getElementById('advancedBtn').click()")
    time.sleep(0.5)

    print("  → 打开分类选择弹框...")
    page.evaluate("document.getElementById('selectCategory').click()")
    time.sleep(1.0)

    print(f"  → 勾选 {len(TARGET_CATEGORIES)} 个分类...")
    categories_js = str(TARGET_CATEGORIES).replace("'", '"')  # 转成 JS 数组字面量
    page.evaluate(f"""
        (function() {{
            var targets = {categories_js};
            var divs = document.querySelectorAll('.checkBoxDiv');
            divs.forEach(function(d) {{
                var span = d.querySelector('span');
                if (span && targets.indexOf(span.textContent.trim()) >= 0) {{
                    d.click();
                }}
            }});
        }})()
    """)

    # 验证勾选状态
    check_result = page.evaluate(f"""
        (function() {{
            var targets = {categories_js};
            var divs = document.querySelectorAll('.checkBoxDiv');
            var r = [];
            divs.forEach(function(d) {{
                var s = d.querySelector('span');
                if (s && targets.indexOf(s.textContent.trim()) >= 0) {{
                    r.push(s.textContent.trim() + ':' + d.classList.contains('on'));
                }}
            }});
            return r.join(' | ');
        }})()
    """)
    print(f"    勾选验证: {check_result}")
    if "false" in check_result or check_result.count(":true") < len(TARGET_CATEGORIES):
        print("  ⚠️  部分分类未勾选，尝试重试...")
        # 重试一次
        page.evaluate(f"""
            (function() {{
                var targets = {categories_js};
                var divs = document.querySelectorAll('.checkBoxDiv');
                divs.forEach(function(d) {{
                    var span = d.querySelector('span');
                    if (span && targets.indexOf(span.textContent.trim()) >= 0 && !d.classList.contains('on')) {{
                        d.click();
                    }}
                }});
            }})()
        """)

    # ── 3. 设置日期（切换标签后立刻设置）──
    print(f"  → 设置日期: {target_date}...")
    set_date(page, "开始日期", f"{target_date} 00:00")
    set_date(page, "结束日期", f"{target_date} 23:59")

    print("  → 点击「确定」关闭弹框...")
    click_by_text(page, "确定", "确定弹框")
    time.sleep(0.5)


def query_and_verify(page, target_date: str, max_retries: int = 3) -> int:
    """
    执行查询，并处理日期被重置的问题。
    返回查询到的数据行数。
    """
    print("\n  [查询] 执行查询...")

    for attempt in range(1, max_retries + 1):
        print(f"  查询第 {attempt} 次...")
        page.evaluate("document.getElementById('btnLoadRequestList').click()")
        page.wait_for_load_state("networkidle", timeout=90_000)

        # 验证日期和行数
        result = page.evaluate("""
            (function() {
                var rows = document.querySelectorAll('table tbody tr');
                var r = 'rows:' + rows.length + ' | ';
                document.querySelectorAll('input.timeInput.hasDatepicker').forEach(function(inp) {
                    r += inp.placeholder + '=' + inp.value + ' | ';
                });
                return r;
            })()
        """)
        print(f"    结果: {result}")

        # sleep(30)

        if target_date in result:
            # 日期正确，提取行数
            import re
            m = re.search(r"rows:(\d+)", result)
            row_count = int(m.group(1)) if m else 0
            print(f"  ✅ 查询成功，共 {row_count} 行数据")
            return row_count
        else:
            print(f"  ⚠️  日期被重置！重新设置日期...")
            set_date(page, "开始日期", f"{target_date} 00:00")
            set_date(page, "结束日期", f"{target_date} 23:59")
            verify_dates(page, target_date)

    raise RuntimeError(f"查询 {max_retries} 次后日期仍然不正确，请手动检查")

def query_and_verify_item_board(page, target_date: str, max_retries: int = 3) -> int:
    """
    执行查询，并处理日期被重置的问题。
    返回查询到的数据行数。
    """
    print("\n  [查询] 执行查询...")

    for attempt in range(1, max_retries + 1):
        print(f"  查询第 {attempt} 次...")
        page.evaluate("document.getElementById('btnList').click()")
        page.wait_for_load_state("networkidle", timeout=90_000)

        # 验证日期和行数
        result = page.evaluate("""
            (function() {
                var rows = document.querySelectorAll('table tbody tr');
                var r = 'rows:' + rows.length + ' | ';
                document.querySelectorAll('input.timeInput.hasDatepicker').forEach(function(inp) {
                    r += inp.placeholder + '=' + inp.value + ' | ';
                });
                return r;
            })()
        """)
        print(f"    结果: {result}")

        # sleep(30)

        if target_date in result:
            # 日期正确，提取行数
            import re
            m = re.search(r"rows:(\d+)", result)
            row_count = int(m.group(1)) if m else 0
            print(f"  ✅ 查询成功，共 {row_count} 行数据")
            return row_count
        else:
            print(f"  ⚠️  日期被重置！重新设置日期...")
            set_date(page, "开始日期", f"{target_date} 00:00")
            set_date(page, "结束日期", f"{target_date} 23:59")
            verify_dates(page, target_date)

    raise RuntimeError(f"查询 {max_retries} 次后日期仍然不正确，请手动检查")


def export_and_save(page, target_date: str, file_prefix: str = "订货商品汇总看板") -> Path:
    """
    点击导出，等待下载，并将文件保存到输出目录。
    返回最终保存的文件路径。
    """
    print("\n  [导出] 导出文件...")
    time.sleep(3)  # 确保导出按钮就绪

    # 格式化日期并创建目标文件夹
    date_str = target_date.replace(".", "-")  # 2026.05.25 → 2026-05-25
    daily_output_dir = OUTPUT_DIR / date_str / "原始下载"
    daily_output_dir.mkdir(parents=True, exist_ok=True)
    print(f"  → 输出至: {daily_output_dir}")

    # 使用 Playwright 的 expect_download 上下文来可靠捕获下载
    with page.expect_download(timeout=60_000) as dl_info:
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
        print(f"  点击导出: {result}")
        if result == "not found":
            raise RuntimeError("未找到「导出」按钮")

    download = dl_info.value
    suggested = download.suggested_filename
    print(f"  下载文件名: {suggested}")

    # 目标路径
    dest = daily_output_dir / f"{file_prefix}_{date_str}.xlsx"

    # 如果目标文件被 Excel 占用，用 _v2 后缀
    # if dest.exists():
    #     try:
    #         dest.unlink()
    #     except PermissionError:
    #         dest = daily_output_dir / f"订货商品汇总看板_{date_str}_v2.xlsx"
    #         print(f"  ⚠️  原文件被占用，改用: {dest.name}")

    download.save_as(dest)
    print(f"  ✅ 已保存到: {dest}")
    return dest


def main():
    parser = argparse.ArgumentParser(description="Pospal 订货商品汇总看板自动导出")
    parser.add_argument("--date", type=str, help="指定日期，格式 YYYY.MM.DD，如 2026.05.30")
    parser.add_argument("--days", type=int, default=2, help="今天之后第N天（默认2=后天）")
    parser.add_argument("--headless", action="store_true", help="无头模式（不显示浏览器窗口）")
    args = parser.parse_args()

    # 确定目标日期
    if args.date:
        target_date = args.date
    else:
        target_dt = datetime.now() + timedelta(days=args.days)
        target_date = target_dt.strftime("%Y.%m.%d")

    print(f"{'=' * 55}")
    print(f"  Pospal 订货看板导出")
    print(f"  目标日期：{target_date}")
    print(f"  输出根目录：{OUTPUT_DIR}")
    print(f"{'=' * 55}")

    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        print("\n❌ 未安装 playwright，请先运行：")
        print("   pip install playwright")
        print("   playwright install chromium")
        sys.exit(1)

    with sync_playwright() as pw:
        browser = pw.chromium.launch(
            headless=args.headless,
            args=["--start-maximized"],
        )
        context = browser.new_context(
            viewport={"width": 1600, "height": 900},
            # 设置下载行为：允许下载
            accept_downloads=True,
        )
        page = context.new_page()

        try:
            login(page)

            # ── 任务 1：订货商品汇总看板 ──────────────────────────────────
            print(f"\n{'─' * 55}")
            print(f"  任务 1/2：订货商品汇总看板")
            print(f"{'─' * 55}")

            navigate_to_board(page, SUMMARY_BOARD_URL, "订货商品汇总看板")
            setup_filters(page, target_date)
            summary_row_count = query_and_verify(page, target_date)
            summary_path = export_and_save(page, target_date, "订货商品汇总看板")

            # ── 任务 2：订货商品明细看板 ──────────────────────────────────
            print(f"\n{'─' * 55}")
            print(f"  任务 2/2：订货商品明细看板")
            print(f"{'─' * 55}")

            navigate_to_board(page, ITEM_BOARD_URL, "订货商品明细看板")
            setup_filters(page, target_date)
            item_row_count = query_and_verify_item_board(page, target_date)
            item_path = export_and_save(page, target_date, "订货商品明细看板")

            print(f"\n{'=' * 55}")
            print(f"  ✅ 全部完成！")
            print(f"  订货商品汇总看板：{summary_row_count} 行 → {summary_path}")
            print(f"  订货商品明细看板：{item_row_count} 行 → {item_path}")
            print(f"{'=' * 55}\n")

        except Exception as e:
            print(f"\n❌ 任务失败: {e}")
            # 截图便于调试
            try:
                screenshot = OUTPUT_DIR / "error_screenshot.png"
                page.screenshot(path=str(screenshot))
                print(f"  错误截图已保存: {screenshot}")
            except Exception:
                pass
            raise
        finally:
            context.close()
            # browser.close()


if __name__ == "__main__":
    main()
