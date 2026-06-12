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
import csv
import json
import logging
import logging.handlers
import re
import subprocess
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
WORKER_ID = "M008"
PASSWORD = "Maianyan88"

LOGIN_URL = "https://beta69.pospal.cn/"
BUSINESS_SUMMARY_URL = "https://beta69.pospal.cn/Report/BusinessSummaryV2"
UNIONPAY_BILL_URL = "https://cloudapp-pay69.pospal.cn/#/additional/fund-summary?oem=0"
CUSTOMER_SUMMARY_URL = "https://beta69.pospal.cn/CustomerReport/CustomerConsumerSummary"
MEITUAN_DOWNLOAD_URL = "https://waimaieapp.meituan.com/finance/static/gray_html_pc/billReconciliation.html#/daily-bill"

MEITUAN_STORE_CONFIG = [
    {
        "store_short": "宝泰店",
        "port": 9223,
        "user_data_dir": r"C:\ChromeDebug_BT",
    },
    {
        "store_short": "龙江店",
        "port": 9224,
        "user_data_dir": r"C:\ChromeDebug_LJ",
    },
    {
        "store_short": "杏坛店",
        "port": 9225,
        "user_data_dir": r"C:\ChromeDebug_XT",
    },
]

OUTPUT_DIR = Path(__file__).resolve().parent / "麦安研营业统计"
TEMPLATE_FILE = Path(__file__).resolve().parent / "麦安研营业统计_格式化模板.xlsx"

STORES = [
    {"full": "3 - 麦安研（东站宝泰店）", "short": "宝泰店", "template_name": "东站宝泰店"},
    {"full": "5 - 麦安研（顺德龙江店）", "short": "龙江店", "template_name": "顺德龙江店"},
    {"full": "2 - 麦安研（顺德杏坛店）", "short": "杏坛店", "template_name": "顺德杏坛店"},
]

UNIONPAY_STORE_CONFIG = [
    {
        "store_short": "宝泰店",
        "parent_node": "总部",
        "select_items": ["麦安研", "麦安研（东站宝泰店）"],
    },
    {
        "store_short": "龙江店",
        "parent_node": None,
        "select_items": ["麦安研（顺德龙江店）"],
    },
    {
        "store_short": "杏坛店",
        "parent_node": "总部",
        "select_items": ["麦安研（顺德杏坛店）"],
    },
]

CUSTOMER_SUMMARY_STORE_CONFIG = [
    {
        "store_short": "宝泰店",
        "select_items": ["1 - 麦安研", "3 - 麦安研（东站宝泰店）"],
    },
    {
        "store_short": "龙江店",
        "select_items": ["5 - 麦安研（顺德龙江店）"],
    },
    {
        "store_short": "杏坛店",
        "select_items": ["2 - 麦安研（顺德杏坛店）"],
    },
]

TEMPLATE_STORE_NAME = "东方宝泰店"

UNIONPAY_CSV_FIELDS = ["交易金额", "交易退款金额", "有效交易金额", "交易手续费", "优惠金额", "优惠退款金额"]

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


def _parse_numeric(val):
    if val is None:
        return None
    try:
        return float(str(val).strip())
    except (ValueError, TypeError):
        return val


def _parse_recharge_amount(text):
    if not text or not isinstance(text, str):
        return None
    m = re.search(r"充值\s*([\d.]+)", text)
    return float(m.group(1)) if m else None


def read_business_summary(data_file: Path):
    from openpyxl import load_workbook as _lwb
    wb = _lwb(data_file, data_only=True)
    ws1 = wb.worksheets[0]
    ws2 = wb.worksheets[1] if len(wb.worksheets) > 1 else None
    result = {
        "c16": _parse_numeric(ws1.cell(row=16, column=3).value),
        "ws2_c4": _parse_numeric(ws2.cell(row=4, column=3).value) if ws2 else None,
        "e4": _parse_numeric(ws1.cell(row=4, column=5).value),
        "b4_recharge": _parse_recharge_amount(ws1.cell(row=4, column=2).value),
    }
    wb.close()
    return result


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
        logger.warning(f"  营业概况文件不存在，跳过填写: {summary_file.name}")
        summary = None
    else:
        summary = read_business_summary(summary_file)
        logger.info(f"  读取 sheet2 C4 (现金充值) = {summary['ws2_c4']}")
        logger.info(f"  读取 C16 (现金合计) = {summary['c16']}")
        logger.info(f"  读取 E4 (充值金额) = {summary['e4']}")
        logger.info(f"  读取 B4 解析充值金额 = {summary['b4_recharge']}")

    unionpay_file = daily_download_dir / f"银豹付交易账单_{store['short']}_{date_label}.csv"
    if not unionpay_file.exists():
        logger.warning(f"  银豹付交易账单不存在，跳过填写: {unionpay_file.name}")
        unionpay = None
    else:
        unionpay = read_unionpay_bill_csv(unionpay_file)
        logger.info(f"  读取银豹付 交易金额 = {unionpay.get('交易金额')}")
        logger.info(f"  读取银豹付 交易退款金额 = {unionpay.get('交易退款金额')}")
        logger.info(f"  读取银豹付 交易手续费 = {unionpay.get('交易手续费')}")

    customer_file = daily_download_dir / f"会员消费汇总表_{store['short']}_{date_label}.xlsx"
    if not customer_file.exists():
        logger.warning(f"  会员消费汇总表不存在，跳过填写: {customer_file.name}")
        customer = None
    else:
        customer = read_customer_summary(customer_file)

    meituan_file = daily_download_dir / f"账单明细_{store['short']}_{date_label}.csv"
    if not meituan_file.exists():
        logger.warning(f"  美团账单明细不存在，跳过填写: {meituan_file.name}")
        meituan = None
    else:
        meituan = read_unionpay_bill_csv(meituan_file)
        logger.info(f"  读取美团 商品总价 = {meituan.get('商品总价')}")
        logger.info(f"  读取美团 打包费 = {meituan.get('打包费')}")
        logger.info(f"  读取美团 其他类 = {meituan.get('其他类')}")

    wb = _lwb(monthly_file)
    ws1 = wb.worksheets[0]

    row_s1 = day + 4
    if summary:
        if summary["ws2_c4"] is not None:
            ws1.cell(row=row_s1, column=5).value = summary["ws2_c4"]
        if summary["c16"] is not None:
            ws1.cell(row=row_s1, column=6).value = summary["c16"]
        if summary["e4"] is not None:
            ws1.cell(row=row_s1, column=21).value = summary["e4"]
        if summary["b4_recharge"] is not None:
            ws1.cell(row=row_s1, column=22).value = summary["b4_recharge"]

    if unionpay:
        if unionpay.get("交易金额") is not None:
            ws1.cell(row=row_s1, column=9).value = unionpay["交易金额"]
        if unionpay.get("交易退款金额") is not None:
            ws1.cell(row=row_s1, column=10).value = unionpay["交易退款金额"]
        if unionpay.get("交易手续费") is not None:
            ws1.cell(row=row_s1, column=12).value = unionpay["交易手续费"]

    if customer:
        if customer["principal"] is not None:
            ws1.cell(row=row_s1, column=25).value = customer["principal"]
        if customer["gift"] is not None:
            ws1.cell(row=row_s1, column=26).value = customer["gift"]

    if meituan:
        if meituan.get("商品总价") is not None:
            ws1.cell(row=row_s1, column=30).value = abs(meituan["商品总价"])
        if meituan.get("打包费") is not None:
            ws1.cell(row=row_s1, column=31).value = abs(meituan["打包费"])
        if meituan.get("商家对顾客的活动补贴") is not None:
            ws1.cell(row=row_s1, column=32).value = abs(meituan["商家对顾客的活动补贴"])
        if meituan.get("佣金") is not None:
            ws1.cell(row=row_s1, column=33).value = abs(meituan["佣金"])
        if meituan.get("配送服务费") is not None:
            ws1.cell(row=row_s1, column=34).value = abs(meituan["配送服务费"])
        if meituan.get("其他类") is not None:
            ws1.cell(row=row_s1, column=35).value = abs(meituan["其他类"])

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


def scrape_unionpay_bill(page):
    """从银豹付交易账单页面抓取汇总数据。"""
    data = page.evaluate("""
        (function() {
            var items = document.querySelectorAll('.data-dz-item');
            var result = {};
            for (var i = 0; i < items.length; i++) {
                var titleEl = items[i].querySelector('.title');
                if (!titleEl) continue;
                var title = '';
                var childNodes = titleEl.childNodes;
                for (var j = 0; j < childNodes.length; j++) {
                    if (childNodes[j].nodeType === 3) title += childNodes[j].textContent;
                }
                title = title.replace(/\\s+/g, ' ').trim();
                var numEl = items[i].querySelector('.num');
                if (!numEl) continue;
                var numText = numEl.textContent.replace(/[￥\\s]/g, '').trim();
                result[title] = parseFloat(numText) || 0;
            }
            return result;
        })()
    """)
    if not data:
        logger.warning("  未抓取到银豹付数据，可能页面未加载完成")
    else:
        logger.info(f"  抓取到银豹付数据: {json.dumps(data, ensure_ascii=False)}")
    return data


def save_unionpay_bill_csv(data, store_short, date_label, output_dir):
    """将银豹付交易账单数据保存为 CSV。"""
    csv_file = output_dir / f"银豹付交易账单_{store_short}_{date_label}.csv"
    with open(csv_file, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        writer.writerow(["项目", "金额"])
        for key in UNIONPAY_CSV_FIELDS:
            writer.writerow([key, data.get(key, 0)])
    logger.info(f"  已保存CSV: {csv_file.name}")
    return csv_file


def read_unionpay_bill_csv(csv_file):
    """读取银豹付交易账单 CSV。"""
    result = {}
    with open(csv_file, "r", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        for row in reader:
            result[row["项目"]] = _parse_numeric(row["金额"])
    return result


def read_customer_summary(data_file: Path):
    """读取会员消费汇总表，返回合计行的 F列（本金消费）和 G列（赠送消费）。"""
    from openpyxl import load_workbook as _lwb
    wb = _lwb(data_file, data_only=True)
    ws = wb.worksheets[0]

    total_row = None
    for row in range(ws.max_row, 0, -1):
        cell_val = ws.cell(row=row, column=1).value
        if cell_val and "合计" in str(cell_val):
            total_row = row
            break

    if total_row is None:
        total_row = ws.max_row
        while total_row > 1 and ws.cell(row=total_row, column=1).value is None:
            total_row -= 1

    result = {
        "principal": _parse_numeric(ws.cell(row=total_row, column=6).value),
        "gift": _parse_numeric(ws.cell(row=total_row, column=7).value),
    }
    wb.close()
    logger.info(f"  合计行={total_row}, 本金消费(F)={result['principal']}, 赠送消费(G)={result['gift']}")
    return result


def select_unionpay_store(page, config):
    """在银豹付交易账单页面选择门店（Element UI Cascader 多选组件）。"""
    logger.info(f"  → 选择银豹付门店: {config['select_items']}")

    cascader = page.locator(".el-cascader")
    existing_tags = cascader.locator(".el-tag__close")
    tag_count = existing_tags.count()
    if tag_count > 0:
        logger.info(f"    清除已有选项 ({tag_count} 个)...")
        for _ in range(tag_count):
            close_btn = cascader.locator(".el-tag__close").first
            if close_btn.count() > 0:
                close_btn.click()
                time.sleep(0.3)

    search_input = cascader.locator(".el-cascader__search-input")
    if search_input.count() > 0:
        search_input.click()
    else:
        cascader.locator(".el-input__inner").click(force=True)
    time.sleep(0.5)

    panel = page.locator(".el-cascader__dropdown:visible, .el-popper:visible .el-cascader-panel")
    if panel.count() == 0:
        if search_input.count() > 0:
            search_input.click()
        else:
            cascader.locator(".el-input__inner").click(force=True)
        time.sleep(0.5)

    cleared = page.evaluate("""
        (function() {
            var count = 0;
            var checked = document.querySelectorAll('.el-cascader-node .el-checkbox__input.is-checked');
            for (var i = 0; i < checked.length; i++) {
                checked[i].click();
                count++;
            }
            return count;
        })()
    """)
    if cleared > 0:
        logger.info(f"    清除面板中已选复选框: {cleared} 个")
        time.sleep(0.5)

    if config.get("parent_node"):
        parent = config["parent_node"]
        result = page.evaluate(f"""
            (function() {{
                var menus = document.querySelectorAll('.el-cascader-menu');
                if (menus.length === 0) return 'no cascader menu found';
                var firstMenu = menus[menus.length > 1 ? menus.length - 1 : 0];
                for (var m = 0; m < menus.length; m++) {{
                    var nodes = menus[m].querySelectorAll('.el-cascader-node');
                    for (var i = 0; i < nodes.length; i++) {{
                        var label = nodes[i].querySelector('.el-cascader-node__label');
                        var text = label ? label.textContent.trim() : nodes[i].textContent.trim();
                        if (text === '{parent}') {{
                            nodes[i].click();
                            return 'clicked parent: ' + text;
                        }}
                    }}
                }}
                var allLabels = [];
                var nodes = document.querySelectorAll('.el-cascader-node');
                for (var i = 0; i < nodes.length; i++) {{
                    var label = nodes[i].querySelector('.el-cascader-node__label');
                    allLabels.push(label ? label.textContent.trim() : nodes[i].textContent.trim());
                }}
                return 'parent not found: {parent}. Available: ' + allLabels.join(', ');
            }})()
        """)
        logger.info(f"    展开父节点: {result}")
        time.sleep(0.5)

    for item_name in config["select_items"]:
        result = page.evaluate(f"""
            (function() {{
                var nodes = document.querySelectorAll('.el-cascader-node');
                for (var i = 0; i < nodes.length; i++) {{
                    var label = nodes[i].querySelector('.el-cascader-node__label');
                    var text = label ? label.textContent.trim() : nodes[i].textContent.trim();
                    if (text === '{item_name}') {{
                        var cb = nodes[i].querySelector('.el-checkbox__input');
                        if (cb) {{
                            cb.click();
                            return 'checked: ' + text;
                        }}
                        nodes[i].click();
                        return 'clicked: ' + text;
                    }}
                }}
                var allLabels = [];
                for (var i = 0; i < nodes.length; i++) {{
                    var label = nodes[i].querySelector('.el-cascader-node__label');
                    allLabels.push(label ? label.textContent.trim() : nodes[i].textContent.trim());
                }}
                return 'not found: {item_name}. Available: ' + allLabels.join(', ');
            }})()
        """)
        logger.info(f"    选择: {result}")
        time.sleep(0.5)

    page.locator("body").click(position={"x": 0, "y": 0})
    time.sleep(0.3)


def select_store_type_consumption(page):
    """在会员消费汇总表页面选择门店类型为'消费门店'。"""
    logger.info("  → 选择门店类型: 消费门店")

    selector = page.locator("[p-single-selector='userTypeOpts']")
    if selector.count() > 0:
        selector.click()
        time.sleep(0.5)

    result = page.evaluate("""
        (function() {
            var box = document.querySelector('[p-single-selector="userTypeOpts"] .selectBox');
            if (box) box.style.display = 'block';
            var lis = document.querySelectorAll('[p-single-selector="userTypeOpts"] .selectBox li');
            for (var i = 0; i < lis.length; i++) {
                if (lis[i].textContent.trim() === '消费门店') {
                    lis[i].click();
                    return 'clicked: 消费门店 (optionvalue=' + lis[i].getAttribute('optionvalue') + ')';
                }
            }
            var names = [];
            for (var i = 0; i < lis.length; i++) names.push(lis[i].textContent.trim());
            return 'not found in: ' + names.join(', ');
        })()
    """)
    logger.info(f"  门店类型: {result}")
    time.sleep(0.5)


def select_stores_multi(page, store_names):
    """在 #queryStoreDiv 中选择多个门店。"""
    logger.info(f"  → 选择门店范围: {store_names}")

    page.evaluate("""
        (function() {
            var lis = document.querySelectorAll('#queryStoreDiv li');
            for (var i = 0; i < lis.length; i++) {
                if (lis[i].classList.contains('on') || lis[i].classList.contains('selected')) {
                    lis[i].click();
                }
            }
        })()
    """)
    time.sleep(0.3)

    for name in store_names:
        result = page.evaluate(f"""
            (function() {{
                var lis = document.querySelectorAll('#queryStoreDiv li');
                for (var i = 0; i < lis.length; i++) {{
                    var text = lis[i].textContent.trim();
                    if (text === '{name}') {{
                        lis[i].click();
                        return 'clicked: ' + text + ' (data=' + lis[i].getAttribute('data') + ')';
                    }}
                }}
                var names = [];
                for (var i = 0; i < lis.length; i++) names.push(lis[i].textContent.trim());
                return 'not found: ' + name + '. Available: ' + names.join(', ');
            }})()
        """)
        logger.info(f"    {result}")
        time.sleep(0.3)


def set_vue_date(page, target_str):
    """在 Vue/Element UI 页面设置日期。"""
    date_dash = target_str.replace(".", "-")

    range_inputs = page.locator("input.el-range-input")
    if range_inputs.count() >= 2:
        for idx in range(2):
            range_inputs.nth(idx).click()
            time.sleep(0.2)
            range_inputs.nth(idx).press("Control+a")
            range_inputs.nth(idx).type(date_dash, delay=50)
        range_inputs.nth(1).press("Enter")
        time.sleep(0.5)
        logger.info(f"  日期已设置(range): {date_dash}")
        return

    date_inputs = page.locator(".el-date-editor input.el-input__inner")
    if date_inputs.count() >= 2:
        for idx in range(min(date_inputs.count(), 2)):
            date_inputs.nth(idx).click()
            time.sleep(0.2)
            date_inputs.nth(idx).press("Control+a")
            date_inputs.nth(idx).type(date_dash, delay=50)
            date_inputs.nth(idx).press("Enter")
            time.sleep(0.5)
        logger.info(f"  日期已设置(editor): {date_dash}")
        return

    logger.warning("  未找到日期输入框，尝试 fallback...")
    page.evaluate(f"""
        (function() {{
            var inputs = document.querySelectorAll('input');
            for (var i = 0; i < inputs.length; i++) {{
                var ph = (inputs[i].placeholder || '');
                if (ph.indexOf('日期') >= 0 || ph.indexOf('date') >= 0) {{
                    var nativeSet = Object.getOwnPropertyDescriptor(
                        window.HTMLInputElement.prototype, 'value').set;
                    nativeSet.call(inputs[i], '{date_dash}');
                    inputs[i].dispatchEvent(new Event('input', {{bubbles: true}}));
                    inputs[i].dispatchEvent(new Event('change', {{bubbles: true}}));
                }}
            }}
        }})()
    """)
    logger.info(f"  日期已设置(fallback): {date_dash}")


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

            output_dir = OUTPUT_DIR / "原始下载" / date_label
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
            logger.info(f"  营业概况下载全部完成！")
            logger.info(f"  输出目录: {output_dir}")
            logger.info(f"{'=' * 55}\n")

            # ── Part 1.5：下载银豹付交易账单 ─────────────────────────────
            logger.info(f"{'─' * 55}")
            logger.info(f"  下载银豹付交易账单")
            logger.info(f"{'─' * 55}")

            for i, up_config in enumerate(UNIONPAY_STORE_CONFIG):
                store_short = up_config["store_short"]
                logger.info(f"{'─' * 40}")
                logger.info(f"  门店 {i + 1}/{len(UNIONPAY_STORE_CONFIG)}: {store_short}")
                logger.info(f"{'─' * 40}")

                logger.info("  [导航] 前往银豹付交易账单页面...")
                page.goto(UNIONPAY_BILL_URL)
                page.wait_for_load_state("networkidle", timeout=120_000)
                time.sleep(2)

                logger.info(f"  → 设置日期: {target_str}...")
                set_vue_date(page, target_str)
                time.sleep(1)

                select_unionpay_store(page, up_config)
                time.sleep(0.5)

                logger.info("  [搜索] 点击搜索...")
                click_by_text(page, "搜索", "搜索")
                page.wait_for_load_state("networkidle", timeout=150_000)
                time.sleep(3)

                bill_data = scrape_unionpay_bill(page)
                save_unionpay_bill_csv(bill_data, store_short, date_label, output_dir)

            logger.info(f"{'=' * 55}")
            logger.info(f"  银豹付交易账单下载全部完成！")
            logger.info(f"  输出目录: {output_dir}")
            logger.info(f"{'=' * 55}\n")

            # ── Part 1.6：下载会员消费汇总表 ─────────────────────────────
            logger.info(f"{'─' * 55}")
            logger.info(f"  下载会员消费汇总表")
            logger.info(f"{'─' * 55}")

            for i, cs_config in enumerate(CUSTOMER_SUMMARY_STORE_CONFIG):
                store_short = cs_config["store_short"]
                logger.info(f"{'─' * 40}")
                logger.info(f"  门店 {i + 1}/{len(CUSTOMER_SUMMARY_STORE_CONFIG)}: {store_short}")
                logger.info(f"{'─' * 40}")

                logger.info("  [导航] 前往会员消费汇总表页面...")
                page.goto(CUSTOMER_SUMMARY_URL)
                page.wait_for_load_state("networkidle", timeout=120_000)
                time.sleep(2)

                select_store_type_consumption(page)

                select_stores_multi(page, cs_config["select_items"])

                date_dash = target_str.replace(".", "-")
                logger.info(f"  → 设置统计时间: {date_dash}...")
                result = page.evaluate(f"""
                    (function() {{
                        function setVal(id, val) {{
                            var inp = document.getElementById(id);
                            if (!inp) return 'not found: ' + id;
                            if (window.jQuery && jQuery.fn.datepicker) {{
                                jQuery(inp).datepicker('setDate', val);
                            }}
                            var nativeSet = Object.getOwnPropertyDescriptor(
                                HTMLInputElement.prototype, 'value').set;
                            nativeSet.call(inp, val);
                            inp.dispatchEvent(new Event('input', {{bubbles: true}}));
                            inp.dispatchEvent(new Event('change', {{bubbles: true}}));
                            inp.dispatchEvent(new Event('blur', {{bubbles: true}}));
                            return inp.value;
                        }}
                        var r1 = setVal('txt_startDatetime', '{date_dash}');
                        var r2 = setVal('txt_endDatetime', '{date_dash}');
                        return 'start=' + r1 + ', end=' + r2;
                    }})()
                """)
                logger.info(f"  统计时间: {result}")

                logger.info("  [查询] 执行查询...")
                page.locator("#btnSearch").click()
                page.wait_for_load_state("networkidle", timeout=150_000)
                time.sleep(3)

                logger.info("  [导出] 点击导出销售单据...")
                click_by_text(page, "导出销售单据", "导出销售单据")
                time.sleep(2)

                logger.info("  [导出] 点击弹窗中的导出...")
                with page.expect_download(timeout=180_000) as dl_info:
                    click_by_text(page, "导出", "弹窗导出")

                download = dl_info.value
                logger.info(f"  下载文件名: {download.suggested_filename}")

                dest = output_dir / f"会员消费汇总表_{store_short}_{date_label}.xlsx"
                download.save_as(dest)
                logger.info(f"  已保存到: {dest}")

            logger.info(f"{'=' * 55}")
            logger.info(f"  会员消费汇总表下载全部完成！")
            logger.info(f"  输出目录: {output_dir}")
            logger.info(f"{'=' * 55}\n")

            # ── Part 3：下载美团外卖账单明细 ─────────────────────────────
            logger.info(f"{'─' * 55}")
            logger.info(f"  Part 3：下载美团外卖账单明细")
            logger.info(f"{'─' * 55}")

            for mt_idx, mt_config in enumerate(MEITUAN_STORE_CONFIG):
                mt_store_short = mt_config["store_short"]
                mt_port = mt_config["port"]
                mt_user_data_dir = mt_config["user_data_dir"]

                logger.info(f"{'─' * 40}")
                logger.info(f"  美团门店 {mt_idx + 1}/{len(MEITUAN_STORE_CONFIG)}: {mt_store_short}")
                logger.info(f"{'─' * 40}")

                logger.info(f"  启动 Chrome (port={mt_port}, profile={mt_user_data_dir})...")
                chrome_process = subprocess.Popen([
                    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
                    "--headless=new",
                    f"--remote-debugging-port={mt_port}",
                    f"--user-data-dir={mt_user_data_dir}",
                ])
                time.sleep(5)

                chrome_browser = None
                chrome_page = None
                try:
                    logger.info("  连接到 Chrome...")
                    chrome_browser = pw.chromium.connect_over_cdp(f"http://localhost:{mt_port}")
                    chrome_context = chrome_browser.contexts[0]
                    chrome_page = chrome_context.new_page()
                    chrome_page.set_default_timeout(120000)
                    chrome_page.set_default_navigation_timeout(120000)

                    logger.info("  [导航] 前往美团外卖账单明细页...")
                    chrome_page.goto(MEITUAN_DOWNLOAD_URL)
                    chrome_page.wait_for_load_state("networkidle", timeout=120_000)
                    time.sleep(3)

                    # 设置日期
                    date_input_value = f"{date_label} 日账单"
                    logger.info(f"  → 设置日期: {date_input_value}...")
                    date_input = chrome_page.locator(".select-input-wrapper .roo-input")
                    date_input.click()
                    time.sleep(0.5)
                    date_input.press("Control+a")
                    date_input.type(date_input_value, delay=50)
                    date_input.press("Enter")
                    time.sleep(0.5)
                    chrome_page.locator("body").click(position={"x": 0, "y": 0})
                    logger.info(f"  已设置日期: {date_input_value}")

                    # 等待数据刷新
                    logger.info("  等待数据刷新...")
                    time.sleep(5)

                    # 从 tfoot 总计行抓取数据
                    logger.info("  → 抓取账单数据...")
                    bill_data = chrome_page.evaluate("""
                        (function() {
                            var result = {};
                            var tfoot = document.querySelector('.bill-charge-table tfoot tr');
                            if (tfoot) {
                                var tds = tfoot.querySelectorAll('td');
                                result['商品总价'] = parseFloat(tds[1].textContent.trim()) || 0;
                                result['打包费'] = parseFloat(tds[2].textContent.trim()) || 0;
                                result['商家对顾客的活动补贴'] = parseFloat(tds[3].textContent.trim()) || 0;
                                result['佣金'] = parseFloat(tds[6].textContent.trim()) || 0;
                                result['配送服务费'] = parseFloat(tds[7].textContent.trim()) || 0;
                            }
                            var tabs = document.querySelectorAll('.roo-tabs-nav .tab-item a');
                            for (var i = 0; i < tabs.length; i++) {
                                var text = tabs[i].textContent.trim();
                                if (text.indexOf('其它类') >= 0 || text.indexOf('其他类') >= 0) {
                                    var match = text.match(/([-\\d.]+)\\s*$/);
                                    result['其他类'] = match ? parseFloat(match[1]) : 0;
                                    break;
                                }
                            }
                            return result;
                        })()
                    """)
                    logger.info(f"  抓取数据: {json.dumps(bill_data, ensure_ascii=False)}")

                    # 保存为 CSV
                    meituan_csv_fields = ["商品总价", "打包费", "商家对顾客的活动补贴", "佣金", "配送服务费", "其他类"]
                    csv_file = output_dir / f"账单明细_{mt_store_short}_{date_label}.csv"
                    with open(csv_file, "w", newline="", encoding="utf-8-sig") as f:
                        writer = csv.writer(f)
                        writer.writerow(["项目", "金额"])
                        for key in meituan_csv_fields:
                            writer.writerow([key, bill_data.get(key, 0)])
                    logger.info(f"  已保存CSV: {csv_file}")

                except Exception as e:
                    logger.error(f"美团外卖账单明细下载失败 ({mt_store_short}): {e}")
                    if chrome_page:
                        try:
                            screenshot = OUTPUT_DIR / f"meituan_error_{mt_store_short}.png"
                            chrome_page.screenshot(path=str(screenshot))
                            logger.info(f"  错误截图已保存: {screenshot}")
                        except Exception:
                            pass
                finally:
                    if chrome_page:
                        try:
                            chrome_page.close()
                        except Exception:
                            pass
                    if chrome_browser:
                        try:
                            chrome_browser.close()
                        except Exception:
                            pass
                    try:
                        chrome_process.terminate()
                    except Exception:
                        pass

            logger.info(f"{'=' * 55}")
            logger.info(f"  美团外卖账单明细下载全部完成！")
            logger.info(f"{'=' * 55}\n")

            # ── Part 4：格式化数据并写入月度统计表 ──────────────────────────
            logger.info(f"{'─' * 55}")
            logger.info(f"  Part 4：格式化数据并写入月度统计表")
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
