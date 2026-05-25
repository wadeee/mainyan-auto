"""
订货商品汇总看板 - 数据填入格式模板脚本
用途：将每日导出的订货数据 Excel 填入格式模板，删除示例行，更新汇总公式
依赖：pip install openpyxl
"""

import copy
from pathlib import Path
from datetime import date

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter


# ============================================================
# 配置区 - 按需修改
# ============================================================
BASE_DIR = Path("C:/Users/Wadec/Desktop/订货商品汇总看板/2026-05-27/原始下载")

# 数据源文件（每日导出的原始看板）
# today_str = date.today().strftime("%Y-%m-%d")
DATA_FILE = BASE_DIR / f"订货商品汇总看板_2026-05-27.xlsx"

# 格式模板文件（固定，不会变）
TEMPLATE_FILE = Path("订货商品汇总看板_格式化模板.xlsx")

# 输出文件
OUTPUT_FILE = BASE_DIR / f"订货商品汇总看板_格式化_2026-05-27.xlsx"

# 模板中示例数据所在行（第几行开始、共几行）
SAMPLE_ROW_START = 8    # 示例数据起始行
SAMPLE_ROW_COUNT = 2    # 示例数据行数

# 数据写入起始行（删除示例行后，从此行开始写入真实数据）
DATA_START_ROW = 8

# 表头行号（用于计算序号偏移量：序号 = ROW() - HEADER_ROW）
HEADER_ROW = 7

# 固定列映射（数据源 0-based 索引 → 模板列号 1-based）
# 门店列（G列起）由数据源中"合计订货量"之后的列动态决定
FIXED_COLUMN_MAP = [
    (1, 2),  # 商品分类 -> B
    (3, 3),  # 商品名称 -> C
    (4, 4),  # 规格 -> D
    (5, 5),  # 单位 -> E
]

TOTAL_ORDER_HEADER = "合计订货量"
# ============================================================


def copy_cell_style(src_cell, dst_cell):
    """将源单元格的样式（字体/填充/边框/对齐/数字格式）复制到目标单元格"""
    if src_cell.has_style:
        dst_cell.font = copy.copy(src_cell.font)
        dst_cell.fill = copy.copy(src_cell.fill)
        dst_cell.border = copy.copy(src_cell.border)
        dst_cell.alignment = copy.copy(src_cell.alignment)
        dst_cell.number_format = src_cell.number_format


def read_data_file(data_file: Path):
    """从数据源文件读取表头信息和数据行，动态识别门店列"""
    wb = load_workbook(data_file)
    ws = wb.active

    headers = [cell.value for cell in list(ws.iter_rows(min_row=1, max_row=1))[0]]

    total_col_idx = None
    for i, h in enumerate(headers):
        if h and TOTAL_ORDER_HEADER in str(h):
            total_col_idx = i
            break
    if total_col_idx is None:
        raise ValueError(f"数据源中找不到'{TOTAL_ORDER_HEADER}'列")

    store_columns = []
    for i in range(total_col_idx + 1, len(headers)):
        if headers[i] is not None and str(headers[i]).strip():
            store_columns.append((i, str(headers[i]).strip()))

    rows = []
    for row in ws.iter_rows(min_row=2):
        vals = [cell.value for cell in row]
        if any(v is not None and str(v).strip() != "" for v in vals):
            rows.append(vals)

    return total_col_idx, store_columns, rows


def merge_into_template(data_rows, total_col_idx, store_columns, template_file, output_file):
    wb = load_workbook(template_file)
    ws = wb.active

    store_count = len(store_columns)
    last_col = 6 + store_count
    template_max_col = ws.max_column

    column_map = list(FIXED_COLUMN_MAP)
    column_map.append((total_col_idx, 6))
    for offset, (src_idx, _) in enumerate(store_columns):
        column_map.append((src_idx, 7 + offset))

    # 更新第7行门店列标题
    for offset, (_, name) in enumerate(store_columns):
        ws.cell(row=HEADER_ROW, column=7 + offset).value = name

    # 清除模板中多余的门店列
    for col in range(last_col + 1, template_max_col + 1):
        for row in range(1, ws.max_row + 1):
            cell = ws.cell(row=row, column=col)
            cell.value = None
            cell.font = copy.copy(cell.font)
            cell.border = copy.copy(cell.border)

    # 如果数据门店列多于模板，从模板最后一个门店列复制样式到新列
    if last_col > template_max_col:
        for row in range(1, ws.max_row + 1):
            src = ws.cell(row=row, column=template_max_col)
            for new_col in range(template_max_col + 1, last_col + 1):
                copy_cell_style(src, ws.cell(row=row, column=new_col))

    # 更新顶部汇总区公式
    first_store = get_column_letter(7)
    last_store = get_column_letter(last_col)
    for r in range(2, 6):
        ws.cell(row=r, column=6).value = f"=SUM({first_store}{r}:{last_store}{r})"
    ws.cell(row=6, column=6).value = f"=SUM({first_store}6:{last_store}6)"

    # 保存示例行样式
    sample_cells = list(ws.iter_rows(
        min_row=SAMPLE_ROW_START,
        max_row=SAMPLE_ROW_START
    ))[0]

    ws.delete_rows(SAMPLE_ROW_START, SAMPLE_ROW_COUNT)
    ws.insert_rows(DATA_START_ROW, len(data_rows))

    for i, row_data in enumerate(data_rows):
        excel_row = DATA_START_ROW + i

        ws.cell(row=excel_row, column=1).value = f"=ROW()-{HEADER_ROW}"

        for src_idx, dst_col in column_map:
            value = row_data[src_idx] if src_idx < len(row_data) else None
            ws.cell(row=excel_row, column=dst_col).value = value

        for col_idx in range(1, last_col + 1):
            if col_idx <= len(sample_cells):
                src = sample_cells[col_idx - 1]
            else:
                src = sample_cells[len(sample_cells) - 1]
            copy_cell_style(src, ws.cell(row=excel_row, column=col_idx))

    last_data_row = DATA_START_ROW + len(data_rows) - 1

    # 更新第6行门店列公式：统一格式 =SUM({col}2:{col}5)
    for col_num in [7 + i for i in range(store_count)]:
        letter = get_column_letter(col_num)
        ws[f"{letter}6"] = f"=SUM({letter}2:{letter}5)"

    wb.save(output_file)
    return last_data_row


def main():
    print(f"数据源: {DATA_FILE}")
    print(f"模板  : {TEMPLATE_FILE}")
    print(f"输出  : {OUTPUT_FILE}")
    print()

    if not DATA_FILE.exists():
        raise FileNotFoundError(f"数据源文件不存在: {DATA_FILE}")
    if not TEMPLATE_FILE.exists():
        raise FileNotFoundError(f"模板文件不存在: {TEMPLATE_FILE}")

    print("读取数据源...")
    total_col_idx, store_columns, data_rows = read_data_file(DATA_FILE)
    print(f"  共 {len(data_rows)} 行数据")
    print(f"  合计订货量列: {get_column_letter(total_col_idx + 1)} (索引 {total_col_idx})")
    print(f"  门店列 ({len(store_columns)}): {', '.join(name for _, name in store_columns)}")

    print("填入模板...")
    last_row = merge_into_template(data_rows, total_col_idx, store_columns, TEMPLATE_FILE, OUTPUT_FILE)
    last_letter = get_column_letter(6 + len(store_columns))
    print(f"  数据区: A{DATA_START_ROW}:{last_letter}{last_row}")

    print()
    print(f"已生成: {OUTPUT_FILE}")


if __name__ == "__main__":
    main()
