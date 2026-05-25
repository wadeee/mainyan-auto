"""
订货商品汇总看板 - 数据填入格式模板脚本
用途：将每日导出的订货数据 Excel 填入格式模板，删除示例行，更新汇总公式
依赖：pip install openpyxl
"""

import copy
from pathlib import Path
from datetime import date

from openpyxl import load_workbook


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

# 数据源列映射（0-based 索引）→ 模板目标列（1-based）
# 格式: (数据源列索引, 模板列号, 说明)
COLUMN_MAP = [
    (1,  2, "品类名称"),
    (3,  3, "商品名称"),
    (4,  4, "规格"),
    (5,  5, "单位"),
    (6,  6, "合计订量"),
    (7,  7, "顺义总仓(春)"),
    (8,  8, "天津站(泰春)"),
    (9,  9, "顺义仓库(春)"),
]

# 顶部汇总区需要更新 SUM 公式的列号（1-based）
SUM_FORMULA_COLS = [6, 7, 8, 9]   # F、G、H、I 列
# ============================================================


def copy_cell_style(src_cell, dst_cell):
    """将源单元格的样式（字体/填充/边框/对齐/数字格式）复制到目标单元格"""
    if src_cell.has_style:
        dst_cell.font = copy.copy(src_cell.font)
        dst_cell.fill = copy.copy(src_cell.fill)
        dst_cell.border = copy.copy(src_cell.border)
        dst_cell.alignment = copy.copy(src_cell.alignment)
        dst_cell.number_format = src_cell.number_format


def read_data_rows(data_file: Path) -> list[list]:
    """从数据源文件读取所有数据行（跳过第1行表头，跳过空行）"""
    wb = load_workbook(data_file)
    ws = wb.active
    rows = []
    for row in ws.iter_rows(min_row=2):
        vals = [cell.value for cell in row]
        if any(v is not None and str(v).strip() != "" for v in vals):
            rows.append(vals)
    return rows


def merge_into_template(data_rows: list[list], template_file: Path, output_file: Path):
    wb = load_workbook(template_file)
    ws = wb.active

    # 保存示例行样式（用于新数据行的格式复制）
    sample_cells = list(ws.iter_rows(
        min_row=SAMPLE_ROW_START,
        max_row=SAMPLE_ROW_START
    ))[0]

    # 删除示例数据行
    ws.delete_rows(SAMPLE_ROW_START, SAMPLE_ROW_COUNT)

    # 在数据起始行插入足够的空行
    ws.insert_rows(DATA_START_ROW, len(data_rows))

    # 逐行写入数据
    for i, row_data in enumerate(data_rows):
        excel_row = DATA_START_ROW + i

        # 序号列（A列）用公式自动编号
        ws.cell(row=excel_row, column=1).value = f"=ROW()-{HEADER_ROW}"

        # 按列映射写入各字段
        for src_idx, dst_col, _ in COLUMN_MAP:
            value = row_data[src_idx] if src_idx < len(row_data) else None
            ws.cell(row=excel_row, column=dst_col).value = value

        # 复制示例行样式到新行
        for col_idx, sample_cell in enumerate(sample_cells, start=1):
            copy_cell_style(sample_cell, ws.cell(row=excel_row, column=col_idx))

    # 更新顶部汇总区（第6行）的 SUM 公式范围
    last_data_row = DATA_START_ROW + len(data_rows) - 1
    col_letters = {1: "A", 2: "B", 3: "C", 4: "D", 5: "E",
                   6: "F", 7: "G", 8: "H", 9: "I", 10: "J"}
    for col_num in SUM_FORMULA_COLS:
        letter = col_letters[col_num]
        ws[f"{letter}6"] = f"=SUM({letter}{DATA_START_ROW}:{letter}{last_data_row})"

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
    data_rows = read_data_rows(DATA_FILE)
    print(f"  共 {len(data_rows)} 行数据")

    print("填入模板...")
    last_row = merge_into_template(data_rows, TEMPLATE_FILE, OUTPUT_FILE)
    print(f"  数据区: A{DATA_START_ROW}:I{last_row}")

    print()
    print(f"✅ 已生成: {OUTPUT_FILE}")


if __name__ == "__main__":
    main()