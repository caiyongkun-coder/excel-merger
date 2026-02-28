#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel 表格整理程序 1.2

功能：
- 读取指定文件夹下的所有 Excel 文件（.xlsx / .xls）
- 识别每个文件首列的"用途"合并单元格
- 提取合并单元格覆盖的所有行数据
- 格式化输出到新的 Excel 文件（使用制表符对齐）
- 支持多 sheet：同一文件的所有 sheet 合并到目标的一个 sheet（不同行）
- 不同源文件 → 不同的目标 sheet
- 优化 Office 兼容性（设置 workbook 属性）

使用方法：
    python excel_merger.py <输入文件夹> <输出文件名（不含扩展名）>
    
示例：
    python excel_merger.py /home/caiyongkun/caishufen/2026-02-28/ 2026-02-28
"""

import os
import sys
from pathlib import Path
from datetime import datetime

try:
    import openpyxl
    from openpyxl.styles import Alignment
except ImportError:
    print("错误：请安装 openpyxl 库：pip install openpyxl")
    sys.exit(1)

try:
    import xlrd
except ImportError:
    print("错误：请安装 xlrd 库：pip install xlrd")
    sys.exit(1)


def contains_keyword(text, keyword="用途"):
    """
    检查文本是否包含关键词（支持空格变体）
    例如："用途"、"用 途"、"用  途" 等
    """
    if not text:
        return False
    text_str = str(text)
    # 移除所有空格后检查
    text_no_space = text_str.replace(" ", "")
    return keyword in text_no_space


def find_merged_cell_range(ws, keyword="用途"):
    """
    在工作表中查找包含关键词的合并单元格（在 A 列）
    返回合并单元格的起始行和结束行
    """
    for merged_range in ws.merged_cells.ranges:
        min_col = merged_range.min_col
        max_col = merged_range.max_col
        min_row = merged_range.min_row
        max_row = merged_range.max_row
        
        if min_col == 1 and max_col == 1:
            cell_value = ws.cell(row=min_row, column=1).value
            if contains_keyword(cell_value, keyword):
                return (min_row, max_row)
    
    return None


def extract_data_from_xlsx(file_path, keyword="用途"):
    """从 xlsx 文件中提取数据，支持文件中的所有 sheet"""
    wb = openpyxl.load_workbook(file_path, data_only=True)
    
    all_rows_data = []
    max_column = 0
    
    # 遍历文件中的所有 sheet
    for ws in wb.worksheets:
        merged_range = find_merged_cell_range(ws, keyword)
        
        if not merged_range:
            continue
        
        start_row, end_row = merged_range
        
        # 计算该 sheet 的最大列数
        sheet_max_col = 0
        for row in ws.iter_rows(min_row=start_row, max_row=end_row):
            for cell in row:
                if cell.value is not None and str(cell.value).strip():
                    sheet_max_col = max(sheet_max_col, cell.column)
        
        max_column = max(max_column, sheet_max_col)
        
        # 提取该 sheet 的数据
        for row_idx in range(start_row, end_row + 1):
            row_data = []
            for col_idx in range(1, sheet_max_col + 1):
                cell_value = ws.cell(row=row_idx, column=col_idx).value
                if cell_value is None:
                    row_data.append("")
                else:
                    row_data.append(str(cell_value).strip())
            all_rows_data.append(row_data)
    
    if not all_rows_data:
        print(f"  警告：在 {file_path.name} 中未找到包含'{keyword}'的合并单元格")
        return [], 0
    
    return all_rows_data, max_column


def extract_data_from_xls(file_path, keyword="用途"):
    """从 xls 文件中提取数据（使用 xlrd）"""
    wb = xlrd.open_workbook(file_path)
    ws = wb.sheet_by_index(0)
    
    start_row = None
    end_row = None
    
    merged_cells = ws.merged_cells
    for merged in merged_cells:
        row_start, row_end, col_start, col_end = merged
        
        if col_start == 0 and col_end == 1:
            cell_value = ws.cell_value(row_start, 0)
            if contains_keyword(cell_value, keyword):
                start_row = row_start
                end_row = row_end - 1
                break
    
    if start_row is None:
        for row_idx in range(ws.nrows):
            cell_value = ws.cell_value(row_idx, 0)
            if contains_keyword(cell_value, keyword):
                start_row = row_idx
                end_row = row_idx
                for next_row in range(row_idx + 1, ws.nrows):
                    next_value = ws.cell_value(next_row, 0)
                    if next_value is None or (next_value and not contains_keyword(next_value, keyword)):
                        break
                    end_row = next_row
                break
    
    if start_row is None:
        print(f"  警告：在 {file_path.name} 中未找到包含'{keyword}'的合并单元格")
        return [], 0
    
    max_column = 0
    for row_idx in range(start_row, end_row + 1):
        for col_idx in range(ws.ncols):
            cell_value = ws.cell_value(row_idx, col_idx)
            if cell_value is not None and str(cell_value).strip():
                max_column = max(max_column, col_idx + 1)
    
    if max_column == 0:
        return [], 0
    
    rows_data = []
    for row_idx in range(start_row, end_row + 1):
        row_data = []
        for col_idx in range(max_column):
            cell_value = ws.cell_value(row_idx, col_idx)
            if cell_value is None:
                row_data.append("")
            else:
                row_data.append(str(cell_value).strip())
        rows_data.append(row_data)
    
    return rows_data, max_column


def format_data_to_text(rows_data, max_column):
    """将数据格式化为对齐的文本（使用制表符）"""
    if not rows_data or max_column == 0:
        return ""
    
    # 计算每列最大宽度
    col_widths = [0] * max_column
    for row in rows_data:
        for col_idx, cell_value in enumerate(row):
            col_widths[col_idx] = max(col_widths[col_idx], len(cell_value))
    
    formatted_lines = []
    for row in rows_data:
        # 用制表符连接，保持对齐
        line_parts = []
        for col_idx, cell_value in enumerate(row):
            line_parts.append(cell_value.ljust(col_widths[col_idx]))
        formatted_lines.append('\t'.join(line_parts))
    
    return "\n".join(formatted_lines)


def process_excel_files(input_folder, output_filename):
    """处理文件夹中的所有 Excel 文件并生成输出文件"""
    input_path = Path(input_folder)
    
    if not input_path.exists():
        print(f"错误：输入文件夹不存在：{input_folder}")
        sys.exit(1)
    
    if not input_path.is_dir():
        print(f"错误：输入路径不是文件夹：{input_folder}")
        sys.exit(1)
    
    excel_files = []
    for ext in ['*.xlsx', '*.xls']:
        excel_files.extend(input_path.glob(ext))
    
    if not excel_files:
        print(f"错误：在 {input_folder} 中未找到 Excel 文件")
        sys.exit(1)
    
    print(f"找到 {len(excel_files)} 个 Excel 文件")
    
    # 创建 Workbook 并设置属性以确保 Office 兼容性
    output_wb = openpyxl.Workbook()
    
    # 设置 workbook 属性（修复 Office 兼容性告警）
    output_wb.properties.creator = "Excel Merger"
    output_wb.properties.lastModifiedBy = "Excel Merger"
    output_wb.properties.created = datetime.now()
    output_wb.properties.modified = datetime.now()
    
    default_sheet = output_wb.active
    
    first_sheet = True
    
    for excel_file in sorted(excel_files):
        print(f"\n处理：{excel_file.name}")
        
        if excel_file.suffix.lower() == '.xlsx':
            rows_data, max_column = extract_data_from_xlsx(excel_file)
        elif excel_file.suffix.lower() == '.xls':
            rows_data, max_column = extract_data_from_xls(excel_file)
        else:
            print(f"  跳过不支持的文件格式：{excel_file.suffix}")
            continue
        
        if not rows_data:
            print(f"  警告：未提取到数据，跳过此文件")
            continue
        
        formatted_text = format_data_to_text(rows_data, max_column)
        
        sheet_name = excel_file.stem
        if len(sheet_name) > 31:
            sheet_name = sheet_name[:31]
        
        if first_sheet:
            output_ws = default_sheet
            output_ws.title = sheet_name
            first_sheet = False
        else:
            output_ws = output_wb.create_sheet(title=sheet_name)
        
        # 写入 A1 单元格，设置自动换行
        cell = output_ws.cell(row=1, column=1, value=formatted_text)
        cell.alignment = Alignment(wrap_text=True)
        
        # 自动计算列宽（根据内容长度 + 余量）
        # 估算：每行平均字符数，最大不超过 200
        avg_line_len = 0
        if formatted_text:
            lines = formatted_text.split('\n')
            if lines:
                avg_line_len = sum(len(line) for line in lines) // len(lines)
        
        estimated_width = min(avg_line_len + 10, 200)  # 留 10 个字符余量，最大 200
        output_ws.column_dimensions['A'].width = estimated_width
        
        print(f"  提取 {len(rows_data)} 行数据，写入 sheet: {sheet_name}")
    
    if first_sheet:
        output_ws = default_sheet
        cell = output_ws.cell(row=1, column=1, value="未找到任何数据")
        cell.alignment = Alignment(wrap_text=True)
        output_ws.column_dimensions['A'].width = 200
    
    output_path = Path(f"{output_filename}.xlsx")
    
    # 保存文件（使用标准参数）
    output_wb.save(output_path)
    
    print(f"\n✓ 完成！输出文件：{output_path.absolute()}")


def main():
    if len(sys.argv) != 3:
        print("Excel 表格整理程序 1.2")
        print("\n使用方法：")
        print(f"  python {sys.argv[0]} <输入文件夹> <输出文件名（不含扩展名）>")
        print("\n示例：")
        print(f"  python {sys.argv[0]} /home/caiyongkun/caishufen/2026-02-28/ 2026-02-28")
        sys.exit(1)
    
    input_folder = sys.argv[1]
    output_filename = sys.argv[2]
    
    print("=" * 60)
    print("Excel 表格整理程序 1.2")
    print("=" * 60)
    print(f"输入文件夹：{input_folder}")
    print(f"输出文件：{output_filename}.xlsx")
    print("=" * 60)
    
    process_excel_files(input_folder, output_filename)


if __name__ == "__main__":
    main()
