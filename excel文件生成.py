import os
import pandas as pd
import tkinter as tk
from tkinter import simpledialog
import xlsxwriter  # 使用 xlsxwriter 替代 openpyxl

# 定义文件夹路径
folder_path = r'C:\excel'  # 使用原始字符串（r''）避免路径中的转义字符

def parse_column_range(range_str):
    """
    解析用户输入的列范围字符串，返回列号列表。
    例如，'1-3,5-9' 解析为 [0, 1, 2, 4, 5, 6, 7, 8]（Python中的列索引从0开始）。
    """
    column_indices = []
    for part in range_str.split(','):
        if '-' in part:
            start, end = map(int, part.split('-'))
            column_indices.extend(range(start-1, end))  # 转换为0-based索引
        else:
            column_indices.append(int(part)-1)  # 单个列，转为0-based索引
    return column_indices

def process_csv(file_path, n_columns_range, m_columns_range):
    # 读取整个CSV文件
    df = pd.read_csv(file_path, encoding='gbk')

    # 提取 N 列的数据用于第一个工作表
    df_original = df.iloc[:, n_columns_range].copy()

    # 提取 M 列的数据用于第二个工作表
    df_processed = df.iloc[:, m_columns_range].copy()

    # 定义输出文件名（保持原文件名，但扩展名为xlsx）
    output_file_path = os.path.splitext(file_path)[0] + '.xlsx'

    # 使用 xlsxwriter 写入数据
    with pd.ExcelWriter(output_file_path, engine='xlsxwriter') as writer:
        # 将原始数据写入第一个工作表
        df_original.to_excel(writer, sheet_name='KPI,MR处理', index=False)
        # 将处理后的数据写入第二个工作表
        df_processed.to_excel(writer, sheet_name='数据处理', index=False)

    return output_file_path


def process_files_in_folder(n_columns_range, m_columns_range):
    """遍历文件夹中的所有 CSV 文件并处理"""
    for filename in os.listdir(folder_path):
        if filename.endswith('.csv'):
            file_path = os.path.join(folder_path, filename)
            output_file_path = process_csv(file_path, n_columns_range, m_columns_range)
            print(f'Processed and saved {filename} as {os.path.basename(output_file_path)}')


def main():
    """
    统一入口，调用 process_files_in_folder()，用于外部调用
    """
    # 创建一个Tkinter的隐藏根窗口
    root = tk.Tk()
    root.withdraw()

    # 弹出对话框获取用户输入的列范围
    n_range_str = simpledialog.askstring("输入", "请输入提取到第一个表的列数据范围，用逗号分隔（例如 1-3,5-9）：",parent=root)
    m_range_str = simpledialog.askstring("输入", "请输入提取到第二个表的列数据范围，用逗号分隔（例如 1-5,7-9）：",parent=root)

    # 解析列范围
    n_columns_range = parse_column_range(n_range_str)
    m_columns_range = parse_column_range(m_range_str)

    # 处理文件夹中的所有CSV文件
    process_files_in_folder(n_columns_range, m_columns_range)
    root.destroy()

if __name__ == "__main__":
    main()
