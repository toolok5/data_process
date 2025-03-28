import tkinter as tk
from tkinter import filedialog
import pandas as pd
import numpy as np

# 为每个组生成随机数的函数
def generate_random_rru_traffic(df, group_by_column):
    # 获取唯一的组标识符
    unique_groups = df[group_by_column].unique()
    # 为每个组生成一个随机数，范围在111到1234568之间，并保留5位小数
    random_traffic = {group: np.round(np.random.uniform(111, 1234568.00001), 5) for group in unique_groups}
    # 创建一个新的DataFrame来存储随机数和组标识符
    random_df = pd.DataFrame.from_dict(random_traffic, orient='index', columns=['总流量MB']).reset_index()
    random_df.columns = [group_by_column, '总流量MB']
    # 将随机数DataFrame合并回原始DataFrame，使用left join以确保所有原始数据都保留
    return df.merge(random_df, on=group_by_column, how='left')

# 尝试使用不同编码读取CSV文件的函数
def read_csv_with_encodings(file_path):
    encodings = ['utf-8', 'gbk']  # 定义要尝试的编码列表
    for encoding in encodings:
        try:
            df = pd.read_csv(file_path, encoding=encoding)
            print(f"Successfully read {file_path} with encoding {encoding}.")
            return df
        except UnicodeDecodeError:
            print(f"Failed to read {file_path} with encoding {encoding}. Trying next encoding.")
    raise Exception(f"Failed to read {file_path} with all provided encodings.")

# 处理CSV文件的函数
def process_csv_files(files):
    for file_path in files:
        # 尝试读取CSV文件
        df = read_csv_with_encodings(file_path)

        # 检查'网元标识'列是否存在，否则使用'源小区CGI'作为分组依据
        group_by_column = '网元标识' if '网元标识' in df.columns else '源小区CGI'

        # 确保分组依据的列不是空的，否则会引发错误
        if df[group_by_column].isnull().all():
            raise ValueError(f"Column '{group_by_column}' in file {file_path} is empty or not found, cannot process the file.")

        # 生成随机数并合并回DataFrame
        df = generate_random_rru_traffic(df, group_by_column)

        # 将处理后的数据保存回原文件（注意：这会覆盖原文件），使用utf-8编码以避免潜在编码问题
        df.to_csv(file_path, index=False, encoding='gbk')

    print("处理完成！")

# 显示文件选择对话框的函数
def select_files():
    root = tk.Tk()
    root.withdraw()  # 隐藏Tkinter主窗口
    files = filedialog.askopenfilenames(filetypes=[("CSV files", "*.csv")])  # 只允许选择CSV文件
    return files

# 统一入口函数
def main():
    # 选择文件
    files = select_files()

    if files:
        # 处理选定的CSV文件
        process_csv_files(files)
    else:
        print("未选择任何文件，程序结束。")

if __name__ == "__main__":
    main()
