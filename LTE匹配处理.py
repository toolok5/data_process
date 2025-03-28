import pandas as pd
import tkinter as tk
from tkinter import filedialog


def read_csv_with_fallback(file_path):
    """
    尝试使用不同的编码读取CSV文件
    """
    encodings = ['utf-8', 'gbk', 'latin1']
    for encoding in encodings:
        try:
            df = pd.read_csv(file_path, encoding=encoding)
            print(f"成功读取文件 {file_path}，编码方式为 {encoding}")
            return df
        except UnicodeDecodeError:
            continue
    raise ValueError(f"Failed to read the file {file_path} with available encodings.")


def select_csv_files(title="选择文件", initialdir=None):
    """
    弹出文件选择框，返回选中文件路径
    """
    root = tk.Tk()
    root.withdraw()  # 隐藏主窗口
    file_paths = filedialog.askopenfilenames(title=title,
                                             initialdir=initialdir,
                                             filetypes=[("CSV files", "*.csv")],parent=root)
    return file_paths


def process_files():
    """
    主逻辑：选择CSV文件、合并数据并保存
    """
    mr_files = select_csv_files(title="选择MR的CSV文件", initialdir="D:\\新建文件夹 (2)\\2404MR")
    performance_files = select_csv_files(title="选择性能CSV文件", initialdir="D:\新建文件夹 (2)\性能")

    # 创建字典保存文件路径，按照日期分组
    mr_files_dict = {}
    performance_files_dict = {}

    # 按日期分组文件
    for file in mr_files:
        date = file.split('_')[-1].split('.')[0]
        if date not in mr_files_dict:
            mr_files_dict[date] = []
        mr_files_dict[date].append(file)

    for file in performance_files:
        date = file.split('_')[-1].split('.')[0]
        if date not in performance_files_dict:
            performance_files_dict[date] = []
        performance_files_dict[date].append(file)

    # 合并每日数据并保存
    for date in mr_files_dict.keys():
        if date in performance_files_dict:
            # 读取每日的MR和性能指标数据
            mr_data_list = [read_csv_with_fallback(file) for file in mr_files_dict[date]]
            performance_data_list = [read_csv_with_fallback(file) for file in performance_files_dict[date]]

            # 合并同一天的所有文件
            mr_data = pd.concat(mr_data_list, ignore_index=True)
            performance_data = pd.concat(performance_data_list, ignore_index=True)

            # 合并两个表格，依据列名“小区CGI”
            merged_table = pd.merge(performance_data, mr_data, on='小区CGI', how='left')

            # 定义一个函数来判断值是否满足指定条件
            def check_success_rate(value, threshold, comparator='>='):
                if pd.isna(value):
                    return '是'
                if comparator == '>=' and value >= threshold:
                    return '是'
                elif comparator == '<=' and value <= threshold:
                    return '是'
                return '否'

            # 应用函数到指定列
            merged_table['无线接通率≥98%'] = merged_table['无线接通率'].apply(lambda x: check_success_rate(x, 98))
            merged_table['切换成功率≥97%'] = merged_table['切换成功率'].apply(lambda x: check_success_rate(x, 97))
            merged_table['RRC连接建立成功率≥98%'] = merged_table['RRC连接建立成功率'].apply(lambda x: check_success_rate(x, 98))
            merged_table['E-RAB建立成功率≥98%'] = merged_table['E-RAB建立成功率'].apply(lambda x: check_success_rate(x, 98))
            merged_table['无线掉线率≤2%'] = merged_table['无线掉线率'].apply(lambda x: check_success_rate(x, 2, '<='))
            merged_table['E-RAB掉线率(小区级)≤2%'] = merged_table['E-RAB掉线率(小区级)'].apply(lambda x: check_success_rate(x, 2, '<='))
            merged_table['无线接通率(QCI=1)≥98%'] = merged_table['无线接通率(QCI=1)'].apply(lambda x: check_success_rate(x, 98))
            merged_table['E-RAB掉线率(QCI=1)(小区级)≤2%'] = merged_table['E-RAB掉线率(QCI=1)(小区级)'].apply(
                lambda x: check_success_rate(x, 2, '<='))
            merged_table['小区RB上行平均干扰电平_全天<-105dBm'] = merged_table['小区RB上行平均干扰电平_全天'].apply(
                lambda x: check_success_rate(x, -105, '<='))
            merged_table['RSRP大于等于-110覆盖率(%)≥80%'] = merged_table['RSRP大于等于-110覆盖率(%)'].apply(
                lambda x: check_success_rate(x, 80))

            # 保存合并后的表格到新的CSV文件
            merged_filename = f"C:\\excel\\{date}.csv"
            merged_table.to_csv(merged_filename, index=False, encoding='gbk')

            print(f"文件 {merged_filename} 已成功合并并保存")

    print("所有文件已成功合并并保存")


def main():
    """
    主入口：供外部调用的统一接口
    """
    process_files()

if __name__ == "__main__":
    main()