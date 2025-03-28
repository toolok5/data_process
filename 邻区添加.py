import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog, simpledialog
from datetime import datetime, timedelta

# 固定的输出路径
output_dir = r"C:\\excel"
if not os.path.exists(output_dir):
    os.makedirs(output_dir)

def get_week_range(start_date, date):
    """
    获取日期所属周的起始和结束日期
    """
    week_start = start_date + timedelta(weeks=((date - start_date).days // 7))
    week_end = week_start + timedelta(days=6)
    return week_start, week_end

def process_files(root):
    """ 处理文件并生成汇总结果 """
    file_paths = filedialog.askopenfilenames(title="选择多个CSV文件", filetypes=[("CSV files", "*.csv")], parent=root)
    if not file_paths:
        print("未选择任何文件！")
        return

    start_date = datetime(2024, 4, 1)
    aggregated_data = {}

    filter_condition = simpledialog.askstring("筛选条件", "请输入地市筛选条件（多个地市用逗号分隔，如'杭州,温州,宁波'）：", parent=root)
    if not filter_condition:
        print("未输入筛选条件，将处理所有地市的数据。")
        filter_cities = []  # 不设置筛选条件，相当于筛选所有地市
    else:
        filter_cities = filter_condition.split(',')
        filter_cities = [city.strip() for city in filter_cities]  # 去除可能的空格

    # 逐文件处理
    for file_path in file_paths:
        print(f"读取文件: {file_path}")

        try:
            # 指定 '创建时间' 列为日期格式，其他列作为字符串读取
            dtype = {col: 'str' for col in range(1, 100)}  # 假设最多100列，如果有更多列，可以调整
            df = pd.read_csv(file_path, encoding='gbk', usecols=range(15), dtype=dtype, parse_dates=['创建时间'],
                             date_format='%Y-%m-%d %H:%M:%S')
        except Exception as e:
            print(f"读取文件 {file_path} 时出错: {e}")
            continue

        # 确保列名一致
        if '创建时间' not in df.columns or '地市' not in df.columns:
            print(f"文件 {file_path} 缺少 '创建时间' 列，跳过！")
            continue

        # 填充 10-14 列的空值为 ""
        columns_to_check = df.columns[9:14]  # 假设列从1开始计数
        df[columns_to_check] = df[columns_to_check].fillna("")

        # 过滤出有效数据
        df = df.dropna(subset=['创建时间'])
        df = df[df['创建时间'] >= start_date]

        # 根据地市进行筛选
        if filter_cities:
            df = df[df['地市'].str.contains('|'.join(filter_cities), na=False, case=False)]

        # 检查解析后的数据
        print(f"文件 {file_path} 中有效的创建时间且符合地市筛选条件的数据行数: {len(df)}")

        if len(df) == 0:
            print(f"文件 {file_path} 没有符合条件的有效创建时间数据或地市筛选条件，跳过！")
            continue

        # 计算每个行的所属周
        df['周开始'], df['周结束'] = zip(*df['创建时间'].apply(lambda x: get_week_range(start_date, x)))

        # 将数据根据周进行分组
        for week_key, group in df.groupby(['周开始', '周结束']):
            week_key_str = f"{week_key[0].strftime('%Y-%m-%d')}_{week_key[1].strftime('%Y-%m-%d')}"
            if week_key_str not in aggregated_data:
                aggregated_data[week_key_str] = []

            aggregated_data[week_key_str].append(group)

    # 保存数据到不同的CSV文件
    if aggregated_data:
        print(f"分组完成，共有 {len(aggregated_data)} 周的数据。")
        for week_key, groups in aggregated_data.items():
            file_name = f"集中参数数据处理-邻区添加{week_key}.csv"
            output_path = os.path.join(output_dir, file_name)

            # 合并所有行数据
            week_df = pd.concat(groups, ignore_index=True)
            week_df = week_df.drop(columns=['周开始', '周结束'])

            # 替换列名中的 '基站名称' 为 '网元名称'
            week_df.columns = week_df.columns.str.replace('基站名称', '网元名称')

            # 写入文件，如果已存在则追加
            if os.path.exists(output_path):
                week_df.to_csv(output_path, mode='a', index=False, header=False, encoding='gbk')
            else:
                week_df.to_csv(output_path, index=False, encoding='gbk')

            print(f"保存文件: {output_path}")
    else:
        print("没有找到符合条件的数据，未生成文件。")

def main():
    """
    统一入口，调用 process_files()，用于外部调用
    """
    # 创建并隐藏根窗口
    root = tk.Tk()
    root.withdraw()  # 隐藏根窗口

    # 调用文件处理
    process_files(root)

    # 销毁根窗口
    root.destroy()

if __name__ == "__main__":
    main()
