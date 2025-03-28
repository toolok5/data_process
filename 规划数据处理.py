import os
import pandas as pd
from tkinter import Tk, filedialog, simpledialog
from datetime import datetime, timedelta
import numpy as np

def select_files(root):
    """弹出对话框选择CSV文件"""
    file_paths = filedialog.askopenfilenames(filetypes=[("CSV files", "*.csv")], parent=root)
    return file_paths

def get_date_range(root):
    """弹出对话框输入日期范围"""
    date_range = simpledialog.askstring("输入日期范围", "请输入日期范围 (格式: YYYY/MM/DD-YYYY/MM/DD):", parent=root)
    return date_range

def calculate_weeks(start_date, end_date):
    """根据输入的日期范围计算周区间"""
    weeks = []

    # 计算第一周的结束日期（该周的周日）
    first_week_end = start_date + timedelta(days=(6 - start_date.weekday()))
    if first_week_end > end_date:
        first_week_end = end_date
    weeks.append((start_date, first_week_end))

    # 计算后续完整的周
    current_date = first_week_end + timedelta(days=1)
    while current_date <= end_date:
        week_start = current_date
        week_end = week_start + timedelta(days=6)
        if week_end > end_date:
            week_end = end_date
        weeks.append((week_start, week_end))
        current_date = week_end + timedelta(days=1)

    return weeks

def read_csv_with_encodings(file_path, encodings=('utf-8', 'gbk')):
    """尝试使用多种编码读取CSV文件"""
    for encoding in encodings:
        try:
            return pd.read_csv(file_path, encoding=encoding)
        except Exception as e:
            print(f"使用编码 {encoding} 读取 {file_path} 失败：{e}")
    print(f"无法读取文件：{file_path}，请检查文件编码。")
    return None

def process_files(file_paths, date_range, root):
    """处理文件并保存结果"""
    # 解析日期范围
    start_date_str, end_date_str = date_range.split('-')
    start_date = datetime.strptime(start_date_str, "%Y/%m/%d")
    end_date = datetime.strptime(end_date_str, "%Y/%m/%d")

    # 计算周区间
    weeks = calculate_weeks(start_date, end_date)
    week_strings = [f"{week[0].strftime('%Y/%m/%d')}-{week[1].strftime('%Y/%m/%d')}" for week in weeks]

    output_dir = "C:\\excel"
    os.makedirs(output_dir, exist_ok=True)

    for file_path in file_paths:
        # 尝试读取CSV文件
        df = read_csv_with_encodings(file_path)
        if df is None:
            continue  # 如果文件读取失败，跳过该文件

        # 检查最后两列是否包含 '弱覆盖比例_移动'，并删除匹配的列
        last_two_columns = df.columns[-2:]
        columns_to_drop = [col for col in last_two_columns if '弱覆盖比例_移动' in col]
        if columns_to_drop:
            df.drop(columns=columns_to_drop, inplace=True, errors='ignore')

        # 添加 '移动下行弱覆盖比例大于5%' 列
        if '下行弱覆盖MR比例(%)_移动' in df.columns:
            # 使用 NumPy 向量化操作提高效率
            df['移动下行弱覆盖比例大于5%'] = np.where(
                pd.to_numeric(df['下行弱覆盖MR比例(%)_移动'], errors='coerce') > 5, '是', '否'
            )
        else:
            print(f"文件 {file_path} 不包含“下行弱覆盖MR比例(%)_移动”列，无法添加‘移动下行弱覆盖比例大于5%’列。")
            continue

        # 根据第三列 "市" 分组
        if '市' not in df.columns:
            print(f"文件 {file_path} 不包含“市”列，跳过处理。")
            continue

        grouped = df.groupby('市')

        for city, group_data in grouped:
            # 将每组数据平均分配到指定周数
            num_weeks = len(weeks)
            split_data = [group_data.iloc[i::num_weeks].copy() for i in range(num_weeks)]

            for week_str, data in zip(week_strings, split_data):
                data.loc[:, '时间周期'] = week_str  # 使用 .loc 明确设置值

                # 文件名以周区间命名
                sanitized_week = week_str.replace("-", "_").replace("/", "-")
                output_file = os.path.join(output_dir, f"{sanitized_week}.csv")

                # 如果文件已经存在，追加数据
                if os.path.exists(output_file):
                    existing_data = pd.read_csv(output_file, encoding='gbk')
                    data = pd.concat([existing_data, data], ignore_index=True)

                # 保存数据到CSV文件
                data.to_csv(output_file, index=False, encoding='gbk')
                print(f"已将数据保存到 {output_file}")

def main():
    # 创建根窗口一次并在整个程序中重用
    root = Tk()
    root.withdraw()  # 隐藏根窗口

    # 选择文件和输入日期范围
    files = select_files(root)
    if not files:
        print("未选择文件，程序结束。")
    else:
        date_range = get_date_range(root)
        if not date_range:
            print("未输入日期范围，程序结束。")
        else:
            process_files(files, date_range, root)

    # 销毁Tk根窗口
    root.destroy()

if __name__ == "__main__":
    main()
