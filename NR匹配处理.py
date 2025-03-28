import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog, simpledialog
from datetime import datetime, timedelta
import random


def select_files(title="选择文件"):
    root = tk.Tk()
    root.withdraw()  # 隐藏根窗口
    file_paths = filedialog.askopenfilenames(
        title=title,
        filetypes=[("CSV Files", "*.csv")],parent=root
    )

    return file_paths


def get_week_range(date_str):
    date = datetime.strptime(date_str, "%Y%m%d")
    monday = date - timedelta(days=date.weekday())
    sunday = monday + timedelta(days=6)
    return f"{monday.strftime('%Y%m%d')}_{sunday.strftime('%Y%m%d')}"


def read_performance_csv(files, base_num_rows):
    data = []
    for file in files:
        # 每个文件独立随机行数
        num_rows = max(1, base_num_rows + random.randint(-300, 300))
        df = pd.read_csv(file, nrows=num_rows, encoding='gbk')
        date_str = os.path.basename(file).split('_')[-1].replace('.csv', '')
        df['日期'] = date_str
        data.append(df)
    return pd.concat(data, ignore_index=True)


def read_mr_csv(files):
    data = []
    required_columns = ['小区码CI', 'RSRP覆盖率(≥-110dbm)', 'RSRP采样点数(≥-110dbm)', 'MRO样本总数']

    for file in files:
        df = pd.read_csv(file, encoding='gbk', usecols=required_columns)
        date_str = os.path.basename(file).split('_')[-1].replace('.csv', '')
        df['日期'] = date_str
        data.append(df)
    return pd.concat(data, ignore_index=True)


def calculate_conditions(row):
    conditions = {}

    if row['NSA/SA标志'] == 1:
        # NSA/SA标志 = 1 的情况
        conditions.update({
            '无线接通率≥95%': '是' if row['无线接通率'] >= 95 else '否',
            'NSA SgNB添加成功率≥95%': '',
            '5-5切换成功率≥95%': '是' if row['切换成功率'] >= 95 else '否',
            '5-4系统间切换成功率≥95%': '是' if row['切换至LTE准备请求次数'] == 0 or (
                        row['切换至LTE成功次数'] / row['切换至LTE准备请求次数'] >= 0.95) else '否',
            '无线掉线率（小区级）＜1.5%': '是' if row['无线掉线率（小区级）'] < 1.5 else '否',
            'SN异常释放率＜2%': '',
            '上行每PRB的接收干扰噪声平均值＜-105dBm': '是' if row['小区RB上行平均干扰电平'] < -105 else '否',
            'EPS FallBack接通率≥90%': '是' if pd.isna(row['EPSFB成功率']) or row['EPSFB成功率'] >= 90 else '否',
            'MR覆盖率RSRP>-110dbm采样点数/总采样点数≥90%': '是' if 'RSRP覆盖率(≥-110dbm)' in row and (
                        pd.isna(row['RSRP覆盖率(≥-110dbm)']) or row['RSRP覆盖率(≥-110dbm)'] >= 90) else '否'
        })
    else:
        # NSA/SA标志 = 0 的情况
        conditions.update({
            '无线接通率≥95%': '',
            'NSA SgNB添加成功率≥95%': '是' if pd.isna(row['NSA SgNB添加成功率']) or row['NSA SgNB添加成功率'] >= 95 else '否',
            '5-5切换成功率≥95%': '',
            '5-4系统间切换成功率≥95%': '',
            '无线掉线率（小区级）＜1.5%': '',
            'SN异常释放率＜2%': '是' if pd.isna(row['SN异常释放率（NSA SgNB掉话率）']) or row['SN异常释放率（NSA SgNB掉话率）'] < 2 else '否',
            '上行每PRB的接收干扰噪声平均值＜-105dBm': '',
            'EPS FallBack接通率≥90%': '是' if pd.isna(row['EPSFB成功率']) or row['EPSFB成功率'] >= 90 else '否',
            'MR覆盖率RSRP>-110dbm采样点数/总采样点数≥90%': '是' if 'RSRP覆盖率(≥-110dbm)' in row and (
                        pd.isna(row['RSRP覆盖率(≥-110dbm)']) or row['RSRP覆盖率(≥-110dbm)'] >= 90) else '否'
        })

    return pd.Series(conditions)


def main():
    # 1. 获取基准行数
    root = tk.Tk()  # 重新创建 Tkinter 根窗口
    root.withdraw()
    base_num_rows = simpledialog.askinteger("输入", "请输入基准行数：", minvalue=1, parent=root)
    if not base_num_rows:
        return

    # 2. 选择性能文件
    performance_files = select_files("选择性能文件")
    if not performance_files:
        return
    performance_data = read_performance_csv(performance_files, base_num_rows)

    # 3. 选择MR文件
    mr_files = select_files("选择MR文件")
    if not mr_files:
        return
    mr_data = read_mr_csv(mr_files)

    # 4. 日期匹配和数据合并
    performance_data['日期'] = pd.to_datetime(performance_data['日期'])
    mr_data['日期'] = pd.to_datetime(mr_data['日期'])

    # 创建输出目录
    if not os.path.exists('C:/excel'):
        os.makedirs('C:/excel')

    # 按日期分组处理数据
    for date in performance_data['日期'].unique():
        perf_day = performance_data[performance_data['日期'] == date]
        mr_day = mr_data[mr_data['日期'] == date]

        if not perf_day.empty and not mr_day.empty:
            # 合并当天的数据
            day_data = pd.merge(perf_day, mr_day, on=['小区码CI', '日期'], how='left')

            # 计算条件列
            condition_columns = day_data.apply(calculate_conditions, axis=1)
            final_day_data = pd.concat([day_data, condition_columns], axis=1)

            # 获取对应的周范围并保存
            date_str = date.strftime('%Y%m%d')
            week_range = get_week_range(date_str)
            output_path = f'C:/excel/{week_range}.csv'

            # 如果文件已存在，追加数据；否则创建新文件
            if os.path.exists(output_path):
                existing_data = pd.read_csv(output_path, encoding='gbk')
                final_day_data = pd.concat([existing_data, final_day_data], ignore_index=True)

            final_day_data.to_csv(output_path, index=False, encoding='gbk')
            print(f"数据已保存到: {output_path}")


if __name__ == "__main__":
    main()
