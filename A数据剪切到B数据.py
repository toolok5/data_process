import os
import pandas as pd
import random
from tkinter import Tk, filedialog, simpledialog
from datetime import datetime, timedelta
import re


def select_file(title, filetypes, root):
    """弹出文件选择对话框并返回选中的文件路径。"""
    file_path = filedialog.askopenfilename(parent=root, title=title, filetypes=filetypes)
    return file_path


def modify_date_column(row, column_name, order_date):
    """
    修改指定行中指定列的日期部分，将日期部分替换为指定的日期。

    Args:
        row (pd.Series): 单行数据。
        column_name (str): 要修改的列名。
        order_date (str): 用于替换的日期（格式：YYYYMMDD）。

    Returns:
        str: 修改后的日期字符串。
    """
    if pd.notna(row[column_name]):
        try:
            # 将 order_date 格式化为 YYYY/MM/DD 格式
            formatted_order_date = f"{order_date[:4]}/{order_date[4:6]}/{order_date[6:8]}"
            print(f"替换 '{column_name}' 中的日期：{row[column_name]} -> {formatted_order_date}")
            # 尝试解析 '创建时间' 格式（第一种格式：YYYY/MM/DD HH:MM:SS）
            create_time = datetime.strptime(row[column_name], '%Y-%m-%d %H:%M:%S')
            # 将 '工单编号' 中提取的日期（YYYYMMDD）替换 '创建时间' 中的年月日
            create_time = create_time.replace(year=int(formatted_order_date[:4]),
                                               month=int(formatted_order_date[5:7]),
                                               day=int(formatted_order_date[8:10]))
            return create_time.strftime('%Y/%m/%d %H:%M:%S')
            print(f"替换后的日期：{create_time.strftime('%Y-%m-%d %H:%M:%S')}")
        except ValueError:


            try:
                # 如果解析失败（第二种格式：YYYY-MM-DD HH:MM:SS）
                create_time = datetime.strptime(row[column_name], '%Y/%m/%d %H:%M')
                # 同样将 '工单编号' 中提取的日期替换 '创建时间' 中的年月日
                create_time = create_time.replace(year=int(formatted_order_date[:4]),
                                                   month=int(formatted_order_date[5:7]),
                                                   day=int(formatted_order_date[8:10]))
                return create_time.strftime('%Y/%m/%d %H:%M')
                print(f"替换后的日期：{create_time.strftime('%Y/%m/%d %H:%M')}")
            except ValueError:
                # 如果两种格式都无法解析，返回原始时间
                return row[column_name]

    return row[column_name]


def modify_work_order_column(df, start_date, end_date, columns_to_modify):
    """
    修改 '工单编号' 列和其他指定列中的日期部分为 B.csv 文件名中提取的日期范围。

    Args:
        df (pd.DataFrame): 要处理的数据。
        start_date (datetime): 文件名中提取的开始日期。
        end_date (datetime): 文件名中提取的结束日期。
        columns_to_modify (list): 需要修改的列名列表。

    Returns:
        pd.DataFrame: 修改后的数据。
    """

    # 随机生成日期范围内的日期
    def random_date():
        delta = (end_date - start_date).days
        return (start_date + timedelta(days=random.randint(0, delta))).strftime('%Y%m%d')

    # 提取后四位数字
    if '工单编号' in df.columns:
        df['后四位'] = df['工单编号'].str[-4:] if '工单编号' in df.columns else None  # 仅在 '工单编号' 存在时提取

    # 创建映射表，为每组后四位分配一个随机日期
    if '工单编号' in df.columns:
        unique_last_digits = df['后四位'].unique()
        date_mapping = {digits: random_date() for digits in unique_last_digits}
    else:
        date_mapping = {}

    # 修改工单编号的日期部分
    if '工单编号' in df.columns:
        def modify_order(row):
            last_digits = row['后四位']
            new_date = date_mapping.get(last_digits, random_date())  # 如果没有映射，则生成新日期
            parts = row['工单编号'].split('-')  # 假设格式为 TZ-YYYYMMDD-####
            parts[1] = new_date  # 替换日期部分
            return '-'.join(parts)

        df['工单编号'] = df.apply(modify_order, axis=1)

    # 同时修改指定的列
    for column in columns_to_modify:
        if column in df.columns:
            # 如果 '工单编号' 存在，确保所有这些列的日期与 '工单编号' 相同
            if '工单编号' in df.columns:
                def modify_column_date(row):
                    order_date = row['工单编号'].split('-')[1]  # 从工单编号中提取日期
                    return modify_date_column(row, column, order_date)  # 将 '工单编号' 中的日期应用到指定列

                df[column] = df.apply(modify_column_date, axis=1)
            else:
                # 如果 '工单编号' 不存在，使用随机日期替换这些列的日期
                date_to_apply = random_date()
                # 直接调用 modify_date_column，而不是传递 row
                df[column] = df[column].apply(lambda x: modify_date_column(pd.Series({column: x}), column, date_to_apply))

    # 删除辅助列
    if '后四位' in df.columns:
        df.drop(columns=['后四位'], inplace=True)

    return df


def extract_dates_from_filename(filename):
    """
    从文件名中提取日期范围。
    假设文件名格式为：任意内容YYYY-MM-DD_YYYY-MM-DD.csv
    """
    basename = os.path.basename(filename)
    match = re.search(r"(\d{4}-\d{2}-\d{2})_(\d{4}-\d{2}-\d{2})", basename)
    if match:
        start_date_str, end_date_str = match.groups()
        start_date = datetime.strptime(start_date_str, "%Y-%m-%d")
        end_date = datetime.strptime(end_date_str, "%Y-%m-%d")
        return start_date, end_date
    else:
        raise ValueError(f"文件名中未找到有效的日期范围: {basename}")


def main():
    # 创建 Tk 窗口并隐藏
    root = Tk()
    root.withdraw()

    # 选择 A.csv 文件
    a_file = select_file("选择 A.csv 文件", [("CSV Files", "*.csv")], root)
    if not a_file:
        print("未选择 A.csv 文件，程序退出。")
        # root.quit()  # 关闭 Tk 窗口
        return

    # 弹出输入框让用户输入 N
    n = simpledialog.askinteger("输入", "请输入随机抽取的行数 N：", minvalue=1, parent=root)
    if n is None:
        print("未输入 N，程序退出。")
        # root.quit()  # 关闭 Tk 窗口
        return

    # 从 A.csv 文件中随机抽取 N 行
    df_a = pd.read_csv(a_file, encoding='gbk')
    if len(df_a) < n:
        print(f"A.csv 文件中的行数不足 {n} 行，程序退出。")
        # root.quit()  # 关闭 Tk 窗口
        return

    sampled_df = df_a.sample(n=n)

    # 选择 B.csv 文件
    b_file = select_file("选择 B.csv 文件", [("CSV Files", "*.csv")], root)
    if not b_file:
        print("未选择 B.csv 文件，程序退出。")
        # root.quit()  # 关闭 Tk 窗口
        return

    # 从 B.csv 文件名中提取日期范围
    try:
        start_date, end_date = extract_dates_from_filename(b_file)
    except ValueError as e:
        print(e)
        # root.quit()  # 关闭 Tk 窗口
        return

    # 弹出输入框让用户输入要修改的列名（比如 '创建时间,工单编号,开发日期'）
    columns_input = simpledialog.askstring("输入", "请输入要修改的列名（逗号分隔）：", parent=root)
    if not columns_input:
        print("未输入列名，程序退出。")
        # root.quit()  # 关闭 Tk 窗口
        return

    columns_to_modify = [col.strip() for col in columns_input.split(',')]

    # 修改 '工单编号' 和其他指定列
    sampled_df = modify_work_order_column(sampled_df, start_date, end_date, columns_to_modify)

    # 加载 B.csv 数据并追加
    if os.path.exists(b_file):
        df_b = pd.read_csv(b_file, encoding='gbk')
        df_b = pd.concat([df_b, sampled_df], ignore_index=True)
    else:
        df_b = sampled_df

    # 保存更新后的 B.csv 文件
    df_b.to_csv(b_file, index=False, encoding='gbk')
    print(f"已将 {n} 行数据追加到 B.csv 文件并保存。")

    # 删除 A.csv 中已抽取的行并保存
    df_a = df_a.drop(sampled_df.index)
    df_a.to_csv(a_file, index=False, encoding='gbk')
    print(f"已从 A.csv 中删除 {n} 行数据并保存。")

    # 最后退出 Tk 窗口
    root.destroy()


if __name__ == "__main__":
    main()
