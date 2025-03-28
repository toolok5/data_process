import os
import random
import pandas as pd
from tkinter import Tk, filedialog, simpledialog
from datetime import datetime, timedelta


def select_csv_files(root):
    """
    弹出文件选择对话框，选择多个 CSV 文件。
    """
    # root.deiconify()  # 确保窗口显示
    files = filedialog.askopenfilenames(
        title="选择CSV文件",
        filetypes=[("CSV Files", "*.csv")],parent=root
    )
    print("选择的文件:", files)  # 调试输出
    root.withdraw()  # 文件选择框后隐藏主窗口
    return files


def get_user_input(root):
    """
    弹出对话框，输入最小行数和最大行数。
    """
    # root.deiconify()  # 确保窗口显示
    root.update()  # 强制更新窗口

    try:
        # 弹出输入框获取最小行数
        min_rows = simpledialog.askinteger("输入", "请输入最小行数 (min_rows):", minvalue=1,parent=root)
        print("最小行数:", min_rows)  # 调试输出
        if min_rows is None:
            print("最小行数输入无效，程序退出")
            return None, None  # 用户未输入有效的最小行数，返回None

        # 弹出输入框获取最大行数
        max_rows = simpledialog.askinteger("输入", "请输入最大行数 (max_rows):", minvalue=min_rows + 1,parent=root)
        print("最大行数:", max_rows)  # 调试输出
        if max_rows is None:
            print("最大行数输入无效，程序退出")
            return None, None  # 用户未输入有效的最大行数，返回None

    except Exception as e:
        print("发生异常:", e)  # 异常输出
        return None, None

    return min_rows, max_rows


def extract_date_from_filename(filename):
    """
    从文件名中提取日期，假设文件名格式为 xxx_yyy_YYYYMMDD.csv。
    """
    basename = os.path.basename(filename)
    print(f"处理文件: {basename}")
    parts = basename.split('_')
    if len(parts) > 2:
        date_str = parts[-1].split('.')[0]
        try:
            return datetime.strptime(date_str, '%Y%m%d')
        except ValueError:
            print(f"警告: 文件 {basename} 日期格式无效，无法解析")
            return None
    else:
        print(f"警告: 文件 {basename} 中找不到日期部分")
        return None


def group_dates_by_week(dates):
    """
    将日期按周分组，确保每组是从周一到周日。
    """
    dates.sort()
    week_groups = []
    current_week_start = dates[0]

    # 确保第一周的开始日期是周一
    if current_week_start.weekday() != 0:
        current_week_start -= timedelta(days=current_week_start.weekday())  # 调整到周一

    for i, date in enumerate(dates):
        current_week_end = current_week_start + timedelta(days=6)  # 本周结束日期为周日
        if i + 1 < len(dates) and dates[i + 1] > current_week_end:
            week_groups.append((current_week_start, current_week_end))
            current_week_start = dates[i + 1]

    # 添加最后一周
    last_week_end = current_week_start + timedelta(days=6)
    week_groups.append((current_week_start, last_week_end))

    return week_groups


def sample_csv_data(csv_file, min_rows, max_rows):
    """
    从 CSV 文件中按随机行数抽样，行数范围由 min_rows 和 max_rows 指定。
    """
    df = pd.read_csv(csv_file, encoding='gbk')
    total_rows = len(df)

    if total_rows <= min_rows:
        print(f"文件 {csv_file} 行数不足 {min_rows} 行，读取所有数据。")
        return df
    else:
        rows_to_read = random.randint(min_rows, max_rows)
        rows_to_sample = random.sample(range(total_rows), min(rows_to_read, total_rows))
        return df.iloc[rows_to_sample]


def main():
    """
    主程序入口。
    """
    root = Tk()  # 创建 Tk 实例
    root.withdraw()  # 隐藏主窗口，避免在文件选择时显示

    try:
        # 选择 CSV 文件
        csv_files = select_csv_files(root)
        if not csv_files:
            print("未选择任何文件，程序退出。")
            return

        # 获取用户输入的行数范围
        min_rows, max_rows = get_user_input(root)
        if min_rows is None or max_rows is None:
            print("未输入有效的行数，程序退出。")
            return  # 如果用户未输入有效的行数，直接退出

        all_dates = []
        file_date_map = {}

        # 提取文件名中的日期
        for file in csv_files:
            date = extract_date_from_filename(file)
            if date:
                all_dates.append(date)
                if date not in file_date_map:
                    file_date_map[date] = []
                file_date_map[date].append(file)

        if not all_dates:
            print("未能从任何文件提取日期信息，请检查文件名格式。")
            return

        # 按周分组日期
        week_groups = group_dates_by_week(all_dates)

        output_dir = r'C:\excel'
        os.makedirs(output_dir, exist_ok=True)

        # 按周合并数据
        for start_date, end_date in week_groups:
            week_str = f"指标-{start_date.strftime('%Y%m%d')}_{end_date.strftime('%Y%m%d')}"
            output_file = os.path.join(output_dir, f"{week_str}.csv")

            week_data = pd.DataFrame()

            for date in file_date_map:
                if start_date <= date <= end_date:
                    for file in file_date_map[date]:
                        sampled_data = sample_csv_data(file, min_rows, max_rows)
                        week_data = pd.concat([week_data, sampled_data], ignore_index=True)

            # 如果文件已存在，加载并追加数据
            if os.path.exists(output_file):
                existing_data = pd.read_csv(output_file, encoding='gbk')
                week_data = pd.concat([existing_data, week_data], ignore_index=True)

            # 保存合并后的数据到 CSV
            week_data.to_csv(output_file, index=False, encoding='gbk')
            print(f"保存到 {output_file}")

    finally:
        # root.quit()  # 确保 Tk 实例被正确销毁
        root.destroy()  # 销毁窗口


if __name__ == "__main__":
    main()
