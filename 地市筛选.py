import tkinter as tk
from tkinter import filedialog, simpledialog
import pandas as pd

def read_csv_with_encodings(file_path):
    """尝试使用不同编码读取CSV文件"""
    encodings = ['utf-8', 'gbk']  # 尝试两种编码
    for encoding in encodings:
        try:
            df = pd.read_csv(file_path, encoding=encoding)
            return df
        except UnicodeDecodeError:
            continue
    raise Exception(f"无法读取文件 {file_path}，所有尝试的编码都失败。")

def get_filter_condition():
    """弹出输入框让用户输入筛选条件"""
    # 显示输入框，要求用户输入筛选的城市
    input_str = simpledialog.askstring("筛选条件", "请输入要筛选的城市名称，用逗号分隔：")
    if input_str:
        # 将输入的城市字符串按逗号分隔，并去除两端空格
        return [city.strip() for city in input_str.split(',')]
    else:
        return []

def process_csv_files(files):
    """处理CSV文件，进行筛选和保存"""
    for file_path in files:
        try:
            df = read_csv_with_encodings(file_path)

            # 检查文件中是否存在“市”这一列
            if '市' not in df.columns:
                print(f"文件 {file_path} 缺少 '市' 列，跳过该文件。")
                continue

            # 弹出输入框让用户输入筛选条件
            filter_condition = get_filter_condition()
            if not filter_condition:  # 如果没有选择筛选条件，跳过此文件
                print(f"没有输入筛选条件，跳过文件 {file_path}。")
                continue  # 用户没有输入筛选条件时跳过该文件

            # 进行包含匹配筛选
            filtered_df = df[df['市'].apply(lambda x: any(city in str(x) for city in filter_condition))]

            # 输出筛选后的数据（可以选择保存或显示）
            if not filtered_df.empty:
                # 构造保存路径
                save_path = file_path.replace(".csv", "_筛选后.csv")
                filtered_df.to_csv(save_path, index=False, encoding='gbk')
                print(f"筛选后的数据已保存到: {save_path}")
            else:
                print(f"文件 {file_path} 没有符合条件的数据，跳过保存。")

        except Exception as e:
            print(f"处理文件 {file_path} 时出错: {e}")

def main():
    """主函数，统一入口"""
    # 创建并隐藏Tkinter主窗口
    root = tk.Tk()
    root.withdraw()

    # 弹出文件选择框，选择多个CSV文件
    files = filedialog.askopenfilenames(filetypes=[("CSV files", "*.csv")], title="选择CSV文件")
    if not files:
        print("没有选择文件，程序退出。")
        return

    # 处理选择的CSV文件
    process_csv_files(files)

if __name__ == "__main__":
    main()
