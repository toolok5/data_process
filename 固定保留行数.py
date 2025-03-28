import os
import pandas as pd
import random
from tkinter import Tk, simpledialog
from tkinter.filedialog import askopenfilenames

def get_random_rows(df, n):
    """随机选择n行数据"""
    return df.sample(n=n, random_state=random.randint(1, 10000))

def process_files():
    # 创建Tk根窗口
    root = Tk()
    root.withdraw()  # 隐藏主窗口

    # 弹出对话框输入行数范围
    range_input = simpledialog.askstring("输入范围", 
                                       "请输入行数范围 (格式: 最小值-最大值):",
                                       parent=root)
    
    if not range_input:
        print("未输入范围，程序退出")
        return

    try:
        min_rows, max_rows = map(int, range_input.split('-'))
        if min_rows >= max_rows:
            print("最小值必须小于最大值")
            return
    except ValueError:
        print("输入格式错误，请使用格式：最小值-最大值")
        return

    # 选择文件
    file_paths = askopenfilenames(
        title="选择CSV文件",
        filetypes=[("CSV files", "*.csv")],
        parent=root
    )

    if not file_paths:
        print("未选择文件，程序退出")
        return

    # 处理每个文件
    for file_path in file_paths:
        filename = os.path.basename(file_path)
        print(f"\n处理文件: {filename}")

        try:
            # 读取CSV文件，使用gbk编码
            df = pd.read_csv(file_path, encoding='gbk')
            original_rows = len(df)
            print(f"原始行数: {original_rows}")

            # 检查行数
            if original_rows < max_rows:
                print(f"行数小于{max_rows}，跳过处理")
                continue

            # 随机选择保留行数
            keep_rows = random.randint(min_rows, max_rows)
            print(f"随机保留行数: {keep_rows}")

            # 随机选择行
            df_sampled = get_random_rows(df, keep_rows)

            # 保存回原文件，使用gbk编码
            df_sampled.to_csv(file_path, index=False, encoding='gbk')
            print(f"已保存回原文件: {file_path}")
            print(f"保留行数: {keep_rows}")

        except Exception as e:
            print(f"处理文件 {filename} 时出错: {str(e)}")

    print("\n所有文件处理完成")

def main():
    process_files()


if __name__ == "__main__":
    main()