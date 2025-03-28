import tkinter as tk
from tkinter import filedialog, simpledialog
import pandas as pd
import random

def try_read_file(file_path, encodings=['gbk','utf-8']):
    """
    尝试读取文件（支持CSV和Excel格式）。

    参数:
    file_path (str): 文件的路径。
    encodings (list of str): 要尝试的编码列表（仅对CSV有效）。

    返回:
    pd.DataFrame: 如果成功读取文件，则返回DataFrame。
    """
    # 检查文件扩展名
    if file_path.endswith('.csv'):
        for encoding in encodings:
            try:
                df = pd.read_csv(file_path, encoding=encoding)
                return df  # 如果成功读取，则返回DataFrame并退出循环
            except (UnicodeDecodeError, pd.errors.ParserError) as e:
                print(f"尝试使用 {encoding} 编码读取 {file_path} 失败: {e}")
        raise Exception(f"无法读取文件 {file_path}，所有尝试的编码都失败。")
    elif file_path.endswith('.xlsx'):
        try:
            df = pd.read_excel(file_path)
            return df  # 如果成功读取Excel文件
        except Exception as e:
            raise Exception(f"无法读取文件 {file_path}: {e}")
    else:
        raise Exception(f"不支持的文件类型: {file_path}")

def save_file(df, file_path):
    """
    根据文件扩展名将DataFrame保存到文件。

    参数:
    df (pd.DataFrame): 要保存的DataFrame。
    file_path (str): 文件的路径。
    """
    if file_path.endswith('.csv'):
        df.to_csv(file_path, index=False, encoding='gbk')  # CSV保存为GBK编码
    elif file_path.endswith('.xlsx'):
        df.to_excel(file_path, index=False, engine='openpyxl')  # Excel保存
    else:
        raise Exception(f"不支持的文件类型: {file_path}")

def remove_random_rows(file_path, num_rows_to_remove):
    """
    随机删除文件中的指定行数。

    参数:
    file_path (str): 文件路径。
    num_rows_to_remove (int): 要删除的行数。
    """
    # 尝试读取文件
    df = try_read_file(file_path)

    # 检查文件是否有足够的行数可以删除
    if len(df) <= num_rows_to_remove:
        print(f"文件 {file_path} 的行数少于 {num_rows_to_remove}，无法删除指定数量的行。")
        return

    # 随机选择要删除的行索引
    rows_to_drop = random.sample(range(len(df)), num_rows_to_remove)

    # 删除行
    df_dropped = df.drop(rows_to_drop).reset_index(drop=True)

    # 保存文件
    save_file(df_dropped, file_path)

def main():
    # 创建Tkinter主窗口
    root = tk.Tk()
    root.withdraw()  # 隐藏主窗口

    # 弹出文件选择对话框
    files = filedialog.askopenfilenames(filetypes=[("CSV or Excel files", "*.csv;*.xlsx")])
    if not files:
        print("未选择文件，程序将退出。")
    
        return

    # 弹出输入框填一个数字
    try:
        num_rows_to_remove = simpledialog.askinteger("输入行数", "请输入要随机删除的行数：", parent=root)
        if num_rows_to_remove is None:
            print("取消输入，程序将退出。")
        
            return
    except ValueError:
        print("输入的不是有效的整数，程序将退出。")
        
        return

    # 对每个文件根据数字随机删除指定行数
    for file_path in files:
        remove_random_rows(file_path, num_rows_to_remove)

    print("处理完成！")
    root.destroy()

if __name__ == "__main__":
    main()
