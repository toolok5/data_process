import os
import pandas as pd
import random
from tkinter import Tk, Toplevel, Label, Entry, Button
from tkinter.filedialog import askopenfilenames

# 提取文件名中的日期
def extract_date_from_filename(filename):
    import re
    match = re.search(r'\d{8}', filename)
    return match.group(0) if match else None

# 提取地市名称
def extract_city_from_filename(filename):
    import re
    match = re.search(r'(\w+)_.*?(指标数据|性能指标)', filename)
    return match.group(1) if match else "未知地市"

# 弹出对话框填写地市和行数
def configure_city_row_counts(cities, root):
    config_window = Toplevel(root)
    config_window.title("设置每个地市的读取行数")
    config_window.grab_set()
    
    # 将窗口居中显示
    config_window.update_idletasks()
    width = config_window.winfo_width()
    height = config_window.winfo_height()
    x = (config_window.winfo_screenwidth() // 2) - (width // 2)
    y = (config_window.winfo_screenheight() // 2) - (height // 2)
    config_window.geometry(f'+{x}+{y}')

    entries = {}
    for i, city in enumerate(cities):
        Label(config_window, text=city).grid(row=i, column=0, padx=5, pady=5)
        entry = Entry(config_window)
        entry.grid(row=i, column=1, padx=5, pady=5)
        entries[city] = entry

    def save_and_close():
        for city, entry in entries.items():
            try:
                city_row_counts[city] = int(entry.get()) if entry.get() else None
            except ValueError:
                print(f"地市 {city} 的行数无效，将读取所有行。")
                city_row_counts[city] = None
        config_window.destroy()
        root.quit()

    Button(config_window, text="确认", command=save_and_close).grid(row=len(cities), column=0, columnspan=2, pady=10)

# 主函数
def main():
    # 创建主窗口
    root = Tk()
    root.withdraw()  # 隐藏主窗口
    
    print("正在打开文件选择对话框...")
    file_paths = askopenfilenames(title="选择多个MR或者性能 CSV 文件", filetypes=[("CSV files", "*.csv")],parent=root)

    if not file_paths:
        print("未选择任何文件。")
        root.destroy()
        return

    print(f"已选择 {len(file_paths)} 个文件")

    # 提取所有文件中的地市名称
    file_names = [os.path.basename(path) for path in file_paths]
    cities = list(set(extract_city_from_filename(name) for name in file_names))

    print(f"检测到以下地市：{cities}")

    global city_row_counts
    city_row_counts = {city: None for city in cities}

    # 弹出对话框配置行数
    configure_city_row_counts(cities, root)
    root.mainloop()

    # 如果用户没有设置任何行数，直接返回
    if all(count is None for count in city_row_counts.values()):
        print("未设置任何行数，程序退出。")
        root.destroy()
        return

    print("开始处理文件...")

    # 输出目录
    output_dir = r'C:\excel'
    os.makedirs(output_dir, exist_ok=True)

    for file_path in file_paths:
        filename = os.path.basename(file_path)
        city = extract_city_from_filename(filename)
        base_row_count = city_row_counts.get(city)
        
        # 在这里为每个文件生成随机行数
        if base_row_count is not None:
            row_count = base_row_count + random.randint(-300, 300)
        else:
            row_count = None

        # 读取文件
        try:
            df = pd.read_csv(file_path, encoding='gbk', nrows=row_count)
            print(f"文件 {filename} 实际读取行数：{row_count}")  # 添加调试输出
        except Exception as e:
            print(f"读取文件 {file_path} 时出错：{e}")
            continue

        # 提取日期
        file_date = extract_date_from_filename(filename)

        if not file_date:
            print(f"文件 {filename} 中未找到日期，跳过处理。")
            continue

        # 保存路径
        output_file = os.path.join(output_dir, f"XX-{file_date}.csv")

        # 检查是否已有文件存在，若存在则加载并删除重复内容
        if os.path.exists(output_file):
            try:
                existing_df = pd.read_csv(output_file, encoding='gbk')
                df = pd.concat([existing_df, df], ignore_index=True).drop_duplicates()
            except Exception as e:
                print(f"读取或合并文件 {output_file} 时出错：{e}")
                continue

        # 保存数据
        try:
            df.to_csv(output_file, encoding='gbk', index=False)
            print(f"文件 {filename} 的数据已保存到 {output_file}，读取了 {row_count if row_count else '所有'} 行，已去除重复内容。")
        except Exception as e:
            print(f"保存文件 {output_file} 时出错：{e}")

if __name__ == "__main__":
    main()
