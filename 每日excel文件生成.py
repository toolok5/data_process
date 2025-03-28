import os
import pandas as pd
from tkinter import Tk
from tkinter.filedialog import askopenfilenames

# 配置文件保存路径
save_folder = "C:\\excel"
os.makedirs(save_folder, exist_ok=True)

def extract_date_from_filename(filename, prefix):
    """
    从文件名中提取日期部分，假设日期前有固定前缀。
    """
    start_idx = filename.find(prefix)
    if start_idx == -1:
        return None
    return filename[start_idx + len(prefix):start_idx + len(prefix) + 8]

def main():
    # 初始化Tkinter对话框
    root = Tk()
    root.withdraw()

    # 选择MR文件
    mr_files = askopenfilenames(title="选择MR文件", filetypes=[("CSV文件", "*.csv")],parent=root)
    if not mr_files:
        print("未选择MR文件！")
        return

    # 选择指标文件
    index_files = askopenfilenames(title="选择指标文件", filetypes=[("CSV文件", "*.csv")])
    if not index_files:
        print("未选择指标文件！")
        return

    # 按日期分组处理
    mr_groups = {}
    index_groups = {}

    # 处理MR文件，按日期分组
    for file in mr_files:
        date = extract_date_from_filename(os.path.basename(file), "MR-")
        if date:
            if date not in mr_groups:
                mr_groups[date] = []
            mr_groups[date].append(file)

    # 处理指标文件，按日期分组
    for file in index_files:
        date = extract_date_from_filename(os.path.basename(file), "指标-")
        if date:
            if date not in index_groups:
                index_groups[date] = []
            index_groups[date].append(file)

    # 遍历所有日期，合并并保存文件
    common_dates = set(mr_groups.keys()) & set(index_groups.keys())
    if not common_dates:
        print("未找到MR文件和指标文件的共同日期！")
        return

    for date in common_dates:
        # 合并MR文件
        mr_data = [pd.read_csv(file, encoding='gbk') for file in mr_groups[date]]
        mr_combined = pd.concat(mr_data, ignore_index=True)

        # 合并指标文件
        index_data = [pd.read_csv(file, encoding='gbk') for file in index_groups[date]]
        index_combined = pd.concat(index_data, ignore_index=True)

        # 保存到Excel
        save_path = os.path.join(save_folder, f"{date}.xlsx")
        with pd.ExcelWriter(save_path, engine='openpyxl') as writer:
            mr_combined.to_excel(writer, sheet_name="MR", index=False)
            index_combined.to_excel(writer, sheet_name="指标", index=False)

        print(f"文件已保存到: {save_path}")

if __name__ == "__main__":
    main()
