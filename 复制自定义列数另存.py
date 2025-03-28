import os
import pandas as pd
from tkinter import Tk, filedialog, simpledialog, messagebox


def main():
    # 创建Tkinter根窗口
    root = Tk()
    root.withdraw() # 隐藏主窗口显示

    try:
        # 弹出对话框让用户选择多个CSV文件
        file_paths = filedialog.askopenfilenames(
            parent=root,  # 确保传递父窗口
            title="选择CSV文件",
            filetypes=[("CSV Files", "*.csv"), ("All Files", "*.*")]
        )

        # 检查是否选择了文件
        if not file_paths:
            messagebox.showinfo("提示", "未选择任何文件")
        else:
            # 输入列数范围
            try:
                # 弹出输入框，获取列范围
                column_range_input = simpledialog.askstring(
                    "输入列范围",
                    "请输入需要读取的列范围（如 1-3, 5-8）：",
                    parent=root  # 确保传递父窗口
                )

                if not column_range_input:
                    raise ValueError("列范围不能为空")

                # 解析列范围
                ranges = column_range_input.split(',')
                cols_to_read = []
                for r in ranges:
                    start_end = r.split('-')
                    if len(start_end) == 1:
                        cols_to_read.append(int(start_end[0]) - 1)  # 单列，减去1以适应0索引
                    elif len(start_end) == 2:
                        start, end = map(int, start_end)
                        cols_to_read.extend(range(start - 1, end))  # 多列，减去1以适应0索引
                    else:
                        raise ValueError("列范围格式错误")

                print(f"列范围输入: {cols_to_read}")

            except Exception as e:
                messagebox.showerror("错误", f"输入列范围时发生错误: {e}")
                root.quit()  # 确保退出程序
                return

            # 遍历文件
            for file_path in file_paths:
                # 选择编码方式：尝试使用utf-8和gbk
                for encoding in ["gbk", "utf-8"]:
                    try:
                        # 尝试读取CSV文件
                        df = pd.read_csv(file_path, usecols=cols_to_read, encoding=encoding)
                        print(f"文件 {file_path} 使用 {encoding} 编码成功读取")
                        break
                    except Exception as e:
                        print(f"文件 {file_path} 使用 {encoding} 编码读取失败: {e}")
                else:
                    # 如果都失败了，显示错误提示并跳过该文件
                    messagebox.showerror("错误", f"文件 {file_path} 无法读取，跳过")
                    continue

                # 构造新的文件名
                base_name = os.path.basename(file_path)  # 原文件名
                new_name = f"规划数据采集_{base_name}"
                new_path = os.path.join(os.path.dirname(file_path), new_name)

                # 保存新的CSV文件
                try:
                    df.to_csv(new_path, index=False, encoding="gbk")
                    print(f"处理完成: {new_path}")
                except Exception as e:
                    print(f"保存文件 {new_path} 时出错: {e}")

    except Exception as e:
        messagebox.showerror("错误", f"程序出错: {e}")

    finally:
        root.destroy()  # 销毁窗口


# 执行关联程序的主函数
if __name__ == "__main__":
    main()
