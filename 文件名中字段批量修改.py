import os
from tkinter import Tk, filedialog, simpledialog, messagebox


def main():
    # 创建 Tkinter 根窗口
    root = Tk()
    root.withdraw() # 设定窗口的尺寸

    try:
        # 弹出对话框让用户选择多个文件（支持 CSV 和 XLSX）
        file_paths = filedialog.askopenfilenames(
            parent=root,  # 显式传递 root 窗口作为父窗口
            title="选择文件",
            filetypes=[("CSV Files", "*.csv"), ("Excel Files", "*.xlsx"), ("All Files", "*.*")]
        )

        if not file_paths:
            messagebox.showinfo("提示", "未选择任何文件")
            return  # 退出程序，避免继续执行

        try:
            # 弹出对话框输入替换规则（例如 '参数,数据'）
            replace_input = simpledialog.askstring(
                "输入替换规则",
                "请输入要替换的字符串（例如 '参数,数据'）：",
                parent=root  # 显式传递 root 窗口作为父窗口
            )

            if not replace_input:
                raise ValueError("替换规则不能为空")

            # 分割规则：旧字符串与新字符串
            old_str, new_str = replace_input.split(',')
            print(f"替换规则：将 '{old_str}' 替换为 '{new_str}'")

        except Exception as e:
            messagebox.showerror("错误", f"输入替换规则时发生错误: {e}")
            return  # 退出程序，避免继续执行

        # 遍历文件并进行文件名替换
        for file_path in file_paths:
            try:
                # 获取文件的路径和文件名
                dir_path, file_name = os.path.split(file_path)

                # 替换文件名中的字符串
                new_file_name = file_name.replace(old_str, new_str)

                if new_file_name != file_name:
                    # 构造新的完整文件路径
                    new_file_path = os.path.join(dir_path, new_file_name)

                    # 重命名文件
                    os.rename(file_path, new_file_path)
                    print(f"文件名已修改: {file_name} -> {new_file_name}")
                else:
                    print(f"文件名没有变化: {file_name}")

            except Exception as e:
                print(f"处理文件 {file_path} 时出错: {e}")

    except Exception as e:
        messagebox.showerror("错误", f"程序出错: {e}")

    finally:
        root.destroy()  # 确保 Tkinter 窗口正确销毁


# 执行关联程序的主函数
if __name__ == "__main__":
    main()
