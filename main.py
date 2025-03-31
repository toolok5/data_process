import tkinter as tk
from tkinter import ttk, messagebox
import threading
import os
import requests
from datetime import datetime
import re
import uuid
import tkinter
import ctypes
# 示例模块导入，替换为你的实际模块
import 地市筛选
import 表格行数统计
import 规划数据处理
import LTE匹配处理
import NR匹配处理
import excel文件生成
import 邻区参数
import 邻区添加
import 普通参数
import 总流量添加
import 删减行数
import MR性能数据采集
import 文件名中字段批量修改
import 复制自定义列数另存
import A数据剪切到B数据
import 每日excel文件生成
import 固定保留行数
import 采集每日数据

# 本地文件路径
file_path = "deadline.txt"  # 使用txt后缀

# 下载文件的URL
url = "https://kdocs.cn/l/ceif9McrDxQ8"  # 替换为实际的文件URL

def show_warning(message):
    """弹出告警对话框"""
    root = tkinter.Tk()
    root.withdraw()  # 隐藏主窗口
    messagebox.showerror("告警", message)  # 显示错误对话框
    root.destroy()

def download_txt_file():
    """下载txt文件并保存，并设置为隐藏状态"""
    try:
        response = requests.get(url, timeout=9)
        response.raise_for_status()  # 确保请求成功
        with open(file_path, "wb") as f:
            f.write(response.content)
        
        # 设置文件为隐藏属性
        FILE_ATTRIBUTE_HIDDEN = 0x02
        ret = ctypes.windll.kernel32.SetFileAttributesW(file_path, FILE_ATTRIBUTE_HIDDEN)
        if ret:
            print("文件下载成功并已设置为隐藏")
        else:
            print("文件下载成功，但设置隐藏属性失败")
    except requests.exceptions.RequestException as e:
        print(f"下载文件时出错：{e}")   # 输出错误信息
        # 下载出错弹出告警
        show_warning("网络有问题，请检查并重新运行软件。")
        exit()  # 退出程序

def check_and_download_file():
    """检查文件是否存在，若不存在或需要更新则下载"""
    if not os.path.exists(file_path):
        # 文件不存在，下载
        print("文件不存在，正在下载...")
        download_txt_file()
    else:
        # 文件存在，检查更新时间
        last_modified_time = os.path.getmtime(file_path)
        file_last_modified = datetime.utcfromtimestamp(last_modified_time)
        # 假设每月检查一次文件更新
        now = datetime.utcnow()
        if (now - file_last_modified).days > 1:  # 例如1天未更新就重新下载
            print("文件已过期，删除并重新下载...")
            try:
                os.remove(file_path)  # 先删除过期文件
                print("已删除过期文件")
            except Exception as e:
                print(f"删除文件时出错：{e}")
            download_txt_file()
        else:
            print("文件已经是最新的，不需要重新下载。")

def get_local_mac_address():
    """获取本机 WLAN 的 MAC 地址"""
    mac = ":".join(re.findall('..', '%012x' % uuid.getnode()))
    print("本机 MAC 地址：", mac)
    return mac

def extract_and_check_authorization():
    """从 txt 文件中提取内容并检查授权状态"""
    try:
        with open(file_path, "r", encoding="utf-8") as file:
            content = file.read()

        # 使用正则表达式提取所有匹配 "font-style=\"normal\">222" 和 "</text>" 之间的内容
        matches = re.findall(r'>2t2(.*?)</text>', content)
        print("匹配结果：", matches)

        if not matches:
            print("未找到有效的授权内容。")
            show_warning("授权失败：未找到有效的授权内容。")
            exit()

        # 遍历所有匹配的内容，进行逻辑检查
        mac_address = get_local_mac_address()
        for match in matches:
            if 'yesyes' in match:
                print("授权成功")
                return True
            elif 'nono' in match:
                print("授权失败")
                show_warning("授权失败")
                os._exit(1)
            elif mac_address in match:
                print("授权成功")
                return True

        # 如果所有匹配中都不满足条件
        print("授权失败")
        show_warning("授权失败")
        os._exit(1)

    except Exception as e:
        print(f"检查授权时出错：{e}")
        show_warning(f"授权时出错")
        exit()

# 主逻辑
check_and_download_file()  # 确保文件是最新的

try:
    extract_and_check_authorization()
except Exception as e:
    show_warning(f"程序运行异常：{e}")

# 按钮功能
def run_task(module):
    """运行指定模块的主函数"""
    try:
        update_progress(0)  # 重置进度条
        module.main()
        update_progress(100)  # 完成进度条更新
    except Exception as e:
        messagebox.showerror("运行错误", f"运行模块时出错：{e}")

def update_progress(value):
    """更新进度条"""
    progress_bar["value"] = value
    window.update_idletasks()

def run_in_thread(func):
    """在单独线程中运行函数"""
    threading.Thread(target=func, daemon=True).start()

def show_instructions(number):
    """显示使用说明"""
    instructions = {
        1: "MR性能数据处理步骤说明：\n---------------------------------注意事项------------------------------\n"
            "1. 首选将MR和性能原始数据分别放在2个文件夹里面。\n"
            "2. 把'去后缀名.bat'也放在文件夹目录下双击运行后自动去除后缀名，然后将文件全部解压。\n"
            "3. 运行'LTE匹配处理'、'NR匹配处理'程序填写行数（每个文件大概需要读取多少行来凑够你需要的总行数），弹出选择框分别选择需要操作的MR文件和性能文件（MR文件和性能文件日期和个数互相对应）。\n"
            "4. 匹配结束后可以查看结果保存在'C:\excel'文件夹里。\n"
            "5. 如果需要可以对csv文件进行选择并删除你需要删除的行数。\n"
            "6. 运行'xlsx文件生成'选择sheet1和sheet2工作表对应保存的列数程序自动对'C:\excel'文件夹里的csv文件进行处理转换成xlsx文件（最好只放要处理的文件）。\n"
            "7. 运行'MR性能数据采集'程序对选择的csv文件根据日期归纳到周期里面按照周名进行合并保存（生成文件后第一周和最后一周的日期可能需要手动更改到实际需要的日期）。\n"
            "8. 运行'A数据剪切到B数据'程序将A文件里的N行剪切到B文件中（默认检查A文件中是否有'工单编号'列并修改工单编号里的日期以对应到B文件名中的日期，还能手动输入列名'创建时间,生产时间'将这些列中的日期也一并修改如果需要的话，B文件名中要包含周日期比如：2024-04-29_2024-05-05）。\n"
            "9. 程序运行可能会有未知的bug，如有问题欢迎指正探讨，微信手机号：15057337780，感谢您的理解和支持",
        2: "规划数据处理步骤说明：\n---------------------------------注意事项------------------------------\n"
            "1. 首先源文件需要各位自己整理每个月要多少条数据和对应的文件，按照月度来处理，比如4月16号到4月28号算在4月份，那就处理对应的文件，日期就填写'2024/04/16-2024/04/28'。\n"
            "2. 如果源文件是按照地市去下载的那就不需要筛选地市了，筛选地市时要注意地市之间的逗号，比如'温州,丽水'里面的逗号一定要英文的','。\n"
            "3. '规划数据处理'是把选择的多个源文件按照日期内的周数和地市平均分到每周文件中，日期格式为'2024/04/16-2024/04/28'。\n"
            "4. 固定全部输出文件结果都在'C:\excel'。\n"
            "5. 运行'采集每日数据'程序，填写每个地市需要保留的行数，程序会自动处理文件夹里的所有文件，并保存到原文件夹，把XX改成MR或者性能。\n"
            "6. 运行'固定保留行数'程序，选择文件夹，填写需要保留的行数，程序会自动处理文件夹里的所有文件，并保存到原文件夹。\n"
            "7. 运行'每日excel文件生成'程序，对应MR和性能文件，程序会自动处理所选文件，并保存到'C:\excel'。\n"
            "8. 程序运行时，请耐心等待，直到进度条提示完成。\n"
            "9. 程序可能存在一些未知bug，如有问题欢迎指正探讨,微信手机：15057337780，感谢您的支持和理解！。\n",
        3: "参数数据处理步骤说明：\n--------------------------------注意事项------------------------------\n"
            "1. 建议按照'邻区参数'→'邻区添加'→'普通参数'→'总流量添加'→'表格行数统计'→'删减行数'顺序按照各自需求运行。\n"
            "2. (邻区添加)一定要先全部选择4到4的csv文件运行完后,再次运行程序选其他文件（表头有bug）。\n"
            "3. 筛选地市时要注意地市之间的逗号，比如'温州,丽水'里面的逗号一定要英文的','。\n"
            "4. 删减行数根据自己需求选择那些要减的源文件和填写每个文件要减多少行。\n"
            "5. 固定全部输出文件结果都在'C:\excel'。\n"
            "6. 程序可能存在一些未知bug，如有问题欢迎指正探讨，微信手机：15057337780，感谢您的支持和理解！\n",
    }
    messagebox.showinfo(f"使用说明{number}", instructions.get(number, "无相关说明"))

# 创建主窗口
window = tk.Tk()
window.title("集中数据处理工具合并版")
window.geometry("700x650")
window.config(bg="#f4f4f4")

# 添加进度条
progress_label = tk.Label(window, text="完成进度提示", font=("Arial", 10), bg="#f4f4f4")
progress_label.pack(pady=5)

progress_bar = ttk.Progressbar(window, orient="horizontal", length=600, mode="determinate")
progress_bar.pack(pady=10)

# 按钮样式
btn_style = {
    "font": ("Arial", 12),
    "width": 18,
    "height": 2,
    "bd": 3,
    "relief": "raised",
    "bg": "#4CAF50",
    "fg": "black",
    "activebackground": "#45a049",
    "activeforeground": "white",
}

# 按钮配置
column_buttons = [
    # 第一列
    [
        ("MR数据说明", lambda: show_instructions(1),{"bg": "#ADD8E6", "activebackground": "#ffff99"}),
        ("LTE匹配处理", lambda: run_in_thread(lambda: run_task(LTE匹配处理))),
        ("NR匹配处理", lambda: run_in_thread(lambda: run_task(NR匹配处理))),
        ("删减行数", lambda: run_in_thread(lambda: run_task(删减行数))),
        ("excel文件生成", lambda: run_in_thread(lambda: run_task(excel文件生成))),
        ("MR性能数据采集", lambda: run_in_thread(lambda: run_task(MR性能数据采集))),
        ("A数据剪切到B数据", lambda: run_in_thread(lambda: run_task(A数据剪切到B数据))),

    ],
    # 第二列
    [
        ("规划数据说明", lambda: show_instructions(2),{"bg": "#ADD8E6", "activebackground": "#ffff99"}),
        ("地市筛选", lambda: run_in_thread(lambda: run_task(地市筛选))),
        ("规划数据处理", lambda: run_in_thread(lambda: run_task(规划数据处理))),
        ("复制自定义列数另存", lambda: run_in_thread(lambda: run_task(复制自定义列数另存))),
        ("文件名中字段批量修改", lambda: run_in_thread(lambda: run_task(文件名中字段批量修改))),
        ("采集每日数据", lambda: run_in_thread(lambda: run_task(采集每日数据))),
        ("固定保留行数", lambda: run_in_thread(lambda: run_task(固定保留行数))),
        ("每日excel文件生成", lambda: run_in_thread(lambda: run_task(每日excel文件生成))),
    ],
    # 第三列
    [
        ("参数数据说明", lambda: show_instructions(3),{"bg": "#ADD8E6", "activebackground": "#ffff99"}),
        ("邻区参数", lambda: run_in_thread(lambda: run_task(邻区参数))),
        ("邻区添加", lambda: run_in_thread(lambda: run_task(邻区添加))),
        ("普通参数", lambda: run_in_thread(lambda: run_task(普通参数))),
        ("总流量添加", lambda: run_in_thread(lambda: run_task(总流量添加))),
        ("表格行数统计", lambda: run_in_thread(lambda: run_task(表格行数统计))),
        ("删减行数", lambda: run_in_thread(lambda: run_task(删减行数))),
    ],
]

# 按钮布局
frame = tk.Frame(window, bg="#f4f4f4")
frame.pack(pady=10)

# 计算最大行数以对齐列
max_rows = max(len(col) for col in column_buttons)

# 创建网格布局
for col_idx, button_list in enumerate(column_buttons):
    for row_idx in range(max_rows):
        if row_idx < len(button_list):
            # 解包按钮配置，支持额外的样式
            text, command, *style_overrides = button_list[row_idx]
            # 合并默认样式和特定样式
            btn_style_updated = {**btn_style, **(style_overrides[0] if style_overrides else {})}
            # 创建按钮
            btn = tk.Button(frame, text=text, command=command, **btn_style_updated)
            btn.grid(row=row_idx, column=col_idx, padx=10, pady=5)
        else:
            # 添加空白标签占位以对齐列
            placeholder = tk.Label(frame, text="", bg="#f4f4f4", width=18, height=2)
            placeholder.grid(row=row_idx, column=col_idx, padx=10, pady=5)

# 运行主窗口
window.mainloop()