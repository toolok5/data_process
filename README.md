这是一个将csv原始数据进行多种特定操作程序，比如将同一天的MR文件和性能文件通过列'小区码CI'进行匹配合并，增加某些列并通过前面几列数据数据内容的比较返回特定的值或者内容。
以下是一些具体的操作详情：

       1: "MR数据处理步骤说明：\n---------------------------------注意事项------------------------------\n"
            "1. 首选将MR和性能原始数据分别放在2个文件夹里面。\n"
            "2. 把“去后缀名.bat”也放在文件夹目录下双击运行后自动去除后缀名，然后将文件全部解压。\n"
            "3. 运行“csv匹配处理”程序弹出选择框分别选择需要操作的MR文件和性能文件（MR文件和性能文件日期和个数互相对应）。\n"
            "4. 匹配结束后可以查看结果保存在“C:\excel”文件夹里。\n"
            "5. 如果需要可以对csv文件进行选择删除你需要删除的行数。\n"
            "6. 运行“xlsx文件生成”程序自动对“C:\excel”文件夹里的csv文件进行处理（100M一个文件大概需要2分钟处理）。\n"
            "7. 程序运行可能会有未知的bug，如有问题欢迎指正探讨，微信手机号：15057337780，感谢您的理解和支持",
        2: "规划数据处理步骤说明：\n---------------------------------注意事项------------------------------\n"
            "1. 首先源文件需要各位自己整理每个月要多少条数据和对应的文件，按照月度来处理，比如4月16号到4月28号算在4月份，那就处理对应的文件，日期就填写’2024/04/16-2024/04/28’。\n"
            "2. 如果源文件是按照地市去下载的那就不需要筛选地市了。\n"
            "3. 筛选地市时要注意地市之间的逗号，比如‘温州,丽水’里面的逗号一定要英文的‘,’。\n"
            "4. 固定全部输出文件结果都在‘C:\excel’。\n"
            "5. 程序运行时，请耐心等待，直到进度条提示完成。\n"
            "6. 程序可能存在一些未知bug，如有问题欢迎指正探讨，微信手机：15057337780，感谢您的支持和理解！。\n",
        3: "参数数据处理步骤说明：\n--------------------------------注意事项------------------------------\n"
            "1. 建议按照‘邻区参数’→‘邻区添加’→’普通参数’→‘总流量添加’→‘表格行数统计’→‘删减行数’顺序按照各自需求运行。\n"
            "2. ‘邻区添加’一定要先全部选择4到4的csv文件运行完后，再次运行程序选其他文件（表头有bug）。\n"
            "3. 筛选地市时要注意地市之间的逗号，比如‘温州,丽水’里面的逗号一定要英文的‘,’。\n"
            "4. 删减行数根据自己需求选择那些要减的源文件和填写每个文件要减多少行。\n"
            "5. 固定全部输出文件结果都在‘C:\excel’。\n"
            "6. 程序可能存在一些未知bug，如有问题欢迎指正探讨，微信手机：15057337780，感谢您的支持和理解！\n"