import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os


class CSVConverterGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("CSV 转 Excel 工具")
        self.root.geometry("800x600")

        # 定义转换参数选项
        self.separators = [
            {"type": "逗号 (,)", "sym": ","},
            {"type": "分号 (;)", "sym": ";"},
            {"type": "竖线 (|)", "sym": "|"},
            {"type": "制表符 (Tab)", "sym": "\t"},
        ]

        self.decimals = [
            {"type": "点号 (.)", "sym": "."},
            {"type": "逗号 (,)", "sym": ","},
        ]

        self.quotechars = [
            {"type": "单引号 (')", "sym": "'"},
            {"type": "双引号 (\")", "sym": '"'},
            {"type": "无", "sym": ""},
        ]

        self.headers = [
            {"type": "第一行", "sym": 0},
            {"type": "无表头", "sym": None},
        ]

        self.encodings = [
            {"type": "UTF-8", "sym": "utf-8"},
            {"type": "UTF-16", "sym": "utf-16"},
            {"type": "Latin-1", "sym": "latin_1"},
            {"type": "ASCII", "sym": "ascii"},
        ]

        # 添加格式处理选项
        self.format_options = [
            {"type": "自动 (自动检测数据类型)", "sym": "auto"},
            {"type": "文本 (所有列作为文本处理)", "sym": "text"},
        ]

        self.create_widgets()

    def create_widgets(self):
        # 文件选择框架
        file_frame = ttk.LabelFrame(self.root, text="文件选择", padding=10)
        file_frame.pack(fill="x", padx=10, pady=5)

        self.file_path = tk.StringVar()
        ttk.Entry(file_frame, textvariable=self.file_path, width=60).pack(side="left", padx=5)
        ttk.Button(file_frame, text="选择CSV文件", command=self.select_file).pack(side="left", padx=5)

        # 参数设置框架 - 保存为实例变量
        self.params_frame = ttk.LabelFrame(self.root, text="参数设置", padding=10)
        self.params_frame.pack(fill="x", padx=10, pady=5)

        # 分隔符
        ttk.Label(self.params_frame, text="分隔符:").grid(row=0, column=0, padx=5, pady=5)
        self.separator_var = tk.StringVar(value=self.separators[0]["type"])
        ttk.Combobox(self.params_frame, textvariable=self.separator_var,
                     values=[s["type"] for s in self.separators]).grid(row=0, column=1, padx=5, pady=5)

        # 小数点
        ttk.Label(self.params_frame, text="小数点:").grid(row=1, column=0, padx=5, pady=5)
        self.decimal_var = tk.StringVar(value=self.decimals[0]["type"])
        ttk.Combobox(self.params_frame, textvariable=self.decimal_var,
                     values=[d["type"] for d in self.decimals]).grid(row=1, column=1, padx=5, pady=5)

        # 引号字符
        ttk.Label(self.params_frame, text="引号字符:").grid(row=2, column=0, padx=5, pady=5)
        self.quotechar_var = tk.StringVar(value=self.quotechars[0]["type"])
        ttk.Combobox(self.params_frame, textvariable=self.quotechar_var,
                     values=[q["type"] for q in self.quotechars]).grid(row=2, column=1, padx=5, pady=5)

        # 表头设置
        ttk.Label(self.params_frame, text="表头设置:").grid(row=3, column=0, padx=5, pady=5)
        self.header_var = tk.StringVar(value=self.headers[0]["type"])
        ttk.Combobox(self.params_frame, textvariable=self.header_var,
                     values=[h["type"] for h in self.headers]).grid(row=3, column=1, padx=5, pady=5)

        # 编码设置
        ttk.Label(self.params_frame, text="文件编码:").grid(row=4, column=0, padx=5, pady=5)
        self.encoding_var = tk.StringVar(value=self.encodings[0]["type"])
        ttk.Combobox(self.params_frame, textvariable=self.encoding_var,
                     values=[e["type"] for e in self.encodings]).grid(row=4, column=1, padx=5, pady=5)

        # 数据格式
        ttk.Label(self.params_frame, text="数据格式:").grid(row=5, column=0, padx=5, pady=5)
        self.format_var = tk.StringVar(value=self.format_options[0]["type"])
        ttk.Combobox(self.params_frame, textvariable=self.format_var,
                     values=[f["type"] for f in self.format_options]).grid(row=5, column=1, padx=5, pady=5)

        # 添加高级选项框架(可折叠)
        self.show_advanced = tk.BooleanVar(value=False)
        ttk.Checkbutton(self.params_frame, text="显示高级选项", variable=self.show_advanced,
                        command=self.toggle_advanced_options).grid(row=6, column=0, columnspan=2, pady=5)

        self.advanced_frame = ttk.LabelFrame(self.root, text="高级选项", padding=10)

        # 列格式预览按钮
        ttk.Button(self.advanced_frame, text="预览并设置列格式",
                   command=self.preview_columns).pack(pady=5)

        # 转换按钮
        ttk.Button(self.root, text="转换", command=self.convert).pack(pady=10)

        # 日志框
        log_frame = ttk.LabelFrame(self.root, text="转换日志", padding=10)
        log_frame.pack(fill="both", expand=True, padx=10, pady=5)

        self.log_text = tk.Text(log_frame, height=10)
        self.log_text.pack(fill="both", expand=True)

    def select_file(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if file_path:
            self.file_path.set(file_path)

    def get_parameters(self):
        params = {}
        params["separator"] = next(s["sym"] for s in self.separators if s["type"] == self.separator_var.get())
        params["decimal"] = next(d["sym"] for d in self.decimals if d["type"] == self.decimal_var.get())
        params["quotechar"] = next(q["sym"] for q in self.quotechars if q["type"] == self.quotechar_var.get())
        params["header"] = next(h["sym"] for h in self.headers if h["type"] == self.header_var.get())
        params["encoding"] = next(e["sym"] for e in self.encodings if e["type"] == self.encoding_var.get())
        params["format_type"] = next(f["sym"] for f in self.format_options if f["type"] == self.format_var.get())
        return params

    def log(self, message):
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)

    def toggle_advanced_options(self):
        if self.show_advanced.get():
            self.advanced_frame.pack(fill="x", padx=10, pady=5, after=self.params_frame)
        else:
            self.advanced_frame.pack_forget()

    def preview_columns(self):
        """预览CSV文件的列并允许设置每列的格式"""
        if not self.file_path.get():
            messagebox.showerror("错误", "请先选择CSV文件")
            return

        try:
            params = self.get_parameters()

            # 使用 csv 模块读取数据
            import csv
            preview_data = []
            headers = None

            with open(self.file_path.get(), 'r', encoding=params["encoding"]) as f:
                reader = csv.reader(f,
                                    delimiter=params["separator"],
                                    quotechar=params["quotechar"] if params["quotechar"] else '"')

                # 读取表头
                if params["header"] == 0:
                    headers = next(reader)

                # 读取前10行数据
                for _ in range(10):
                    try:
                        row = next(reader)
                        preview_data.append(row)
                    except StopIteration:
                        break

                # 如果没有表头，生成默认表头
                if not headers:
                    headers = [f"Column_{i + 1}" for i in range(len(preview_data[0]))]

            # 创建预览窗口
            preview_window = tk.Toplevel(self.root)
            preview_window.title("列格式设置")
            preview_window.geometry("800x600")

            # 创建滚动条和画布
            canvas = tk.Canvas(preview_window)
            scrollbar = ttk.Scrollbar(preview_window, orient="vertical", command=canvas.yview)
            scrollable_frame = ttk.Frame(canvas)

            # 配置画布
            canvas.configure(yscrollcommand=scrollbar.set)
            scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
            canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")

            # 添加说明标签
            ttk.Label(scrollable_frame, text="此功能允许为每列设置特定格式",
                      wraplength=700).pack(pady=10)

            # 存储每列的格式选择
            self.column_formats = {}

            # 列出所有列及其示例数据
            for i, col in enumerate(headers):
                frame = ttk.LabelFrame(scrollable_frame, text=f"列 {i + 1}: {col}")
                frame.pack(fill="x", padx=5, pady=5)

                # 显示数据样例
                sample_data = [row[i] for row in preview_data[:3]]
                sample_text = "样例数据:\n" + "\n".join(str(x) for x in sample_data)
                ttk.Label(frame, text=sample_text, wraplength=600).pack(anchor="w", padx=5, pady=5)

                # 格式选择
                format_var = tk.StringVar(value="auto")
                self.column_formats[col] = format_var

                formats_frame = ttk.Frame(frame)
                formats_frame.pack(fill="x", padx=5, pady=5)

                ttk.Radiobutton(formats_frame, text="自动检测",
                                variable=format_var, value="auto").pack(side="left", padx=5)
                ttk.Radiobutton(formats_frame, text="文本格式",
                                variable=format_var, value="text").pack(side="left", padx=5)
                ttk.Radiobutton(formats_frame, text="数值格式",
                                variable=format_var, value="number").pack(side="left", padx=5)
                ttk.Radiobutton(formats_frame, text="日期格式",
                                variable=format_var, value="date").pack(side="left", padx=5)

            # 确认按钮
            def save_formats():
                self.saved_column_formats = {col: var.get() for col, var in self.column_formats.items()}
                preview_window.destroy()
                self.log("已保存列格式设置")

            ttk.Button(scrollable_frame, text="保存格式设置",
                       command=save_formats).pack(pady=10)

            # 布局滚动条和画布
            scrollbar.pack(side="right", fill="y")
            canvas.pack(side="left", fill="both", expand=True)

        except Exception as e:
            import traceback
            error_msg = f"预览失败: {str(e)}\n{traceback.format_exc()}"
            self.log(error_msg)
            messagebox.showerror("错误", error_msg)

    def convert(self):
        if not self.file_path.get():
            messagebox.showerror("错误", "请选择CSV文件")
            return

        try:
            params = self.get_parameters()
            self.log(f"开始转换文件: {self.file_path.get()}")

            # 读取CSV文件
            import csv
            all_data = []
            headers = None

            with open(self.file_path.get(), 'r', encoding=params["encoding"]) as f:
                reader = csv.reader(f,
                                    delimiter=params["separator"],
                                    quotechar=params["quotechar"] if params["quotechar"] else '"')

                # 读取表头
                if params["header"] == 0:
                    headers = next(reader)

                # 读取所有数据
                all_data = list(reader)

                # 如果没有表头，生成默认表头
                if not headers:
                    headers = [f"Column_{i + 1}" for i in range(len(all_data[0]))]

            self.log(f"读取完成，行数: {len(all_data)}, 列数: {len(headers)}")
            self.log(f"列名: {headers}")

            # 设置输出路径
            output_path = os.path.splitext(self.file_path.get())[0] + '.xlsx'

            # 创建Excel写入器
            import xlsxwriter
            workbook = xlsxwriter.Workbook(output_path)
            worksheet = workbook.add_worksheet('Sheet1')

            # 定义格式
            formats = {
                'number': workbook.add_format({
                    'num_format': '#,##0',
                    'align': 'right'
                }),
                'float': workbook.add_format({
                    'num_format': '#,##0.00',
                    'align': 'right'
                }),
                'text': workbook.add_format({
                    'num_format': '@',
                    'align': 'left'
                }),
                'date': workbook.add_format({
                    'num_format': 'yyyy-mm-dd',
                    'align': 'center'
                }),
                'header': workbook.add_format({
                    'bold': True,
                    'text_wrap': True,
                    'align': 'center',
                    'valign': 'vcenter',
                    'bg_color': '#F0F0F0'
                })
            }

            # 写入表头
            for col_num, value in enumerate(headers):
                worksheet.write(0, col_num, value, formats['header'])

            # 检查是否有保存的列格式设置
            has_column_formats = hasattr(self, 'saved_column_formats') and self.saved_column_formats

            # 处理每列数据
            for col_num, col_name in enumerate(headers):
                # 计算列宽
                max_length = max(
                    max(len(str(row[col_num])) for row in all_data),
                    len(str(col_name))
                )
                worksheet.set_column(col_num, col_num, min(max_length + 2, 100))

                # 获取该列的格式设置
                column_format = (self.saved_column_formats.get(col_name, 'auto')
                                 if has_column_formats else 'auto')

                self.log(f"处理列 '{col_name}' (格式: {column_format})")

                # 写入数据
                for row_num, row in enumerate(all_data, start=1):
                    value = row[col_num]
                    if not value:  # 跳过空值
                        continue

                    try:
                        if column_format == 'text' or (column_format == 'auto' and ':' in str(value)):
                            worksheet.write_string(row_num, col_num, str(value), formats['text'])
                        elif column_format == 'number':
                            if '.' in str(value):
                                worksheet.write_number(row_num, col_num, float(value), formats['float'])
                            else:
                                worksheet.write_number(row_num, col_num, int(value), formats['number'])
                        elif column_format == 'date':
                            try:
                                import datetime
                                date_value = datetime.datetime.strptime(value, '%Y-%m-%d')
                                worksheet.write_datetime(row_num, col_num, date_value, formats['date'])
                            except:
                                worksheet.write_string(row_num, col_num, str(value), formats['text'])
                        else:
                            worksheet.write_string(row_num, col_num, str(value), formats['text'])

                    except Exception as e:
                        self.log(f"警告: 处理单元格 ({row_num}, {col_num}) 时出错: {str(e)}")
                        worksheet.write_string(row_num, col_num, str(value), formats['text'])

            workbook.close()
            self.log(f"转换完成，已保存到: {output_path}")
            messagebox.showinfo("成功", "文件转换完成！")

        except Exception as e:
            import traceback
            error_msg = f"转换失败: {str(e)}\n{traceback.format_exc()}"
            self.log(error_msg)
            messagebox.showerror("错误", error_msg)


def main():
    root = tk.Tk()
    app = CSVConverterGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
