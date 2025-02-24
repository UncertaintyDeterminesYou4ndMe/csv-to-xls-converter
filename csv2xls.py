import pandas as pd
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

        self.create_widgets()

    def create_widgets(self):
        # 文件选择框架
        file_frame = ttk.LabelFrame(self.root, text="文件选择", padding=10)
        file_frame.pack(fill="x", padx=10, pady=5)

        self.file_path = tk.StringVar()
        ttk.Entry(file_frame, textvariable=self.file_path, width=60).pack(side="left", padx=5)
        ttk.Button(file_frame, text="选择CSV文件", command=self.select_file).pack(side="left", padx=5)

        # 参数设置框架
        params_frame = ttk.LabelFrame(self.root, text="参数设置", padding=10)
        params_frame.pack(fill="x", padx=10, pady=5)

        # 分隔符
        ttk.Label(params_frame, text="分隔符:").grid(row=0, column=0, padx=5, pady=5)
        self.separator_var = tk.StringVar(value=self.separators[0]["type"])
        ttk.Combobox(params_frame, textvariable=self.separator_var,
                     values=[s["type"] for s in self.separators]).grid(row=0, column=1, padx=5, pady=5)

        # 小数点
        ttk.Label(params_frame, text="小数点:").grid(row=1, column=0, padx=5, pady=5)
        self.decimal_var = tk.StringVar(value=self.decimals[0]["type"])
        ttk.Combobox(params_frame, textvariable=self.decimal_var,
                     values=[d["type"] for d in self.decimals]).grid(row=1, column=1, padx=5, pady=5)

        # 引号字符
        ttk.Label(params_frame, text="引号字符:").grid(row=2, column=0, padx=5, pady=5)
        self.quotechar_var = tk.StringVar(value=self.quotechars[0]["type"])
        ttk.Combobox(params_frame, textvariable=self.quotechar_var,
                     values=[q["type"] for q in self.quotechars]).grid(row=2, column=1, padx=5, pady=5)

        # 表头设置
        ttk.Label(params_frame, text="表头设置:").grid(row=3, column=0, padx=5, pady=5)
        self.header_var = tk.StringVar(value=self.headers[0]["type"])
        ttk.Combobox(params_frame, textvariable=self.header_var,
                     values=[h["type"] for h in self.headers]).grid(row=3, column=1, padx=5, pady=5)

        # 编码设置
        ttk.Label(params_frame, text="文件编码:").grid(row=4, column=0, padx=5, pady=5)
        self.encoding_var = tk.StringVar(value=self.encodings[0]["type"])
        ttk.Combobox(params_frame, textvariable=self.encoding_var,
                     values=[e["type"] for e in self.encodings]).grid(row=4, column=1, padx=5, pady=5)

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
        return params

    def log(self, message):
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)

    def convert(self):
        if not self.file_path.get():
            messagebox.showerror("错误", "请选择CSV文件")
            return

        try:
            params = self.get_parameters()
            self.log(f"开始转换文件: {self.file_path.get()}")

            # 读取CSV文件
            df = pd.read_csv(
                self.file_path.get(),
                sep=params["separator"],
                decimal=params["decimal"],
                quotechar=params["quotechar"],
                header=params["header"],
                encoding=params["encoding"]
            )

            # 设置输出路径
            output_path = os.path.splitext(self.file_path.get())[0] + '.xls'

            # 创建Excel写入器
            writer = pd.ExcelWriter(output_path, engine='xlsxwriter')

            # 获取workbook和worksheet对象
            df.to_excel(writer, index=False)
            workbook = writer.book
            worksheet = writer.sheets['Sheet1']

            # 设置格式
            formats = {
                'number': workbook.add_format({'num_format': '#,##0'}),
                'float': workbook.add_format({'num_format': '#,##0.00'}),
                'text': workbook.add_format({'text_wrap': True}),
            }

            # 设置列宽和格式
            for i, column in enumerate(df.columns):
                # 计算列宽
                max_length = max(
                    df[column].astype(str).apply(len).max(),
                    len(str(column))
                )
                worksheet.set_column(i, i, min(max_length + 2, 100))

                # 设置数据格式
                if df[column].dtype in ['int64', 'Int64']:
                    worksheet.set_column(i, i, None, formats['number'])
                elif df[column].dtype == 'float64':
                    worksheet.set_column(i, i, None, formats['float'])
                else:
                    worksheet.set_column(i, i, None, formats['text'])

            writer.close()

            self.log(f"转换完成，已保存到: {output_path}")
            messagebox.showinfo("成功", "文件转换完成！")

        except Exception as e:
            error_msg = f"转换失败: {str(e)}"
            self.log(error_msg)
            messagebox.showerror("错误", error_msg)


def main():
    root = tk.Tk()
    app = CSVConverterGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
