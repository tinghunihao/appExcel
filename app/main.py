import queue
import threading
import traceback
from pathlib import Path
from tkinter import BOTH, LEFT, RIGHT, X, filedialog, messagebox
from tkinter import Tk, StringVar
from tkinter import ttk

from logic import (
    MatchColumns,
    complete_address_workbook,
    complete_anhui_address,
    extract_and_merge_workbook,
    match_order_workbooks,
)


class ExcelHelperApp(Tk):
    def __init__(self):
        super().__init__()
        self.title("门店 Excel 智能处理工具")
        self.geometry("860x620")
        self.minsize(760, 540)
        self.jobs = queue.Queue()
        self._build()
        self.after(100, self._poll_jobs)

    def _build(self):
        root = ttk.Frame(self, padding=12)
        root.pack(fill=BOTH, expand=True)

        title = ttk.Label(root, text="门店 Excel 智能处理工具", font=("Microsoft YaHei UI", 16, "bold"))
        title.pack(anchor="w")

        tabs = ttk.Notebook(root)
        tabs.pack(fill=BOTH, expand=True, pady=(12, 8))

        self.address_tab = ttk.Frame(tabs, padding=12)
        self.extract_tab = ttk.Frame(tabs, padding=12)
        self.match_tab = ttk.Frame(tabs, padding=12)
        tabs.add(self.address_tab, text="地址行政区划补全")
        tabs.add(self.extract_tab, text="提取采购单所需信息")
        tabs.add(self.match_tab, text="A/B 表发货单号匹配")

        self._build_address_tab()
        self._build_extract_tab()
        self._build_match_tab()

        self.log_text = ttk.Label(root, text="就绪", foreground="#555")
        self.log_text.pack(anchor="w")

    def _build_address_tab(self):
        preview = ttk.LabelFrame(self.address_tab, text="单条地址测试", padding=10)
        preview.pack(fill=X)
        self.address_input = StringVar(value="阜南县朱寨镇刘锋手机大卖场")
        self.address_output = StringVar()
        row = ttk.Frame(preview)
        row.pack(fill=X)
        ttk.Entry(row, textvariable=self.address_input).pack(side=LEFT, fill=X, expand=True)
        ttk.Button(row, text="测试补全", command=self._preview_address).pack(side=RIGHT, padx=(8, 0))
        ttk.Label(preview, textvariable=self.address_output, foreground="#0F766E").pack(anchor="w", pady=(8, 0))

        batch = ttk.LabelFrame(self.address_tab, text="批量处理 Excel", padding=10)
        batch.pack(fill=X, pady=(14, 0))
        self.addr_input_file = StringVar()
        self.addr_output_file = StringVar()
        self.addr_column = StringVar(value="门店")
        self._file_row(batch, "输入 Excel", self.addr_input_file, self._pick_addr_input)
        self._entry_row(batch, "地址列名", self.addr_column)
        self._file_row(batch, "输出 Excel", self.addr_output_file, self._pick_addr_output, save=True)
        ttk.Button(batch, text="开始批量补全", command=self._run_address_batch).pack(anchor="e", pady=(10, 0))

    def _build_extract_tab(self):
        files = ttk.LabelFrame(self.extract_tab, text="选择文件", padding=10)
        files.pack(fill=X)
        self.raw_file = StringVar()
        self.master_file = StringVar()
        self.extract_output_file = StringVar()
        self._file_row(files, "原始表", self.raw_file, self._pick_raw_file)
        self._file_row(files, "机型主数据表", self.master_file, self._pick_master_file)
        self._file_row(files, "输出 Excel", self.extract_output_file, self._pick_extract_output, save=True)

        settings = ttk.LabelFrame(self.extract_tab, text="提取与匹配设置", padding=10)
        settings.pack(fill=X, pady=(14, 0))
        self.extract_columns = StringVar(value="客户姓名,手机号,产品名称/规格,69国标码,客户实付价,补贴金额,s/n码,Imei-1号,Imei-2号,核销门店,还货地址,联系人,电话")
        self.rename_columns = StringVar(value="")
        self.raw_code_header = StringVar(value="69国标码")
        self.master_code_header = StringVar(value="69码")
        self._entry_row(settings, "提取列", self.extract_columns)
        self._entry_row(settings, "列名修改", self.rename_columns)
        self._entry_row(settings, "原始表69码列", self.raw_code_header)
        self._entry_row(settings, "主数据69码列", self.master_code_header)

        note = (
            "提取列可填列序号或列名，例如：1,3,5,8,10,69国标码；"
            "列名修改格式：1=订单号,69国标码=商品条码"
        )
        ttk.Label(settings, text=note, foreground="#555", wraplength=780).pack(anchor="w", pady=(8, 0))
        ttk.Button(settings, text="开始提取并匹配属性", command=self._run_extract_merge).pack(anchor="e", pady=(10, 0))

    def _build_match_tab(self):
        files = ttk.LabelFrame(self.match_tab, text="选择文件", padding=10)
        files.pack(fill=X)
        self.a_file = StringVar()
        self.b_file = StringVar()
        self.match_output_file = StringVar()
        self._file_row(files, "A表：商户订单表", self.a_file, self._pick_a_file)
        self._file_row(files, "B表：厂家发货表", self.b_file, self._pick_b_file)
        self._file_row(files, "输出 Excel", self.match_output_file, self._pick_match_output, save=True)

        columns = ttk.LabelFrame(self.match_tab, text="列名设置", padding=10)
        columns.pack(fill=X, pady=(14, 0))
        self.a_phone = StringVar(value="电话")
        self.a_name = StringVar(value="联系人")
        self.a_model = StringVar(value="厂家机型")
        self.a_fuzzy_model = StringVar(value="系统机型")
        self.b_phone = StringVar(value="联系方式")
        self.b_name = StringVar(value="收货人")
        self.b_model = StringVar(value="机型版本系列（填写产品匹配里面的格式）")
        self.b_ship_no = StringVar(value="发货通知单号")
        grid = ttk.Frame(columns)
        grid.pack(fill=X)
        self._grid_entry(grid, 0, 0, "A表电话列", self.a_phone)
        self._grid_entry(grid, 0, 1, "A表姓名列", self.a_name)
        self._grid_entry(grid, 0, 2, "A表厂家机型列", self.a_model)
        self._grid_entry(grid, 1, 0, "A表模糊机型列", self.a_fuzzy_model)
        self._grid_entry(grid, 1, 1, "B表电话列", self.b_phone)
        self._grid_entry(grid, 1, 2, "B表姓名列", self.b_name)
        self._grid_entry(grid, 2, 0, "B表机型列", self.b_model)
        self._grid_entry(grid, 2, 1, "B表发货单号列", self.b_ship_no)

        ttk.Button(self.match_tab, text="开始匹配并生成结果", command=self._run_match).pack(anchor="e", pady=(16, 0))

    def _file_row(self, parent, label, var, command, save=False):
        row = ttk.Frame(parent)
        row.pack(fill=X, pady=4)
        ttk.Label(row, text=label, width=16).pack(side=LEFT)
        ttk.Entry(row, textvariable=var).pack(side=LEFT, fill=X, expand=True)
        ttk.Button(row, text="选择" if not save else "保存为", command=command).pack(side=RIGHT, padx=(8, 0))

    def _entry_row(self, parent, label, var):
        row = ttk.Frame(parent)
        row.pack(fill=X, pady=4)
        ttk.Label(row, text=label, width=16).pack(side=LEFT)
        ttk.Entry(row, textvariable=var).pack(side=LEFT, fill=X, expand=True)

    def _grid_entry(self, parent, row, col, label, var):
        frame = ttk.Frame(parent)
        frame.grid(row=row, column=col, sticky="ew", padx=(0, 12), pady=5)
        parent.columnconfigure(col, weight=1)
        ttk.Label(frame, text=label).pack(anchor="w")
        ttk.Entry(frame, textvariable=var).pack(fill=X)

    def _preview_address(self):
        completed, status = complete_anhui_address(self.address_input.get())
        self.address_output.set(f"{completed}    [{status}]")

    def _pick_addr_input(self):
        path = filedialog.askopenfilename(filetypes=[("Excel 文件", "*.xlsx")])
        if path:
            self.addr_input_file.set(path)
            if not self.addr_output_file.get():
                self.addr_output_file.set(str(Path(path).with_name(Path(path).stem + "_地址补全.xlsx")))

    def _pick_addr_output(self):
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel 文件", "*.xlsx")])
        if path:
            self.addr_output_file.set(path)

    def _pick_raw_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel 文件", "*.xlsx")])
        if path:
            self.raw_file.set(path)
            if not self.extract_output_file.get():
                self.extract_output_file.set(str(Path(path).with_name(Path(path).stem + "_字段提取属性匹配.xlsx")))

    def _pick_master_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel 文件", "*.xlsx")])
        if path:
            self.master_file.set(path)

    def _pick_extract_output(self):
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel 文件", "*.xlsx")])
        if path:
            self.extract_output_file.set(path)

    def _pick_a_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel 文件", "*.xlsx")])
        if path:
            self.a_file.set(path)
            if not self.match_output_file.get():
                self.match_output_file.set(str(Path(path).with_name(Path(path).stem + "_已匹配发货单号.xlsx")))

    def _pick_b_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel 文件", "*.xlsx")])
        if path:
            self.b_file.set(path)

    def _pick_match_output(self):
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel 文件", "*.xlsx")])
        if path:
            self.match_output_file.set(path)

    def _run_address_batch(self):
        self._start_job(
            "正在补全地址...",
            lambda: complete_address_workbook(
                Path(self.addr_input_file.get()),
                Path(self.addr_output_file.get()),
                self.addr_column.get().strip(),
            ),
            "地址补全完成",
        )

    def _run_extract_merge(self):
        self._start_job(
            "正在提取字段并匹配机型属性...",
            lambda: extract_and_merge_workbook(
                Path(self.raw_file.get()),
                Path(self.master_file.get()),
                Path(self.extract_output_file.get()),
                self.extract_columns.get(),
                self.rename_columns.get(),
                self.raw_code_header.get().strip(),
                self.master_code_header.get().strip(),
            ),
            "字段提取与属性匹配完成",
        )

    def _run_match(self):
        columns = MatchColumns(
            self.a_phone.get().strip(),
            self.a_name.get().strip(),
            self.a_model.get().strip(),
            self.a_fuzzy_model.get().strip(),
            self.b_phone.get().strip(),
            self.b_name.get().strip(),
            self.b_model.get().strip(),
            self.b_ship_no.get().strip(),
        )
        self._start_job(
            "正在匹配 A/B 表...",
            lambda: match_order_workbooks(
                Path(self.a_file.get()),
                Path(self.b_file.get()),
                Path(self.match_output_file.get()),
                columns,
            ),
            "A/B 表匹配完成",
        )

    def _start_job(self, message, func, success_title):
        self.log_text.config(text=message)
        threading.Thread(target=self._worker, args=(func, success_title), daemon=True).start()

    def _worker(self, func, success_title):
        try:
            result = func()
            self.jobs.put(("success", success_title, result))
        except Exception as exc:
            self.jobs.put(("error", str(exc), traceback.format_exc()))

    def _poll_jobs(self):
        try:
            kind, title, payload = self.jobs.get_nowait()
        except queue.Empty:
            self.after(100, self._poll_jobs)
            return

        if kind == "success":
            summary = "，".join(f"{key}={value}" for key, value in payload.items())
            self.log_text.config(text=f"{title}：{summary}")
            messagebox.showinfo(title, summary)
        else:
            self.log_text.config(text=f"处理失败：{title}")
            messagebox.showerror("处理失败", f"{title}\n\n详细信息已显示在终端。")
            print(payload)
        self.after(100, self._poll_jobs)


if __name__ == "__main__":
    ExcelHelperApp().mainloop()
