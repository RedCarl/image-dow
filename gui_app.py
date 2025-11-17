#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os

from download_images import process_excel


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("商品图片下载器")
        self.geometry("720x420")

        self.cancel_event = None
        self.worker = None

        self._build_ui()

    def _build_ui(self):
        frm = ttk.Frame(self, padding=12)
        frm.pack(fill=tk.BOTH, expand=True)

        # Excel 文件
        ttk.Label(frm, text="Excel 文件").grid(row=0, column=0, sticky=tk.W, pady=4)
        self.input_var = tk.StringVar()
        ent_input = ttk.Entry(frm, textvariable=self.input_var, width=60)
        ent_input.grid(row=0, column=1, sticky=tk.EW, padx=6)
        frm.columnconfigure(1, weight=1)
        ttk.Button(frm, text="选择...", command=self.pick_file).grid(row=0, column=2, padx=6)

        # 工作表
        ttk.Label(frm, text="工作表").grid(row=1, column=0, sticky=tk.W, pady=4)
        self.sheet_var = tk.StringVar(value="Sheet1")
        ttk.Entry(frm, textvariable=self.sheet_var, width=20).grid(row=1, column=1, sticky=tk.W, padx=6)

        # 输出目录
        ttk.Label(frm, text="输出目录").grid(row=2, column=0, sticky=tk.W, pady=4)
        self.out_var = tk.StringVar(value=os.path.join(os.getcwd(), "images"))
        ttk.Entry(frm, textvariable=self.out_var, width=60).grid(row=2, column=1, sticky=tk.EW, padx=6)
        ttk.Button(frm, text="浏览...", command=self.pick_folder).grid(row=2, column=2, padx=6)

        # 范围
        ttk.Label(frm, text="开始行").grid(row=3, column=0, sticky=tk.W, pady=4)
        self.start_var = tk.StringVar(value="2")
        ttk.Entry(frm, textvariable=self.start_var, width=10).grid(row=3, column=1, sticky=tk.W, padx=6)

        ttk.Label(frm, text="结束行(可选)").grid(row=4, column=0, sticky=tk.W, pady=4)
        self.end_var = tk.StringVar(value="")
        ttk.Entry(frm, textvariable=self.end_var, width=10).grid(row=4, column=1, sticky=tk.W, padx=6)

        ttk.Label(frm, text="并发数").grid(row=5, column=0, sticky=tk.W, pady=4)
        self.concurrency_var = tk.StringVar(value="4")
        ttk.Entry(frm, textvariable=self.concurrency_var, width=10).grid(row=5, column=1, sticky=tk.W, padx=6)

        # 进度条
        self.progress = ttk.Progressbar(frm, mode="determinate")
        self.progress.grid(row=6, column=0, columnspan=3, sticky=tk.EW, pady=10)

        # 状态日志
        self.log = tk.Text(frm, height=10)
        self.log.grid(row=7, column=0, columnspan=3, sticky=tk.NSEW)
        frm.rowconfigure(7, weight=1)

        # 按钮
        btn_frame = ttk.Frame(frm)
        btn_frame.grid(row=8, column=0, columnspan=3, sticky=tk.E, pady=8)
        self.start_btn = ttk.Button(btn_frame, text="开始下载", command=self.start_download)
        self.start_btn.pack(side=tk.LEFT, padx=6)
        self.cancel_btn = ttk.Button(btn_frame, text="取消", command=self.cancel_download, state=tk.DISABLED)
        self.cancel_btn.pack(side=tk.LEFT)

        # 拖拽提示
        ttk.Label(frm, text="提示：可点击“选择...”或将Excel路径粘贴到输入框").grid(row=9, column=0, columnspan=3, sticky=tk.W)

    def pick_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx"), ("所有文件", "*.*")])
        if path:
            self.input_var.set(path)

    def pick_folder(self):
        path = filedialog.askdirectory()
        if path:
            self.out_var.set(path)

    def start_download(self):
        input_path = self.input_var.get().strip()
        sheet = self.sheet_var.get().strip() or "Sheet1"
        out_dir = self.out_var.get().strip() or os.path.join(os.getcwd(), "images")
        try:
            start_row = int(self.start_var.get() or "2")
        except Exception:
            messagebox.showerror("错误", "开始行必须是数字")
            return
        end_val = self.end_var.get().strip()
        end_row = int(end_val) if end_val else None

        try:
            concurrency = int(self.concurrency_var.get() or "4")
            if concurrency <= 0:
                raise ValueError()
        except Exception:
            messagebox.showerror("错误", "并发数必须是正整数")
            return

        if not input_path or not os.path.isfile(input_path):
            messagebox.showerror("错误", "请选择有效的Excel文件")
            return

        self.log_delete()
        self.progress.configure(value=0, maximum=100)
        self.start_btn.configure(state=tk.DISABLED)
        self.cancel_btn.configure(state=tk.NORMAL)

        self.cancel_event = threading.Event()

        def on_progress(info):
            status = info.get("status")
            processed = info.get("processed", 0)
            total = info.get("total", 0)
            filename = info.get("filename", "")
            pct = 0
            if total:
                pct = int(processed * 100 / total)
            self.after(0, self.update_progress, pct, status, filename, processed, total)

        def worker():
            try:
                process_excel(
                    input_path=input_path,
                    sheet_name=sheet,
                    output_dir=out_dir,
                    start_row=start_row,
                    end_row=end_row,
                    limit=None,
                    on_progress=on_progress,
                    cancel_event=self.cancel_event,
                    concurrency=concurrency,
                )
            except Exception as e:
                self.after(0, lambda: messagebox.showerror("错误", str(e)))
            finally:
                self.after(0, self.finish_download)

        self.worker = threading.Thread(target=worker, daemon=True)
        self.worker.start()

    def cancel_download(self):
        if self.cancel_event:
            self.cancel_event.set()
            self.append_log("已请求取消...\n")

    def finish_download(self):
        self.start_btn.configure(state=tk.NORMAL)
        self.cancel_btn.configure(state=tk.DISABLED)
        self.append_log("完成\n")

    def update_progress(self, pct, status, filename, processed, total):
        self.progress.configure(value=pct, maximum=100)
        if status == "success":
            self.append_log(f"下载成功：{filename}\n")
        elif status == "skip":
            self.append_log(f"已存在，跳过：{filename}\n")
        elif status == "fail":
            self.append_log(f"下载失败：{filename}\n")
        elif status == "done":
            self.append_log(f"完成：{processed}/{total}\n")

    def append_log(self, text):
        self.log.insert(tk.END, text)
        self.log.see(tk.END)

    def log_delete(self):
        self.log.delete("1.0", tk.END)


def main():
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()