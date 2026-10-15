"""Actual PDF page preview with paging and fit-to-width rendering."""
import tkinter as tk
from tkinter import ttk

import pypdfium2 as pdfium
from PIL import ImageTk


class PagePreview(tk.Toplevel):
    def __init__(self, parent, pdf_path, cleanup):
        super().__init__(parent)
        self.title("文档排版预览")
        self.geometry(f"{min(parent.px(850), self.winfo_screenwidth()-80)}x{min(parent.px(750), self.winfo_screenheight()-100)}")
        self.document = pdfium.PdfDocument(pdf_path)
        self.page_index = 0
        self.cleanup = cleanup
        self.photo = None
        self.resize_timer = None
        toolbar = ttk.Frame(self, padding=parent.px(12))
        toolbar.pack(fill="x")
        ttk.Button(toolbar, text="上一页", command=lambda: self.change(-1)).pack(side="left")
        self.page_label = ttk.Label(toolbar)
        self.page_label.pack(side="left", padx=20)
        ttk.Button(toolbar, text="下一页", command=lambda: self.change(1)).pack(side="left")
        ttk.Label(toolbar, text="按窗口宽度显示，可滚动查看页面").pack(side="right")
        frame = ttk.Frame(self)
        frame.pack(fill="both", expand=True)
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(0, weight=1)
        self.canvas = tk.Canvas(frame, bg="#E8ECE1", highlightthickness=0)
        self.canvas.grid(row=0, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(frame, orient="vertical", command=self.canvas.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.canvas.configure(yscrollcommand=scrollbar.set)
        self.canvas.bind("<Configure>", self.resize)
        self.canvas.bind("<MouseWheel>", lambda event: self.canvas.yview_scroll(int(-event.delta / 120), "units"))
        self.protocol("WM_DELETE_WINDOW", self.close)
        self.after(100, self.render)

    def resize(self, event=None):
        if self.resize_timer:
            self.after_cancel(self.resize_timer)
        self.resize_timer = self.after(150, self.render)

    def change(self, offset):
        new_index = max(0, min(len(self.document) - 1, self.page_index + offset))
        if new_index != self.page_index:
            self.page_index = new_index
            self.render()
            self.canvas.yview_moveto(0)

    def render(self):
        self.resize_timer = None
        if self.document is None:
            return
        page = self.document[self.page_index]
        bitmap = page.render(scale=min(2, max(0.3, (self.canvas.winfo_width() - 30) / page.get_width())))
        picture = bitmap.to_pil()
        self.photo = ImageTk.PhotoImage(picture, master=self)
        self.canvas.delete("all")
        center = max(15, self.canvas.winfo_width() // 2)
        self.canvas.create_rectangle(center - picture.width / 2 + 3, 18, center + picture.width / 2 + 3,
                                     picture.height + 18, fill="#CCD5C6", outline="")
        self.canvas.create_image(center, 15, anchor="n", image=self.photo)
        self.canvas.configure(scrollregion=(0, 0, self.canvas.winfo_width(), picture.height + 30))
        self.page_label.configure(text=f"第 {self.page_index + 1} / {len(self.document)} 页")
        bitmap.close()
        page.close()

    def close(self):
        if self.resize_timer:
            self.after_cancel(self.resize_timer)
        if self.document is not None:
            self.document.close()
            self.document = None
        self.photo = None
        self.destroy()
        self.cleanup()
