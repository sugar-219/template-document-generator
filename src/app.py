"""Windows desktop entry point. Tk widgets stay on the main thread."""
from __future__ import annotations

import ctypes
import json
import os
import queue
import shutil
import sys
import tempfile
import threading
import time
import traceback
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, font as tkfont, messagebox, ttk

from core import InputError, Template, check_records, continue_batch, file_stem, generate_batch, pdf_converter, preview_filenames, read_data, retry_batch, sheet_names, values_for
from check_report import write_check_report
from motion import Motion
from pdf_worker import PdfSession
from preferences import Preferences
from preview_cache import PreviewCache

TITLE = "模板文档批量生成软件"
BG, INK, MUTED, GREEN, SIDE = "#F5F7F4", "#203B34", "#6B8176", "#187C67", "#173F37"
BORDER, SOFT = "#DFE7DF", "#EFF4EE"
SELECT = "#E1F0E5"
FONT = "Microsoft YaHei UI"
PAGE_SIZE = 200
RESULT_COLUMNS = ("数据行", "Word 文件", "PDF 文件", "状态", "说明")


def app_directory() -> Path:
    return Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parent.parent


def resource(name: str) -> Path:
    return Path(getattr(sys, "_MEIPASS", app_directory())) / name


class Application(tk.Tk):
    def __init__(self, settings_path=None):
        super().__init__()
        self.title(TITLE)
        self.configure(bg=BG)
        self.ui_scale = max(1, float(self.tk.call("tk", "scaling")) / (96 / 72))
        self.px = lambda value: round(value * self.ui_scale)
        self.minsize(min(self.px(1020), self.winfo_screenwidth()-40), min(self.px(700), self.winfo_screenheight()-80))
        width = min(self.px(1180), self.winfo_screenwidth() - 60)
        height = min(self.px(860), self.winfo_screenheight() - 80)
        self.geometry(f"{width}x{height}+{max(0, (self.winfo_screenwidth()-width)//2)}+{max(0, (self.winfo_screenheight()-height)//2-20)}")
        try:
            self.iconbitmap(str(resource("assets/app.ico")))
        except tk.TclError:
            pass
        self.template = None
        self.dataset = None
        self.data_path = None
        self.mapping = {}
        self.drawing_mapping = False
        self.preferences = Preferences(settings_path)
        self.save_timer = None
        self.preview_windows = []
        self.preview_cache = PreviewCache(self.preferences.path.parent / "previews")
        self.preview_root = self.preview_cache.path
        self.preview_folders = set()
        self.active_preview_folders = set()
        self.selected_indices = set()
        self.data_page_number = 0
        self.data_indices = []
        self.data_search_text = ""
        self.filling_data = False
        self.result_items = []
        self.result_page_number = 0
        self.last_batch = None
        self.last_job = None
        self.issues = []
        self.check_completed = False
        self.issue_page_number = 0
        self.busy = False
        self.close_requested = False
        self.last_folder = None
        self.events = queue.Queue()
        self.cancel_event = threading.Event()
        self.controls = []
        self.naming_buttons = []
        self.pdf_checks = []
        self.reason_popup = None
        self.operation_started = None
        self.operation_phase = ""
        self.operation_current = 0
        self.operation_total = 0
        self.operation_completed = 0
        self.operation_sample_elapsed = 0
        self.last_operation_elapsed = None
        self.base = app_directory()
        self.output = tk.StringVar(value=self.preferences.data.get("output", str(self.base / "输出文档")))
        self.pattern = tk.StringVar(value=self.preferences.data.get("pattern", "{{序号}}_文档"))
        self.sheet = tk.StringVar()
        self.preview_index = tk.StringVar(value="1")
        self.scope = tk.StringVar(value="全部记录")
        self.data_query = tk.StringVar()
        self.locate_row = tk.StringVar()
        self.result_filter = tk.StringVar(value="全部结果")
        self.pdf = tk.BooleanVar(value=bool(self.preferences.data.get("pdf", False)))
        self.motion_enabled = tk.BooleanVar(value=bool(self.preferences.data.get("animations", True)))
        self.status = tk.StringVar(value="先选择模板和数据，或点击“载入示例”体验完整流程。")
        self.action_hint = tk.StringVar()
        self.converter = pdf_converter()
        if not self.converter:
            self.pdf.set(False)
        self._styles()
        self._build()
        for variable in (self.output, self.pattern, self.pdf):
            variable.trace_add("write", lambda *args: self._settings_changed())
        self.scope.trace_add("write", lambda *args: self._refresh_ready())
        self.scope.trace_add("write", lambda *args: self._show_issues([]))
        self.pattern.trace_add("write", lambda *args: self._show_issues([]))
        self.motion_enabled.trace_add("write", lambda *args: self._motion_changed())
        self.protocol("WM_DELETE_WINDOW", self._close)
        self.after(100, self._poll)
        self.after(30000, self._cleanup_previews)

    def _styles(self):
        self.option_add("*Font", (FONT, 9))
        self.table_font = tkfont.Font(root=self, font=(FONT, 9))
        self.metric_font = "Bahnschrift" if "Bahnschrift" in tkfont.families(self) else FONT
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("TFrame", background=BG)
        style.configure("Card.TFrame", background="white")
        style.configure("TLabel", background=BG, foreground=INK)
        style.configure("TButton", font=(FONT, 9), width=0, padding=(self.px(13), self.px(6)), background="white", foreground=INK,
                        bordercolor=BORDER, lightcolor="white", darkcolor="white", borderwidth=1, relief="raised", focusthickness=1, focuscolor=GREEN)
        style.map("TButton", background=[("disabled", "#F6F8F9"), ("pressed", SELECT), ("active", SOFT)],
                  foreground=[("disabled", "#8B9AA5")], bordercolor=[("focus", GREEN), ("active", "#A7BFC9")])
        style.configure("Accent.TButton", background=GREEN, foreground="white", bordercolor=GREEN,
                        lightcolor=GREEN, darkcolor=GREEN, font=(FONT, 10, "bold"), padding=(self.px(20), self.px(9)))
        style.map("Accent.TButton", background=[("disabled", "#B5C9CD"), ("pressed", "#0A5F5B"), ("active", "#106E69")],
                  foreground=[("disabled", "white")], bordercolor=[("disabled", "#B5C9CD"), ("active", "#106E69")])
        style.configure("Quiet.TButton", background=BG, bordercolor=BG, lightcolor=BG, darkcolor=BG, foreground=MUTED, relief="flat")
        style.configure("Sidebar.TButton", background="#245449", foreground="#F4FBF5", bordercolor="#356B5E",
                        lightcolor="#245449", darkcolor="#245449", padding=(self.px(12), self.px(10)), anchor="w")
        style.map("Sidebar.TButton", background=[("disabled", "#20473E"), ("pressed", "#316B58"), ("active", "#2C6151")],
                  foreground=[("disabled", "#89A99A")], bordercolor=[("focus", "#B5D9C5"), ("active", "#83AB95")])
        style.configure("Sidebar.TCombobox", fieldbackground="#245449", foreground="#F4FBF5", background="#245449",
                        bordercolor="#356B5E", lightcolor="#245449", darkcolor="#245449", arrowcolor="#CAE2D0")
        style.map("Sidebar.TCombobox", fieldbackground=[("disabled", "#20473E"), ("readonly", "#245449")],
                  foreground=[("disabled", "#89A99A"), ("readonly", "#F4FBF5")], bordercolor=[("focus", "#B5D9C5")])
        for name in ("TEntry", "TCombobox", "TSpinbox"):
            style.configure(name, padding=self.px(5), fieldbackground="white", foreground=INK, background=SOFT, arrowcolor=MUTED, bordercolor=BORDER,
                            lightcolor="white", darkcolor="white", arrowsize=self.px(12))
            style.map(name, bordercolor=[("focus", GREEN)], fieldbackground=[("disabled", BG), ("readonly", "white")],
                      foreground=[("disabled", "#8B9AA5"), ("readonly", INK)])
        style.configure("Treeview", font=(FONT, 9), rowheight=self.px(34), background="white", fieldbackground="white", foreground=INK,
                        bordercolor=BORDER, lightcolor="white", darkcolor="white", borderwidth=0)
        style.configure("Treeview.Heading", font=(FONT, 9, "bold"), padding=(self.px(10), self.px(9)), background=SOFT,
                        foreground="#486577", bordercolor=SOFT, lightcolor=SOFT, darkcolor=SOFT, relief="flat")
        style.map("Treeview.Heading", background=[("active", "#E1EBF0")])
        style.map("Treeview", background=[("selected", SELECT)], foreground=[("selected", "#105B57")])
        style.configure("TNotebook", background=BG, bordercolor=BORDER, lightcolor=BORDER, darkcolor=BORDER, borderwidth=1, tabmargins=(0, 0, 0, 0))
        style.configure("TNotebook.Tab", font=(FONT, 9, "bold"), padding=(self.px(18), self.px(10)), background=BG,
                        foreground=MUTED, bordercolor=BG, lightcolor=BG, darkcolor=BG, focuscolor=GREEN)
        style.map("TNotebook.Tab", background=[("selected", GREEN), ("active", SOFT)], foreground=[("selected", "white")],
                  bordercolor=[("selected", GREEN)], lightcolor=[("selected", GREEN)], darkcolor=[("selected", GREEN)])
        style.layout("Workspace.TNotebook.Tab", [])
        for name, background, foreground in (("Navigation.TButton", BG, MUTED), ("ActiveNavigation.TButton", "white", GREEN)):
            style.configure(name, background=background, foreground=foreground, bordercolor=background, lightcolor=background,
                            darkcolor=background, borderwidth=0, relief="flat", font=(FONT, 9, "bold"), padding=(self.px(12), self.px(8)))
            style.map(name, background=[("pressed", SELECT), ("active", SOFT)], bordercolor=[("focus", GREEN)])
        style.configure("Horizontal.TProgressbar", background=GREEN, troughcolor="#DFE8EC", bordercolor="#DFE8EC", lightcolor=GREEN, darkcolor=GREEN, borderwidth=0, thickness=self.px(5))
        for name in ("TCheckbutton", "TRadiobutton"):
            style.configure(name, font=(FONT, 9), background="white", foreground=INK, indicatorbackground="white", indicatorforeground=GREEN, focuscolor=GREEN)
            style.map(name, background=[("active", "white")], foreground=[("disabled", "#8B9AA5")])
        style.configure("Header.TCheckbutton", background=BG, foreground=MUTED)
        style.map("Header.TCheckbutton", background=[("active", BG)])
        for orientation, sticky in (("Vertical", "ns"), ("Horizontal", "ew")):
            name = orientation + ".TScrollbar"
            style.layout(name, [(orientation + ".Scrollbar.trough", {"sticky": sticky, "children": [
                (orientation + ".Scrollbar.thumb", {"expand": 1, "sticky": "nsew"})]})])
            style.configure(name, background="#C4D3DC", troughcolor="#F3F6F8", bordercolor="#F3F6F8",
                            lightcolor="#C4D3DC", darkcolor="#C4D3DC", borderwidth=0, width=self.px(9))
            style.map(name, background=[("active", "#9AB4C3"), ("pressed", "#7596A8")])

    def _button(self, parent, text, command, style="TButton"):
        widget = ttk.Button(parent, text=text, command=command, style=style)
        self.controls.append((widget, "normal"))
        self._bind_disabled_reason(widget)
        return widget

    def _bind_disabled_reason(self, widget):
        widget.bind("<Enter>", lambda _: self._show_disabled_reason(widget))
        widget.bind("<Leave>", lambda _: self._hide_disabled_reason())
        widget.bind("<Button-1>", lambda _: self._hide_disabled_reason())

    def _availability(self, widget, enabled, reason):
        widget.configure(state="normal" if enabled and not self.busy else "disabled")
        widget.disabled_reason = reason

    def _show_disabled_reason(self, widget):
        self._hide_disabled_reason()
        if not widget.instate(["disabled"]):
            return None
        reason = "正在处理，请等待当前操作结束。" if self.busy else getattr(widget, "disabled_reason", "")
        if not reason:
            return None
        window = self.reason_popup = tk.Toplevel(self)
        window.overrideredirect(True)
        window.attributes("-topmost", True)
        window.reason_text = self._label(window, reason, 9, INK, SOFT, wraplength=self.px(300), justify="left")
        window.reason_text.pack(padx=self.px(10), pady=self.px(8))
        window.update_idletasks()
        x = min(widget.winfo_rootx(), self.winfo_screenwidth() - window.winfo_reqwidth() - 10)
        y = min(widget.winfo_rooty() + widget.winfo_height() + 5, self.winfo_screenheight() - window.winfo_reqheight() - 10)
        window.geometry(f"+{max(0, x)}+{max(0, y)}")
        return window

    def _hide_disabled_reason(self):
        if self.reason_popup is not None:
            self.reason_popup.destroy()
            self.reason_popup = None

    def _label(self, parent, text="", size=9, color=INK, bg="white", bold=False, **kwargs):
        return tk.Label(parent, text=text, font=(FONT, size, "bold" if bold else "normal"), fg=color, bg=bg, anchor="w", **kwargs)

    def _build(self):
        sidebar = tk.Frame(self, bg=SIDE, width=self.px(218))
        sidebar.pack(side="left", fill="y")
        sidebar.pack_propagate(False)
        brand = tk.Frame(sidebar, bg=SIDE)
        self.brand = brand
        brand.pack(fill="x", padx=self.px(20), pady=(self.px(26), self.px(28)))
        mark = tk.Canvas(brand, width=self.px(32), height=self.px(36), bg=SIDE, highlightthickness=0)
        self.brand_mark = mark
        mark.pack(side="left", padx=(0, self.px(12)))
        mark.create_polygon(*[self.px(value) for value in (3, 2, 21, 2, 29, 10, 29, 33, 3, 33)], outline="#83C8BC", fill=SIDE, width=self.px(1))
        mark.create_line(*[self.px(value) for value in (21, 2, 21, 10, 29, 10)], fill="#83C8BC", width=self.px(1))
        for line in (16, 22, 28):
            mark.create_line(self.px(9), self.px(line), self.px(22 if line != 28 else 17), self.px(line), fill="#83C8BC", width=self.px(1))
        self.brand_text = tk.Frame(brand, bg=SIDE)
        self.brand_text.pack(side="left")
        self.brand_title = self._label(self.brand_text, "模板文档", 16, "#F0F7FA", SIDE, True)
        self.brand_title.pack(anchor="w")
        self.brand_subtitle = self._label(self.brand_text, "批量生成工作台", 8, "#9BB4C2", SIDE)
        self.brand_subtitle.pack(anchor="w", pady=(self.px(5), 0))
        tk.Frame(sidebar, bg="#304957", height=1).pack(fill="x", padx=self.px(20), pady=(0, self.px(20)))
        self._label(sidebar, "文件来源", 8, "#9BB4C2", SIDE).pack(anchor="w", padx=self.px(20))
        self._button(sidebar, "01   选择 Word 模板", self.choose_template, "Sidebar.TButton").pack(fill="x", padx=self.px(16), pady=(self.px(10), self.px(7)))
        self.template_label = self._label(sidebar, "尚未选择 .docx", 8, "#BACED8", SIDE, wraplength=self.px(178))
        self.template_label.pack(fill="x", padx=self.px(20), pady=(0, self.px(20)))
        self._button(sidebar, "02   选择 Excel / CSV", self.choose_data, "Sidebar.TButton").pack(fill="x", padx=self.px(16), pady=(0, self.px(7)))
        self.data_label = self._label(sidebar, "尚未选择数据", 8, "#BACED8", SIDE, wraplength=self.px(178))
        self.data_label.pack(fill="x", padx=self.px(20), pady=(0, self.px(16)))
        self._label(sidebar, "工作表", 8, "#9BB4C2", SIDE).pack(anchor="w", padx=self.px(20), pady=(0, self.px(6)))
        self.sheet_combo = ttk.Combobox(sidebar, textvariable=self.sheet, state="disabled", style="Sidebar.TCombobox")
        self.sheet_combo.pack(fill="x", padx=self.px(16))
        self.sheet_combo.bind("<<ComboboxSelected>>", lambda _: self.reload_sheet())
        self.controls.append((self.sheet_combo, "readonly"))
        file_actions = tk.Frame(sidebar, bg=SIDE)
        file_actions.pack(fill="x", padx=self.px(16), pady=(self.px(10), 0))
        file_actions.columnconfigure((0, 1), weight=1, uniform="file_actions")
        self.refresh_button = self._button(file_actions, "刷新数据", self.refresh_data, "Sidebar.TButton")
        self.refresh_button.grid(row=0, column=0, sticky="ew", padx=(0, self.px(3)))
        self.recent_button = self._button(file_actions, "最近文件", self.open_recent_files, "Sidebar.TButton")
        self.recent_button.grid(row=0, column=1, sticky="ew", padx=(self.px(3), 0))
        bottom = tk.Frame(sidebar, bg=SIDE)
        bottom.pack(side="bottom", fill="x", padx=self.px(20), pady=self.px(20))
        self.sidebar_tips = tk.Frame(bottom, bg=SIDE)
        self.sidebar_tips.pack(fill="x")
        tk.Frame(self.sidebar_tips, bg="#304957", height=1).pack(fill="x", pady=(0, self.px(14)))
        self._label(self.sidebar_tips, "占位符写法", 8, "#9BB4C2", SIDE).pack(anchor="w")
        self._label(self.sidebar_tips, "{{姓名}}  {{部门}}", 10, "#83C8BC", SIDE).pack(anchor="w", pady=(self.px(6), self.px(8)))
        self._label(self.sidebar_tips, "数据首行为列标题\n一行数据生成一份文档", 8, "#BACED8", SIDE, justify="left").pack(anchor="w")
        self._button(bottom, "使用说明", self.show_help, "Sidebar.TButton").pack(fill="x", pady=(self.px(14), 0))

        main = tk.Frame(self, bg=BG)
        self.main = main
        main.pack(side="left", fill="both", expand=True, padx=self.px(22), pady=self.px(22))
        main.columnconfigure(0, weight=1)
        main.rowconfigure(2, weight=1)
        header = tk.Frame(main, bg=BG)
        self.header = header
        header.grid(row=0, column=0, sticky="ew", pady=(0, self.px(18)))
        self.demo_button = self._button(header, "载入示例", self.load_demo)
        self.demo_button.pack(side="right", padx=(10, 0))
        self.motion_check = ttk.Checkbutton(header, text="界面动画", variable=self.motion_enabled, style="Header.TCheckbutton")
        self.motion_check.pack(side="right", padx=(self.px(12), 0))
        self.compact_export_button = self._button(header, "导出设置…", self.open_export_settings)
        self._label(header, "批量文档工作台", 21, bg=BG, bold=True).pack(anchor="w")
        self.header_subtitle = self._label(header, "准备一份模板，让每一行数据都有自己的文档。", 9, MUTED, BG)
        self.header_subtitle.pack(anchor="w", pady=(self.px(5), 0))
        cards = tk.Frame(main, bg=BG)
        self.cards = cards
        cards.grid(row=1, column=0, sticky="ew", pady=(0, self.px(16)))
        self.metrics = []
        for index, title in enumerate(("模板字段", "数据记录", "字段匹配")):
            cards.columnconfigure(index, weight=1, uniform="metrics")
            card_bg = SELECT if index == 2 else "white"
            card = tk.Frame(cards, bg=card_bg, highlightbackground="#C2DED9" if index == 2 else BORDER, highlightthickness=1)
            card.grid(row=0, column=index, sticky="nsew", padx=(0 if index == 0 else self.px(10), 0))
            tk.Frame(card, bg=GREEN if index == 2 else "#ABC4B4", width=self.px(3)).pack(side="left", fill="y")
            content = tk.Frame(card, bg=card_bg)
            content.pack(fill="both", expand=True)
            self._label(content, title, 8, MUTED, card_bg).pack(anchor="w", padx=self.px(16), pady=(self.px(8), 0))
            number = self._label(content, "—", 22, GREEN if index == 2 else INK, card_bg, bold=True)
            number.configure(font=(self.metric_font, 22, "bold"))
            number.pack(anchor="w", padx=self.px(16), pady=(self.px(2), self.px(5)))
            self.metrics.append(number)

        workspace = tk.Frame(main, bg=BG)
        workspace.grid(row=2, column=0, sticky="nsew")
        workspace.columnconfigure(0, weight=1)
        workspace.rowconfigure(1, weight=1)
        self.navigation = tk.Frame(workspace, bg=BG)
        self.navigation.grid(row=0, column=0, sticky="ew")
        self.notebook = ttk.Notebook(workspace, height=self.px(150), style="Workspace.TNotebook")
        self.notebook.grid(row=1, column=0, sticky="nsew")
        self.mapping_page = ttk.Frame(self.notebook, style="Card.TFrame")
        self.data_page = ttk.Frame(self.notebook, style="Card.TFrame")
        self.check_page = ttk.Frame(self.notebook, style="Card.TFrame")
        self.preview_page = ttk.Frame(self.notebook, style="Card.TFrame")
        self.result_page = ttk.Frame(self.notebook, style="Card.TFrame")
        self.nav_buttons = []
        for index, (frame, name) in enumerate(((self.mapping_page, "字段匹配"), (self.data_page, "数据预览"), (self.check_page, "数据检查"), (self.preview_page, "文档预览"), (self.result_page, "生成结果"))):
            self.notebook.add(frame, text=name)
            self.navigation.columnconfigure(index, weight=1, uniform="navigation")
            button = ttk.Button(self.navigation, text=f"0{index+1}  {name}", command=lambda page=frame: self.notebook.select(page), style="Navigation.TButton")
            button.grid(row=0, column=index, sticky="ew")
            self.nav_buttons.append(button)
        self.nav_line = tk.Canvas(self.navigation, height=self.px(3), bg=BG, highlightthickness=0)
        self.nav_line.grid(row=1, column=0, columnspan=5, sticky="ew")
        self.nav_indicator = self.nav_line.create_rectangle(0, 0, 0, self.px(3), fill=GREEN, outline="")
        motion_enabled = self.motion_enabled.get()
        self.nav_motion = Motion(self, motion_enabled)
        self.nav_line.bind("<Configure>", lambda _: self.after_idle(lambda: self._sync_navigation(animate=False)))
        self.notebook.bind("<<NotebookTabChanged>>", lambda _: self._sync_navigation())
        self.notebook.enable_traversal()
        self._mapping_view()
        self._data_view()
        self._check_view()
        self._preview_view()
        self._results_view()

        self.export_panel, self.pdf_check = self._export_settings(main)
        self.export_panel.grid(row=3, column=0, sticky="ew", pady=(self.px(14), self.px(12)))
        action = tk.Frame(main, bg=BG)
        action.grid(row=4, column=0, sticky="ew")
        self.generate_button = self._button(action, "生成全部 0 份", self.start_generation, "Accent.TButton")
        self.generate_button.pack(side="right")
        self.cancel_button = ttk.Button(action, text="停止生成", command=self.cancel_event.set, state="disabled", style="Quiet.TButton")
        self.cancel_button.pack(side="right", padx=10)
        self.progressbar = ttk.Progressbar(action, mode="determinate", length=130)
        self.progress_motion = Motion(self, motion_enabled)
        self.progressbar.pack(side="left", fill="x", expand=True, padx=(0, 16))
        self.action_hint_label = tk.Label(main, textvariable=self.action_hint, bg=BG, fg=GREEN, font=(FONT, 8), anchor="w", justify="left")
        self.action_hint_label.grid(row=5, column=0, sticky="ew", pady=(self.px(6), 0))
        self.action_hint_label.bind("<Configure>", lambda event: self.action_hint_label.configure(wraplength=event.width))
        self.status_label = tk.Label(main, textvariable=self.status, bg=BG, fg=MUTED, font=(FONT, 8), anchor="w", justify="left", wraplength=self.px(850))
        self.status_label.grid(row=6, column=0, sticky="ew", pady=(self.px(4), 0))
        self.status_label.bind("<Configure>", lambda event: self.status_label.configure(wraplength=event.width))
        self.compact_layout = None
        self.bind("<Configure>", self._resize_layout)
        self.after_idle(self._resize_layout)
        self._refresh_ready()

    def _sync_navigation(self, animate=True):
        selected = self.notebook.index(self.notebook.select())
        for index, button in enumerate(self.nav_buttons):
            button.configure(style="ActiveNavigation.TButton" if index == selected else "Navigation.TButton")
        button = self.nav_buttons[selected]
        destination = (button.winfo_x()+self.px(8), button.winfo_x()+button.winfo_width()-self.px(8))
        coordinates = self.nav_line.coords(self.nav_indicator)
        origin = (coordinates[0], coordinates[2])
        def draw(position):
            self.nav_line.coords(self.nav_indicator, position[0], 0, position[1], self.px(3))
        if animate and self.nav_line.winfo_ismapped():
            self.nav_motion.start(origin, destination, draw)
        else:
            self.nav_motion.cancel()
            draw(destination)

    def _animate_progress(self, value):
        self.progress_motion.start((float(self.progressbar.cget("value")),), (float(value),),
                                   lambda position: self.progressbar.configure(value=position[0]), duration=220)

    def _motion_changed(self):
        enabled = self.motion_enabled.get()
        self.nav_motion.enabled = self.progress_motion.enabled = enabled
        if not enabled:
            self._sync_navigation(animate=False)
            self.progress_motion.finish()
        self._schedule_save()

    def _export_settings(self, parent):
        self.controls = [(widget, state) for widget, state in self.controls if widget.winfo_exists()]
        settings = tk.Frame(parent, bg="white", highlightthickness=1, highlightbackground=BORDER)
        settings.columnconfigure(1, weight=1)
        self._label(settings, "文件命名", 9).grid(row=0, column=0, padx=(self.px(14), self.px(10)), pady=(self.px(10), self.px(4)))
        naming = ttk.Entry(settings, textvariable=self.pattern)
        naming.grid(row=0, column=1, sticky="ew", pady=(self.px(10), self.px(4)))
        self.controls.append((naming, "normal"))
        pdf_check = ttk.Checkbutton(settings, text="同时导出 PDF", variable=self.pdf, state="normal" if self.converter else "disabled")
        pdf_check.grid(row=0, column=2, padx=self.px(14), pady=(self.px(10), self.px(4)))
        self.controls.append((pdf_check, "normal" if self.converter else "disabled"))
        self._bind_disabled_reason(pdf_check)
        self.pdf_checks = [widget for widget in self.pdf_checks if widget.winfo_exists()]
        self.pdf_checks.append(pdf_check)
        preview = self._button(settings, "预览文件名", self.show_filename_preview)
        preview.grid(row=0, column=3, padx=(0, self.px(14)), pady=(self.px(10), self.px(4)))
        self.naming_buttons = [widget for widget in self.naming_buttons if widget.winfo_exists()]
        self.naming_buttons.append(preview)
        self._label(settings, "输出位置", 9).grid(row=1, column=0, padx=(self.px(14), self.px(10)), pady=self.px(4))
        out = ttk.Entry(settings, textvariable=self.output)
        out.grid(row=1, column=1, sticky="ew", pady=self.px(4))
        self.controls.append((out, "normal"))
        self._button(settings, "选择文件夹", self.choose_output).grid(row=1, column=2, columnspan=2, padx=self.px(14), pady=self.px(4), sticky="ew")
        converter_text = f"PDF：{self.converter[0]} 已检测到" if self.converter else "PDF：需安装 Word 或 LibreOffice"
        self._label(settings, "支持 {{序号}} 和模板字段；每次创建独立批次目录。", 8, MUTED).grid(row=2, column=0, columnspan=2, sticky="w", padx=self.px(14), pady=(self.px(3), self.px(10)))
        self._label(settings, converter_text, 8, MUTED).grid(row=2, column=2, columnspan=2, padx=self.px(14), pady=(self.px(3), self.px(10)), sticky="w")
        return settings, pdf_check

    def _resize_layout(self, event=None):
        if event is not None and event.widget is not self:
            return
        # 留出结果表头和至少两行；100% 缩放同样需要足够的工具栏空间。
        compact = self.winfo_height() < max(840, self.px(790))
        if compact == self.compact_layout:
            return
        self.compact_layout = compact
        self.main.pack_configure(pady=self.px(12 if compact else 22))
        self.header.grid_configure(pady=(0, self.px(10 if compact else 18)))
        self.brand.pack_configure(pady=(self.px(18 if compact else 26), self.px(18 if compact else 28)))
        if compact:
            self.cards.grid_remove()
            self.export_panel.grid_remove()
            self.header_subtitle.pack_forget()
            self.brand_mark.pack_forget()
            self.brand_subtitle.pack_forget()
            self.sidebar_tips.pack_forget()
            self.compact_export_button.pack(side="right", before=self.demo_button, padx=(10, 0))
        else:
            self.cards.grid()
            self.export_panel.grid()
            self.header_subtitle.pack(anchor="w", pady=(self.px(5), 0))
            self.brand_mark.pack(side="left", before=self.brand_text, padx=(0, self.px(12)))
            self.brand_subtitle.pack(anchor="w", pady=(self.px(5), 0))
            self.sidebar_tips.pack(fill="x", before=self.sidebar_tips.master.winfo_children()[-1])
            self.compact_export_button.pack_forget()

    def open_export_settings(self):
        if self.busy:
            return None
        window = getattr(self, "export_window", None)
        if window is not None and window.winfo_exists():
            window.lift()
            window.focus_set()
            return window
        window = self.export_window = tk.Toplevel(self)
        window.title("导出设置")
        window.configure(bg=BG)
        window.transient(self)
        window.geometry(f"{min(self.px(760), self.winfo_screenwidth()-60)}x{self.px(270)}")
        window.resizable(True, False)
        self._label(window, "导出设置", 15, bg=BG, bold=True).pack(anchor="w", padx=self.px(18), pady=(self.px(16), 0))
        panel, _ = self._export_settings(window)
        panel.pack(fill="x", padx=self.px(16), pady=self.px(16))
        ttk.Button(window, text="完成", command=window.destroy, style="Accent.TButton").pack(side="right", padx=self.px(16), pady=(0, self.px(12)))
        self._refresh_ready()
        return window

    def show_filename_preview(self):
        if self.busy:
            return None
        window = getattr(self, "naming_window", None)
        if window is not None and window.winfo_exists():
            self._refresh_filename_preview()
            window.lift()
            return window
        window = self.naming_window = tk.Toplevel(self)
        window.title("生成文件名预览")
        window.configure(bg=BG)
        window.transient(self)
        window.geometry(f"{min(self.px(760), self.winfo_screenwidth()-80)}x{min(self.px(420), self.winfo_screenheight()-100)}")
        ttk.Button(window, text="关闭", command=window.destroy).pack(side="bottom", anchor="e", padx=self.px(16), pady=self.px(12))
        self._label(window, "文件名预览", 15, bg=BG, bold=True).pack(anchor="w", padx=self.px(18), pady=(self.px(16), 0))
        box = tk.Frame(window, bg="white")
        box.pack(fill="both", expand=True, padx=self.px(16), pady=(self.px(16), 0))
        window.preview_text = tk.Text(box, wrap="word", bg="white", fg=INK, font=(FONT, 9), relief="flat",
                                      padx=self.px(16), pady=self.px(16), spacing3=self.px(4), selectbackground=SELECT, selectforeground=INK)
        scroll = ttk.Scrollbar(box, command=window.preview_text.yview)
        scroll.pack(side="right", fill="y")
        window.preview_text.configure(yscrollcommand=scroll.set)
        window.preview_text.pack(fill="both", expand=True)
        self._refresh_filename_preview()
        return window

    def _refresh_filename_preview(self):
        window = getattr(self, "naming_window", None)
        if window is None or not window.winfo_exists():
            return
        try:
            if not self.template or not self.dataset:
                raise InputError("请先选择模板和数据。")
            indices = sorted(self.selected_indices) if self.scope.get() == "选中记录" else None
            count = len(indices) if indices is not None else len(self.dataset.records)
            if not count:
                raise InputError("请先选择记录。")
            items = preview_filenames(self.template, self.dataset, self._mapping(), self.pattern.get(), indices)
            lines = [f"生成{'选中' if indices is not None else '全部'} {count} 份 · 预览前 {len(items)} 份", "",
                     "这是本次新批次的文件名；继续和重试沿用原批次。", ""]
            for item in items:
                lines.append(f"数据行 {item['数据行']}：{item['Word 文件']}")
                if self.pdf.get():
                    lines.append("PDF：" + Path(item["Word 文件"]).with_suffix(".pdf").name)
                if item["说明"]:
                    lines.append("提示：" + item["说明"])
                lines.append("")
            content = "\n".join(lines)
        except InputError as exc:
            content = "暂时无法预览：" + str(exc) + "\n请调整字段匹配、生成范围或文件命名规则。"
        window.preview_text.configure(state="normal")
        window.preview_text.delete("1.0", "end")
        window.preview_text.insert("1.0", content)
        window.preview_text.configure(state="disabled")

    def _settings_changed(self):
        self._schedule_save()
        self._refresh_ready()

    def _mapping_view(self):
        bar = tk.Frame(self.mapping_page, bg="white")
        bar.pack(fill="x", padx=self.px(16), pady=(self.px(12), self.px(8)))
        self._button(bar, "检查数据", self.check_data).pack(side="right")
        self._label(bar, "同名字段自动匹配；手动匹配会保存供下次使用。", 9, MUTED).pack(side="left")
        container = tk.Frame(self.mapping_page, bg="white")
        container.pack(fill="both", expand=True, padx=self.px(16), pady=(0, self.px(14)))
        self.map_canvas = tk.Canvas(container, bg="white", borderwidth=0, highlightthickness=0)
        scroll = ttk.Scrollbar(container, orient="vertical", command=self.map_canvas.yview)
        self.map_canvas.configure(yscrollcommand=scroll.set)
        scroll.pack(side="right", fill="y")
        self.map_canvas.pack(side="left", fill="both", expand=True)
        self.map_inner = tk.Frame(self.map_canvas, bg="white")
        self.map_window = self.map_canvas.create_window((0, 0), window=self.map_inner, anchor="nw")
        self.map_inner.bind("<Configure>", lambda _: self.map_canvas.configure(scrollregion=self.map_canvas.bbox("all")))
        self.map_canvas.bind("<Configure>", lambda event: self.map_canvas.itemconfigure(self.map_window, width=event.width))
        self.map_canvas.bind("<Enter>", lambda _: self.bind_all("<MouseWheel>", self._map_scroll))
        self.map_canvas.bind("<Leave>", lambda _: self.unbind_all("<MouseWheel>"))
        self._draw_mapping()

    def _map_scroll(self, event):
        self.map_canvas.yview_scroll(int(-event.delta / 120), "units")

    def _draw_mapping(self):
        previous = self._mapping()
        self.drawing_mapping = True
        for widget in self.map_inner.winfo_children():
            widget.destroy()
        self.controls = [(w, state) for w, state in self.controls if w.winfo_exists()]
        self.mapping = {}
        if not self.template:
            self._label(self.map_inner, "从一份 Word 模板开始", 15, INK, bold=True).pack(anchor="w", pady=(self.px(28), self.px(10)))
            self._label(self.map_inner, "在 Word 中写入 {{字段名}}，再选择保存后的 .docx 文件。", 9, MUTED, wraplength=self.px(520)).pack(anchor="w")
            self._label(self.map_inner, "也可以点击右上角“载入示例”，体验字段匹配与批量生成。", 9, MUTED, wraplength=self.px(520)).pack(anchor="w", pady=(self.px(7), 0))
            self.drawing_mapping = False
            return
        for column, heading in enumerate(("模板字段", "对应数据列", "首条记录")):
            self._label(self.map_inner, heading, 9, "#486577", bg=SOFT, bold=True).grid(row=0, column=column, sticky="ew", ipady=self.px(8), padx=(0, self.px(12)))
        self.map_inner.columnconfigure(1, weight=1)
        self.map_inner.columnconfigure(2, weight=1)
        headers = self.dataset.headers if self.dataset else []
        saved = self.preferences.profile(self.template.fingerprint).get("mapping", {})
        if not isinstance(saved, dict):
            saved = {}
        for row, field in enumerate(self.template.fields, 1):
            choice = previous.get(field) or saved.get(field) or (field if field in headers else "")
            variable = tk.StringVar(value=choice)
            self.mapping[field] = variable
            self._label(self.map_inner, "{{" + field + "}}", 9, GREEN, bold=True).grid(row=row, column=0, sticky="w", padx=(0, self.px(20)), pady=self.px(8))
            combo = ttk.Combobox(self.map_inner, textvariable=variable, values=["", *headers], state="readonly", width=19)
            combo.grid(row=row, column=1, sticky="ew", padx=(0, self.px(16)), pady=self.px(7))
            self.controls.append((combo, "readonly"))
            sample = self._label(self.map_inner, "", 9, MUTED, wraplength=self.px(220))
            sample.grid(row=row, column=2, sticky="w", pady=self.px(7))
            variable.trace_add("write", lambda *args, value=variable, label=sample: self._mapping_changed(value, label))
            self._mapping_changed(variable, sample)
        self.drawing_mapping = False
        self._refresh_ready()
        self._update_preview()

    def _mapping_changed(self, variable, label):
        value = variable.get()
        text = self.dataset.records[0].get(value, "") if self.dataset else ""
        missing = bool(value and self.dataset and value not in self.dataset.headers)
        label.configure(text="当前数据缺少此列" if missing else (text[:60] or "（空值）") if value else "尚未匹配",
                        fg=MUTED if value and not missing else "#B26C37")
        if self.drawing_mapping:
            return
        self._show_issues([])
        self._refresh_ready()
        self._update_preview()
        self._schedule_save()

    def _tree(self, frame):
        box = tk.Frame(frame, bg="white")
        box.pack(fill="both", expand=True, padx=self.px(16), pady=(0, self.px(14)))
        box.columnconfigure(0, weight=1)
        box.rowconfigure(0, weight=1)
        tree = ttk.Treeview(box, show="headings", selectmode="extended")
        tree.grid(row=0, column=0, sticky="nsew")
        vertical = ttk.Scrollbar(box, orient="vertical", command=tree.yview)
        vertical.grid(row=0, column=1, sticky="ns")
        horizontal = ttk.Scrollbar(box, orient="horizontal", command=tree.xview)
        horizontal.grid(row=1, column=0, sticky="ew")
        tree.configure(yscrollcommand=vertical.set, xscrollcommand=horizontal.set)
        tree.tag_configure("odd", background="#F7F9FB")
        tree.cell_menu = tk.Menu(tree, tearoff=False)
        tree.bind("<Button-3>", lambda event: self._cell_menu(tree, event))
        tree.bind("<Configure>", lambda _: self._fit_table(tree))
        return tree

    def _fit_table(self, tree):
        available = tree.winfo_width() - 4
        columns = [(key, tree.column(key)) for key in tree.cget("columns")]
        fixed = sum(info["width"] for _, info in columns if not info["stretch"])
        flexible = [(key, info) for key, info in columns if info["stretch"]]
        minimum = fixed + sum(info["minwidth"] for _, info in flexible)
        total = sum(info["width"] for _, info in columns)
        if not flexible or available < minimum or total <= available:
            return
        room = total - minimum
        extra = available - minimum
        remaining = extra
        for index, (key, info) in enumerate(flexible):
            share = remaining if index == len(flexible)-1 else int((info["width"]-info["minwidth"]) * extra / room)
            tree.column(key, width=info["minwidth"] + share)
            remaining -= share

    @staticmethod
    def _cell_summary(value, limit=140):
        text = " ".join(str(value).split())
        return text if len(text) <= limit else text[:limit-1] + "…"

    def _cell_menu(self, tree, event):
        if self.busy or tree.identify_region(event.x, event.y) not in ("cell", "heading"):
            return "break"
        row = tree.identify_row(event.y)
        column = tree.identify_column(event.x)
        if not column:
            return "break"
        tree.cell_menu.delete(0, "end")
        tree.cell_menu.add_command(label="查看完整内容", command=lambda: self.show_cell_content(tree, row, column))
        try:
            tree.cell_menu.tk_popup(event.x_root, event.y_root)
        finally:
            tree.cell_menu.grab_release()
        return "break"

    def show_cell_content(self, tree, row, column):
        if row and not tree.exists(row):
            return None
        position = int(column[1:]) - 1
        key = tree.cget("columns")[position]
        heading = tree.heading(key, "text")
        source_row = None
        if tree is self.data_tree:
            heading = "数据行" if position == 0 else self.dataset.headers[position-1]
            if row:
                source_row = self.dataset.row_numbers[int(row)]
                value = str(source_row) if position == 0 else self.dataset.records[int(row)][heading]
            else:
                value = heading
        elif tree is self.result_tree:
            heading = key
            value = heading if not row else self.result_items[int(row)][key]
            if row:
                source_row = self.result_items[int(row)]["数据行"]
        else:
            value = heading if not row else self.issues[int(row)][key]
            if row:
                source_row = self.issues[int(row)]["行号"]
        window = self.cell_window = tk.Toplevel(self)
        window.title("完整内容 · " + self._cell_summary(heading, 40))
        window.configure(bg=BG)
        window.transient(self)
        window.geometry(f"{min(self.px(720), self.winfo_screenwidth()-80)}x{min(self.px(440), self.winfo_screenheight()-100)}")
        self._label(window, "完整内容", 15, bg=BG, bold=True).pack(anchor="w", padx=self.px(18), pady=(self.px(16), self.px(8)))
        context = f"数据第 {source_row} 行" if row else "列标题"
        self._label(window, context + " · " + self._cell_summary(heading, 48), 9, MUTED, BG).pack(anchor="w", padx=self.px(18))
        actions = tk.Frame(window, bg=BG)
        actions.pack(side="bottom", fill="x", padx=self.px(16), pady=self.px(12))
        ttk.Button(actions, text="关闭", command=window.destroy).pack(side="right")
        def copy_content():
            self.clipboard_clear()
            self.clipboard_append(str(value))
            window.copy_button.configure(text="已复制")
        window.copy_button = ttk.Button(actions, text="复制全部", command=copy_content, style="Accent.TButton")
        window.copy_button.pack(side="right", padx=(0, self.px(8)))
        box = tk.Frame(window, bg="white")
        box.pack(fill="both", expand=True, padx=self.px(16), pady=(self.px(12), 0))
        window.content_text = tk.Text(box, wrap="word", bg="white", fg=INK, font=(FONT, 9), relief="flat",
                                      padx=self.px(16), pady=self.px(16), spacing3=self.px(4), selectbackground=SELECT, selectforeground=INK)
        scroll = ttk.Scrollbar(box, command=window.content_text.yview)
        scroll.pack(side="right", fill="y")
        window.content_text.configure(yscrollcommand=scroll.set)
        window.content_text.pack(fill="both", expand=True)
        window.content_text.insert("1.0", str(value))
        window.content_text.configure(state="disabled")
        return window

    def _data_view(self):
        bar = tk.Frame(self.data_page, bg="white")
        bar.pack(fill="x", padx=self.px(16), pady=(self.px(12), self.px(6)))
        self._label(bar, "生成范围", 9, MUTED).pack(side="left", padx=(0, 12))
        for text in ("全部记录", "选中记录"):
            button = ttk.Radiobutton(bar, text=text, variable=self.scope, value=text)
            button.pack(side="left", padx=(0, 12))
            self.controls.append((button, "normal"))
        self._button(bar, "下一页", lambda: self._change_data_page(1)).pack(side="right")
        self._button(bar, "上一页", lambda: self._change_data_page(-1)).pack(side="right", padx=6)
        self.clear_selection_button = self._button(bar, "清空全部选择", self.clear_selection)
        self.clear_selection_button.pack(side="left", padx=(0, 6))
        search = tk.Frame(self.data_page, bg="white")
        search.pack(fill="x", padx=self.px(16), pady=self.px(4))
        search.columnconfigure(1, weight=1)
        self._label(search, "搜索", 9, MUTED).grid(row=0, column=0, padx=(0, 8))
        query = ttk.Entry(search, textvariable=self.data_query, width=12)
        query.grid(row=0, column=1, sticky="ew")
        query.bind("<Return>", lambda _: self._apply_data_search())
        self.controls.append((query, "normal"))
        self._button(search, "搜索", self._apply_data_search).grid(row=0, column=2, padx=6)
        self._button(search, "清空", self._clear_data_search).grid(row=0, column=3)
        self._label(search, "数据行", 9, MUTED).grid(row=0, column=4, padx=(12, 6))
        row = ttk.Entry(search, textvariable=self.locate_row, width=7)
        row.grid(row=0, column=5)
        row.bind("<Return>", lambda _: self._locate_row())
        self.controls.append((row, "normal"))
        self._button(search, "定位", self._locate_row).grid(row=0, column=6, padx=(6, 0))
        self.data_hint = self._label(self.data_page, "导入后可查看数据；按 Ctrl 或 Shift 选择多条记录。", 9, MUTED)
        self.data_hint.pack(fill="x", padx=self.px(16), pady=(self.px(4), self.px(10)))
        self.data_tree = self._tree(self.data_page)
        self.data_tree.bind("<<TreeviewSelect>>", self._data_selection)
        self.data_tree.bind("<Double-1>", lambda _: self.notebook.select(self.preview_page))

    def _check_view(self):
        bar = tk.Frame(self.check_page, bg="white")
        bar.pack(fill="x", padx=self.px(16), pady=self.px(12))
        self._button(bar, "下一页", lambda: self._change_issue_page(1)).pack(side="right")
        self._button(bar, "上一页", lambda: self._change_issue_page(-1)).pack(side="right", padx=6)
        self._button(bar, "重新检查", self.check_data).pack(side="right", padx=6)
        self.export_report_button = self._button(bar, "导出检查报告", self.export_check_report)
        self.export_report_button.pack(side="right", padx=(0, 6))
        self.check_hint = self._label(bar, "生成前会自动检查；双击问题可定位记录。", 9, MUTED)
        self.check_hint.pack(side="left")
        self.check_tree = self._tree(self.check_page)
        columns = ("行号", "字段", "数据列", "级别", "说明")
        self.check_tree.configure(columns=columns)
        for column in columns:
            self.check_tree.heading(column, text=column, anchor="w")
            self.check_tree.column(column, width=self.px(300 if column == "说明" else 100),
                                   minwidth=self.px(180 if column == "说明" else 80), stretch=column == "说明")
        self.check_tree.bind("<Double-1>", self._locate_issue)

    def _preview_view(self):
        bar = tk.Frame(self.preview_page, bg="white")
        bar.pack(fill="x", padx=self.px(16), pady=(self.px(14), self.px(10)))
        self._label(bar, "记录").pack(side="left")
        self.record_spin = ttk.Spinbox(bar, from_=1, to=1, textvariable=self.preview_index, width=6, command=self._update_preview)
        self.record_spin.pack(side="left", padx=(10, 12))
        self.controls.append((self.record_spin, "normal"))
        self.record_spin.bind("<Return>", lambda _: self._update_preview())
        self.record_spin.bind("<FocusOut>", lambda _: self._update_preview())
        self.record_hint = self._label(bar, "", 9, MUTED)
        self.record_hint.pack(side="left")
        self._button(bar, "打开预览 Word", self.open_preview).pack(side="right")
        self.page_preview_button = self._button(bar, "查看排版", self.open_page_preview)
        self.page_preview_button.pack(side="right", padx=8)
        self._label(self.preview_page, "此处为文本预览；“查看排版”显示真实页面，需要可用的 PDF 转换器。", 9, MUTED).pack(fill="x", padx=self.px(16), pady=(0, self.px(10)))
        frame = tk.Frame(self.preview_page, bg="white", highlightbackground=BORDER, highlightthickness=1)
        frame.pack(fill="both", expand=True, padx=self.px(16), pady=(0, self.px(16)))
        self.preview_text = tk.Text(frame, bg="white", fg=INK, font=(FONT, 10), relief="flat", padx=self.px(24), pady=self.px(20), wrap="word", width=1, height=1, spacing3=self.px(6), selectbackground=SELECT, selectforeground=INK, state="disabled")
        self.preview_text.pack(side="left", fill="both", expand=True)
        scroll = ttk.Scrollbar(frame, orient="vertical", command=self.preview_text.yview)
        scroll.pack(side="right", fill="y")
        self.preview_text.configure(yscrollcommand=scroll.set)
        self._set_preview("导入模板和数据并完成字段匹配后，在这里检查填充内容。")

    def _set_preview(self, text):
        if not hasattr(self, "preview_text"):
            return
        self.preview_text.configure(state="normal")
        self.preview_text.delete("1.0", "end")
        self.preview_text.insert("1.0", text)
        self.preview_text.configure(state="disabled")

    def _preview_values(self):
        if not self.template or not self.dataset:
            raise InputError("请先导入模板和数据。")
        try:
            index = int(self.preview_index.get()) - 1
        except ValueError:
            raise InputError("请输入有效的记录序号。")
        if not 0 <= index < len(self.dataset.records):
            raise InputError(f"记录序号应在 1 到 {len(self.dataset.records)} 之间。")
        return values_for(self.template, self._mapping(), self.dataset.records[index])

    def _update_preview(self):
        if not hasattr(self, "preview_text"):
            return
        try:
            values = self._preview_values()
            index = int(self.preview_index.get()) - 1
            self.record_hint.configure(text=f"共 {len(self.dataset.records)} 条 · 数据第 {self.dataset.row_numbers[index]} 行")
            self._set_preview(self.template.preview(values))
        except (InputError, ValueError) as exc:
            self._set_preview(str(exc))

    def _results_view(self):
        self.result_hint = self._label(self.result_page, "生成完成后显示结果；右键可查看单元格完整内容。", 9, MUTED)
        self.result_hint.pack(fill="x", padx=self.px(16), pady=(self.px(10), 0))
        bar = tk.Frame(self.result_page, bg="white")
        bar.pack(fill="x", padx=self.px(16), pady=self.px(6))
        self.resume_button = self._button(bar, "继续剩余记录", self.start_continue)
        self.resume_button.pack(side="left", padx=(0, 8))
        self.retry_button = self._button(bar, "重试失败项", self.start_retry)
        self.retry_button.pack(side="left")
        self.detail_button = self._button(bar, "查看详情", self.show_result_details)
        self.detail_button.pack(side="left", padx=8)
        self._button(bar, "打开输出文件夹", self.open_output).pack(side="right")
        tools = tk.Frame(self.result_page, bg="white")
        tools.pack(fill="x", padx=self.px(16), pady=(0, self.px(6)))
        choice = ttk.Combobox(tools, textvariable=self.result_filter, values=("全部结果", "仅失败和部分成功"), state="readonly", width=18)
        choice.pack(side="left")
        choice.bind("<<ComboboxSelected>>", self._filter_results)
        self.controls.append((choice, "readonly"))
        self.manifest_button = self._button(tools, "打开最新清单", self.open_manifest)
        self.manifest_button.pack(side="right")
        self.pdf_result_button = self._button(tools, "打开 PDF", lambda: self._open_result(kind="PDF 文件"))
        self.pdf_result_button.pack(side="right", padx=6)
        self.word_result_button = self._button(tools, "打开 Word", lambda: self._open_result(kind="Word 文件"))
        self.word_result_button.pack(side="right")
        pager = tk.Frame(self.result_page, bg="white")
        pager.pack(fill="x", padx=self.px(16), pady=(0, self.px(8)))
        self._button(pager, "下一页", lambda: self._change_result_page(1)).pack(side="right")
        self._button(pager, "上一页", lambda: self._change_result_page(-1)).pack(side="right", padx=6)
        self.result_page_hint = self._label(pager, "每页最多显示 200 条结果", 9, MUTED)
        self.result_page_hint.pack(side="left")
        self.result_tree = self._tree(self.result_page)
        self.result_tree.configure(selectmode="browse")
        columns = RESULT_COLUMNS
        self.result_tree.configure(columns=columns)
        for column in columns:
            self.result_tree.heading(column, text="说明（摘要）" if column == "说明" else column, anchor="w")
            flexible = "文件" in column or column == "说明"
            self.result_tree.column(column, width=self.px(220 if flexible else 90),
                                    minwidth=self.px(160 if column == "说明" else 120 if flexible else 90), stretch=flexible)
        self.result_tree.tag_configure("成功", foreground=GREEN)
        self.result_tree.tag_configure("部分成功", foreground="#966123", background="#FFF8EB")
        self.result_tree.tag_configure("失败", foreground="#A03E41", background="#FFF3F2")
        self.result_tree.bind("<Double-1>", self._open_result)
        self.result_tree.bind("<<TreeviewSelect>>", lambda _: self._refresh_ready())

    def _mapping(self):
        return {field: variable.get() for field, variable in self.mapping.items()}

    def _schedule_save(self):
        if self.drawing_mapping:
            return
        if self.save_timer:
            self.after_cancel(self.save_timer)
        self.save_timer = self.after(500, self._save_preferences)

    def _save_preferences(self, recent_file=None, clear_recent=False):
        self.save_timer = None
        try:
            self.preferences.save(self.output.get(), self.pattern.get(), self.pdf.get(),
                                  self.template.fingerprint if self.template and self.dataset else None,
                                  self._mapping(), recent_file=recent_file, clear_recent=clear_recent,
                                  animations=self.motion_enabled.get())
            return True
        except OSError as exc:
            self.status.set("无法保存本地设置：" + str(exc))
            return False

    def _refresh_ready(self):
        if not hasattr(self, "generate_button"):
            return
        total = len(self.template.fields) if self.template else 0
        matched = sum(variable.get() in self.dataset.headers for variable in self.mapping.values()) if self.dataset else 0
        self.metrics[0].configure(text=str(total) if self.template else "—")
        self.metrics[1].configure(text=str(len(self.dataset.records)) if self.dataset else "—")
        self.metrics[2].configure(text=f"{matched} / {total}" if total else "—")
        selected_scope = self.scope.get() == "选中记录"
        count = len(self.selected_indices) if selected_scope else len(self.dataset.records) if self.dataset else 0
        if not self.template:
            input_reason = "请先选择 Word 模板，或点击“载入示例”。"
        elif not self.dataset:
            input_reason = "请先选择 Excel / CSV 数据。"
        elif matched != total:
            input_reason = f"还有 {total-matched} 个字段未匹配，请到“字段匹配”中选择对应数据列。"
        elif not count:
            input_reason = "请到“数据预览”选择记录，或改为“全部记录”。"
        else:
            input_reason = ""
        self.ready_reason = input_reason
        if not self.ready_reason:
            if not self.output.get().strip():
                self.ready_reason = "请在导出设置中填写输出位置。"
            else:
                try:
                    index = min(self.selected_indices) if selected_scope else 0
                    file_stem(self.pattern.get(), values_for(self.template, self._mapping(), self.dataset.records[index]), 1)
                except InputError as exc:
                    self.ready_reason = str(exc) + " 请调整文件命名规则。"
        self.generate_button.configure(text=f"生成{'选中' if selected_scope else '全部'} {count} 份")
        self._availability(self.generate_button, not self.ready_reason, self.ready_reason)
        for button in self.naming_buttons:
            if button.winfo_exists():
                self._availability(button, not input_reason, input_reason)
        pdf_reason = "未检测到 Word 或 LibreOffice；安装后重新打开软件，即可使用 PDF 和排版预览。"
        for widget in self.pdf_checks:
            if widget.winfo_exists():
                self._availability(widget, bool(self.converter), pdf_reason)
        preview_ready = self.template and self.dataset and matched == total
        self._availability(self.page_preview_button, preview_ready and self.converter,
                           pdf_reason if not self.converter else "请先导入模板和数据，并完成字段匹配。")
        retryable = self.last_batch and (any(item["状态"] != "成功" for item in self.last_batch.items)
                                        or (self.last_batch.manifest_warning and self.last_batch.manifest_path is None))
        self._availability(self.retry_button, retryable, "当前没有失败项或需要补保存的清单；请先生成文档。")
        resumable = self.last_batch and self.last_batch.remaining_indices
        self._availability(self.resume_button, resumable, "当前批次没有剩余未处理记录；停止后有剩余记录时可继续。")
        manifest = self.last_batch.manifest_path if self.last_batch else None
        self._availability(self.manifest_button, manifest and manifest.is_file(), "当前批次清单尚未保存，或文件已被移动。")
        selected = self.result_tree.selection()
        self._availability(self.detail_button, selected, "请先在结果列表中选择一条记录。")
        item = self.result_items[int(selected[0])] if selected else {}
        for button, key in ((self.word_result_button, "Word 文件"), (self.pdf_result_button, "PDF 文件")):
            available = self.last_folder and item.get(key) and (self.last_folder / item[key]).is_file()
            self._availability(button, available, "请先选择一条结果。" if not selected else f"所选记录的{key}尚未生成，或文件已被移动；可查看详情。")
        if not self.data_path:
            self.sheet_combo.configure(state="disabled")
        self._availability(self.refresh_button, self.data_path, "请先选择 Excel / CSV 数据，再刷新已保存的改动。")
        report_reason = "检查完成，未发现问题，无需导出。" if self.check_completed else "请先检查数据，再导出发现的问题。"
        self._availability(self.export_report_button, self.issues and self.data_path, report_reason)
        window = getattr(self, "recent_window", None)
        if window is not None and window.winfo_exists():
            self._availability(window.open_button, window.file_tree.selection(), "请选择一个最近使用的文件。")
        self._refresh_filename_preview()
        self._update_activity()

    @staticmethod
    def _duration(seconds):
        seconds = max(0, int(seconds))
        minutes, seconds = divmod(seconds, 60)
        return f"{minutes} 分 {seconds} 秒" if minutes else f"{seconds} 秒"

    def _update_activity(self):
        if self.busy and self.operation_started is not None:
            elapsed = max(0, time.monotonic() - self.operation_started)
            phase = self.operation_phase or "正在处理"
            parts = [("正在停止 · " if self.cancel_event.is_set() else "") + phase]
            if self.operation_total:
                parts.append(f"{self.operation_current} / {self.operation_total}")
            parts.append("已用 " + self._duration(elapsed))
            if not self.cancel_event.is_set() and self.operation_completed >= 2 and self.operation_total > self.operation_completed and self.operation_phase != "保存清单":
                remaining = (self.operation_sample_elapsed / self.operation_completed) * (self.operation_total-self.operation_completed) - (elapsed-self.operation_sample_elapsed)
                parts.append("预计剩余约 " + self._duration(max(1, remaining)) if remaining > 0 else "剩余时间重新估算中")
            content = " · ".join(parts)
        elif self.busy:
            content = "正在处理，请等待当前操作结束。"
        else:
            content = self.ready_reason or "准备就绪，可先预览文件名再生成。"
            if self.last_operation_elapsed is not None:
                content += " · 上次操作耗时 " + self._duration(self.last_operation_elapsed)
        if content != self.action_hint.get():
            self.action_hint.set(content)

    def _report_phase(self, phase, current, total):
        self.events.put(("phase", phase, current, total))

    def _set_busy(self, busy):
        self.busy = busy
        self._hide_disabled_reason()
        for widget, state in self.controls:
            if widget.winfo_exists():
                widget.configure(state="disabled" if busy else state)
        self._refresh_ready()

    def _run(self, function, done, label, generation=False):
        if self.busy:
            return
        self.progress_motion.cancel()
        self.operation_started = time.monotonic()
        self.operation_phase = label.rstrip("…")
        self.operation_current = self.operation_total = self.operation_completed = 0
        self.operation_sample_elapsed = 0
        self._set_busy(True)
        self.status.set(label)
        self.cancel_event.clear()
        if generation:
            self.cancel_button.configure(state="normal")
            self.progressbar.configure(mode="determinate", value=0)
        else:
            self.progressbar.configure(mode="indeterminate")
            self.progressbar.start(12)
        def work():
            try:
                self.events.put(("done", done, function()))
            except Exception as exc:
                self.events.put(("error", str(exc), traceback.format_exc()))
        threading.Thread(target=work, daemon=True).start()

    def _poll(self):
        filtered_results_changed = False
        try:
            for _ in range(100):
                event = self.events.get_nowait()
                if event[0] == "progress":
                    _, number, total, item, slot = event
                    self.operation_completed, self.operation_total = number, total
                    self.operation_current = number
                    if self.operation_started is not None:
                        self.operation_sample_elapsed = max(0, time.monotonic() - self.operation_started)
                    self._animate_progress(number * 100 / total)
                    self.status.set(f"正在处理 {number} / {total} · 数据第 {item['数据行']} 行")
                    if slot < len(self.result_items):
                        self.result_items[slot] = item
                    else:
                        self.result_items.append(item)
                    iid = str(slot)
                    if self.result_filter.get() != "全部结果":
                        filtered_results_changed = True
                    elif self.result_page_number * PAGE_SIZE <= slot < (self.result_page_number + 1) * PAGE_SIZE:
                        values = [item[column] for column in RESULT_COLUMNS]
                        if self.result_tree.exists(iid):
                            self.result_tree.item(iid, values=values)
                        else:
                            self.result_tree.insert("", "end", iid=iid, values=values, tags=("odd",) if slot % 2 else ())
                elif event[0] == "phase":
                    _, self.operation_phase, self.operation_current, self.operation_total = event
                else:
                    self.progressbar.stop()
                    self.progress_motion.cancel()
                    self.progressbar.configure(mode="determinate")
                    self.cancel_button.configure(state="disabled")
                    if self.operation_started is not None:
                        self.last_operation_elapsed = max(0, time.monotonic() - self.operation_started)
                    self.operation_started = None
                    self._set_busy(False)
                    if self.close_requested:
                        self._finish_close()
                        return
                    if event[0] == "done":
                        try:
                            event[1](event[2])
                        except Exception as exc:
                            self.status.set("操作失败：" + str(exc))
                            messagebox.showerror("操作未完成", str(exc), parent=self)
                    else:
                        self.status.set("操作失败：" + event[1])
                        messagebox.showerror("操作未完成", event[1], parent=self)
        except queue.Empty:
            pass
        if filtered_results_changed:
            self._fill_results()
        self._update_activity()
        self.after(100, self._poll)

    def choose_template(self):
        path = filedialog.askopenfilename(parent=self, title="选择 Word 模板", filetypes=[("Word 模板", "*.docx")])
        if path:
            self._open_template(Path(path))

    def _open_template(self, path):
        def read():
            if not path.is_file():
                raise InputError("模板文件不存在或已移动，请重新选择：" + str(path))
            return Template(path)
        self._run(read, self._template_loaded, "正在读取模板…")

    def _template_loaded(self, template):
        self._save_preferences()
        self.template = template
        self._show_issues([])
        self.mapping = {}
        self.template_label.configure(text=template.path.name)
        self._draw_mapping()
        candidate = next((field for field in template.fields if field in ("姓名", "名称", "客户名称", "编号")), None)
        saved_pattern = self.preferences.profile(template.fingerprint).get("pattern")
        self.pattern.set(saved_pattern if isinstance(saved_pattern, str) and saved_pattern else "{{序号}}_{{" + candidate + "}}" if candidate else "{{序号}}_文档")
        self.status.set(f"已识别 {len(template.fields)} 个字段。请选择数据文件或检查字段匹配。")
        self._refresh_ready()
        self._update_preview()
        self._save_preferences(recent_file=("template", template.path))

    def choose_data(self):
        path = filedialog.askopenfilename(parent=self, title="选择数据文件", filetypes=[("Excel / CSV", "*.xlsx *.csv"), ("Excel", "*.xlsx"), ("CSV", "*.csv")])
        if path:
            self._open_data(Path(path))

    @staticmethod
    def _read_source(source, sheet=None):
        if not source.is_file():
            raise InputError("数据文件不存在或已移动，请重新选择：" + str(source))
        sheets = sheet_names(source)
        if sheet and sheet not in sheets:
            raise InputError(f"工作表“{sheet}”已不存在，请重新选择数据文件和工作表。")
        return source, sheets, read_data(source, sheet)

    def _open_data(self, source):
        self._run(lambda: self._read_source(source), self._data_loaded, "正在导入数据…")

    def _data_loaded(self, result):
        path, sheets, dataset = result
        self.data_path, self.dataset = path, dataset
        self._show_issues([])
        self.data_label.configure(text=path.name)
        self.sheet_combo.configure(values=sheets, state="readonly")
        self.sheet.set(dataset.sheet)
        self.preview_index.set("1")
        self.record_spin.configure(to=len(dataset.records))
        self._draw_mapping()
        self.data_page_number = 0
        self.selected_indices.clear()
        self._fill_data(reset=True)
        self._refresh_ready()
        self._update_preview()
        blanks = sum(any(not value.strip() for value in row.values()) for row in dataset.records)
        self.status.set(f"已导入 {len(dataset.records)} 条记录、{len(dataset.headers)} 列。" + (f"其中 {blanks} 条包含空值，生成时按空白填充，请先检查。" if blanks else "同名字段已自动匹配。"))
        return self._save_preferences(recent_file=("data", path))

    def reload_sheet(self):
        if self.data_path and not self.busy:
            source, sheet = self.data_path, self.sheet.get()
            self._run(lambda: self._read_source(source, sheet), self._data_loaded, "正在读取工作表…")

    def refresh_data(self):
        if not self.data_path or self.busy:
            return
        source, sheet = self.data_path, self.dataset.sheet
        previous_headers = list(self.dataset.headers)
        def refreshed(result):
            saved = self._data_loaded(result)
            headers = self.dataset.headers
            added = [name for name in headers if name not in previous_headers]
            removed = [name for name in previous_headers if name not in headers]
            changes = (["新增列：" + "、".join(added)] if added else []) + (["删除列：" + "、".join(removed)] if removed else [])
            note = "列结构已变化（" + "；".join(changes) + "），请检查字段匹配。" if changes else "字段匹配已保留。"
            if saved:
                self.status.set(f"已刷新 {len(self.dataset.records)} 条记录、{len(headers)} 列，旧选择、搜索和检查已清空。" + note)
        self._run(lambda: self._read_source(source, sheet), refreshed, "正在刷新数据…")

    def open_recent_files(self):
        if self.busy:
            return None
        window = getattr(self, "recent_window", None)
        if window is not None and window.winfo_exists():
            window.lift()
            return window
        try:
            entries = [(kind, path) for kind in ("template", "data") for path in self.preferences.recent_files(kind)]
        except OSError as exc:
            messagebox.showerror("无法读取最近文件", str(exc), parent=self)
            return None
        window = self.recent_window = tk.Toplevel(self)
        window.title("最近使用文件")
        window.configure(bg=BG)
        window.transient(self)
        window.geometry(f"{min(self.px(820), self.winfo_screenwidth()-80)}x{min(self.px(420), self.winfo_screenheight()-100)}")
        self._label(window, "最近使用文件", 15, bg=BG, bold=True).pack(anchor="w", padx=self.px(18), pady=(self.px(16), self.px(6)))
        self._label(window, "模板和数据各保留 5 个；只在本机记录路径，打开时重新读取文件。", 9, MUTED, BG).pack(anchor="w", padx=self.px(18))
        actions = tk.Frame(window, bg=BG)
        actions.pack(side="bottom", fill="x", padx=self.px(16), pady=self.px(12))
        ttk.Button(actions, text="关闭", command=window.destroy).pack(side="right")
        def open_selected(event=None):
            selected = window.file_tree.selection()
            if not selected or self.busy:
                return
            kind, path = entries[int(selected[0])]
            window.destroy()
            self._open_template(Path(path)) if kind == "template" else self._open_data(Path(path))
        window.open_button = self._button(actions, "打开所选", open_selected, "Accent.TButton")
        window.open_button.pack(side="right", padx=self.px(8))
        def clear():
            if not self.busy and self._save_preferences(clear_recent=True):
                window.file_tree.delete(*window.file_tree.get_children())
                self.status.set("最近文件记录已清空。")
                self._refresh_ready()
        window.clear_button = self._button(actions, "清空记录", clear)
        window.clear_button.pack(side="left")
        box = tk.Frame(window, bg="white")
        box.pack(fill="both", expand=True, padx=self.px(16), pady=(self.px(14), 0))
        box.columnconfigure(0, weight=1)
        box.rowconfigure(0, weight=1)
        tree = window.file_tree = ttk.Treeview(box, columns=("kind", "name", "path"), show="headings", selectmode="browse", height=1)
        tree.grid(row=0, column=0, sticky="nsew")
        for column, title, width in (("kind", "类型", 70), ("name", "文件名", 200), ("path", "完整路径", 470)):
            tree.heading(column, text=title, anchor="w")
            content_width = max([self.px(width), *[self.table_font.measure(Path(path).name if column == "name" else path) + self.px(28)
                                                for _, path in entries if column != "kind"]])
            tree.column(column, width=content_width, minwidth=content_width, stretch=column == "path")
        for index, (kind, path) in enumerate(entries):
            tree.insert("", "end", iid=str(index), values=("模板" if kind == "template" else "数据", Path(path).name, path))
        vertical = ttk.Scrollbar(box, command=tree.yview)
        vertical.grid(row=0, column=1, sticky="ns")
        horizontal = ttk.Scrollbar(box, orient="horizontal", command=tree.xview)
        horizontal.grid(row=1, column=0, sticky="ew")
        tree.configure(yscrollcommand=vertical.set, xscrollcommand=horizontal.set)
        tree.bind("<<TreeviewSelect>>", lambda _: self._refresh_ready())
        tree.bind("<Double-1>", open_selected)
        self._refresh_ready()
        return window

    def _fill_data(self, reset=False):
        if not self.dataset:
            return
        if reset:
            self.data_indices = list(range(len(self.dataset.records)))
            self.data_search_text = ""
            self.data_query.set("")
            self.locate_row.set("")
        self.filling_data = True
        self.data_tree.delete(*self.data_tree.get_children())
        columns = ["row", *(f"c{i}" for i in range(len(self.dataset.headers)))]
        self.data_tree.configure(columns=columns)
        start = self.data_page_number * PAGE_SIZE
        visible = self.data_indices[start:start + PAGE_SIZE]
        for column, name in zip(columns, ["数据行", *self.dataset.headers]):
            self.data_tree.heading(column, text=self._cell_summary(name), anchor="w")
            if column == "row":
                self.data_tree.column(column, width=self.px(74), minwidth=self.px(74), stretch=False)
            else:
                widths = [self.table_font.measure(self._cell_summary(name))]
                widths.extend(self.table_font.measure(self._cell_summary(self.dataset.records[index][name])) for index in visible)
                width = max(self.px(80), min(self.px(280), max(widths) + self.px(28)))
                self.data_tree.column(column, width=width, minwidth=self.px(80), stretch=True)
        for index in visible:
            number, record = self.dataset.row_numbers[index], self.dataset.records[index]
            self.data_tree.insert("", "end", iid=str(index), values=[number, *(self._cell_summary(value) for value in record.values())], tags=("odd",) if index % 2 else ())
        self.data_tree.selection_set([str(index) for index in visible if index in self.selected_indices])
        self.after_idle(lambda: self._fit_table(self.data_tree))
        self._data_page_hint()
        self.after_idle(lambda: setattr(self, "filling_data", False))

    def _data_page_hint(self):
        self._refresh_ready()
        if self.dataset:
            pages = max(1, (len(self.data_indices) + PAGE_SIZE - 1) // PAGE_SIZE)
            total = len(self.dataset.records)
            count = f"匹配 {len(self.data_indices)} / {total} 条" if self.data_search_text else f"共 {total} 条"
            self.data_hint.configure(text=f"第 {self.data_page_number + 1} / {pages} 页 · {count} · 已选 {len(self.selected_indices)} 条，跨页保留选择 · 右键查看完整内容")

    def _apply_data_search(self):
        if not self.dataset or self.busy:
            return
        self._capture_selection()
        self.data_search_text = self.data_query.get().strip().casefold()
        query = self.data_search_text
        self.data_indices = [index for index, record in enumerate(self.dataset.records)
                             if not query or query in str(self.dataset.row_numbers[index])
                             or any(query in value.casefold() for value in record.values())]
        self.data_page_number = 0
        self._fill_data()
        self.status.set(f"搜索匹配 {len(self.data_indices)} 条。搜索只筛选显示，生成仍按“全部记录”或“选中记录”执行。")

    def _clear_data_search(self):
        if self.busy:
            return
        self.data_query.set("")
        self._apply_data_search()

    def _locate_row(self):
        if not self.dataset or self.busy:
            return
        try:
            row = int(self.locate_row.get().strip())
            index = self.dataset.row_numbers.index(row)
        except ValueError:
            self.status.set("请输入已导入的原始数据行号；标题行和完全空白行不会作为记录导入。")
            return
        self._locate_data_index(index)

    def _locate_data_index(self, index):
        self._capture_selection()
        cleared = index not in self.data_indices
        if cleared:
            self.data_query.set("")
            self.data_search_text = ""
            self.data_indices = list(range(len(self.dataset.records)))
        self.selected_indices.add(index)
        self.data_page_number = self.data_indices.index(index) // PAGE_SIZE
        self._fill_data()
        self.data_tree.focus(str(index))
        self.data_tree.see(str(index))
        self.preview_index.set(str(index + 1))
        self._update_preview()
        self.notebook.select(self.data_page)
        self.status.set(f"已定位并选中数据第 {self.dataset.row_numbers[index]} 行。" + ("已清空搜索。" if cleared else ""))

    def _capture_selection(self):
        if self.filling_data:
            return
        previous = set(self.selected_indices)
        visible = {int(iid) for iid in self.data_tree.get_children()}
        self.selected_indices.difference_update(visible)
        self.selected_indices.update(int(iid) for iid in self.data_tree.selection())
        if previous != self.selected_indices and self.scope.get() == "选中记录":
            self._show_issues([])
        self._data_page_hint()

    def clear_selection(self):
        if self.busy:
            return
        self.selected_indices.clear()
        if self.scope.get() == "选中记录":
            self._show_issues([])
        self.data_tree.selection_remove(*self.data_tree.selection())
        self._data_page_hint()
        self.status.set("已清空全部选择，包括其他页和被搜索隐藏的记录。")

    def _change_data_page(self, offset):
        if not self.dataset or self.busy:
            return
        self._capture_selection()
        last = max(0, (len(self.data_indices) - 1) // PAGE_SIZE)
        self.data_page_number = max(0, min(last, self.data_page_number + offset))
        self._fill_data()

    def _data_selection(self, event=None):
        if self.filling_data:
            return
        self._capture_selection()
        selected = self.data_tree.selection()
        if selected:
            self.preview_index.set(str(int(selected[0]) + 1))
            self._update_preview()

    def _fill_results(self):
        self.result_tree.delete(*self.result_tree.get_children())
        slots = self._result_slots()
        pages = max(1, (len(slots) + PAGE_SIZE - 1) // PAGE_SIZE)
        self.result_page_number = min(self.result_page_number, pages - 1)
        start = self.result_page_number * PAGE_SIZE
        for slot in slots[start:start + PAGE_SIZE]:
            item = self.result_items[slot]
            tags = (item["状态"], "odd") if slot % 2 else (item["状态"],)
            self.result_tree.insert("", "end", iid=str(slot), values=[self._cell_summary(item[column]) for column in RESULT_COLUMNS], tags=tags)
        self.result_page_hint.configure(text=f"第 {self.result_page_number + 1} / {pages} 页 · 显示 {len(slots)} / {len(self.result_items)} 条结果")
        self._refresh_ready()

    def _result_slots(self):
        return [slot for slot, item in enumerate(self.result_items)
                if self.result_filter.get() == "全部结果" or item["状态"] != "成功"]

    def _filter_results(self, event=None):
        self.result_page_number = 0
        self._fill_results()

    def _change_result_page(self, offset):
        last = max(0, (len(self._result_slots()) - 1) // PAGE_SIZE)
        self.result_page_number = max(0, min(last, self.result_page_number + offset))
        self._fill_results()

    def _show_issues(self, issues, completed=None):
        self.issues = issues
        self.check_completed = bool(issues) if completed is None else completed
        self.issue_page_number = 0
        self._fill_issues()

    def _fill_issues(self):
        self.check_tree.delete(*self.check_tree.get_children())
        start = self.issue_page_number * PAGE_SIZE
        for slot in range(start, min(start + PAGE_SIZE, len(self.issues))):
            issue = self.issues[slot]
            self.check_tree.insert("", "end", iid=str(slot), values=[self._cell_summary(issue[column]) for column in ("行号", "字段", "数据列", "级别", "说明")])
        errors = sum(issue["级别"] == "错误" for issue in self.issues)
        pages = max(1, (len(self.issues) + PAGE_SIZE - 1) // PAGE_SIZE)
        text = f"{errors} 项错误、{len(self.issues)-errors} 项提示 · 第 {self.issue_page_number + 1} / {pages} 页"
        self.check_hint.configure(text=text if self.issues else "检查完成，未发现问题。" if self.check_completed else "尚未检查，请点击“重新检查”。")
        self._refresh_ready()

    def _change_issue_page(self, offset):
        last = max(0, (len(self.issues) - 1) // PAGE_SIZE)
        self.issue_page_number = max(0, min(last, self.issue_page_number + offset))
        self._fill_issues()

    def _locate_issue(self, event=None):
        selected = self.check_tree.selection()
        if not selected or not self.dataset:
            return
        index = self.issues[int(selected[0])]["index"]
        self._locate_data_index(index)

    def check_data(self):
        if not self.template or not self.dataset:
            self.status.set("请先导入模板和数据并完成字段匹配。")
            return
        template, dataset, mapping, pattern = self.template, self.dataset, self._mapping(), self.pattern.get()
        def done(issues):
            self._show_issues(issues, completed=True)
            self.notebook.select(self.check_page)
            self.status.set("检查完成：错误需要先修正；空值按空白填充，重名文件自动编号。")
        self._run(lambda: check_records(template, dataset, mapping, pattern), done, "正在检查各条记录…")

    def export_check_report(self):
        if self.busy or not self.issues or not self.data_path:
            return
        path = filedialog.asksaveasfilename(parent=self, title="导出全部检查结果", defaultextension=".csv",
                                          filetypes=[("CSV 检查报告", "*.csv")], initialdir=str(self.data_path.parent),
                                          initialfile="数据检查报告_" + time.strftime("%Y%m%d_%H%M%S") + ".csv")
        if not path:
            return
        destination = Path(path)
        try:
            if destination.resolve() in (self.data_path.resolve(), self.template.path.resolve()):
                raise InputError("检查报告不能覆盖源文件，请选择其他文件名。")
        except (InputError, OSError) as exc:
            messagebox.showerror("无法导出检查报告", str(exc), parent=self)
            return
        source, sheet, issues = self.data_path, self.dataset.sheet, list(self.issues)
        def done(saved):
            self.status.set(f"已导出全部 {len(issues)} 项检查结果：{saved}")
        self._run(lambda: write_check_report(destination, source, sheet, issues), done, "正在导出检查报告…")

    def load_demo(self):
        folder = self.base / "示例文件"
        if not (folder / "通知模板.docx").is_file():
            folder = resource("示例文件")
        if not (folder / "通知模板.docx").is_file():
            messagebox.showerror("示例缺失", "未找到示例文件，请保留随软件附带的示例文件目录。", parent=self)
            return
        def load():
            source = folder / "人员数据.xlsx"
            return Template(folder / "通知模板.docx"), (source, sheet_names(source), read_data(source))
        def done(result):
            self._template_loaded(result[0])
            self._data_loaded(result[1])
            self.status.set("示例已载入：3 条人员记录。可查看预览，或直接开始批量生成。")
        self._run(load, done, "正在载入示例…")

    def choose_output(self):
        path = filedialog.askdirectory(parent=self, title="选择输出文件夹")
        if path:
            self.output.set(path)

    def open_preview(self):
        try:
            values = self._preview_values()
            folder = self._new_preview_folder("word-")
            path = folder / "文档预览.docx"
            self.template.render(values, path)
            os.startfile(str(path))
            self.status.set("已打开预览 Word，文件保存在本软件的本地预览缓存中。")
        except Exception as exc:
            messagebox.showerror("无法预览", str(exc), parent=self)

    def _new_preview_folder(self, prefix):
        self.preview_cache.ensure()
        folder = Path(tempfile.mkdtemp(prefix=prefix, dir=self.preview_root))
        self.preview_folders.add(folder)
        return folder

    def _remove_preview(self, folder):
        folder = Path(folder)
        if folder.is_symlink() or folder.is_junction() or folder.resolve().parent != self.preview_root.resolve():
            return
        try:
            shutil.rmtree(folder)
            self.preview_folders.discard(folder)
        except OSError:
            pass

    def _cleanup_previews(self, schedule=True):
        active = {getattr(window, "folder", None) for window in self.preview_windows if window.winfo_exists()}
        active.update(self.active_preview_folders)
        if self.preview_root.is_dir():
            for folder in self.preview_root.iterdir():
                if not folder.is_dir() or folder in active or not folder.name.startswith(("word-", "page-")):
                    continue
                try:
                    if time.time() - folder.stat().st_mtime < 30 or any(path.name.startswith("~$") for path in folder.iterdir()):
                        continue
                    self._remove_preview(folder)
                except OSError:
                    pass
        if schedule:
            self.after(30000, self._cleanup_previews)
        self.preview_cache.cleanup_stale()

    def open_page_preview(self):
        try:
            values, template = self._preview_values(), self.template
            if not self.converter:
                raise InputError("排版预览需要 Microsoft Word 或 LibreOffice。")
        except InputError as exc:
            messagebox.showinfo("无法查看排版", str(exc), parent=self)
            return
        def render():
            folder = self._new_preview_folder("page-")
            self.active_preview_folders.add(folder)
            try:
                source = folder / "文档预览.docx"
                template.render(values, source)
                with PdfSession(self.converter, cancel=self.cancel_event) as converter:
                    pdf = converter.convert(source)
                return pdf, folder
            except Exception:
                self.active_preview_folders.discard(folder)
                self._remove_preview(folder)
                raise
        def done(result):
            pdf, folder = result
            try:
                from page_preview import PagePreview
                window = PagePreview(self, pdf, lambda: self._remove_preview(folder))
                window.folder = folder
                self.preview_windows.append(window)
                self.status.set("排版预览已打开；关闭预览窗口后自动清理本次预览文件。")
            except Exception:
                self._remove_preview(folder)
                raise
            finally:
                self.active_preview_folders.discard(folder)
        self._run(render, done, "正在生成真实页面预览…", generation=True)

    def start_generation(self):
        if self.busy or not self.dataset:
            return
        mapping = self._mapping()
        template, dataset = self.template, self.dataset
        if not self.output.get().strip():
            messagebox.showerror("缺少输出位置", "请选择输出文件夹。", parent=self)
            return
        indices = None
        if self.scope.get() == "选中记录":
            self._capture_selection()
            indices = sorted(self.selected_indices)
            if not indices:
                messagebox.showinfo("请选择记录", "请先在“数据预览”中选择要生成的记录。", parent=self)
                self.notebook.select(self.data_page)
                return
        output, pattern, pdf = Path(self.output.get().strip()), self.pattern.get(), self.pdf.get()
        def checked(issues):
            self._show_issues(issues, completed=True)
            if any(issue["级别"] == "错误" for issue in issues):
                self.notebook.select(self.check_page)
                self.status.set("数据检查发现错误，请根据行号和字段修正源文件后重新导入。")
                return
            self.last_job = (template, dataset, mapping, pattern)
            self.last_batch = None
            self.result_items = []
            self.result_page_number = 0
            self.result_filter.set("全部结果")
            self.result_tree.delete(*self.result_tree.get_children())
            self.notebook.select(self.result_page)
            self.result_hint.configure(text="正在生成，请等待…")
            def progress(number, total, item):
                self.events.put(("progress", number, total, item, number-1))
            self._run(lambda: generate_batch(template, dataset, mapping, output, pattern, indices, pdf, self.cancel_event, progress,
                                            phase=self._report_phase), self._batch_done, "正在准备生成…", generation=True)
        self._save_preferences()
        self._run(lambda: check_records(template, dataset, mapping, pattern, indices), checked, "生成前正在检查数据…")

    def start_retry(self):
        if self.busy or not self.last_batch or not self.last_job:
            return
        previous = self.last_batch
        template, dataset, mapping, pattern = self.last_job
        slots = {item["数据行"]: slot for slot, item in enumerate(previous.items)}
        def progress(number, total, item):
            self.events.put(("progress", number, total, item, slots[item["数据行"]]))
        self.notebook.select(self.result_page)
        self._run(lambda: retry_batch(previous, template, dataset, mapping, pattern, self.cancel_event, progress, phase=self._report_phase),
                  self._batch_done, "正在重试失败部分…", generation=True)

    def start_continue(self):
        if self.busy or not self.last_batch or not self.last_job or not self.last_batch.remaining_indices:
            return
        previous = self.last_batch
        template, dataset, mapping, pattern = self.last_job
        initial_count = len(previous.items)
        def progress(number, total, item):
            self.events.put(("progress", number, total, item, initial_count + number - 1))
        self.notebook.select(self.result_page)
        self._run(lambda: continue_batch(previous, template, dataset, mapping, pattern, self.cancel_event, progress, phase=self._report_phase),
                  self._batch_done, f"正在继续剩余 {len(previous.remaining_indices)} 条记录…", generation=True)

    def _batch_done(self, result):
        self.last_folder = result.folder
        self.last_batch = result
        self.result_items = result.items
        self._fill_results()
        failed = len(result.items) - result.succeeded - result.partial
        text = ("已停止" if result.cancelled else "处理完成") + f"：成功 {result.succeeded}，部分成功 {result.partial}，失败 {failed}，剩余 {len(result.remaining_indices)}。"
        self.result_hint.configure(text=text)
        self.status.set(text + (result.manifest_warning or "结果和生成清单已保存，可点击“打开输出文件夹”。"))
        count = len(result.items) + len(result.remaining_indices)
        self._animate_progress(len(result.items) * 100 / count if count else 0)
        self._refresh_ready()

    def open_output(self):
        folder = self.last_folder or Path(self.output.get().strip())
        if folder.is_dir():
            os.startfile(str(folder))
        else:
            messagebox.showinfo("尚未生成", "生成后即可打开输出文件夹。", parent=self)

    def _open_result(self, event=None, kind=None):
        selected = self.result_tree.selection()
        if selected and self.last_folder:
            item = self.result_items[int(selected[0])]
            name = item.get(kind, "") if kind else item["Word 文件"] or item["PDF 文件"]
            if name and (self.last_folder / name).is_file():
                try:
                    os.startfile(str(self.last_folder / name))
                except OSError as exc:
                    messagebox.showerror("无法打开文件", str(exc), parent=self)
            else:
                self.status.set("该记录尚无所选文件，或文件已被移动。")

    def show_result_details(self):
        selected = self.result_tree.selection()
        if not selected:
            self.status.set("请先选择一条结果，再查看详情。")
            return None
        item = self.result_items[int(selected[0])]
        folder = self.last_folder.absolute() if self.last_folder else None
        lines = [f"数据行：{item['数据行']}", f"状态：{item['状态']}"]
        for key in ("Word 文件", "PDF 文件"):
            name = item.get(key, "")
            path = str(folder / name) if name and folder else name
            lines.append(f"{key}：{path or '未生成'}")
        lines.extend((f"批次目录：{folder or '暂无'}", "", "说明：", item.get("说明", "") or "无"))
        content = "\n".join(lines)
        window = tk.Toplevel(self)
        window.title(f"结果详情 · 数据第 {item['数据行']} 行")
        window.configure(bg=BG)
        window.transient(self)
        window.geometry(f"{min(self.px(760), self.winfo_screenwidth()-80)}x{min(self.px(480), self.winfo_screenheight()-100)}")
        self._label(window, "结果详情", 15, bg=BG, bold=True).pack(anchor="w", padx=self.px(18), pady=(self.px(16), 0))
        box = tk.Frame(window, bg="white")
        box.pack(fill="both", expand=True, padx=self.px(16), pady=self.px(16))
        window.detail_text = tk.Text(box, wrap="word", font=(FONT, 9), bg="white", fg=INK, relief="flat",
                                     padx=self.px(16), pady=self.px(16), spacing3=self.px(6), selectbackground=SELECT, selectforeground=INK)
        scroll = ttk.Scrollbar(box, command=window.detail_text.yview)
        scroll.pack(side="right", fill="y")
        window.detail_text.configure(yscrollcommand=scroll.set)
        window.detail_text.pack(fill="both", expand=True)
        window.detail_text.insert("1.0", content)
        window.detail_text.configure(state="disabled")
        actions = tk.Frame(window, bg=BG)
        actions.pack(side="bottom", fill="x", before=box, padx=self.px(16), pady=(0, self.px(12)))
        ttk.Button(actions, text="关闭", command=window.destroy).pack(side="right")
        def copy_details():
            self.clipboard_clear()
            self.clipboard_append(content)
            window.copy_button.configure(text="已复制")
        window.copy_button = ttk.Button(actions, text="复制全部", command=copy_details, style="Accent.TButton")
        window.copy_button.pack(side="right", padx=(0, self.px(8)))
        return window

    def open_manifest(self):
        path = self.last_batch.manifest_path if self.last_batch else None
        if path and path.is_file():
            try:
                os.startfile(str(path))
            except OSError as exc:
                messagebox.showerror("无法打开清单", str(exc), parent=self)
        else:
            self.status.set("最新清单尚未保存，或文件已被移动。")

    def show_help(self):
        window = tk.Toplevel(self)
        window.title("使用说明")
        window.geometry(f"{min(self.px(660), self.winfo_screenwidth()-80)}x{min(self.px(640), self.winfo_screenheight()-100)}")
        window.configure(bg=BG)
        window.transient(self)
        ttk.Button(window, text="知道了", command=window.destroy, style="Accent.TButton").pack(side="bottom", anchor="e", padx=self.px(24), pady=self.px(18))
        self._label(window, "四步生成文档", 17, bg=BG, bold=True).pack(anchor="w", padx=self.px(24), pady=(self.px(22), self.px(14)))
        help_text = (
            "1. 准备 Word 模板\n在需要替换的位置写 {{姓名}}、{{部门}} 等占位符，保存为 .docx。支持正文、表格、页眉页脚中的普通文字字段。\n\n"
            "2. 准备 Excel 或 CSV\n第一行填写列标题，每行填写一条记录。支持 .xlsx、UTF-8 / GB18030 CSV。编号的前导零应保存为文本或设置为 000 等数字格式。修改后先保存文件，再点击刷新数据；成功刷新清空旧选择、搜索和检查，保留匹配。最近文件可重新打开已用的模板和数据，也可清空路径记录。\n\n"
            "3. 检查匹配和预览\n同名字段自动匹配；手动映射会保留供下次使用。可搜索关键词、定位原始数据行号，跨页保留选择，点击清空全部选择可一次取消。搜索只筛选显示；生成范围仍为全部或选中记录。检查问题可双击定位，或导出全部检查结果为 CSV，包含源行号及完整原因。修改匹配、命名或生成范围后请重新检查。查看排版显示真实页面。\n\n"
            "4. 设置输出并生成\n命名可用 {{序号}} 和模板字段，可先预览前 5 份文件名。主按钮下提示尚缺步骤，灰色操作可悬停查看原因。生成时显示阶段、耗时和剩余时间估算。每次创建独立批次目录和清单。PDF 需要 Word 或 LibreOffice，单份最多等 60 秒。停止后可继续，失败部分可重试。关闭软件后不恢复任务。\n\n"
            "注意：不支持 .doc / .xls、图片占位符或循环表格；公式数据需要先在 Excel / WPS 中计算并保存。软件保留模板结构，长文本仍可能使文档重新分页。所有数据在本机处理。"
        )
        box = tk.Frame(window, bg="white", highlightbackground=BORDER, highlightthickness=1)
        box.pack(fill="both", expand=True, padx=self.px(24))
        window.help_text = tk.Text(box, wrap="word", bg="white", fg=INK, font=(FONT, 9), relief="flat",
                                   padx=self.px(18), pady=self.px(16), spacing3=self.px(5), selectbackground=SELECT, selectforeground=INK)
        scroll = ttk.Scrollbar(box, command=window.help_text.yview)
        scroll.pack(side="right", fill="y")
        window.help_text.configure(yscrollcommand=scroll.set)
        window.help_text.pack(fill="both", expand=True)
        window.help_text.insert("1.0", help_text)
        window.help_text.tag_configure("heading", font=(FONT, 10, "bold"), foreground=GREEN)
        for line, text in enumerate(help_text.splitlines(), 1):
            if text.startswith(("1. ", "2. ", "3. ", "4. ")):
                window.help_text.tag_add("heading", f"{line}.0", f"{line}.end")
        window.help_text.configure(state="disabled")
        return window

    def _close(self):
        if self.busy:
            self.close_requested = True
            self.cancel_event.set()
            self.status.set("已请求停止，正在结束本次操作，随后自动关闭。")
            return
        self._finish_close()

    def _finish_close(self):
        for timer in self.tk.call("after", "info"):
            self.after_cancel(timer)
        self._save_preferences()
        for window in self.preview_windows:
            if window.winfo_exists():
                window.close()
        for folder in list(self.preview_folders):
            if folder.name.startswith("page-"):
                self._remove_preview(folder)
        self._cleanup_previews(schedule=False)
        self.preview_cache.close()
        self.destroy()


def verify_cell_menu(app, tree, row, column):
    """Exercise a Windows menu and dismiss its native message loop in UI checks."""
    from ctypes import wintypes
    class MenuInfo(ctypes.Structure):
        _fields_ = [("cbSize", wintypes.DWORD), ("flags", wintypes.DWORD),
                    ("hwndActive", wintypes.HWND), ("hwndFocus", wintypes.HWND), ("hwndCapture", wintypes.HWND),
                    ("hwndMenuOwner", wintypes.HWND), ("hwndMoveSize", wintypes.HWND), ("hwndCaret", wintypes.HWND),
                    ("rcCaret", wintypes.RECT)]
    api = ctypes.windll.user32
    thread_id = api.GetWindowThreadProcessId(app.winfo_id(), None)
    shown = threading.Event()
    def dismiss():
        for _ in range(100):
            info = MenuInfo()
            info.cbSize = ctypes.sizeof(info)
            if api.GetGUIThreadInfo(thread_id, ctypes.byref(info)) and info.hwndMenuOwner:
                if api.GetWindowThreadProcessId(info.hwndMenuOwner, None) == thread_id:
                    shown.set()
                    api.PostMessageW(info.hwndMenuOwner, 0x0100, 0x1B, 0)
                    return
            time.sleep(0.02)
    x, y, width, height = tree.bbox(row, column)
    x, y = x + width // 2, y + height // 2
    helper = threading.Thread(target=dismiss, daemon=True)
    helper.start()
    tree.event_generate("<Button-3>", x=x, y=y, rootx=tree.winfo_rootx()+x, rooty=tree.winfo_rooty()+y)
    helper.join(timeout=3)
    assert shown.is_set(), "原生右键菜单未显示"
    assert tree.cell_menu.entrycget(0, "label") == "查看完整内容"
    tree.cell_menu.invoke(0)
    return app.cell_window


def verify_file_workflow(app, folder):
    """Verify refreshed input and recent-file actions in source and packaged UI checks."""
    def wait_idle():
        deadline = time.monotonic() + 10
        while app.busy:
            assert time.monotonic() < deadline, app.status.get()
            app.update()
            time.sleep(0.01)
        app.update()
    source = folder / "refresh-source.csv"
    source.write_text(resource("示例文件/人员数据.csv").read_text(encoding="utf-8-sig"), encoding="utf-8-sig")
    app._data_loaded((source, ["CSV"], read_data(source)))
    original = app.dataset
    app.mapping["姓名"].set("部门")
    app.data_tree.selection_set("0", "1")
    app.update()
    source.write_text(source.read_text(encoding="utf-8-sig").replace("林晓", "刷新验收姓名"), encoding="utf-8-sig")
    app.refresh_button.invoke()
    wait_idle()
    assert app.dataset.records[0]["姓名"] == "刷新验收姓名"
    assert original.records[0]["姓名"] == "林晓"
    assert app._mapping()["姓名"] == "部门"
    assert not app.selected_indices and not app.data_tree.selection()
    assert "已刷新" in app.status.get()
    window = app.open_recent_files()
    app.update()
    data_item = next(iid for iid in window.file_tree.get_children() if window.file_tree.set(iid, "path") == str(source))
    window.file_tree.selection_set(data_item)
    app.update()
    window.open_button.invoke()
    wait_idle()
    assert app.data_path == source and app.dataset.records[0]["姓名"] == "刷新验收姓名"
    window = app.open_recent_files()
    template_item = next(iid for iid in window.file_tree.get_children() if window.file_tree.set(iid, "kind") == "模板")
    window.file_tree.selection_set(template_item)
    app.update()
    window.open_button.invoke()
    wait_idle()
    window = app.open_recent_files()
    window.clear_button.invoke()
    assert not window.file_tree.get_children()
    assert not Preferences(app.preferences.path).recent_files("data")
    assert not Preferences(app.preferences.path).recent_files("template")
    assert source.is_file()
    window.destroy()
    source = resource("示例文件/人员数据.xlsx")
    app._data_loaded((source, sheet_names(source), read_data(source)))
    app.mapping["姓名"].set("姓名")
    app._save_preferences()
    return {"refresh_reads_saved_file_and_keeps_mapping": True, "refresh_clears_old_selection": True,
            "recent_template_and_data_reopened": True, "recent_history_cleared_without_deleting_files": True}


def verify_check_report(app, folder):
    """Export genuine multi-page check results through the desktop controls."""
    import csv
    from unittest.mock import patch
    def wait_idle():
        deadline = time.monotonic() + 10
        while app.busy:
            assert time.monotonic() < deadline, app.status.get()
            app.update()
            time.sleep(0.01)
        app.update()
    source = folder / "check-source.csv"
    sample = read_data(resource("示例文件/人员数据.xlsx"))
    with source.open("w", encoding="utf-8-sig", newline="") as stream:
        writer = csv.DictWriter(stream, fieldnames=sample.headers)
        writer.writeheader()
        for index in range(251):
            writer.writerow(dict(sample.records[0], 姓名=f"报告验证{index:04d}", 备注=""))
            writer.writerow({})
    app._data_loaded((source, ["CSV"], read_data(source)))
    app.check_data()
    wait_idle()
    assert len(app.issues) == 251, app.issues[:3]
    app._change_issue_page(1)
    assert len(app.check_tree.get_children()) == 51
    target = folder / "数据检查报告.csv"
    with patch.object(filedialog, "asksaveasfilename", return_value=str(target)):
        app.export_report_button.invoke()
        wait_idle()
    assert target.read_bytes().startswith(b"\xef\xbb\xbf")
    with target.open(encoding="utf-8-sig", newline="") as stream:
        rows = list(csv.DictReader(stream))
    assert len(rows) == 251 and rows[-1]["行号"] == "502"
    assert rows[-1]["说明"] == app.issues[-1]["说明"]
    assert all(row["数据文件"] == str(source) and row["工作表"] == "CSV" for row in rows)
    source = resource("示例文件/人员数据.xlsx")
    app._data_loaded((source, sheet_names(source), read_data(source)))
    assert app.export_report_button.instate(["disabled"])
    return {"actual_check_report_export": True, "all_251_issues_exported": True,
            "source_row_numbers_and_context_correct": True, "report_cleared_after_data_reload": True}


def main():
    if "--pdf-worker" in sys.argv:
        from pdf_worker import run_worker
        run_worker(sys.argv[sys.argv.index("--pdf-worker") + 1])
        return
    if os.name == "nt":
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("Local.TemplateBatch.1")
        except (AttributeError, OSError):
            pass
    report_path = Path(sys.argv[sys.argv.index("--smoke-test") + 1]) if "--smoke-test" in sys.argv else None
    peer_root = Path(sys.argv[sys.argv.index("--smoke-peer") + 1]) if report_path and "--smoke-peer" in sys.argv else None
    app = Application(settings_path=peer_root / "preferences.json" if peer_root else report_path.parent / "smoke-settings.json" if report_path else None)
    # UI smoke testing runs the same callbacks and background workers as a normal session.
    if "--smoke-test" in sys.argv:
        report_path = Path(sys.argv[sys.argv.index("--smoke-test") + 1])
        if peer_root:
            app.withdraw()
            role = report_path.parent.name
            assert role in ("a", "b")
            (peer_root / (role + ".started")).write_text("started")
            def load_when_both_started():
                if all((peer_root / (name + ".started")).exists() for name in ("a", "b")):
                    app.load_demo()
                else:
                    app.after(100, load_when_both_started)
            app.after(250, load_when_both_started)
        else:
            app.after(250, app.load_demo)
        deadline = __import__("time").monotonic() + 120
        stage = [0]
        summary = {}
        if "--smoke-workflow" in sys.argv:
            summary["worker_phases"] = []
            report_phase = app._report_phase
            def capture_phase(phase, current, total):
                if phase not in summary["worker_phases"]:
                    summary["worker_phases"].append(phase)
                report_phase(phase, current, total)
            app._report_phase = capture_phase
        def smoke():
            try:
                if __import__("time").monotonic() > deadline:
                    raise TimeoutError("界面测试超时")
                if app.busy and "已用" in app.action_hint.get():
                    summary["live_elapsed_display"] = True
                if stage[0] == 0 and not app.busy and app.dataset:
                    if "--smoke-files" in sys.argv:
                        summary.update(verify_file_workflow(app, report_path.parent))
                    if "--smoke-report" in sys.argv:
                        summary.update(verify_check_report(app, report_path.parent))
                    if "--smoke-template" in sys.argv:
                        path = Path(sys.argv[sys.argv.index("--smoke-template") + 1])
                        app._template_loaded(Template(path))
                        summary["template_source"] = str(path)
                    summary["template_fingerprint"] = app.template.fingerprint
                    assert not app.generate_button.instate(["disabled"]), app._mapping()
                    app.notebook.select(app.preview_page)
                    assert "林晓" in app.preview_text.get("1.0", "end")
                    app.output.set(str(report_path.parent / "gui-output"))
                    app.pdf.set("--smoke-pdf" in sys.argv)
                    if "--smoke-workflow" in sys.argv:
                        app.nav_buttons[1].invoke()
                        app.update()
                        assert app.notebook.select() == str(app.data_page)
                        assert app.nav_buttons[1].cget("style") == "ActiveNavigation.TButton"
                        summary["navigation_buttons_and_active_state"] = True
                        summary["animations_enabled"] = app.nav_motion.enabled
                        assert app.generate_button.cget("text") == "生成全部 3 份"
                        geometry, minimum = app.geometry(), app.minsize()
                        app.minsize(1, 1)
                        app.geometry(f"{app.winfo_width()}x{app.px(550)}")
                        app.update()
                        for page, tree in ((app.data_page, app.data_tree), (app.result_page, app.result_tree)):
                            app.notebook.select(page)
                            app.update()
                            assert tree.winfo_height() >= app.px(34) * 3
                            assert tree.xview() == (0.0, 1.0)
                        summary["sample_tables_fit_width"] = True
                        assert not app.export_panel.winfo_ismapped()
                        settings = app.open_export_settings()
                        app.update()
                        assert settings.winfo_ismapped()
                        settings.destroy()
                        app.minsize(*minimum)
                        app.geometry(geometry)
                        summary["compact_layout_and_export_settings"] = True
                        app.notebook.select(app.data_page)
                        app.update()
                        tree = app.data_tree
                        tree.selection_set("0", "1")
                        app.update()
                        cell = verify_cell_menu(app, tree, "0", "c6")
                        content = app.dataset.records[0]["备注"]
                        assert cell.content_text.get("1.0", "end-1c") == content
                        assert app.selected_indices == {0, 1}
                        from unittest.mock import patch
                        with patch.object(app, "clipboard_clear"), patch.object(app, "clipboard_append") as copied:
                            cell.copy_button.invoke()
                            copied.assert_called_once_with(content)
                        cell.destroy()
                        app.clear_selection()
                        summary["native_cell_menu_and_full_copy"] = True
                        app.data_query.set("陈雨")
                        app._apply_data_search()
                        assert app.data_indices == [1]
                        assert app.scope.get() == "全部记录"
                        app.locate_row.set("4")
                        app._locate_row()
                        assert app.data_query.get() == ""
                        assert app.preview_index.get() == "3"
                        app.scope.set("选中记录")
                        assert app.generate_button.cget("text") == "生成选中 1 份"
                        app.clear_selection()
                        assert not app.selected_indices and not app.data_tree.selection()
                        assert app.generate_button.cget("text") == "生成选中 0 份"
                        assert app.generate_button.instate(["disabled"])
                        assert "选择记录" in app.action_hint.get()
                        app.generate_button.event_generate("<Enter>")
                        app.update()
                        assert app.reason_popup is not None
                        assert "选择记录" in app.reason_popup.reason_text.cget("text")
                        app.generate_button.event_generate("<Leave>")
                        summary["disabled_action_reason"] = True
                        app.scope.set("全部记录")
                        preview = app.show_filename_preview()
                        names = preview_filenames(app.template, app.dataset, app._mapping(), app.pattern.get())
                        summary["filename_preview_names"] = [item["Word 文件"] for item in names]
                        content = preview.preview_text.get("1.0", "end-1c")
                        assert all(name in content for name in summary["filename_preview_names"])
                        assert app.last_folder is None
                        preview.destroy()
                        summary["generation_scope_and_count"] = True
                        summary["clear_all_selection"] = True
                        summary["search_and_source_row_location"] = True
                        template, dataset, mapping, pattern = app.template, app.dataset, app._mapping(), app.pattern.get()
                        output, pdf = Path(app.output.get()), app.pdf.get()
                        app.last_job = (template, dataset, mapping, pattern)
                        def stopped_progress(number, total, item):
                            app.events.put(("progress", number, total, item, number - 1))
                            if number == 1:
                                app.cancel_event.set()
                        app._run(lambda: generate_batch(template, dataset, mapping, output, pattern,
                                                        export_pdf=pdf, cancel=app.cancel_event, progress=stopped_progress, phase=app._report_phase),
                                 app._batch_done, "正在验证停止与继续…", generation=True)
                        stage[0] = 3
                    else:
                        app.start_generation()
                        stage[0] = 1
                elif stage[0] == 3 and not app.busy and app.last_batch:
                    assert app.last_batch.succeeded == 1, app.last_batch.items
                    assert app.last_batch.remaining_indices == [1, 2]
                    assert not app.resume_button.instate(["disabled"])
                    original_files = [path for path in app.last_folder.iterdir() if path.suffix in (".docx", ".pdf")]
                    summary["_original_state"] = {path: (path.read_bytes(), path.stat().st_mtime_ns) for path in original_files}
                    summary["continued_folder"] = str(app.last_folder)
                    app.start_continue()
                    stage[0] = 1
                elif stage[0] == 1 and not app.busy and app.last_folder:
                    if "--smoke-workflow" in sys.argv:
                        assert str(app.last_folder) == summary["continued_folder"]
                        assert app.last_batch.succeeded == 3, app.last_batch.items
                        assert all("成功" in app.result_tree.item(str(i), "tags") for i in range(3))
                        colors = {str(app.result_tree.tag_configure(status, "foreground")) for status in ("成功", "部分成功", "失败")}
                        assert len(colors) == 3 and "" not in colors
                        summary["result_status_display"] = True
                        assert not app.last_batch.remaining_indices
                        assert app.resume_button.instate(["disabled"])
                        for path, snapshot in summary.pop("_original_state").items():
                            assert (path.read_bytes(), path.stat().st_mtime_ns) == snapshot
                        app.result_filter.set("仅失败和部分成功")
                        app._filter_results()
                        assert not app.result_tree.get_children()
                        app.result_filter.set("全部结果")
                        app._filter_results()
                        app.result_tree.selection_set("1")
                        app._refresh_ready()
                        assert not app.word_result_button.instate(["disabled"])
                        assert not app.manifest_button.instate(["disabled"])
                        if app.pdf.get():
                            assert not app.pdf_result_button.instate(["disabled"])
                        summary["same_batch_continued_and_success_files_unchanged"] = True
                        summary["result_filter_and_file_actions"] = True
                        details = app.show_result_details()
                        content = details.detail_text.get("1.0", "end-1c")
                        assert str(app.last_folder / app.result_items[1]["Word 文件"]) in content
                        assert "数据行：3" in content
                        from unittest.mock import patch
                        with patch.object(app, "clipboard_clear"), patch.object(app, "clipboard_append") as copied:
                            details.copy_button.invoke()
                            copied.assert_called_once_with(content)
                        details.destroy()
                        summary["result_details_and_copy"] = True
                        assert summary["filename_preview_names"] == [item["Word 文件"] for item in app.result_items]
                        assert all(phase in summary["worker_phases"] for phase in ("填充 Word", "保存清单"))
                        if app.pdf.get():
                            assert "转换 PDF" in summary["worker_phases"]
                        assert summary["live_elapsed_display"]
                        assert "上次操作耗时" in app.action_hint.get()
                        summary["filename_preview_matches_generation"] = True
                        summary["phase_and_elapsed_feedback"] = True
                    files = list(app.last_folder.glob("*.docx"))
                    assert len(files) == 3
                    pdf_count = len(list(app.last_folder.glob("*.pdf")))
                    if "--smoke-pdf" in sys.argv:
                        assert pdf_count == 3, app.status.get()
                    summary.update(status="PASS", documents=len(files), pdf_documents=pdf_count, folder=str(app.last_folder))
                    if "--smoke-preview" in sys.argv:
                        app.open_page_preview()
                        stage[0] = 2
                    else:
                        report_path.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
                        app._close()
                        return
                elif stage[0] == 2 and not app.busy and app.preview_windows and app.preview_windows[-1].photo:
                    window = app.preview_windows[-1]
                    folder = window.folder
                    summary["preview_pages"] = len(window.document)
                    if peer_root:
                        from pdf_worker import write_json
                        os.utime(folder, (time.time()-60, time.time()-60))
                        os.utime(app.preview_root, (time.time()-60, time.time()-60))
                        write_json(peer_root / (role + ".ready.json"), {"folder": str(folder), "instance": str(app.preview_root)})
                        stage[0] = 4
                    else:
                        window.close()
                        summary["preview_cleaned"] = not folder.exists()
                        assert summary["preview_cleaned"]
                        report_path.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
                        app._close()
                        return
                elif stage[0] == 4 and all((peer_root / (name + ".ready.json")).exists() for name in ("a", "b")):
                    peer = "b" if role == "a" else "a"
                    peer_info = json.loads((peer_root / (peer + ".ready.json")).read_text(encoding="utf-8"))
                    assert peer_info["instance"] != str(app.preview_root)
                    app._cleanup_previews(schedule=False)
                    assert (Path(peer_info["folder"]) / "文档预览.docx").is_file()
                    assert (Path(peer_info["folder"]) / "文档预览.pdf").is_file()
                    summary["other_process_preview_kept"] = True
                    (peer_root / (role + ".checked")).write_text("checked")
                    stage[0] = 5
                elif stage[0] == 5 and all((peer_root / (name + ".checked")).exists() for name in ("a", "b")):
                    window = app.preview_windows[-1]
                    folder = window.folder
                    window.close()
                    summary["preview_cleaned"] = not folder.exists()
                    assert summary["preview_cleaned"]
                    report_path.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
                    app._close()
                    return
            except Exception:
                report_path.write_text(traceback.format_exc(), encoding="utf-8")
                app._close()
                return
            app.after(100, smoke)
        app.after(400, smoke)
    app.mainloop()


if __name__ == "__main__":
    main()
