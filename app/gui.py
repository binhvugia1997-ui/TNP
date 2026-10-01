"""Tkinter desktop GUI (Vietnamese labels).

Thin view over :class:`app.gui_controller.GuiController`: all validation, discovery,
state and event handling lives in the controller (headless, unit-tested); this module
only builds widgets, forwards clicks and renders controller state.  The processing is
the same production ``BatchProcessor`` used by the CLI, running on a worker thread; the
GUI thread only drains the controller's event queue every 100 ms.
"""
from __future__ import annotations

import os
import subprocess
import sys
import threading
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import Dict, Optional

from . import APP_NAME, __version__
from .config import AppConfig
from .diagnostics import format_diagnostics, run_diagnostics
from .gui_controller import (DEFAULT_OUTPUT_NAME, FINAL_STATUSES, PERIOD_MODE_VI, SCAN_FILTERS_VI, GuiController,
                             default_output_path)

SCAN_TREE_ROWS = 12      # visible rows of "Danh sách file"; more rows scroll inside the table (rule #66)
RESULT_TREE_ROWS = 10    # visible rows of "Kết quả xử lý" (rule #69)
SCAN_COLUMNS = (("stt", "STT", 45), ("mgmt", "Management Number", 150), ("date", "Ngày phát sinh", 100),
                ("file", "Tên file", 320), ("vendor", "Vendor", 90), ("scan", "Trạng thái quét", 220),
                ("path", "Đường dẫn", 260))
from .scanner import parse_dnd_paths

try:  # optional drag & drop support
    from tkinterdnd2 import DND_FILES, TkinterDnD  # type: ignore
    _DND_OK = True
except Exception:  # noqa: BLE001
    _DND_OK = False

STATUS_ICON = {"waiting": "○", "completed": "✓", "needs_review": "⚠", "error": "✗", "skipped": "•", "not_written": "⚠",
               "outside_period": "–", "source_duplicate": "•", "fast_skip": "•"}
COLUMNS = (("stt", "STT", 45), ("mgmt", "Management Number", 140), ("file", "Tên file", 330), ("vendor", "Vendor", 110),
           ("model", "Model", 70), ("item", "Item", 70), ("status", "Trạng thái", 210), ("note", "Ghi chú", 320))


def open_path(path: str) -> None:
    try:
        if sys.platform == "win32":
            os.startfile(path)  # type: ignore[attr-defined]
        elif sys.platform == "darwin":
            subprocess.Popen(["open", path])
        else:
            subprocess.Popen(["xdg-open", path])
    except Exception as e:  # noqa: BLE001
        messagebox.showerror(APP_NAME, f"Không mở được: {path}\n{e}")


class ReportExtractorApp:
    def __init__(self, cfg: Optional[AppConfig] = None):
        self.ctl = GuiController(cfg or AppConfig.load())
        self.root = TkinterDnD.Tk() if _DND_OK else tk.Tk()
        self.root.title(f"{APP_NAME} v{__version__} – Báo cáo → Kiểm chứng đối sách")
        self.root.geometry("1180x800")
        self.root.minsize(960, 640)
        self.row_items: Dict[int, str] = {}
        self._config_widgets = []
        self._build()
        self._load_from_controller()
        self.root.after(100, self._poll)
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    # ------------------------------------------------------------------ layout
    def _build(self) -> None:
        r = self.root
        style = ttk.Style(r)
        try:
            style.theme_use("vista" if sys.platform == "win32" else "clam")
        except Exception:
            pass
        style.configure("Title.TLabel", font=("Segoe UI", 13, "bold"))
        style.configure("Ok.TLabel", foreground="#1a7f37", font=("Segoe UI", 10, "bold"))
        style.configure("Bad.TLabel", foreground="#c62828", font=("Segoe UI", 10, "bold"))
        style.configure("Start.TButton", font=("Segoe UI", 10, "bold"))
        pad = {"padx": 8, "pady": 3}
        ttk.Label(r, text="BÁO CÁO PPTX  →  BẢNG KIỂM CHỨNG ĐỐI SÁCH", style="Title.TLabel").pack(anchor="w", padx=10, pady=(10, 4))
        # two tabs: processing (inputs, period, results, progress) | configuration & Ollama
        self.nb = ttk.Notebook(r)
        self.nb.pack(fill="both", expand=True)
        self.tab_run = ttk.Frame(self.nb)
        self.tab_cfg = ttk.Frame(self.nb)
        self.nb.add(self.tab_run, text="Xử lý báo cáo")
        self.nb.add(self.tab_cfg, text="Cấu hình & Ollama")
        # ---- whole-page vertical scroll layer for tab 1 (rule #70): sections scroll, tables scroll internally
        self.page_canvas = tk.Canvas(self.tab_run, highlightthickness=0, borderwidth=0)
        self.page_vsb = ttk.Scrollbar(self.tab_run, orient="vertical", command=self.page_canvas.yview)
        self.page_canvas.configure(yscrollcommand=self.page_vsb.set)
        self.page_vsb.pack(side="right", fill="y")
        self.page_canvas.pack(side="left", fill="both", expand=True)
        self.page = ttk.Frame(self.page_canvas)
        self._page_win = self.page_canvas.create_window((0, 0), window=self.page, anchor="nw")
        self.page.bind("<Configure>", self._on_page_configure)
        self.page_canvas.bind("<Configure>", self._on_canvas_configure)
        self.page_canvas.bind_all("<MouseWheel>", self._on_page_wheel, add="+")
        self.page_canvas.bind_all("<Button-4>", self._on_page_wheel, add="+")
        self.page_canvas.bind_all("<Button-5>", self._on_page_wheel, add="+")
        r = self.page

        # ---- 1. Nguồn dữ liệu ------------------------------------------------------------------------
        f = ttk.LabelFrame(r, text="1. Nguồn dữ liệu")
        f.pack(fill="x", padx=10, pady=4)
        f.columnconfigure(1, weight=1)

        ttk.Label(f, text="Thư mục báo cáo:").grid(row=0, column=0, sticky="w", **pad)
        self.var_folder = tk.StringVar()
        e1 = ttk.Entry(f, textvariable=self.var_folder)
        e1.grid(row=0, column=1, sticky="ew", **pad)
        b1 = ttk.Button(f, text="Chọn thư mục…", command=self.browse_folder)
        b1.grid(row=0, column=2, sticky="w", **pad)
        self.lbl_found = ttk.Label(f, text="Chưa chọn thư mục")
        self.lbl_found.grid(row=0, column=3, sticky="w", **pad)

        ttk.Label(f, text="File Kiểm chứng (.xlsx):").grid(row=1, column=0, sticky="w", **pad)
        self.var_template = tk.StringVar()
        e2 = ttk.Entry(f, textvariable=self.var_template)
        e2.grid(row=1, column=1, sticky="ew", **pad)
        b3 = ttk.Button(f, text="Chọn file…", command=self.browse_template)
        b3.grid(row=1, column=2, sticky="w", **pad)

        ttk.Label(f, text="File kết quả:").grid(row=2, column=0, sticky="w", **pad)
        self.var_output = tk.StringVar()
        e3 = ttk.Entry(f, textvariable=self.var_output)
        e3.grid(row=2, column=1, sticky="ew", **pad)
        b4 = ttk.Button(f, text="Chọn file…", command=self.browse_output)
        b4.grid(row=2, column=2, sticky="w", **pad)
        ttk.Label(f, text="(không bao giờ ghi đè file Kiểm chứng gốc)").grid(row=2, column=3, sticky="w", **pad)

        # ---- 2. Thời gian xử lý (lọc theo ngày phát sinh từ Management Number, trước khi mở PPTX) ------
        tf = ttk.LabelFrame(r, text="2. Thời gian xử lý")
        tf.pack(fill="x", padx=10, pady=4)
        self.var_period_mode = tk.StringVar(value="auto")
        self._period_radios = []
        for col, (mode, label) in enumerate(PERIOD_MODE_VI.items()):
            rb = ttk.Radiobutton(tf, text=label, value=mode, variable=self.var_period_mode,
                                 command=self._on_period_changed)
            rb.grid(row=0, column=col, sticky="w", **pad)
            self._period_radios.append(rb)
        self.frm_month = ttk.Frame(tf)
        ttk.Label(self.frm_month, text="Tháng:").pack(side="left", padx=(0, 4))
        self.var_period_month = tk.StringVar()
        self.sb_month = ttk.Spinbox(self.frm_month, from_=1, to=12, width=4, format="%02.0f",
                                    textvariable=self.var_period_month, command=self._on_period_changed)
        self.sb_month.pack(side="left", padx=(0, 10))
        ttk.Label(self.frm_month, text="Năm:").pack(side="left", padx=(0, 4))
        self.var_period_year = tk.StringVar()
        self.sb_year = ttk.Spinbox(self.frm_month, from_=2020, to=2099, width=6,
                                   textvariable=self.var_period_year, command=self._on_period_changed)
        self.sb_year.pack(side="left")
        self.frm_range = ttk.Frame(tf)
        ttk.Label(self.frm_range, text="Từ ngày:").pack(side="left", padx=(0, 4))
        self.var_period_from = tk.StringVar()
        self.e_from = ttk.Entry(self.frm_range, textvariable=self.var_period_from, width=12)
        self.e_from.pack(side="left", padx=(0, 10))
        ttk.Label(self.frm_range, text="Đến ngày:").pack(side="left", padx=(0, 4))
        self.var_period_to = tk.StringVar()
        self.e_to = ttk.Entry(self.frm_range, textvariable=self.var_period_to, width=12)
        self.e_to.pack(side="left")
        ttk.Label(self.frm_range, text="(dd/mm/yyyy, bao gồm cả hai đầu)").pack(side="left", padx=(8, 0))
        self.lbl_period = ttk.Label(tf, text="")
        self.lbl_period.grid(row=2, column=0, columnspan=6, sticky="w", **pad)
        self.lbl_period_warn = ttk.Label(tf, text="", style="Bad.TLabel")
        self.lbl_period_warn.grid(row=3, column=0, columnspan=6, sticky="w", **pad)
        for var in (self.var_period_month, self.var_period_year, self.var_period_from, self.var_period_to,
                    self.var_output, self.var_template):
            var.trace_add("write", lambda *_: self._on_period_changed())

        # ---- 3. AI (status only; editable Ollama settings live on tab 2) ------------------------------
        af = ttk.LabelFrame(r, text="3. AI")
        af.pack(fill="x", padx=10, pady=4)
        self.lbl_ai_run = ttk.Label(af, text=self.ctl.ai_status_text())
        self.lbl_ai_run.pack(side="left", padx=8, pady=3)
        ttk.Button(af, text="Cấu hình Ollama…", command=lambda: self.nb.select(self.tab_cfg)).pack(side="right", padx=8, pady=3)

        # ---- 4. Các nút chức năng (always above the file list – rules #63–#65) --------------------------
        xf = ttk.LabelFrame(r, text="4. Các nút chức năng")
        xf.pack(fill="x", padx=10, pady=4)
        row_list = ttk.Frame(xf)
        row_list.pack(fill="x", padx=6, pady=(4, 2))
        ttk.Label(row_list, text="Danh sách:", width=10).pack(side="left")
        self.btn_scan = ttk.Button(row_list, text="Quét file", command=self.scan_reports)
        self.btn_scan.pack(side="left", padx=2)
        self.btn_rescan = ttk.Button(row_list, text="Quét lại", command=self.scan_reports)
        self.btn_rescan.pack(side="left", padx=2)
        self.btn_exclude = ttk.Button(row_list, text="Xóa khỏi danh sách", command=self.exclude_selected)
        self.btn_exclude.pack(side="left", padx=(12, 2))
        self.btn_restore = ttk.Button(row_list, text="Khôi phục", command=self.restore_selected)
        self.btn_restore.pack(side="left", padx=2)
        ttk.Label(row_list, text="Hiển thị:").pack(side="left", padx=(16, 2))
        self.var_scan_filter = tk.StringVar(value=SCAN_FILTERS_VI[0])
        self.cb_scan_filter = ttk.Combobox(row_list, textvariable=self.var_scan_filter, values=list(SCAN_FILTERS_VI),
                                           state="readonly", width=20)
        self.cb_scan_filter.pack(side="left")
        self.cb_scan_filter.bind("<<ComboboxSelected>>", lambda _e: self._render_scan())
        ttk.Separator(xf, orient="horizontal").pack(fill="x", padx=6, pady=2)
        row_proc = ttk.Frame(xf)
        row_proc.pack(fill="x", padx=6, pady=(2, 4))
        ttk.Label(row_proc, text="Xử lý:", width=10).pack(side="left")
        self.var_force = tk.BooleanVar(value=False)
        c1 = ttk.Checkbutton(row_proc, text="Xử lý lại báo cáo đã xử lý (ghi đè các trường tự động)", variable=self.var_force)
        c1.pack(side="left", padx=2)
        self.btn_start = ttk.Button(row_proc, text="Bắt đầu xử lý", style="Start.TButton", command=self.start)
        self.btn_start.pack(side="left", padx=(16, 2))
        self.btn_stop = ttk.Button(row_proc, text="Dừng sau báo cáo hiện tại", command=self.stop, state="disabled")
        self.btn_stop.pack(side="left", padx=2)

        # ---- 5. Thống kê quét ------------------------------------------------------------------------
        kf = ttk.LabelFrame(r, text="5. Thống kê quét")
        kf.pack(fill="x", padx=10, pady=4)
        self.lbl_scan = ttk.Label(kf, text="")
        self.lbl_scan.pack(anchor="w", padx=8)
        self.lbl_queue = ttk.Label(kf, text="")
        self.lbl_queue.pack(anchor="w", padx=8, pady=(0, 3))

        # ---- 6. Danh sách file (own vertical + horizontal scrollbars – rules #66–#68) ----------------
        sf = ttk.LabelFrame(r, text="6. Danh sách file  (kiểm tra / loại file trước khi xử lý – không xoá file gốc; nháy đúp để xem đường dẫn)")
        sf.pack(fill="x", padx=10, pady=4)
        self.scan_tree, _svsb, _shsb = self._make_table(sf, SCAN_COLUMNS, height=SCAN_TREE_ROWS, selectmode="extended",
                                                      center=("stt", "date"))
        for tag, color in (("candidate", "#0b57d0"), ("excluded", "#9e9e9e"), ("skip", "#777777")):
            self.scan_tree.tag_configure(tag, foreground=color)
        self.scan_tree.bind("<Delete>", lambda _e: self.exclude_selected())
        self.scan_tree.bind("<Double-1>", self._on_scan_double_click)
        self.scan_menu = tk.Menu(self.root, tearoff=0)
        self.scan_menu.add_command(label="Xóa khỏi danh sách xử lý", command=self.exclude_selected)
        self.scan_menu.add_command(label="Khôi phục", command=self.restore_selected)
        self.scan_tree.bind("<Button-3>", self._on_scan_right_click)
        self.scan_items: Dict[str, int] = {}

        # ---- 7. Tiến trình ---------------------------------------------------------------------------
        pf = ttk.LabelFrame(r, text="7. Tiến trình")
        pf.pack(fill="x", padx=10, pady=4)
        top = ttk.Frame(pf)
        top.pack(fill="x", padx=6)
        self.lbl_progress = ttk.Label(top, text="Sẵn sàng.", font=("Segoe UI", 10, "bold"))
        self.lbl_progress.pack(side="left")
        self.lbl_eta = ttk.Label(top, text="")
        self.lbl_eta.pack(side="right", padx=(12, 0))
        self.lbl_elapsed = ttk.Label(top, text="")
        self.lbl_elapsed.pack(side="right")
        self.lbl_stage = ttk.Label(pf, text="")
        self.lbl_stage.pack(anchor="w", padx=6)
        bar = ttk.Frame(pf)
        bar.pack(fill="x", padx=6, pady=2)
        self.pb = ttk.Progressbar(bar, mode="determinate", maximum=100)
        self.pb.pack(side="left", fill="x", expand=True)
        self.lbl_percent = ttk.Label(bar, text="0%", width=5, anchor="e")
        self.lbl_percent.pack(side="left", padx=(6, 0))
        self.lbl_counts = ttk.Label(pf, text=self.ctl.counts_text())
        self.lbl_counts.pack(anchor="w", padx=6, pady=(0, 3))

        # ---- 8. Kết quả xử lý (own vertical + horizontal scrollbars – rule #69) ---------------------
        lf = ttk.LabelFrame(r, text="8. Kết quả xử lý  (nháy đúp để xem chi tiết)" + ("  – có thể kéo thả thư mục/file vào đây" if _DND_OK else ""))
        lf.pack(fill="x", padx=10, pady=4)
        self.tree, _vsb, _hsb = self._make_table(lf, COLUMNS, height=RESULT_TREE_ROWS, selectmode="browse",
                                               center=("stt", "model", "item"))
        for tag, color in (("completed", "#1a7f37"), ("needs_review", "#b26a00"), ("error", "#c62828"),
                           ("skipped", "#666666"), ("not_written", "#8e24aa"), ("working", "#0b57d0"),
                           ("outside_period", "#9e9e9e"), ("source_duplicate", "#9e9e9e"), ("fast_skip", "#666666")):
            self.tree.tag_configure(tag, foreground=color)
        self.tree.bind("<Double-1>", self._on_row_double_click)
        if _DND_OK:
            for w in (self.tree, lf, r):
                try:
                    w.drop_target_register(DND_FILES)
                    w.dnd_bind("<<Drop>>", self._on_drop)
                except Exception:
                    pass

        self.txt_log = tk.Text(r, height=4, wrap="word", state="disabled", font=("Consolas", 9))
        self.txt_log.pack(fill="x", padx=10, pady=(2, 4))

        bf = ttk.Frame(r)
        bf.pack(fill="x", padx=10, pady=(0, 10))
        ttk.Button(bf, text="Xem log", command=self.open_log).pack(side="left", padx=2)
        ttk.Button(bf, text="Mở file kết quả", command=self.open_output_file).pack(side="left", padx=2)
        ttk.Button(bf, text="Mở thư mục kết quả", command=self.open_output_folder).pack(side="left", padx=2)

        # ---- tab 2: Ollama settings (editable config must stay here) ----------------------------------
        of = ttk.LabelFrame(self.tab_cfg, text="Kết nối Ollama")
        of.pack(fill="x", padx=10, pady=4)
        ttk.Label(of, text="IP / Server:").grid(row=0, column=0, sticky="w", **pad)
        self.var_host = tk.StringVar()
        e4 = ttk.Entry(of, textvariable=self.var_host, width=26)
        e4.grid(row=0, column=1, sticky="w", **pad)
        ttk.Label(of, text="Port:").grid(row=0, column=2, sticky="w", **pad)
        self.var_port = tk.StringVar()
        e5 = ttk.Entry(of, textvariable=self.var_port, width=8)
        e5.grid(row=0, column=3, sticky="w", **pad)
        ttk.Label(of, text="Model:").grid(row=0, column=4, sticky="w", **pad)
        self.var_model = tk.StringVar()
        self.cb_model = ttk.Combobox(of, textvariable=self.var_model, state="normal", width=22)
        self.cb_model.grid(row=0, column=5, sticky="w", **pad)
        b5 = ttk.Button(of, text="Kiểm tra kết nối", command=self.check_ollama)
        b5.grid(row=0, column=6, sticky="w", **pad)
        b6 = ttk.Button(of, text="Làm mới model", command=self.refresh_models)
        b6.grid(row=0, column=7, sticky="w", **pad)
        b7 = ttk.Button(of, text="Lưu cấu hình", command=self.save_ollama)
        b7.grid(row=0, column=8, sticky="w", **pad)
        self.lbl_conn = ttk.Label(of, text="● Chưa kiểm tra kết nối", style="Bad.TLabel")
        self.lbl_conn.grid(row=1, column=0, columnspan=5, sticky="w", **pad)
        self.lbl_ai = ttk.Label(of, text=self.ctl.ai_status_text())
        self.lbl_ai.grid(row=1, column=5, columnspan=4, sticky="w", **pad)
        for var in (self.var_host, self.var_port, self.var_model):
            var.trace_add("write", lambda *_: self._on_endpoint_edited())
        cf = ttk.Frame(self.tab_cfg)
        cf.pack(fill="x", padx=10, pady=(0, 10))
        ttk.Button(cf, text="Chẩn đoán hệ thống", command=self.show_system_diagnostics).pack(side="left", padx=2)

        self._config_widgets = [e1, e2, e3, e4, e5, b1, b3, b4, b5, b6, b7, self.cb_model, c1,
                                self.sb_month, self.sb_year, self.e_from, self.e_to, *self._period_radios,
                                self.btn_scan, self.btn_rescan, self.btn_exclude, self.btn_restore, self.cb_scan_filter]

    # ------------------------------------------------------------------ scroll helpers (layout only)
    @staticmethod
    def _make_table(parent, columns, *, height: int, selectmode: str, center=()):
        """Treeview with its OWN vertical + horizontal scrollbars (grid: tree / vsb / hsb).

        Columns do not stretch, so long file names / paths scroll horizontally instead of being squeezed.
        """
        tree = ttk.Treeview(parent, columns=[c[0] for c in columns], show="headings", selectmode=selectmode, height=height)
        for key, title, width in columns:
            tree.heading(key, text=title)
            tree.column(key, width=width, minwidth=min(width, 60), anchor="center" if key in center else "w", stretch=False)
        vsb = ttk.Scrollbar(parent, orient="vertical", command=tree.yview)
        hsb = ttk.Scrollbar(parent, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        tree.grid(row=0, column=0, sticky="nsew", padx=(6, 0), pady=(6, 0))
        vsb.grid(row=0, column=1, sticky="ns", pady=(6, 0))
        hsb.grid(row=1, column=0, sticky="ew", padx=(6, 0), pady=(0, 6))
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(0, weight=1)
        return tree, vsb, hsb

    def _on_page_configure(self, _event=None) -> None:
        bbox = self.page_canvas.bbox("all")
        if bbox:
            self.page_canvas.configure(scrollregion=bbox)

    def _on_canvas_configure(self, event) -> None:
        # inner frame always as wide as the canvas -> sections stretch horizontally, only vertical page scroll
        try:
            self.page_canvas.itemconfigure(self._page_win, width=event.width)
        except tk.TclError:
            pass

    def _on_page_wheel(self, event) -> None:
        """Page scroll only when the pointer is NOT over a table/log (those scroll themselves)."""
        try:
            w = event.widget if not isinstance(event.widget, str) else self.root.nametowidget(event.widget)
        except Exception:  # noqa: BLE001
            return
        while w is not None:
            if isinstance(w, (ttk.Treeview, tk.Text, ttk.Combobox, ttk.Spinbox)):
                return
            if w is self.tab_cfg:
                return
            w = getattr(w, "master", None)
        if getattr(event, "num", None) == 4 or getattr(event, "delta", 0) > 0:
            self.page_canvas.yview_scroll(-1, "units")
        elif getattr(event, "num", None) == 5 or getattr(event, "delta", 0) < 0:
            self.page_canvas.yview_scroll(1, "units")

    # ------------------------------------------------------------------ controller <-> widgets
    def _load_from_controller(self) -> None:
        c = self.ctl
        self.var_folder.set(c.report_folder)
        self.var_template.set(c.template)
        self.var_output.set(c.output)
        self._loading = True
        self.var_host.set(c.host)
        self.var_port.set(str(c.port))
        self.var_model.set(c.model)
        self.cb_model["values"] = c.available_models or ([c.model] if c.model else [])
        self._loading = False
        self.lbl_ai.configure(text=c.ai_status_text())
        self.var_force.set(c.force_reprocess)
        self._loading = True
        self.var_period_mode.set(c.period_mode)
        self.var_period_month.set(c.period_month or f"{datetime.now():%m}")
        self.var_period_year.set(c.period_year or f"{datetime.now():%Y}")
        self.var_period_from.set(c.period_from)
        self.var_period_to.set(c.period_to)
        self._loading = False
        self._on_period_changed()
        if c.report_folder and Path(c.report_folder).is_dir():
            self.scan_reports()

    def _push_to_controller(self) -> None:
        c = self.ctl
        c.report_folder = self.var_folder.get().strip()
        c.set_template(self.var_template.get())
        c.set_output(self.var_output.get())
        self._push_endpoint()
        c.set_force(bool(self.var_force.get()))
        self._push_period()
        self._render_scan_state()

    def _push_period(self) -> str:
        return self.ctl.set_period(self.var_period_mode.get(), self.var_period_month.get(), self.var_period_year.get(),
                                   self.var_period_from.get(), self.var_period_to.get())

    def _on_period_changed(self) -> None:
        """Show the ACTIVE period (or the Vietnamese problem) before the batch starts; manual values win over
        the Excel file name; the auto month is recomputed from the CURRENT Excel path."""
        if getattr(self, "_loading", False) or not hasattr(self, "lbl_period"):
            return
        self.ctl.set_template(self.var_template.get())
        self.ctl.set_output(self.var_output.get())
        mode = self.var_period_mode.get()
        self.frm_month.grid_forget()
        self.frm_range.grid_forget()
        if mode == "month":
            self.frm_month.grid(row=1, column=0, columnspan=6, sticky="w", padx=6, pady=2)
        elif mode == "range":
            self.frm_range.grid(row=1, column=0, columnspan=6, sticky="w", padx=6, pady=2)
        err = self._push_period()
        self.lbl_period.configure(text=self.ctl.period_label(), style="Bad.TLabel" if err else "TLabel")
        self.lbl_period_warn.configure(text=self.ctl.period_warning())

    def _push_endpoint(self) -> str:
        """Host/port/model from the widgets into the controller (takes effect immediately)."""
        problem = self.ctl.set_endpoint(self.var_host.get(), self.var_port.get(), self.var_model.get())
        if not problem:
            # set_endpoint may have split "host:port" / "http://host:port" typed into the IP field
            if self.var_host.get().strip() != self.ctl.host:
                self.var_host.set(self.ctl.host)
            if self.var_port.get().strip() != str(self.ctl.port):
                self.var_port.set(str(self.ctl.port))
        self.lbl_ai.configure(text=self.ctl.ai_status_text())
        if hasattr(self, "lbl_ai_run"):
            self.lbl_ai_run.configure(text=self.ctl.ai_status_text())
        return problem

    def _on_endpoint_edited(self) -> None:
        if getattr(self, "_loading", False):
            return
        self.ctl.ollama_ok = None
        self.lbl_conn.configure(text="● Chưa kiểm tra kết nối (đã thay đổi)", style="Bad.TLabel")
        self.lbl_ai.configure(text=f"AI: {self.var_model.get().strip() or '(chưa chọn model)'} @ "
                                   f"{self.var_host.get().strip()}:{self.var_port.get().strip()} — Chưa kiểm tra")

    def log(self, msg: str) -> None:
        self.txt_log.configure(state="normal")
        self.txt_log.insert("end", msg + "\n")
        self.txt_log.see("end")
        self.txt_log.configure(state="disabled")

    # ------------------------------------------------------------------ pickers
    def browse_folder(self) -> None:
        p = filedialog.askdirectory(title="Chọn thư mục chứa báo cáo PPTX")
        if p:
            self.var_folder.set(p)
            if not self.var_output.get().strip() or Path(self.var_output.get()).name == DEFAULT_OUTPUT_NAME:
                self.var_output.set(default_output_path(p))
            self.scan_reports()
            self._push_to_controller()
            self.ctl.save_settings()

    def browse_template(self) -> None:
        p = filedialog.askopenfilename(title="Chọn file Kiểm chứng", filetypes=[("Excel", "*.xlsx *.xlsm"), ("Tất cả", "*.*")])
        if p:
            self.var_template.set(p)
            self._push_to_controller()
            self.ctl.save_settings()

    def browse_output(self) -> None:
        init = self.var_output.get() or default_output_path(self.var_folder.get()) or DEFAULT_OUTPUT_NAME
        p = filedialog.asksaveasfilename(title="File kết quả", defaultextension=".xlsx", initialfile=Path(init).name,
                                         initialdir=str(Path(init).parent) if init else None,
                                         filetypes=[("Excel", "*.xlsx")], confirmoverwrite=False)
        if p:
            self.var_output.set(p)
            self._push_to_controller()
            self.ctl.save_settings()

    def scan_reports(self) -> None:
        """Quét lại: cheap discovery now, the REAL pre-scan in a worker thread (GUI never freezes)."""
        self._push_to_controller()
        if self.ctl.is_running():
            return
        n = self.ctl.discover()
        self._render_rows()
        self.lbl_found.configure(text=f"Đã tìm thấy {n} báo cáo")
        self.lbl_progress.configure(text=f"Đã tìm thấy {n} báo cáo .pptx" if n else "Không tìm thấy báo cáo .pptx nào.")
        self.pb["value"] = 0
        self.lbl_counts.configure(text=self.ctl.counts_text())
        if n and not self.ctl.effective_period()[1]:
            self.lbl_scan.configure(text="Đang quét thư mục...")
            self.ctl.scan_async()
        else:
            self._render_scan()

    # ------------------------------------------------------------------ scanned-file list
    def _render_scan(self) -> None:
        self.scan_tree.delete(*self.scan_tree.get_children())
        self.scan_items.clear()
        rows = self.ctl.scan_rows(self.var_scan_filter.get())
        for n, row in enumerate(rows, start=1):
            tag = "excluded" if row.excluded else ("candidate" if row.is_candidate else "skip")
            iid = self.scan_tree.insert("", "end", values=row.as_values(n), tags=(tag,))
            self.scan_items[iid] = row.index
        self._render_scan_state()

    def _render_scan_state(self) -> None:
        if not hasattr(self, "lbl_queue"):
            return
        if self.ctl.scan_stale or (self.ctl.scan_result is None and self.ctl.scan_message):
            self.lbl_scan.configure(text=self.ctl.scan_message, style="Bad.TLabel")
        elif self.ctl.scan_result is not None:
            c = self.ctl.scan_result.counts()
            self.lbl_scan.configure(text=f"Đã quét {c['discovered']} file", style="TLabel")
        else:
            self.lbl_scan.configure(text="", style="TLabel")
        self.lbl_queue.configure(text=self.ctl.queue_text() if self.ctl.scan_result else "")
        running = self.ctl.is_running()
        for b in (self.btn_exclude, self.btn_restore):
            b.configure(state="disabled" if running else "normal")

    def _selected_scan_indexes(self):
        return [self.scan_items[iid] for iid in self.scan_tree.selection() if iid in self.scan_items]

    def exclude_selected(self) -> None:
        """Button, Delete key and context menu all end here: remove from the CURRENT queue only."""
        problem = self.ctl.exclude(self._selected_scan_indexes())
        if problem:
            self.log(problem)
        self._render_scan()

    def restore_selected(self) -> None:
        problem = self.ctl.restore(self._selected_scan_indexes())
        if problem:
            self.log(problem)
        self._render_scan()

    def _on_scan_right_click(self, event) -> None:
        try:
            iid = self.scan_tree.identify_row(event.y)
            if iid and iid not in self.scan_tree.selection():
                self.scan_tree.selection_set(iid)
            self.scan_menu.tk_popup(event.x_root, event.y_root)
        finally:
            try:
                self.scan_menu.grab_release()
            except Exception:
                pass

    def _on_scan_double_click(self, _event=None) -> None:
        idx = self._selected_scan_indexes()
        if not idx:
            return
        row = next((r for r in self.ctl.scan_rows() if r.index == idx[0]), None)
        if row:
            messagebox.showinfo(APP_NAME, f"{row.path.name}\n\nManagement Number: {row.management_number or '(trống)'}\n"
                                          f"Ngày phát sinh: {row.occurrence_date or '(trống)'}\n"
                                          f"Trạng thái quét: {row.status_vi}\n{row.reason}\n\nĐường dẫn:\n{row.path}")

    def _on_drop(self, event) -> None:
        paths = parse_dnd_paths(event.data)
        folders = [p for p in paths if Path(p).is_dir()]
        if folders:
            self.var_folder.set(folders[0])
            if not self.var_output.get().strip():
                self.var_output.set(default_output_path(folders[0]))
            self.scan_reports()

    # ------------------------------------------------------------------ Ollama
    def check_ollama(self) -> None:
        problem = self._push_endpoint()
        if problem:
            self.lbl_conn.configure(text=f"● {problem}", style="Bad.TLabel")
            return
        self.lbl_conn.configure(text=f"● Đang kiểm tra {self.ctl.endpoint_label}…", style="Bad.TLabel")
        self.ctl.check_ollama_async()                      # worker thread; result arrives in _poll()

    def refresh_models(self) -> None:
        problem = self._push_endpoint()
        if problem:
            self.lbl_conn.configure(text=f"● {problem}", style="Bad.TLabel")
            return
        self.lbl_conn.configure(text=f"● Đang lấy danh sách model từ {self.ctl.endpoint_label}…", style="Bad.TLabel")
        self.ctl.refresh_models_async()

    def save_ollama(self) -> None:
        problem = self._push_endpoint()
        if problem:
            messagebox.showwarning(APP_NAME, problem)
            return
        err = self.ctl.save_ollama_settings()
        if err:
            messagebox.showwarning(APP_NAME, err)
            return
        self.log(f"Đã lưu cấu hình Ollama: {self.ctl.model} @ {self.ctl.endpoint_label}")
        self.lbl_ai.configure(text=self.ctl.ai_status_text())

    # ------------------------------------------------------------------ run
    def start(self) -> None:
        if not self.ctl.can_start():
            return
        self._push_to_controller()
        if not self.ctl.files:
            self.scan_reports()
        errs = self.ctl.validate()
        if errs:
            messagebox.showwarning(APP_NAME, "Chưa thể bắt đầu:\n\n• " + "\n• ".join(errs))
            self._render_scan_state()
            return
        if self.ctl.scan_result is None:
            self.ctl.scan()                      # synchronous – the list was never scanned (e.g. keyboard-only user)
            self._render_scan()
        use_ai = True
        ok, msg = self.ctl.check_ollama(timeout=15)
        self.lbl_conn.configure(text=msg, style="Ok.TLabel" if ok else "Bad.TLabel")
        self.lbl_ai.configure(text=self.ctl.ai_status_text())
        if not ok:
            if not messagebox.askyesno(APP_NAME, f"{msg}\n\nOllama không khả dụng. Tiếp tục xử lý với heuristic fallback "
                                                 f"(nhận diện theo cấu trúc/từ khoá, không dùng AI)?"):
                return
            use_ai = False
        if not self.ctl.start(use_ollama=use_ai):
            return
        self._render_rows()
        self._set_running(True)
        self._render_scan_state()
        self._render_progress()
        self.lbl_stage.configure(text="Đang quét thư mục...")
        self.root.after(1000, self._tick)

    def _tick(self) -> None:
        """Once-a-second refresh of elapsed / ETA from controller state (never blocks the GUI)."""
        if not self.ctl.is_running():
            return
        self.lbl_elapsed.configure(text=self.ctl.elapsed_text())
        self.lbl_eta.configure(text=self.ctl.eta_text())
        self.root.after(1000, self._tick)

    def _render_progress(self) -> None:
        """Bar, percentage and text all come from the SAME controller value."""
        p = self.ctl.progress
        pct = p.percent
        self.pb["value"] = pct
        self.lbl_percent.configure(text=f"{p.percent_int}%")
        if p.text:
            self.lbl_progress.configure(text=p.text)
        self.lbl_elapsed.configure(text=self.ctl.elapsed_text())
        self.lbl_eta.configure(text=self.ctl.eta_text())

    def stop(self) -> None:
        if self.ctl.request_stop():
            self.btn_stop.configure(state="disabled")
            self.log("Sẽ dừng sau khi xử lý xong báo cáo hiện tại…")

    def _set_running(self, running: bool) -> None:
        self.btn_start.configure(state="disabled" if running else "normal")
        self.btn_stop.configure(state="normal" if running else "disabled")
        for w in self._config_widgets:
            try:
                w.configure(state="disabled" if running else "normal")
            except tk.TclError:
                pass

    # ------------------------------------------------------------------ rendering
    def _render_rows(self) -> None:
        self.tree.delete(*self.tree.get_children())
        self.row_items.clear()
        for row in self.ctl.rows:
            self.row_items[row.index] = self.tree.insert("", "end", values=self._values(row), tags=self._tags(row))

    @staticmethod
    def _values(row):
        v = list(row.as_values())
        icon = STATUS_ICON.get(row.stage, "⟳" if row.stage != "waiting" else "○")
        v[6] = f"{icon} {row.status_vi}"
        return tuple(v)

    @staticmethod
    def _tags(row):
        if row.stage in FINAL_STATUSES:
            return (row.stage,)
        return ("working",) if row.stage != "waiting" else ()

    def _render_row(self, i: int) -> None:
        iid = self.row_items.get(i)
        if iid is None or i >= len(self.ctl.rows):
            return
        row = self.ctl.rows[i]
        self.tree.item(iid, values=self._values(row), tags=self._tags(row))
        if row.stage != "waiting":
            self.tree.see(iid)

    def _poll(self) -> None:
        for ev in self.ctl.pump():
            if ev.kind == "row":
                i = ev.payload[0]
                self._render_row(i)
                p = self.ctl.progress
                self._render_progress()
                if p.current_index is not None and not self.ctl.rows[i].is_final:
                    self.lbl_stage.configure(text=f"{self.ctl.rows[i].path.name} — {self.ctl.rows[i].status_vi} "
                                                  f"{p.current_detail}".strip())
            elif ev.kind == "progress":
                self._render_progress()
                self.lbl_counts.configure(text=self.ctl.counts_text())
            elif ev.kind == "log":
                self.log(str(ev.payload))
            elif ev.kind == "scan":
                self._render_scan()
                if self.ctl.scan_result is not None:
                    self.log(self.ctl.queue_text())
                elif self.ctl.scan_message:
                    self.log(self.ctl.scan_message)
            elif ev.kind == "prescan":
                self._render_progress()
                self.lbl_counts.configure(text=self.ctl.counts_text())
                for i in self.row_items:
                    self._render_row(i)
                c = ev.payload.counts()
                self.lbl_stage.configure(text=f"Quét nhanh: {c['discovered']} file, cần xử lý thực tế {c['candidates']} "
                                              f"(ngoài thời gian {c['outside_period']}, trùng {c['source_duplicates']}, "
                                              f"đã xử lý gần đây {c['fast_skipped']}, Excel đầy đủ {c['master_complete']}, "
                                              f"không có trong Excel {c['master_not_found']})")
            elif ev.kind == "ollama":
                ok, msg = ev.payload
                self.lbl_conn.configure(text=msg, style="Ok.TLabel" if ok else "Bad.TLabel")
                self.lbl_ai.configure(text=self.ctl.ai_status_text())
                if self.ctl.available_models:
                    self.cb_model["values"] = self.ctl.available_models
                self.log(msg)
            elif ev.kind == "models":
                ok, msg, models = ev.payload
                self.lbl_conn.configure(text=msg, style="Ok.TLabel" if ok else "Bad.TLabel")
                if ok and models:
                    self.cb_model["values"] = models            # real installed models, user's value kept
                    if not self.var_model.get().strip():
                        self.var_model.set(self.ctl.model)
                    if self.var_model.get().strip() not in models:
                        self.log(f"Model '{self.var_model.get().strip()}' không có trên máy chủ; "
                                 f"có: {', '.join(models)}")
                self.log(msg)
            elif ev.kind == "done":
                self._on_done()
        self.root.after(100, self._poll)

    def _on_done(self) -> None:
        s = self.ctl.summary
        self._set_running(False)
        self._render_scan_state()
        for i in self.row_items:
            self._render_row(i)
        self.lbl_counts.configure(text=self.ctl.counts_text())
        self._render_progress()                      # 100% only when every report reached a terminal state
        self.lbl_stage.configure(text="Đã dừng theo yêu cầu." if (s and s.stopped) else "Hoàn thành.")
        self._show_summary_dialog()

    def _show_summary_dialog(self) -> None:
        win = tk.Toplevel(self.root)
        win.title("Kết quả xử lý")
        win.transient(self.root)
        win.resizable(False, False)
        frm = ttk.Frame(win, padding=14)
        frm.pack(fill="both", expand=True)
        for ln in self.ctl.summary_lines():
            ttk.Label(frm, text=ln, font=("Segoe UI", 10)).pack(anchor="w")
        out = self.ctl.summary.output_file if self.ctl.summary else self.ctl.output
        ttk.Label(frm, text=f"\nFile kết quả:\n{out}", wraplength=520).pack(anchor="w")
        bf = ttk.Frame(frm)
        bf.pack(fill="x", pady=(12, 0))
        ttk.Button(bf, text="Mở file kết quả", command=self.open_output_file).pack(side="left", padx=2)
        ttk.Button(bf, text="Mở thư mục kết quả", command=self.open_output_folder).pack(side="left", padx=2)
        ttk.Button(bf, text="Xem log", command=self.open_log).pack(side="left", padx=2)
        ttk.Button(bf, text="Đóng", command=win.destroy).pack(side="right", padx=2)

    # ------------------------------------------------------------------ dialogs
    def _on_row_double_click(self, _event=None) -> None:
        sel = self.tree.selection()
        if not sel:
            return
        idx = next((i for i, iid in self.row_items.items() if iid == sel[0]), None)
        if idx is None:
            return
        d = self.ctl.diagnostics_for(idx)
        if d is None:
            messagebox.showinfo(APP_NAME, "Báo cáo này chưa được xử lý.")
            return
        win = tk.Toplevel(self.root)
        win.title(f"Chi tiết – {d['Tên file']}")
        win.geometry("860x560")
        t = tk.Text(win, wrap="word", font=("Segoe UI", 10))
        t.pack(fill="both", expand=True)
        for k, v in d.items():
            if k.startswith("_"):
                continue
            t.insert("end", f"{k}: ", ("b",))
            t.insert("end", f"{v}\n")
        t.tag_configure("b", font=("Segoe UI", 10, "bold"))
        t.configure(state="disabled")

    def show_system_diagnostics(self) -> None:
        self._push_to_controller()
        cfg = self.ctl.cfg
        cfg.ollama_server, cfg.model = self.ctl.server, self.ctl.model
        win = tk.Toplevel(self.root)
        win.title("Chẩn đoán hệ thống")
        win.geometry("760x420")
        t = tk.Text(win, wrap="word", font=("Consolas", 10))
        t.pack(fill="both", expand=True)
        t.insert("end", "Đang kiểm tra…\n")

        def work():
            text = format_diagnostics(run_diagnostics(cfg, template=self.ctl.template or None,
                                                      output_folder=str(Path(self.ctl.output).parent) if self.ctl.output else None,
                                                      smoke=True))
            self.root.after(0, lambda: (t.delete("1.0", "end"), t.insert("end", text)))
        threading.Thread(target=work, daemon=True).start()

    def open_output_file(self) -> None:
        p = self.ctl.output_file()
        if p:
            open_path(str(p))
        else:
            messagebox.showinfo(APP_NAME, "Chưa có file kết quả.")

    def open_output_folder(self) -> None:
        p = self.ctl.output_folder()
        if p:
            open_path(str(p))
        else:
            messagebox.showinfo(APP_NAME, "Chưa có thư mục kết quả.")

    def open_log(self) -> None:
        p = self.ctl.log_file()
        if p:
            open_path(str(p))
        else:
            messagebox.showinfo(APP_NAME, "Chưa có file log (chưa chạy lần nào).")

    def _on_close(self) -> None:
        if self.ctl.is_running():
            if not messagebox.askyesno(APP_NAME, "Đang xử lý. Thoát sẽ dừng sau báo cáo hiện tại. Thoát?"):
                return
            self.ctl.request_stop()
        self._push_to_controller()
        self.ctl.save_settings()
        self.root.destroy()

    def run(self) -> None:
        self.root.mainloop()


def launch() -> None:
    ReportExtractorApp().run()
