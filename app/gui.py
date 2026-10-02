"""Tkinter desktop GUI (Vietnamese labels) – PROMPT-002 redesign.

Thin view over :class:`app.gui_controller.GuiController`: all validation, discovery, state, progress/ETA and
event handling live in the controller (headless, unit-tested); this module only builds widgets, forwards
clicks and renders controller state.  Processing is the same production ``BatchProcessor`` used by the CLI,
on a worker thread; the GUI thread only drains the controller's event queue every 100 ms.

Layout (tab "Xử lý báo cáo", responsive grid inside one vertical page scroller):
    header  →  Nguồn dữ liệu  →  Thời gian xử lý  →  Danh sách báo cáo (toolbar + table, grows with the window)
            →  Xử lý (Start / Stop, progress, summary cards)  →  Kết quả xử lý (table)
Tab "Cấu hình & Ollama": Ollama connection, processing options, log viewer, system diagnostics.
Only ``tkinter`` / ``ttk`` are used (optional ``tkinterdnd2`` drag & drop kept); styling is centralised in
:func:`configure_styles`.
"""
from __future__ import annotations

import logging
import os
import subprocess
import sys
import threading
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import Any, Dict, List, Optional

from . import APP_NAME, APP_TITLE, BUILD_ID, __version__
from .config import AppConfig
from .diagnostics import format_diagnostics, run_diagnostics
from .gui_controller import (DEFAULT_OUTPUT_NAME, FINAL_STATUSES, PERIOD_MODE_VI, SCAN_FILTERS_VI, USER_EXCLUDED,
                             GuiController, default_output_path)
from .prescan import (ACTION_FAST_SKIP, ACTION_INVALID_MGMT, ACTION_MASTER_COMPLETE, ACTION_OUTSIDE_PERIOD,
                      ACTION_MASTER_NOT_FOUND, ACTION_PROCESS, ACTION_SOURCE_DUPLICATE)
from .scanner import parse_dnd_paths

try:  # optional drag & drop support
    from tkinterdnd2 import DND_FILES, TkinterDnD  # type: ignore
    _DND_OK = True
except Exception:  # noqa: BLE001
    _DND_OK = False

LOG = logging.getLogger("report_extractor.gui")

# ----------------------------------------------------------------------------- design tokens
XS, S, M, L, XL = 4, 6, 10, 16, 24                      # spacing tokens (px)
FONT_FAMILY = "Segoe UI"                                # Tk substitutes the default font when unavailable
MONO_FAMILY = "Consolas"
PALETTE = {
    "bg": "#f3f4f6",            # window background (very light neutral)
    "card": "#ffffff",          # cards
    "border": "#d9dce1",
    "text": "#1f2328",
    "secondary": "#6b7280",
    "primary": "#0b57d0",       # blue accent
    "primary_active": "#0947a8",
    "success": "#1a7f37",
    "warning": "#b26a00",
    "error": "#c62828",
    "muted": "#9e9e9e",
    "stripe": "#f7f8fa",
    "selection": "#dbe7ff",
}
DEFAULT_GEOMETRY = (1400, 850)
MIN_GEOMETRY = (1100, 700)
SCAN_TREE_ROWS = 12      # visible rows of "Danh sách báo cáo"; more rows scroll inside the table
RESULT_TREE_ROWS = 9     # visible rows of "Kết quả xử lý"

# scanned-file table: (key, title, width, stretch)
SCAN_COLUMNS = (("stt", "STT", 48, False), ("mgmt", "Management Number", 160, False), ("date", "Ngày phát sinh", 110, False),
                ("vendor", "Vendor", 100, False), ("file", "Tên file", 340, True), ("scan", "Trạng thái quét", 240, False),
                ("path", "Đường dẫn", 420, False))
# processed-results table (order used by RowState.as_values)
COLUMNS = (("stt", "STT", 48), ("mgmt", "Management Number", 150), ("file", "Tên file", 320), ("vendor", "Vendor", 110),
           ("model", "Model", 80), ("item", "Item", 80), ("status", "Trạng thái", 220), ("note", "Ghi chú", 360))
RESULT_STRETCH = ("file", "note")

STATUS_ICON = {"waiting": "○", "completed": "✓", "needs_review": "⚠", "error": "✗", "skipped": "•", "not_written": "⚠",
               "outside_period": "–", "source_duplicate": "•", "fast_skip": "•"}
# scanned-file row tag per pre-scan state (text stays the status; colour is supplemental only)
SCAN_TAGS = {ACTION_PROCESS: "candidate", ACTION_MASTER_NOT_FOUND: "invalid", USER_EXCLUDED: "excluded",
             ACTION_OUTSIDE_PERIOD: "outside", ACTION_SOURCE_DUPLICATE: "duplicate", ACTION_FAST_SKIP: "fast_skip",
             ACTION_MASTER_COMPLETE: "complete", ACTION_INVALID_MGMT: "invalid"}
SCAN_TAG_COLORS = {"candidate": PALETTE["primary"], "excluded": PALETTE["muted"], "outside": PALETTE["muted"],
                   "duplicate": PALETTE["warning"], "fast_skip": PALETTE["secondary"], "complete": PALETTE["success"],
                   "invalid": PALETTE["error"]}
RESULT_TAG_COLORS = {"completed": PALETTE["success"], "needs_review": PALETTE["warning"], "error": PALETTE["error"],
                     "skipped": PALETTE["secondary"], "not_written": "#8e24aa", "working": PALETTE["primary"],
                     "outside_period": PALETTE["muted"], "source_duplicate": PALETTE["muted"], "fast_skip": PALETTE["secondary"]}
SEARCH_PLACEHOLDER = "Tìm Management Number / tên file..."
SUMMARY_CARDS = (("completed", "Hoàn thành", "Success.TLabel"), ("needs_review", "Cần kiểm tra", "Warning.TLabel"),
                 ("failed", "Lỗi", "Error.TLabel"), ("skipped", "Bỏ qua", "Secondary.TLabel"))


def prescan_stage_text(counts) -> str:
    """Stage line after the pre-scan, built from the canonical ``PreScanResult.counts()`` schema with
    default-safe access (``master_not_found``: keys absent from the master – reported, never created)."""
    c = dict(counts or {})
    g = lambda k: c.get(k, 0)  # noqa: E731
    return (f"Quét nhanh: {g('discovered')} file, cần xử lý thực tế {g('candidates')} "
            f"(ngoài thời gian {g('outside_period')}, trùng {g('source_duplicates')}, "
            f"đã xử lý gần đây {g('fast_skipped')}, Excel đầy đủ {g('master_complete')}, "
            f"không có trong Excel {g('master_not_found')}, cần bổ sung {g('incomplete')})")


def summary_counts(summary) -> Dict[str, int]:
    """Default-safe counters for the summary cards (canonical BatchSummary fields only)."""
    out = {}
    for key in ("total", "completed", "needs_review", "failed", "skipped", "not_written"):
        try:
            out[key] = int(getattr(summary, key, 0) or 0)
        except (TypeError, ValueError):
            out[key] = 0
    return out


def ollama_indicator_text(ctl) -> str:
    """Compact '● Ollama: …' indicator for the processing page (status only; configuration lives on tab 2)."""
    if not ctl.server:
        return "● Ollama: Không kết nối — heuristic fallback"
    model = ctl.model or "(chưa chọn model)"
    if ctl.ollama_ok is True:
        return f"● Ollama: {model} — Sẵn sàng"
    if ctl.ollama_ok is False:
        return f"● Ollama: {model} — Không kết nối được (heuristic fallback)"
    return f"● Ollama: {model} — Chưa kiểm tra"


def ollama_indicator_style(ctl) -> str:
    if ctl.ollama_ok is True:
        return "Success.TLabel"
    if ctl.ollama_ok is False or not ctl.server:
        return "Warning.TLabel"
    return "Secondary.TLabel"


def row_matches_search(row, needle: str) -> bool:
    """GUI-only filter of the scanned list: Management Number / file name / path, case-insensitive."""
    n = (needle or "").strip().lower()
    if not n:
        return True
    hay = f"{row.management_number} {row.path.name} {row.path}".lower()
    return n in hay


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


# ----------------------------------------------------------------------------- styling system
def pick_theme(style) -> str:
    """Best available built-in theme: Windows native when usable, otherwise the controllable 'clam'."""
    try:
        names = tuple(style.theme_names() or ())
    except Exception:  # noqa: BLE001
        names = ()
    order = ("vista", "winnative", "clam") if sys.platform == "win32" else ("clam", "alt", "default")
    for name in order:
        if name in names:
            try:
                style.theme_use(name)
                return name
            except Exception:  # noqa: BLE001
                continue
    return ""


def configure_styles(root) -> str:
    """Central semantic ttk styles (no colours/fonts scattered through widget construction)."""
    style = ttk.Style(root)
    theme = pick_theme(style)
    base = (FONT_FAMILY, 10)
    bold = (FONT_FAMILY, 10, "bold")
    try:
        root.configure(background=PALETTE["bg"])
    except Exception:  # noqa: BLE001
        pass
    style.configure(".", font=base, background=PALETTE["bg"], foreground=PALETTE["text"])
    style.configure("App.TFrame", background=PALETTE["bg"])
    style.configure("Card.TFrame", background=PALETTE["card"], relief="flat")
    style.configure("CardBorder.TFrame", background=PALETTE["border"])
    style.configure("TLabel", background=PALETTE["bg"])
    style.configure("Card.TLabel", background=PALETTE["card"])
    style.configure("Header.TLabel", font=(FONT_FAMILY, 17, "bold"), background=PALETTE["bg"])
    style.configure("SubHeader.TLabel", font=(FONT_FAMILY, 10), foreground=PALETTE["secondary"], background=PALETTE["bg"])
    style.configure("Version.TLabel", font=(FONT_FAMILY, 9), foreground=PALETTE["secondary"], background=PALETTE["bg"])
    style.configure("Section.TLabel", font=(FONT_FAMILY, 11, "bold"), background=PALETTE["card"])
    style.configure("Secondary.TLabel", font=(FONT_FAMILY, 9), foreground=PALETTE["secondary"], background=PALETTE["card"])
    style.configure("Field.TLabel", font=base, background=PALETTE["card"])
    style.configure("Success.TLabel", foreground=PALETTE["success"], font=bold, background=PALETTE["card"])
    style.configure("Warning.TLabel", foreground=PALETTE["warning"], font=bold, background=PALETTE["card"])
    style.configure("Error.TLabel", foreground=PALETTE["error"], font=bold, background=PALETTE["card"])
    style.configure("Progress.TLabel", font=(FONT_FAMILY, 12, "bold"), background=PALETTE["card"])
    style.configure("Percent.TLabel", font=(FONT_FAMILY, 16, "bold"), foreground=PALETTE["primary"], background=PALETTE["card"])
    style.configure("CardValue.TLabel", font=(FONT_FAMILY, 15, "bold"), background=PALETTE["card"])
    style.configure("TButton", font=base, padding=(M, XS))
    style.configure("Primary.TButton", font=(FONT_FAMILY, 11, "bold"), padding=(L, S), foreground="#ffffff",
                    background=PALETTE["primary"], borderwidth=0)
    style.map("Primary.TButton", background=[("disabled", "#9bb6e6"), ("active", PALETTE["primary_active"]),
                                             ("pressed", PALETTE["primary_active"])],
              foreground=[("disabled", "#f0f0f0")])
    style.configure("Danger.TButton", font=bold, padding=(L, S), foreground=PALETTE["error"])
    style.configure("Card.TCheckbutton", background=PALETTE["card"])
    style.configure("Card.TRadiobutton", background=PALETTE["card"])
    style.configure("Status.Treeview", font=base, rowheight=24, fieldbackground=PALETTE["card"], background=PALETTE["card"])
    style.configure("Status.Treeview.Heading", font=bold)
    style.map("Status.Treeview", background=[("selected", PALETTE["selection"])], foreground=[("selected", PALETTE["text"])])
    style.configure("TNotebook", background=PALETTE["bg"], tabmargins=(M, S, M, 0))
    style.configure("TNotebook.Tab", font=bold, padding=(L, S))
    style.configure("Horizontal.TProgressbar", thickness=14, background=PALETTE["primary"], troughcolor="#e5e7eb")
    return theme


def apply_dpi_awareness() -> None:
    """Per-monitor DPI awareness on Windows so Tk fonts/controls scale at 125 % / 150 % (no-op elsewhere)."""
    if sys.platform != "win32":
        return
    try:
        import ctypes
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(1)  # type: ignore[attr-defined]
        except Exception:  # noqa: BLE001
            ctypes.windll.user32.SetProcessDPIAware()  # type: ignore[attr-defined]
    except Exception:  # noqa: BLE001
        pass


def valid_geometry(width, height, screen_w, screen_h) -> bool:
    """Persisted window size is reused only when it is sane for the current screen."""
    try:
        w, h, sw, sh = int(width), int(height), int(screen_w), int(screen_h)
    except (TypeError, ValueError):
        return False
    return MIN_GEOMETRY[0] <= w <= sw and MIN_GEOMETRY[1] <= h <= sh


class ReportExtractorApp:
    def __init__(self, cfg: Optional[AppConfig] = None):
        self.ctl = GuiController(cfg or AppConfig.load())
        self.root = TkinterDnD.Tk() if _DND_OK else tk.Tk()
        self.root.title(f"{APP_TITLE} – PPTX → Bảng kiểm chứng")
        self.theme = configure_styles(self.root)
        self._apply_window_geometry()
        self.row_items: Dict[int, str] = {}
        self._config_widgets: List = []
        self._done_handled = True                    # no batch yet; start() arms it
        self.discovery_window = None
        self.discovery_tree = None
        self._discovery_items: Dict[str, Any] = {}
        self._search_placeholder = True
        self._build()
        self._bind_shortcuts()
        self._load_from_controller()
        self.lbl_conn.configure(text="● Đang kiểm tra Ollama local (127.0.0.1:11434)…", style="Secondary.TLabel")
        self.ctl.auto_connect_async()                 # local → saved server; short probe timeout, no LAN scan
        self.root.after(100, self._poll)
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    # ------------------------------------------------------------------ window geometry
    def _apply_window_geometry(self) -> None:
        w, h = DEFAULT_GEOMETRY
        try:
            sw, sh = int(self.root.winfo_screenwidth() or 0), int(self.root.winfo_screenheight() or 0)
        except Exception:  # noqa: BLE001
            sw, sh = 0, 0
        extra = getattr(self.ctl.cfg, "extra", {}) or {}
        if sw and sh:
            if valid_geometry(extra.get("window_width"), extra.get("window_height"), sw, sh):
                w, h = int(extra["window_width"]), int(extra["window_height"])
            w, h = min(w, sw - 40), min(h, sh - 80)
        self.root.geometry(f"{max(w, 800)}x{max(h, 600)}")
        self.root.minsize(*MIN_GEOMETRY)
        if extra.get("window_maximized"):
            try:
                self.root.state("zoomed")
            except Exception:  # noqa: BLE001
                pass

    def _remember_window_geometry(self) -> None:
        try:
            extra = self.ctl.cfg.extra
            maximized = self.root.state() == "zoomed"
            extra["window_maximized"] = bool(maximized)
            if not maximized:
                w, h = int(self.root.winfo_width() or 0), int(self.root.winfo_height() or 0)
                sw, sh = int(self.root.winfo_screenwidth() or 0), int(self.root.winfo_screenheight() or 0)
                if valid_geometry(w, h, sw, sh):
                    extra["window_width"], extra["window_height"] = w, h
        except Exception:  # noqa: BLE001
            pass

    # ------------------------------------------------------------------ small layout helpers
    def _card(self, parent, title: str, row: int, *, weight: int = 0):
        """White card with a 1px border and a section title; returns the inner body frame (grid-managed)."""
        border = ttk.Frame(parent, style="CardBorder.TFrame")
        border.grid(row=row, column=0, sticky="nsew", padx=L, pady=(0, M))
        border.columnconfigure(0, weight=1)
        border.rowconfigure(0, weight=1)
        card = ttk.Frame(border, style="Card.TFrame", padding=(L, M, L, M))
        card.grid(row=0, column=0, sticky="nsew", padx=1, pady=1)
        card.columnconfigure(0, weight=1)
        head = ttk.Frame(card, style="Card.TFrame")
        head.grid(row=0, column=0, sticky="ew", pady=(0, S))
        head.columnconfigure(1, weight=1)
        ttk.Label(head, text=title, style="Section.TLabel").grid(row=0, column=0, sticky="w")
        body = ttk.Frame(card, style="Card.TFrame")
        body.grid(row=1, column=0, sticky="nsew")
        body.columnconfigure(0, weight=1)
        card.rowconfigure(1, weight=1)
        if weight:
            parent.rowconfigure(row, weight=weight)
        body.head = head  # type: ignore[attr-defined]
        return body

    @staticmethod
    def _make_table(parent, columns, *, height: int, selectmode: str, center=(), stretch=()):
        """Treeview with its OWN vertical + horizontal scrollbars (grid: tree / vsb / hsb).

        Non-stretch columns keep their width so long names / paths scroll horizontally; the stretch columns
        absorb extra width when the window grows.
        """
        tree = ttk.Treeview(parent, columns=[c[0] for c in columns], show="headings", selectmode=selectmode,
                            height=height, style="Status.Treeview")
        for col in columns:
            key, title, width = col[0], col[1], col[2]
            do_stretch = key in stretch or (len(col) > 3 and bool(col[3]))
            tree.heading(key, text=title)
            tree.column(key, width=width, minwidth=min(width, 60), anchor="center" if key in center else "w",
                        stretch=do_stretch)
        vsb = ttk.Scrollbar(parent, orient="vertical", command=tree.yview)
        hsb = ttk.Scrollbar(parent, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(0, weight=1)
        return tree, vsb, hsb

    # ------------------------------------------------------------------ layout
    def _build(self) -> None:
        r = self.root
        r.columnconfigure(0, weight=1)
        r.rowconfigure(1, weight=1)

        # ---- header ----------------------------------------------------------------------------------
        hdr = ttk.Frame(r, style="App.TFrame", padding=(L, M, L, S))
        hdr.grid(row=0, column=0, sticky="ew")
        hdr.columnconfigure(0, weight=1)
        ttk.Label(hdr, text=APP_NAME, style="Header.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(hdr, text="PPTX → Bảng kiểm chứng", style="SubHeader.TLabel").grid(row=1, column=0, sticky="w")
        self.lbl_version = ttk.Label(hdr, text=f"v{__version__}", style="Version.TLabel")
        self.lbl_version.grid(row=0, column=1, sticky="e")
        self.lbl_build = ttk.Label(hdr, text=BUILD_ID, style="Version.TLabel")
        self.lbl_build.grid(row=1, column=1, sticky="e")

        # ---- two tabs --------------------------------------------------------------------------------
        self.nb = ttk.Notebook(r)
        self.nb.grid(row=1, column=0, sticky="nsew", padx=0, pady=(0, 0))
        self.tab_run = ttk.Frame(self.nb, style="App.TFrame")
        self.tab_cfg = ttk.Frame(self.nb, style="App.TFrame")
        self.nb.add(self.tab_run, text="  ▶  Xử lý báo cáo  ")
        self.nb.add(self.tab_cfg, text="  ⚙  Cấu hình & Ollama  ")
        self._build_run_tab()
        self._build_cfg_tab()

    def _build_run_tab(self) -> None:
        # ---- whole-page vertical scroll layer: sections scroll, tables scroll internally ------------
        self.tab_run.columnconfigure(0, weight=1)
        self.tab_run.rowconfigure(0, weight=1)
        self.page_canvas = tk.Canvas(self.tab_run, highlightthickness=0, borderwidth=0, background=PALETTE["bg"])
        self.page_vsb = ttk.Scrollbar(self.tab_run, orient="vertical", command=self.page_canvas.yview)
        self.page_canvas.configure(yscrollcommand=self.page_vsb.set)
        self.page_canvas.grid(row=0, column=0, sticky="nsew")
        self.page_vsb.grid(row=0, column=1, sticky="ns")
        self.page = ttk.Frame(self.page_canvas, style="App.TFrame", padding=(0, M, 0, S))
        self._page_win = self.page_canvas.create_window((0, 0), window=self.page, anchor="nw")
        self.page.bind("<Configure>", self._on_page_configure)
        self.page_canvas.bind("<Configure>", self._on_canvas_configure)
        self.page_canvas.bind_all("<MouseWheel>", self._on_page_wheel, add="+")
        self.page_canvas.bind_all("<Button-4>", self._on_page_wheel, add="+")
        self.page_canvas.bind_all("<Button-5>", self._on_page_wheel, add="+")
        self.page.columnconfigure(0, weight=1)
        r = self.page

        # ---- Nguồn dữ liệu ---------------------------------------------------------------------------
        f = self._card(r, "Nguồn dữ liệu", 0)
        f.columnconfigure(1, weight=1)
        ttk.Label(f, text="Thư mục báo cáo", style="Field.TLabel").grid(row=0, column=0, sticky="w", padx=(0, M), pady=XS)
        self.var_folder = tk.StringVar()
        e1 = ttk.Entry(f, textvariable=self.var_folder)
        e1.grid(row=0, column=1, sticky="ew", pady=XS)
        b1 = ttk.Button(f, text="Chọn...", command=self.browse_folder, width=9)
        b1.grid(row=0, column=2, sticky="w", padx=(S, 0), pady=XS)
        self.btn_scan = ttk.Button(f, text="Quét", command=self.scan_reports, width=9)
        self.btn_scan.grid(row=0, column=3, sticky="w", padx=(XS, 0), pady=XS)
        self.lbl_found = ttk.Label(f, text="Chưa chọn thư mục", style="Secondary.TLabel")
        self.lbl_found.grid(row=1, column=1, sticky="w", pady=(0, XS))

        ttk.Label(f, text="File kiểm chứng", style="Field.TLabel").grid(row=2, column=0, sticky="w", padx=(0, M), pady=XS)
        self.var_template = tk.StringVar()
        e2 = ttk.Entry(f, textvariable=self.var_template)
        e2.grid(row=2, column=1, sticky="ew", pady=XS)
        b3 = ttk.Button(f, text="Chọn...", command=self.browse_template, width=9)
        b3.grid(row=2, column=2, sticky="w", padx=(S, 0), pady=XS)

        ttk.Label(f, text="File kết quả", style="Field.TLabel").grid(row=3, column=0, sticky="w", padx=(0, M), pady=XS)
        self.var_output = tk.StringVar()
        e3 = ttk.Entry(f, textvariable=self.var_output)
        e3.grid(row=3, column=1, sticky="ew", pady=XS)
        b4 = ttk.Button(f, text="Chọn...", command=self.browse_output, width=9)
        b4.grid(row=3, column=2, sticky="w", padx=(S, 0), pady=XS)
        ttk.Label(f, text="Không bao giờ ghi đè file kiểm chứng gốc.", style="Secondary.TLabel").grid(
            row=4, column=1, sticky="w")

        # ---- Thời gian xử lý (lọc theo ngày phát sinh từ Management Number, trước khi mở PPTX) ------
        tf = self._card(r, "Thời gian xử lý", 1)
        radios = ttk.Frame(tf, style="Card.TFrame")
        radios.grid(row=0, column=0, sticky="w")
        self.var_period_mode = tk.StringVar(value="auto")
        self._period_radios = []
        for col, (mode, label) in enumerate(PERIOD_MODE_VI.items()):
            rb = ttk.Radiobutton(radios, text=label, value=mode, variable=self.var_period_mode,
                                 command=self._on_period_changed, style="Card.TRadiobutton")
            rb.grid(row=0, column=col, sticky="w", padx=(0, L), pady=XS)
            self._period_radios.append(rb)
        self.frm_month = ttk.Frame(tf, style="Card.TFrame")
        ttk.Label(self.frm_month, text="Tháng", style="Field.TLabel").pack(side="left", padx=(0, XS))
        self.var_period_month = tk.StringVar()
        self.sb_month = ttk.Spinbox(self.frm_month, from_=1, to=12, width=4, format="%02.0f",
                                    textvariable=self.var_period_month, command=self._on_period_changed)
        self.sb_month.pack(side="left", padx=(0, M))
        ttk.Label(self.frm_month, text="Năm", style="Field.TLabel").pack(side="left", padx=(0, XS))
        self.var_period_year = tk.StringVar()
        self.sb_year = ttk.Spinbox(self.frm_month, from_=2020, to=2099, width=6,
                                   textvariable=self.var_period_year, command=self._on_period_changed)
        self.sb_year.pack(side="left")
        self.frm_range = ttk.Frame(tf, style="Card.TFrame")
        ttk.Label(self.frm_range, text="Từ ngày", style="Field.TLabel").pack(side="left", padx=(0, XS))
        self.var_period_from = tk.StringVar()
        self.e_from = ttk.Entry(self.frm_range, textvariable=self.var_period_from, width=12)
        self.e_from.pack(side="left", padx=(0, M))
        ttk.Label(self.frm_range, text="Đến ngày", style="Field.TLabel").pack(side="left", padx=(0, XS))
        self.var_period_to = tk.StringVar()
        self.e_to = ttk.Entry(self.frm_range, textvariable=self.var_period_to, width=12)
        self.e_to.pack(side="left")
        ttk.Label(self.frm_range, text="dd/mm/yyyy, bao gồm cả hai đầu", style="Secondary.TLabel").pack(side="left", padx=(M, 0))
        self.lbl_period = ttk.Label(tf, text="", style="Secondary.TLabel")
        self.lbl_period.grid(row=2, column=0, sticky="w", pady=(XS, 0))
        self.lbl_period_warn = ttk.Label(tf, text="", style="Warning.TLabel")
        self.lbl_period_warn.grid(row=3, column=0, sticky="w")
        for var in (self.var_period_month, self.var_period_year, self.var_period_from, self.var_period_to,
                    self.var_output, self.var_template):
            var.trace_add("write", lambda *_: self._on_period_changed())

        # ---- Danh sách báo cáo: toolbar + table (largest, grows with the window) ---------------------
        sf = self._card(r, "Danh sách báo cáo", 2, weight=3)
        self.lbl_list_count = ttk.Label(sf.head, text="0 file", style="Secondary.TLabel")
        self.lbl_list_count.grid(row=0, column=1, sticky="w", padx=(M, 0))
        self.lbl_ai_run = ttk.Label(sf.head, text=ollama_indicator_text(self.ctl), style=ollama_indicator_style(self.ctl))
        self.lbl_ai_run.grid(row=0, column=2, sticky="e")
        tb = ttk.Frame(sf, style="Card.TFrame")
        tb.grid(row=0, column=0, sticky="ew", pady=(0, S))
        tb.columnconfigure(0, weight=1)
        self.var_search = tk.StringVar()
        self.e_search = ttk.Entry(tb, textvariable=self.var_search)
        self.e_search.grid(row=0, column=0, sticky="ew", padx=(0, S))
        self.e_search.bind("<FocusIn>", self._on_search_focus_in)
        self.e_search.bind("<FocusOut>", self._on_search_focus_out)
        self._set_search_placeholder()
        self.var_search.trace_add("write", lambda *_: self._on_search_changed())
        ttk.Label(tb, text="Hiển thị", style="Field.TLabel").grid(row=0, column=1, padx=(0, XS))
        self.var_scan_filter = tk.StringVar(value=SCAN_FILTERS_VI[0])
        self.cb_scan_filter = ttk.Combobox(tb, textvariable=self.var_scan_filter, values=list(SCAN_FILTERS_VI),
                                           state="readonly", width=20)
        self.cb_scan_filter.grid(row=0, column=2, padx=(0, M))
        self.cb_scan_filter.bind("<<ComboboxSelected>>", lambda _e: self._render_scan())
        self.btn_rescan = ttk.Button(tb, text="Quét lại", command=self.scan_reports)
        self.btn_rescan.grid(row=0, column=3, padx=(0, XS))
        self.btn_restore = ttk.Button(tb, text="Khôi phục", command=self.restore_selected)
        self.btn_restore.grid(row=0, column=4, padx=(0, XS))
        self.btn_exclude = ttk.Button(tb, text="Xóa khỏi danh sách", command=self.exclude_selected)
        self.btn_exclude.grid(row=0, column=5)
        self.lbl_scan = ttk.Label(sf, text="", style="Secondary.TLabel")
        self.lbl_scan.grid(row=1, column=0, sticky="w")
        self.lbl_queue = ttk.Label(sf, text="", style="Secondary.TLabel")
        self.lbl_queue.grid(row=2, column=0, sticky="w", pady=(0, XS))
        tbl = ttk.Frame(sf, style="Card.TFrame")
        tbl.grid(row=3, column=0, sticky="nsew")
        sf.rowconfigure(3, weight=1)
        self.scan_tree, self.scan_vsb, self.scan_hsb = self._make_table(
            tbl, SCAN_COLUMNS, height=SCAN_TREE_ROWS, selectmode="extended", center=("stt", "date"))
        for tag, color in SCAN_TAG_COLORS.items():
            self.scan_tree.tag_configure(tag, foreground=color)
        self.scan_tree.tag_configure("stripe", background=PALETTE["stripe"])
        self.scan_tree.bind("<Delete>", lambda _e: self.exclude_selected())
        self.scan_tree.bind("<Double-1>", self._on_scan_double_click)
        self.scan_menu = tk.Menu(self.root, tearoff=0)
        self.scan_menu.add_command(label="Xóa khỏi danh sách xử lý", command=self.exclude_selected)
        self.scan_menu.add_command(label="Khôi phục", command=self.restore_selected)
        self.scan_tree.bind("<Button-3>", self._on_scan_right_click)
        self.scan_items: Dict[str, int] = {}
        ttk.Label(sf, text="Chỉ xóa khỏi danh sách xử lý, không xóa file gốc. Nháy đúp để xem đường dẫn đầy đủ.",
                  style="Secondary.TLabel").grid(row=4, column=0, sticky="w", pady=(XS, 0))

        # ---- Xử lý: Start / Stop + progress + summary cards ------------------------------------------
        pf = self._card(r, "Xử lý", 3)
        actions = ttk.Frame(pf, style="Card.TFrame")
        actions.grid(row=0, column=0, sticky="ew", pady=(0, S))
        self.btn_start = ttk.Button(actions, text="▶  BẮT ĐẦU XỬ LÝ", style="Primary.TButton", command=self.start)
        self.btn_start.pack(side="left")
        self.btn_stop = ttk.Button(actions, text="■  DỪNG SAU FILE HIỆN TẠI", style="Danger.TButton", command=self.stop,
                                   state="disabled")
        self.btn_stop.pack(side="left", padx=(M, 0))
        self.var_force = tk.BooleanVar(value=False)
        c1 = ttk.Checkbutton(actions, text="Xử lý lại báo cáo đã xử lý (ghi đè các trường tự động)",
                             variable=self.var_force, style="Card.TCheckbutton")
        c1.pack(side="left", padx=(XL, 0))
        prog = ttk.Frame(pf, style="Card.TFrame")
        prog.grid(row=1, column=0, sticky="ew")
        prog.columnconfigure(0, weight=1)
        self.lbl_progress = ttk.Label(prog, text="Sẵn sàng.", style="Progress.TLabel")
        self.lbl_progress.grid(row=0, column=0, sticky="w")
        self.lbl_percent = ttk.Label(prog, text="0%", style="Percent.TLabel", anchor="e")
        self.lbl_percent.grid(row=0, column=1, sticky="e")
        self.pb = ttk.Progressbar(prog, mode="determinate", maximum=100, style="Horizontal.TProgressbar")
        self.pb.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(XS, S))
        self.lbl_stage = ttk.Label(prog, text="", style="Field.TLabel")
        self.lbl_stage.grid(row=2, column=0, columnspan=2, sticky="w")
        self.lbl_file = ttk.Label(prog, text="", style="Secondary.TLabel")
        self.lbl_file.grid(row=3, column=0, columnspan=2, sticky="w")
        times = ttk.Frame(prog, style="Card.TFrame")
        times.grid(row=4, column=0, columnspan=2, sticky="ew", pady=(XS, 0))
        self.lbl_elapsed = ttk.Label(times, text="", style="Field.TLabel")
        self.lbl_elapsed.pack(side="left")
        self.lbl_eta = ttk.Label(times, text="", style="Field.TLabel")
        self.lbl_eta.pack(side="left", padx=(XL, 0))
        cards = ttk.Frame(pf, style="Card.TFrame")
        cards.grid(row=2, column=0, sticky="ew", pady=(M, 0))
        self.card_values: Dict[str, object] = {}
        for col, (key, title, style_name) in enumerate(SUMMARY_CARDS):
            cell = ttk.Frame(cards, style="CardBorder.TFrame")
            cell.grid(row=0, column=col, sticky="ew", padx=(0, M))
            inner = ttk.Frame(cell, style="Card.TFrame", padding=(M, S))
            inner.pack(fill="both", expand=True, padx=1, pady=1)
            val = ttk.Label(inner, text="0", style="CardValue.TLabel")
            val.pack(anchor="w")
            ttk.Label(inner, text=title, style=style_name).pack(anchor="w")
            self.card_values[key] = val
            cards.columnconfigure(col, weight=1, uniform="cards")
        self.lbl_counts = ttk.Label(pf, text=self.ctl.counts_text(), style="Secondary.TLabel")
        self.lbl_counts.grid(row=3, column=0, sticky="w", pady=(S, 0))

        # ---- Kết quả xử lý ---------------------------------------------------------------------------
        lf = self._card(r, "Kết quả xử lý", 4, weight=2)
        hint = "Nháy đúp một dòng để xem chi tiết." + ("  Có thể kéo thả thư mục/file vào bảng." if _DND_OK else "")
        ttk.Label(lf.head, text=hint, style="Secondary.TLabel").grid(row=0, column=1, sticky="w", padx=(M, 0))
        links = ttk.Frame(lf.head, style="Card.TFrame")
        links.grid(row=0, column=2, sticky="e")
        ttk.Button(links, text="Mở file kết quả", command=self.open_output_file).pack(side="left", padx=(0, XS))
        ttk.Button(links, text="Mở thư mục kết quả", command=self.open_output_folder).pack(side="left", padx=(0, XS))
        ttk.Button(links, text="Xem log", command=self.open_log).pack(side="left")
        rtbl = ttk.Frame(lf, style="Card.TFrame")
        rtbl.grid(row=0, column=0, sticky="nsew")
        lf.rowconfigure(0, weight=1)
        self.tree, self.result_vsb, self.result_hsb = self._make_table(
            rtbl, COLUMNS, height=RESULT_TREE_ROWS, selectmode="browse", center=("stt", "model", "item"),
            stretch=RESULT_STRETCH)
        for tag, color in RESULT_TAG_COLORS.items():
            self.tree.tag_configure(tag, foreground=color)
        self.tree.bind("<Double-1>", self._on_row_double_click)
        if _DND_OK:
            for w in (self.tree, rtbl, r):
                try:
                    w.drop_target_register(DND_FILES)
                    w.dnd_bind("<<Drop>>", self._on_drop)
                except Exception:
                    pass

        self._config_widgets.extend([e1, e2, e3, b1, b3, b4, c1, self.sb_month, self.sb_year, self.e_from, self.e_to,
                                     *self._period_radios, self.btn_scan, self.btn_rescan, self.btn_exclude,
                                     self.btn_restore, self.cb_scan_filter])

    def _build_cfg_tab(self) -> None:
        page = self.tab_cfg
        page.columnconfigure(0, weight=1)
        page.rowconfigure(2, weight=1)
        # ---- Kết nối Ollama (editable configuration lives ONLY here) ---------------------------------
        of = self._card(page, "Kết nối Ollama", 0)
        of.columnconfigure(1, weight=1)
        ttk.Label(of, text="Server", style="Field.TLabel").grid(row=0, column=0, sticky="w", padx=(0, M), pady=XS)
        self.var_host = tk.StringVar()
        e4 = ttk.Entry(of, textvariable=self.var_host, width=34)
        e4.grid(row=0, column=1, sticky="w", pady=XS)
        ttk.Label(of, text="host, host:port hoặc http://host:port", style="Secondary.TLabel").grid(
            row=0, column=2, sticky="w", padx=(M, 0))
        ttk.Label(of, text="Port", style="Field.TLabel").grid(row=1, column=0, sticky="w", padx=(0, M), pady=XS)
        self.var_port = tk.StringVar()
        e5 = ttk.Entry(of, textvariable=self.var_port, width=8)
        e5.grid(row=1, column=1, sticky="w", pady=XS)
        ttk.Label(of, text="Model", style="Field.TLabel").grid(row=2, column=0, sticky="w", padx=(0, M), pady=XS)
        self.var_model = tk.StringVar()
        self.cb_model = ttk.Combobox(of, textvariable=self.var_model, state="normal", width=31)
        self.cb_model.grid(row=2, column=1, sticky="w", pady=XS)
        btns = ttk.Frame(of, style="Card.TFrame")
        btns.grid(row=3, column=0, columnspan=3, sticky="w", pady=(S, XS))
        b5 = ttk.Button(btns, text="Kiểm tra kết nối", command=self.check_ollama)
        b5.pack(side="left", padx=(0, XS))
        b6 = ttk.Button(btns, text="Làm mới model", command=self.refresh_models)
        b6.pack(side="left", padx=(0, XS))
        b7 = ttk.Button(btns, text="Lưu cấu hình", command=self.save_ollama)
        b7.pack(side="left", padx=(0, XS))
        ttk.Button(btns, text="Chẩn đoán hệ thống", command=self.show_system_diagnostics).pack(side="left")
        self.lbl_conn = ttk.Label(of, text="● Chưa kiểm tra kết nối", style="Warning.TLabel")
        self.lbl_conn.grid(row=4, column=0, columnspan=3, sticky="w", pady=(XS, 0))
        self.lbl_ai = ttk.Label(of, text=self.ctl.ai_status_text(), style="Secondary.TLabel")
        self.lbl_ai.grid(row=5, column=0, columnspan=3, sticky="w")
        # ---- LAN discovery: user-initiated only, same card, same visual system ----------
        dbtns = ttk.Frame(of, style="Card.TFrame")
        dbtns.grid(row=6, column=0, columnspan=3, sticky="w", pady=(S, XS))
        self.btn_discover = ttk.Button(dbtns, text="Tìm Ollama trong mạng LAN", command=self.discover_ollama)
        self.btn_discover.pack(side="left", padx=(0, XS))
        self.btn_discover_stop = ttk.Button(dbtns, text="Dừng tìm", command=self.cancel_discovery, state="disabled")
        self.btn_discover_stop.pack(side="left", padx=(0, XS))
        self.btn_discover_results = ttk.Button(dbtns, text="Xem kết quả", command=self.show_discovery_results,
                                               state="disabled")
        self.btn_discover_results.pack(side="left")
        self.lbl_discovery = ttk.Label(of, text="Chỉ quét cổng 11434 trong mạng LAN nội bộ khi bạn bấm nút; "
                                                "không tự động đổi server.", style="Secondary.TLabel", wraplength=900)
        self.lbl_discovery.grid(row=7, column=0, columnspan=3, sticky="w")
        for var in (self.var_host, self.var_port, self.var_model):
            var.trace_add("write", lambda *_: self._on_endpoint_edited())

        # ---- Tùy chọn xử lý (existing persisted settings only) --------------------------------------
        xf = self._card(page, "Tùy chọn xử lý", 1)
        c2 = ttk.Checkbutton(xf, text="Xử lý lại dữ liệu đã có (ghi đè các trường tự động – tương đương --force)",
                             variable=self.var_force, style="Card.TCheckbutton")
        c2.grid(row=0, column=0, sticky="w")
        ttk.Label(xf, text="Thời gian xử lý, thư mục và file được chọn trên tab “Xử lý báo cáo”; mọi thiết lập được lưu "
                           "vào config.json khi thay đổi hoặc khi đóng chương trình.", style="Secondary.TLabel",
                  wraplength=900).grid(row=1, column=0, sticky="w", pady=(XS, 0))

        # ---- Nhật ký xử lý (GUI copy of the log stream; clearing never touches app.log) --------------
        lg = self._card(page, "Nhật ký xử lý", 2, weight=1)
        lbtn = ttk.Frame(lg.head, style="Card.TFrame")
        lbtn.grid(row=0, column=2, sticky="e")
        ttk.Button(lbtn, text="Mở file log", command=self.open_log).pack(side="left", padx=(0, XS))
        ttk.Button(lbtn, text="Mở thư mục log", command=self.open_log_folder).pack(side="left", padx=(0, XS))
        ttk.Button(lbtn, text="Xóa phần hiển thị", command=self.clear_log_view).pack(side="left")
        logf = ttk.Frame(lg, style="Card.TFrame")
        logf.grid(row=0, column=0, sticky="nsew")
        lg.rowconfigure(0, weight=1)
        logf.columnconfigure(0, weight=1)
        logf.rowconfigure(0, weight=1)
        self.txt_log = tk.Text(logf, height=12, wrap="word", state="disabled", font=(MONO_FAMILY, 9),
                               relief="flat", background=PALETTE["card"], foreground=PALETTE["text"])
        self.txt_log.grid(row=0, column=0, sticky="nsew")
        log_vsb = ttk.Scrollbar(logf, orient="vertical", command=self.txt_log.yview)
        log_vsb.grid(row=0, column=1, sticky="ns")
        self.txt_log.configure(yscrollcommand=log_vsb.set)

        self._config_widgets.extend([e4, e5, b5, b6, b7, self.cb_model, c2, self.btn_discover])

    def _bind_shortcuts(self) -> None:
        """Safe conveniences: Ctrl+F focuses the search box, F5 rescans only while idle."""
        try:
            self.root.bind("<Control-f>", lambda _e: self._focus_search())
            self.root.bind("<Control-F>", lambda _e: self._focus_search())
            self.root.bind("<F5>", lambda _e: self._rescan_if_idle())
        except Exception:  # noqa: BLE001
            pass

    def _focus_search(self) -> None:
        try:
            self.nb.select(self.tab_run)
            self.e_search.focus_set()
        except Exception:  # noqa: BLE001
            pass

    def _rescan_if_idle(self) -> None:
        if self.ctl.is_running() or self.btn_rescan["state"] == "disabled":
            return
        self.scan_reports()

    # ------------------------------------------------------------------ scroll helpers (layout only)
    def _on_page_configure(self, _event=None) -> None:
        bbox = self.page_canvas.bbox("all")
        if bbox:
            self.page_canvas.configure(scrollregion=bbox)

    def _on_canvas_configure(self, event) -> None:
        """Inner page always as wide as the canvas; as tall as the canvas when the content fits (so the tables
        grow with the window) and as tall as the content otherwise (page scrolls)."""
        try:
            req_h = int(self.page.winfo_reqheight() or 0)
            self.page_canvas.itemconfigure(self._page_win, width=event.width, height=max(event.height, req_h))
        except Exception:  # noqa: BLE001
            pass

    def _on_page_wheel(self, event) -> None:
        """Page scroll only when the pointer is NOT over a table/log (those scroll themselves) or on tab 2."""
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
        self._loading = True                      # traces on template/output must not push the period before
        self.var_folder.set(c.report_folder)     # the period widgets carry the saved mode (kept "auto" otherwise)
        self.var_template.set(c.template)
        self.var_output.set(c.output)
        self.var_host.set(c.host)
        self.var_port.set(str(c.port))
        self.var_model.set(c.model)
        self.cb_model["values"] = c.available_models or ([c.model] if c.model else [])
        self._render_ai_status()
        self.var_force.set(c.force_reprocess)
        self.var_period_mode.set(c.period_mode)
        self.var_period_month.set(c.period_month or f"{datetime.now():%m}")
        self.var_period_year.set(c.period_year or f"{datetime.now():%Y}")
        self.var_period_from.set(c.period_from)
        self.var_period_to.set(c.period_to)
        self._loading = False
        self._on_period_changed()
        self._render_counts()
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
            self.frm_month.grid(row=1, column=0, sticky="w", pady=XS)
        elif mode == "range":
            self.frm_range.grid(row=1, column=0, sticky="w", pady=XS)
        err = self._push_period()
        label = self.ctl.period_label()
        if not err:
            label = "Đang áp dụng: " + label.replace("Thời gian xử lý: ", "", 1)
        self.lbl_period.configure(text=label, style="Error.TLabel" if err else "Secondary.TLabel")
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
        self._render_ai_status()
        return problem

    def _render_ai_status(self) -> None:
        self.lbl_ai.configure(text=self.ctl.ai_status_text())
        if hasattr(self, "lbl_ai_run"):
            self.lbl_ai_run.configure(text=ollama_indicator_text(self.ctl), style=ollama_indicator_style(self.ctl))

    def _sync_endpoint_widgets(self) -> None:
        """Controller → Server/Port/Model widgets (used after the local-first auto-connect selected an endpoint);
        the user may still edit the controls afterwards."""
        c = self.ctl
        if (self.var_host.get().strip(), self.var_port.get().strip(), self.var_model.get().strip()) == \
                (c.host, str(c.port), c.model):
            return
        self._loading = True
        try:
            self.var_host.set(c.host)
            self.var_port.set(str(c.port))
            self.var_model.set(c.model)
        finally:
            self._loading = False

    def _on_endpoint_edited(self) -> None:
        if getattr(self, "_loading", False):
            return
        self.ctl.ollama_ok = None
        self.lbl_conn.configure(text="● Chưa kiểm tra kết nối (đã thay đổi)", style="Warning.TLabel")
        self.lbl_ai.configure(text=f"AI: {self.var_model.get().strip() or '(chưa chọn model)'} @ "
                                   f"{self.var_host.get().strip()}:{self.var_port.get().strip()} — Chưa kiểm tra")
        if hasattr(self, "lbl_ai_run"):
            self.lbl_ai_run.configure(text=f"● Ollama: {self.var_model.get().strip() or '(chưa chọn model)'} — Chưa kiểm tra",
                                      style="Secondary.TLabel")

    def log(self, msg: str) -> None:
        self.txt_log.configure(state="normal")
        self.txt_log.insert("end", msg + "\n")
        self.txt_log.see("end")
        self.txt_log.configure(state="disabled")

    def clear_log_view(self) -> None:
        """Clears ONLY the GUI text widget – app.log / errors.log / batch diagnostics are never touched."""
        self.txt_log.configure(state="normal")
        self.txt_log.delete("1.0", "end")
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
        self.lbl_found.configure(text=f"Đã tìm thấy {n} báo cáo" if n else "Không tìm thấy báo cáo .pptx nào")
        self.lbl_progress.configure(text=f"Đã tìm thấy {n} báo cáo .pptx" if n else "Không tìm thấy báo cáo .pptx nào.")
        self.pb["value"] = 0
        self._render_counts()
        if n and not self.ctl.effective_period()[1]:
            self.lbl_scan.configure(text="Đang quét thư mục...", style="Secondary.TLabel")
            self.ctl.scan_async()
        else:
            self._render_scan()

    # ------------------------------------------------------------------ search (GUI-only filter)
    def _set_search_placeholder(self) -> None:
        self._search_placeholder = True
        self.var_search.set(SEARCH_PLACEHOLDER)
        try:
            self.e_search.configure(foreground=PALETTE["secondary"])
        except Exception:  # noqa: BLE001
            pass

    def _on_search_focus_in(self, _event=None) -> None:
        if self._search_placeholder:
            self._search_placeholder = False
            self.var_search.set("")
            try:
                self.e_search.configure(foreground=PALETTE["text"])
            except Exception:  # noqa: BLE001
                pass

    def _on_search_focus_out(self, _event=None) -> None:
        if not self.var_search.get().strip():
            self._set_search_placeholder()

    def _on_search_changed(self) -> None:
        if getattr(self, "_loading", False) or not hasattr(self, "scan_tree"):
            return
        if self._search_placeholder and self.var_search.get() == SEARCH_PLACEHOLDER:
            return
        self._search_placeholder = False
        self._render_scan()

    def search_text(self) -> str:
        return "" if self._search_placeholder else self.var_search.get().strip()

    def set_search(self, text: str) -> None:
        """Programmatic search (tests / shortcuts): empty text restores every visible row."""
        self._search_placeholder = False
        self.var_search.set(text)
        if not text:
            self._set_search_placeholder()
        self._render_scan()

    # ------------------------------------------------------------------ scanned-file list
    @staticmethod
    def _scan_values(row, stt: int):
        return (stt, row.management_number, row.occurrence_date, "", row.path.name, row.status_vi, str(row.path))

    def _render_scan(self) -> None:
        self.scan_tree.delete(*self.scan_tree.get_children())
        self.scan_items.clear()
        rows = self.ctl.scan_rows(self.var_scan_filter.get())
        needle = self.search_text()
        shown = 0
        for n, row in enumerate(rows, start=1):
            if not row_matches_search(row, needle):
                continue
            shown += 1
            tags = [SCAN_TAGS.get(row.state, "skip")]
            if shown % 2 == 0:
                tags.append("stripe")
            iid = self.scan_tree.insert("", "end", values=self._scan_values(row, n), tags=tuple(tags))
            self.scan_items[iid] = row.index
        total = len(rows)
        self.lbl_list_count.configure(text=f"{shown} / {total} file" if needle else f"{total} file")
        self._render_scan_state()

    def _render_scan_state(self) -> None:
        if not hasattr(self, "lbl_queue"):
            return
        if self.ctl.scan_stale or (self.ctl.scan_result is None and self.ctl.scan_message):
            self.lbl_scan.configure(text=self.ctl.scan_message, style="Warning.TLabel")
        elif self.ctl.scan_result is not None:
            c = self.ctl.scan_result.counts()
            self.lbl_scan.configure(text=f"Đã quét {c.get('discovered', 0)} file", style="Secondary.TLabel")
        else:
            self.lbl_scan.configure(text="", style="Secondary.TLabel")
        self.lbl_queue.configure(text=self.ctl.queue_text() if self.ctl.scan_result else "")
        running = self.ctl.is_running()
        for b in (self.btn_exclude, self.btn_restore):
            b.configure(state="disabled" if running else "normal")

    def _selected_scan_indexes(self):
        return [self.scan_items[iid] for iid in self.scan_tree.selection() if iid in self.scan_items]

    def exclude_selected(self) -> None:
        """Button, Delete key and context menu all end here: remove from the CURRENT queue only (never the file)."""
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
            self.lbl_conn.configure(text=f"● {problem}", style="Error.TLabel")
            return
        self.lbl_conn.configure(text=f"● Đang kiểm tra {self.ctl.endpoint_label}…", style="Secondary.TLabel")
        self.ctl.check_ollama_async(local_first=True)      # worker thread; result arrives in _poll()

    def refresh_models(self) -> None:
        problem = self._push_endpoint()
        if problem:
            self.lbl_conn.configure(text=f"● {problem}", style="Error.TLabel")
            return
        self.lbl_conn.configure(text=f"● Đang lấy danh sách model từ {self.ctl.endpoint_label}…", style="Secondary.TLabel")
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
        self._render_ai_status()

    # ------------------------------------------------------------------ LAN discovery
    def discover_ollama(self) -> None:
        """User clicked "Tìm Ollama trong mạng LAN" – never called automatically."""
        if self.ctl.is_running():
            messagebox.showinfo(APP_NAME, "Đang xử lý báo cáo – vui lòng đợi xong rồi mới tìm Ollama trong mạng LAN.")
            return
        if not self.ctl.discover_ollama_async():
            return
        self.discovery_window = None
        self.btn_discover.configure(state="disabled")
        self.btn_discover_stop.configure(state="normal")
        self.btn_discover_results.configure(state="disabled")
        self.lbl_discovery.configure(text=self.ctl.discovery_progress_text(), style="Secondary.TLabel")
        self.log("Bắt đầu tìm Ollama trong mạng LAN (cổng 11434)…")

    def cancel_discovery(self) -> None:
        if self.ctl.cancel_discovery():
            self.btn_discover_stop.configure(state="disabled")
            self.lbl_discovery.configure(text="Đang dừng tìm Ollama…", style="Secondary.TLabel")

    def _render_discovery_progress(self) -> None:
        self.lbl_discovery.configure(text=self.ctl.discovery_progress_text(), style="Secondary.TLabel")

    def _on_discovery_done(self) -> None:
        results = self.ctl.discovery_results
        self.btn_discover.configure(state="normal" if not self.ctl.is_running() else "disabled")
        self.btn_discover_stop.configure(state="disabled")
        self.btn_discover_results.configure(state="normal" if results else "disabled")
        first = self.ctl.discovery_message.splitlines()[0] if self.ctl.discovery_message else ""
        if results:
            self.lbl_discovery.configure(text=f"{first} – tìm thấy {len(results)} server Ollama.", style="Success.TLabel")
            self.log(f"Tìm thấy {len(results)} server Ollama: " + ", ".join(r.endpoint for r in results))
            self.show_discovery_results()
        else:
            self.lbl_discovery.configure(text=f"{first} – không tìm thấy Ollama trong mạng LAN.", style="Warning.TLabel")
            self.log("Không tìm thấy Ollama trong mạng LAN.")
            messagebox.showinfo(APP_NAME, self.ctl.discovery_message)

    def show_discovery_results(self) -> None:
        """Result list: the user must pick a row and confirm "Sử dụng server này" – even for a single result."""
        results = list(self.ctl.discovery_results)
        if not results:
            return
        win = tk.Toplevel(self.root)
        win.title("Ollama trong mạng LAN")
        win.transient(self.root)
        win.columnconfigure(0, weight=1)
        win.rowconfigure(1, weight=1)
        self.discovery_window = win
        ttk.Label(win, text=f"Tìm thấy {len(results)} server Ollama. Chọn một server rồi bấm “Sử dụng server này”.",
                  style="Secondary.TLabel", wraplength=700).grid(row=0, column=0, sticky="w", padx=M, pady=(M, XS))
        cols = ("server", "models", "status")
        tree = ttk.Treeview(win, columns=cols, show="headings", height=min(10, max(3, len(results))), selectmode="browse")
        for key, title, width in (("server", "Server", 180), ("models", "Models", 420), ("status", "Trạng thái", 160)):
            tree.heading(key, text=title)
            tree.column(key, width=width, anchor="w")
        tree.grid(row=1, column=0, sticky="nsew", padx=M)
        self.discovery_tree = tree
        self._discovery_items = {}
        for r in results:
            iid = tree.insert("", "end", values=(r.endpoint, ", ".join(r.models) or "(chưa có model)",
                                                 f"Sẵn sàng ({r.latency_ms} ms)"))
            self._discovery_items[iid] = r
        btns = ttk.Frame(win)
        btns.grid(row=2, column=0, sticky="e", padx=M, pady=M)
        ttk.Button(btns, text="Quét lại", command=lambda: (win.destroy(), self.discover_ollama())).pack(side="left", padx=(0, XS))
        ttk.Button(btns, text="Sử dụng server này", style="TButton",
                   command=lambda: self.use_discovered_server(win)).pack(side="left", padx=(0, XS))
        ttk.Button(btns, text="Đóng", command=win.destroy).pack(side="left")
        tree.bind("<Double-1>", lambda _e: self.use_discovered_server(win))

    def _selected_discovery_result(self):
        tree = getattr(self, "discovery_tree", None)
        if tree is None:
            return None
        sel = tree.selection()
        if not sel:
            return None
        return self._discovery_items.get(sel[0])

    def use_discovered_server(self, win=None) -> None:
        res = self._selected_discovery_result()
        if res is None:
            messagebox.showinfo(APP_NAME, "Vui lòng chọn một server trong danh sách.")
            return
        if self.ctl.is_running():
            messagebox.showinfo(APP_NAME, "Đang xử lý báo cáo – không thể đổi server lúc này.")
            return
        if not messagebox.askyesno(APP_NAME, f"Sử dụng server Ollama {res.endpoint} cho chương trình?"):
            return
        ok, msg = self.ctl.apply_discovered_server(res)
        self._loading = True
        try:
            self.var_host.set(self.ctl.host)
            self.var_port.set(str(self.ctl.port))
            self.var_model.set(self.ctl.model)
        finally:
            self._loading = False
        if self.ctl.available_models:
            self.cb_model["values"] = self.ctl.available_models
        self._render_ai_status()
        self.lbl_conn.configure(text=f"● Đã chọn server {self.ctl.endpoint_label} (đã lưu cấu hình)",
                                style="Success.TLabel" if ok else "Warning.TLabel")
        self.log(msg)
        if win is not None:
            try:
                win.destroy()
            except tk.TclError:
                pass
        if not ok:
            messagebox.showwarning(APP_NAME, msg)

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
        ok, msg = self.ctl.check_ollama_local_first(timeout=15)
        self._sync_endpoint_widgets()
        self.lbl_conn.configure(text=msg, style="Success.TLabel" if ok else "Error.TLabel")
        self._render_ai_status()
        if not ok:
            if not messagebox.askyesno(APP_NAME, f"{msg}\n\nOllama không khả dụng. Tiếp tục xử lý với heuristic fallback "
                                                 f"(nhận diện theo cấu trúc/từ khoá, không dùng AI)?"):
                return
            use_ai = False
        if not self.ctl.start(use_ollama=use_ai):
            return
        self._done_handled = False
        self._render_rows()
        self._set_running(True)
        self._render_scan_state()
        self._render_counts()
        self._render_progress()
        self.lbl_stage.configure(text="Đang quét thư mục...")
        self.lbl_file.configure(text="")
        self.root.after(1000, self._tick)

    def _tick(self) -> None:
        """Once-a-second refresh of elapsed / ETA from controller state (never blocks the GUI)."""
        if not self.ctl.is_running():
            return
        self.lbl_elapsed.configure(text=self.ctl.elapsed_text())
        self.lbl_eta.configure(text=self.ctl.eta_text())
        self.root.after(1000, self._tick)

    def _render_progress(self) -> None:
        """Bar, percentage and text all come from the SAME controller value (no second calculation)."""
        p = self.ctl.progress
        pct = p.percent
        self.pb["value"] = pct
        self.lbl_percent.configure(text=f"{p.percent_int}%")
        if p.text:
            self.lbl_progress.configure(text=p.text)
        self.lbl_elapsed.configure(text=self.ctl.elapsed_text())
        self.lbl_eta.configure(text=self.ctl.eta_text())

    def _render_counts(self) -> None:
        """Summary cards + secondary line from the canonical BatchSummary (default-safe)."""
        self.lbl_counts.configure(text=self.ctl.counts_text())
        s = self.ctl.summary or (self.ctl.processor.summary if self.ctl.processor else None)
        c = summary_counts(s) if s is not None else {k: 0 for k, _, _ in SUMMARY_CARDS}
        for key, lbl in self.card_values.items():
            lbl.configure(text=str(c.get(key, 0)))

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
        if running and self.ctl.discovery_running:          # a batch never runs with a LAN scan in flight
            self.ctl.cancel_discovery()
        if not running and self.ctl.discovery_running:
            self.btn_discover.configure(state="disabled")

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
        """Drain controller events every 100 ms.  Rendering one malformed/non-critical display event must never
        kill the loop (the loop is re-armed in ``finally``), lose a later ``done`` or leave the GUI running."""
        try:
            for ev in self.ctl.pump():
                try:
                    self._render_event(ev)
                except Exception as e:  # noqa: BLE001 – display-only failure, logged and reported, never fatal
                    msg = f"GUI_EVENT_ERROR event={ev.kind} error={type(e).__name__}: {e}"
                    LOG.exception(msg)
                    self.log(msg)
                if ev.kind == "done":
                    self._on_done()                       # authoritative; runs even if rendering above failed
            if self.ctl.is_running() and not self._done_handled:
                reason = self.ctl.reconcile()             # safety net: dead worker without "done"
                if reason:
                    self.log(reason)
                    self._on_done()
        except Exception as e:  # noqa: BLE001
            LOG.exception("GUI_POLL_ERROR %s", e)
            self.log(f"GUI_POLL_ERROR {type(e).__name__}: {e}")
        finally:
            self.root.after(100, self._poll)              # exactly one polling loop, always re-armed

    def _render_event(self, ev) -> None:
        if ev.kind == "row":
            i = ev.payload[0]
            self._render_row(i)
            p = self.ctl.progress
            self._render_progress()
            if p.current_index is not None and 0 <= i < len(self.ctl.rows) and not self.ctl.rows[i].is_final:
                row = self.ctl.rows[i]
                self.lbl_stage.configure(text=f"{row.status_vi} {p.current_detail}".strip())
                self.lbl_file.configure(text=row.path.name)
        elif ev.kind == "progress":
            self._render_progress()
            self._render_counts()
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
            self._render_counts()
            for i in self.row_items:
                self._render_row(i)
            self.lbl_stage.configure(text=prescan_stage_text(ev.payload.counts()))
        elif ev.kind in ("ollama", "autoconnect"):
            ok, msg = ev.payload
            self._sync_endpoint_widgets()
            style = "Success.TLabel" if ok else ("Warning.TLabel" if self.ctl.ollama_source == "local" else "Error.TLabel")
            self.lbl_conn.configure(text=msg, style=style)
            self._render_ai_status()
            if self.ctl.available_models:
                self.cb_model["values"] = self.ctl.available_models
            if ev.kind == "autoconnect":
                self.log(f"Kết nối Ollama tự động ({self.ctl.ollama_source or 'none'}): {msg}")
                if not ok and self.ctl.ollama_source == "local" and self.ctl.available_models:
                    self.log("Model đã cài trên Ollama local: " + ", ".join(self.ctl.available_models))
            else:
                self.log(msg)
        elif ev.kind == "models":
            ok, msg, models = ev.payload
            self.lbl_conn.configure(text=msg, style="Success.TLabel" if ok else "Error.TLabel")
            if ok and models:
                self.cb_model["values"] = models            # real installed models, user's value kept
                if not self.var_model.get().strip():
                    self.var_model.set(self.ctl.model)
                if self.var_model.get().strip() not in models:
                    self.log(f"Model '{self.var_model.get().strip()}' không có trên máy chủ; "
                             f"có: {', '.join(models)}")
            self.log(msg)
        elif ev.kind == "discovery_progress":
            self._render_discovery_progress()
        elif ev.kind == "discovery_found":
            self._render_discovery_progress()
        elif ev.kind == "discovery_done":
            self._on_discovery_done()
        elif ev.kind == "done":
            pass                                            # handled by _poll -> _on_done (idempotent)

    def _on_done(self) -> None:
        """Finalize the view ONCE (idempotent): internal state first, modal dialog last."""
        if self._done_handled:
            return
        self._done_handled = True
        s = self.ctl.summary
        self._set_running(False)                     # Start enabled / Stop disabled / config editable
        self._render_scan_state()
        for i in self.row_items:
            self._render_row(i)                      # transient "Đang ..." statuses replaced by final ones
        self._render_counts()
        self._render_progress()                      # 100% when every report reached a terminal state; timer frozen
        if self.ctl.worker_failure:
            self.lbl_stage.configure(text=self.ctl.worker_failure)
        else:
            self.lbl_stage.configure(text="Đã dừng theo yêu cầu." if (s and s.stopped) else "Hoàn thành.")
        self.lbl_file.configure(text="")
        self._show_summary_dialog()                  # never opens Excel automatically

    def _show_summary_dialog(self) -> None:
        win = tk.Toplevel(self.root)
        win.title("Kết quả xử lý")
        win.transient(self.root)
        win.resizable(False, False)
        frm = ttk.Frame(win, padding=L)
        frm.pack(fill="both", expand=True)
        for ln in self.ctl.summary_lines():
            ttk.Label(frm, text=ln).pack(anchor="w")
        out = self.ctl.summary.output_file if self.ctl.summary else self.ctl.output
        ttk.Label(frm, text=f"\nFile kết quả:\n{out}", wraplength=560).pack(anchor="w")
        bf = ttk.Frame(frm)
        bf.pack(fill="x", pady=(L, 0))
        ttk.Button(bf, text="Mở file kết quả", command=self.open_output_file).pack(side="left", padx=(0, XS))
        ttk.Button(bf, text="Mở thư mục kết quả", command=self.open_output_folder).pack(side="left", padx=(0, XS))
        ttk.Button(bf, text="Xem log", command=self.open_log).pack(side="left")
        ttk.Button(bf, text="Đóng", command=win.destroy).pack(side="right")

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
        win.geometry("900x600")
        frm = ttk.Frame(win, padding=M)
        frm.pack(fill="both", expand=True)
        frm.columnconfigure(0, weight=1)
        frm.rowconfigure(0, weight=1)
        t = tk.Text(frm, wrap="word", font=(FONT_FAMILY, 10), relief="flat", padx=M, pady=M)
        t.grid(row=0, column=0, sticky="nsew")
        sb = ttk.Scrollbar(frm, orient="vertical", command=t.yview)
        sb.grid(row=0, column=1, sticky="ns")
        t.configure(yscrollcommand=sb.set)
        for k, v in d.items():
            if k.startswith("_"):
                continue
            t.insert("end", f"{k}: ", ("b",))
            t.insert("end", f"{v}\n")
        t.tag_configure("b", font=(FONT_FAMILY, 10, "bold"))
        t.configure(state="disabled")

    def show_system_diagnostics(self) -> None:
        self._push_to_controller()
        cfg = self.ctl.cfg
        cfg.ollama_server, cfg.model = self.ctl.server, self.ctl.model
        win = tk.Toplevel(self.root)
        win.title("Chẩn đoán hệ thống")
        win.geometry("800x460")
        t = tk.Text(win, wrap="word", font=(MONO_FAMILY, 10), padx=M, pady=M)
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

    def open_log_folder(self) -> None:
        p = self.ctl.log_file()
        if p:
            open_path(str(Path(p).parent))
        else:
            messagebox.showinfo(APP_NAME, "Chưa có thư mục log (chưa chạy lần nào).")

    def _on_close(self) -> None:
        if self.ctl.is_running():
            if not messagebox.askyesno(APP_NAME, "Đang xử lý. Thoát sẽ dừng sau báo cáo hiện tại. Thoát?"):
                return
            self.ctl.request_stop()
        self._push_to_controller()
        self._remember_window_geometry()
        self.ctl.save_settings()
        self.root.destroy()

    def run(self) -> None:
        self.root.mainloop()


def launch() -> None:
    apply_dpi_awareness()
    ReportExtractorApp().run()
