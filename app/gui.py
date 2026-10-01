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
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import Dict, Optional

from . import APP_NAME, __version__
from .config import AppConfig
from .diagnostics import format_diagnostics, run_diagnostics
from .gui_controller import DEFAULT_OUTPUT_NAME, FINAL_STATUSES, GuiController, default_output_path
from .scanner import parse_dnd_paths

try:  # optional drag & drop support
    from tkinterdnd2 import DND_FILES, TkinterDnD  # type: ignore
    _DND_OK = True
except Exception:  # noqa: BLE001
    _DND_OK = False

STATUS_ICON = {"waiting": "○", "completed": "✓", "needs_review": "⚠", "error": "✗", "skipped": "•", "not_written": "⚠"}
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

        f = ttk.LabelFrame(r, text="Cấu hình")
        f.pack(fill="x", padx=10, pady=4)
        f.columnconfigure(1, weight=1)

        ttk.Label(f, text="Thư mục báo cáo:").grid(row=0, column=0, sticky="w", **pad)
        self.var_folder = tk.StringVar()
        e1 = ttk.Entry(f, textvariable=self.var_folder)
        e1.grid(row=0, column=1, sticky="ew", **pad)
        fb = ttk.Frame(f)
        fb.grid(row=0, column=2, sticky="w")
        b1 = ttk.Button(fb, text="Chọn thư mục…", command=self.browse_folder)
        b1.pack(side="left", padx=(8, 2), pady=3)
        b2 = ttk.Button(fb, text="Quét lại", command=self.scan_reports)
        b2.pack(side="left", padx=2, pady=3)
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

        self.var_force = tk.BooleanVar(value=False)
        c1 = ttk.Checkbutton(f, text="Xử lý lại báo cáo đã xử lý (ghi đè các trường tự động)", variable=self.var_force)
        c1.grid(row=3, column=1, sticky="w", **pad)

        # ---- Kết nối Ollama -------------------------------------------------------
        of = ttk.LabelFrame(r, text="Kết nối Ollama")
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
        self._config_widgets = [e1, e2, e3, e4, e5, b1, b2, b3, b4, b5, b6, b7, self.cb_model, c1]

        lf = ttk.LabelFrame(r, text="Kết quả từng báo cáo  (nháy đúp để xem chi tiết)" + ("  – có thể kéo thả thư mục/file vào đây" if _DND_OK else ""))
        lf.pack(fill="both", expand=True, padx=10, pady=4)
        self.tree = ttk.Treeview(lf, columns=[c[0] for c in COLUMNS], show="headings", selectmode="browse")
        for key, title, width in COLUMNS:
            self.tree.heading(key, text=title)
            self.tree.column(key, width=width, anchor="center" if key in ("stt", "model", "item") else "w",
                             stretch=key in ("file", "note"))
        vsb = ttk.Scrollbar(lf, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side="left", fill="both", expand=True, padx=(6, 0), pady=6)
        vsb.pack(side="left", fill="y", pady=6)
        for tag, color in (("completed", "#1a7f37"), ("needs_review", "#b26a00"), ("error", "#c62828"),
                           ("skipped", "#666666"), ("not_written", "#8e24aa"), ("working", "#0b57d0")):
            self.tree.tag_configure(tag, foreground=color)
        self.tree.bind("<Double-1>", self._on_row_double_click)
        if _DND_OK:
            for w in (self.tree, lf, r):
                try:
                    w.drop_target_register(DND_FILES)
                    w.dnd_bind("<<Drop>>", self._on_drop)
                except Exception:
                    pass

        pf = ttk.Frame(r)
        pf.pack(fill="x", padx=10, pady=2)
        top = ttk.Frame(pf)
        top.pack(fill="x")
        self.lbl_progress = ttk.Label(top, text="Sẵn sàng.", font=("Segoe UI", 10, "bold"))
        self.lbl_progress.pack(side="left")
        self.lbl_eta = ttk.Label(top, text="")
        self.lbl_eta.pack(side="right", padx=(12, 0))
        self.lbl_elapsed = ttk.Label(top, text="")
        self.lbl_elapsed.pack(side="right")
        self.lbl_stage = ttk.Label(pf, text="")
        self.lbl_stage.pack(anchor="w")
        bar = ttk.Frame(pf)
        bar.pack(fill="x", pady=2)
        self.pb = ttk.Progressbar(bar, mode="determinate", maximum=100)
        self.pb.pack(side="left", fill="x", expand=True)
        self.lbl_percent = ttk.Label(bar, text="0%", width=5, anchor="e")
        self.lbl_percent.pack(side="left", padx=(6, 0))
        self.lbl_counts = ttk.Label(pf, text=self.ctl.counts_text())
        self.lbl_counts.pack(anchor="w")

        self.txt_log = tk.Text(r, height=4, wrap="word", state="disabled", font=("Consolas", 9))
        self.txt_log.pack(fill="x", padx=10, pady=(2, 4))

        bf = ttk.Frame(r)
        bf.pack(fill="x", padx=10, pady=(0, 10))
        ttk.Button(bf, text="Chẩn đoán hệ thống", command=self.show_system_diagnostics).pack(side="left", padx=2)
        ttk.Button(bf, text="Xem log", command=self.open_log).pack(side="left", padx=2)
        ttk.Button(bf, text="Mở file kết quả", command=self.open_output_file).pack(side="left", padx=2)
        ttk.Button(bf, text="Mở thư mục kết quả", command=self.open_output_folder).pack(side="left", padx=2)
        self.btn_start = ttk.Button(bf, text="Bắt đầu xử lý", style="Start.TButton", command=self.start)
        self.btn_start.pack(side="right", padx=2)
        self.btn_stop = ttk.Button(bf, text="Dừng sau báo cáo hiện tại", command=self.stop, state="disabled")
        self.btn_stop.pack(side="right", padx=2)

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
        if c.report_folder and Path(c.report_folder).is_dir():
            self.scan_reports()

    def _push_to_controller(self) -> None:
        c = self.ctl
        c.report_folder = self.var_folder.get().strip()
        c.set_template(self.var_template.get())
        c.set_output(self.var_output.get())
        self._push_endpoint()
        c.force_reprocess = bool(self.var_force.get())

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
        self._push_to_controller()
        n = self.ctl.discover()
        self._render_rows()
        self.lbl_found.configure(text=f"Đã tìm thấy {n} báo cáo")
        self.lbl_progress.configure(text=f"Đã tìm thấy {n} báo cáo .pptx" if n else "Không tìm thấy báo cáo .pptx nào.")
        self.pb["value"] = 0
        self.lbl_counts.configure(text=self.ctl.counts_text())

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
            return
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
        self._render_progress()
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
