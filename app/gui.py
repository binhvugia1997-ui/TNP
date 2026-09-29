"""Tkinter desktop GUI (Vietnamese labels). All heavy work runs in worker threads;
the GUI thread only consumes a queue of status events."""
from __future__ import annotations

import os
import queue
import subprocess
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import Dict, List, Optional

from . import APP_NAME, __version__
from .batch_processor import STAGE_LABELS_VI, BatchOptions, BatchProcessor, BatchSummary
from .config import AppConfig, normalize_ollama_url
from .diagnostics import format_diagnostics, run_diagnostics
from .ollama_client import OllamaClient, OllamaError, preferred_model
from .scanner import parse_dnd_paths, scan_inputs

try:  # optional drag & drop support
    from tkinterdnd2 import DND_FILES, TkinterDnD  # type: ignore
    _DND_OK = True
except Exception:  # noqa: BLE001
    _DND_OK = False

DEFAULT_OUTPUT_NAME = "Kiem_chung_doi_sach_TONG_HOP.xlsx"
STATUS_ICON = {
    "waiting": "○", "reading": "⟳", "analyzing": "⟳", "extracting_qpn": "⟳",
    "extracting_images": "⟳", "writing_excel": "⟳", "completed": "✓",
    "needs_review": "⚠", "error": "✗", "skipped": "•",
}


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
        self.cfg = cfg or AppConfig.load()
        self.root = TkinterDnD.Tk() if _DND_OK else tk.Tk()
        self.root.title(f"{APP_NAME} v{__version__} – Báo cáo → Kiểm chứng đối sách")
        self.root.geometry("980x760")
        self.root.minsize(860, 620)
        self.queue: "queue.Queue" = queue.Queue()
        self.files: List[Path] = []
        self.file_items: Dict[int, str] = {}
        self.processor: Optional[BatchProcessor] = None
        self.summary: Optional[BatchSummary] = None
        self.error_details: List[str] = []
        self._build()
        self._load_config()
        self.root.after(100, self._poll_queue)
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    # ------------------------------------------------------------------
    # Layout
    # ------------------------------------------------------------------
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
        ttk.Label(r, text="BÁO CÁO  →  KIỂM CHỨNG ĐỐI SÁCH", style="Title.TLabel").pack(anchor="w", padx=10, pady=(10, 4))

        # ---- settings frame -------------------------------------------------
        f = ttk.LabelFrame(r, text="Cấu hình")
        f.pack(fill="x", padx=10, pady=4)
        f.columnconfigure(1, weight=1)

        ttk.Label(f, text="Máy AI (Ollama):").grid(row=0, column=0, sticky="w", **pad)
        self.var_server = tk.StringVar()
        ttk.Entry(f, textvariable=self.var_server).grid(row=0, column=1, sticky="ew", **pad)
        ttk.Button(f, text="Kiểm tra kết nối", command=self.test_connection).grid(row=0, column=2, **pad)

        ttk.Label(f, text="Model AI:").grid(row=1, column=0, sticky="w", **pad)
        self.var_model = tk.StringVar()
        self.cb_model = ttk.Combobox(f, textvariable=self.var_model, state="normal")
        self.cb_model.grid(row=1, column=1, sticky="ew", **pad)
        self.lbl_conn = ttk.Label(f, text="● Chưa kết nối", style="Bad.TLabel")
        self.lbl_conn.grid(row=1, column=2, sticky="w", **pad)

        ttk.Label(f, text="Form Excel:").grid(row=2, column=0, sticky="w", **pad)
        self.var_template = tk.StringVar()
        ttk.Entry(f, textvariable=self.var_template).grid(row=2, column=1, sticky="ew", **pad)
        ttk.Button(f, text="Chọn file…", command=self.browse_template).grid(row=2, column=2, **pad)

        ttk.Label(f, text="Thư mục báo cáo:").grid(row=3, column=0, sticky="w", **pad)
        self.var_folder = tk.StringVar()
        ttk.Entry(f, textvariable=self.var_folder).grid(row=3, column=1, sticky="ew", **pad)
        fb = ttk.Frame(f)
        fb.grid(row=3, column=2, sticky="w")
        ttk.Button(fb, text="Chọn thư mục", command=self.browse_folder).pack(side="left", padx=(8, 2), pady=3)
        ttk.Button(fb, text="Quét báo cáo", command=self.scan_reports).pack(side="left", padx=2, pady=3)

        ttk.Label(f, text="File kết quả:").grid(row=4, column=0, sticky="w", **pad)
        self.var_output = tk.StringVar()
        ttk.Entry(f, textvariable=self.var_output).grid(row=4, column=1, sticky="ew", **pad)
        ttk.Button(f, text="Chọn file…", command=self.browse_output).grid(row=4, column=2, **pad)

        self.var_force = tk.BooleanVar(value=False)
        ttk.Checkbutton(f, text="Xử lý lại các báo cáo đã xử lý trước đó (bỏ qua kiểm tra trùng)",
                        variable=self.var_force).grid(row=5, column=1, sticky="w", **pad)

        # ---- file list --------------------------------------------------------
        lf = ttk.LabelFrame(r, text="Danh sách báo cáo  (kéo thả thư mục hoặc file .pptx vào đây)"
                            if _DND_OK else "Danh sách báo cáo")
        lf.pack(fill="both", expand=True, padx=10, pady=4)
        cols = ("stt", "file", "status", "detail")
        self.tree = ttk.Treeview(lf, columns=cols, show="headings", selectmode="browse")
        self.tree.heading("stt", text="#")
        self.tree.heading("file", text="Báo cáo")
        self.tree.heading("status", text="Trạng thái")
        self.tree.heading("detail", text="Chi tiết")
        self.tree.column("stt", width=45, anchor="center", stretch=False)
        self.tree.column("file", width=420)
        self.tree.column("status", width=170, stretch=False)
        self.tree.column("detail", width=280)
        vsb = ttk.Scrollbar(lf, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side="left", fill="both", expand=True, padx=(6, 0), pady=6)
        vsb.pack(side="left", fill="y", pady=6)
        self.tree.tag_configure("completed", foreground="#1a7f37")
        self.tree.tag_configure("needs_review", foreground="#b26a00")
        self.tree.tag_configure("error", foreground="#c62828")
        self.tree.tag_configure("skipped", foreground="#666666")
        self.tree.tag_configure("working", foreground="#0b57d0")
        if _DND_OK:
            for w in (self.tree, lf, r):
                try:
                    w.drop_target_register(DND_FILES)
                    w.dnd_bind("<<Drop>>", self._on_drop)
                except Exception:
                    pass

        # ---- progress -----------------------------------------------------------
        pf = ttk.Frame(r)
        pf.pack(fill="x", padx=10, pady=2)
        self.lbl_current = ttk.Label(pf, text="Sẵn sàng.")
        self.lbl_current.pack(anchor="w")
        self.pb = ttk.Progressbar(pf, mode="determinate", maximum=100)
        self.pb.pack(fill="x", pady=2)
        self.lbl_counts = ttk.Label(pf, text="Tổng: 0   Hoàn thành: 0   Cần kiểm tra: 0   Lỗi: 0   Bỏ qua: 0")
        self.lbl_counts.pack(anchor="w")

        # ---- log ----------------------------------------------------------------
        self.txt_log = tk.Text(r, height=5, wrap="word", state="disabled", font=("Consolas", 9))
        self.txt_log.pack(fill="x", padx=10, pady=(2, 4))

        # ---- buttons ------------------------------------------------------------
        bf = ttk.Frame(r)
        bf.pack(fill="x", padx=10, pady=(0, 10))
        ttk.Button(bf, text="Chẩn đoán", command=self.show_diagnostics).pack(side="left", padx=2)
        ttk.Button(bf, text="Xem chi tiết lỗi", command=self.show_errors).pack(side="left", padx=2)
        ttk.Button(bf, text="Mở file kết quả", command=self.open_output_file).pack(side="left", padx=2)
        ttk.Button(bf, text="Mở thư mục kết quả", command=self.open_output_folder).pack(side="left", padx=2)
        self.btn_start = ttk.Button(bf, text="BẮT ĐẦU TRÍCH XUẤT", style="Start.TButton", command=self.start)
        self.btn_start.pack(side="right", padx=2)
        self.btn_stop = ttk.Button(bf, text="Dừng sau file hiện tại", command=self.stop, state="disabled")
        self.btn_stop.pack(side="right", padx=2)

    # ------------------------------------------------------------------
    # Config
    # ------------------------------------------------------------------
    def _load_config(self) -> None:
        c = self.cfg
        self.var_server.set(c.ollama_server)
        self.var_model.set(c.model)
        if c.model:
            self.cb_model["values"] = [c.model]
        self.var_template.set(c.last_template)
        self.var_folder.set(c.last_report_folder)
        out = c.last_output_file or (str(Path(c.last_output_folder) / DEFAULT_OUTPUT_NAME) if c.last_output_folder else "")
        self.var_output.set(out)
        self.var_force.set(bool(c.force_reprocess))
        if c.last_report_folder and Path(c.last_report_folder).is_dir():
            self.scan_reports(silent=True)

    def _save_config(self) -> None:
        c = self.cfg
        c.ollama_server = normalize_ollama_url(self.var_server.get())
        c.model = self.var_model.get().strip()
        c.last_template = self.var_template.get().strip()
        c.last_report_folder = self.var_folder.get().strip()
        out = self.var_output.get().strip()
        c.last_output_file = out
        c.last_output_folder = str(Path(out).parent) if out else c.last_output_folder
        c.force_reprocess = bool(self.var_force.get())
        try:
            c.save()
        except Exception as e:  # noqa: BLE001
            self.log(f"Không lưu được config.json: {e}")

    # ------------------------------------------------------------------
    # Actions
    # ------------------------------------------------------------------
    def log(self, msg: str) -> None:
        self.txt_log.configure(state="normal")
        self.txt_log.insert("end", msg + "\n")
        self.txt_log.see("end")
        self.txt_log.configure(state="disabled")

    def browse_template(self) -> None:
        p = filedialog.askopenfilename(title="Chọn form Excel", filetypes=[("Excel", "*.xlsx *.xlsm"), ("Tất cả", "*.*")])
        if p:
            self.var_template.set(p)
            self._save_config()

    def browse_folder(self) -> None:
        p = filedialog.askdirectory(title="Chọn thư mục báo cáo")
        if p:
            self.var_folder.set(p)
            self._save_config()
            self.scan_reports()

    def browse_output(self) -> None:
        init = self.var_output.get() or DEFAULT_OUTPUT_NAME
        p = filedialog.asksaveasfilename(title="File kết quả", defaultextension=".xlsx",
                                         initialfile=Path(init).name,
                                         initialdir=str(Path(init).parent) if init else None,
                                         filetypes=[("Excel", "*.xlsx")], confirmoverwrite=False)
        if p:
            self.var_output.set(p)
            self._save_config()

    def set_files(self, files: List[Path]) -> None:
        self.files = files
        self.tree.delete(*self.tree.get_children())
        self.file_items.clear()
        for i, p in enumerate(files):
            iid = self.tree.insert("", "end", values=(i + 1, self._display_name(p), "○ Đang chờ", ""))
            self.file_items[i] = iid
        self.pb["value"] = 0
        self._update_counts(BatchSummary(total=len(files)))
        self.lbl_current.configure(text=f"Đã tìm thấy {len(files)} báo cáo .pptx")

    def _display_name(self, p: Path) -> str:
        folder = self.var_folder.get().strip()
        try:
            if folder:
                return str(p.relative_to(folder))
        except ValueError:
            pass
        return p.name

    def scan_reports(self, silent: bool = False) -> None:
        folder = self.var_folder.get().strip()
        if not folder or not Path(folder).exists():
            if not silent:
                messagebox.showwarning(APP_NAME, "Thư mục báo cáo không tồn tại.")
            return
        files = scan_inputs([folder])
        self.set_files(files)
        self.log(f"Quét {folder}: {len(files)} file .pptx (kể cả thư mục con)")
        if not files and not silent:
            messagebox.showinfo(APP_NAME, "Không tìm thấy file .pptx nào trong thư mục (và các thư mục con).")

    def _on_drop(self, event) -> None:
        paths = parse_dnd_paths(event.data)
        if not paths:
            return
        dirs = [p for p in paths if Path(p).is_dir()]
        if len(dirs) == 1 and len(paths) == 1:
            self.var_folder.set(dirs[0])
        elif paths:
            common = os.path.commonpath([str(Path(p).parent if Path(p).is_file() else p) for p in paths])
            self.var_folder.set(common)
        files = scan_inputs(paths)
        self.set_files(files)
        self.log(f"Kéo thả: {len(paths)} mục → {len(files)} file .pptx")
        self._save_config()

    # ---- Ollama -----------------------------------------------------------------
    def test_connection(self) -> None:
        server = normalize_ollama_url(self.var_server.get())
        self.var_server.set(server)
        self.lbl_conn.configure(text="● Đang kiểm tra…", style="Bad.TLabel")

        def work():
            try:
                info = OllamaClient(server).test_connection()
                self.queue.put(("conn_ok", info))
            except OllamaError as e:
                self.queue.put(("conn_fail", str(e)))
        threading.Thread(target=work, daemon=True).start()

    def _apply_models(self, info: Dict) -> None:
        models = info.get("models", [])
        self.cb_model["values"] = models
        cur = self.var_model.get().strip()
        if cur not in models:
            pick = preferred_model(models) or (models[0] if models else "")
            self.var_model.set(pick)
        qwen4 = [m for m in models if "qwen" in m.lower() and "4b" in m.lower()]
        self.lbl_conn.configure(text=f"● Đã kết nối ({len(models)} model)", style="Ok.TLabel")
        self.log(f"Kết nối Ollama OK: {info['server']} – model: {', '.join(models) or '(không có)'}")
        if not models:
            messagebox.showwarning(APP_NAME, "Máy chủ Ollama không có model nào. Hãy chạy: ollama pull qwen3:4b")
        elif len(qwen4) > 1:
            self.log(f"Có {len(qwen4)} model Qwen 4B, hãy chọn trong danh sách: {', '.join(qwen4)}")
        self._save_config()

    # ---- batch -------------------------------------------------------------------
    def start(self) -> None:
        if self.processor and self.processor.is_running():
            return
        if not self.files:
            self.scan_reports()
            if not self.files:
                return
        template = self.var_template.get().strip()
        output = self.var_output.get().strip()
        if not template or not Path(template).exists():
            messagebox.showwarning(APP_NAME, "Hãy chọn form Excel (file .xlsx mẫu).")
            return
        if not output:
            messagebox.showwarning(APP_NAME, "Hãy chọn file kết quả.")
            return
        if Path(output).resolve() == Path(template).resolve():
            messagebox.showwarning(APP_NAME, "File kết quả không được trùng với form Excel gốc.")
            return
        self._save_config()
        server = normalize_ollama_url(self.var_server.get())
        model = self.var_model.get().strip()
        use_ai = bool(model)
        if not use_ai:
            if not messagebox.askyesno(APP_NAME, "Chưa chọn model AI. Tiếp tục với nhận diện theo từ khoá (không dùng AI)?"):
                return
        opts = BatchOptions(files=list(self.files), template=Path(template), output_file=Path(output),
                            ollama_server=server, model=model, force_reprocess=bool(self.var_force.get()),
                            request_timeout=self.cfg.request_timeout, use_ollama=use_ai)
        self.error_details = []
        self.summary = None
        for i in self.file_items:
            self._set_row(i, "waiting", "")
        self.processor = BatchProcessor(
            opts,
            on_file=lambda i, s, d: self.queue.put(("file", i, s, d)),
            on_batch=lambda d, t: self.queue.put(("batch", d, t)),
            on_done=lambda s: self.queue.put(("done", s)),
            on_log=lambda m: self.queue.put(("log", m)),
        )
        self.btn_start.configure(state="disabled")
        self.btn_stop.configure(state="normal")
        self.pb["value"] = 0
        self.processor.start()

    def stop(self) -> None:
        if self.processor and self.processor.is_running():
            self.processor.request_stop()
            self.btn_stop.configure(state="disabled")
            self.log("Sẽ dừng sau khi xử lý xong file hiện tại…")

    # ------------------------------------------------------------------
    # Queue / UI updates (GUI thread only)
    # ------------------------------------------------------------------
    def _poll_queue(self) -> None:
        try:
            while True:
                ev = self.queue.get_nowait()
                kind = ev[0]
                if kind == "file":
                    _, i, stage, detail = ev
                    self._set_row(i, stage, detail)
                    name = self.files[i].name if i < len(self.files) else ""
                    if stage in ("reading", "analyzing", "extracting_qpn", "extracting_images", "writing_excel"):
                        self.lbl_current.configure(text=f"[{i + 1}/{len(self.files)}] {name} – {STAGE_LABELS_VI[stage]} {detail}")
                    if stage == "error":
                        self.error_details.append(f"{name}: {detail}")
                    elif stage == "needs_review":
                        self.error_details.append(f"{name} (cần kiểm tra): {detail}")
                elif kind == "batch":
                    _, done, total = ev
                    self.pb["value"] = (done / total * 100) if total else 0
                    if self.processor:
                        self._update_counts(self.processor.summary)
                elif kind == "done":
                    self._on_done(ev[1])
                elif kind == "log":
                    self.log(ev[1])
                elif kind == "conn_ok":
                    self._apply_models(ev[1])
                elif kind == "conn_fail":
                    self.lbl_conn.configure(text="● Chưa kết nối", style="Bad.TLabel")
                    self.log(ev[1])
                    messagebox.showerror(APP_NAME, ev[1])
        except queue.Empty:
            pass
        self.root.after(100, self._poll_queue)

    def _set_row(self, i: int, stage: str, detail: str) -> None:
        iid = self.file_items.get(i)
        if not iid:
            return
        label = f"{STATUS_ICON.get(stage, '')} {STAGE_LABELS_VI.get(stage, stage)}"
        tag = stage if stage in ("completed", "needs_review", "error", "skipped") else ("working" if stage != "waiting" else "")
        self.tree.item(iid, values=(i + 1, self._display_name(self.files[i]), label, detail), tags=(tag,) if tag else ())
        if stage != "waiting":
            self.tree.see(iid)

    def _update_counts(self, s: BatchSummary) -> None:
        self.lbl_counts.configure(text=f"Tổng: {s.total}   Hoàn thành: {s.completed}   Cần kiểm tra: {s.needs_review}"
                                       f"   Lỗi: {s.failed}   Bỏ qua: {s.skipped}")

    def _on_done(self, s: BatchSummary) -> None:
        self.summary = s
        self._update_counts(s)
        self.pb["value"] = 100 if not s.stopped else self.pb["value"]
        self.btn_start.configure(state="normal")
        self.btn_stop.configure(state="disabled")
        self.lbl_current.configure(text="Đã dừng." if s.stopped else "Hoàn tất.")
        msg = (f"Tổng: {s.total}\nHoàn thành: {s.completed}\nCần kiểm tra: {s.needs_review}\n"
               f"Lỗi: {s.failed}\nBỏ qua (đã xử lý): {s.skipped}\n\nFile kết quả:\n{s.output_file}")
        if s.stopped:
            msg = "Đã dừng theo yêu cầu.\n\n" + msg
        messagebox.showinfo(APP_NAME, msg)

    # ------------------------------------------------------------------
    def show_errors(self) -> None:
        lines = list(self.error_details)
        if not lines:
            messagebox.showinfo(APP_NAME, "Chưa có lỗi nào.")
            return
        win = tk.Toplevel(self.root)
        win.title("Chi tiết lỗi / cần kiểm tra")
        win.geometry("760x420")
        t = tk.Text(win, wrap="word", font=("Consolas", 9))
        t.pack(fill="both", expand=True)
        t.insert("end", "\n".join(lines))
        out = self.var_output.get().strip()
        if out:
            t.insert("end", f"\n\nXem thêm: {Path(out).parent / 'logs' / 'errors.log'}")
        t.configure(state="disabled")

    def show_diagnostics(self) -> None:
        self._save_config()
        win = tk.Toplevel(self.root)
        win.title("Chẩn đoán hệ thống")
        win.geometry("760x360")
        t = tk.Text(win, wrap="word", font=("Consolas", 10))
        t.pack(fill="both", expand=True)
        t.insert("end", "Đang kiểm tra…")
        out = self.var_output.get().strip()

        def work():
            rows = run_diagnostics(self.cfg, self.var_template.get().strip(),
                                   str(Path(out).parent) if out else "")
            self.root.after(0, lambda: (t.delete("1.0", "end"), t.insert("end", format_diagnostics(rows))))
        threading.Thread(target=work, daemon=True).start()

    def open_output_file(self) -> None:
        out = self.var_output.get().strip()
        if out and Path(out).exists():
            open_path(out)
        else:
            messagebox.showinfo(APP_NAME, "Chưa có file kết quả.")

    def open_output_folder(self) -> None:
        out = self.var_output.get().strip()
        folder = str(Path(out).parent) if out else ""
        if folder and Path(folder).exists():
            open_path(folder)
        else:
            messagebox.showinfo(APP_NAME, "Chưa có thư mục kết quả.")

    def _on_close(self) -> None:
        if self.processor and self.processor.is_running():
            if not messagebox.askyesno(APP_NAME, "Đang xử lý. Thoát sẽ dừng ngay lập tức (file hiện tại có thể không được ghi). Thoát?"):
                return
        self._save_config()
        self.root.destroy()

    def run(self) -> None:
        self.root.mainloop()


def launch() -> None:
    ReportExtractorApp().run()
