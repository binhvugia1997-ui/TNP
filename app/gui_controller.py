"""Headless controller behind the Tkinter GUI.

Everything the window needs (validation, report discovery, start/stop state, status
mapping, result rows, summary, diagnostics, settings persistence) lives here without any
Tk dependency so it can be unit-tested on a machine without a display.  The Tk view only
renders ``rows`` / ``progress`` and forwards button clicks.  The processing itself is the
SAME production pipeline as the CLI (``BatchProcessor``); no second implementation.
"""
from __future__ import annotations

import queue
import threading
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional, Tuple

from .batch_processor import STAGE_LABELS_VI, BatchOptions, BatchProcessor, BatchSummary, format_file_diagnostics
from .config import DEFAULT_MODEL, DEFAULT_OLLAMA, AppConfig, normalize_ollama_url
from .excel_writer import validate_template
from .extractor import management_number_from_filename
from .logger import FileResult
from .ollama_client import OllamaClient, OllamaError
from .scanner import scan_inputs

DEFAULT_OUTPUT_NAME = "Kiem_chung_Ket_qua.xlsx"

# final statuses shown in the result table (Vietnamese)
STATUS_VI: Dict[str, str] = {
    "waiting": "Đang chờ",
    "completed": "Hoàn thành",
    "needs_review": "Cần kiểm tra",
    "not_written": "Không tìm thấy Management Number",
    "error": "Lỗi",
    "skipped": "Bỏ qua",
}
FINAL_STATUSES = ("completed", "needs_review", "not_written", "error", "skipped")
WORKING_STAGES = ("reading", "analyzing", "analyzing_heuristic", "extracting", "extracting_qpn",
                  "extracting_images", "writing_excel")


def status_label(stage: str) -> str:
    """Vietnamese label for a final status or an in-progress pipeline stage."""
    return STATUS_VI.get(stage) or STAGE_LABELS_VI.get(stage) or stage


def default_output_path(report_folder: str) -> str:
    return str(Path(report_folder) / DEFAULT_OUTPUT_NAME) if report_folder else ""


def format_elapsed(seconds: float) -> str:
    seconds = int(round(seconds))
    h, rem = divmod(seconds, 3600)
    m, s = divmod(rem, 60)
    return f"{h}:{m:02d}:{s:02d}" if h else f"{m}:{s:02d}"


@dataclass
class RowState:
    index: int
    path: Path
    management_number: str = ""
    vendor: str = ""
    model: str = ""
    item: str = ""
    stage: str = "waiting"          # pipeline stage or final status key
    note: str = ""

    @property
    def status_vi(self) -> str:
        return status_label(self.stage)

    @property
    def is_final(self) -> bool:
        return self.stage in FINAL_STATUSES

    def as_values(self) -> Tuple[Any, ...]:
        return (self.index + 1, self.management_number, self.path.name, self.vendor.replace("\n", " / "),
                self.model, self.item, self.status_vi, self.note)


@dataclass
class Progress:
    done: int = 0
    total: int = 0
    current_index: Optional[int] = None
    current_stage: str = ""
    current_detail: str = ""

    @property
    def percent(self) -> float:
        return (self.done / self.total * 100.0) if self.total else 0.0

    @property
    def text(self) -> str:
        if self.total == 0:
            return ""
        cur = min(self.total, (self.current_index or 0) + 1) if self.current_index is not None else self.done
        return f"Đang xử lý: {cur} / {self.total}"


@dataclass
class UiEvent:
    kind: str                      # row | progress | log | done | ollama
    payload: Any = None


class GuiController:
    """State machine + adapter between the GUI widgets and the production pipeline."""

    def __init__(self, cfg: Optional[AppConfig] = None,
                 processor_factory: Callable[..., BatchProcessor] = BatchProcessor,
                 client_factory: Callable[..., OllamaClient] = OllamaClient,
                 config_path: Optional[Path] = None):
        self.cfg = cfg or AppConfig.load(config_path)
        self._config_path = config_path
        self._processor_factory = processor_factory
        self._client_factory = client_factory
        # editable settings (mirrors of the widgets)
        self.report_folder: str = self.cfg.last_report_folder or ""
        self.template: str = self.cfg.last_template or ""
        self.output: str = self.cfg.last_output_file or ""
        self.server: str = normalize_ollama_url(self.cfg.ollama_server or DEFAULT_OLLAMA)
        self.model: str = self.cfg.model or DEFAULT_MODEL
        self.force_reprocess: bool = bool(self.cfg.force_reprocess)
        # runtime state
        self.state: str = "idle"                 # idle | running | stopping
        self.files: List[Path] = []
        self.rows: List[RowState] = []
        self.progress = Progress()
        self.summary: Optional[BatchSummary] = None
        self.processor: Optional[BatchProcessor] = None
        self.ollama_status: str = ""
        self.ollama_ok: Optional[bool] = None
        self.log_lines: List[str] = []
        self.started_at: Optional[float] = None
        self.finished_at: Optional[float] = None
        self._queue: "queue.Queue[UiEvent]" = queue.Queue()

    # ------------------------------------------------------------------ inputs
    def set_report_folder(self, folder: str) -> int:
        """Select the report folder, discover PPT/PPTX recursively, propose the output path."""
        self.report_folder = folder.strip()
        if not self.output or Path(self.output).name == DEFAULT_OUTPUT_NAME:
            self.output = default_output_path(self.report_folder)
        return self.discover()

    def discover(self) -> int:
        folder = Path(self.report_folder) if self.report_folder else None
        self.files = scan_inputs([folder]) if folder and folder.is_dir() else []
        self.rows = [RowState(index=i, path=p, management_number=management_number_from_filename(p.name))
                     for i, p in enumerate(self.files)]
        self.progress = Progress(total=len(self.files))
        self.summary = None
        return len(self.files)

    def set_template(self, path: str) -> None:
        self.template = path.strip()

    def set_output(self, path: str) -> None:
        self.output = path.strip()

    def set_ollama(self, server: str, model: str) -> None:
        self.server = normalize_ollama_url(server) if server.strip() else ""
        self.model = model.strip()

    # ------------------------------------------------------------------ validation
    def validate(self) -> List[str]:
        """Vietnamese problems that prevent a run (empty list = OK). Ollama is NOT required."""
        errs: List[str] = []
        if not self.report_folder or not Path(self.report_folder).is_dir():
            errs.append("Thư mục báo cáo không tồn tại.")
        elif not self.files:
            errs.append("Không tìm thấy file .ppt/.pptx nào trong thư mục báo cáo.")
        tpl = Path(self.template) if self.template else None
        if tpl is None or not tpl.is_file():
            errs.append("File Kiểm chứng (.xlsx) không tồn tại.")
        out = Path(self.output) if self.output else None
        if out is None:
            errs.append("Chưa chọn file kết quả.")
        else:
            if out.suffix.lower() not in (".xlsx", ".xlsm"):
                errs.append("File kết quả phải có đuôi .xlsx.")
            if tpl is not None and tpl.exists():
                try:
                    same = out.resolve() == tpl.resolve()
                except OSError:
                    same = str(out) == str(tpl)
                if same:
                    errs.append("File kết quả không được trùng với file Kiểm chứng gốc.")
            if not self._writable(out):
                errs.append(f"Không ghi được vào thư mục kết quả: {out.parent}")
        if tpl is not None and tpl.is_file():
            ok, msg = validate_template(tpl)
            if not ok:
                errs.append(f"File Kiểm chứng không hợp lệ: {msg}")
        return errs

    @staticmethod
    def _writable(out: Path) -> bool:
        try:
            out.parent.mkdir(parents=True, exist_ok=True)
            probe = out.parent / ".re_write_test.tmp"
            probe.write_text("x", encoding="utf-8")
            probe.unlink()
            if out.exists():
                with open(out, "ab"):
                    pass
            return True
        except OSError:
            return False

    # ------------------------------------------------------------------ Ollama
    def check_ollama(self) -> Tuple[bool, str]:
        """Same production client; returns (ok, Vietnamese message)."""
        server = normalize_ollama_url(self.server) if self.server else ""
        model = self.model.strip()
        if not server:
            self.ollama_ok, self.ollama_status = False, "Ollama: chưa nhập địa chỉ máy chủ"
            return False, self.ollama_status
        try:
            info = self._client_factory(server, timeout=int(self.cfg.request_timeout or 180)).test_connection()
        except OllamaError as e:
            self.ollama_ok, self.ollama_status = False, f"Ollama: không kết nối được – {e}"
            return False, self.ollama_status
        models = info.get("models", [])
        if not model:
            self.ollama_ok, self.ollama_status = False, "Ollama: chưa chọn model (ví dụ qwen3:4b)"
        elif model not in models:
            self.ollama_ok = False
            self.ollama_status = (f"Ollama: model '{model}' không có trên máy chủ"
                                  + (f" (có: {', '.join(models)})" if models else " (chưa có model nào, chạy: ollama pull qwen3:4b)"))
        else:
            self.ollama_ok, self.ollama_status = True, f"Ollama: Sẵn sàng — {model}"
        self.available_models = models
        return bool(self.ollama_ok), self.ollama_status

    def check_ollama_async(self) -> None:
        def work():
            ok, msg = self.check_ollama()
            self._queue.put(UiEvent("ollama", (ok, msg)))
        threading.Thread(target=work, daemon=True).start()

    # ------------------------------------------------------------------ run control
    def can_start(self) -> bool:
        return self.state == "idle"

    def build_options(self, use_ollama: bool) -> BatchOptions:
        return BatchOptions(files=list(self.files), template=Path(self.template), output_file=Path(self.output),
                            ollama_server=self.server, model=self.model if use_ollama else "",
                            force_reprocess=self.force_reprocess, request_timeout=int(self.cfg.request_timeout or 180),
                            use_ollama=use_ollama and bool(self.model),
                            fill_temporary_column=bool(self.cfg.fill_temporary_column),
                            vendors=list(self.cfg.vendors or []), row_mode=self.cfg.row_mode or "match")

    def start(self, use_ollama: bool = True, in_thread: bool = True) -> bool:
        """Start the production batch; returns False when validation fails or already running."""
        if not self.can_start():
            return False
        if self.validate():
            return False
        self.save_settings()
        for r in self.rows:
            r.stage, r.note, r.vendor, r.model, r.item = "waiting", "", "", "", ""
        self.progress = Progress(total=len(self.files))
        self.summary = None
        self.started_at, self.finished_at = time.perf_counter(), None
        self.processor = self._processor_factory(
            self.build_options(use_ollama),
            on_file=lambda i, s, d: self._queue.put(UiEvent("row", (i, s, d))),
            on_batch=lambda d, t: self._queue.put(UiEvent("progress", (d, t))),
            on_done=lambda s: self._queue.put(UiEvent("done", s)),
            on_log=lambda m: self._queue.put(UiEvent("log", m)),
        )
        self.state = "running"
        if in_thread:
            self.processor.start()
        else:
            self.processor.run()
        return True

    def request_stop(self) -> bool:
        """Stop AFTER the current report (never in the middle of an Excel write)."""
        if self.state != "running" or not self.processor:
            return False
        self.processor.request_stop()
        self.state = "stopping"
        self.log_lines.append("Sẽ dừng sau khi xử lý xong báo cáo hiện tại…")
        return True

    def is_running(self) -> bool:
        return self.state in ("running", "stopping")

    def controls_enabled(self) -> bool:
        """Configuration widgets are editable only while idle."""
        return self.state == "idle"

    # ------------------------------------------------------------------ events
    def pump(self, max_events: int = 500) -> List[UiEvent]:
        """Drain worker events (call from the GUI thread); returns the events applied."""
        applied: List[UiEvent] = []
        for _ in range(max_events):
            try:
                ev = self._queue.get_nowait()
            except queue.Empty:
                break
            self.apply_event(ev)
            applied.append(ev)
        return applied

    def apply_event(self, ev: UiEvent) -> None:
        if ev.kind == "row":
            i, stage, detail = ev.payload
            if 0 <= i < len(self.rows):
                row = self.rows[i]
                row.stage = stage
                if stage in FINAL_STATUSES:
                    row.note = detail or ""
                    self._fill_row_from_result(row)
                else:
                    self.progress.current_index, self.progress.current_stage = i, stage
                    self.progress.current_detail = detail or ""
        elif ev.kind == "progress":
            done, total = ev.payload
            self.progress.done, self.progress.total = done, total
            if self.processor:
                self.summary = self.processor.summary
        elif ev.kind == "log":
            self.log_lines.append(str(ev.payload))
        elif ev.kind == "done":
            self.summary = ev.payload
            self.finished_at = time.perf_counter()
            self.state = "idle"
            self.progress.current_index, self.progress.current_stage = None, "done"
            for r in self.rows:                      # rows the batch never reached (stopped early)
                if not r.is_final and r.stage != "waiting":
                    r.stage = "error"
        elif ev.kind == "ollama":
            self.ollama_ok, self.ollama_status = ev.payload

    def _fill_row_from_result(self, row: RowState) -> None:
        fr = self.result_for(row.index)
        if fr is None:
            return
        row.management_number = fr.management_number or row.management_number
        row.vendor, row.model, row.item = fr.vendor, fr.model, fr.item
        if fr.status == "error" and fr.error:
            row.note = fr.error
        elif fr.review_reasons and not row.note:
            row.note = "; ".join(fr.review_reasons)

    def result_for(self, index: int) -> Optional[FileResult]:
        if not self.processor or index >= len(self.files):
            return None
        src = str(self.files[index])
        return next((r for r in self.processor.results if r.source_file == src), None)

    # ------------------------------------------------------------------ summary / diagnostics
    def elapsed_seconds(self) -> float:
        if self.started_at is None:
            return 0.0
        end = self.finished_at if self.finished_at is not None else time.perf_counter()
        return end - self.started_at

    def summary_lines(self) -> List[str]:
        s = self.summary or BatchSummary(total=len(self.files))
        lines = [f"Tổng: {s.total}", f"Hoàn thành: {s.completed}", f"Cần kiểm tra: {s.needs_review}",
                 f"Không tìm thấy Management Number: {s.not_written}", f"Lỗi: {s.failed}", f"Bỏ qua: {s.skipped}"]
        if self.started_at is not None:
            lines.append(f"Thời gian: {format_elapsed(self.elapsed_seconds())}")
        if s.stopped:
            lines.insert(0, "Đã dừng theo yêu cầu.")
        return lines

    def summary_text(self) -> str:
        return "\n".join(self.summary_lines())

    def counts_text(self) -> str:
        s = self.summary or (self.processor.summary if self.processor else None) or BatchSummary(total=len(self.files))
        return (f"Tổng: {s.total}   Hoàn thành: {s.completed}   Cần kiểm tra: {s.needs_review}   "
                f"Không tìm thấy Management Number: {s.not_written}   Lỗi: {s.failed}   Bỏ qua: {s.skipped}")

    def output_file(self) -> Optional[Path]:
        p = Path(self.summary.output_file) if self.summary and self.summary.output_file else (Path(self.output) if self.output else None)
        return p if p and p.exists() else None

    def output_folder(self) -> Optional[Path]:
        p = Path(self.output).parent if self.output else None
        return p if p and p.exists() else None

    def log_file(self) -> Optional[Path]:
        folder = self.output_folder()
        if not folder:
            return None
        for name in ("app.log", "errors.log", "batch_result.json"):
            p = folder / "logs" / name
            if p.exists():
                return p
        return None

    def diagnostics_for(self, index: int) -> Optional[Dict[str, Any]]:
        """Human-readable diagnostics of one processed report (no hidden model reasoning)."""
        fr = self.result_for(index)
        if fr is None:
            return None
        timing = next((n for n in fr.classifier_notes if n.startswith("Ollama ok in") or "heuristic fallback" in n), "")
        return {
            "Tên file": Path(fr.source_file).name,
            "Trạng thái": status_label(fr.status) + (f" (dòng Excel {fr.excel_row})" if fr.excel_row else ""),
            "Management Number": fr.management_number or "(trống)",
            "Ngày phát sinh": fr.occurrence_date or "(trống)",
            "Vendor": fr.vendor.replace("\n", " / ") if fr.vendor else "(trống)",
            "Model": fr.model or "(trống)",
            "Item": fr.item or "(trống)",
            "Slide QPN": f"{fr.qpn_slide or '(không thấy)'}" + (f"  [{fr.qpn_source}]" if fr.qpn_source else ""),
            "Slide nguyên nhân": str(fr.cause_slides or "(không thấy)"),
            "Slide xử lý tạm thời": str(fr.temporary_slides or "-"),
            "Slide đối sách cải tiến": str(fr.improvement_slides or "(không thấy)"),
            "Slide ảnh cải tiến": str(fr.improvement_image_slides or "(không có)"),
            "Bộ phân loại": fr.classifier or "-",
            "Độ tin cậy AI": f"{fr.confidence:.2f}" if fr.confidence is not None else "-",
            "Lý do cần kiểm tra": "\n".join(fr.review_reasons) if fr.review_reasons else "-",
            "Lỗi": fr.error or "-",
            "Thời gian Ollama": timing or "-",
            "Ghi chú phân loại": "\n".join(fr.classifier_notes) if fr.classifier_notes else "-",
            "_text": format_file_diagnostics(fr),
        }

    # ------------------------------------------------------------------ settings
    def save_settings(self) -> None:
        c = self.cfg
        c.last_report_folder = self.report_folder
        c.last_template = self.template
        c.last_output_file = self.output
        c.last_output_folder = str(Path(self.output).parent) if self.output else c.last_output_folder
        c.ollama_server = normalize_ollama_url(self.server) if self.server else DEFAULT_OLLAMA
        c.model = self.model
        c.force_reprocess = bool(self.force_reprocess)
        try:
            c.save(self._config_path)
        except OSError as e:
            self.log_lines.append(f"Không lưu được config.json: {e}")
