"""Headless controller behind the Tkinter GUI.

Everything the window needs (validation, report discovery, start/stop state, status
mapping, result rows, summary, diagnostics, settings persistence) lives here without any
Tk dependency so it can be unit-tested on a machine without a display.  The Tk view only
renders ``rows`` / ``progress`` and forwards button clicks.  The processing itself is the
SAME production pipeline as the CLI (``BatchProcessor``); no second implementation.
"""
from __future__ import annotations

import logging
import queue
import threading
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional, Tuple

from .batch_processor import STAGE_LABELS_VI, BatchOptions, BatchProcessor, BatchSummary, format_file_diagnostics
from .config import DEFAULT_MODEL, DEFAULT_OLLAMA, AppConfig, endpoint_problem, normalize_ollama_url, split_endpoint
from .excel_writer import validate_template
from .extractor import management_number_from_filename
from .logger import FileResult
from .ollama_client import OllamaClient, OllamaError, preferred_model
from .prescan import (AUTO_PERIOD_FAIL_VI, PERIOD_DIFFERS_VI, ALL_PERIOD, PreScanResult, ProcessingPeriod,
                      auto_period_from_excel, month_period, range_period)
from .scanner import scan_inputs

LOG = logging.getLogger("report_extractor.gui")
DEFAULT_OUTPUT_NAME = "Kiem_chung_Ket_qua.xlsx"

# final statuses shown in the result table (Vietnamese)
STATUS_VI: Dict[str, str] = {
    "waiting": "Đang chờ",
    "completed": "Hoàn thành",
    "needs_review": "Cần kiểm tra",
    "not_written": "Không tìm thấy Management Number",
    "error": "Lỗi",
    "skipped": "Bỏ qua — đã cập nhật",
    "outside_period": "Bỏ qua ngoài thời gian xử lý",
    "source_duplicate": "Trùng Management Number trong folder",
    "fast_skip": "Bỏ qua nhanh — đã xử lý gần đây",
}
PRESCAN_STATUSES = ("outside_period", "source_duplicate", "fast_skip")
FINAL_STATUSES = ("completed", "needs_review", "not_written", "error", "skipped") + PRESCAN_STATUSES
PERIOD_MODES = ("auto", "month", "range", "all")
PERIOD_MODE_VI = {"auto": "Tự động theo file Excel", "month": "Chọn tháng", "range": "Khoảng thời gian", "all": "Tất cả"}
WORKING_STAGES = ("reading", "analyzing", "analyzing_heuristic", "extracting", "extracting_qpn",
                  "extracting_images", "writing_excel")

# Deterministic share of ONE report that is considered done when a REAL pipeline stage is reached
# (order = order in which BatchProcessor emits them).  Nothing is interpolated while a stage runs:
# during a long Qwen call the value stays at the "analyzing" boundary until the next real event.
STAGE_WEIGHTS: Dict[str, float] = {
    "reading": 0.05,
    "analyzing": 0.15,
    "analyzing_heuristic": 0.15,
    "extracting": 0.45,
    "extracting_qpn": 0.55,
    "extracting_images": 0.70,
    "writing_excel": 0.85,
}
MIN_ETA_SAMPLES = 1          # ETA is shown once at least this many reports finished with measured durations
ETA_WINDOW = 10              # rolling window of report durations


def status_label(stage: str) -> str:
    """Vietnamese label for a final status or an in-progress pipeline stage."""
    return STATUS_VI.get(stage) or STAGE_LABELS_VI.get(stage) or stage


def default_output_path(report_folder: str) -> str:
    return str(Path(report_folder) / DEFAULT_OUTPUT_NAME) if report_folder else ""


def format_elapsed(seconds: float) -> str:
    """MM:SS under one hour, HH:MM:SS from one hour on."""
    seconds = max(0, int(seconds))
    h, rem = divmod(seconds, 3600)
    m, s = divmod(rem, 60)
    return f"{h:02d}:{m:02d}:{s:02d}" if h else f"{m:02d}:{s:02d}"


def progress_bar_text(percent: float, width: int = 20) -> str:
    """Text rendering of the bar ("████████░░░░ 36%") – same value as the ttk bar."""
    pct = max(0.0, min(100.0, percent))
    filled = int(round(pct / 100.0 * width))
    return "█" * filled + "░" * (width - filled) + f" {int(pct)}%"


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
    """Single source of truth for the progress display (text, bar, percentage)."""
    done: int = 0                       # reports in a terminal state
    total: int = 0
    current_index: Optional[int] = None
    current_stage: str = ""
    current_detail: str = ""
    current_fraction: float = 0.0       # share of the current report reached (STAGE_WEIGHTS, monotonic)
    finished: bool = False              # batch ended (completed or stopped)

    @property
    def percent(self) -> float:
        if self.total <= 0:
            return 0.0
        value = (self.done + (self.current_fraction if self.current_index is not None else 0.0)) / self.total * 100.0
        return max(0.0, min(100.0, value))

    @property
    def percent_int(self) -> int:
        return int(self.percent)

    @property
    def text(self) -> str:
        if self.total == 0:
            return ""
        if self.finished:
            return f"Đã xử lý: {self.done} / {self.total} — {self.percent_int}%"
        cur = min(self.total, self.current_index + 1) if self.current_index is not None else min(self.total, self.done + 1)
        return f"Đang xử lý: {cur} / {self.total} — {self.percent_int}%"

    @property
    def bar_text(self) -> str:
        return progress_bar_text(self.percent)

    def reach_stage(self, index: int, stage: str, detail: str = "") -> None:
        if self.current_index != index:
            self.current_index, self.current_fraction = index, 0.0
        self.current_stage, self.current_detail = stage, detail or ""
        self.current_fraction = max(self.current_fraction, STAGE_WEIGHTS.get(stage, 0.0))

    def finish_report(self, index: int) -> None:
        if self.current_index == index:
            self.current_index, self.current_fraction, self.current_stage = None, 0.0, ""


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
        self.host, self.port = split_endpoint(self.cfg.ollama_server or DEFAULT_OLLAMA)
        self.model: str = self.cfg.model or DEFAULT_MODEL
        self.available_models: List[str] = []
        self.force_reprocess: bool = bool(self.cfg.force_reprocess)
        # processing period (fast pre-scan)
        self.period_mode: str = self.cfg.period_mode if self.cfg.period_mode in PERIOD_MODES else "auto"
        self.period_month: str = str(self.cfg.period_month or "")
        self.period_year: str = str(self.cfg.period_year or "")
        self.period_from: str = self.cfg.period_from or ""
        self.period_to: str = self.cfg.period_to or ""
        self.prescan: Optional[PreScanResult] = None
        self.queue_indexes: Optional[set] = None       # indexes that really enter the processing pipeline
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
        self.started_at: Optional[float] = None      # time.monotonic() when the batch really started
        self.finished_at: Optional[float] = None
        self.report_durations: List[float] = []      # measured seconds per finished report (rolling window)
        self._report_started_at: Optional[float] = None
        self._clock: Callable[[], float] = time.monotonic
        self._queue: "queue.Queue[UiEvent]" = queue.Queue()

    # ------------------------------------------------------------------ Ollama endpoint
    @property
    def server(self) -> str:
        """Normalized endpoint used by BOTH the connection test and the production batch."""
        return normalize_ollama_url(f"{self.host}:{self.port}") if str(self.host).strip() else ""

    @server.setter
    def server(self, value: str) -> None:
        self.host, self.port = split_endpoint(value) if (value or "").strip() else ("", 11434)

    @property
    def endpoint_label(self) -> str:
        return f"{self.host}:{self.port}" if str(self.host).strip() else "(chưa nhập)"

    def set_endpoint(self, host: str, port: Any, model: Optional[str] = None) -> str:
        """Accept '127.0.0.1', '192.168.1.50:11434', 'http://192.168.1.50:11434' (host field may carry the
        port / scheme); returns a Vietnamese problem text or '' – takes effect immediately."""
        host = (host or "").strip()
        port_s = str(port or "").strip()
        if "://" in host or "/" in host or ":" in host:
            h, p = split_endpoint(host)
            if ":" in host.split("://", 1)[-1].split("/", 1)[0]:
                port_s = str(p)
            host = h
        problem = endpoint_problem(host, port_s or 11434)
        if problem:
            return problem
        self.host, self.port = host, int(port_s or 11434)
        if model is not None:
            self.model = model.strip()
        self.ollama_ok, self.ollama_status = None, ""        # endpoint changed -> status unknown
        return ""

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
        self.server = server
        self.model = model.strip()

    # ------------------------------------------------------------------ processing period
    def set_period(self, mode: str, month: Any = None, year: Any = None, start: Optional[str] = None,
                   end: Optional[str] = None) -> str:
        """Change the period mode/values; returns a Vietnamese problem text ('' when fine)."""
        if mode not in PERIOD_MODES:
            return f"Chế độ thời gian không hợp lệ: {mode}"
        self.period_mode = mode
        if month is not None:
            self.period_month = str(month).strip()
        if year is not None:
            self.period_year = str(year).strip()
        if start is not None:
            self.period_from = start.strip()
        if end is not None:
            self.period_to = end.strip()
        _, err = self.effective_period()
        return err

    def detected_period(self) -> Optional[ProcessingPeriod]:
        """Month detected from the CURRENT destination / master Excel file name (recomputed every call, so a
        newly selected workbook is never mixed up with the previous one)."""
        for candidate in (self.output, self.template):
            if candidate:
                per = auto_period_from_excel(candidate)
                if per:
                    return per
        return None

    def effective_period(self) -> Tuple[Optional[ProcessingPeriod], str]:
        """(period, error).  Manual month/range always wins over the Excel file name."""
        mode = self.period_mode
        if mode == "all":
            return ALL_PERIOD, ""
        if mode == "auto":
            per = self.detected_period()
            return (per, "") if per else (None, AUTO_PERIOD_FAIL_VI)
        if mode == "month":
            try:
                m, y = int(self.period_month), int(self.period_year)
                if not 1 <= m <= 12:
                    raise ValueError
                if not 2000 <= y <= 2099:
                    return None, f"Năm không hợp lệ: {self.period_year!r}"
                return month_period(y, m), ""
            except (TypeError, ValueError):
                return None, f"Tháng/Năm không hợp lệ: {self.period_month!r}/{self.period_year!r}"
        per, err = range_period(self.period_from, self.period_to)
        return (per, "") if per else (None, err)

    def period_label(self) -> str:
        per, err = self.effective_period()
        return per.label_vi() if per else f"Thời gian xử lý: {err}"

    def period_warning(self) -> str:
        """Non-blocking notice when a manual period differs from the month in the Excel file name."""
        if self.period_mode not in ("month", "range"):
            return ""
        per, _ = self.effective_period()
        det = self.detected_period()
        if per and det and (per.start, per.end) != (det.start, det.end):
            return PERIOD_DIFFERS_VI
        return ""

    # ------------------------------------------------------------------ validation
    def validate(self) -> List[str]:
        """Vietnamese problems that prevent a run (empty list = OK). Ollama is NOT required."""
        errs: List[str] = []
        _, perr = self.effective_period()
        if perr:
            errs.append(perr)
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
    def _classify_ollama_error(self, e: Exception) -> str:
        cause = getattr(e, "__cause__", None)
        name = type(cause).__name__ if cause is not None else type(e).__name__
        text = f"{name} {e}".lower()
        if "timeout" in text or "timed out" in text:
            return "timeout"
        return "unreachable"

    def check_ollama(self, timeout: Optional[int] = None) -> Tuple[bool, str]:
        """Same production client (server reachable → API answers → model present). (ok, message)."""
        server = self.server
        model = self.model.strip()
        where = self.endpoint_label
        if not server:
            self.ollama_ok, self.ollama_status = False, "● Chưa nhập IP / Server của Ollama"
            return False, self.ollama_status
        problem = endpoint_problem(self.host, self.port)
        if problem:
            self.ollama_ok, self.ollama_status = False, f"● {problem}"
            return False, self.ollama_status
        try:
            client = self._client_factory(server, timeout=int(timeout or self.cfg.request_timeout or 180))
            info = client.test_connection()
        except OllamaError as e:
            LOG.warning("Ollama check failed (%s): %s", server, e)
            kind = self._classify_ollama_error(e)
            self.ollama_ok = False
            self.ollama_status = ("● Kết nối Ollama quá thời gian" if kind == "timeout"
                                  else f"● Không kết nối được Ollama tại {where}")
            return False, self.ollama_status
        except Exception as e:  # noqa: BLE001  – never show a traceback to the user
            LOG.exception("Ollama check crashed (%s)", server)
            self.ollama_ok, self.ollama_status = False, f"● Không kết nối được Ollama tại {where} ({type(e).__name__})"
            return False, self.ollama_status
        models = list(info.get("models", []) or [])
        if models:
            self.available_models = models
        if not model:
            self.ollama_ok, self.ollama_status = False, "● Đã kết nối Ollama nhưng chưa chọn model (ví dụ qwen3:4b)"
        elif model not in models:
            self.ollama_ok = False
            self.ollama_status = f"● Đã kết nối Ollama nhưng không tìm thấy model {model}"
        else:
            self.ollama_ok, self.ollama_status = True, f"● Đã kết nối — {model}"
        return bool(self.ollama_ok), self.ollama_status

    def check_ollama_async(self) -> None:
        """Run the check on a worker thread; the result arrives as an 'ollama' event in pump()."""
        def work():
            ok, msg = self.check_ollama()
            self._queue.put(UiEvent("ollama", (ok, msg)))
        threading.Thread(target=work, daemon=True, name="ollama-check").start()

    def refresh_models(self, timeout: Optional[int] = None) -> Tuple[bool, str, List[str]]:
        """Query the configured server for its INSTALLED models (no hard-coded list).

        On failure the user's current model is kept untouched. Returns (ok, message, models)."""
        server = self.server
        if not server or endpoint_problem(self.host, self.port):
            return False, f"● {endpoint_problem(self.host, self.port) or 'Chưa nhập IP / Server của Ollama'}", []
        try:
            info = self._client_factory(server, timeout=int(timeout or self.cfg.request_timeout or 180)).test_connection()
        except Exception as e:  # noqa: BLE001
            LOG.warning("Model discovery failed (%s): %s", server, e)
            kind = self._classify_ollama_error(e) if isinstance(e, OllamaError) else "unreachable"
            msg = ("● Kết nối Ollama quá thời gian – giữ nguyên model hiện tại" if kind == "timeout"
                   else f"● Không lấy được danh sách model từ {self.endpoint_label} – giữ nguyên model hiện tại")
            return False, msg, []
        models = list(info.get("models", []) or [])
        if not models:
            return False, f"● Ollama tại {self.endpoint_label} chưa có model nào (chạy: ollama pull qwen3:4b)", []
        self.available_models = models
        if not self.model.strip():
            self.model = DEFAULT_MODEL if DEFAULT_MODEL in models else (preferred_model(models) or models[0])
        return True, f"● Đã tìm thấy {len(models)} model trên {self.endpoint_label}", models

    def refresh_models_async(self) -> None:
        def work():
            self._queue.put(UiEvent("models", self.refresh_models()))
        threading.Thread(target=work, daemon=True, name="ollama-models").start()

    def ai_status_text(self) -> str:
        """What the program is about to use: 'AI: qwen3:4b @ host:port — Đã kết nối' etc."""
        if not self.server:
            return "AI: Không kết nối — sẽ sử dụng heuristic fallback"
        if self.ollama_ok is True:
            return f"AI: {self.model} @ {self.endpoint_label} — Đã kết nối"
        if self.ollama_ok is False:
            return "AI: Không kết nối — sẽ sử dụng heuristic fallback"
        return f"AI: {self.model or '(chưa chọn model)'} @ {self.endpoint_label} — Chưa kiểm tra"

    def save_ollama_settings(self) -> str:
        """Persist host/port/model with the existing config mechanism (no registry / env / manual JSON)."""
        problem = endpoint_problem(self.host, self.port)
        if problem:
            return problem
        self.save_settings()
        return ""

    # ------------------------------------------------------------------ run control
    def can_start(self) -> bool:
        return self.state == "idle"

    def build_options(self, use_ollama: bool) -> BatchOptions:
        return BatchOptions(files=list(self.files), template=Path(self.template), output_file=Path(self.output),
                            ollama_server=self.server, model=self.model if use_ollama else "",
                            force_reprocess=self.force_reprocess, request_timeout=int(self.cfg.request_timeout or 180),
                            use_ollama=use_ollama and bool(self.model),
                            fill_temporary_column=bool(self.cfg.fill_temporary_column),
                            vendors=list(self.cfg.vendors or []), row_mode=self.cfg.row_mode or "match",
                            period=self.effective_period()[0])

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
        self.prescan, self.queue_indexes = None, None
        self.report_durations = []
        self._report_started_at = None
        self.started_at, self.finished_at = self._clock(), None
        self.log_lines.append(self.period_label())
        if self.period_warning():
            self.log_lines.append(self.period_warning())
        self.processor = self._processor_factory(
            self.build_options(use_ollama),
            on_prescan=lambda r: self._queue.put(UiEvent("prescan", r)),
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

    def _in_queue(self, index: int) -> bool:
        return self.queue_indexes is None or index in self.queue_indexes

    def _count_done(self) -> int:
        return sum(1 for r in self.rows if r.is_final and self._in_queue(r.index))

    def apply_event(self, ev: UiEvent) -> None:
        if ev.kind == "prescan":
            res: PreScanResult = ev.payload
            self.prescan = res
            self.queue_indexes = {it.index for it in res.candidates}
            self.progress = Progress(total=len(self.queue_indexes))     # denominator = real candidates only
            self._report_started_at = self._clock()                      # pre-scan time never enters the ETA
            for line in res.summary_lines_vi():
                self.log_lines.append(line)
        elif ev.kind == "row":
            i, stage, detail = ev.payload
            if 0 <= i < len(self.rows):
                row = self.rows[i]
                was_final = row.is_final
                row.stage = stage
                now = self._clock()
                if not self._in_queue(i):
                    row.note = detail or ""                              # pre-scan rejection: no progress / ETA
                    self._fill_row_from_result(row)
                    return
                if stage in FINAL_STATUSES:
                    row.note = detail or ""
                    self._fill_row_from_result(row)
                    if not was_final:
                        start = self._report_started_at if self._report_started_at is not None else self.started_at
                        if start is not None:
                            self.report_durations.append(max(0.0, now - start))
                            del self.report_durations[:-ETA_WINDOW]
                        self._report_started_at = now          # next report starts right away
                    self.progress.finish_report(i)
                    self.progress.done = self._count_done()
                else:
                    if self.progress.current_index != i or self._report_started_at is None:
                        self._report_started_at = now if self._report_started_at is None else self._report_started_at
                    self.progress.reach_stage(i, stage, detail)
        elif ev.kind == "progress":
            done, total = ev.payload
            self.progress.total = total
            self.progress.done = max(done, self._count_done())
            if self.processor:
                self.summary = self.processor.summary
        elif ev.kind == "log":
            self.log_lines.append(str(ev.payload))
        elif ev.kind == "done":
            self.summary = ev.payload
            self.finished_at = self._clock()
            self.state = "idle"
            for r in self.rows:                      # rows the batch never reached (stopped early)
                if not r.is_final and r.stage != "waiting":
                    r.stage = "error"
            self.progress.current_index, self.progress.current_stage, self.progress.current_fraction = None, "", 0.0
            self.progress.done = self._count_done()
            self.progress.finished = True
        elif ev.kind == "ollama":
            self.ollama_ok, self.ollama_status = ev.payload
        elif ev.kind == "models":
            pass                                     # payload (ok, msg, models) is rendered by the view

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
        """Seconds since the batch actually started (monotonic); frozen once the batch ended."""
        if self.started_at is None:
            return 0.0
        end = self.finished_at if self.finished_at is not None else self._clock()
        return max(0.0, end - self.started_at)

    def elapsed_text(self) -> str:
        if self.started_at is None:
            return ""
        if self.finished_at is not None:
            stopped = bool(self.summary and self.summary.stopped) or self.progress.done < self.progress.total
            return (f"Thời gian đã chạy: {format_elapsed(self.elapsed_seconds())}" if stopped
                    else f"Tổng thời gian: {format_elapsed(self.elapsed_seconds())}")
        return f"Đã chạy: {format_elapsed(self.elapsed_seconds())}"

    def average_report_seconds(self) -> Optional[float]:
        if len(self.report_durations) < MIN_ETA_SAMPLES:
            return None
        return sum(self.report_durations) / len(self.report_durations)

    def eta_seconds(self) -> Optional[float]:
        """Approximate remaining seconds from MEASURED report durations; None when unknown / not running."""
        if not self.is_running() or self.progress.total <= 0:
            return None
        avg = self.average_report_seconds()
        if avg is None:
            return None
        p = self.progress
        remaining_reports = p.total - p.done - (p.current_fraction if p.current_index is not None else 0.0)
        return max(0.0, remaining_reports * avg)

    def eta_text(self) -> str:
        """'Còn khoảng: MM:SS' while running, 'Còn khoảng: Đang tính...' before data exists, '' when not running."""
        if not self.is_running():
            return ""
        eta = self.eta_seconds()
        return "Còn khoảng: Đang tính..." if eta is None else f"Còn khoảng: {format_elapsed(eta)}"

    def status_line(self) -> str:
        """One-line status for the window: progress text + bar + elapsed + ETA (controller state only)."""
        parts = [self.progress.text, self.progress.bar_text, self.elapsed_text(), self.eta_text()]
        return "   ".join(x for x in parts if x)

    def prescan_lines(self) -> List[str]:
        """Pre-scan counters (Vietnamese) – empty before the pre-scan ran."""
        return self.prescan.summary_lines_vi() if self.prescan else []

    def summary_lines(self) -> List[str]:
        s = self.summary or BatchSummary(total=len(self.files))
        lines = [f"Tổng: {s.total}", f"Hoàn thành: {s.completed}", f"Cần kiểm tra: {s.needs_review}",
                 f"Không tìm thấy Management Number: {s.not_written}", f"Lỗi: {s.failed}", f"Bỏ qua: {s.skipped}"]
        if self.prescan:
            lines.extend(self.prescan_lines())
        if self.started_at is not None:
            lines.append(f"Tổng thời gian xử lý: {format_elapsed(self.elapsed_seconds())}")
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
            "Ảnh Sau cải tiến": (", ".join(fr.after_pictures) + f" (slide {fr.after_picture_slides})")
            if fr.after_pictures else "(không có)",
            "Quyết định từng ảnh": "\n".join(fr.picture_notes) if fr.picture_notes else "-",
            "Bộ phân loại": fr.classifier or "-",
            "Độ tin cậy AI": f"{fr.confidence:.2f}" if fr.confidence is not None else "-",
            "Trường đã điền": ", ".join(fr.filled_fields) if fr.filled_fields else "-",
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
        c.period_mode = self.period_mode
        try:
            c.period_month = int(self.period_month) if str(self.period_month).strip() else 0
            c.period_year = int(self.period_year) if str(self.period_year).strip() else 0
        except ValueError:
            c.period_month, c.period_year = 0, 0
        c.period_from, c.period_to = self.period_from, self.period_to
        try:
            c.save(self._config_path)
        except OSError as e:
            self.log_lines.append(f"Không lưu được config.json: {e}")
