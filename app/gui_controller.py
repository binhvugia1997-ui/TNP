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
from .config import (DEFAULT_MODEL, DEFAULT_OLLAMA, DEFAULT_PROBE_TIMEOUT, LOCAL_OLLAMA, AppConfig, endpoint_problem,
                     is_local_host, normalize_ollama_url, split_endpoint)
from .excel_writer import validate_template
from .extractor import management_number_from_filename
from .logger import FileResult
from .ollama_client import OllamaClient, OllamaError, preferred_model
from .ollama_discovery import (OllamaDiscovery, OllamaDiscoveryResult, discovery_summary_vi, missing_model_message,
                               model_available)
from .prescan import (ACTION_INVALID_MGMT, ACTION_MASTER_COMPLETE, ACTION_OUTSIDE_PERIOD, ACTION_PROCESS,
                      ACTION_PROCESS_NEW_ROW, ACTION_SOURCE_DUPLICATE, ACTION_FAST_SKIP, AUTO_PERIOD_FAIL_VI,
                      CACHE_FILE_NAME, PERIOD_DIFFERS_VI, ALL_PERIOD, FastScanCache, MasterLookup, PreScanItem,
                      PreScanResult, ProcessingPeriod, auto_period_from_excel, month_period, normalize_source_path,
                      prescan, range_period)
from .excel_writer import ExcelWriter
from .scanner import scan_inputs
from . import updater
from .runtime_paths import portable_root
from .updater import UpdateCheck, version_label

LOG = logging.getLogger("report_extractor.gui")
DEFAULT_OUTPUT_NAME = "Kiem_chung_Ket_qua.xlsx"

# final statuses shown in the result table (Vietnamese)
STATUS_VI: Dict[str, str] = {
    "waiting": "Đang chờ",
    "completed": "Hoàn thành",
    "needs_review": "Cần kiểm tra",
    "completed_new": "Hoàn thành — đã thêm Management Number mới",
    "not_written": "Cần kiểm tra — Không xác định được Management Number từ tên file",
    "error": "Lỗi",
    "skipped": "Bỏ qua — đã cập nhật",
    "outside_period": "Bỏ qua ngoài thời gian xử lý",
    "source_duplicate": "Trùng Management Number trong folder",
    "fast_skip": "Bỏ qua nhanh — đã xử lý gần đây",
}
PRESCAN_STATUSES = ("outside_period", "source_duplicate", "fast_skip")
FINAL_STATUSES = ("completed", "completed_new", "needs_review", "not_written", "error", "skipped") + PRESCAN_STATUSES
PERIOD_MODES = ("auto", "month", "range", "all")
# scanned-file list: Vietnamese label of each pre-scan decision + the manual exclusion
AUTO_UPDATE_CHECK_KEY = "auto_update_check"            # cfg.extra key (persisted; default enabled)
AUTO_CHECK_FAILED_VI = "Không thể kiểm tra cập nhật tự động — chương trình vẫn hoạt động bình thường."
UPDATE_RESULT_STATUSES = ("available", "latest", "older")   # real comparisons; everything else = not reachable


def _as_bool(v) -> bool:
    if isinstance(v, str):
        return v.strip().lower() not in ("0", "false", "no", "off", "")
    return bool(v)


USER_EXCLUDED = "USER_EXCLUDED"
SCAN_LABELS_VI: Dict[str, str] = {
    ACTION_PROCESS: "Sẽ xử lý — cần bổ sung dữ liệu",
    ACTION_PROCESS_NEW_ROW: "Sẽ xử lý — Management Number mới",
    ACTION_FAST_SKIP: "Bỏ qua — đã xử lý gần đây",
    ACTION_MASTER_COMPLETE: "Bỏ qua — Excel đã đầy đủ",
    ACTION_OUTSIDE_PERIOD: "Ngoài thời gian xử lý",
    ACTION_SOURCE_DUPLICATE: "Trùng Management Number trong folder",
    ACTION_INVALID_MGMT: "Không xác định được Management Number từ tên file",
    USER_EXCLUDED: "Đã loại thủ công",
}
SCAN_FILTERS_VI = ("File cần xử lý", "Tất cả file đã quét", "File bị bỏ qua", "File đã loại thủ công")
# display groups of the scanned list (ascending = top to bottom); manual exclusions are always last
GROUP_PROCESSING, GROUP_WAITING, GROUP_REVIEW, GROUP_COMPLETED, GROUP_SKIPPED, GROUP_EXCLUDED = range(6)
DEFAULT_SCAN_FILTER_VI = SCAN_FILTERS_VI[1]       # every discovered report is visible first; filtering is opt-in
ATTENTION_ACTIONS = (ACTION_INVALID_MGMT,)
STALE_LIST_VI = "Danh sách file đã thay đổi điều kiện. Vui lòng quét lại."
NO_SCAN_VI = "Chưa quét thư mục. Vui lòng bấm Quét lại."
RUNNING_VI = "Đang xử lý – không thể thay đổi danh sách. Dùng 'Dừng sau báo cáo hiện tại'."
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
    high_water: float = 0.0             # monotonic guard: a stale stage event never lowers the overall percentage

    @property
    def percent(self) -> float:
        if self.total <= 0:
            return 0.0
        value = (self.done + (self.current_fraction if self.current_index is not None else 0.0)) / self.total * 100.0
        value = max(0.0, min(100.0, value))
        if value > self.high_water:
            self.high_water = value
        return self.high_water

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
class ScanRow:
    """One line of the scanned-file list (view model built from the real PreScanItem)."""
    index: int
    key: str                          # stable identity: normalized path + Management Number
    path: Path
    management_number: str
    occurrence_date: str
    action: str                       # pre-scan action
    excluded: bool
    reason: str
    excel_row: Optional[int] = None
    run_stage: str = ""               # live/final pipeline stage of the current batch for this file ('' = none)
    vendor: str = ""                  # Vendor known from the current batch (display only)
    note: str = ""                    # batch note (diagnostics only – not a main-table column)

    @property
    def is_candidate(self) -> bool:
        return self.action in (ACTION_PROCESS, ACTION_PROCESS_NEW_ROW)

    @property
    def is_new_row(self) -> bool:
        return self.action == ACTION_PROCESS_NEW_ROW

    @property
    def needs_attention(self) -> bool:
        """Discovered report that cannot be processed as-is (key absent from the master / unreadable key):
        never hidden from the default list – the user must see it with its status."""
        return self.action in ATTENTION_ACTIONS

    @property
    def state(self) -> str:
        return USER_EXCLUDED if self.excluded else self.action

    @property
    def scan_status_vi(self) -> str:
        """Pure pre-scan label (ignores live batch status and manual exclusion)."""
        return SCAN_LABELS_VI.get(self.state, self.state)

    @property
    def status_vi(self) -> str:
        if self.excluded:
            return SCAN_LABELS_VI[USER_EXCLUDED]
        if self.run_stage and self.run_stage != "waiting":
            return status_label(self.run_stage)
        return SCAN_LABELS_VI.get(self.state, self.state)

    @property
    def will_process(self) -> bool:
        return self.is_candidate and not self.excluded

    @property
    def group(self) -> int:
        """Deterministic display group (GROUP_* ranks); manual exclusions are ALWAYS last."""
        if self.excluded:
            return GROUP_EXCLUDED
        st = self.run_stage
        if st and st != "waiting":
            if st in ("completed", "completed_new"):
                return GROUP_COMPLETED
            if st in ("needs_review", "not_written", "error"):
                return GROUP_REVIEW
            if st in ("skipped", "fast_skip", "outside_period", "source_duplicate"):
                return GROUP_SKIPPED
            return GROUP_PROCESSING                                  # live pipeline stage
        if self.is_candidate:
            return GROUP_WAITING
        if self.needs_attention:
            return GROUP_REVIEW
        return GROUP_SKIPPED

    def as_values(self, stt: int) -> Tuple[Any, ...]:
        """Compact main-table values (STT = current display order; path/note stay internal -> details())."""
        return (stt, self.management_number, self.occurrence_date, self.vendor.replace("\n", " / "), self.path.name,
                self.status_vi)

    def details(self) -> str:
        """Full diagnostics text (double-click): full path, status, pre-scan reason, batch note."""
        lines = [self.path.name, "", f"Management Number: {self.management_number or '(trống)'}",
                 f"Ngày phát sinh: {self.occurrence_date or '(trống)'}", f"Trạng thái: {self.status_vi}"]
        if self.run_stage and not self.excluded:
            lines.append(f"Trạng thái quét: {self.scan_status_vi}")
        if self.reason:
            lines.append(self.reason)
        if self.note:
            lines.append(f"Ghi chú: {self.note}")
        if self.excel_row:
            lines.append(f"Dòng Excel: {self.excel_row}")
        lines += ["", "Đường dẫn:", str(self.path)]
        return "\n".join(lines)


@dataclass
class ContentReviewOutcome:
    """Separate counters of the content-review workflow (labels vs Excel) – never report prepared as committed."""
    labels_saved: int = 0
    excel_rows_prepared: int = 0
    excel_rows_committed: int = 0
    locked_path: Optional[str] = None
    excel_error: str = ""

    @property
    def excel_pending(self) -> bool:
        return self.excel_rows_prepared > 0 and self.excel_rows_committed == 0 and bool(self.excel_error)


@dataclass
class UiEvent:
    kind: str                      # row | progress | log | done | ollama | models | discovery_*
    payload: Any = None


UPDATE_STAGES_VI = {                      # PROMPT-011 progress window texts (stage -> label)
    "PREPARING": "Đang chuẩn bị cập nhật...",
    "COPYING": "Đang sao chép bản cập nhật...",
    "VERIFYING": "Đang kiểm tra tính toàn vẹn...",
    "READY": "Gói cập nhật đã được kiểm tra và sẵn sàng.",
    "APPLYING": "Đang chuẩn bị cài đặt...",
    "BACKUP": "Đang sao lưu phiên bản hiện tại...",      # performed by the external updater (after this GUI exits)
    "INSTALLING": "Đang cài đặt...",                       # idem – never shown with an invented percentage here
    "HANDOFF": "Gói cập nhật đã sẵn sàng.\nĐang khởi động trình cài đặt...",
    "COMPLETE": "Cập nhật hoàn tất — đang khởi động lại...",
    "ERROR": "Cập nhật thất bại",
}


def format_bytes(n: float) -> str:
    n = float(n or 0)
    if n >= 1 << 30:
        return f"{n / (1 << 30):.2f} GB"
    if n >= 1 << 20:
        return f"{n / (1 << 20):.1f} MB"
    if n >= 1 << 10:
        return f"{n / (1 << 10):.0f} KB"
    return f"{int(n)} B"


@dataclass
class UpdateProgress:
    """Snapshot published by the install worker (immutable per event; rendered on the Tk thread)."""
    stage: str = ""
    bytes_copied: int = 0
    total_bytes: int = 0
    percent: float = 0.0
    speed_bps: float = 0.0
    error: str = ""
    determinate: bool = False

    @property
    def label(self) -> str:
        return UPDATE_STAGES_VI.get(self.stage, self.stage)

    @property
    def bytes_text(self) -> str:
        if self.stage in ("COPYING", "VERIFYING", "READY", "HANDOFF", "COMPLETE") and self.total_bytes:
            return f"{format_bytes(self.bytes_copied)} / {format_bytes(self.total_bytes)}"
        return ""

    @property
    def speed_text(self) -> str:
        return f"Tốc độ: {format_bytes(self.speed_bps)}/s" if self.stage == "COPYING" and self.speed_bps > 0 else ""


class GuiController:
    """State machine + adapter between the GUI widgets and the production pipeline."""

    def __init__(self, cfg: Optional[AppConfig] = None,
                 processor_factory: Callable[..., BatchProcessor] = BatchProcessor,
                 client_factory: Callable[..., OllamaClient] = OllamaClient,
                 config_path: Optional[Path] = None, learning_dir_override: Optional[Path] = None):
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
        # PROMPT-005 offline update (optional infrastructure – never blocks processing)
        self.update_path: str = (self.cfg.update_path or "").strip()
        self.update_check: Optional[UpdateCheck] = None
        self.update_busy: bool = False
        self.update_dirty: bool = False          # GUI refreshes the update card when set
        self.update_notice: str = ""             # one-shot "Cập nhật thành công lên phiên bản X."
        # PROMPT-009: automatic LAN check on start-up (persisted in cfg.extra – no schema change), soft failures,
        # one notification per remote build per session
        self.auto_update_check: bool = _as_bool((self.cfg.extra or {}).get(AUTO_UPDATE_CHECK_KEY, True))
        self.update_check_startup: bool = False  # last check was the automatic start-up one (soft failure text)
        self.offered_update_builds: set = set()  # remote builds already announced in this session
        self._update_thread: Optional[threading.Thread] = None
        self.update_check_done = threading.Event()   # set when a check has COMPLETED (result published); cleared at start
        self.update_check_done.set()
        self.update_installing: bool = False         # PROMPT-011: one install at a time (buttons disabled meanwhile)
        self.update_progress = UpdateProgress()
        self._install_thread: Optional[threading.Thread] = None
        self._install_lock = threading.Lock()
        self.prescan: Optional[PreScanResult] = None
        self.queue_indexes: Optional[set] = None       # indexes that really enter the processing pipeline
        # reviewed scan list (before Start): real pre-scan result + manual exclusions
        self.all_files: List[Path] = []                # every discovered PPT/PPTX (recursive)
        self.scan_result: Optional[PreScanResult] = None
        self.scan_stale: bool = False
        self.scan_message: str = ""
        self.excluded_keys: set = set()
        # runtime state
        self.state: str = "idle"                 # idle | running | stopping
        self.files: List[Path] = []
        self.rows: List[RowState] = []
        self.progress = Progress()
        self.summary: Optional[BatchSummary] = None
        self.processor: Optional[BatchProcessor] = None
        self.ollama_status: str = ""
        self.ollama_ok: Optional[bool] = None
        self.ollama_source: str = ""                 # "" | local | saved | none  (how the endpoint was chosen)
        self.autoconnect_runs = 0
        self.autoconnect_stale = 0                   # autoconnect results dropped because the user chose a server
        self.endpoint_epoch = 0                      # bumped by every explicit endpoint choice (typed / discovery)
        self.log_lines: List[str] = []
        self.started_at: Optional[float] = None      # time.monotonic() when the batch really started
        self.finished_at: Optional[float] = None
        self.report_durations: List[float] = []      # measured seconds per finished report (rolling window)
        self._report_started_at: Optional[float] = None
        self._clock: Callable[[], float] = time.monotonic
        self.done_events = 0                         # how many "done" events arrived (first one finalizes)
        self.stale_events = 0                        # worker events ignored after finalization
        self.worker_failure: str = ""                # set by reconcile() when the worker died without "done"
        self._today: Callable[[], Any] = __import__("datetime").date.today   # overridable for tests
        self._queue: "queue.Queue[UiEvent]" = queue.Queue()
        # LAN discovery (user-initiated only; completely separate from the batch progress)
        self._discovery_factory: Callable[..., OllamaDiscovery] = OllamaDiscovery
        self.discovery: Optional[OllamaDiscovery] = None
        self.discovery_state: str = "idle"           # idle | running
        self.discovery_checked = 0
        self.discovery_total = 0
        self.discovery_results: List[OllamaDiscoveryResult] = []
        self.discovery_message: str = ""
        self.discovery_runs = 0
        # PROMPT-006 image learning (user data in learning_data/; independent from Ollama)
        self.learning_dir: Optional[Path] = learning_dir_override
        self._learning = None
        self.review_index = 0
        self.pending_labels: Dict[str, str] = {}      # candidate_id -> label chosen in the review window (unsaved)
        self.pending_image_reapply: list = []         # saved image labels whose region-crop Excel reapply is still owed
        self.pending_content_labels: Dict[str, str] = {}   # PROMPT-006B content review (separate dataset)
        self.pending_content_reapply: list = []              # candidates whose Excel re-apply is still owed (locked)
        self.last_content_review = ContentReviewOutcome()
        self.review_content_index = 0

    # ------------------------------------------------------------------ PROMPT-006 image learning
    @property
    def learning(self):
        if self._learning is None:
            from .image_learning import ImageLearning
            try:
                self._learning = ImageLearning(self.learning_dir)
            except Exception as e:  # noqa: BLE001
                LOG.warning("image learning unavailable: %s", e)
                return None
        return self._learning

    def learning_status_text(self) -> str:
        lrn = self.learning
        if lrn is None:
            from .image_learning import MSG_MODEL_UNAVAILABLE
            return MSG_MODEL_UNAVAILABLE
        return lrn.status_text()

    def content_status_text(self) -> str:
        lrn = self.learning
        if lrn is None:
            from .content_learning import MSG_MODEL_UNAVAILABLE
            return MSG_MODEL_UNAVAILABLE
        return lrn.content.status_text()

    def learning_model_status_text(self) -> str:
        """Compact status for the separate learning tab; image/content models stay independent."""
        lrn = self.learning
        if lrn is None:
            return "Mô hình ảnh và nội dung: không khả dụng."

        def state_text(model, state: str) -> str:
            if model is not None:
                return f"đã huấn luyện ({model.n_examples} mẫu)"
            if state in ("corrupt", "incompatible"):
                return "không khả dụng — dùng quy tắc hiện tại"
            return "chưa huấn luyện — dùng quy tắc hiện tại"

        content = lrn.content
        return (f"Mô hình ảnh: {state_text(lrn.model, lrn.model_status)}.  "
                f"Mô hình nội dung: {state_text(content.model, content.model_status)}.")

    # ---- PROMPT-006B content review -------------------------------------------------------------
    def review_content_candidates(self) -> list:
        """Reviewable text blocks of the LAST finished batch (hard-excluded title/sidebar/footer never offered)."""
        if self.is_running() or not self.processor:
            return []
        out = []
        for fr in getattr(self.processor, "results", []):
            out.extend(c for c in getattr(fr, "content_candidates", []) if not c.hard_excluded)
        return out

    def content_review_summary_text(self) -> str:
        cands = self.review_content_candidates()
        if not cands:
            return "Chưa có khối nội dung để kiểm tra (chạy xử lý trước)."
        n_rev = sum(1 for c in cands if c.decision == "review")
        return f"{len(cands)} khối chữ ứng viên, {n_rev} khối chưa chắc chắn"

    def set_pending_content_label(self, cand, label: str) -> None:
        from .content_learning import CONTENT_LABELS
        if label not in CONTENT_LABELS:
            raise ValueError(label)
        self.pending_content_labels[cand.candidate_id] = label

    def save_content_confirmations(self, cands) -> Tuple[bool, str]:
        """Two independent transactions: (1) labels -> content_labels.jsonl (committed first, never rolled back);
        (2) Excel re-apply, which may fail (e.g. the master is open in Excel) without touching (1).  A failed
        Excel step keeps the touched candidates in ``pending_content_reapply`` so *Thử lại* only redoes Excel."""
        lrn = self.learning
        if lrn is None:
            return False, "Không ghi được dữ liệu học (thư mục learning_data không khả dụng)."
        if not self.pending_content_labels:
            return False, "Chưa chọn nhãn nào."
        by_id = {c.candidate_id: c for c in cands}
        written, touched = 0, []
        for cid, label in list(self.pending_content_labels.items()):
            c = by_id.get(cid)
            if c is None:
                continue
            if lrn.content.store.label(c, label) is not None:
                written += 1
            touched.append(c)
        self.pending_content_labels.clear()
        LOG.info("CONTENT_REVIEW_LABELS_SAVED count=%d", written)
        self.last_content_review = ContentReviewOutcome(labels_saved=written)
        msg = f"Đã lưu {written} nhãn nội dung."
        # merge rows still waiting from an earlier locked attempt (labels already on disk)
        for c in self.pending_content_reapply:
            if c.candidate_id not in {t.candidate_id for t in touched}:
                touched.append(c)
        self.pending_content_reapply = []
        if touched and self.template and self.output and Path(self.output).exists():
            msg = self._content_reapply(touched, msg)
        return True, msg

    def _content_reapply(self, touched, msg: str) -> str:
        from .image_review import reapply_content_labels
        res = reapply_content_labels(Path(self.template), Path(self.output), touched, self.learning)
        out = self.last_content_review
        out.excel_rows_prepared = len(res.prepared_rows)
        out.excel_rows_committed = len(res.updated_rows)
        out.locked_path = res.locked_path
        out.excel_error = "; ".join(res.errors)
        self.log_lines.extend(res.messages + res.errors)
        if res.technical_error:
            self.log_lines.append(f"Chi tiết kỹ thuật: {res.technical_error}")
        if res.updated_rows:
            msg += f" Đã cập nhật nội dung cải tiến cho {len(res.updated_rows)} dòng Excel."
        if res.locked:
            self.pending_content_reapply = list(touched)          # retry = Excel only, labels stay saved
            n = len(res.prepared_rows)
            msg += (f"\n\nChưa thể cập nhật {n} dòng Excel vì file đang được sử dụng:\n\n{res.locked_path}"
                    "\n\nHãy đóng file rồi thử lại.")
        elif res.errors:
            self.pending_content_reapply = list(touched)
            msg += "\n\nChưa cập nhật được Excel: " + "; ".join(res.errors)
        return msg

    def retry_content_reapply(self) -> Tuple[bool, str]:
        """*Thử lại*: Excel transaction only – uses the labels already saved, no review, no re-labelling."""
        touched = list(self.pending_content_reapply)
        if not touched:
            return False, "Không có dòng Excel nào đang chờ cập nhật."
        if self.learning is None or not (self.template and self.output and Path(self.output).exists()):
            return False, "Chưa có file Excel kết quả để cập nhật."
        self.pending_content_reapply = []
        saved = self.last_content_review.labels_saved
        self.last_content_review = ContentReviewOutcome(labels_saved=saved)
        msg = self._content_reapply(touched, f"Nhãn nội dung đã lưu trước đó ({saved} nhãn) được dùng lại.")
        return self.last_content_review.excel_rows_committed > 0 or not self.last_content_review.excel_error, msg

    @property
    def content_reapply_pending(self) -> int:
        return len(self.pending_content_reapply)

    def train_models(self) -> Tuple[bool, str]:
        """'Cập nhật mô hình học': image and content models independently, two result lines."""
        lrn = self.learning
        if lrn is None:
            from .image_learning import MSG_MODEL_UNAVAILABLE
            return False, MSG_MODEL_UNAVAILABLE
        ok, msg = lrn.train_all()
        self.log_lines.extend(msg.splitlines())
        return ok, msg

    def learning_counts(self) -> Dict[str, int]:
        lrn = self.learning
        return lrn.store.counts() if lrn else {"total": 0, "after": 0, "non_after": 0}

    def learning_folder(self) -> Path:
        lrn = self.learning
        from .image_learning import learning_dir as _ld
        return lrn.dir if lrn else _ld(self.learning_dir)

    def review_candidates(self) -> list:
        """Reviewable pictures of the LAST finished batch (hard-excluded logos/arrows are never offered)."""
        if self.is_running() or not self.processor:
            return []
        out = []
        for fr in getattr(self.processor, "results", []):
            out.extend(c for c in getattr(fr, "image_candidates", []) if not c.hard_excluded)
        return out

    def review_summary_text(self) -> str:
        cands = self.review_candidates()
        if not cands:
            return "Chưa có ảnh để kiểm tra (chạy xử lý trước)."
        n_rev = sum(1 for c in cands if c.decision == "review")
        return f"{len(cands)} ảnh ứng viên, {n_rev} ảnh cần xác nhận"

    @staticmethod
    def candidate_blob(c) -> Optional[bytes]:
        """Raw picture bytes of a candidate (re-read from the PPTX; preview only)."""
        try:
            from .pptx_parser import parse_pptx
            sl = parse_pptx(c.source_file).slide(c.slide)
            for p in (sl.pictures if sl else []):
                if p.shape_id == c.picture_id:
                    return p.image_blob
        except Exception as e:  # noqa: BLE001
            LOG.debug("candidate blob unavailable: %s", e)
        return None

    def set_pending_label(self, cand, label: str) -> None:
        from .image_learning import LABELS
        if label not in LABELS:
            raise ValueError(label)
        self.pending_labels[cand.candidate_id] = label

    def save_confirmations(self, cands) -> Tuple[bool, str]:
        """Persist labels first, then re-apply region crops to the Excel image cells."""
        lrn = self.learning
        if lrn is None:
            return False, "Không ghi được dữ liệu học (thư mục learning_data không khả dụng)."
        if not self.pending_labels:
            return False, "Chưa chọn nhãn nào."
        by_id = {c.candidate_id: c for c in cands}
        written, touched = 0, []
        for cid, label in list(self.pending_labels.items()):
            c = by_id.get(cid)
            if c is None:
                continue
            if lrn.store.label(c, label) is not None:
                written += 1
            touched.append(c)
        self.pending_labels.clear()
        # Include reports owed from a previous locked attempt; the label JSONL stays the source of truth.
        by_touched_id = {c.candidate_id: c for c in self.pending_image_reapply}
        by_touched_id.update({c.candidate_id: c for c in touched})
        touched = list(by_touched_id.values())
        self.pending_image_reapply = []
        msg = f"Đã lưu {written} nhãn xác nhận."
        if touched:
            if self.template and self.output and Path(self.output).exists():
                msg = self._image_reapply(touched, lrn, msg)
            else:
                self.pending_image_reapply = list(touched)
                msg += "\n\nNhãn đã lưu; chưa có file Excel kết quả để tạo lại ảnh, có thể thử lại sau."
        return True, msg

    def _image_reapply(self, touched, learning, msg: str) -> str:
        from .image_review import reapply_labels
        try:
            res = reapply_labels(Path(self.template), Path(self.output), touched, learning)
        except Exception as exc:  # noqa: BLE001 – saved labels remain safe; the Excel step is retryable
            self.pending_image_reapply = list(touched)
            self.log_lines.append(f"IMAGE_REAPPLY_FAILED: {exc}")
            return msg + f"\n\nChưa cập nhật được Excel: {exc} (có thể thử lại từ nhãn đã lưu)."
        self.log_lines.extend(res.messages + res.errors)
        if res.updated_rows:
            msg += f" Đã cập nhật ảnh cải tiến cho {len(res.updated_rows)} dòng Excel ({res.pictures} ảnh)."
        if res.locked:
            self.pending_image_reapply = list(touched)     # labels are already saved; retry re-renders region crops
            msg += (f"\n\nChưa thể cập nhật {len(res.prepared_rows)} dòng Excel vì file đang được sử dụng:"
                    f"\n\n{res.locked_path}\n\nĐóng file rồi chọn 'Thử lại cập nhật ảnh đã lưu'.")
        elif res.errors:
            self.pending_image_reapply = list(touched)
            msg += "\n\nChưa cập nhật được Excel: " + "; ".join(res.errors)
        return msg

    def retry_image_reapply(self) -> Tuple[bool, str]:
        """Retry a saved-label Excel transaction; re-parses PPTX and regenerates visual-region evidence."""
        touched = list(self.pending_image_reapply)
        if not touched:
            return False, "Không có ảnh Excel nào đang chờ cập nhật."
        if self.learning is None or not (self.template and self.output and Path(self.output).exists()):
            return False, "Chưa có file Excel kết quả để cập nhật."
        self.pending_image_reapply = []
        msg = self._image_reapply(touched, self.learning, "Nhãn ảnh đã lưu được dùng lại.")
        return not bool(self.pending_image_reapply), msg

    @property
    def image_reapply_pending(self) -> int:
        return len(self.pending_image_reapply)

    def train_image_model(self) -> Tuple[bool, str]:
        lrn = self.learning
        if lrn is None:
            from .image_learning import MSG_MODEL_UNAVAILABLE
            return False, MSG_MODEL_UNAVAILABLE
        ok, msg = lrn.train()
        self.log_lines.append(msg)
        return ok, msg

    def export_learning_data(self, target: str) -> Tuple[bool, str]:
        lrn = self.learning
        if lrn is None:
            return False, "Không có dữ liệu học."
        try:
            p = lrn.store.export(Path(target))
        except OSError as e:
            return False, f"Không xuất được dữ liệu học: {e}"
        return True, f"Đã xuất {lrn.store.counts()['total']} mẫu ra {p}"

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
        if (host, int(port_s or 11434)) != (self.host, self.port):
            self.endpoint_epoch += 1                 # explicit user choice outranks any in-flight autoconnect
        self.host, self.port = host, int(port_s or 11434)
        if model is not None:
            self.model = model.strip()
        self.ollama_ok, self.ollama_status = None, ""        # endpoint changed -> status unknown
        self.ollama_source = ""
        return ""

    # ------------------------------------------------------------------ inputs
    def set_report_folder(self, folder: str) -> int:
        """Select the report folder, discover PPT/PPTX recursively, propose the output path."""
        changed = folder.strip() != self.report_folder
        self.report_folder = folder.strip()
        if changed:
            self.excluded_keys = set()
        if not self.output or Path(self.output).name == DEFAULT_OUTPUT_NAME:
            self.output = default_output_path(self.report_folder)
        n = self.discover()
        if changed:
            self._invalidate_scan()
        return n

    def discover(self) -> int:
        """Cheap recursive discovery (names only – nothing is opened)."""
        folder = Path(self.report_folder) if self.report_folder else None
        self.scan_rejections: List[Tuple[Path, str]] = []
        self.all_files = (scan_inputs([folder], on_reject=lambda p, why: self.scan_rejections.append((p, why)))
                          if folder and folder.is_dir() else [])
        for p, why in self.scan_rejections:                  # every rejected PowerPoint-like entry is reported
            self.log_lines.append(f"Bỏ qua khi quét thư mục: {p.name} – {why}")
        self._set_files(self.all_files)
        return len(self.all_files)

    def _set_files(self, files: List[Path]) -> None:
        self.files = list(files)
        self.rows = [RowState(index=i, path=p, management_number=management_number_from_filename(p.name))
                     for i, p in enumerate(self.files)]
        self.progress = Progress(total=len(self.files))
        self.summary = None

    def _invalidate_scan(self) -> None:
        """Any input that changes the candidate set makes the reviewed list stale (never run the old queue)."""
        if self.scan_result is not None:
            self.scan_stale = True
            self.scan_message = STALE_LIST_VI

    def set_template(self, path: str) -> None:
        if path.strip() != self.template:
            self.template = path.strip()
            self._invalidate_scan()

    def set_output(self, path: str) -> None:
        if path.strip() != self.output:
            self.output = path.strip()
            self._invalidate_scan()

    def set_force(self, value: bool) -> None:
        if bool(value) != self.force_reprocess:
            self.force_reprocess = bool(value)
            self._invalidate_scan()

    # ------------------------------------------------------------------ scanned-file list (review before Start)
    @staticmethod
    def candidate_key(path: Path, mgmt: str) -> str:
        return f"{normalize_source_path(Path(path))}|{(mgmt or '').strip().upper()}"

    def scan(self) -> Optional[PreScanResult]:
        """Run the REAL pre-scan (same function the batch uses) on the discovered files with the current folder /
        Excel / period / cache / force settings.  Resets manual exclusions.  Nothing is opened or written.
        Never runs while a batch is active: ``discover()`` would replace the live rows/progress from another thread."""
        if self.is_running():
            return self.scan_result
        self.discover()
        per, err = self.effective_period()
        # manual exclusions survive an ordinary rescan of the same folder (keyed by normalised path + Management
        # Number, so a renamed/moved file simply becomes a new candidate); a folder change clears them
        self.scan_stale, self.scan_message = False, ""
        if err:
            self.scan_result, self.scan_message = None, err
            return None
        master = None
        writer = None
        try:
            tpl, out = Path(self.template) if self.template else None, Path(self.output) if self.output else None
            if tpl and tpl.is_file() and out and (self.cfg.row_mode or "match") == "match":
                writer = ExcelWriter(tpl, out, probe=True)
                master = MasterLookup.from_writer(writer)
        except Exception as e:  # noqa: BLE001
            self.log_lines.append(f"Không đọc được file Excel để quét: {e}")
        cache = None
        try:
            if self.output:
                cache = FastScanCache(Path(self.output).parent / "logs" / CACHE_FILE_NAME)
        except Exception:  # noqa: BLE001
            cache = None
        try:
            self.scan_result = prescan(self.all_files, per, cache, master, force=self.force_reprocess,
                                       today=self._today(), on_stage=lambda m: self.log_lines.append(m))
        finally:
            if writer is not None:
                writer.close()
        for line in self.scan_result.summary_lines_vi():
            self.log_lines.append(line)
        return self.scan_result

    def scan_async(self) -> None:
        def work():
            try:
                res = self.scan()
                self._queue.put(UiEvent("scan", res))
            except Exception as e:  # noqa: BLE001
                self.scan_result, self.scan_message = None, f"Quét thư mục thất bại: {e}"
                self._queue.put(UiEvent("scan", None))
        self.log_lines.append("Đang quét thư mục...")
        threading.Thread(target=work, name="prescan", daemon=True).start()

    def scan_rows(self, filter_name: str = SCAN_FILTERS_VI[1]) -> List[ScanRow]:
        """View rows of the scanned list.  Filter: 'File cần xử lý' (candidates incl. manually excluded ones),
        'Tất cả file đã quét', 'File bị bỏ qua' (business skips + manual exclusions)."""
        if not self.scan_result:
            return []
        rows: List[ScanRow] = []
        live = {normalize_source_path(Path(r.path)): r for r in self.rows} if self.rows else {}
        for it in self.scan_result.items:
            key = self.candidate_key(it.path, it.management_number)
            rs = live.get(normalize_source_path(Path(it.path)))
            rows.append(ScanRow(index=it.index, key=key, path=it.path, management_number=it.management_number,
                                occurrence_date=f"{it.occurrence_date:%d/%m/%Y}" if it.occurrence_date else "",
                                action=it.action, excluded=key in self.excluded_keys, reason=it.reason,
                                excel_row=it.excel_row, run_stage=rs.stage if rs else "",
                                vendor=rs.vendor if rs else "", note=rs.note if rs else ""))
        if filter_name == SCAN_FILTERS_VI[0]:
            rows = [r for r in rows if (r.is_candidate or r.needs_attention) and not r.excluded]   # problems never hidden
        elif filter_name == SCAN_FILTERS_VI[2]:
            rows = [r for r in rows if not r.will_process]
        elif filter_name == SCAN_FILTERS_VI[3]:
            rows = [r for r in rows if r.excluded]
        # deterministic grouping: stable sort keeps the original scan order inside every group
        rows.sort(key=lambda r: r.group)
        return rows

    def _scan_row(self, index: int) -> Optional[ScanRow]:
        return next((r for r in self.scan_rows() if r.index == index), None)

    def exclude(self, indexes) -> str:
        """Remove processing candidates from the CURRENT queue (never touches the files).  '' or a problem."""
        if self.is_running():
            return RUNNING_VI
        if not self.scan_result:
            return NO_SCAN_VI
        n = 0
        for i in list(indexes):
            r = self._scan_row(int(i))
            if r and r.is_candidate and not r.excluded:
                self.excluded_keys.add(r.key)
                n += 1
                LOG.info("USER_EXCLUDED management_number=%s file=%s", r.management_number, r.path)
        return "" if n else "Không có file cần xử lý nào được chọn."

    def restore(self, indexes) -> str:
        """Undo a manual exclusion; business exclusions (period, duplicate, invalid key, complete row, cache) stay."""
        if self.is_running():
            return RUNNING_VI
        if not self.scan_result:
            return NO_SCAN_VI
        n = 0
        for i in list(indexes):
            r = self._scan_row(int(i))
            if r and r.excluded:
                self.excluded_keys.discard(r.key)
                n += 1
                LOG.info("USER_RESTORED management_number=%s file=%s", r.management_number, r.path)
        return "" if n else "Chỉ khôi phục được file đã loại thủ công."

    def final_queue(self) -> List[PreScanItem]:
        """Candidates of the reviewed list minus manual exclusions (source order) = the real processing queue."""
        if not self.scan_result:
            return []
        return [it for it in self.scan_result.candidates
                if self.candidate_key(it.path, it.management_number) not in self.excluded_keys]

    def queue_counts(self) -> Dict[str, int]:
        cands = self.scan_result.candidates if self.scan_result else []
        excluded = sum(1 for it in cands if self.candidate_key(it.path, it.management_number) in self.excluded_keys)
        pc = self.scan_result.counts() if self.scan_result else {}
        return {"candidates": len(cands), "excluded": excluded, "will_process": len(cands) - excluded,
                "discovered": len(self.all_files),
                "skipped": pc.get("master_complete", 0) + pc.get("fast_skipped", 0) + pc.get("outside_period", 0)
                + pc.get("source_duplicates", 0),
                "new_rows": pc.get("new_rows", 0),
                "invalid": pc.get("invalid_management_number", 0)}

    def queue_text(self) -> str:
        """Discovery vs eligibility at a glance: every discovered file is accounted for in exactly one bucket
        (+ manual exclusions), so a report that is not going to be processed is visible, never silently gone."""
        c = self.queue_counts()
        return (f"Tổng file phát hiện: {c['discovered']}   Sẽ xử lý: {c['will_process']}   "
                f"Bỏ qua/đã cập nhật: {c['skipped']}   Management Number mới: {c['new_rows']}   "
                f"Lỗi/không hợp lệ: {c['invalid']}   Đã loại thủ công: {c['excluded']}")

    def set_ollama(self, server: str, model: str) -> None:
        self.server = server
        self.model = model.strip()

    # ------------------------------------------------------------------ processing period
    def set_period(self, mode: str, month: Any = None, year: Any = None, start: Optional[str] = None,
                   end: Optional[str] = None) -> str:
        """Change the period mode/values; returns a Vietnamese problem text ('' when fine)."""
        if mode not in PERIOD_MODES:
            return f"Chế độ thời gian không hợp lệ: {mode}"
        before = (self.period_mode, self.period_month, self.period_year, self.period_from, self.period_to)
        self.period_mode = mode
        if month is not None:
            self.period_month = str(month).strip()
        if year is not None:
            self.period_year = str(year).strip()
        if start is not None:
            self.period_from = start.strip()
        if end is not None:
            self.period_to = end.strip()
        if (self.period_mode, self.period_month, self.period_year, self.period_from, self.period_to) != before:
            self._invalidate_scan()
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
        elif not self.all_files and not self.files:
            errs.append("Không tìm thấy file .ppt/.pptx nào trong thư mục báo cáo.")
        if self.scan_stale:
            errs.append(STALE_LIST_VI)
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

    def _make_client(self, server: str, timeout: Optional[int] = None):
        """Production client: inference timeout = request_timeout (unchanged), probe timeout = short.
        An explicit ``timeout`` (manual check) caps the probe only."""
        probe = min(self.probe_timeout, float(timeout)) if timeout else self.probe_timeout
        try:
            return self._client_factory(server, timeout=int(self.cfg.request_timeout or 180), probe_timeout=probe)
        except TypeError:                                       # factories that only accept (server, timeout)
            return self._client_factory(server, timeout=int(timeout or self.cfg.request_timeout or 180))

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
            client = self._make_client(server, timeout)
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

    # ------------------------------------------------------------------ Ollama local-first (PROMPT-004A)
    @property
    def probe_timeout(self) -> float:
        """Connection / model-list probe timeout – separate from the Qwen inference timeout (request_timeout)."""
        try:
            t = float(getattr(self.cfg, "probe_timeout", DEFAULT_PROBE_TIMEOUT) or DEFAULT_PROBE_TIMEOUT)
        except (TypeError, ValueError):
            t = float(DEFAULT_PROBE_TIMEOUT)
        return max(0.5, min(t, 15.0))

    def _probe(self, server: str) -> Tuple[Optional[List[str]], str]:
        """Real Ollama API (/api/tags) with the SHORT probe timeout. (models | None, error kind)."""
        try:
            info = self._client_factory(server, timeout=int(self.cfg.request_timeout or 180),
                                        probe_timeout=self.probe_timeout).test_connection()
        except TypeError:                                   # client factory without probe_timeout (tests/legacy)
            try:
                info = self._client_factory(server, timeout=int(self.probe_timeout) or 1).test_connection()
            except Exception as e:  # noqa: BLE001
                return None, self._classify_ollama_error(e)
        except Exception as e:  # noqa: BLE001
            return None, self._classify_ollama_error(e)
        return list((info or {}).get("models", []) or []), ""

    def _autoconnect_stale(self, epoch: int) -> bool:
        """True when the user picked a server (typed / confirmed discovery) while this autoconnect was probing:
        the stale result must not overwrite the newer explicit selection."""
        if epoch != self.endpoint_epoch:
            self.autoconnect_stale += 1
            LOG.info("OLLAMA_AUTOCONNECT ignored=stale server=%s (user selection took precedence)", self.endpoint_label)
            return True
        return False

    def _select_local(self, models: List[str]) -> Tuple[bool, str]:
        """Local Ollama answered: show 127.0.0.1 / 11434 in the controls; never switch to an unrelated model."""
        self.host, self.port = split_endpoint(LOCAL_OLLAMA)
        self.ollama_source = "local"
        if models:
            self.available_models = models
        model = self.model.strip()
        if not model and DEFAULT_MODEL in models:
            self.model = model = DEFAULT_MODEL
        if model and model in models:
            self.ollama_ok, self.ollama_status = True, f"Ollama local: Sẵn sàng — {model}"
        else:
            self.ollama_ok = False
            self.ollama_status = f"Ollama local đang chạy nhưng không tìm thấy model {model or DEFAULT_MODEL}"
        LOG.info("OLLAMA_AUTOCONNECT source=local ok=%s model=%s models=%d", self.ollama_ok, model, len(models))
        return bool(self.ollama_ok), self.ollama_status

    def auto_connect(self) -> Tuple[bool, str]:
        """Connection priority: 1) local 127.0.0.1:11434  2) saved/configured server  3) nothing automatic
        (manual Server/Port, "Kiểm tra kết nối", "Tìm Ollama trong mạng LAN" stay available).  Never starts a LAN
        scan; every probe uses the short probe timeout so GUI start-up never waits for the inference timeout."""
        self.autoconnect_runs += 1
        epoch = self.endpoint_epoch                  # explicit selections made after this point take precedence
        models, _kind = self._probe(LOCAL_OLLAMA)
        if self._autoconnect_stale(epoch):
            return bool(self.ollama_ok), self.ollama_status
        if models is not None:
            return self._select_local(models)
        saved = self.server
        saved_label = self.endpoint_label
        if saved and not is_local_host(self.host) and not endpoint_problem(self.host, self.port):
            models, _kind = self._probe(saved)
            if self._autoconnect_stale(epoch):
                return bool(self.ollama_ok), self.ollama_status
            if models is not None:
                self.ollama_source = "saved"
                if models:
                    self.available_models = models
                model = self.model.strip()
                if model and model in models:
                    self.ollama_ok, self.ollama_status = True, f"● Đã kết nối — {model} @ {saved_label}"
                else:
                    self.ollama_ok = False
                    self.ollama_status = f"● Đã kết nối Ollama tại {saved_label} nhưng không tìm thấy model {model or DEFAULT_MODEL}"
                LOG.info("OLLAMA_AUTOCONNECT source=saved server=%s ok=%s", saved_label, self.ollama_ok)
                return bool(self.ollama_ok), self.ollama_status
        self.ollama_source = "none"
        self.ollama_ok, self.ollama_status = False, "Không kết nối được Ollama local hoặc server đã lưu."
        LOG.info("OLLAMA_AUTOCONNECT source=none saved=%s", saved_label)
        return False, self.ollama_status

    def auto_connect_async(self) -> None:
        def work():
            epoch = self.endpoint_epoch
            try:
                ok, msg = self.auto_connect()
            except Exception as e:  # noqa: BLE001
                LOG.exception("auto_connect crashed")
                ok, msg = False, f"Không kết nối được Ollama local hoặc server đã lưu. ({type(e).__name__})"
                if epoch == self.endpoint_epoch:
                    self.ollama_ok, self.ollama_status, self.ollama_source = False, msg, "none"
            if epoch != self.endpoint_epoch:
                return                               # stale: the GUI already shows the user's newer selection
            self._queue.put(UiEvent("autoconnect", (ok, msg)))
        threading.Thread(target=work, daemon=True, name="ollama-autoconnect").start()

    def check_ollama_local_first(self, timeout: Optional[int] = None) -> Tuple[bool, str]:
        """"Kiểm tra kết nối": check the endpoint in the controls; if it is a remote address that does not answer,
        a running local Ollama is detected and selected instead (a stale LAN IP must never hide local Ollama)."""
        ok, msg = self.check_ollama(timeout=timeout)
        if ok or is_local_host(self.host):
            return ok, msg
        models, _kind = self._probe(LOCAL_OLLAMA)
        if models is None:
            return ok, msg
        return self._select_local(models)

    def check_ollama_async(self, local_first: bool = False) -> None:
        """Run the check on a worker thread; the result arrives as an 'ollama' event in pump()."""
        def work():
            ok, msg = self.check_ollama_local_first() if local_first else self.check_ollama()
            self._queue.put(UiEvent("ollama", (ok, msg)))
        threading.Thread(target=work, daemon=True, name="ollama-check").start()

    def refresh_models(self, timeout: Optional[int] = None) -> Tuple[bool, str, List[str]]:
        """Query the configured server for its INSTALLED models (no hard-coded list).

        On failure the user's current model is kept untouched. Returns (ok, message, models)."""
        server = self.server
        if not server or endpoint_problem(self.host, self.port):
            return False, f"● {endpoint_problem(self.host, self.port) or 'Chưa nhập IP / Server của Ollama'}", []
        try:
            info = self._make_client(server, timeout).test_connection()
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

    # ------------------------------------------------------------------ LAN discovery (PROMPT-003)
    @property
    def discovery_running(self) -> bool:
        return self.discovery_state == "running"

    def can_discover(self) -> bool:
        """Only while no batch is running / stopping and no scan is already in flight."""
        return self.state == "idle" and not self.discovery_running

    def discovery_progress_text(self) -> str:
        if self.discovery_running:
            return (f"Đang tìm Ollama trong mạng LAN... Đã kiểm tra {self.discovery_checked} / "
                    f"{self.discovery_total} địa chỉ")
        return self.discovery_message

    def discover_ollama(self) -> List[OllamaDiscoveryResult]:
        """Synchronous scan (worker thread body). Never touches the configured server."""
        disc = self._discovery_factory()
        self.discovery = disc
        self.discovery_results = []
        self.discovery_runs += 1
        LOG.info("DISCOVERY_START port=%s", disc.port)

        def progress(checked: int, total: int) -> None:
            self._queue.put(UiEvent("discovery_progress", (checked, total)))

        def found(res: OllamaDiscoveryResult) -> None:
            self._queue.put(UiEvent("discovery_found", res))

        try:
            results = disc.run(progress=progress, found=found)
        except Exception as e:  # noqa: BLE001 – the scan must never crash the GUI
            LOG.exception("DISCOVERY_ERROR %s", e)
            results = list(disc.results)
        LOG.info("DISCOVERY_DONE checked=%s total=%s found=%s cancelled=%s networks=%s", disc.checked, disc.total,
                 len(results), disc.cancelled, [str(n) for n in disc.networks])
        self._queue.put(UiEvent("discovery_done", (results, disc.checked, disc.total, disc.cancelled)))
        return results

    def discover_ollama_async(self) -> bool:
        if not self.can_discover():
            return False
        self.discovery_state = "running"
        self.discovery_checked, self.discovery_total = 0, 0
        self.discovery_message = ""
        threading.Thread(target=self.discover_ollama, daemon=True, name="ollama-discovery").start()
        return True

    def cancel_discovery(self) -> bool:
        if not self.discovery_running or self.discovery is None:
            return False
        self.discovery.cancel()
        return True

    def apply_discovered_server(self, result: OllamaDiscoveryResult, timeout: Optional[int] = None) -> Tuple[bool, str]:
        """User confirmed "Sử dụng server này": set + persist the endpoint through the normal config path, refresh
        the model list and keep the configured model if the server has it (otherwise warn – never auto-switch,
        never download).  Returns (model_available, message)."""
        problem = self.set_endpoint(result.host, result.port)
        if problem:
            return False, problem
        self.endpoint_epoch += 1                     # confirmed discovery always wins over a pending autoconnect
        self.ollama_ok, self.ollama_status = None, ""
        self.save_ollama_settings()
        ok, msg, models = self.refresh_models(timeout=timeout)
        if not ok and result.models:
            models = list(result.models)
            self.available_models = models
        LOG.info("DISCOVERY_APPLY server=%s models=%s", self.server, len(models))
        if not self.model.strip():
            return False, missing_model_message(DEFAULT_MODEL)
        if model_available(self.model, models):
            return True, f"Đã chuyển sang server {self.endpoint_label}. Model {self.model} có sẵn."
        return False, missing_model_message(self.model)

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
                            period=self.effective_period()[0], learning_dir=self.learning_dir)

    def start(self, use_ollama: bool = True, in_thread: bool = True) -> bool:
        """Start the production batch; returns False when validation fails or already running."""
        if not self.can_start():
            return False
        if self.validate():
            return False
        if self.scan_result is None:
            self.scan()                               # no reviewed list yet -> run the real pre-scan now
            if self.scan_result is None:
                return False
        if self.scan_stale:
            return False
        # the processing queue = the FINAL reviewed list (manual exclusions never reach the pipeline; a new
        # Excel row is created only when its candidate is actually processed)
        self._set_files([it.path for it in self.final_queue()])
        self.save_settings()
        for r in self.rows:
            r.stage, r.note, r.vendor, r.model, r.item = "waiting", "", "", "", ""
        self.progress = Progress(total=len(self.files))
        self.summary = None
        self.prescan, self.queue_indexes = None, None
        self.report_durations = []
        self._report_started_at = None
        self.started_at, self.finished_at = self._clock(), None
        self.done_events, self.stale_events, self.worker_failure = 0, 0, ""
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

    def worker_alive(self) -> bool:
        return bool(self.processor and self.processor.is_running())

    def reconcile(self) -> str:
        """Safety net (#9): the controller thinks a batch is running, the worker thread is dead and the queue is
        drained, yet no "done" was applied.  Finalize from the worker's own summary when it ended normally,
        otherwise surface a worker failure.  Returns the reason text ("" when nothing was reconciled)."""
        if not self.is_running() or self.processor is None or self.processor.is_running() or not self._queue.empty():
            return ""
        if getattr(self.processor, "_thread", None) is None:
            return ""                                 # synchronous run (tests) – no thread to watch
        summary = getattr(self.processor, "summary", None)
        total = len(self.files)
        finished_rows = self._count_done()
        if summary is not None and (summary.total >= total or finished_rows >= total) and total > 0:
            reason = "WORKER_RECONCILE: worker ended without a done event – finalized from its summary"
            self.log_lines.append(reason)
            self.apply_event(UiEvent("done", summary))
            return reason
        self.worker_failure = "Lỗi tiến trình xử lý: luồng xử lý đã dừng bất thường (không có kết quả kết thúc)."
        reason = f"WORKER_DIED: {self.worker_failure}"
        self.log_lines.append(reason)
        for r in self.rows:
            if not r.is_final and r.stage != "waiting":
                r.stage, r.note = "error", self.worker_failure
        self.apply_event(UiEvent("done", summary or BatchSummary(total=total, failed=total - finished_rows)))
        return reason

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
            try:
                self.apply_event(ev)
            except Exception as e:  # noqa: BLE001 – one malformed event must not stop the stream (#2)
                self.log_lines.append(f"GUI_EVENT_ERROR event={ev.kind} error={type(e).__name__}: {e}")
                LOG.exception("GUI_EVENT_ERROR event=%s", ev.kind)
            applied.append(ev)
        return applied

    def _in_queue(self, index: int) -> bool:
        return self.queue_indexes is None or index in self.queue_indexes

    def _count_done(self) -> int:
        return sum(1 for r in self.rows if r.is_final and self._in_queue(r.index))

    def apply_event(self, ev: UiEvent) -> None:
        if self.progress.finished and ev.kind in ("prescan", "row", "progress"):
            # stale worker event after finalization: must never put the GUI back into "Đang xử lý"
            if ev.kind == "row" and ev.payload[1] in FINAL_STATUSES and 0 <= ev.payload[0] < len(self.rows):
                self.rows[ev.payload[0]].stage = ev.payload[1]          # a late final status is still truth
                self._fill_row_from_result(self.rows[ev.payload[0]])
            else:
                self.stale_events += 1
            return
        if ev.kind == "scan":
            pass                                     # scan_result already set by scan(); the view re-renders
        elif ev.kind == "prescan":
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
                if was_final and stage not in FINAL_STATUSES:
                    self.stale_events += 1           # late stage event for a finished report: never regress (#6)
                    return
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
            if self.finished_at is not None:         # duplicate done: idempotent (timer/summary/progress untouched)
                self.done_events += 1
                return
            self.done_events += 1
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
        elif ev.kind == "discovery_progress":
            self.discovery_checked, self.discovery_total = ev.payload
        elif ev.kind == "discovery_found":
            if all(r.endpoint != ev.payload.endpoint for r in self.discovery_results):
                self.discovery_results.append(ev.payload)
        elif ev.kind == "discovery_done":
            results, checked, total, cancelled = ev.payload
            self.discovery_results = list(results)
            self.discovery_checked, self.discovery_total = checked, total
            self.discovery_state = "idle"
            self.discovery_message = discovery_summary_vi(results, checked, total, cancelled)

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
        """'Còn khoảng: MM:SS' while running, 'Còn khoảng: Đang tính...' before data exists, 'Đã hoàn thành' after a complete batch, '' otherwise."""
        if not self.is_running():
            if self.finished_at is not None and self.progress.total > 0 and self.progress.done >= self.progress.total:
                return "Đã hoàn thành"
            return ""
        eta = self.eta_seconds()
        return "Còn khoảng: Đang tính..." if eta is None else f"Còn khoảng: {format_elapsed(eta)}"

    def status_line(self) -> str:
        """One-line status for the window: progress text + bar + elapsed + ETA (controller state only)."""
        parts = [self.progress.text, self.progress.bar_text, self.elapsed_text(), self.eta_text()]
        return "   ".join(x for x in parts if x)

    def prescan_lines(self) -> List[str]:
        """Pre-scan counters (Vietnamese) of the reviewed list (+ manual decisions) – empty before any scan."""
        res = self.scan_result or self.prescan
        if not res:
            return []
        lines = res.summary_lines_vi()
        if self.scan_result:
            lines.append(self.queue_text())
        return lines

    def summary_lines(self) -> List[str]:
        s = self.summary or BatchSummary(total=len(self.files))
        lines = [f"Tổng: {s.total}", f"Hoàn thành: {s.completed}", f"Cần kiểm tra: {s.needs_review}",
                 f"Management Number mới: {s.new_rows}", f"Chưa ghi: {s.not_written}", f"Lỗi: {s.failed}",
                 f"Bỏ qua: {s.skipped}"]
        if self.scan_result or self.prescan:
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
                f"Management Number mới: {s.new_rows}   Chưa ghi: {s.not_written}   Lỗi: {s.failed}   Bỏ qua: {s.skipped}")

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
            "Dòng sử dụng": str(fr.excel_row) if fr.excel_row else "(không ghi)",
            "Management Number bị trùng tại dòng": (", ".join(map(str, fr.duplicate_rows)) + " – Đã đánh dấu đỏ các dòng trùng."
                                                  if fr.duplicate_rows else "(không)"),
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
    # ------------------------------------------------------------------ PROMPT-005 offline update
    def current_version_text(self) -> str:
        return f"Phiên bản hiện tại: {version_label()}"

    def set_update_path(self, path: str) -> None:
        new = (path or "").strip().strip('"')
        if new != self.update_path:
            self.update_path = new
            self.update_check = None
            self.update_dirty = True

    def update_status_text(self) -> str:
        if self.update_busy:
            return "Đang kiểm tra cập nhật…"
        if self.update_check is None:
            return "Chưa kiểm tra cập nhật." if self.update_path else updater.MSG["no_path"]
        if self.update_check_startup and self.update_check.status not in UPDATE_RESULT_STATUSES:
            return AUTO_CHECK_FAILED_VI                   # start-up check: never alarming, app keeps working
        return self.update_check.message

    def set_auto_update_check(self, enabled: bool) -> None:
        self.auto_update_check = bool(enabled)
        self.cfg.extra[AUTO_UPDATE_CHECK_KEY] = self.auto_update_check
        self.save_settings()

    def pending_update_offer(self) -> Optional[UpdateCheck]:
        """The newer remote build to announce now, at most ONCE per remote build per session (None otherwise)."""
        chk = self.update_check
        if not (chk and chk.available and chk.info) or self.update_busy:
            return None
        if chk.info.build in self.offered_update_builds:
            return None
        self.offered_update_builds.add(chk.info.build)
        return chk

    def update_available(self) -> bool:
        return bool(self.update_check and self.update_check.available) and not self.update_busy \
            and not self.update_installing

    def update_target_text(self) -> str:
        """'Build 010 → Build 011' for the progress window header."""
        cur = updater.format_build(updater.BUILD_NUMBER)
        new = updater.format_build(self.update_check.info.build) if self.update_check and self.update_check.info else "?"
        return f"Build {cur} → Build {new}"

    def check_update(self, startup: bool = False) -> UpdateCheck:
        """Synchronous read-only check – reads version.json only, never copies the ZIP (worker thread / tests)."""
        self.update_busy = True
        self.update_check_startup = bool(startup)
        try:
            res = updater.check_for_update(self.update_path)
        except Exception as e:  # noqa: BLE001 – infrastructure problem must never crash the GUI
            LOG.exception("UPDATE_CHECK unexpected error")
            res = UpdateCheck("inaccessible", f"{updater.MSG['inaccessible']} ({type(e).__name__})", self.update_path)
        self.update_check = res                      # result published BEFORE busy=False / done event (state contract)
        self.update_busy = False
        self.update_dirty = True
        self.update_check_done.set()
        return res

    def check_update_async(self, startup: bool = False) -> bool:
        """Non-blocking check; the startup variant silently does nothing when no path is configured."""
        if self._update_thread and self._update_thread.is_alive():
            return False
        if startup and (not self.update_path or not self.auto_update_check):
            return False
        self.update_busy = True
        self.update_dirty = True
        self.update_check_done.clear()
        self._update_thread = threading.Thread(target=self.check_update, args=(startup,), name="update-check",
                                               daemon=True)
        self._update_thread.start()
        return True

    def consume_update_notice(self) -> str:
        """One-shot message after a restart performed by the updater (empty when there is none)."""
        data = updater.consume_result(portable_root())
        if not data:
            return ""
        if data.get("status") == "ok":
            self.update_notice = f"Cập nhật thành công lên phiên bản {data.get('version', '')}."
            updater.cleanup_staging(portable_root())
        elif data.get("status") in ("rolled_back", "failed"):
            self.update_notice = f"Cập nhật không thành công – đã khôi phục phiên bản trước. {data.get('detail', '')}".strip()
        else:
            self.update_notice = f"Cập nhật thất bại ({data.get('status')}). {data.get('detail', '')}".strip()
        return self.update_notice

    def install_update(self, spawn=None) -> Tuple[bool, str]:
        """Synchronous stage + verify + launch of the external updater (tests / CLI).  The GUI uses
        install_update_async(); both share _install_work.  On success the caller must close the application."""
        if self.is_running():
            return False, "Đang xử lý báo cáo – hãy đợi xong rồi cập nhật."
        if not self.update_available():
            return False, "Không có bản cập nhật để cài."
        with self._install_lock:
            if self.update_installing:
                return False, "Đang cập nhật – vui lòng đợi."
            self.update_installing = True
        try:
            return self._install_work(spawn, emit=False)
        finally:
            if not (self.update_progress.stage == "HANDOFF"):
                self.update_installing = False

    def install_update_async(self, spawn=None) -> bool:
        """Start the install worker (LAN copy / verify / extract off the Tk thread); False when one is already
        running, nothing is available or a batch is processing.  Progress -> UiEvent('update_progress'),
        result -> UiEvent('update_done', (ok, msg))."""
        if self.is_running() or not self.update_available():
            return False
        with self._install_lock:
            if self.update_installing or (self._install_thread and self._install_thread.is_alive()):
                return False
            self.update_installing = True
        self.update_progress = UpdateProgress(stage="PREPARING")
        self.update_dirty = True

        def work():
            ok, msg = self._install_work(spawn, emit=True)
            if not ok:
                self.update_installing = False
                self.update_dirty = True
            self._queue.put(UiEvent("update_done", (ok, msg)))
        self._install_thread = threading.Thread(target=work, name="update-install", daemon=True)
        self._install_thread.start()
        return True

    def _publish_update_progress(self, stage: str, done: int, total: int, pct: float, emit: bool,
                                 t0: List[Optional[float]]) -> None:
        """Worker-side: build an immutable snapshot and queue it (NO widget access here)."""
        now = self._clock()
        if stage == "COPYING" and t0[0] is None:
            t0[0] = now
        elapsed = now - t0[0] if t0[0] is not None else 0.0
        speed = (done / elapsed) if (stage == "COPYING" and elapsed > 0 and done) else 0.0
        snap = UpdateProgress(stage=stage, bytes_copied=int(done), total_bytes=int(total), percent=float(pct),
                              speed_bps=speed, determinate=(stage in ("COPYING", "VERIFYING", "READY", "HANDOFF",
                                                                      "COMPLETE")))
        self.update_progress = snap
        if emit:
            self._queue.put(UiEvent("update_progress", snap))

    def _install_work(self, spawn, emit: bool) -> Tuple[bool, str]:
        t0: List[Optional[float]] = [None]
        try:
            staged = updater.stage_update(self.update_check, portable_root(),
                                          progress=lambda st, d, t, p: self._publish_update_progress(st, d, t, p,
                                                                                                     emit, t0))
            size = Path(staged.zip_path).stat().st_size
            self._publish_update_progress("HANDOFF", size, size, 100.0, emit, t0)   # verified: 100%, then hand off
            cmd = updater.launch_updater(staged, spawn) if spawn else updater.launch_updater(staged)
        except Exception as e:  # noqa: BLE001
            msg = str(e) or type(e).__name__
            LOG.error("UPDATE_INSTALL failed: %s", msg)
            updater.cleanup_staging(portable_root())
            prev = self.update_progress
            self.update_progress = UpdateProgress(stage="ERROR", bytes_copied=prev.bytes_copied,
                                                  total_bytes=prev.total_bytes, percent=prev.percent, error=msg)
            if emit:
                self._queue.put(UiEvent("update_progress", self.update_progress))
            return False, msg
        LOG.info("UPDATE_INSTALL updater launched: %s", cmd)
        return True, f"Đang cài đặt {self.update_check.info.label()} – ứng dụng sẽ tự khởi động lại."

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
        c.update_path = self.update_path
        try:
            c.save(self._config_path)
        except OSError as e:
            self.log_lines.append(f"Không lưu được config.json: {e}")
