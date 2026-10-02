"""Batch pipeline: scan -> parse -> classify -> extract -> render -> Excel -> log.

Runs in a worker thread (see gui.py) and reports progress through callbacks.
One failing report never stops the batch; the workbook is saved after every
record so a crash loses at most the current file.
"""
from __future__ import annotations

import logging
import threading
import traceback
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path
from typing import Callable, Dict, List, Optional

from .classifier import classify
from .excel_writer import ExcelWriter
from .extractor import extract_record
from .history import History, fingerprint
from .extractor import management_number_from_filename
from .excel_writer import MANAGED_FIELDS
from .image_extractor import export_after_pictures
from .logger import BatchResultLog, FileResult, setup_logging
from .ollama_client import OllamaClient
from .pptx_parser import parse_pptx
from .prescan import (ACTION_FAST_SKIP, ACTION_INVALID_MGMT, ACTION_MASTER_COMPLETE, ACTION_MASTER_NOT_FOUND,
                      PROCESS_ACTIONS, CACHE_FILE_NAME,
                      FastScanCache, MasterLookup, PreScanItem, PreScanResult, ProcessingPeriod, prescan)
from .qpn_renderer import SlideRenderer, render_qpn_panel

LOG = logging.getLogger("report_extractor.batch")

# Per-file stage identifiers (GUI maps them to Vietnamese labels)
STAGES = ["waiting", "reading", "analyzing", "extracting_qpn", "extracting_images",
          "writing_excel", "completed", "needs_review", "error", "skipped"]

STAGE_LABELS_VI = {
    "waiting": "Đang chờ",
    "reading": "Đang đọc PPTX",
    "analyzing": "Đang phân tích Qwen",
    "analyzing_heuristic": "Đang phân tích (từ khoá, không AI)",
    "extracting": "Đang trích xuất nguyên nhân / đối sách cải tiến",
    "extracting_qpn": "Đang trích xuất QPN",
    "extracting_images": "Đang trích xuất hình ảnh cải tiến",
    "writing_excel": "Đang ghi Excel",
    "completed": "Hoàn thành",
    "needs_review": "Cần kiểm tra",
    "error": "Lỗi",
    "skipped": "Bỏ qua — đã cập nhật",
    "not_written": "Không tìm thấy Management Number",
    # pre-scan outcomes (decided from file name / cache / master workbook – PPTX never opened)
    "outside_period": "Bỏ qua ngoài thời gian xử lý",
    "source_duplicate": "Trùng Management Number trong folder",
    "fast_skip": "Bỏ qua nhanh — đã xử lý gần đây",
}
PRESCAN_STATUSES = ("outside_period", "source_duplicate", "fast_skip")


@dataclass
class BatchOptions:
    files: List[Path]
    template: Path
    output_file: Path
    ollama_server: str = ""
    model: str = ""
    force_reprocess: bool = False
    request_timeout: int = 180
    render_width: int = 1920
    use_ollama: bool = True
    fill_temporary_column: bool = False
    # "match": Management Number from the file name locates the EXISTING Excel row (production);
    # "append": every report becomes a new row (legacy / empty template)
    row_mode: str = "match"
    vendors: List[str] = field(default_factory=list)   # controlled Vendor list override (empty -> built-in)
    # fast pre-scan: processing period (None = all), recent-success cache switch, fixed "today" for tests
    period: Optional[ProcessingPeriod] = None
    use_fast_cache: bool = True
    today: Optional[date] = None


@dataclass
class BatchSummary:
    total: int = 0
    completed: int = 0
    needs_review: int = 0
    failed: int = 0
    skipped: int = 0
    not_written: int = 0
    stopped: bool = False
    output_file: str = ""
    output_folder: str = ""
    errors: List[Dict[str, str]] = field(default_factory=list)
    # pre-scan counters = PreScanResult.counts() (schema: prescan.PRESCAN_COUNT_KEYS) + the period actually used
    prescan: Dict[str, int] = field(default_factory=dict)
    candidates: int = 0
    period: str = ""
    outside_period: int = 0
    source_duplicates: int = 0
    fast_skipped: int = 0
    master_not_found: int = 0            # pre-scan: valid keys absent from the master (reported, never created)
    new_rows: int = 0                    # always 0 since PROMPT-004 (kept for result-log compatibility)
    backup_file: str = ""                # safety copy taken before the first real workbook modification

    def as_dict(self) -> Dict:
        return self.__dict__.copy()


ProgressCB = Callable[[int, str, str], None]           # (file_index, stage, detail)
BatchCB = Callable[[int, int], None]                    # (done, total)
DoneCB = Callable[[BatchSummary], None]
PreScanCB = Callable[[PreScanResult], None]


class BatchProcessor:
    def __init__(self, opts: BatchOptions,
                 on_file: Optional[ProgressCB] = None,
                 on_batch: Optional[BatchCB] = None,
                 on_done: Optional[DoneCB] = None,
                 on_log: Optional[Callable[[str], None]] = None,
                 on_prescan: Optional[PreScanCB] = None):
        self.opts = opts
        self.on_file = on_file or (lambda i, s, d: None)
        self.on_batch = on_batch or (lambda d, t: None)
        self.on_done = on_done or (lambda s: None)
        self.on_log = on_log or (lambda m: None)
        self.on_prescan = on_prescan or (lambda r: None)
        self.prescan_result: Optional[PreScanResult] = None
        self.cache: Optional[FastScanCache] = None
        self.stop_event = threading.Event()
        self.summary = BatchSummary(total=len(opts.files))
        self.results: List[FileResult] = []
        self._thread: Optional[threading.Thread] = None

    # ------------------------------------------------------------------
    def start(self) -> threading.Thread:
        self._thread = threading.Thread(target=self.run, name="batch-worker", daemon=True)
        self._thread.start()
        return self._thread

    def request_stop(self) -> None:
        self.stop_event.set()

    def is_running(self) -> bool:
        return bool(self._thread and self._thread.is_alive())

    # ------------------------------------------------------------------
    def run(self) -> BatchSummary:
        opts = self.opts
        out_dir = Path(opts.output_file).parent
        assets = out_dir / "assets"
        logs = out_dir / "logs"
        assets.mkdir(parents=True, exist_ok=True)
        setup_logging(logs)
        result_log = BatchResultLog(logs / "batch_result.json")
        history = History(logs / "history.json")
        self.summary.output_file = str(opts.output_file)
        self.summary.output_folder = str(out_dir)
        self._log(f"Bắt đầu xử lý {len(opts.files)} báo cáo → {opts.output_file}")

        client: Optional[OllamaClient] = None
        if opts.use_ollama and opts.ollama_server and opts.model:
            client = OllamaClient(opts.ollama_server, timeout=opts.request_timeout)
        else:
            self._log("Không dùng AI (chưa cấu hình máy chủ/model) – dùng nhận diện theo từ khoá")

        try:
            writer = ExcelWriter(opts.template, opts.output_file)
        except Exception as e:  # noqa: BLE001
            LOG.error("Cannot open template/output: %s", e)
            self._log(f"LỖI form Excel: {e}")
            for i in range(len(opts.files)):
                self.on_file(i, "error", str(e))
            self.summary.failed = len(opts.files)
            self.summary.errors.append({"file": "", "error": str(e)})
            result_log.finish(self.summary.as_dict())
            self.on_done(self.summary)
            return self.summary

        renderer = SlideRenderer(width_px=opts.render_width)
        seq = writer.used_count()
        done = 0
        try:
            # --- fast pre-scan: file name / period / duplicates / 7-day cache / master row (no PPTX) ---
            queue = self._prescan(writer, logs, result_log)
            total = len(queue)
            self.summary.candidates = total
            self.on_batch(0, total)
            for item in queue:
                idx, path = item.index, item.path
                if self.stop_event.is_set():
                    self.summary.stopped = True
                    self._log("Đã dừng theo yêu cầu (trước file tiếp theo).")
                    break
                seq += 1
                fr = self._process_one(idx, Path(path), seq, writer, renderer, client, history, assets)
                self.results.append(fr)
                result_log.add(fr)
                done += 1
                self.on_batch(done, total)
        finally:
            try:
                writer.save()
            except Exception as e:  # noqa: BLE001
                LOG.error("Final save failed: %s", e)
            self.summary.backup_file = str(writer.backup_path) if writer.backup_path else ""
            if writer.backup_path:
                self._log(f"Bản sao lưu Excel trước khi ghi: {writer.backup_path}")
            writer.close()
            order = {str(p): i for i, p in enumerate(opts.files)}
            self.results.sort(key=lambda r: order.get(r.source_file, len(order)))   # file order, like the GUI table
            result_log.finish(self.summary.as_dict())
            try:
                write_review_report(self.results, logs / "review_report.txt", self.summary)
            except Exception as e:  # noqa: BLE001
                LOG.warning("review_report.txt failed: %s", e)
            self._log(f"Kết thúc. Tổng: {self.summary.total}  Hoàn thành: {self.summary.completed}  "
                      f"Cần kiểm tra: {self.summary.needs_review}  Chưa ghi: {self.summary.not_written}  "
                      f"Lỗi: {self.summary.failed}  Bỏ qua: {self.summary.skipped}")
            self.on_done(self.summary)
        return self.summary

    # ------------------------------------------------------------------
    def _prescan(self, writer: ExcelWriter, logs: Path, result_log: BatchResultLog) -> List[PreScanItem]:
        """Run the cheap pre-scan, report every rejected file immediately, return the real queue."""
        opts = self.opts
        today = opts.today or date.today()
        period = opts.period
        if period is not None:
            LOG.info("%s period=%s%s", "PERIOD_AUTO" if period.mode == "auto" else "PERIOD_MANUAL", period.iso(),
                     f" excel={period.source}" if period.source else "")
        self.cache = None
        if opts.use_fast_cache:
            try:
                self.cache = FastScanCache(logs / CACHE_FILE_NAME)
                self.cache.cleanup(today)
            except Exception as e:  # noqa: BLE001 – cache is optional performance state
                LOG.warning("CACHE_PROBLEM reason=%s", e)
                self.cache = None
        if self.cache is not None and self.cache.problem:
            self._log(f"Cache quét nhanh không dùng được ({self.cache.problem}) – xử lý bình thường")
        master = MasterLookup.from_writer(writer) if opts.row_mode == "match" else None
        self._log("Đang quét thư mục...")
        res = prescan(opts.files, period, self.cache, master, force=opts.force_reprocess, today=today,
                      on_stage=self._log)
        self.prescan_result = res
        c = res.counts()
        self.summary.prescan = c
        self.summary.period = period.label_vi() if period else "Thời gian xử lý: Tất cả"
        self.summary.outside_period = c["outside_period"]
        self.summary.source_duplicates = c["source_duplicates"]
        self.summary.fast_skipped = c["fast_skipped"]
        self.summary.master_not_found = c["master_not_found"]
        self.on_prescan(res)
        for line in res.summary_lines_vi():
            self._log(line)
        # report the rejected files right away (they never enter the processing queue)
        for it in res.items:
            if it.action in PROCESS_ACTIONS:
                continue
            fr = FileResult(source_file=str(it.path), started_at=datetime.now().isoformat(timespec="seconds"),
                            management_number=it.management_number)
            fr.occurrence_date = f"{it.occurrence_date:%d/%m/%Y}" if it.occurrence_date else ""
            if it.action in (ACTION_INVALID_MGMT, ACTION_MASTER_NOT_FOUND):
                self._not_written(it.index, fr, None, it.reason.replace("Cần kiểm tra: ", ""))
            elif it.action == ACTION_MASTER_COMPLETE:
                self._skip_complete(it.index, fr, writer, it.management_number, it.path)
            else:
                fr.status = it.status                    # outside_period | source_duplicate | fast_skip
                fr.error = it.reason or it.label_vi
                fr.excel_row = it.excel_row
                fr.finished_at = datetime.now().isoformat(timespec="seconds")
                if it.action == ACTION_FAST_SKIP:
                    self.summary.skipped += 1
                self.on_file(it.index, it.status, it.label_vi)
            self.results.append(fr)
            result_log.add(fr)
        return res.candidates

    def _skip_complete(self, idx: int, fr: FileResult, writer: ExcelWriter, mgmt: str, path: Path) -> FileResult:
        """Complete master row -> `Bỏ qua — đã cập nhật` (no parse, no Qwen, no rewrite); duplicate Excel rows
        are still highlighted red; the outcome is safe for the 7-day cache."""
        rows = writer.find_rows_by_management_number(mgmt)
        row = rows[0] if rows else None
        fr.duplicate_rows = self._handle_duplicate_rows(writer, mgmt, rows, path)
        fr.status = "skipped"
        fr.excel_row = row
        fr.error = f"Bỏ qua — đã cập nhật (dòng {row})"
        fr.finished_at = datetime.now().isoformat(timespec="seconds")
        self.summary.skipped += 1
        LOG.info("%s: row %s already complete – skipped (no Qwen, no extraction)", path.name, row)
        self._remember(mgmt, path, "skipped")
        self.on_file(idx, "skipped", fr.error)
        return fr

    def _handle_duplicate_rows(self, writer: ExcelWriter, mgmt: str, rows: List[int], path: Path) -> List[int]:
        """Duplicate Management Number rows (PROMPT-004 §15): the TOPMOST data row is the canonical destination,
        the other rows keep their business data untouched and are marked red (the batch continues, the report may
        still finish as Hoàn thành).  Returns the extra (red) rows."""
        if len(rows) <= 1:
            return []
        extra = rows[1:]
        writer.mark_rows_red(extra)                        # backup is taken by the writer before the first change
        LOG.warning("%s: Management Number %s – Dòng sử dụng: %s; Management Number bị trùng tại dòng: %s; "
                    "Đã đánh dấu đỏ các dòng trùng.", path.name, mgmt, rows[0], ", ".join(map(str, extra)))
        return extra

    def _remember(self, mgmt: str, path: Path, status: str) -> None:
        """Feed the recent-success cache; any failure is logged and ignored (never blocks the batch)."""
        if self.cache is None:
            return
        try:
            self.cache.record(mgmt, path, status)
        except Exception as e:  # noqa: BLE001
            LOG.warning("CACHE_PROBLEM management_number=%s reason=%s", mgmt, e)

    # ------------------------------------------------------------------
    def _process_one(self, idx: int, path: Path, seq: int, writer: ExcelWriter, renderer: SlideRenderer,
                     client: Optional[OllamaClient], history: History, assets: Path) -> FileResult:
        fr = FileResult(source_file=str(path), started_at=datetime.now().isoformat(timespec="seconds"))
        prefix = f"{seq:04d}"
        try:
            # --- duplicate protection --------------------------------------
            try:
                fr.fingerprint = fingerprint(path)
            except Exception as e:  # noqa: BLE001
                raise RuntimeError(f"Không đọc được file: {e}") from e
            row: Optional[int] = None
            missing: Optional[List[str]] = None
            if self.opts.row_mode == "match":
                # Fixed-master rules: the destination row decides BEFORE any parsing / Qwen call.
                mgmt = management_number_from_filename(path.name)
                fr.management_number = mgmt
                if not mgmt:
                    return self._not_written(idx, fr, None, "Không xác định được Management Number từ tên file")
                rows = writer.find_rows_by_management_number(mgmt)
                if not rows:
                    # PROMPT-004: never create a production row automatically – report and leave the file for a rerun
                    return self._not_written(idx, fr, None,
                                             f"Không tìm thấy Management Number {mgmt} trong Excel – không tạo dòng mới")
                fr.duplicate_rows = self._handle_duplicate_rows(writer, mgmt, rows, path)
                row = rows[0]
                missing = writer.missing_managed_fields(row)
                if not missing and not self.opts.force_reprocess:
                    fr.status = "skipped"
                    fr.excel_row = row
                    fr.error = f"Bỏ qua — đã cập nhật (dòng {row})"
                    self.summary.skipped += 1
                    LOG.info("%s: row %s already complete – skipped (no Qwen, no extraction)", path.name, row)
                    self._remember(mgmt, path, "skipped")
                    self.on_file(idx, "skipped", fr.error)
                    return fr
                if missing and len(missing) < len([f for f in MANAGED_FIELDS if f in writer.columns]) \
                        and not self.opts.force_reprocess:
                    LOG.info("%s: row %s partially complete – filling only %s", path.name, row, missing)
                    if _row_looks_created_by_us(writer, row):
                        LOG.info("MASTER_NEW_RETRY management_number=%s existing_partial_row=%s", mgmt, row)
            else:
                prev = history.lookup(fr.fingerprint)
                if prev and not self.opts.force_reprocess and Path(prev.get("output", "")) == Path(self.opts.output_file):
                    fr.status = "skipped"
                    fr.excel_row = prev.get("row")
                    fr.error = f"Đã xử lý trước đó (dòng {prev.get('row')})"
                    self.summary.skipped += 1
                    self.on_file(idx, "skipped", fr.error)
                    return fr

            # --- 1. read PPTX ------------------------------------------------
            self.on_file(idx, "reading", "")
            if path.suffix.lower() == ".ppt":
                raise RuntimeError("Định dạng .ppt cũ không đọc được – hãy mở bằng PowerPoint và lưu lại thành .pptx")
            report = parse_pptx(path)
            if not report.slides:
                raise RuntimeError("PPTX không có slide nào")

            # --- 2. classify (AI decides WHERE) ------------------------------
            self.on_file(idx, "analyzing" if client else "analyzing_heuristic", f"{len(report.slides)} slide")
            cls = classify(report, client, self.opts.model if client else "")
            fr.classifier = cls.source
            fr.qpn_slide = cls.qpn_slide
            fr.cause_slides = list(cls.cause_slides)
            fr.improvement_slides = list(cls.improvement_slides)
            fr.improvement_image_slides = list(cls.improvement_image_slides)
            fr.temporary_slides = list(cls.temporary_slides)
            fr.verify_slides = list(cls.verify_slides)
            fr.defect_slide = cls.defect_slide
            fr.confidence = cls.confidence
            fr.classifier_notes = list(cls.notes)
            fr.qpn_source = cls.qpn_source
            fr.qpn_override = cls.qpn_override
            fr.image_slides_structural = list(cls.image_slides_structural)
            fr.image_slides_llm = list(cls.image_slides_llm)
            for n in cls.notes:
                LOG.info("%s: %s", path.name, n)
            LOG.info("%s: classifier=%s qpn=%s cause=%s temp=%s improvement=%s images=%s conf=%s",
                     path.name, cls.source, cls.qpn_slide, cls.cause_slides, cls.temporary_slides,
                     cls.improvement_slides, cls.improvement_image_slides, cls.confidence)

            # --- 3. extract original content (program copies WHAT) ----------
            self.on_file(idx, "extracting", "")
            rec = extract_record(report, cls, writer.item_mapping, writer.known_models, self.opts.vendors or None)
            fr.management_number = rec.management_number
            fr.after_picture_slides = list(rec.after_picture_slides)
            fr.vendor = rec.vendor
            fr.occurrence_date = rec.occurrence_date_text
            fr.model = rec.model
            fr.item = rec.item

            # --- 4. QPN = the Quality Problem Notice PANEL only (object/region based, fail-closed) ---------
            qpn_png: Optional[Path] = None
            if rec.qpn_slide:
                self.on_file(idx, "extracting_qpn", f"slide {rec.qpn_slide}")
                try:
                    qr = render_qpn_panel(report, rec.qpn_slide, assets / f"{prefix}_QPN.png", renderer)
                    if qr.ok:
                        qpn_png = qr.path
                        fr.qpn_image = str(qr.path)
                        fr.qpn_renderer = qr.backend
                        fr.qpn_region = qr.region.describe(report.slide_width, report.slide_height)
                        fr.qpn_excluded = list(qr.region.excluded)
                    else:
                        fr.qpn_region = f"không tách được: {qr.reason}"
                        rec.review_reasons.append(f"Cần kiểm tra: Không tách được QPN khỏi slide {rec.qpn_slide} "
                                                  f"({qr.reason}) – không chèn ảnh cả slide")
                except Exception as e:  # noqa: BLE001
                    LOG.warning("%s: QPN render failed: %s", path.name, e)
                    rec.review_reasons.append(f"Không render được QPN: {e}")

            # --- 5. improvement images: ONLY "Sau cải tiến" pictures -----------
            imp_jpg: List[Path] = []           # one PNG per After picture -> independent Excel images
            fr.after_pictures = [r.label for r in rec.after_pictures]
            fr.picture_notes = list(rec.picture_notes)
            fr.excluded_sections = list(rec.excluded_sections)
            if rec.after_pictures:
                self.on_file(idx, "extracting_images", f"{len(rec.after_pictures)} ảnh Sau cải tiến")
                try:
                    imp_jpg, problems = export_after_pictures(report, rec.after_pictures,
                                                              assets / f"{prefix}_IMPROVEMENT_pics")
                    rec.review_reasons.extend(problems)
                except Exception as e:  # noqa: BLE001
                    LOG.warning("%s: improvement image failed: %s", path.name, e)
                    rec.review_reasons.append(f"Không tạo được hình ảnh cải tiến: {e}")

            # --- 6. write Excel ------------------------------------------------
            self.on_file(idx, "writing_excel", "")
            if self.opts.row_mode == "match":
                assert row is not None and missing is not None
                if self.opts.force_reprocess:
                    # explicit user request: rewrite every extractor-managed field
                    conflicts = writer.update_record(row, rec, qpn_png, imp_jpg,
                                                     fill_temporary=self.opts.fill_temporary_column)
                else:
                    # auto fill: ONLY the blank fields, populated cells are preserved
                    conflicts = writer.fill_missing_fields(row, rec, missing, qpn_png, imp_jpg)
                    rec.review_reasons = filter_review_reasons(rec.review_reasons, missing)
                    fr.filled_fields = [f for f in missing if _rec_has(rec, f, qpn_png, imp_jpg)]
                rec.review_reasons.extend(conflicts)
                note = "; ".join(rec.review_reasons)
                if "status" in writer.columns:
                    cell, _ = writer._anchor(row, writer.columns["status"])
                    if rec.review_reasons:
                        writer._set_cell(row, "status", "Cần kiểm tra")
                    elif str(cell.value or "").strip() == "Cần kiểm tra":
                        writer._set_cell(row, "status", "")        # our earlier flag, now resolved
                        if "note" in writer.columns:
                            writer._set_cell(row, "note", "")
                if note and "note" in writer.columns:
                    writer._set_cell(row, "note", note)
            else:
                note = "; ".join(rec.review_reasons)
                row = writer.append_record(rec, qpn_png, imp_jpg,
                                           status_text="Cần kiểm tra" if rec.review_reasons else "",
                                           note_text=note, fill_temporary=self.opts.fill_temporary_column)
            status = "needs_review" if rec.review_reasons else "completed"
            note = "; ".join(rec.review_reasons)
            writer.save()                       # save after every successful record
            fr.excel_row = row
            fr.review_reasons = list(rec.review_reasons)
            fr.blank_fields = list(rec.blank_fields)
            fr.status = status
            history.record(fr.fingerprint, path, Path(self.opts.output_file), row, status)
            if status == "completed":
                self._remember(fr.management_number, path, "completed")   # needs_review is never safe-cached
                self.summary.completed += 1
            else:
                self.summary.needs_review += 1
            self.on_file(idx, status, note or f"dòng {row}")
            LOG.info("%s -> row %s (%s) %s", path.name, row, status, note)
        except Exception as e:  # noqa: BLE001
            fr.status = "error"
            fr.error = f"{type(e).__name__}: {e}"
            self.summary.failed += 1
            self.summary.errors.append({"file": str(path), "error": fr.error})
            LOG.error("%s: %s\n%s", path.name, fr.error, traceback.format_exc())
            self.on_file(idx, "error", fr.error)
        fr.finished_at = datetime.now().isoformat(timespec="seconds")
        return fr

    def _not_written(self, idx: int, fr: FileResult, rec, reason: str) -> FileResult:
        """Match mode: destination row cannot be determined -> nothing is written, report it.

        Nothing has been parsed or sent to Qwen at this point; the report stays eligible for a rerun once the
        user adds the Management Number to the master workbook.
        """
        fr.status = "not_written"
        fr.review_reasons = [f"Cần kiểm tra: {reason}", *(rec.review_reasons if rec else [])]
        fr.blank_fields = list(rec.blank_fields) if rec else []
        fr.error = f"Cần kiểm tra: {reason}"
        fr.finished_at = datetime.now().isoformat(timespec="seconds")
        self.summary.not_written += 1
        self.summary.errors.append({"file": fr.source_file, "error": fr.error})
        LOG.warning("%s: not written – %s", Path(fr.source_file).name, reason)
        self.on_file(idx, "not_written", reason)
        return fr

    def _log(self, msg: str) -> None:
        LOG.info(msg)
        self.on_log(msg)


# review reason -> managed field it concerns (reasons about already-populated fields are dropped in fill mode)
_REASON_FIELD = (("Nội dung lỗi", "defect_content"), ("QPN", "qpn"), ("Nguyên nhân", "root_cause"),
                 ("Nội dung đối sách cải tiến", "improvement"), ("hình ảnh cải tiến", "improvement_image"),
                 ("Model", "model"), ("Vendor", "vendor"), ("vendor", "vendor"))


def filter_review_reasons(reasons: List[str], missing: List[str]) -> List[str]:
    """Keep only reasons that concern a field we actually had to fill (or generic ones)."""
    kept: List[str] = []
    for r in reasons:
        fld = next((f for key, f in _REASON_FIELD if key in r), None)
        if fld is None or fld in missing:
            kept.append(r)
    return kept


def _row_looks_created_by_us(writer: ExcelWriter, row: int) -> bool:
    """A row holding ONLY the key (every managed text field blank) is typically a row created by a previous run
    whose extraction failed – it is simply retried through the incremental logic (never appended again)."""
    try:
        missing = writer.missing_managed_fields(row)
        return all(f in missing for f in MANAGED_FIELDS if f in writer.columns and f != "occurrence_date")
    except Exception:  # noqa: BLE001
        return False


def _rec_has(rec, field: str, qpn_png, imp_jpg) -> bool:
    if field == "qpn":
        return bool(qpn_png)
    if field == "improvement_image":
        return bool(imp_jpg)
    if field == "occurrence_date":
        return bool(rec.occurrence_date)
    return bool(getattr(rec, field, ""))


def format_file_diagnostics(fr: FileResult) -> str:
    """Human readable per-report diagnostics (also shown by the GUI)."""
    lines = [
        f"Báo cáo            : {fr.source_file}",
        f"Trạng thái         : {fr.status}" + (f"  (dòng Excel {fr.excel_row})" if fr.excel_row else ""),
        f"Dòng sử dụng       : {fr.excel_row if fr.excel_row else '(không ghi)'}",
        f"Management Number bị trùng tại dòng: {', '.join(map(str, fr.duplicate_rows)) if fr.duplicate_rows else '(không)'}"
        + ("  – Đã đánh dấu đỏ các dòng trùng." if fr.duplicate_rows else ""),
        f"Bộ phân loại       : {fr.classifier or '-'}" + (f"  (độ tin cậy AI {fr.confidence:.2f})" if fr.confidence is not None else ""),
        f"Management number  : {fr.management_number or '(trống)'}  (từ tên file – khoá dòng Excel)",
        f"Vendor (danh sách chuẩn, từ đối sách): {fr.vendor.replace(chr(10), ' / ') if fr.vendor else '(trống)'}",
        f"Ngày phát sinh     : {fr.occurrence_date or '(trống)'}  (suy ra từ Management Number YYMMDD)",
        f"Model              : {fr.model or '(trống)'}",
        f"Item               : {fr.item or '(trống)'}",
        f"Slide QPN          : {fr.qpn_slide or '(không thấy)'}",
        f"Slide nội dung lỗi : {fr.defect_slide or '-'}",
        f"Slide nguyên nhân  : {fr.cause_slides or '(không thấy)'}",
        f"Slide xử lý tạm thời: {fr.temporary_slides or '-'}  (loại khỏi đối sách)",
        f"Slide đối sách     : {fr.improvement_slides or '(không thấy)'}",
        f"Slide hình cải tiến: {fr.improvement_image_slides or '(không có)'}  (cấu trúc {fr.image_slides_structural}, AI {fr.image_slides_llm})",
        f"Nguồn QPN          : {fr.qpn_source or '-'}" + (f"  – {fr.qpn_override}" if fr.qpn_override else ""),
        f"Slide kiểm chứng   : {fr.verify_slides or '-'}  (WEEK +1..+8 luôn để trống)",
        f"Renderer QPN       : {fr.qpn_renderer or '-'}",
        f"Vùng QPN           : {fr.qpn_region or '-'}",
        f"Loại khỏi QPN      : {'; '.join(fr.qpn_excluded) if fr.qpn_excluded else '-'}",
        f"Mục loại khỏi cải tiến: {'; '.join(fr.excluded_sections) if fr.excluded_sections else '-'}",
        f"Trường để trống    : {', '.join(fr.blank_fields) if fr.blank_fields else '(không)'}",
        f"Trường đã điền     : {', '.join(fr.filled_fields) if fr.filled_fields else '(không)'}",
        f"Ảnh Sau cải tiến   : {', '.join(fr.after_pictures) if fr.after_pictures else '(không)'} "
        f"(slide {fr.after_picture_slides})",
        f"Cần kiểm tra thủ công: {'; '.join(fr.review_reasons) if fr.review_reasons else '(không)'}",
    ]
    if fr.classifier_notes:
        lines.append("Ghi chú phân loại  : " + " | ".join(fr.classifier_notes))
    if fr.error:
        lines.append(f"Lỗi                : {fr.error}")
    return "\n".join(lines)


def write_review_report(results: List[FileResult], path: Path, summary: BatchSummary) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    head = (f"BÁO CÁO KIỂM TRA TRÍCH XUẤT – {datetime.now():%Y-%m-%d %H:%M}\n"
            f"Tổng: {summary.total}  Hoàn thành: {summary.completed}  Cần kiểm tra: {summary.needs_review}  "
            f"Lỗi: {summary.failed}  Bỏ qua: {summary.skipped}\n"
            f"Nguyên tắc: thông tin không có trong báo cáo được để trống + 'Cần kiểm tra' (không suy đoán); "
            f"WEEK +1..+8 luôn giữ nguyên.\n")
    body = "\n\n".join(("=" * 78) + "\n" + format_file_diagnostics(fr) for fr in results)
    path.write_text(head + "\n" + body + "\n", encoding="utf-8")
    return path
