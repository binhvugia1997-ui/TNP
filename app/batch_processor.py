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
from datetime import datetime
from pathlib import Path
from typing import Callable, Dict, List, Optional

from .classifier import classify
from .excel_writer import ExcelWriter
from .extractor import extract_record
from .history import History, fingerprint
from .image_extractor import build_improvement_image
from .logger import BatchResultLog, FileResult, setup_logging
from .ollama_client import OllamaClient
from .pptx_parser import parse_pptx
from .qpn_renderer import SlideRenderer, render_qpn

LOG = logging.getLogger("report_extractor.batch")

# Per-file stage identifiers (GUI maps them to Vietnamese labels)
STAGES = ["waiting", "reading", "analyzing", "extracting_qpn", "extracting_images",
          "writing_excel", "completed", "needs_review", "error", "skipped"]

STAGE_LABELS_VI = {
    "waiting": "Đang chờ",
    "reading": "Đang đọc PPTX",
    "analyzing": "Đang phân tích (AI)",
    "extracting_qpn": "Đang trích xuất QPN",
    "extracting_images": "Đang trích xuất hình ảnh",
    "writing_excel": "Đang ghi Excel",
    "completed": "Hoàn thành",
    "needs_review": "Cần kiểm tra",
    "error": "Lỗi",
    "skipped": "Đã xử lý trước đó",
    "not_written": "Cần kiểm tra (chưa ghi)",
}


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

    def as_dict(self) -> Dict:
        return self.__dict__.copy()


ProgressCB = Callable[[int, str, str], None]           # (file_index, stage, detail)
BatchCB = Callable[[int, int], None]                    # (done, total)
DoneCB = Callable[[BatchSummary], None]


class BatchProcessor:
    def __init__(self, opts: BatchOptions,
                 on_file: Optional[ProgressCB] = None,
                 on_batch: Optional[BatchCB] = None,
                 on_done: Optional[DoneCB] = None,
                 on_log: Optional[Callable[[str], None]] = None):
        self.opts = opts
        self.on_file = on_file or (lambda i, s, d: None)
        self.on_batch = on_batch or (lambda d, t: None)
        self.on_done = on_done or (lambda s: None)
        self.on_log = on_log or (lambda m: None)
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
            for idx, path in enumerate(opts.files):
                if self.stop_event.is_set():
                    self.summary.stopped = True
                    self._log("Đã dừng theo yêu cầu (trước file tiếp theo).")
                    break
                seq += 1
                fr = self._process_one(idx, Path(path), seq, writer, renderer, client, history, assets)
                self.results.append(fr)
                result_log.add(fr)
                done += 1
                self.on_batch(done, len(opts.files))
        finally:
            try:
                writer.save()
            except Exception as e:  # noqa: BLE001
                LOG.error("Final save failed: %s", e)
            writer.close()
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
            report = parse_pptx(path)
            if not report.slides:
                raise RuntimeError("PPTX không có slide nào")

            # --- 2. classify (AI decides WHERE) ------------------------------
            self.on_file(idx, "analyzing", f"{len(report.slides)} slide")
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
            rec = extract_record(report, cls, writer.item_mapping, writer.known_models, self.opts.vendors or None)
            fr.management_number = rec.management_number
            fr.vendor = rec.vendor
            fr.occurrence_date = rec.occurrence_date_text
            fr.model = rec.model
            fr.item = rec.item

            # --- 4. QPN full slide render -----------------------------------
            qpn_png: Optional[Path] = None
            if rec.qpn_slide:
                self.on_file(idx, "extracting_qpn", f"slide {rec.qpn_slide}")
                try:
                    qpn_png, backend = render_qpn(report, rec.qpn_slide, assets / f"{prefix}_QPN.png", renderer)
                    fr.qpn_renderer = backend
                except Exception as e:  # noqa: BLE001
                    LOG.warning("%s: QPN render failed: %s", path.name, e)
                    rec.review_reasons.append(f"Không render được QPN: {e}")

            # --- 5. improvement images ---------------------------------------
            imp_jpg: Optional[Path] = None
            if rec.improvement_image_slides:
                self.on_file(idx, "extracting_images", f"slide {rec.improvement_image_slides}")
                try:
                    imp_jpg, _ = build_improvement_image(
                        report, rec.improvement_image_slides, assets / f"{prefix}_IMPROVEMENT.jpg",
                        renderer, pictures_dir=assets / f"{prefix}_IMPROVEMENT_pics")
                except Exception as e:  # noqa: BLE001
                    LOG.warning("%s: improvement image failed: %s", path.name, e)
                    rec.review_reasons.append(f"Không render được hình ảnh cải tiến: {e}")

            # --- 6. write Excel ------------------------------------------------
            self.on_file(idx, "writing_excel", "")
            if self.opts.row_mode == "match":
                # Management Number (from file name) is the key of the destination row
                if not rec.management_number:
                    return self._not_written(idx, fr, rec, "Không xác định được Management Number từ tên file")
                rows = writer.find_rows_by_management_number(rec.management_number)
                if not rows:
                    return self._not_written(idx, fr, rec,
                                             f"Không tìm thấy Management Number {rec.management_number} trong file Kiểm chứng")
                if len(rows) > 1:
                    return self._not_written(idx, fr, rec,
                                             f"Management Number {rec.management_number} xuất hiện nhiều dòng ({rows})")
                row = rows[0]
                conflicts = writer.update_record(row, rec, qpn_png, imp_jpg, fill_temporary=self.opts.fill_temporary_column)
                rec.review_reasons.extend(conflicts)
                note = "; ".join(rec.review_reasons)
                if note and "note" in writer.columns:
                    writer._set_cell(row, "note", note)
                if rec.review_reasons and "status" in writer.columns:
                    writer._set_cell(row, "status", "Cần kiểm tra")
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
        """Match mode: destination row cannot be determined -> nothing is written, report it."""
        fr.status = "not_written"
        fr.review_reasons = [f"Cần kiểm tra: {reason}", *rec.review_reasons]
        fr.blank_fields = list(rec.blank_fields)
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


def format_file_diagnostics(fr: FileResult) -> str:
    """Human readable per-report diagnostics (also shown by the GUI)."""
    lines = [
        f"Báo cáo            : {fr.source_file}",
        f"Trạng thái         : {fr.status}" + (f"  (dòng Excel {fr.excel_row})" if fr.excel_row else ""),
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
        f"Trường để trống    : {', '.join(fr.blank_fields) if fr.blank_fields else '(không)'}",
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
