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


@dataclass
class BatchSummary:
    total: int = 0
    completed: int = 0
    needs_review: int = 0
    failed: int = 0
    skipped: int = 0
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
            self._log(f"Kết thúc. Tổng: {self.summary.total}  Hoàn thành: {self.summary.completed}  "
                      f"Cần kiểm tra: {self.summary.needs_review}  Lỗi: {self.summary.failed}  "
                      f"Bỏ qua: {self.summary.skipped}")
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
            for n in cls.notes:
                LOG.info("%s: %s", path.name, n)

            # --- 3. extract original content (program copies WHAT) ----------
            rec = extract_record(report, cls, writer.item_mapping, writer.known_models)
            fr.management_number = rec.management_number
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
            status = "needs_review" if rec.review_reasons else "completed"
            note = "; ".join(rec.review_reasons)
            row = writer.append_record(rec, qpn_png, imp_jpg,
                                       status_text="Cần kiểm tra" if rec.review_reasons else "",
                                       note_text=note)
            writer.save()                       # save after every successful record
            fr.excel_row = row
            fr.review_reasons = list(rec.review_reasons)
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

    def _log(self, msg: str) -> None:
        LOG.info(msg)
        self.on_log(msg)
