"""Local application services for the React + pywebview development runtime.

This module is an adapter around :class:`GuiController`, not a second processing implementation.  All PPTX,
pre-scan, extraction, Excel, learning, Ollama, log and update decisions continue to live in the existing Python
backend.  The browser receives bounded DTOs and opaque IDs; it never receives an arbitrary file-read/write API.
"""
from __future__ import annotations

import base64
import io
import json
import logging
import os
import re
import secrets
import subprocess
import sys
import threading
import zipfile
from datetime import date, datetime
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional

from . import APP_NAME, APP_TITLE, BUILD_ID, BUILD_NUMBER, __version__
from .excel_writer import ExcelLockedError, ExcelWriter, TemplateError
from .gui_controller import GuiController
from .extractor import management_number_from_filename
from .prescan import (ACTION_FAST_SKIP, ACTION_INVALID_MGMT, ACTION_MASTER_COMPLETE, ACTION_OUTSIDE_PERIOD,
                      ACTION_PROCESS, ACTION_PROCESS_NEW_ROW, ACTION_SOURCE_DUPLICATE, normalize_source_path)
from .preview_geometry import slide_fraction_box
from .runtime_paths import config_dir, logs_dir

LOG = logging.getLogger("report_extractor.webview")


class ServiceError(RuntimeError):
    """Expected, user-actionable service failure with a stable bridge error code."""

    def __init__(self, message: str, code: str = "OPERATION_FAILED"):
        super().__init__(message)
        self.code = code


class ApplicationService:
    """Small, JSON-safe application facade around the production ``GuiController``."""

    def __init__(self, controller: Optional[GuiController] = None, config_path: Optional[Path] = None,
                 manual_fields_path: Optional[Path] = None):
        self.controller = controller or GuiController(config_path=config_path)
        self._lock = threading.RLock()
        self._manual_fields_path = Path(manual_fields_path) if manual_fields_path else (
            Path(config_path).parent / "manual_fields.json" if config_path else config_dir() / "manual_fields.json")
        self._manual_fields: Dict[str, Dict[str, str]] = {}
        self._pending_manual: Dict[str, Dict[str, str]] = {}
        self._report_paths: Dict[str, Path] = {}
        self._file_tokens: Dict[str, Path] = {}
        self._active_file_token = ""
        self._log_entries: List[dict] = []
        self._log_cursor = 0
        self._log_sequence = 0
        self._last_excel_result = {"kind": "none", "message": ""}
        self._diagnostics: List[dict] = []
        self._diagnostics_running = False
        self._diagnostics_error = ""
        self._diagnostics_thread: Optional[threading.Thread] = None
        self._training = False
        self._training_message = ""
        self._training_thread: Optional[threading.Thread] = None
        self._server_apply_running = False
        self._server_apply_message = ""
        self._server_apply_thread: Optional[threading.Thread] = None
        self._preview_cache: Dict[str, str] = {}
        # PROMPT-024R: FULL rendered slide per (report path, slide, mtime) for the review workspace.
        # Rendered once and shared by every candidate of that report+slide; never shared across
        # reports (§40/§54). Value: (data_uri, width_px, height_px).
        self._slide_preview_cache: Dict[tuple, tuple] = {}
        self._load_manual_fields()

    # ------------------------------------------------------------------ static app/config DTOs
    @staticmethod
    def app_version() -> dict:
        return {"name": APP_NAME, "title": APP_TITLE, "version": __version__,
                "build": BUILD_ID, "buildNumber": BUILD_NUMBER, "runtime": "pywebview-development"}

    def current_config(self) -> dict:
        with self._lock:
            c = self.controller
            return {
                "paths": {"reportFolder": c.report_folder, "template": c.template, "output": c.output},
                "period": {"mode": c.period_mode, "month": c.period_month, "year": c.period_year,
                           "from": c.period_from, "to": c.period_to},
                "forceReprocess": bool(c.force_reprocess),
                "ollama": {"host": c.host, "port": str(c.port), "model": c.model,
                           "checked": "ok" if c.ollama_ok is True else "fail" if c.ollama_ok is False else "unchecked",
                           "message": c.ollama_status or c.ai_status_text(),
                           "models": list(c.available_models)},
                "update": {"path": c.update_path, "autoCheck": bool(c.auto_update_check),
                           "status": "checking" if c.update_busy else "idle",
                           "message": c.update_status_text(), "available": None},
                "app": self.app_version(),
            }

    def save_configuration(self, dto: dict) -> dict:
        data = self._object(dto, "configuration")
        with self._lock:
            self._assert_idle("Không thể thay đổi cài đặt khi đang xử lý báo cáo.")
            c = self.controller
            paths = self._object(data.get("paths", {}), "paths")
            folder = self._text(paths.get("reportFolder", c.report_folder), "reportFolder", 32767)
            template = self._text(paths.get("template", c.template), "template", 32767)
            output = self._text(paths.get("output", c.output), "output", 32767)
            if folder != c.report_folder:
                c.input_files_override = None
                c.report_folder = folder
                c.excluded_keys.clear()
                c._invalidate_scan()
            c.set_template(template)
            c.set_output(output)
            c.set_force(self._boolean(data.get("forceReprocess", c.force_reprocess), "forceReprocess"))

            period = self._object(data.get("period", {}), "period")
            mode = self._choice(period.get("mode", c.period_mode), "period.mode", {"auto", "all", "month", "range"})
            problem = c.set_period(mode, period.get("month", c.period_month), period.get("year", c.period_year),
                                   period.get("from", c.period_from), period.get("to", c.period_to))
            if problem:
                raise ServiceError(problem, "INVALID_PERIOD")

            ollama = self._object(data.get("ollama", {}), "ollama")
            host = self._text(ollama.get("host", c.host), "ollama.host", 512)
            port = self._text(ollama.get("port", c.port), "ollama.port", 8)
            model = self._text(ollama.get("model", c.model), "ollama.model", 256)
            problem = c.set_endpoint(host, port, model)
            if problem:
                raise ServiceError(problem, "INVALID_OLLAMA_ENDPOINT")

            update = self._object(data.get("update", {}), "update")
            update_path = self._text(update.get("path", c.update_path), "update.path", 32767)
            c.set_update_path(update_path)
            c.set_auto_update_check(self._boolean(update.get("autoCheck", c.auto_update_check), "update.autoCheck"))
            c.save_settings()
            return self.current_config()

    # ------------------------------------------------------------------ file-selection capabilities
    def register_pptx_selection(self, selected_path: str) -> dict:
        path = Path(selected_path).expanduser()
        if path.suffix.lower() != ".pptx" or not path.is_file():
            raise ServiceError("Chọn một file PPTX hiện có.", "INVALID_SELECTION")
        resolved = path.resolve()
        token = secrets.token_urlsafe(24)
        with self._lock:
            self._file_tokens.clear()
            self._file_tokens[token] = resolved
            self._active_file_token = token
        return {"token": token, "name": resolved.name, "folder": str(resolved.parent)}

    def clear_pptx_selection(self) -> None:
        with self._lock:
            self._file_tokens.clear()
            self._active_file_token = ""

    # ------------------------------------------------------------------ scan and reports
    def scan_reports(self, config: dict, input_file_token: str = "") -> dict:
        data = self._object(config, "configuration")
        with self._lock:
            self._assert_idle("Không thể quét trong khi đang xử lý báo cáo.")
            self._apply_configuration_for_scan(data)
            selected: Optional[Path] = None
            if input_file_token:
                token = self._text(input_file_token, "inputFileToken", 128)
                selected = self._file_tokens.get(token)
                if selected is None or token != self._active_file_token:
                    raise ServiceError("Lựa chọn file đã hết hạn. Hãy chọn lại file PPTX.", "INVALID_SELECTION")
                if selected.suffix.lower() != ".pptx" or not selected.is_file():
                    raise ServiceError("File PPTX đã chọn không còn tồn tại.", "INVALID_SELECTION")
                self.controller.input_files_override = [selected]
                self.controller.report_folder = str(selected.parent)
            else:
                self.controller.input_files_override = None
                folder = Path(self.controller.report_folder).expanduser() if self.controller.report_folder else None
                if folder is None or not folder.is_dir():
                    raise ServiceError("Thư mục báo cáo không tồn tại hoặc chưa được chọn.", "INVALID_REPORT_FOLDER")
                self.controller.report_folder = str(folder.resolve())
            result = self.controller.scan()
            self.controller.save_settings()
            if result is None:
                message = self.controller.scan_message or "Không thể hoàn tất bước quét trước xử lý."
            else:
                message = self.controller.queue_text()
            LOG.info("WEBVIEW_SCAN finished=%s files=%d", bool(result), len(self.controller.all_files))
            return self.dashboard_state(message_override=message)

    def dashboard_state(self, message_override: Optional[str] = None) -> dict:
        with self._lock:
            c = self.controller
            c.pump()
            c.reconcile()
            c.pump()
            self._collect_logs()
            rows = c.scan_rows("Tất cả file đã quét")
            reports = []
            self._report_paths = {}
            for row in rows:
                report_id = str(row.index)
                self._report_paths[report_id] = Path(row.path)
                self._manual_fields.setdefault(normalize_source_path(Path(row.path)), {})
                fields = self._manual_fields.get(normalize_source_path(Path(row.path)), {})
                result = self._result_for_path(Path(row.path))
                reports.append(self._report_dto(report_id, row, result, fields, len(reports) + 1))
            progress = c.progress
            # PROMPT-024R: distinct job states for the two stop modes. "cancelling" is shown while the
            # worker has not yet acknowledged cancel-all; the UI only shows "Đã dừng" after "done".
            run_status = ("cancelling" if c.state == "cancelling" else
                          "stopping" if c.state == "stopping" else
                          "processing" if c.state == "running" else
                          ("done" if progress.finished else "idle"))
            current_name = ""
            if progress.current_index is not None and 0 <= progress.current_index < len(c.rows):
                current_name = c.rows[progress.current_index].path.name
            candidate_ids = [str(row.index) for row in rows
                             if row.action in (ACTION_PROCESS, ACTION_PROCESS_NEW_ROW) and not row.excluded]
            done_at = int(c.finished_at * 1000) if c.finished_at is not None else None
            started_at = int(c.started_at * 1000) if c.started_at is not None else None
            return {
                "app": self.app_version(),
                "config": self.current_config(),
                "scan": {"scanned": c.scan_result is not None,
                         "message": self._redact_source_paths(message_override or c.scan_message or
                                                             (c.queue_text() if c.scan_result else "Chưa quét thư mục."))},
                "reports": reports,
                "job": {"status": run_status, "queue": candidate_ids if run_status == "idle" else
                        [str(i) for i in range(progress.total)],
                        "index": (progress.current_index or 0), "doneCount": progress.done,
                        "stage": progress.current_stage or "waiting", "percent": progress.percent,
                        "currentFile": current_name, "elapsedSec": c.elapsed_seconds(),
                        "remainSec": c.eta_seconds(), "startedAt": started_at, "finishedAt": done_at,
                        "hasSamples": c.average_report_seconds() is not None,
                        "stopped": bool(c.summary and c.summary.stopped),
                        # PROMPT-024R: cancel-all acknowledgement is surfaced separately from graceful stop.
                        "cancelRequested": (c.state == "cancelling"
                                            or bool(c.summary and getattr(c.summary, "cancel_requested", False))),
                        "cancelledCount": int(getattr(c.summary, "cancelled", 0)) if c.summary else 0,
                        "error": self._redact_source_paths(c.worker_failure or "")},
                "logs": list(self._log_entries[-500:]),
                "ollama": {"host": c.host, "port": c.port, "model": c.model,
                           "checked": "ok" if c.ollama_ok is True else "fail" if c.ollama_ok is False else "unchecked",
                           "message": c.ollama_status or c.ai_status_text(), "aiStatus": c.ai_status_text(),
                           "models": list(c.available_models), "discovering": c.discovery_running,
                           "discoveryMessage": c.discovery_progress_text(),
                           "discoveryResults": [self._discovery_dto(x) for x in c.discovery_results],
                           "discoveryChecked": c.discovery_checked, "discoveryTotal": c.discovery_total,
                           "serverApplyRunning": self._server_apply_running,
                           "serverApplyMessage": self._server_apply_message},
                "update": self._update_dto(),
                "diagnostics": {"running": self._diagnostics_running,
                                "rows": [{**row, "label": self._redact_source_paths(row.get("label", "")),
                                          "value": self._redact_source_paths(row.get("value", ""))}
                                         for row in self._diagnostics],
                                "error": self._redact_source_paths(self._diagnostics_error)},
            }

    def _apply_configuration_for_scan(self, data: dict) -> None:
        c = self.controller
        paths = self._object(data.get("paths", {}), "paths")
        folder = self._text(paths.get("reportFolder", c.report_folder), "reportFolder", 32767)
        c.report_folder = folder
        c.set_template(self._text(paths.get("template", c.template), "template", 32767))
        c.set_output(self._text(paths.get("output", c.output), "output", 32767))
        c.set_force(self._boolean(data.get("forceReprocess", c.force_reprocess), "forceReprocess"))
        period = self._object(data.get("period", {}), "period")
        mode = self._choice(period.get("mode", c.period_mode), "period.mode", {"auto", "all", "month", "range"})
        problem = c.set_period(mode, period.get("month", c.period_month), period.get("year", c.period_year),
                               period.get("from", c.period_from), period.get("to", c.period_to))
        if problem:
            raise ServiceError(problem, "INVALID_PERIOD")
        ollama = self._object(data.get("ollama", {}), "ollama")
        problem = c.set_endpoint(self._text(ollama.get("host", c.host), "ollama.host", 512),
                                 self._text(ollama.get("port", c.port), "ollama.port", 8),
                                 self._text(ollama.get("model", c.model), "ollama.model", 256))
        if problem:
            raise ServiceError(problem, "INVALID_OLLAMA_ENDPOINT")
        update = self._object(data.get("update", {}), "update")
        c.set_update_path(self._text(update.get("path", c.update_path), "update.path", 32767))
        c.set_auto_update_check(self._boolean(update.get("autoCheck", c.auto_update_check), "update.autoCheck"))

    def _report_dto(self, report_id: str, row, fr, manual_fields: dict, stt: int) -> dict:
        status = self._status_key(row.action, row.excluded, row.run_stage)
        warning = row.reason or row.note or ""
        if fr is not None and fr.review_reasons:
            warning = "; ".join(fr.review_reasons)
        date_text = row.occurrence_date or ""
        vendor_text = row.vendor or ""
        if fr is not None:
            date_text = fr.occurrence_date or date_text
            vendor_text = fr.vendor or vendor_text
        if "occurrence_date" in manual_fields:
            date_text = self._iso_to_dmy(manual_fields["occurrence_date"]) if manual_fields["occurrence_date"] else ""
        if "vendor" in manual_fields:
            vendor_text = manual_fields["vendor"]
        short = row.note or row.reason or row.scan_status_vi
        if fr is not None:
            if fr.status in ("error", "not_written"):
                short = fr.error or "; ".join(fr.review_reasons) or short
            elif fr.review_reasons:
                short = "; ".join(fr.review_reasons)
            elif fr.status == "skipped":
                short = fr.error or short
        short = self._redact_source_paths(short)
        warning = self._redact_source_paths(warning)
        return {"id": report_id, "stt": stt, "managementNumber": row.management_number,
                "fileName": row.path.name, "path": "", "occurrenceDate": date_text,
                "manualFields": {k: v for k, v in manual_fields.items()},
                "manualPending": normalize_source_path(Path(row.path)) in self._pending_manual, "vendor": vendor_text,
                "slides": fr.slide_count if fr is not None else 0, "status": status,
                "shortResult": short, "warning": warning or None,
                "processedAt": getattr(fr, "finished_at", "") if fr is not None else "",
                "durationText": self._duration_text(fr) if fr is not None else "",
                "excelRow": row.excel_row or (fr.excel_row if fr is not None else None),
                "excelPath": self.controller.output if row.excel_row or (fr and fr.excel_row) else None,
                "results": {"qpn": None, "causes": [], "countermeasures": [], "temporaryRemoved": False,
                            "afterImages": []}}

    @staticmethod
    def _status_key(action: str, excluded: bool, run_stage: str) -> str:
        if excluded:
            return "excluded"
        final_map = {"completed": "completed", "completed_new": "completed", "needs_review": "needs_review",
                     "not_written": "needs_review", "error": "error", "skipped": "skipped",
                     # PROMPT-024R: user cancellation is its own terminal status – never mapped to "error".
                     "cancelled": "cancelled",
                     "outside_period": "outside_period", "source_duplicate": "source_duplicate", "fast_skip": "fast_skip"}
        if run_stage in final_map:
            return final_map[run_stage]
        if run_stage and run_stage != "waiting":
            return "processing"
        action_map = {ACTION_PROCESS: "waiting", ACTION_PROCESS_NEW_ROW: "new_row",
                      ACTION_MASTER_COMPLETE: "skipped", ACTION_OUTSIDE_PERIOD: "outside_period",
                      ACTION_SOURCE_DUPLICATE: "source_duplicate", ACTION_FAST_SKIP: "fast_skip",
                      ACTION_INVALID_MGMT: "needs_review"}
        return action_map.get(action, "waiting")

    # ------------------------------------------------------------------ processing/worker lifecycle
    def start_processing(self, config: dict, use_ollama: bool = True) -> dict:
        data = self._object(config, "configuration")
        with self._lock:
            self._assert_idle("Đang xử lý báo cáo.")
            self._apply_configuration_for_scan(data)
            self._retry_pending_manual_fields()
            c = self.controller
            if c.scan_result is None:
                raise ServiceError("Hãy quét thư mục và xem lại danh sách trước khi xử lý.", "SCAN_REQUIRED")
            errors = c.validate()
            if errors:
                raise ServiceError("\n".join(errors), "VALIDATION_FAILED")
            if not c.start(use_ollama=bool(use_ollama), in_thread=True):
                raise ServiceError(c.scan_message or "Không thể bắt đầu xử lý. Hãy quét lại danh sách.",
                                   "START_FAILED")
            LOG.info("WEBVIEW_BATCH_STARTED candidates=%d", len(c.files))
            return self.dashboard_state()

    def stop_after_current(self) -> dict:
        with self._lock:
            if not self.controller.request_stop():
                raise ServiceError("Không có file nào đang được xử lý.", "NOT_RUNNING")
            return self.dashboard_state()

    def cancel_all(self) -> dict:
        """PROMPT-024R "Dừng tất cả": cooperative cancel-all.

        Returns a JSON-safe acknowledgement. Cancel while idle is harmless and deterministic (§49);
        duplicate cancel clicks are idempotent (§48). No threading primitive ever crosses the bridge.
        """
        with self._lock:
            requested = self.controller.request_cancel()
            if requested:
                LOG.info("WEBVIEW_CANCEL_ALL_REQUESTED")
            return {"ok": True, "requested": bool(requested), "mode": "cancel_all",
                    **self.dashboard_state()}

    def exclude_reports(self, report_ids: Iterable[str], restore: bool = False) -> dict:
        with self._lock:
            self._assert_idle("Không thể đổi hàng đợi khi đang xử lý báo cáo.")
            ids = self._validated_report_ids(report_ids)
            indexes = [int(x) for x in ids]
            message = self.controller.restore(indexes) if restore else self.controller.exclude(indexes)
            if message:
                raise ServiceError(message, "QUEUE_UPDATE_FAILED")
            return self.dashboard_state()

    def report_details(self, report_id: str) -> dict:
        with self._lock:
            rid = self._text(report_id, "reportId", 24)
            if not rid.isdigit() or rid not in self._report_paths:
                raise ServiceError("Báo cáo không còn trong danh sách đã quét.", "REPORT_NOT_FOUND")
            path = self._report_paths[rid]
            row = next((r for r in self.controller.scan_rows("Tất cả file đã quét") if str(r.index) == rid), None)
            if row is None:
                raise ServiceError("Báo cáo không còn trong danh sách đã quét.", "REPORT_NOT_FOUND")
            fr = self._result_for_path(path)
            details = {"id": rid, "causes": [], "countermeasures": [], "afterImages": [], "qpn": None,
                       "defectText": "", "vendor": "", "occurrenceDate": row.occurrence_date or "",
                       "model": "", "item": "", "slides": 0, "excelRow": row.excel_row,
                       "temporaryRemoved": False}
            if fr is not None:
                details.update({"causes": fr.cause_sections or ([fr.cause_text] if fr.cause_text else []),
                                "countermeasures": fr.improvement_sections or
                                ([fr.improvement_text] if fr.improvement_text else []),
                                "defectText": fr.defect_text, "vendor": fr.vendor,
                                "occurrenceDate": fr.occurrence_date or details["occurrenceDate"],
                                "model": fr.model, "item": fr.item, "slides": fr.slide_count,
                                "excelRow": fr.excel_row or row.excel_row,
                                "temporaryRemoved": bool(fr.temporary_slides),
                                "warning": "; ".join(fr.review_reasons) if fr.review_reasons else fr.error})
                if fr.qpn_image:
                    image = self._asset_data_uri(Path(fr.qpn_image))
                    if image:
                        details["qpn"] = {"src": image, "slide": fr.qpn_slide or 0, "caption": "QPN — rendered evidence"}
                details["afterImages"] = [{"src": src, "slide": fr.after_picture_slides[i] if i < len(fr.after_picture_slides) else 0,
                                           "caption": f"Ảnh Sau cải tiến {i + 1}"}
                                          for i, src in enumerate(self._asset_data_uri(Path(p)) for p in fr.after_assets) if src]
                if not fr.cause_text and not fr.improvement_text and not fr.defect_text:
                    self._excel_report_details(details, row)
            else:
                self._excel_report_details(details, row)
            if details.get("warning"):
                details["warning"] = self._redact_source_paths(details["warning"])
            return details

    def _excel_report_details(self, details: dict, row) -> None:
        c = self.controller
        if not c.output or not Path(c.output).is_file() or not c.template or not Path(c.template).is_file():
            return
        writer = None
        try:
            writer = ExcelWriter(Path(c.template), Path(c.output), probe=True)
            excel_rows = writer.find_rows_by_management_number(row.management_number)
            if not excel_rows:
                return
            excel_row = excel_rows[0]
            details["excelRow"] = excel_row
            for field, key in (("root_cause", "causes"), ("improvement", "countermeasures"),
                               ("defect_content", "defectText"), ("vendor", "vendor"),
                               ("occurrence_date", "occurrenceDate"), ("model", "model"), ("item", "item")):
                col = writer.columns.get(field)
                if not col:
                    continue
                cell, _ = writer._anchor(excel_row, col)
                value = cell.value
                if value is None:
                    continue
                if field == "occurrence_date":
                    parsed = writer._as_date(value)
                    details[key] = parsed.strftime("%d/%m/%Y") if parsed else str(value)
                elif field in ("root_cause", "improvement"):
                    details[key] = [str(value)] if str(value).strip() else []
                else:
                    details[key] = str(value)
        except Exception as exc:  # read-only detail must not break report listing
            LOG.warning("WEBVIEW_EXCEL_DETAIL_FAILED file=%s error=%s", row.path.name, exc)
        finally:
            if writer is not None:
                writer.close()

    # ------------------------------------------------------------------ explicit manual Vendor/date edits
    def save_manual_fields(self, report_id: str, fields: dict) -> dict:
        rid = self._text(report_id, "reportId", 24)
        if not rid.isdigit():
            raise ServiceError("Báo cáo không hợp lệ.", "INVALID_REPORT_ID")
        submitted = self._object(fields, "manualFields")
        allowed = {"vendor", "occurrence_date"}
        if not submitted or set(submitted) - allowed:
            raise ServiceError("Chỉ Vendor và Ngày phát sinh có thể chỉnh sửa.", "INVALID_MANUAL_FIELDS")
        validated: Dict[str, str] = {}
        if "vendor" in submitted:
            validated["vendor"] = self._validate_vendor(submitted["vendor"])
        if "occurrence_date" in submitted:
            validated["occurrence_date"] = self._validate_date_to_iso(submitted["occurrence_date"])
        with self._lock:
            self._assert_idle("Không thể chỉnh sửa khi đang xử lý báo cáo.")
            if rid not in self._report_paths:
                raise ServiceError("Báo cáo không còn trong danh sách đã quét.", "REPORT_NOT_FOUND")
            path = self._report_paths[rid]
            key = normalize_source_path(path)
            merged = dict(self._manual_fields.get(key, {}))
            merged.update(validated)
            self._manual_fields[key] = merged
            self._pending_manual[key] = dict(merged)
            self.controller.manual_fields_by_path = {k: dict(v) for k, v in self._manual_fields.items()}
            self._save_manual_fields()
            result = self._try_apply_manual_fields(path, merged)
            if result["status"] == "applied":
                self._pending_manual.pop(key, None)
                message = f"Đã lưu và cập nhật trường đã chọn vào dòng Excel {result['row']}."
            elif result["status"] == "locked":
                message = "Nhãn chỉnh sửa đã lưu an toàn; file Excel đang bị khóa. Đóng Excel rồi bấm Thử lại."
            elif result["status"] == "pending":
                self._pending_manual.pop(key, None)
                message = "Đã lưu. Trường sẽ được ghi khi báo cáo được thêm vào file kết quả."
            else:
                message = result.get("message", "Đã lưu trường thủ công.")
            self._save_manual_fields()
            LOG.info("WEBVIEW_MANUAL_FIELDS report=%s fields=%s status=%s", path.name, sorted(validated), result["status"])
            return {"status": result["status"], "message": message, "row": result.get("row"),
                    "state": self.dashboard_state()}

    def retry_manual_fields(self, report_id: str) -> dict:
        rid = self._text(report_id, "reportId", 24)
        with self._lock:
            self._assert_idle("Không thể cập nhật Excel khi đang xử lý báo cáo.")
            path = self._report_paths.get(rid)
            if path is None:
                raise ServiceError("Báo cáo không còn trong danh sách đã quét.", "REPORT_NOT_FOUND")
            key = normalize_source_path(path)
            fields = self._pending_manual.get(key) or self._manual_fields.get(key)
            if not fields:
                raise ServiceError("Không có trường thủ công nào đang chờ cập nhật.", "NO_PENDING_MANUAL_FIELDS")
            result = self._try_apply_manual_fields(path, fields)
            if result["status"] == "applied":
                self._pending_manual.pop(key, None)
                message = f"Đã cập nhật Excel dòng {result['row']}."
            elif result["status"] == "locked":
                message = "Excel vẫn đang bị khóa; các trường đã lưu và có thể thử lại."
            else:
                self._pending_manual.pop(key, None)
                message = "Chưa có dòng Excel cho báo cáo này; giá trị đã lưu để xử lý sau."
            self._save_manual_fields()
            return {"status": result["status"], "message": message, "state": self.dashboard_state()}

    def _try_apply_manual_fields(self, path: Path, fields: dict) -> dict:
        c = self.controller
        if not c.output or not Path(c.output).is_file() or not c.template or not Path(c.template).is_file():
            return {"status": "pending"}
        mgmt = management_number_from_filename(path.name)
        if not mgmt:
            return {"status": "pending"}
        writer = None
        try:
            writer = ExcelWriter(Path(c.template), Path(c.output))
            missing = [field for field in fields if field not in writer.columns]
            if missing:
                return {"status": "error", "message": "Form Excel thiếu cột: " + ", ".join(missing)}
            rows = writer.find_rows_by_management_number(mgmt)
            if not rows:
                return {"status": "pending"}
            writer.apply_manual_fields(rows[0], fields)
            writer.save()
            return {"status": "applied", "row": rows[0]}
        except ExcelLockedError as exc:
            return {"status": "locked", "message": str(exc)}
        except (OSError, TemplateError, ValueError) as exc:
            LOG.exception("WEBVIEW_MANUAL_FIELDS_FAILED file=%s", path.name)
            return {"status": "error", "message": f"Không cập nhật được Excel: {exc}"}
        finally:
            if writer is not None:
                writer.close()

    def _retry_pending_manual_fields(self) -> None:
        if not self._pending_manual:
            return
        failures = []
        for key, fields in list(self._pending_manual.items()):
            path = next((p for p in self._report_paths.values() if normalize_source_path(p) == key), None)
            if path is None:
                continue
            result = self._try_apply_manual_fields(path, fields)
            if result["status"] in ("applied", "pending"):
                self._pending_manual.pop(key, None)
            elif result["status"] == "locked":
                failures.append(f"{path.name}: file Excel đang được sử dụng")
            elif result["status"] == "error":
                failures.append(f"{path.name}: {result.get('message', 'lỗi Excel')}")
        self._save_manual_fields()
        if failures:
            raise ServiceError("\n".join(failures), "MANUAL_EXCEL_PENDING")

    # ------------------------------------------------------------------ Ollama, settings, update and diagnostics
    def check_ollama_connection(self) -> dict:
        with self._lock:
            if self.controller.is_running():
                raise ServiceError("Không thể kiểm tra Ollama trong khi đang xử lý.", "BUSY")
            self.controller.check_ollama_async(local_first=True)
            return self.dashboard_state()

    def refresh_ollama_models(self) -> dict:
        with self._lock:
            if self.controller.is_running():
                raise ServiceError("Không thể làm mới model trong khi đang xử lý.", "BUSY")
            self.controller.refresh_models_async()
            return self.dashboard_state()

    def start_ollama_discovery(self) -> dict:
        with self._lock:
            if not self.controller.discover_ollama_async():
                raise ServiceError("Không thể bắt đầu tìm kiếm Ollama lúc này.", "BUSY")
            return self.dashboard_state()

    def stop_ollama_discovery(self) -> dict:
        with self._lock:
            self.controller.cancel_discovery()
            return self.dashboard_state()

    def use_discovered_server(self, host: str, port: int) -> dict:
        host = self._text(host, "host", 512)
        try:
            port_num = int(port)
        except (TypeError, ValueError) as exc:
            raise ServiceError("Port không hợp lệ.", "INVALID_OLLAMA_ENDPOINT") from exc
        with self._lock:
            self._assert_idle("Không thể đổi server khi đang xử lý tác vụ nền.")
            if self._server_apply_running:
                raise ServiceError("Đang áp dụng server Ollama.", "BUSY")
            result = next((r for r in self.controller.discovery_results
                           if r.host == host and r.port == port_num), None)
            if result is None:
                raise ServiceError("Server không còn trong kết quả tìm kiếm.", "SERVER_NOT_FOUND")
            self._server_apply_running = True
            self._server_apply_message = "Đang xác nhận server và lấy model..."

        def work() -> None:
            try:
                ok, message = self.controller.apply_discovered_server(result)
                if ok:
                    self.controller.save_settings()
                with self._lock:
                    self._server_apply_message = message
                    self._server_apply_running = False
            except Exception as exc:  # worker result is surfaced as a safe message
                with self._lock:
                    self._server_apply_message = str(exc)
                    self._server_apply_running = False
                LOG.warning("WEBVIEW_OLLAMA_APPLY_FAILED host=%s error=%s", host, exc)

        self._server_apply_thread = threading.Thread(target=work, daemon=True, name="webview-ollama-apply")
        self._server_apply_thread.start()
        return self.dashboard_state()

    def run_diagnostics(self) -> dict:
        with self._lock:
            if self._diagnostics_running:
                return self.dashboard_state()
            self._assert_idle("Không thể chạy chẩn đoán trong khi một tác vụ nền đang chạy.")
            self._diagnostics_running = True
            self._diagnostics_error = ""
            cfg = self.controller.cfg
            template = self.controller.template or None
            output_folder = str(Path(self.controller.output).parent) if self.controller.output else None

        def work() -> None:
            try:
                from .diagnostics import run_diagnostics
                rows = run_diagnostics(cfg, template, output_folder, check_ollama=True, smoke=False)
                dto = [{"label": label, "value": value,
                        "state": "error" if "LỖI" in value or "Không hợp lệ" in value else
                                 "warn" if "Chưa" in value or "Không kết nối" in value else "ok"}
                       for label, value in rows]
                with self._lock:
                    self._diagnostics = dto
            except Exception as exc:  # diagnostics must not crash the desktop app
                with self._lock:
                    self._diagnostics_error = f"Chạy chẩn đoán thất bại: {exc}"
            finally:
                with self._lock:
                    self._diagnostics_running = False

        self._diagnostics_thread = threading.Thread(target=work, daemon=True, name="webview-diagnostics")
        self._diagnostics_thread.start()
        return self.dashboard_state()

    def check_update(self) -> dict:
        with self._lock:
            self.controller.check_update_async(startup=False)
            return self.dashboard_state()

    def install_update(self) -> dict:
        return {"status": "unavailable", "message":
                "Cài đặt bản cập nhật bị tắt trong tích hợp phát triển; không thay thế hay khởi động lại ứng dụng.",
                "state": self.dashboard_state()}

    # ------------------------------------------------------------------ learning: real candidates, labels, Excel reapply, model training
    def learning_state(self) -> dict:
        with self._lock:
            c = self.controller
            lrn = c.learning
            if lrn is None:
                image_status, content_status = "Không khả dụng.", "Không khả dụng."
                images: List[dict] = []
                contents: List[dict] = []
                image_total_ids: set = set()
                image_labeled_ids: set = set()
                content_total_ids: set = set()
                content_labeled_ids: set = set()
            else:
                image_status = self._model_status(lrn.model, lrn.model_status, "ảnh")
                content_status = self._model_status(lrn.content.model, lrn.content.model_status, "nội dung")
                image_candidates = c.review_candidates()
                content_candidates = c.review_content_candidates()
                images = [self._image_candidate_dto(candidate, lrn) for candidate in image_candidates]
                contents = [self._content_candidate_dto(candidate, lrn) for candidate in content_candidates]
                saved_image_ids = set(lrn.store.latest())
                saved_content_ids = set(lrn.content.store.latest())
                image_labeled_ids = saved_image_ids | set(c.pending_labels)
                content_labeled_ids = saved_content_ids | set(c.pending_content_labels)
                image_total_ids = saved_image_ids | {x.candidate_id for x in image_candidates}
                content_total_ids = saved_content_ids | {x.candidate_id for x in content_candidates}
            pending = c.image_reapply_pending + c.content_reapply_pending
            return {"images": images, "contents": contents,
                    "counts": {"image": {"total": len(image_total_ids), "labeled": len(image_labeled_ids)},
                               "content": {"total": len(content_total_ids), "labeled": len(content_labeled_ids)}},
                    "modelStatus": {"image": image_status, "content": content_status},
                    "excelPending": pending, "excelLastResult": dict(self._last_excel_result),
                    "training": self._training, "trainingMessage": self._training_message,
                    "reviewSummary": c.review_summary_text(), "contentReviewSummary": c.content_review_summary_text()}

    def set_learning_label(self, kind: str, candidate_id: str, label: str, note: str = "") -> dict:
        kind = self._choice(kind, "kind", {"image", "content"})
        candidate_id = self._text(candidate_id, "candidateId", 512)
        note = self._text(note, "note", 1000)
        with self._lock:
            self._assert_idle("Không thể gán nhãn trong khi đang xử lý báo cáo.")
            c = self.controller
            if kind == "image":
                candidate = next((item for item in c.review_candidates() if item.candidate_id == candidate_id), None)
                if candidate is None:
                    raise ServiceError("Ảnh này không còn trong danh sách học hiện tại.", "CANDIDATE_NOT_FOUND")
                c.set_pending_label(candidate, label, note)
            else:
                candidate = next((item for item in c.review_content_candidates()
                                  if item.candidate_id == candidate_id), None)
                if candidate is None:
                    raise ServiceError("Khối nội dung này không còn trong danh sách học hiện tại.", "CANDIDATE_NOT_FOUND")
                c.set_pending_content_label(candidate, label, note)
            return self.learning_state()

    def save_learning_labels(self, kind: str, notes: Optional[dict] = None) -> dict:
        kind = self._choice(kind, "kind", {"image", "content"})
        notes = self._object(notes or {}, "learningNotes")
        with self._lock:
            self._assert_idle("Không thể lưu nhãn trong khi đang xử lý báo cáo.")
            c = self.controller
            if kind == "image":
                for candidate_id in c.pending_labels:
                    if candidate_id in notes:
                        c.pending_label_notes[candidate_id] = self._text(notes[candidate_id], "note", 1000)
            else:
                for candidate_id in c.pending_content_labels:
                    if candidate_id in notes:
                        c.pending_content_notes[candidate_id] = self._text(notes[candidate_id], "note", 1000)
            if kind == "image":
                ok, message = c.save_confirmations(c.review_candidates())
                pending = c.image_reapply_pending
            else:
                ok, message = c.save_content_confirmations(c.review_content_candidates())
                pending = c.content_reapply_pending
            if not ok:
                raise ServiceError(message, "LABEL_SAVE_FAILED")
            status = "locked" if pending or "chưa cập nhật được Excel" in message.lower() or "đang được sử dụng" in message.lower() else "ok"
            self._last_excel_result = {"kind": status, "message": message}
            LOG.info("WEBVIEW_LABELS_SAVED kind=%s pending_excel=%d", kind, pending)
            return {"message": message, "state": self.learning_state()}

    def dismiss_learning_excel_notice(self) -> dict:
        """Clear only the transient result message; preserve saved labels and any pending Excel retry."""
        with self._lock:
            self._last_excel_result = {"kind": "none", "message": ""}
            return self.learning_state()

    def retry_learning_excel(self) -> dict:
        with self._lock:
            self._assert_idle("Không thể cập nhật Excel trong khi đang xử lý báo cáo.")
            c = self.controller
            messages = []
            had_pending = bool(c.image_reapply_pending or c.content_reapply_pending)
            if c.content_reapply_pending:
                ok, msg = c.retry_content_reapply()
                messages.append(msg)
            if c.image_reapply_pending:
                ok_img, msg = c.retry_image_reapply()
                messages.append(msg)
            if not had_pending:
                messages.append("Không có cập nhật Excel nào đang chờ.")
            pending = c.image_reapply_pending + c.content_reapply_pending
            message = "\n".join(x for x in messages if x)
            self._last_excel_result = {"kind": "locked" if pending else "ok", "message": message}
            return {"message": message, "state": self.learning_state()}

    def train_models(self) -> dict:
        with self._lock:
            self._assert_idle("Không thể huấn luyện khi đang xử lý báo cáo.")
            if self._training:
                return self.learning_state()
            self._training = True
            self._training_message = "Đang huấn luyện mô hình ảnh và nội dung…"

        def work() -> None:
            try:
                _ok, msg = self.controller.train_models()
                with self._lock:
                    self._training_message = msg
            except Exception as exc:
                with self._lock:
                    self._training_message = f"Huấn luyện thất bại: {exc}"
                LOG.exception("WEBVIEW_MODEL_TRAIN_FAILED")
            finally:
                with self._lock:
                    self._training = False

        self._training_thread = threading.Thread(target=work, daemon=True, name="webview-model-training")
        self._training_thread.start()
        return self.learning_state()

    def export_learning_data(self, target: str) -> dict:
        path = Path(target).expanduser()
        if path.suffix.lower() != ".zip":
            raise ServiceError("Tệp xuất dữ liệu học phải có phần mở rộng .zip.", "INVALID_EXPORT_PATH")
        with self._lock:
            self._assert_idle("Không thể xuất dữ liệu học khi đang xử lý báo cáo.")
            learning = self.controller.learning
            if learning is None:
                raise ServiceError("Không có dữ liệu học để xuất.", "EXPORT_FAILED")
            image_records = learning.store.records()
            content_records = learning.content.store.records()
            files = {
                "image_labels.jsonl": "".join(json.dumps(x, ensure_ascii=False) + "\n" for x in image_records),
                "content_labels.jsonl": "".join(json.dumps(x, ensure_ascii=False) + "\n" for x in content_records),
            }
            for name in ("image_model.json", "content_model.json"):
                model_path = learning.dir / name
                if model_path.is_file() and model_path.stat().st_size <= 16 * 1024 * 1024:
                    files[name] = model_path.read_text(encoding="utf-8")
            manifest = {"format": "tnp-learning-export-v1", "appVersion": __version__,
                        "imageLabelCount": len(image_records), "contentLabelCount": len(content_records),
                        "files": sorted(files)}
            files["manifest.json"] = json.dumps(manifest, ensure_ascii=False, indent=2)
            path = path.resolve()
            path.parent.mkdir(parents=True, exist_ok=True)
            temp_path = path.with_name(path.name + ".tmp")
            try:
                with zipfile.ZipFile(temp_path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
                    for name, content in files.items():
                        archive.writestr(name, content.encode("utf-8"))
                os.replace(temp_path, path)
            except (OSError, zipfile.BadZipFile, ValueError) as exc:
                try:
                    temp_path.unlink(missing_ok=True)
                except OSError:
                    pass
                raise ServiceError(f"Không xuất được dữ liệu học: {exc}", "EXPORT_FAILED") from exc
            message = (f"Đã xuất {len(image_records)} nhãn ảnh và {len(content_records)} nhãn nội dung "
                       f"ra {path.name}.")
            return {"message": message, "imageLabelCount": len(image_records),
                    "contentLabelCount": len(content_records)}

    # ------------------------------------------------------------------ output/log actions (paths come only from backend config)
    def open_output_file(self) -> dict:
        path = self.controller.output_file()
        return self._open_local_path(path, "OUTPUT_NOT_FOUND", "Chưa có file kết quả để mở.")

    def open_output_folder(self) -> dict:
        path = self.controller.output_folder()
        return self._open_local_path(path, "OUTPUT_NOT_FOUND", "Thư mục kết quả chưa tồn tại.")

    def open_log_file(self) -> dict:
        path = self.controller.log_file()
        return self._open_local_path(path, "LOG_NOT_FOUND", "Chưa có file log. Chạy quét hoặc xử lý trước.")

    def open_log_folder(self) -> dict:
        output = Path(self.controller.output).parent if self.controller.output else None
        path = output / "logs" if output and (output / "logs").is_dir() else logs_dir(create=False)
        return self._open_local_path(path, "LOG_NOT_FOUND", "Thư mục log chưa tồn tại.")

    def open_learning_folder(self) -> dict:
        return self._open_local_path(self.controller.learning_folder(), "LEARNING_NOT_FOUND",
                                     "Thư mục dữ liệu học chưa tồn tại.")

    def _open_local_path(self, path: Optional[Path], code: str, message: str) -> dict:
        if path is None or not Path(path).exists():
            raise ServiceError(message, code)
        target = Path(path).resolve()
        try:
            if sys.platform == "win32":
                os.startfile(str(target))  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                subprocess.Popen(["open", str(target)], close_fds=True)
            else:
                subprocess.Popen(["xdg-open", str(target)], close_fds=True)
        except Exception as exc:
            raise ServiceError(f"Không mở được đường dẫn đã chọn: {exc}", "OPEN_FAILED") from exc
        return {"message": f"Đã mở {target.name}."}

    # ------------------------------------------------------------------ close/log helpers
    def close_requested(self) -> tuple[bool, str]:
        with self._lock:
            if self.controller.is_running():
                self.controller.request_stop()
                return False, "Đã yêu cầu dừng sau báo cáo hiện tại. Hãy đợi xử lý kết thúc rồi đóng cửa sổ."
            if self._training or self._diagnostics_running or self._server_apply_running:
                return False, "Một tác vụ nền đang chạy. Hãy đợi tác vụ hoàn tất rồi đóng cửa sổ."
            return True, ""

    def _collect_logs(self) -> None:
        lines = self.controller.log_lines
        if self._log_cursor > len(lines):
            self._log_cursor = 0
            self._log_entries.clear()
        for line in lines[self._log_cursor:]:
            upper = str(line).upper()
            level = "ERROR" if "ERROR" in upper or "LỖI" in upper or "FAILED" in upper else (
                "WARN" if "WARN" in upper or "WARNING" in upper or "CẢNH BÁO" in upper else "INFO")
            self._log_sequence += 1
            self._log_entries.append({"id": self._log_sequence, "time": datetime.now().strftime("%H:%M:%S"),
                                     "level": level, "text": self._redact_source_paths(line)})
        self._log_cursor = len(lines)
        if len(self._log_entries) > 1000:
            del self._log_entries[:-1000]

    def _redact_source_paths(self, value: Any) -> str:
        """Keep user-selected PPTX source paths in Python; expose filenames, never absolute paths, to React."""
        text = str(value)
        paths = list(self._report_paths.values())
        paths.extend(Path(path) for path in getattr(self.controller, "all_files", []))
        replacements = set()
        for source in paths:
            try:
                path = Path(source)
                name = path.name
                if not name:
                    continue
                variants = {str(path), path.as_posix(), str(path.resolve()), path.resolve().as_posix()}
                variants.update(item.replace("/", chr(92)) for item in tuple(variants))
                for variant in variants:
                    if variant:
                        replacements.add((variant, name))
            except (OSError, RuntimeError, TypeError, ValueError):
                continue
        for source, name in sorted(replacements, key=lambda item: len(item[0]), reverse=True):
            text = re.sub(re.escape(source), lambda _match, replacement=name: replacement, text, flags=re.IGNORECASE)
        return text

    # ------------------------------------------------------------------ DTO helpers
    def _result_for_path(self, path: Path):
        c = self.controller
        processor = c.processor
        if not processor:
            return None
        normalized = normalize_source_path(path)
        return next((r for r in getattr(processor, "results", [])
                     if normalize_source_path(Path(r.source_file)) == normalized), None)

    @staticmethod
    def _duration_text(fr) -> str:
        if not fr.started_at or not fr.finished_at:
            return ""
        try:
            start = datetime.fromisoformat(fr.started_at)
            finish = datetime.fromisoformat(fr.finished_at)
            seconds = max(0, int((finish - start).total_seconds()))
            return f"{seconds // 60:02d}:{seconds % 60:02d}"
        except ValueError:
            return ""

    def _asset_data_uri(self, path: Path) -> str:
        try:
            path = path.resolve()
            outdir = Path(self.controller.output).parent.resolve() / "assets"
            if not path.is_relative_to(outdir) or not path.is_file() or path.stat().st_size > 40 * 1024 * 1024:
                return ""
            from PIL import Image
            with Image.open(path) as source:
                image = source.convert("RGB")
                image.thumbnail((1200, 1000))
                buf = io.BytesIO()
                image.save(buf, format="JPEG", quality=78, optimize=True)
            encoded = base64.b64encode(buf.getvalue()).decode("ascii")
            return "data:image/jpeg;base64," + encoded
        except Exception as exc:  # stale/corrupt image previews do not break the report detail pane
            LOG.debug("WEBVIEW_IMAGE_PREVIEW_UNAVAILABLE file=%s error=%s", path.name, exc)
            return ""

    # ------------------------------------------------------------------ PROMPT-024R learning workspace
    def _slide_preview_for(self, candidate) -> Dict[str, Any]:
        """FULL authored slide of one candidate as a data URI (review display only, §23/§29/§30).

        Rendered ONCE per (report, slide, file mtime) and shared by every candidate of that slide;
        two reports with identical layouts never share a preview (§54). No filesystem path ever reaches
        the browser (§39). Evidence bytes / Excel crops are untouched — this is review UI only.
        """
        empty = {"src": "", "width": 0, "height": 0}
        try:
            source = Path(candidate.source_file)
            if not source or not source.is_file():
                return empty
            key = (str(source), int(candidate.slide), int(source.stat().st_mtime))
            cached = self._slide_preview_cache.get(key)
            if cached is not None:
                return {"src": cached[0], "width": cached[1], "height": cached[2]}
            from .pptx_parser import parse_pptx
            from .qpn_renderer import SlideRenderer
            report = parse_pptx(source)
            if report.slide(int(candidate.slide)) is None:
                self._slide_preview_cache[key] = ("", 0, 0)
                return empty
            renderer = SlideRenderer(width_px=1600)   # best available backend, same chain as production
            import tempfile
            with tempfile.TemporaryDirectory(prefix="re_learning_slide_") as tmp:
                paths = renderer.render(report, [int(candidate.slide)], Path(tmp))
                rendered = paths.get(int(candidate.slide))
                if rendered is None or not Path(rendered).exists():
                    self._slide_preview_cache[key] = ("", 0, 0)
                    return empty
                from PIL import Image
                with Image.open(rendered) as im:
                    image = im.convert("RGB")
                    width, height = image.size
                    buf = io.BytesIO()
                    image.save(buf, format="JPEG", quality=82, optimize=True)
            uri = "data:image/jpeg;base64," + base64.b64encode(buf.getvalue()).decode("ascii")
            self._slide_preview_cache[key] = (uri, width, height)
            if len(self._slide_preview_cache) > 48:
                self._slide_preview_cache.pop(next(iter(self._slide_preview_cache)))
            return {"src": uri, "width": width, "height": height}
        except Exception as exc:  # noqa: BLE001 – preview is display-only; never break the review state
            LOG.debug("WEBVIEW_SLIDE_PREVIEW_UNAVAILABLE candidate=%s error=%s",
                      getattr(candidate, "candidate_id", "?"), exc)
            return empty

    def _image_candidate_dto(self, candidate, learning) -> dict:
        src = self._preview_cache.get(candidate.candidate_id)
        if src is None:
            blob = self.controller.candidate_blob(candidate) or b""
            src = self._bytes_data_uri(blob)
            self._preview_cache[candidate.candidate_id] = src
            if len(self._preview_cache) > 128:
                self._preview_cache.pop(next(iter(self._preview_cache)))
        bw = max(1, int(candidate.slide_size[0] or 1))
        bh = max(1, int(candidate.slide_size[1] or 1))
        x, y, w, h = candidate.bounds
        label = self.controller.pending_labels.get(candidate.candidate_id) or learning.store.current_label(candidate.candidate_id)
        latest = learning.store.latest().get(candidate.candidate_id, {})
        pending_note = self.controller.pending_label_notes.get(candidate.candidate_id)
        # PROMPT-024R: full-slide context + authoritative target geometry for the review workspace.
        fx, fy, fw, fh = slide_fraction_box((x, y, w, h), bw, bh)
        preview = self._slide_preview_for(candidate)
        return {"id": candidate.candidate_id, "sourceFile": candidate.source_name,
                "managementNumber": candidate.management_number, "slide": candidate.slide,
                "pictureId": str(candidate.picture_id), "src": src,
                "bounds": {"x": round(x / bw * 100, 2), "y": round(y / bh * 100, 2),
                           "w": round(w / bw * 100, 2), "h": round(h / bh * 100, 2)},
                # Authoritative PPTX geometry (EMU + fractions) so the UI never guesses coordinates (§27/§38)
                "slideWidth": bw, "slideHeight": bh,
                "targetBbox": {"x": int(x), "y": int(y), "width": int(w), "height": int(h)},
                "targetBboxPct": {"x": round(fx * 100, 3), "y": round(fy * 100, 3),
                                  "w": round(fw * 100, 3), "h": round(fh * 100, 3)},
                "targetKind": "picture",
                # Full authored slide rendered once per report+slide (review display only, §40)
                "slidePreview": preview["src"],
                "slidePreviewWidth": preview["width"], "slidePreviewHeight": preview["height"],
                # PROMPT-025: logical improvement-item identity for the per-item learning preview
                "itemId": candidate.logical_item_owner, "itemIndex": candidate.item_index,
                "itemHeading": candidate.owner_heading,
                "decision": candidate.decision, "confidenceBand": candidate.confidence_band,
                "confidence": float(candidate.confidence), "evidence": list(candidate.evidence),
                "excelEligible": bool(candidate.excel_output_eligible),
                "eligibilityReason": candidate.eligibility_reason, "nearbyText": candidate.nearby_text,
                "userLabel": label, "labelPending": candidate.candidate_id in self.controller.pending_labels,
                "note": str(pending_note if pending_note is not None else latest.get("note", ""))}

    def _content_candidate_dto(self, candidate, learning) -> dict:
        label = self.controller.pending_content_labels.get(candidate.candidate_id) or \
            learning.content.store.current_label(candidate.candidate_id)
        latest = learning.content.store.latest().get(candidate.candidate_id, {})
        pending_note = self.controller.pending_content_notes.get(candidate.candidate_id)
        x = {"id": candidate.candidate_id, "sourceFile": candidate.source_name,
             "managementNumber": candidate.management_number, "slide": candidate.slide,
             "blockId": str(candidate.shape_id), "text": candidate.text, "decision": candidate.decision,
             "confidenceBand": candidate.confidence_band, "confidence": float(candidate.confidence),
             "evidence": list(candidate.evidence), "section": candidate.section_kind or candidate.role,
             "nearestTitle": candidate.nearest_title, "userLabel": label,
             "labelPending": candidate.candidate_id in self.controller.pending_content_labels,
             "note": str(pending_note if pending_note is not None else latest.get("note", ""))}
        return x

    @staticmethod
    def _bytes_data_uri(blob: bytes) -> str:
        if not blob or len(blob) > 32 * 1024 * 1024:
            return ""
        try:
            from PIL import Image
            with Image.open(io.BytesIO(blob)) as source:
                image = source.convert("RGB")
                image.thumbnail((1200, 1000))
                buf = io.BytesIO()
                image.save(buf, format="JPEG", quality=78, optimize=True)
            return "data:image/jpeg;base64," + base64.b64encode(buf.getvalue()).decode("ascii")
        except Exception:
            return ""

    @staticmethod
    def _model_status(model, state: str, kind: str) -> str:
        if model is not None:
            return f"Đã huấn luyện ({model.n_examples} mẫu)."
        if state in ("corrupt", "incompatible"):
            return f"Mô hình {kind} không khả dụng ({state}); backend tiếp tục dùng quy tắc hiện tại."
        return f"Chưa huấn luyện mô hình {kind}; backend tiếp tục dùng quy tắc hiện tại."

    def _update_dto(self) -> dict:
        c = self.controller
        check = c.update_check
        info = check.info if check else None
        status = "checking" if c.update_busy else (
            "available" if check and check.available else "latest" if check and check.status == "latest" else
            "error" if check and check.status not in ("no_path", "older") else "idle")
        return {"path": c.update_path, "autoCheck": bool(c.auto_update_check), "status": status,
                "message": c.update_status_text(),
                "available": ({"version": info.version, "build": f"{info.build:03d}",
                               "package": info.package, "size": ""} if info else None),
                "progress": float(c.update_progress.percent), "progressStage": c.update_progress.label,
                "installEnabled": False}

    @staticmethod
    def _discovery_dto(result) -> dict:
        return {"host": result.host, "port": result.port, "models": len(result.models),
                "version": result.version or "—", "note": result.note or ""}

    # ------------------------------------------------------------------ manual file store and input validation
    def _load_manual_fields(self) -> None:
        try:
            data = json.loads(self._manual_fields_path.read_text(encoding="utf-8"))
            if isinstance(data, dict):
                overrides = data.get("overrides") if isinstance(data.get("overrides"), dict) else data
                pending = data.get("pending", []) if "overrides" in data else []
                for path_key, values in overrides.items():
                    if not isinstance(path_key, str) or not isinstance(values, dict):
                        continue
                    clean = {}
                    if "vendor" in values and isinstance(values["vendor"], str):
                        clean["vendor"] = self._validate_vendor(values["vendor"])
                    if "occurrence_date" in values and isinstance(values["occurrence_date"], str):
                        clean["occurrence_date"] = self._validate_iso_date(values["occurrence_date"])
                    if clean:
                        self._manual_fields[path_key] = clean
                if isinstance(pending, list):
                    self._pending_manual = {k: dict(self._manual_fields[k]) for k in pending
                                            if isinstance(k, str) and k in self._manual_fields}
            self.controller.manual_fields_by_path = {k: dict(v) for k, v in self._manual_fields.items()}
        except FileNotFoundError:
            self.controller.manual_fields_by_path = {}
        except Exception as exc:
            LOG.warning("WEBVIEW_MANUAL_FIELDS_LOAD_FAILED error=%s", exc)
            self.controller.manual_fields_by_path = {}

    def _save_manual_fields(self) -> None:
        self._manual_fields_path.parent.mkdir(parents=True, exist_ok=True)
        tmp = self._manual_fields_path.with_suffix(".tmp")
        payload = {"version": 1, "overrides": self._manual_fields, "pending": sorted(self._pending_manual)}
        tmp.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
        os.replace(tmp, self._manual_fields_path)

    def _validate_vendor(self, value: Any) -> str:
        text = self._text(value, "vendor", 500).replace("\r\n", "\n").replace("\r", "\n")
        if any(ord(ch) < 32 and ch not in "\n\t" for ch in text):
            raise ServiceError("Vendor có ký tự điều khiển không hợp lệ.", "INVALID_VENDOR")
        lines = [line.strip() for line in text.splitlines()]
        if len(lines) > 20 or any(len(line) > 120 for line in lines):
            raise ServiceError("Vendor tối đa 20 dòng, mỗi dòng không quá 120 ký tự.", "INVALID_VENDOR")
        return "\n".join(lines).strip()

    @classmethod
    def _validate_date_to_iso(cls, value: Any) -> str:
        text = cls._text(value, "occurrenceDate", 32).strip()
        if not text:
            return ""
        try:
            parsed = datetime.strptime(text, "%d/%m/%Y").date()
        except ValueError as exc:
            raise ServiceError("Ngày phát sinh phải có dạng dd/mm/yyyy và là ngày hợp lệ.", "INVALID_DATE") from exc
        if not 1900 <= parsed.year <= 2199:
            raise ServiceError("Năm Ngày phát sinh phải nằm trong khoảng 1900–2199.", "INVALID_DATE")
        return parsed.isoformat()

    @staticmethod
    def _validate_iso_date(value: str) -> str:
        if not value:
            return ""
        try:
            parsed = date.fromisoformat(value)
        except ValueError as exc:
            raise ValueError("Ngày ISO không hợp lệ") from exc
        if not 1900 <= parsed.year <= 2199:
            raise ValueError("Năm ngày ISO không hợp lệ")
        return parsed.isoformat()

    @staticmethod
    def _iso_to_dmy(value: str) -> str:
        if not value:
            return ""
        try:
            return date.fromisoformat(value).strftime("%d/%m/%Y")
        except ValueError:
            return ""

    # ------------------------------------------------------------------ common validation
    @staticmethod
    def _object(value: Any, name: str) -> dict:
        if not isinstance(value, dict):
            raise ServiceError(f"{name} phải là một đối tượng JSON.", "INVALID_DTO")
        if len(value) > 64:
            raise ServiceError(f"{name} có quá nhiều trường.", "INVALID_DTO")
        return value

    @staticmethod
    def _text(value: Any, name: str, max_length: int) -> str:
        if value is None:
            return ""
        if not isinstance(value, (str, int, float)):
            raise ServiceError(f"{name} phải là chuỗi.", "INVALID_DTO")
        text = str(value).strip()
        if len(text) > max_length:
            raise ServiceError(f"{name} quá dài.", "INVALID_DTO")
        return text

    @staticmethod
    def _boolean(value: Any, name: str) -> bool:
        if not isinstance(value, bool):
            raise ServiceError(f"{name} phải là giá trị true/false.", "INVALID_DTO")
        return value

    @classmethod
    def _choice(cls, value: Any, name: str, allowed: set) -> str:
        text = cls._text(value, name, 128)
        if text not in allowed:
            raise ServiceError(f"{name} không hợp lệ.", "INVALID_DTO")
        return text

    def _validated_report_ids(self, report_ids: Iterable[str]) -> List[str]:
        if not isinstance(report_ids, (list, tuple)) or not report_ids or len(report_ids) > 1000:
            raise ServiceError("Chọn ít nhất một báo cáo hợp lệ.", "INVALID_REPORT_IDS")
        ids = [self._text(x, "reportId", 24) for x in report_ids]
        if any(not x.isdigit() or x not in self._report_paths for x in ids):
            raise ServiceError("Danh sách báo cáo không hợp lệ hoặc đã hết hạn.", "INVALID_REPORT_IDS")
        return ids

    def _assert_idle(self, message: str) -> None:
        if self.controller.is_running() or self._training or self._diagnostics_running or self._server_apply_running:
            raise ServiceError(message, "BUSY")
