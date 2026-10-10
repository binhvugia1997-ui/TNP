"""Validated pywebview API surface for the local React desktop UI.

Only explicit native file/folder selections and bounded application operations are exposed. There is deliberately no
method that accepts an arbitrary path for reading, writing, listing, or opening.
"""
from __future__ import annotations

import json
import logging
import time
from pathlib import Path
from typing import Any, Callable, Optional

from .application_service import ApplicationService, ServiceError

LOG = logging.getLogger("report_extractor.webview.bridge")
MAX_DTO_BYTES = 1_000_000
PPTX_FILTER = "PowerPoint presentation (*.pptx)"
XLSX_FILTER = "Excel workbook (*.xlsx)"
ZIP_FILTER = "ZIP archive (*.zip)"


class BridgeService:
    """pywebview JavaScript API with one stable success/error envelope."""

    def __init__(self, service: Optional[ApplicationService] = None):
        self.service = service or ApplicationService()
        self.window = None

    def bind_window(self, window) -> None:
        self.window = window

    # -------------------------------------------------------------- required bridge contract
    def get_app_version(self) -> dict:
        return self._respond(self.service.app_version)

    def choose_pptx_file(self) -> dict:
        return self._respond(self._choose_pptx_file)

    def choose_report_folder(self) -> dict:
        return self._respond(self._choose_report_folder)

    def get_current_config(self) -> dict:
        return self._respond(self.service.current_config)

    # -------------------------------------------------------------- dashboard / report operations
    def get_dashboard_state(self) -> dict:
        return self._respond(self.service.dashboard_state)

    def save_configuration(self, dto: Any) -> dict:
        return self._respond(lambda: self.service.save_configuration(self._bounded_object(dto, "configuration")))

    def scan_reports(self, dto: Any, input_file_token: str = "") -> dict:
        started = time.perf_counter()
        result = self._respond(lambda: self.service.scan_reports(self._bounded_object(dto, "configuration"),
                                                                 self._string(input_file_token, "inputFileToken", 128)))
        elapsed_ms = (time.perf_counter() - started) * 1000.0
        # pywebview serializes the returned envelope only after this Python method returns.  Mark that
        # boundary honestly; timing it here would measure a second json.dumps, not the real transport.
        LOG.info("SCAN_SERIALIZE elapsed_ms=unavailable boundary=pywebview_after_python_return")
        LOG.info("SCAN_RETURN elapsed_ms=%.3f ok=%s bridge_calls=1", elapsed_ms, bool(result.get("ok")))
        return result

    def get_report_details(self, report_id: str) -> dict:
        return self._respond(lambda: self.service.report_details(self._string(report_id, "reportId", 24)))

    def start_processing(self, dto: Any, use_ollama: bool = True) -> dict:
        def call():
            return self.service.start_processing(self._bounded_object(dto, "configuration"),
                                                 self._boolean(use_ollama, "useOllama"))
        return self._respond(call)

    def stop_after_current(self) -> dict:
        return self._respond(self.service.stop_after_current)

    def cancel_all(self) -> dict:
        """PROMPT-024R: cooperative "Dừng tất cả" (never kills threads/processes)."""
        return self._respond(self.service.cancel_all)

    def exclude_reports(self, report_ids: Any) -> dict:
        return self._respond(lambda: self.service.exclude_reports(self._string_list(report_ids, "reportIds"), False))

    def restore_reports(self, report_ids: Any) -> dict:
        return self._respond(lambda: self.service.exclude_reports(self._string_list(report_ids, "reportIds"), True))

    def save_manual_fields(self, report_id: str, fields: Any) -> dict:
        return self._respond(lambda: self.service.save_manual_fields(
            self._string(report_id, "reportId", 24), self._bounded_object(fields, "manualFields")))

    def retry_manual_fields(self, report_id: str) -> dict:
        return self._respond(lambda: self.service.retry_manual_fields(self._string(report_id, "reportId", 24)))

    # -------------------------------------------------------------- native dialogs for other configured paths
    def choose_template_file(self) -> dict:
        return self._respond(lambda: self._choose_file(XLSX_FILTER, ".xlsx", "template"))

    def choose_output_file(self) -> dict:
        return self._respond(lambda: self._choose_save_file("Kiem_chung_Ket_qua.xlsx", XLSX_FILTER, ".xlsx"))

    def choose_update_folder(self) -> dict:
        return self._respond(self._choose_report_folder)

    def choose_learning_export(self) -> dict:
        return self._respond(lambda: self._choose_save_file("learning_data_export.zip", ZIP_FILTER, ".zip"))

    # -------------------------------------------------------------- settings / Ollama / updates / diagnostics
    def check_ollama_connection(self, dto: Any) -> dict:
        def call():
            self.service.save_configuration(self._bounded_object(dto, "configuration"))
            return self.service.check_ollama_connection()
        return self._respond(call)

    def refresh_ollama_models(self, dto: Any) -> dict:
        def call():
            self.service.save_configuration(self._bounded_object(dto, "configuration"))
            return self.service.refresh_ollama_models()
        return self._respond(call)

    def start_ollama_discovery(self) -> dict:
        return self._respond(self.service.start_ollama_discovery)

    def stop_ollama_discovery(self) -> dict:
        return self._respond(self.service.stop_ollama_discovery)

    def use_discovered_server(self, host: str, port: int) -> dict:
        return self._respond(lambda: self.service.use_discovered_server(
            self._string(host, "host", 512), self._integer(port, "port", 1, 65535)))

    def check_update(self, dto: Any) -> dict:
        def call():
            self.service.save_configuration(self._bounded_object(dto, "configuration"))
            return self.service.check_update()
        return self._respond(call)

    def install_update(self) -> dict:
        return self._respond(self.service.install_update)

    def run_diagnostics(self) -> dict:
        return self._respond(self.service.run_diagnostics)

    # -------------------------------------------------------------- learning operations
    def get_learning_state(self) -> dict:
        return self._respond(self.service.learning_state)

    def set_learning_label(self, kind: str, candidate_id: str, label: str, note: str = "") -> dict:
        return self._respond(lambda: self.service.set_learning_label(
            self._string(kind, "kind", 16), self._string(candidate_id, "candidateId", 512),
            self._string(label, "label", 64), self._string(note, "note", 1000)))

    def save_learning_labels(self, kind: str, notes: Any = None) -> dict:
        if notes is None:
            notes = {}
        return self._respond(lambda: self.service.save_learning_labels(
            self._string(kind, "kind", 16), self._bounded_object(notes, "learningNotes")))

    def retry_learning_excel(self) -> dict:
        return self._respond(self.service.retry_learning_excel)

    def dismiss_learning_excel_notice(self) -> dict:
        return self._respond(self.service.dismiss_learning_excel_notice)

    def train_models(self) -> dict:
        return self._respond(self.service.train_models)

    def export_learning_data(self) -> dict:
        def call():
            selected = self._choose_save_file("learning_data_export.zip", ZIP_FILTER, ".zip")
            if selected["cancelled"]:
                return selected
            return self.service.export_learning_data(selected["path"])
        return self._respond(call)

    # -------------------------------------------------------------- fixed backend-owned output actions
    def open_output_file(self) -> dict:
        return self._respond(self.service.open_output_file)

    def open_output_folder(self) -> dict:
        return self._respond(self.service.open_output_folder)

    def open_log_file(self) -> dict:
        return self._respond(self.service.open_log_file)

    def open_log_folder(self) -> dict:
        return self._respond(self.service.open_log_folder)

    def open_learning_folder(self) -> dict:
        return self._respond(self.service.open_learning_folder)

    def ping(self) -> dict:
        """Minimal health-check also useful to verify a newly created native window."""
        return self._respond(lambda: "Python bridge OK")

    # -------------------------------------------------------------- dialog internals
    def _choose_pptx_file(self) -> dict:
        path = self._dialog_path(self._webview_constant("OPEN_DIALOG"), file_types=(PPTX_FILTER,))
        if not path:
            return {"cancelled": True}
        selected = self.service.register_pptx_selection(path)
        return {"cancelled": False, **selected}

    def _choose_report_folder(self) -> dict:
        path = self._dialog_path(self._webview_constant("FOLDER_DIALOG"))
        if not path:
            return {"cancelled": True}
        if not Path(path).is_dir():
            raise ServiceError("Đường dẫn đã chọn không phải thư mục hiện có.", "INVALID_SELECTION")
        return {"cancelled": False, "path": str(Path(path).resolve())}

    def _choose_file(self, file_type: str, expected_suffix: str, name: str) -> dict:
        path = self._dialog_path(self._webview_constant("OPEN_DIALOG"), file_types=(file_type,))
        if not path:
            return {"cancelled": True}
        selected = Path(path)
        if selected.suffix.lower() != expected_suffix.lower() or not selected.is_file():
            raise ServiceError(f"Chọn một file {expected_suffix} hiện có.", "INVALID_SELECTION")
        return {"cancelled": False, "path": str(selected), "kind": name}

    def _choose_save_file(self, filename: str, file_type: str, expected_suffix: str) -> dict:
        path = self._dialog_path(self._webview_constant("SAVE_DIALOG"), save_filename=filename,
                                 file_types=(file_type,))
        if not path:
            return {"cancelled": True}
        selected = Path(path)
        if selected.suffix.lower() != expected_suffix.lower() or selected.is_dir():
            raise ServiceError(f"Tên tệp phải có phần mở rộng {expected_suffix}.", "INVALID_SELECTION")
        return {"cancelled": False, "path": str(selected)}

    def _dialog_path(self, dialog_type: Any, **kwargs) -> str:
        if self.window is None:
            raise ServiceError("Cửa sổ native chưa sẵn sàng để mở hộp thoại.", "DIALOG_UNAVAILABLE")
        result = self.window.create_file_dialog(dialog_type, **kwargs)
        if not result:
            return ""
        if isinstance(result, (tuple, list)):
            if len(result) != 1:
                raise ServiceError("Chỉ chọn một đường dẫn mỗi lần.", "INVALID_SELECTION")
            result = result[0]
        if not isinstance(result, (str, Path)):
            raise ServiceError("Hộp thoại trả về lựa chọn không hợp lệ.", "INVALID_SELECTION")
        try:
            return str(Path(result).expanduser().resolve())
        except (OSError, RuntimeError, ValueError) as exc:
            raise ServiceError("Hộp thoại trả về đường dẫn không hợp lệ.", "INVALID_SELECTION") from exc

    @staticmethod
    def _webview_constant(name: str):
        try:
            import webview
        except ImportError as exc:
            raise ServiceError("pywebview chưa được cài đặt. Hãy cài requirements-webview.txt.",
                               "WEBVIEW_NOT_INSTALLED") from exc
        return getattr(webview, name)

    # -------------------------------------------------------------- DTO/error contract
    def _respond(self, action: Callable[[], Any]) -> dict:
        try:
            return {"ok": True, "data": action()}
        except ServiceError as exc:
            return {"ok": False, "error": {"code": exc.code,
                                            "message": self.service._redact_source_paths(str(exc))}}
        except (TypeError, ValueError) as exc:
            return {"ok": False, "error": {"code": "INVALID_REQUEST",
                                            "message": self.service._redact_source_paths(str(exc))}}
        except Exception:  # do not leak tracebacks or raw filesystem paths to JavaScript
            LOG.exception("WEBVIEW_BRIDGE_UNEXPECTED_ERROR")
            return {"ok": False, "error": {"code": "INTERNAL_ERROR",
                                            "message": "Đã xảy ra lỗi nội bộ. Hãy kiểm tra nhật ký ứng dụng."}}

    @staticmethod
    def _bounded_object(value: Any, name: str) -> dict:
        if not isinstance(value, dict):
            raise ServiceError(f"{name} phải là một đối tượng JSON.", "INVALID_DTO")
        try:
            size = len(json.dumps(value, ensure_ascii=False).encode("utf-8"))
        except (TypeError, ValueError) as exc:
            raise ServiceError(f"{name} không phải JSON hợp lệ.", "INVALID_DTO") from exc
        if size > MAX_DTO_BYTES:
            raise ServiceError(f"{name} vượt giới hạn {MAX_DTO_BYTES} byte.", "INVALID_DTO")
        return value

    @staticmethod
    def _string(value: Any, name: str, max_length: int) -> str:
        if not isinstance(value, str):
            raise ServiceError(f"{name} phải là chuỗi.", "INVALID_DTO")
        text = value.strip()
        if len(text) > max_length:
            raise ServiceError(f"{name} quá dài.", "INVALID_DTO")
        return text

    @staticmethod
    def _integer(value: Any, name: str, minimum: int, maximum: int) -> int:
        if isinstance(value, bool) or not isinstance(value, int) or not minimum <= value <= maximum:
            raise ServiceError(f"{name} phải là số nguyên từ {minimum} đến {maximum}.", "INVALID_DTO")
        return value

    @staticmethod
    def _boolean(value: Any, name: str) -> bool:
        if not isinstance(value, bool):
            raise ServiceError(f"{name} phải là giá trị true/false.", "INVALID_DTO")
        return value

    @classmethod
    def _string_list(cls, value: Any, name: str) -> list[str]:
        if not isinstance(value, list) or not value or len(value) > 1000:
            raise ServiceError(f"{name} phải chứa 1–1000 mã báo cáo.", "INVALID_DTO")
        return [cls._string(x, name, 24) for x in value]
