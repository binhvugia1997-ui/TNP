"""PROMPT-037 backend guards.

The hide rule lives in the frontend presentation boundary. These tests pin the backend side of that
contract: the skip classification and the DTO status stay exactly as before (processing, duplicate
handling and incremental transitions depend on them), and building the DTO never touches source files.
"""
from __future__ import annotations

from pathlib import Path
from types import SimpleNamespace

from app import prescan
from app.application_service import ApplicationService
from app.prescan import (ACTION_FAST_SKIP, ACTION_MASTER_COMPLETE, ACTION_OUTSIDE_PERIOD, ACTION_PROCESS,
                         ACTION_PROCESS_NEW_ROW, ACTION_SOURCE_DUPLICATE)


class _IdentityRedactor:
    def redact(self, value):
        return "" if value is None else str(value)


def _service() -> ApplicationService:
    service = ApplicationService.__new__(ApplicationService)
    service._pending_manual = set()
    service.controller = SimpleNamespace(output="")   # only read for rows with an Excel row
    return service


def _row(path: Path, action: str, *, excluded: bool = False, run_stage: str = "", excel_row=None):
    return SimpleNamespace(action=action, excluded=excluded, run_stage=run_stage, reason="Bỏ qua — đã cập nhật (dòng 3)",
                           note="", occurrence_date=None, vendor="", management_number="2610101-VOC", path=path,
                           excel_row=excel_row, normalized_path=str(path).lower(), scan_status_vi="")


def test_skip_classification_is_unchanged_in_the_dto_mapping():
    # Pre-scan "Excel row already complete" is still a skipped classification (prescan pipeline key).
    assert prescan.ACTION_STATUS[ACTION_MASTER_COMPLETE] == "skipped"
    assert ApplicationService._status_key(ACTION_MASTER_COMPLETE, False, "") == "skipped"
    # Batch-time final skip keeps its status key.
    assert ApplicationService._status_key(ACTION_PROCESS, False, "skipped") == "skipped"
    # The other skip kinds keep their own statuses (the frontend hides only "skipped").
    assert ApplicationService._status_key(ACTION_OUTSIDE_PERIOD, False, "") == "outside_period"
    assert ApplicationService._status_key(ACTION_SOURCE_DUPLICATE, False, "") == "source_duplicate"
    assert ApplicationService._status_key(ACTION_FAST_SKIP, False, "") == "fast_skip"
    # Precedence unchanged: manual exclusion wins; processable rows remain waiting / new_row.
    assert ApplicationService._status_key(ACTION_MASTER_COMPLETE, True, "") == "excluded"
    assert ApplicationService._status_key(ACTION_PROCESS, False, "") == "waiting"
    assert ApplicationService._status_key(ACTION_PROCESS_NEW_ROW, False, "") == "new_row"


def test_dto_keeps_skipped_rows_with_status_so_presentation_can_reconcile_transitions(tmp_path):
    service = _service()
    source = tmp_path / "2610101-VOC.pptx"
    source.write_bytes(b"pptx-bytes")
    dto = service._report_dto("7", _row(source, ACTION_MASTER_COMPLETE, excel_row=3), None, {}, 1,
                              redactor=_IdentityRedactor())
    assert dto["id"] == "7"
    assert dto["status"] == "skipped"
    assert dto["fileName"] == "2610101-VOC.pptx"
    # The skipped DTO row is built without touching the source file.
    assert source.read_bytes() == b"pptx-bytes"


def test_building_skipped_dto_rows_never_deletes_source_files(tmp_path):
    service = _service()
    files = []
    for index in range(25):
        source = tmp_path / f"report-{index}.pptx"
        source.write_bytes(b"x")
        files.append(source)
    rows = [_row(path, ACTION_MASTER_COMPLETE if index % 2 else ACTION_PROCESS)
            for index, path in enumerate(files)]
    dtos = [service._report_dto(str(index), row, None, {}, index + 1, redactor=_IdentityRedactor())
            for index, row in enumerate(rows)]
    assert sum(1 for dto in dtos if dto["status"] == "skipped") == 12
    assert all(path.exists() for path in files)
