"""PROMPT-031 regression: bounded Quét work, termination diagnostics, and list semantics.

The synthetic ``.pptx`` files are intentionally not ZIP files.  The list-discovery contract
must use directory/name/stat metadata only; opening one as a presentation would fail this suite.
No wall-clock threshold is a correctness gate.  Work counts are the stable performance gate.
"""
from __future__ import annotations

import logging
import time
from datetime import date
from pathlib import Path

import pytest

import app.batch_processor as batch_processor
import app.qpn_renderer as qpn_renderer
from app.application_service import ApplicationService
from app.bridge import BridgeService
from app.config import AppConfig
from app.gui_controller import GuiController
from app.prescan import ACTION_OUTSIDE_PERIOD, ACTION_PROCESS, IntrinsicScanCache, month_period, prescan
from app.scan_diagnostics import ScanMetrics

INSIDE = 133
OUTSIDE = 67
UNRELATED = 400
STRUCTURAL_REJECTS = 10
POWERPOINT_CANDIDATES = INSIDE + OUTSIDE
DIRECTORY_ENTRIES = POWERPOINT_CANDIDATES + UNRELATED + STRUCTURAL_REJECTS


def _dto(folder: Path) -> dict:
    return {
        "paths": {"reportFolder": str(folder), "template": "", "output": ""},
        "period": {"mode": "month", "month": "9", "year": "2026", "from": "", "to": ""},
        "forceReprocess": False,
        "ollama": {"host": "127.0.0.1", "port": "11434", "model": "qwen3:4b"},
        "update": {"path": "", "autoCheck": False},
    }


def _service(folder: Path, tmp_path: Path) -> ApplicationService:
    config_path = tmp_path / "config" / "settings.json"
    controller = GuiController(AppConfig(period_mode="month", period_month=9, period_year=2026),
                               config_path=config_path)
    return ApplicationService(controller, config_path=config_path,
                              manual_fields_path=tmp_path / "config" / "manual_fields.json")


@pytest.fixture
def realistic_scan_folder(tmp_path):
    """610 deterministic entries: list survivors, month skips, cheap rejects, corrupt bytes."""
    folder = tmp_path / "reports"
    folder.mkdir()
    inside_names = []
    outside_names = []
    for i in range(INSIDE):
        name = f"260901{i + 1:03d}_eligible_{i:04d}.pptx"
        (folder / name).write_bytes(b"not-a-zip-presentation")
        inside_names.append(name)
    for i in range(OUTSIDE):
        name = f"261001{i + 1:03d}_outside_{i:04d}.pptx"
        (folder / name).write_bytes(b"also-not-a-zip")
        outside_names.append(name)
    for i in range(UNRELATED):
        (folder / f"unrelated_{i:04d}.txt").write_text("not a report", encoding="utf-8")
    for i in range(4):
        (folder / f"~$2609019{i:02d}_lock.pptx").write_bytes(b"temp")
    for i in range(3):
        (folder / f".hidden_{i}.pptx").write_bytes(b"hidden")
    for i in range(3):
        (folder / f"deck_{i}.pptx.tmp").write_bytes(b"temp")
    assert len(list(folder.iterdir())) == DIRECTORY_ENTRIES
    return folder, inside_names, outside_names


def _snapshot(envelope: dict) -> list:
    reports = envelope["data"]["reports"]
    return [(row["id"], row["stt"], row["fileName"], row["managementNumber"],
             row["occurrenceDate"], row["status"], row["path"]) for row in reports]


def test_large_quiet_scan_is_one_bounded_batch_with_identical_rescan(realistic_scan_folder, tmp_path,
                                                                       monkeypatch, caplog):
    folder, inside_names, outside_names = realistic_scan_folder
    service = _service(folder, tmp_path)
    bridge = BridgeService(service)

    parse_calls = []
    com_calls = []
    monkeypatch.setattr(batch_processor, "parse_pptx",
                        lambda path: parse_calls.append(path) or pytest.fail("Quét parsed a PPTX"))
    monkeypatch.setattr(qpn_renderer, "render_with_powerpoint",
                        lambda *args, **kwargs: com_calls.append(args) or pytest.fail("Quét started PowerPoint COM"))

    with caplog.at_level(logging.INFO, logger="report_extractor"):
        first = bridge.scan_reports(_dto(folder))
    assert first["ok"] is True
    first_snapshot = _snapshot(first)
    metrics = service._last_scan_metrics
    assert metrics is not None

    # Same list contract as before the optimization: every candidate remains visible,
    # processing rows first, outside-month rows afterwards, stable source order in each group.
    assert len(first_snapshot) == POWERPOINT_CANDIDATES
    assert [row[2] for row in first_snapshot[:INSIDE]] == inside_names
    assert [row[2] for row in first_snapshot[INSIDE:]] == outside_names
    assert {row[5] for row in first_snapshot[:INSIDE]} == {"waiting"}
    assert {row[5] for row in first_snapshot[INSIDE:]} == {"outside_period"}
    assert all(row[4] == "01/09/2026" for row in first_snapshot[:INSIDE])
    assert all(row[4] == "01/10/2026" for row in first_snapshot[INSIDE:])
    assert all(row[6] == "" for row in first_snapshot)  # minimal list DTO still hides source paths

    # Stable algorithmic gates.  Unrelated and structurally ineligible files are rejected
    # before candidate stat/status work; each candidate has exactly one metadata inspection.
    assert metrics.entries_inspected == DIRECTORY_ENTRIES
    assert metrics.candidate_files == POWERPOINT_CANDIDATES
    assert metrics.structural_rejected_files == STRUCTURAL_REJECTS
    assert metrics.unrelated_files == UNRELATED
    assert metrics.accepted_files == INSIDE
    assert metrics.status_skipped_files == OUTSIDE
    assert metrics.directory_traversals == 1
    assert metrics.metadata_reads == POWERPOINT_CANDIDATES
    assert metrics.metadata_errors == 0
    assert metrics.redaction_context_builds == 1
    assert metrics.redaction_sources == POWERPOINT_CANDIDATES
    assert metrics.intrinsic_cache_hits == 0
    assert metrics.intrinsic_cache_misses == POWERPOINT_CANDIDATES
    assert metrics.intrinsic_derivations == POWERPOINT_CANDIDATES
    assert metrics.leaf_folders == metrics.leaf_deliveries == 1
    assert metrics.leaf_rows_delivered == POWERPOINT_CANDIDATES
    assert metrics.pptx_opens == 0 and metrics.powerpoint_com_starts == 0
    assert parse_calls == [] and com_calls == []

    messages = [record.getMessage() for record in caplog.records]
    for marker in ("SCAN_START", "SCAN_ENUMERATE_START", "SCAN_ENUMERATE_END",
                   "SCAN_FILTER_START", "SCAN_FILTER_END", "SCAN_BUILD_RESPONSE_START",
                   "SCAN_BUILD_RESPONSE_END", "SCAN_PER_FILE", "SCAN_METADATA",
                   "SCAN_CACHE_SUMMARY", "SCAN_LEAF_START", "SCAN_LEAF_END", "SCAN_LEAF_DELIVER",
                   "SCAN_INCREMENTAL_SUMMARY", "SCAN_PARSE_PPTX", "SCAN_POWERPOINT_COM", "SCAN_SERIALIZE",
                   "SCAN_RETURN", "SCAN_TOTAL"):
        assert any(message.startswith(marker) for message in messages), marker
    assert any("entries=610" in message and "candidates=200" in message and
               "accepted=133" in message and "traversals=1" in message
               for message in messages if message.startswith("SCAN_TOTAL"))

    # A rescan of unchanged inputs is deterministic and still performs one traversal, not
    # a per-file bridge call or a duplicate walk hidden behind the first result.
    second = bridge.scan_reports(_dto(folder))
    assert second["ok"] is True and _snapshot(second) == first_snapshot
    assert service._last_scan_metrics.directory_traversals == 1
    assert service._last_scan_metrics.metadata_reads == POWERPOINT_CANDIDATES
    assert service._last_scan_metrics.redaction_context_builds == 1
    assert service._last_scan_metrics.intrinsic_cache_hits == POWERPOINT_CANDIDATES
    assert service._last_scan_metrics.intrinsic_cache_misses == 0
    assert service._last_scan_metrics.intrinsic_derivations == 0


def test_redaction_context_is_lexical_and_reused_without_source_io(tmp_path, monkeypatch):
    folder = tmp_path / "reports"
    folder.mkdir()
    sources = [folder / "260901001 alpha.pptx", folder / "260901002 beta.pptx"]
    for source in sources:
        source.write_bytes(b"x")
    service = _service(folder, tmp_path)
    service.controller.all_files = sources
    metrics = ScanMetrics()

    # Building the sanitizer must never resolve/stat source paths.  This is the exact
    # network/removable-drive-hostile operation that dominated the pre-fix profile.
    monkeypatch.setattr(Path, "resolve", lambda self, *a, **k: pytest.fail("redaction resolved a source path"))
    redactor = service._source_path_redactor(metrics)
    text = f"failed: {sources[0]} and {sources[1].as_posix()}"
    safe = service._redact_source_paths(text, redactor)
    assert str(folder) not in safe
    assert sources[0].name in safe and sources[1].name in safe
    assert metrics.redaction_context_builds == 1 and metrics.redaction_sources == 2

    # Reusing the same context for arbitrarily many DTO fields performs no rebuild.
    for _ in range(500):
        assert service._redact_source_paths(text, redactor) == safe
    assert metrics.redaction_context_builds == 1


def test_stat_error_and_corrupt_pptx_are_isolated_and_scan_terminates(tmp_path, monkeypatch):
    good = tmp_path / "260901001_good.pptx"
    bad = tmp_path / "260901002_unreadable.pptx"
    good.write_bytes(b"corrupt-but-scan-does-not-open-it")
    bad.write_bytes(b"also-corrupt")
    real_stat = Path.stat

    def isolated_stat(path, *args, **kwargs):
        if path.name == bad.name:
            raise OSError("synthetic metadata failure")
        return real_stat(path, *args, **kwargs)

    monkeypatch.setattr(Path, "stat", isolated_stat)
    metrics = ScanMetrics()
    intrinsic = IntrinsicScanCache()
    result = prescan([good, bad], month_period(2026, 9), None, None,
                     today=date(2026, 9, 30), metrics=metrics, intrinsic_cache=intrinsic)
    assert [item.path for item in result.items] == [good, bad]
    assert [item.action for item in result.items] == [ACTION_PROCESS, ACTION_PROCESS]
    assert metrics.metadata_reads == 2 and metrics.metadata_errors == 1
    assert metrics.intrinsic_cache_misses == 1 and metrics.intrinsic_cache_unavailable == 1
    assert intrinsic.entry_count == 1
    assert metrics.pptx_opens == 0 and metrics.powerpoint_com_starts == 0


def test_empty_folder_returns_an_empty_scanned_list(tmp_path):
    folder = tmp_path / "empty"
    folder.mkdir()
    service = _service(folder, tmp_path)
    response = BridgeService(service).scan_reports(_dto(folder))
    assert response["ok"] is True
    assert response["data"]["scan"]["scanned"] is True
    assert response["data"]["reports"] == []
    metrics = service._last_scan_metrics
    assert metrics.entries_inspected == 0 and metrics.candidate_files == 0
    assert metrics.metadata_reads == 0 and metrics.directory_traversals == 1


def test_watchdog_identifies_exact_slow_file_without_absolute_path(tmp_path, caplog):
    secret = tmp_path / "private" / "260901001 sensitive report.pptx"
    metrics = ScanMetrics(stall_seconds=0.02)
    with caplog.at_level(logging.WARNING, logger="report_extractor.scan"):
        metrics.start_watchdog()
        token, started = metrics.begin("SCAN_FILE", index=37, path=secret, phase="metadata")
        time.sleep(0.06)
        metrics.end(token, started)
        metrics.stop_watchdog()
    messages = [record.getMessage() for record in caplog.records]
    stalled = next(message for message in messages if message.startswith("SCAN_FILE_START"))
    assert "index=37" in stalled and "phase=metadata" in stalled and "status=still_running" in stalled
    assert secret.name in stalled
    assert str(secret.parent) not in stalled
    assert len([message for message in messages if message.startswith("SCAN_FILE_START")]) == 1


def test_month_filter_contract_stays_filename_metadata_only(tmp_path):
    inside = tmp_path / "260915001_inside.pptx"
    outside = tmp_path / "261015001_outside.pptx"
    unrelated = tmp_path / "notes.txt"
    for path in (inside, outside, unrelated):
        path.write_bytes(b"not-opened")
    metrics = ScanMetrics()
    result = prescan([inside, outside], month_period(2026, 9), None, None,
                     today=date(2026, 9, 30), metrics=metrics)
    assert [(item.path.name, item.action) for item in result.items] == [
        (inside.name, ACTION_PROCESS),
        (outside.name, ACTION_OUTSIDE_PERIOD),
    ]
    assert metrics.metadata_reads == 2 and metrics.accepted_files == 0  # controller owns final aggregate counts
    assert unrelated.name not in {item.path.name for item in result.items}
