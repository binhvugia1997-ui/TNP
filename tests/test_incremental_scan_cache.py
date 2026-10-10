"""PROMPT-031R: safe intrinsic cache identity and folder-level incremental delivery."""
from __future__ import annotations

import json
import os
import threading
from datetime import date
from pathlib import Path

import pytest

from app.application_service import ApplicationService
from app.config import AppConfig
from app.gui_controller import GuiController
from app.prescan import (ACTION_FAST_SKIP, ACTION_MASTER_COMPLETE, ACTION_OUTSIDE_PERIOD, ACTION_PROCESS,
                         ACTION_PROCESS_NEW_ROW, ACTION_SOURCE_DUPLICATE, FastScanCache, IntrinsicScanCache,
                         MasterLookup, month_period, prescan)
from app.scan_diagnostics import ScanMetrics

TODAY = date(2026, 10, 14)
OCTOBER = month_period(2026, 10)
NOVEMBER = month_period(2026, 11)


def _source(parent: Path, name: str, *, payload: bytes = b"source", mtime: float = 1_700_000_000) -> Path:
    parent.mkdir(parents=True, exist_ok=True)
    path = parent / name
    path.write_bytes(payload)
    os.utime(path, (mtime, mtime))
    return path


def _run(files, cache, *, period=OCTOBER, master=None, force=False, recent=None):
    metrics = ScanMetrics()
    result = prescan(files, period, recent, master, force=force, today=TODAY,
                     intrinsic_cache=cache, metrics=metrics)
    return result, metrics


def _semantic_rows(result):
    return [(item.index, item.path.name, item.management_number, item.occurrence_date,
             item.period_decision, item.duplicate_decision, item.cache_decision,
             item.master_decision, item.excel_row, item.action, item.reason)
            for item in result.items]


def test_intrinsic_cache_identity_invalidates_path_size_and_mtime_and_never_caches_dynamic_status(tmp_path):
    source = _source(tmp_path / "A", "261010001-VOC_report.pptx")
    intrinsic = IntrinsicScanCache()

    cold, cold_metrics = _run([source], intrinsic)
    assert cold.items[0].action == ACTION_PROCESS
    assert cold_metrics.intrinsic_cache_hits == 0
    assert cold_metrics.intrinsic_cache_misses == 1
    assert cold_metrics.intrinsic_derivations == 1

    warm, warm_metrics = _run([source], intrinsic)
    assert _semantic_rows(warm) == _semantic_rows(cold)
    assert warm_metrics.intrinsic_cache_hits == 1
    assert warm_metrics.intrinsic_cache_misses == 0
    assert warm_metrics.intrinsic_derivations == 0
    assert warm_metrics.metadata_reads == 1  # full identity still requires current stat metadata

    # Period is current context, not a cached row action.
    outside, outside_metrics = _run([source], intrinsic, period=NOVEMBER)
    assert outside.items[0].action == ACTION_OUTSIDE_PERIOD
    assert outside_metrics.intrinsic_cache_hits == 1

    # Master and force state are always recomputed from current context.
    missing_master = MasterLookup(lambda _manager: [], lambda _row: [])
    new_row, new_metrics = _run([source], intrinsic, master=missing_master)
    assert new_row.items[0].action == ACTION_PROCESS_NEW_ROW
    assert new_metrics.intrinsic_cache_hits == 1

    complete_master = MasterLookup(lambda _manager: [4], lambda _row: [])
    complete, complete_metrics = _run([source], intrinsic, master=complete_master)
    assert complete.items[0].action == ACTION_MASTER_COMPLETE
    assert complete_metrics.intrinsic_cache_hits == 1
    forced, forced_metrics = _run([source], intrinsic, master=complete_master, force=True)
    assert forced.items[0].action == ACTION_PROCESS
    assert forced_metrics.intrinsic_cache_hits == 1

    # The existing recent-success cache remains authoritative and dynamic too.
    recent = FastScanCache(tmp_path / "recent.json")
    assert recent.record(cold.items[0].management_number, source, "completed")
    fast, fast_metrics = _run([source], intrinsic, recent=recent)
    assert fast.items[0].action == ACTION_FAST_SKIP
    assert fast_metrics.intrinsic_cache_hits == 1

    original_mtime = source.stat().st_mtime
    source.write_bytes(b"source-with-a-different-size")
    os.utime(source, (original_mtime, original_mtime))
    resized, resized_metrics = _run([source], intrinsic)
    assert resized.items[0].action == ACTION_PROCESS
    assert resized_metrics.intrinsic_cache_hits == 0
    assert resized_metrics.intrinsic_cache_misses == 1
    assert resized_metrics.intrinsic_cache_invalidations == 1

    os.utime(source, (original_mtime + 10, original_mtime + 10))
    retimed, retimed_metrics = _run([source], intrinsic)
    assert retimed.items[0].action == ACTION_PROCESS
    assert retimed_metrics.intrinsic_cache_hits == 0
    assert retimed_metrics.intrinsic_cache_misses == 1
    assert retimed_metrics.intrinsic_cache_invalidations == 1

    # Same Manager/size/mtime in another path is a distinct source identity, never an alias/hit.
    moved = _source(tmp_path / "B", source.name, payload=source.read_bytes(), mtime=source.stat().st_mtime)
    other_path, other_metrics = _run([moved], intrinsic)
    assert other_path.items[0].action == ACTION_PROCESS
    assert other_metrics.intrinsic_cache_hits == 0
    assert other_metrics.intrinsic_cache_misses == 1
    assert other_metrics.intrinsic_cache_invalidations == 0
    assert intrinsic.entry_count == 2


def test_intrinsic_cache_is_session_local_and_bounded(tmp_path):
    intrinsic = IntrinsicScanCache(max_entries=2)
    sources = [_source(tmp_path / f"leaf-{index}", f"26101002{index}-VOC_report.pptx", mtime=1_700_010_000 + index)
               for index in range(3)]
    for source in sources:
        result, metrics = _run([source], intrinsic)
        assert result.items[0].action == ACTION_PROCESS
        assert metrics.intrinsic_cache_misses == 1
    assert intrinsic.entry_count == 2
    assert intrinsic.evictions == 1

    evicted, metrics = _run([sources[0]], intrinsic)
    assert evicted.items[0].action == ACTION_PROCESS
    assert metrics.intrinsic_cache_hits == 0 and metrics.intrinsic_cache_misses == 1


def test_cold_and_warm_duplicate_rows_are_equivalent_with_distinct_path_identities(tmp_path):
    old = _source(tmp_path / "01-old", "261010010-VOC_old.pptx", mtime=1_700_000_000)
    new = _source(tmp_path / "02-new", "261010010-VOC_new.pptx", mtime=1_700_000_900)
    intrinsic = IntrinsicScanCache()

    cold, cold_metrics = _run([old, new], intrinsic)
    warm, warm_metrics = _run([old, new], intrinsic)

    assert [item.action for item in cold.items] == [ACTION_SOURCE_DUPLICATE, ACTION_PROCESS]
    assert _semantic_rows(warm) == _semantic_rows(cold)
    assert cold_metrics.intrinsic_cache_misses == 2
    assert warm_metrics.intrinsic_cache_hits == 2
    assert warm_metrics.intrinsic_cache_misses == 0
    assert intrinsic.entry_count == 2


def _dto(folder: Path, output: Path) -> dict:
    return {
        "paths": {"reportFolder": str(folder), "template": "", "output": str(output)},
        "period": {"mode": "month", "month": "10", "year": "2026", "from": "", "to": ""},
        "forceReprocess": False,
        "ollama": {"host": "127.0.0.1", "port": "11434", "model": "qwen3:4b"},
        "update": {"path": "", "autoCheck": False},
    }


def _service(folder: Path, tmp_path: Path) -> ApplicationService:
    config_path = tmp_path / "config" / "settings.json"
    controller = GuiController(AppConfig(period_mode="month", period_month=10, period_year=2026),
                               config_path=config_path)
    controller._today = lambda: TODAY
    return ApplicationService(controller, config_path=config_path,
                              manual_fields_path=tmp_path / "config" / "manual_fields.json")


def _dashboard_semantics(state: dict) -> list:
    return [(row["id"], row["stt"], row["fileName"], row["managementNumber"], row["occurrenceDate"],
             row["status"], row["shortResult"], row["warning"], row["excelRow"], row["path"])
            for row in state["reports"]]


def test_incremental_leaf_snapshots_mix_hits_misses_invalidation_and_global_duplicates(tmp_path, monkeypatch):
    root = tmp_path / "reports"
    duplicate_old = _source(root / "01-early", "261010100-VOC_duplicate_old.pptx", mtime=1_700_000_000)
    early_hit = _source(root / "01-early", "261010101-VOC_early.pptx", mtime=1_700_000_100)
    invalidated = _source(root / "02-middle", "261010102-VOC_middle.pptx", mtime=1_700_000_200)
    duplicate_new = _source(root / "03-slow", "261010100-VOC_duplicate_new.pptx", mtime=1_700_000_900)
    slow_hit = _source(root / "03-slow", "261010103-VOC_slow.pptx", mtime=1_700_001_000)
    files = [duplicate_old, early_hit, invalidated, duplicate_new, slow_hit]

    service = _service(root, tmp_path)
    config = _dto(root, tmp_path / "output" / "result.xlsx")
    cold = service.scan_reports(config)
    cold_semantics = _dashboard_semantics(cold)
    cold_metrics = service._last_scan_metrics
    assert cold_metrics.intrinsic_cache_misses == len(files)
    assert cold_metrics.intrinsic_cache_hits == 0
    assert cold_metrics.leaf_folders == cold_metrics.leaf_deliveries == 3
    assert cold_metrics.pptx_opens == cold_metrics.powerpoint_com_starts == 0
    assert {row["fileName"]: row["status"] for row in cold["reports"]}[duplicate_old.name] == "source_duplicate"
    assert {row["fileName"]: row["status"] for row in cold["reports"]}[duplicate_new.name] == "waiting"

    # Preserve mtime while changing size: exactly one authorized identity is invalidated on the warm scan.
    original_mtime = invalidated.stat().st_mtime
    invalidated.write_bytes(b"middle-source-size-changed")
    os.utime(invalidated, (original_mtime, original_mtime))

    slow_started = threading.Event()
    release_slow = threading.Event()
    real_lookup = FastScanCache.lookup

    def gated_lookup(cache, manager, path, today=None, days=7):
        if Path(path).name == slow_hit.name:
            slow_started.set()
            assert release_slow.wait(timeout=10), "test did not release the slow final leaf"
        return real_lookup(cache, manager, path, today, days)

    monkeypatch.setattr(FastScanCache, "lookup", gated_lookup)
    outcome = {}

    def run_scan():
        try:
            outcome["state"] = service.scan_reports(config)
        except BaseException as exc:  # captured for the parent test thread
            outcome["error"] = exc

    thread = threading.Thread(target=run_scan, name="prompt031r-scan", daemon=True)
    thread.start()
    assert slow_started.wait(timeout=10), "scan never reached the deliberately slow final leaf"

    incremental = service.dashboard_state()
    scan = incremental["scan"]
    assert scan["inProgress"] is True
    assert scan["completedLeaves"] == 2 and scan["totalLeaves"] == 3
    assert scan["revision"] == 2
    assert {row["fileName"] for row in incremental["reports"]} == {
        duplicate_old.name, early_hit.name, invalidated.name,
    }
    # Duplicate context was global before delivery: the early leaf already has its correct final loser status.
    early_rows = {row["fileName"]: row for row in incremental["reports"]}
    assert early_rows[duplicate_old.name]["status"] == "source_duplicate"
    assert str(duplicate_new) not in json.dumps(incremental, ensure_ascii=False)
    assert all(row["path"] == "" for row in incremental["reports"])
    assert service.close_requested()[0] is False  # closing cannot tear down a live scan generation

    release_slow.set()
    thread.join(timeout=10)
    assert not thread.is_alive()
    assert "error" not in outcome
    warm = outcome["state"]
    warm_metrics = service._last_scan_metrics

    assert warm["scan"]["inProgress"] is False
    assert warm["scan"]["completedLeaves"] == warm["scan"]["totalLeaves"] == 3
    assert warm["scan"]["revision"] == 4  # three leaves plus one final aggregate snapshot
    assert _dashboard_semantics(warm) == cold_semantics
    assert len({row["id"] for row in warm["reports"]}) == len(files)
    assert warm_metrics.intrinsic_cache_hits == 4
    assert warm_metrics.intrinsic_cache_misses == 1
    assert warm_metrics.intrinsic_cache_invalidations == 1
    assert warm_metrics.metadata_reads == len(files)
    assert warm_metrics.leaf_folders == 3
    assert warm_metrics.leaf_deliveries == 3  # folder granularity, never five per-file deliveries
    assert warm_metrics.leaf_rows_delivered == len(files)  # each incremental row DTO is built once, not cumulatively
    assert warm_metrics.first_leaf_ms is not None
    assert warm_metrics.first_leaf_ms < warm_metrics.elapsed_ms
    assert warm_metrics.pptx_opens == warm_metrics.powerpoint_com_starts == 0
    assert service.dashboard_state()["scan"]["generation"] == warm["scan"]["generation"]


@pytest.mark.parametrize(("candidate_count", "leaf_count"), [(30, 3), (120, 6), (300, 10)])
def test_cold_warm_scaling_uses_work_counts_not_wall_clock(tmp_path, candidate_count, leaf_count):
    root = tmp_path / f"reports-{candidate_count}"
    for index in range(candidate_count):
        leaf = root / f"leaf-{index % leaf_count:02d}"
        _source(leaf, f"261010{index + 1:03d}-VOC_report_{index:04d}.pptx",
                payload=b"not-a-pptx", mtime=1_700_100_000 + index)

    service = _service(root, tmp_path)
    config = _dto(root, tmp_path / "unused.xlsx")
    config["paths"]["output"] = ""  # benchmark intrinsic cache only, not the separate recent-success cache

    cold = service.scan_reports(config)
    cold_metrics = service._last_scan_metrics
    cold_semantics = _dashboard_semantics(cold)
    assert cold_metrics.entries_inspected == candidate_count + leaf_count
    assert cold_metrics.candidate_files == candidate_count
    assert cold_metrics.metadata_reads == candidate_count
    assert cold_metrics.intrinsic_cache_hits == 0
    assert cold_metrics.intrinsic_cache_misses == candidate_count
    assert cold_metrics.intrinsic_derivations == candidate_count
    assert cold_metrics.leaf_folders == cold_metrics.leaf_deliveries == leaf_count
    assert cold_metrics.leaf_rows_delivered == candidate_count
    assert cold_metrics.directory_traversals == 1
    assert cold_metrics.pptx_opens == cold_metrics.powerpoint_com_starts == 0

    warm = service.scan_reports(config)
    warm_metrics = service._last_scan_metrics
    assert _dashboard_semantics(warm) == cold_semantics
    assert warm_metrics.entries_inspected == candidate_count + leaf_count
    assert warm_metrics.metadata_reads == candidate_count  # identity validation is intentionally not skipped
    assert warm_metrics.intrinsic_cache_hits == candidate_count
    assert warm_metrics.intrinsic_cache_misses == 0
    assert warm_metrics.intrinsic_derivations == 0
    assert warm_metrics.leaf_folders == warm_metrics.leaf_deliveries == leaf_count
    assert warm_metrics.leaf_rows_delivered == candidate_count
    assert warm_metrics.redaction_context_builds == 1
    assert warm_metrics.pptx_opens == warm_metrics.powerpoint_com_starts == 0
    assert warm_metrics.first_leaf_ms is not None and warm_metrics.first_leaf_ms < warm_metrics.elapsed_ms
