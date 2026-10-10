"""Bounded diagnostics for the Quét list-discovery path.

The scan remains sequential.  This module observes it; it does not add worker concurrency,
timeouts, retries, or caching.  INFO logs stay aggregate-only during a normal scan.  If one
stage remains active long enough to look hung, a daemon watchdog emits a bounded snapshot
containing only the stage, list index and a safe basename/hash label (never an absolute path).
"""
from __future__ import annotations

import hashlib
import heapq
import logging
import re
import threading
import time
from contextlib import contextmanager
from pathlib import Path
from typing import Callable, Iterator, List, Optional, Tuple

LOG = logging.getLogger("report_extractor.scan")
SLOW_OPERATION_LIMIT = 5
STALL_LOG_LIMIT = 8
DEFAULT_STALL_SECONDS = 2.0


def safe_file_label(path: object) -> str:
    """Bounded, log-safe basename plus a short identity hash; never return a directory."""
    name = Path(path).name if path is not None else ""
    name = re.sub(r"[\x00-\x1f\x7f]+", "?", name).strip() or "(unknown)"
    digest = hashlib.sha256(name.encode("utf-8", "replace")).hexdigest()[:10]
    if len(name) > 96:
        name = name[:93] + "..."
    return f"{name}#{digest}"


class ScanMetrics:
    """One in-memory measurement record for one Quét call."""

    def __init__(self, *, clock: Callable[[], float] = time.perf_counter,
                 stall_seconds: float = DEFAULT_STALL_SECONDS) -> None:
        self.clock = clock
        self.stall_seconds = max(0.01, float(stall_seconds))
        self.started_at = self.clock()

        self.entries_inspected = 0
        self.directories_inspected = 0
        self.directories_skipped = 0
        self.candidate_files = 0
        self.structural_rejected_files = 0
        self.unrelated_files = 0
        self.directory_traversals = 0
        self.enumerate_ms = 0.0
        self.filter_ms = 0.0
        self.accepted_files = 0
        self.status_skipped_files = 0

        self.metadata_reads = 0
        self.metadata_errors = 0
        self.metadata_ms = 0.0
        self.master_open_ms = 0.0
        self.per_file_operations = 0
        self.per_file_ms = 0.0
        self.pptx_opens = 0
        self.parse_pptx_ms = 0.0
        self.powerpoint_com_starts = 0
        self.powerpoint_com_ms = 0.0
        self.build_response_ms = 0.0
        self.redaction_context_builds = 0
        self.redaction_sources = 0

        self._lock = threading.Lock()
        self._active: List[Tuple[int, str, Optional[int], str, str, float]] = []
        self._next_token = 0
        self._slowest: List[Tuple[float, int, str, int, str]] = []
        self._slow_sequence = 0
        self._watchdog_done = threading.Event()
        self._watchdog_thread: Optional[threading.Thread] = None
        self._stall_emitted: set = set()

    @property
    def elapsed_ms(self) -> float:
        return max(0.0, (self.clock() - self.started_at) * 1000.0)

    def begin(self, stage: str, *, index: Optional[int] = None, path: object = None,
              phase: str = "") -> Tuple[int, float]:
        started = self.clock()
        label = safe_file_label(path) if path is not None else ""
        with self._lock:
            self._next_token += 1
            token = self._next_token
            self._active.append((token, stage, index, phase, label, started))
        if stage == "SCAN_FILE":
            LOG.debug("SCAN_FILE_START index=%s phase=%s file=%s", index, phase or "inspect", label)
        return token, started

    def end(self, token: int, started: float) -> float:
        elapsed_ms = max(0.0, (self.clock() - started) * 1000.0)
        ended = None
        with self._lock:
            for pos in range(len(self._active) - 1, -1, -1):
                if self._active[pos][0] == token:
                    ended = self._active.pop(pos)
                    break
        if ended and ended[1] == "SCAN_FILE":
            _token, _stage, index, phase, label, _started = ended
            LOG.debug("SCAN_FILE_END index=%s phase=%s file=%s elapsed_ms=%.3f",
                      index, phase or "inspect", label, elapsed_ms)
        return elapsed_ms

    @contextmanager
    def operation(self, stage: str, *, index: Optional[int] = None, path: object = None,
                  phase: str = "") -> Iterator[None]:
        token, started = self.begin(stage, index=index, path=path, phase=phase)
        try:
            yield
        finally:
            self.end(token, started)

    def record_file_operation(self, index: int, path: object, phase: str, elapsed_ms: float) -> None:
        label = safe_file_label(path)
        self.per_file_operations += 1
        self.per_file_ms += elapsed_ms
        self._slow_sequence += 1
        item = (float(elapsed_ms), self._slow_sequence, phase, int(index), label)
        if len(self._slowest) < SLOW_OPERATION_LIMIT:
            heapq.heappush(self._slowest, item)
        elif item > self._slowest[0]:
            heapq.heapreplace(self._slowest, item)

    def slowest_operations(self) -> List[Tuple[float, str, int, str]]:
        return [(elapsed, phase, index, label)
                for elapsed, _seq, phase, index, label in sorted(self._slowest, reverse=True)]

    def start_watchdog(self) -> None:
        if self._watchdog_thread is not None:
            return
        self._watchdog_done.clear()
        self._watchdog_thread = threading.Thread(target=self._watchdog,
                                                 name="scan-diagnostics-watchdog", daemon=True)
        self._watchdog_thread.start()

    def stop_watchdog(self) -> None:
        self._watchdog_done.set()
        thread = self._watchdog_thread
        if thread is not None and thread is not threading.current_thread():
            thread.join(timeout=min(1.0, self.stall_seconds + 0.1))

    def _watchdog(self) -> None:
        interval = max(0.01, min(0.25, self.stall_seconds / 4.0))
        while not self._watchdog_done.wait(interval):
            now = self.clock()
            with self._lock:
                current = self._active[-1] if self._active else None
            if current is None:
                continue
            _token, stage, index, phase, label, started = current
            elapsed_ms = max(0.0, (now - started) * 1000.0)
            if elapsed_ms < self.stall_seconds * 1000.0:
                continue
            identity = (stage, index, phase, label)
            if identity in self._stall_emitted or len(self._stall_emitted) >= STALL_LOG_LIMIT:
                continue
            self._stall_emitted.add(identity)
            if stage == "SCAN_FILE":
                # This INFO line appears only for a genuinely slow/stuck per-file operation.  Normal
                # per-file START/END diagnostics stay at DEBUG and therefore do not flood app.log.
                LOG.warning("SCAN_FILE_START index=%s phase=%s file=%s elapsed_ms=%.3f status=still_running",
                            index, phase or "inspect", label, elapsed_ms)
            else:
                LOG.warning("SCAN_STALLED stage=%s elapsed_ms=%.3f status=still_running",
                            stage, elapsed_ms)
