"""Fast pre-scan layer for large recursive report folders.

Everything in this module works on CHEAP information only (path, file name, size, mtime, the master
workbook that is already open and a small JSON cache).  No PPTX is parsed, no slide is rendered and
Ollama is never called here.

Decision order (one report file):

    file name -> Management Number -> occurrence date (existing parser) -> processing period
    -> source duplicates (same Management Number) -> 7-day recent-success cache
    -> master Excel row lookup (absent -> MASTER_NOT_FOUND, never a new row) -> master row completeness -> PROCESS

Business strings (Vietnamese) are kept in ``ACTION_LABELS_VI``; code uses the ``ACTION_*`` constants.
"""
from __future__ import annotations

import calendar
import json
import logging
import os
import re
import threading
import unicodedata
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional, Sequence, Tuple

from .extractor import derive_occurrence_date, management_number_from_filename

LOG = logging.getLogger("report_extractor.prescan")

CACHE_DAYS = 7
CACHE_FILE_NAME = "fast_scan_cache.json"
CACHE_VERSION = 1
SAFE_CACHE_STATUSES = ("completed", "skipped")          # Hoàn thành / Hoàn thành — bổ sung / Bỏ qua — đã cập nhật

# final pre-scan actions
# canonical keys of PreScanResult.counts(): "master_not_found" (PROMPT-004: a Management Number absent from the
# Management-Number row creation arrived; consumers must use these names (default-safe access in display code)
PRESCAN_COUNT_KEYS = ("discovered", "outside_period", "source_duplicates", "fast_skipped", "master_complete",
                      "master_not_found", "incomplete", "invalid_management_number", "candidates")

ACTION_PROCESS = "PROCESS"
ACTION_OUTSIDE_PERIOD = "OUTSIDE_PERIOD"
ACTION_SOURCE_DUPLICATE = "SOURCE_DUPLICATE"
ACTION_FAST_SKIP = "FAST_SKIP"
ACTION_MASTER_COMPLETE = "MASTER_COMPLETE"
ACTION_MASTER_NOT_FOUND = "MASTER_NOT_FOUND"     # valid key absent from the master -> NOT written (never a new row)
PROCESS_ACTIONS = (ACTION_PROCESS,)
ACTION_INVALID_MGMT = "INVALID_MANAGEMENT_NUMBER"

ACTION_LABELS_VI = {
    ACTION_PROCESS: "Cần bổ sung dữ liệu",
    ACTION_MASTER_NOT_FOUND: "Không tìm thấy Management Number trong Excel",
    ACTION_OUTSIDE_PERIOD: "Bỏ qua ngoài thời gian xử lý",
    ACTION_SOURCE_DUPLICATE: "Trùng Management Number trong folder",
    ACTION_FAST_SKIP: "Bỏ qua nhanh — đã xử lý gần đây",
    ACTION_MASTER_COMPLETE: "Bỏ qua — đã cập nhật",
    ACTION_INVALID_MGMT: "Không xác định được Management Number từ tên file",
}

# pipeline status keys used for the row table / batch_result.json
ACTION_STATUS = {
    ACTION_OUTSIDE_PERIOD: "outside_period",
    ACTION_SOURCE_DUPLICATE: "source_duplicate",
    ACTION_FAST_SKIP: "fast_skip",
    ACTION_MASTER_COMPLETE: "skipped",
    ACTION_INVALID_MGMT: "not_written",
}


# ----------------------------------------------------------------------------
# Processing period
# ----------------------------------------------------------------------------
def fmt_date(d: Optional[date]) -> str:
    return f"{d:%d/%m/%Y}" if d else ""


def parse_vn_date(text: str) -> Tuple[Optional[date], str]:
    """'dd/mm/yyyy' (also dd-mm-yyyy, dd.mm.yyyy) -> (date, '') or (None, Vietnamese error)."""
    s = (text or "").strip()
    m = re.fullmatch(r"(\d{1,2})[/.\-](\d{1,2})[/.\-](\d{4})", s)
    if not m:
        return None, f"Ngày không hợp lệ: {text!r} (định dạng dd/mm/yyyy)"
    try:
        return date(int(m.group(3)), int(m.group(2)), int(m.group(1))), ""
    except ValueError:
        return None, f"Ngày không hợp lệ: {text!r}"


@dataclass(frozen=True)
class ProcessingPeriod:
    """Inclusive date window used by the pre-scan (``start``/``end`` None = unbounded)."""
    mode: str                                  # auto | month | range | all
    start: Optional[date] = None
    end: Optional[date] = None
    source: str = ""                           # e.g. the Excel file name the period was detected from

    def contains(self, d: Optional[date]) -> bool:
        if self.mode == "all":
            return True
        if d is None:
            return False
        if self.start and d < self.start:
            return False
        if self.end and d > self.end:
            return False
        return True

    @property
    def is_all(self) -> bool:
        return self.mode == "all"

    def iso(self) -> str:
        return "all" if self.is_all else f"{self.start:%Y-%m-%d}..{self.end:%Y-%m-%d}"

    def label_vi(self) -> str:
        if self.is_all:
            return "Thời gian xử lý: Tất cả"
        if self.mode == "auto":
            return f"Thời gian xử lý: {fmt_date(self.start)} → {fmt_date(self.end)} (từ tên file Excel)"
        if self.mode == "month":
            return f"Thời gian xử lý: Tháng {self.start:%m/%Y}"
        return f"Thời gian xử lý: {fmt_date(self.start)} → {fmt_date(self.end)}"

    def month_tuple(self) -> Optional[Tuple[int, int]]:
        """(year, month) when the window is exactly one calendar month."""
        if self.start and self.end and self.start.day == 1 and self.start.month == self.end.month \
                and self.start.year == self.end.year and self.end.day == calendar.monthrange(self.end.year, self.end.month)[1]:
            return self.start.year, self.start.month
        return None


def month_period(year: int, month: int, mode: str = "month", source: str = "") -> ProcessingPeriod:
    """Real calendar month, inclusive (09/2026 -> 01/09/2026..30/09/2026)."""
    year, month = int(year), int(month)
    if not 1 <= month <= 12:
        raise ValueError(f"Tháng không hợp lệ: {month}")
    last = calendar.monthrange(year, month)[1]
    return ProcessingPeriod(mode, date(year, month, 1), date(year, month, last), source)


def range_period(start_text: str, end_text: str) -> Tuple[Optional[ProcessingPeriod], str]:
    """Inclusive range from two dd/mm/yyyy strings; never swaps an invalid order."""
    d1, e1 = parse_vn_date(start_text)
    if e1:
        return None, f"Từ ngày: {e1}"
    d2, e2 = parse_vn_date(end_text)
    if e2:
        return None, f"Đến ngày: {e2}"
    if d1 > d2:
        return None, f"Từ ngày ({fmt_date(d1)}) phải nhỏ hơn hoặc bằng Đến ngày ({fmt_date(d2)})."
    return ProcessingPeriod("range", d1, d2), ""


ALL_PERIOD = ProcessingPeriod("all")

AUTO_PERIOD_FAIL_VI = "Không xác định được tháng từ tên file Excel. Vui lòng chọn tháng hoặc khoảng thời gian."
PERIOD_DIFFERS_VI = "Thời gian đã chọn khác với tháng nhận diện từ tên file Excel."


def _normalize_name(name: str) -> str:
    stem = Path(name).stem
    s = unicodedata.normalize("NFD", stem)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    s = s.replace("đ", "d").replace("Đ", "D").lower()
    s = re.sub(r"[_\-./\\,;()\[\]]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()


_PAT_THANG = re.compile(r"(?<![a-z0-9])(?:thang|t)\s*(\d{1,2})\s+(\d{4})(?![0-9])")
_PAT_MM_YYYY = re.compile(r"(?<![a-z0-9])(\d{1,2})\s+(\d{4})(?![0-9])")
_PAT_YYYY_MM = re.compile(r"(?<![a-z0-9])(\d{4})\s+(\d{1,2})(?![0-9])")


def detect_month_from_excel_name(name: str) -> Optional[Tuple[int, int]]:
    """(year, month) from an unambiguous destination-workbook name, else None (never guesses).

    Kiem_chung_09_2026 / 09-2026 / 09.2026 / 2026_09 / Thang_09_2026 / Tháng 9-2026 / T09_2026 -> (2026, 9).
    A bare number without a reliable 4-digit year (Kiem_chung_v2_09) gives None.
    """
    if not name:
        return None
    s = _normalize_name(name)
    found = set()
    for m in _PAT_THANG.finditer(s):
        mo, yr = int(m.group(1)), int(m.group(2))
        if 1 <= mo <= 12 and 2000 <= yr <= 2099:
            found.add((yr, mo))
    if not found:
        for m in _PAT_MM_YYYY.finditer(s):
            mo, yr = int(m.group(1)), int(m.group(2))
            if 1 <= mo <= 12 and 2000 <= yr <= 2099:
                found.add((yr, mo))
        for m in _PAT_YYYY_MM.finditer(s):
            yr, mo = int(m.group(1)), int(m.group(2))
            if 1 <= mo <= 12 and 2000 <= yr <= 2099:
                found.add((yr, mo))
    if len(found) != 1:
        return None                       # nothing or ambiguous -> no guessing
    return next(iter(found))


def auto_period_from_excel(excel_path: str) -> Optional[ProcessingPeriod]:
    ym = detect_month_from_excel_name(Path(excel_path).name if excel_path else "")
    if not ym:
        return None
    return month_period(ym[0], ym[1], mode="auto", source=Path(excel_path).name)


# ----------------------------------------------------------------------------
# 7-day recent-success cache
# ----------------------------------------------------------------------------
def normalize_source_path(p: Path) -> str:
    try:
        s = str(Path(p).resolve())
    except OSError:
        s = str(Path(p).absolute())
    return os.path.normcase(s).replace("\\", "/")


def cache_cutoff(today: date, days: int = CACHE_DAYS) -> date:
    """Entries whose occurrence date (from the Management Number) is <= cutoff expire."""
    return today - timedelta(days=days)


class FastScanCache:
    """Tiny JSON cache of recently *successful* reports (optimisation only, never business data).

    Schema: {"version": 1, "entries": {<management_number>: {"path", "size", "mtime", "occurrence_date",
    "processed_at", "status"}}}.  Any read/write problem is logged and ignored (cache missing/corrupt/
    read-only must never block the batch).
    """

    def __init__(self, path: Optional[Path]):
        self.path = Path(path) if path else None
        self.entries: Dict[str, Dict[str, Any]] = {}
        self.problem: str = ""
        self._lock = threading.Lock()
        self._load()

    # -- persistence -----------------------------------------------------
    def _load(self) -> None:
        if not self.path or not self.path.exists():
            return
        try:
            data = json.loads(self.path.read_text(encoding="utf-8"))
            if not isinstance(data, dict) or data.get("version") != CACHE_VERSION \
                    or not isinstance(data.get("entries"), dict):
                raise ValueError("incompatible cache format")
            self.entries = {str(k): v for k, v in data["entries"].items() if isinstance(v, dict)}
        except Exception as e:  # noqa: BLE001
            self.problem = f"cache unreadable ({type(e).__name__}: {e}) – rebuilt from future results"
            LOG.warning("CACHE_PROBLEM file=%s reason=%s", self.path, self.problem)
            self.entries = {}

    def flush(self) -> bool:
        """Atomic write (tmp + os.replace); False when the cache could not be written."""
        if not self.path:
            return False
        try:
            self.path.parent.mkdir(parents=True, exist_ok=True)
            tmp = self.path.with_name(self.path.name + ".tmp")
            tmp.write_text(json.dumps({"version": CACHE_VERSION, "entries": self.entries},
                                      ensure_ascii=False, indent=1), encoding="utf-8")
            os.replace(tmp, self.path)
            return True
        except Exception as e:  # noqa: BLE001
            self.problem = f"cache not writable ({type(e).__name__}: {e})"
            LOG.warning("CACHE_PROBLEM file=%s reason=%s", self.path, self.problem)
            return False

    # -- API ----------------------------------------------------------------
    def cleanup(self, today: Optional[date] = None, days: int = CACHE_DAYS) -> List[str]:
        """Drop entries whose Management-Number occurrence date is <= today - 7 days.  Touches ONLY
        the cache file.  Returns the expired Management Numbers."""
        today = today or date.today()
        cutoff = cache_cutoff(today, days)
        expired: List[str] = []
        with self._lock:
            for mgmt, ent in list(self.entries.items()):
                occ, _ = derive_occurrence_date(mgmt)
                if occ is None:
                    try:
                        occ = date.fromisoformat(str(ent.get("occurrence_date", "")))
                    except ValueError:
                        occ = None
                if occ is None or occ <= cutoff:
                    expired.append(mgmt)
                    LOG.info("CACHE_EXPIRE management_number=%s occurrence_date=%s cutoff=%s",
                             mgmt, occ.isoformat() if occ else "?", cutoff.isoformat())
                    del self.entries[mgmt]
        if expired:
            self.flush()
        return expired

    @staticmethod
    def _stat(path: Path) -> Tuple[Optional[int], Optional[float]]:
        try:
            st = Path(path).stat()
            return st.st_size, round(st.st_mtime, 3)
        except OSError:
            return None, None

    def lookup(self, mgmt: str, path: Path, today: Optional[date] = None, days: int = CACHE_DAYS) -> Tuple[bool, str]:
        """(hit, reason).  Reasons: recent_success_cache | not_cached | source_changed | status_not_safe | expired."""
        ent = self.entries.get(mgmt)
        if not ent:
            return False, "not_cached"
        if ent.get("status") not in SAFE_CACHE_STATUSES:
            return False, "status_not_safe"
        today = today or date.today()
        occ, _ = derive_occurrence_date(mgmt)
        if occ is None or occ <= cache_cutoff(today, days):
            return False, "expired"
        size, mtime = self._stat(path)
        if ent.get("path") != normalize_source_path(path) or size is None \
                or ent.get("size") != size or ent.get("mtime") != mtime:
            return False, "source_changed"
        return True, "recent_success_cache"

    def record(self, mgmt: str, path: Path, status: str) -> bool:
        """Store a SAFE terminal outcome; anything else is ignored (stays eligible)."""
        if not mgmt or status not in SAFE_CACHE_STATUSES:
            return False
        size, mtime = self._stat(path)
        if size is None:
            return False
        occ, _ = derive_occurrence_date(mgmt)
        with self._lock:
            self.entries[mgmt] = {
                "path": normalize_source_path(path),
                "size": size,
                "mtime": mtime,
                "occurrence_date": occ.isoformat() if occ else "",
                "processed_at": datetime.now().isoformat(timespec="seconds"),
                "status": status,
            }
        return self.flush()


# ----------------------------------------------------------------------------
# Pre-scan result model
# ----------------------------------------------------------------------------
@dataclass
class PreScanItem:
    index: int
    path: Path
    filename: str = ""
    management_number: str = ""
    occurrence_date: Optional[date] = None
    size: Optional[int] = None
    mtime: Optional[float] = None
    period_decision: str = ""          # inside | outside | no_date | n/a
    duplicate_decision: str = ""       # selected | ignored | unique
    cache_decision: str = ""           # hit | miss:<reason> | bypassed_force | n/a
    master_decision: str = ""          # found_complete | found_incomplete | not_found (-> not written) | n/a
    excel_row: Optional[int] = None
    action: str = ACTION_PROCESS
    reason: str = ""

    @property
    def status(self) -> str:
        return ACTION_STATUS.get(self.action, "waiting")

    @property
    def label_vi(self) -> str:
        return ACTION_LABELS_VI.get(self.action, self.action)


@dataclass
class PreScanResult:
    period: ProcessingPeriod
    items: List[PreScanItem] = field(default_factory=list)
    duplicates: Dict[str, Dict[str, Any]] = field(default_factory=dict)   # mgmt -> {"selected", "ignored"}

    # counters -----------------------------------------------------------
    def count(self, action: str) -> int:
        return sum(1 for it in self.items if it.action == action)

    @property
    def discovered(self) -> int:
        return len(self.items)

    @property
    def candidates(self) -> List[PreScanItem]:
        """Reports that really enter PPT processing: incomplete existing rows + new rows."""
        return [it for it in self.items if it.action in PROCESS_ACTIONS]

    def counts(self) -> Dict[str, int]:
        """Canonical pre-scan summary schema (``PRESCAN_COUNT_KEYS``) – every key is always present."""
        return {
            "discovered": self.discovered,
            "outside_period": self.count(ACTION_OUTSIDE_PERIOD),
            "source_duplicates": self.count(ACTION_SOURCE_DUPLICATE),
            "fast_skipped": self.count(ACTION_FAST_SKIP),
            "master_complete": self.count(ACTION_MASTER_COMPLETE),
            "master_not_found": self.count(ACTION_MASTER_NOT_FOUND),
            "incomplete": self.count(ACTION_PROCESS),
            "invalid_management_number": self.count(ACTION_INVALID_MGMT),
            "candidates": len(self.candidates),
        }

    def summary_lines_vi(self) -> List[str]:
        c = self.counts()
        return [
            self.period.label_vi(),
            f"Tổng file phát hiện: {c['discovered']}",
            f"Ngoài thời gian xử lý: {c['outside_period']}",
            f"Trùng Management Number trong folder: {c['source_duplicates']}",
            f"Bỏ qua nhanh — đã xử lý gần đây: {c['fast_skipped']}",
            f"Bỏ qua — Excel đã đầy đủ: {c['master_complete']}",
            f"Không tìm thấy Management Number trong Excel: {c['master_not_found']}",
            f"Cần bổ sung dữ liệu: {c['incomplete']}",
            f"Không xác định được Management Number từ tên file: {c['invalid_management_number']}",
            f"Cần xử lý thực tế: {c['candidates']}",
        ]


class MasterLookup:
    """Cheap view over the open master workbook: ``rows(mgmt) -> [row...]``, ``missing(row) -> [field...]``."""

    def __init__(self, rows: Callable[[str], List[int]], missing: Callable[[int], List[str]]):
        self.rows = rows
        self.missing = missing

    @classmethod
    def from_writer(cls, writer) -> "MasterLookup":
        return cls(writer.find_rows_by_management_number, writer.missing_managed_fields)


# ----------------------------------------------------------------------------
# The pre-scan itself
# ----------------------------------------------------------------------------
def prescan(files: Sequence[Path], period: Optional[ProcessingPeriod], cache: Optional[FastScanCache],
            master: Optional[MasterLookup], force: bool = False, today: Optional[date] = None,
            on_stage: Optional[Callable[[str], None]] = None) -> PreScanResult:
    """Classify every discovered file WITHOUT opening it (see module doc for the order)."""
    period = period or ALL_PERIOD
    today = today or date.today()
    stage = on_stage or (lambda m: None)
    res = PreScanResult(period=period)
    LOG.info("PRESCAN start files=%d period=%s force=%s", len(files), period.iso(), force)

    # 1. file name -> Management Number -> occurrence date -> period ---------------------------
    stage("Đang lọc theo thời gian...")
    for i, p in enumerate(files):
        p = Path(p)
        it = PreScanItem(index=i, path=p, filename=p.name)
        try:
            st = p.stat()
            it.size, it.mtime = st.st_size, round(st.st_mtime, 3)
        except OSError:
            pass
        it.management_number = management_number_from_filename(p.name)
        res.items.append(it)
        if not it.management_number:
            it.period_decision = "n/a"
            if master is not None or not period.is_all:
                # match mode keeps the existing "not written" outcome; a period filter cannot be applied either
                it.action, it.reason = ACTION_INVALID_MGMT, ACTION_LABELS_VI[ACTION_INVALID_MGMT]
            continue                     # append mode without a period: legacy behaviour, file is processed
        it.occurrence_date, why = derive_occurrence_date(it.management_number)
        if it.occurrence_date is None:
            it.period_decision = "no_date"
            if not period.is_all:
                it.action, it.reason = ACTION_OUTSIDE_PERIOD, f"Bỏ qua ngoài thời gian xử lý ({why})"
                LOG.info("PERIOD_SKIP management_number=%s occurrence_date=invalid reason=%s", it.management_number, why)
            continue
        if period.contains(it.occurrence_date):
            it.period_decision = "inside"
        else:
            it.period_decision = "outside"
            it.action = ACTION_OUTSIDE_PERIOD
            it.reason = f"Bỏ qua ngoài thời gian xử lý (ngày phát sinh {fmt_date(it.occurrence_date)})"
            LOG.info("PERIOD_SKIP management_number=%s occurrence_date=%s period=%s",
                     it.management_number, it.occurrence_date.isoformat(), period.iso())

    # 2. source duplicates (same Management Number in the folder tree) ------------------------
    groups: Dict[str, List[PreScanItem]] = {}
    for it in res.items:
        if it.action == ACTION_PROCESS and it.management_number:
            groups.setdefault(it.management_number, []).append(it)
    for mgmt, grp in groups.items():
        if len(grp) == 1:
            grp[0].duplicate_decision = "unique"
            continue
        # newest modified wins; ties -> deterministic (case-insensitive) path order
        ordered = sorted(grp, key=lambda x: (-(x.mtime or 0.0), normalize_source_path(x.path)))
        sel, ignored = ordered[0], ordered[1:]
        sel.duplicate_decision = "selected"
        for ig in ignored:
            ig.duplicate_decision = "ignored"
            ig.action = ACTION_SOURCE_DUPLICATE
            ig.reason = f"Trùng Management Number {mgmt} trong folder; đã chọn {sel.path}"
        res.duplicates[mgmt] = {"selected": str(sel.path), "ignored": [str(x.path) for x in ignored]}
        LOG.info("SOURCE_DUPLICATE management_number=%s selected=%s ignored=%s",
                 mgmt, sel.path, ";".join(str(x.path) for x in ignored))

    # 3. recent-success cache + 4. master Excel -----------------------------------------------
    stage("Đang kiểm tra dữ liệu đã xử lý...")
    for it in res.items:
        if it.action != ACTION_PROCESS or not it.management_number:
            continue
        mgmt = it.management_number
        cache_hit = False
        if force:
            it.cache_decision = "bypassed_force"
        elif cache is None:
            it.cache_decision = "n/a"
        else:
            hit, why = cache.lookup(mgmt, it.path, today)
            if hit:
                cache_hit = True
                it.cache_decision = "hit"
            else:
                it.cache_decision = f"miss:{why}"
                LOG.info("CACHE_MISS management_number=%s reason=%s", mgmt, why)

        if master is None:
            it.master_decision = "n/a"
            if cache_hit:
                it.action, it.reason = ACTION_FAST_SKIP, ACTION_LABELS_VI[ACTION_FAST_SKIP]
                LOG.info("FAST_SKIP management_number=%s reason=recent_success_cache", mgmt)
            continue

        rows = master.rows(mgmt)
        if not rows:
            # PROMPT-004: the master workbook is the only place a report row may come from – a Management Number
            # absent from it is reported and NOT written (no new row, no parsing, no Qwen).  The CURRENT workbook
            # is authoritative: a cache hit from another/previous workbook never hides it.
            it.master_decision = "not_found"
            if cache_hit:
                it.cache_decision = "miss:master_row_missing"
                LOG.info("CACHE_MISS management_number=%s reason=master_row_missing", mgmt)
            it.action = ACTION_MASTER_NOT_FOUND
            it.reason = f"Cần kiểm tra: Không tìm thấy Management Number {mgmt} trong Excel – không tạo dòng mới"
            LOG.info("MASTER_MISS management_number=%s reason=not_found action=MASTER_NOT_FOUND", mgmt)
            continue
        it.excel_row = rows[0]
        missing = master.missing(rows[0])
        if missing:
            it.master_decision = "found_incomplete"
            if cache_hit:
                it.cache_decision = "miss:master_row_incomplete"        # Excel row state is authoritative
                LOG.info("CACHE_MISS management_number=%s reason=master_row_incomplete", mgmt)
            it.action = ACTION_PROCESS
            continue
        it.master_decision = "found_complete"
        if force:
            it.action = ACTION_PROCESS                                 # --force rewrites complete rows
            continue
        if cache_hit:
            it.action, it.reason = ACTION_FAST_SKIP, ACTION_LABELS_VI[ACTION_FAST_SKIP]
            LOG.info("FAST_SKIP management_number=%s reason=recent_success_cache", mgmt)
        else:
            it.action, it.reason = ACTION_MASTER_COMPLETE, f"Bỏ qua — đã cập nhật (dòng {rows[0]})"
            LOG.info("MASTER_SKIP management_number=%s reason=row_complete row=%s", mgmt, rows[0])

    c = res.counts()
    LOG.info("PRESCAN done %s", " ".join(f"{k}={v}" for k, v in c.items()))
    return res
