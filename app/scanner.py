"""Recursive discovery of .pptx report files."""
from __future__ import annotations

import logging
import os
from pathlib import Path
from typing import Callable, List, Optional, Sequence, Union, TYPE_CHECKING

if TYPE_CHECKING:
    from .scan_diagnostics import ScanMetrics

PathLike = Union[str, os.PathLike]
LOG = logging.getLogger("report_extractor.scanner")

REPORT_SUFFIXES = (".pptx", ".pptm", ".ppt")     # legacy .ppt is listed so the user sees its conversion error
RejectHook = Optional[Callable[[Path, str], None]]


def rejection_reason(path: Path) -> str:
    """Why a PowerPoint-looking file is NOT a report candidate ('' = it is one).  Discovery is purely structural:
    only genuine system artefacts are rejected – NEVER the file name content (Management Number, model, parentheses,
    missing SEV/date, unknown vendor …); eligibility is decided later by the pre-scan and stays visible as a status."""
    name = path.name
    if path.suffix.lower() not in REPORT_SUFFIXES:
        return f"phần mở rộng không phải PowerPoint ({path.suffix or 'không có'})"
    if name.startswith("~$"):
        return "file khoá/tạm của PowerPoint (~$)"
    if name.startswith("."):
        return "file ẩn hệ thống (bắt đầu bằng '.')"
    return ""


def is_report_file(path: Path) -> bool:
    return rejection_reason(path) == ""


def _looks_like_powerpoint(name: str) -> bool:
    low = name.lower()
    return any(low.endswith(sfx) or low.endswith(sfx + ".tmp") for sfx in REPORT_SUFFIXES) or \
        (low.startswith("~$") and any(sfx[1:] in low for sfx in REPORT_SUFFIXES))


def _reject(path: Path, why: str, on_reject: RejectHook) -> None:
    LOG.info("SCAN_REJECT file=%s reason=%s", path.name, why)
    if on_reject is not None:
        on_reject(path, why)


def scan_folder(folder: PathLike, on_reject: RejectHook = None,
                metrics: Optional["ScanMetrics"] = None) -> List[Path]:
    """Recursively find every report candidate in ``folder`` (sorted, deterministic).

    Every PowerPoint-looking entry that is skipped is logged with its exact reason
    (``SCAN_REJECT``) – nothing is discarded silently.  Optional metrics count work only;
    they never alter discovery decisions.
    """
    root = Path(folder)
    if not root.exists():
        return []
    if root.is_file():
        if metrics is not None:
            metrics.entries_inspected += 1
        why = rejection_reason(root)
        if why:
            if metrics is not None:
                metrics.structural_rejected_files += 1
            _reject(root, why, on_reject)
            return []
        if metrics is not None:
            metrics.candidate_files += 1
        return [root]
    if metrics is not None:
        metrics.directory_traversals += 1
    found: List[Path] = []
    for dirpath, dirnames, filenames in os.walk(root):
        if metrics is not None:
            metrics.directories_inspected += 1
            metrics.entries_inspected += len(dirnames) + len(filenames)
        hidden = [d for d in dirnames if d.startswith(".")]
        if metrics is not None:
            metrics.directories_skipped += len(hidden)
        for d in hidden:
            _reject(Path(dirpath) / d, "thư mục ẩn (bắt đầu bằng '.') – không quét", on_reject)
        dirnames[:] = sorted(d for d in dirnames if not d.startswith("."))
        for fn in sorted(filenames):
            p = Path(dirpath) / fn
            why = rejection_reason(p)
            if not why:
                found.append(p)
                if metrics is not None:
                    metrics.candidate_files += 1
            elif _looks_like_powerpoint(fn):
                if metrics is not None:
                    metrics.structural_rejected_files += 1
                _reject(p, why, on_reject)
            elif metrics is not None:
                metrics.unrelated_files += 1
    return found


def scan_inputs(inputs: Sequence[PathLike], on_reject: RejectHook = None,
                metrics: Optional["ScanMetrics"] = None) -> List[Path]:
    """Accept a mix of folders and files (e.g. from drag & drop). Dedupe, keep order."""
    seen = set()
    out: List[Path] = []
    for item in inputs:
        for p in scan_folder(item, on_reject, metrics=metrics):
            key = str(p.resolve()).lower()
            if key not in seen:
                seen.add(key)
                out.append(p)
    return out


def parse_dnd_paths(data: str) -> List[str]:
    """Parse the Tk DND_Files string (``{C:/a b/x.pptx} C:/y.pptx``) into paths."""
    paths: List[str] = []
    buf = ""
    in_brace = False
    for ch in data:
        if ch == "{" and not in_brace:
            in_brace = True
            buf = ""
        elif ch == "}" and in_brace:
            in_brace = False
            paths.append(buf)
            buf = ""
        elif ch == " " and not in_brace:
            if buf:
                paths.append(buf)
                buf = ""
        else:
            buf += ch
    if buf:
        paths.append(buf)
    return [p for p in paths if p]
