"""Logging: application log, errors.log and machine readable batch_result.json."""
from __future__ import annotations

import json
import logging
import os
import threading
from dataclasses import dataclass, field, asdict
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional

LOG = logging.getLogger("report_extractor")


def setup_logging(log_dir: Path, level: int = logging.INFO) -> Path:
    """Configure root ``report_extractor`` logger with app.log + errors.log."""
    log_dir = Path(log_dir)
    log_dir.mkdir(parents=True, exist_ok=True)
    LOG.setLevel(logging.DEBUG)
    # avoid duplicate handlers when called several times (new output folder)
    for h in list(LOG.handlers):
        if getattr(h, "_re_managed", False):
            LOG.removeHandler(h)
            try:
                h.close()
            except Exception:
                pass
    fmt = logging.Formatter("%(asctime)s [%(levelname)s] %(name)s: %(message)s")

    app_h = logging.FileHandler(log_dir / "app.log", encoding="utf-8")
    app_h.setLevel(level)
    app_h.setFormatter(fmt)
    app_h._re_managed = True  # type: ignore[attr-defined]
    LOG.addHandler(app_h)

    err_h = logging.FileHandler(log_dir / "errors.log", encoding="utf-8")
    err_h.setLevel(logging.WARNING)
    err_h.setFormatter(fmt)
    err_h._re_managed = True  # type: ignore[attr-defined]
    LOG.addHandler(err_h)

    if not any(isinstance(h, logging.StreamHandler) and not isinstance(h, logging.FileHandler)
               for h in LOG.handlers):
        sh = logging.StreamHandler()
        sh.setLevel(level)
        sh.setFormatter(fmt)
        LOG.addHandler(sh)
    return log_dir / "errors.log"


@dataclass
class FileResult:
    """One entry of batch_result.json."""
    source_file: str
    status: str = "waiting"            # completed | needs_review | error | skipped
    management_number: str = ""
    vendor: str = ""
    occurrence_date: str = ""
    model: str = ""
    item: str = ""
    qpn_slide: Optional[int] = None
    cause_slides: List[int] = field(default_factory=list)
    improvement_slides: List[int] = field(default_factory=list)
    improvement_image_slides: List[int] = field(default_factory=list)
    temporary_slides: List[int] = field(default_factory=list)
    verify_slides: List[int] = field(default_factory=list)
    defect_slide: Optional[int] = None
    excel_row: Optional[int] = None
    review_reasons: List[str] = field(default_factory=list)
    blank_fields: List[str] = field(default_factory=list)
    filled_fields: List[str] = field(default_factory=list)     # match mode: managed fields written this run
    classifier_notes: List[str] = field(default_factory=list)
    confidence: Optional[float] = None
    error: str = ""
    classifier: str = ""               # "qwen" | "qwen+heuristic" | "heuristic"
    qpn_source: str = ""               # explicit_text | metadata | qpn_heading | defect_structure | structural | llm
    qpn_override: str = ""             # why an LLM QPN suggestion was rejected
    image_slides_structural: List[int] = field(default_factory=list)
    image_slides_llm: List[int] = field(default_factory=list)
    qpn_renderer: str = ""
    fingerprint: str = ""
    started_at: str = ""
    finished_at: str = ""

    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)


class BatchResultLog:
    """Thread-safe writer for logs/batch_result.json (rewritten after each file)."""

    def __init__(self, path: Path):
        self.path = Path(path)
        self._lock = threading.Lock()
        self._results: List[Dict[str, Any]] = []
        self.meta: Dict[str, Any] = {"started_at": datetime.now().isoformat(timespec="seconds")}

    def add(self, result: FileResult) -> None:
        with self._lock:
            self._results.append(result.to_dict())
            self._flush()

    def finish(self, summary: Dict[str, Any]) -> None:
        with self._lock:
            self.meta["finished_at"] = datetime.now().isoformat(timespec="seconds")
            self.meta["summary"] = summary
            self._flush()

    def _flush(self) -> None:
        self.path.parent.mkdir(parents=True, exist_ok=True)
        data = {"meta": self.meta, "results": self._results}
        tmp = self.path.with_suffix(".tmp")
        tmp.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        os.replace(tmp, self.path)
