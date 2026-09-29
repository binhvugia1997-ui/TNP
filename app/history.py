"""Duplicate protection: fingerprints of already-imported reports (per output folder)."""
from __future__ import annotations

import hashlib
import json
import os
import threading
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Optional


def sha256_file(path: Path, chunk: int = 1 << 20) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        while True:
            b = f.read(chunk)
            if not b:
                break
            h.update(b)
    return h.hexdigest()


def fingerprint(path: Path) -> str:
    """SHA-256 of (content hash + size + filename). Content-based, so a moved file
    with identical content is still recognised; a modified report gets a new id."""
    p = Path(path)
    st = p.stat()
    digest = sha256_file(p)
    raw = f"{digest}|{st.st_size}|{p.name.lower()}"
    return hashlib.sha256(raw.encode("utf-8")).hexdigest()[:32]


class History:
    def __init__(self, path: Path):
        self.path = Path(path)
        self._lock = threading.Lock()
        self._data: Dict[str, Dict[str, Any]] = {}
        if self.path.exists():
            try:
                self._data = json.loads(self.path.read_text(encoding="utf-8")) or {}
            except Exception:
                self._data = {}

    def lookup(self, fp: str) -> Optional[Dict[str, Any]]:
        return self._data.get(fp)

    def record(self, fp: str, source: Path, output: Path, row: Optional[int], status: str) -> None:
        with self._lock:
            self._data[fp] = {
                "source": str(source),
                "output": str(output),
                "row": row,
                "status": status,
                "size": Path(source).stat().st_size if Path(source).exists() else None,
                "mtime": Path(source).stat().st_mtime if Path(source).exists() else None,
                "time": datetime.now().isoformat(timespec="seconds"),
            }
            self._flush()

    def _flush(self) -> None:
        self.path.parent.mkdir(parents=True, exist_ok=True)
        tmp = self.path.with_suffix(".tmp")
        tmp.write_text(json.dumps(self._data, ensure_ascii=False, indent=2), encoding="utf-8")
        os.replace(tmp, self.path)
