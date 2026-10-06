"""Transient, filesystem-scoped identity for one source report.

This key is deliberately separate from Management Number (the Excel row key) and from image-learning
candidate IDs (reusable classification data). It is used only to keep report evidence and temporary
image assets isolated when different reports share the same filename, shape IDs, or image bytes.
"""
from __future__ import annotations

import hashlib
import os
from pathlib import Path
from typing import Union

PathLike = Union[str, os.PathLike[str]]


def report_scope_key(source: PathLike) -> str:
    """Opaque stable key for a source-file location; same basenames in different folders stay distinct.

    The absolute path is hashed rather than exposed in asset names or serialized learning records. Resolving the
    path also makes relative/absolute spellings of the same file share a scope. This runtime evidence key is not a
    replacement for the Management Number used to select an Excel row.
    """
    raw = os.fspath(source) if source is not None else ""
    if not str(raw).strip():
        return ""
    path = Path(raw).expanduser()
    try:
        canonical = path.resolve(strict=False)
    except (OSError, RuntimeError):
        canonical = Path(os.path.abspath(os.fspath(path)))
    normalized = os.path.normcase(os.path.normpath(str(canonical)))
    return hashlib.sha256(normalized.encode("utf-8", errors="surrogatepass")).hexdigest()[:20]
