"""Recursive discovery of .pptx report files."""
from __future__ import annotations

import os
from pathlib import Path
from typing import List, Sequence, Union

PathLike = Union[str, os.PathLike]


def is_report_file(path: Path) -> bool:
    name = path.name
    if name.startswith("~$"):          # PowerPoint lock/temp files
        return False
    if name.startswith("."):
        return False
    return path.suffix.lower() in (".pptx", ".pptm")


def scan_folder(folder: PathLike) -> List[Path]:
    """Recursively find every .pptx in ``folder`` (sorted, deterministic)."""
    root = Path(folder)
    if not root.exists():
        return []
    if root.is_file():
        return [root] if is_report_file(root) else []
    found: List[Path] = []
    for dirpath, dirnames, filenames in os.walk(root):
        dirnames[:] = sorted(d for d in dirnames if not d.startswith("."))
        for fn in sorted(filenames):
            p = Path(dirpath) / fn
            if is_report_file(p):
                found.append(p)
    return found


def scan_inputs(inputs: Sequence[PathLike]) -> List[Path]:
    """Accept a mix of folders and files (e.g. from drag & drop). Dedupe, keep order."""
    seen = set()
    out: List[Path] = []
    for item in inputs:
        for p in scan_folder(item):
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
