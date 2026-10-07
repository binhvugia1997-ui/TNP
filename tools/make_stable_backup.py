"""Create (or verify) a STABLE source backup of the project – PROMPT-006 §0.

Copies every Git-tracked source file (``git ls-files``) into ``backup/<name>/`` and refuses to overwrite an
existing backup.  Development artefacts and user/sample data are never copied even when they are tracked by
mistake.  Usage::

    python tools/make_stable_backup.py ReportExtractor_v1.0.4_Build004_STABLE
"""
from __future__ import annotations

import shutil
import subprocess
import sys
from pathlib import Path
from typing import Iterable, List, Sequence

from console_safe import child_env, configure_console

EXCLUDED_DIRS = ("backup", ".venv", ".venv-build", "build", "dist", "Output", "logs", "config", "sample_data",
                 "learning_data", "release", "update_staging", "update_backup", "__pycache__", ".git",
                 ".pytest_cache")
EXCLUDED_SUFFIXES = (".pptx", ".ppt", ".xlsx", ".xlsm", ".log", ".pyc", ".zip")
REQUIRED = ("app/__init__.py", "app/gui.py", "app/extractor.py", "app/improvement_pictures.py", "requirements.txt")


class BackupError(RuntimeError):
    pass


def tracked_files(root: Path) -> List[str]:
    out = subprocess.run(["git", "-C", str(root), "ls-files", "-z"], capture_output=True, check=True,
                         env=child_env())
    return [p for p in out.stdout.decode("utf-8", "replace").split("\0") if p]


def is_excluded(rel: str) -> bool:
    parts = Path(rel).parts
    if any(p in EXCLUDED_DIRS for p in parts[:-1]) or parts[0] in EXCLUDED_DIRS:
        return True
    return rel.lower().endswith(EXCLUDED_SUFFIXES)


def select_files(files: Iterable[str]) -> List[str]:
    return sorted(f for f in files if not is_excluded(f))


def make_backup(root: Path, name: str, files: Sequence[str] | None = None) -> Path:
    root = Path(root)
    target = root / "backup" / name
    if target.exists():
        raise BackupError(f"Backup đã tồn tại, không ghi đè: {target}")
    files = select_files(files if files is not None else tracked_files(root))
    missing = [r for r in REQUIRED if r not in files]
    if missing:
        raise BackupError(f"Thiếu file nguồn bắt buộc: {', '.join(missing)}")
    for rel in files:
        src = root / rel
        if not src.is_file():
            continue
        dst = target / rel
        dst.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(src, dst)
    return target


def verify_backup(target: Path) -> List[str]:
    """Problems found in an existing backup folder (empty list = verified)."""
    target = Path(target)
    problems: List[str] = []
    if not target.is_dir():
        return [f"không tồn tại: {target}"]
    for r in REQUIRED:
        if not (target / r).is_file():
            problems.append(f"thiếu {r}")
    for p in target.rglob("*"):
        rel = p.relative_to(target).as_posix()
        if p.is_dir():
            if p.name in EXCLUDED_DIRS:
                problems.append(f"thư mục không được phép: {rel}")
        elif rel.lower().endswith(EXCLUDED_SUFFIXES):
            problems.append(f"file dữ liệu không được phép: {rel}")
    return problems


def main(argv: Sequence[str]) -> int:
    configure_console()
    if len(argv) != 2:
        print(__doc__)
        return 2
    root = Path(__file__).resolve().parent.parent
    try:
        target = make_backup(root, argv[1])
    except (BackupError, subprocess.CalledProcessError) as e:
        print(f"BACKUP FAILED — NO SOURCE MODIFICATIONS MADE ({e})")
        return 1
    problems = verify_backup(target)
    if problems:
        print("BACKUP FAILED — " + "; ".join(problems))
        return 1
    print(f"STABLE BACKUP VERIFIED {target}")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
