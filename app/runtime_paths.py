"""Central runtime-path resolution (PROMPT-003).

Two different roots must never be confused:

* **resource root** – where read-only bundled files live (``assets/`` …).  Source mode: the project root;
  PyInstaller onedir: ``sys._MEIPASS`` (= ``<portable>/_internal``).
* **portable root** – the user-visible, writable application directory.  Source mode: the project root;
  packaged: the folder that contains ``ReportExtractor.exe``.  ``config/``, ``logs/`` and ``Output/`` are
  created there on demand.  Runtime state is NEVER written inside ``_internal`` and the current working
  directory is never used as a root (double-click, shortcuts, "Start in" folders, spaces and Vietnamese
  characters in the path all resolve the same way).
"""
from __future__ import annotations

import sys
from pathlib import Path
from typing import Optional

CONFIG_DIR_NAME = "config"
LOGS_DIR_NAME = "logs"
OUTPUT_DIR_NAME = "Output"
INTERNAL_DIR_NAME = "_internal"


def is_packaged() -> bool:
    return bool(getattr(sys, "frozen", False))


def _source_root() -> Path:
    return Path(__file__).resolve().parent.parent


def resource_root() -> Path:
    """Read-only bundled resources (assets)."""
    if is_packaged():
        meipass = getattr(sys, "_MEIPASS", None)
        if meipass:
            return Path(meipass)
        return Path(sys.executable).resolve().parent / INTERNAL_DIR_NAME
    return _source_root()


def portable_root() -> Path:
    """Writable application directory (folder of the executable when packaged, project root in source mode)."""
    if is_packaged():
        return Path(sys.executable).resolve().parent
    return _source_root()


def _writable_dir(name: str, create: bool) -> Path:
    root = portable_root()
    if root.name == INTERNAL_DIR_NAME:              # defensive: never treat the bundle folder as writable root
        root = root.parent
    p = root / name
    if create:
        try:
            p.mkdir(parents=True, exist_ok=True)
        except OSError:
            pass                                    # read-only location (e.g. Program Files): caller handles errors
    return p


def config_dir(create: bool = True) -> Path:
    return _writable_dir(CONFIG_DIR_NAME, create)


def logs_dir(create: bool = True) -> Path:
    return _writable_dir(LOGS_DIR_NAME, create)


def output_dir(create: bool = True) -> Path:
    return _writable_dir(OUTPUT_DIR_NAME, create)


def config_file() -> Path:
    return config_dir() / "config.json"


def legacy_config_file() -> Optional[Path]:
    """``config.json`` next to the executable / project root (pre-1.0.3 layout) – read-only migration source."""
    p = portable_root() / "config.json"
    return p if p.exists() else None


def ensure_portable_dirs() -> None:
    """A freshly extracted portable folder may lack config/, logs/, Output/ – create them (no admin needed)."""
    for fn in (config_dir, logs_dir, output_dir):
        fn(create=True)


def describe() -> str:
    """One-line, secret-free summary for the start-up log."""
    return (f"packaged={'true' if is_packaged() else 'false'} executable={sys.executable} "
            f"portable_root={portable_root()} resource_root={resource_root()}")
