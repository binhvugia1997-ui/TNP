"""Helpers behind setup.bat / run.bat (kept in Python so the batch files stay small and testable).

Usage (from setup.bat):
    python tools/setup_support.py check-python        -> exit 0 when THIS interpreter is supported
    python tools/setup_support.py verify              -> runtime import / config smoke test (exit 0 = OK)
    python tools/setup_support.py models-missing M... -> prints the selected models not yet in `ollama list`
    python tools/setup_support.py set-model NAME      -> set the initial model through AppConfig (safe, optional)
    python tools/setup_support.py summary             -> Vietnamese summary lines

Supported Python range is defined ONCE here (``PYTHON_MIN`` / ``PYTHON_MAX_EXCLUSIVE``) and mirrored in README.
"""
from __future__ import annotations

import importlib
import os
import shutil
import subprocess
import sys
from pathlib import Path
from typing import Iterable, List, Sequence, Tuple

from console_safe import child_env, configure_console

ROOT = Path(__file__).resolve().parent.parent

PYTHON_MIN = (3, 10)                 # `X | Y` annotations, dataclass features, current wheels (PyMuPDF, pywin32)
PYTHON_MAX_EXCLUSIVE = (3, 15)       # 3.14 verified: PyMuPDF (cp310-abi3), pywin32 312, Pillow 12 ship win_amd64 wheels
PYTHON_RECOMMENDED = "3.12"
WINGET_PYTHON_ID = "Python.Python.3.12"

# runtime modules that MUST import for the application to run (module name -> pip distribution)
RUNTIME_IMPORTS = (("pptx", "python-pptx"), ("openpyxl", "openpyxl"), ("PIL", "Pillow"), ("requests", "requests"),
                   ("fitz", "PyMuPDF"))
WINDOWS_ONLY_IMPORTS = (("tkinterdnd2", "tkinterdnd2"), ("win32com", "pywin32"))

# models offered by setup.bat (exact Ollama tags); DEFAULT_MODEL of the app stays the recommendation
SETUP_MODELS: Tuple[str, ...] = ("qwen3:1.7b", "qwen3:4b")


def python_version_ok(version: Sequence[int] = sys.version_info) -> bool:
    v = (int(version[0]), int(version[1]))
    return PYTHON_MIN <= v < PYTHON_MAX_EXCLUSIVE


def python_range_text() -> str:
    return (f">= {PYTHON_MIN[0]}.{PYTHON_MIN[1]} và < {PYTHON_MAX_EXCLUSIVE[0]}.{PYTHON_MAX_EXCLUSIVE[1]} "
            f"(khuyến nghị {PYTHON_RECOMMENDED})")


def cmd_check_python() -> int:
    v = sys.version_info
    if python_version_ok(v):
        print(f"Python {v.major}.{v.minor}.{v.micro} — OK ({sys.executable})")
        return 0
    print(f"Python {v.major}.{v.minor} không phù hợp. Cần Python {python_range_text()}")
    return 1


def missing_runtime_imports(extra_windows: bool = (os.name == "nt")) -> List[str]:
    missing: List[str] = []
    checks = list(RUNTIME_IMPORTS) + (list(WINDOWS_ONLY_IMPORTS) if extra_windows else [])
    for module, dist in checks:
        try:
            importlib.import_module(module)
        except Exception:  # noqa: BLE001
            missing.append(dist)
    return missing


def cmd_verify() -> int:
    """Lightweight end-of-setup verification: imports, app modules, config directory. Never processes a report."""
    ok = True
    missing = missing_runtime_imports()
    if missing:
        print("Thiếu thư viện: " + ", ".join(missing))
        ok = False
    else:
        print("Thư viện runtime: OK")
    try:
        import tkinter  # noqa: F401
        print("Tkinter (giao diện): OK")
    except Exception as e:  # noqa: BLE001
        print(f"Tkinter không khả dụng ({e}) – cài lại Python với tuỳ chọn tcl/tk.")
        ok = False
    sys.path.insert(0, str(ROOT))
    try:
        from app import main, batch_processor, excel_writer, gui_controller  # noqa: F401
        from app.config import AppConfig, DEFAULT_MODEL, app_base_dir
        base = app_base_dir()
        base.mkdir(parents=True, exist_ok=True)
        probe = base / ".re_setup_probe.tmp"
        probe.write_text("ok", encoding="utf-8")
        probe.unlink()
        cfg = AppConfig.load()
        print(f"Module ứng dụng: OK  (config: {AppConfig.path()}, model mặc định {cfg.model or DEFAULT_MODEL})")
    except Exception as e:  # noqa: BLE001
        print(f"Không nạp được module ứng dụng: {type(e).__name__}: {e}")
        ok = False
    print("Ollama (local): " + ("Đã cài đặt" if ollama_available() else "Không cài / dùng máy khác"))
    return 0 if ok else 1


# ----------------------------------------------------------------- Ollama helpers (CLI only, no server needed)
def ollama_available() -> bool:
    return shutil.which("ollama") is not None


def parse_ollama_list(output: str) -> List[str]:
    """Model tags from `ollama list` text (header line skipped, first column)."""
    names: List[str] = []
    for line in (output or "").splitlines():
        parts = line.split()
        if not parts or parts[0].upper() == "NAME":
            continue
        names.append(parts[0])
    return names


def installed_models() -> List[str]:
    if not ollama_available():
        return []
    try:
        out = subprocess.run(["ollama", "list"], capture_output=True, text=True, timeout=30, encoding="utf-8",
                             errors="replace", env=child_env())
        return parse_ollama_list(out.stdout) if out.returncode == 0 else []
    except Exception:  # noqa: BLE001
        return []


def normalize_tag(tag: str) -> str:
    t = (tag or "").strip().lower()
    return t if ":" in t else f"{t}:latest"


def models_missing(selected: Iterable[str], installed: Iterable[str]) -> List[str]:
    have = {normalize_tag(m) for m in installed}
    return [m for m in selected if normalize_tag(m) not in have]


def cmd_models_missing(selected: List[str]) -> int:
    installed = installed_models()
    for m in selected:
        if m in models_missing([m], installed):
            print(m)
        else:
            sys.stderr.write(f"{m} đã tồn tại — bỏ qua tải xuống.\n")
    return 0


def cmd_set_model(model: str) -> int:
    """Initial model through the EXISTING config mechanism (AppConfig) – never a second config format."""
    sys.path.insert(0, str(ROOT))
    from app.config import AppConfig
    cfg = AppConfig.load()
    cfg.model = model.strip()
    cfg.save()
    print(f"Đã đặt model ban đầu: {cfg.model} ({AppConfig.path()})")
    return 0


def cmd_summary() -> int:
    v = sys.version_info
    installed = installed_models()
    lines = ["Python:        " + f"{v.major}.{v.minor}.{v.micro}", "Môi trường:    .venv",
             "Thư viện:      " + ("OK" if not missing_runtime_imports() else "THIẾU"),
             "Ollama:        " + ("Đã cài" if ollama_available() else "Không cài / dùng máy khác"),
             "Model local:"]
    for m in SETUP_MODELS:
        lines.append(f"  {m:<12} " + ("Có" if not models_missing([m], installed) else "Không"))
    print("\n".join(lines))
    return 0


def main(argv: List[str]) -> int:
    configure_console()
    if not argv:
        print(__doc__)
        return 2
    cmd, args = argv[0], argv[1:]
    if cmd == "check-python":
        return cmd_check_python()
    if cmd == "verify":
        return cmd_verify()
    if cmd == "models-missing":
        return cmd_models_missing(args)
    if cmd == "set-model":
        return cmd_set_model(args[0]) if args else 2
    if cmd == "summary":
        return cmd_summary()
    print(f"Lệnh không hợp lệ: {cmd}")
    return 2


if __name__ == "__main__":
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")  # type: ignore[attr-defined]
    except Exception:  # noqa: BLE001
        pass
    sys.exit(main(sys.argv[1:]))
