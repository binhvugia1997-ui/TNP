"""Diagnostic checks shown in the GUI and available via ``--diag`` on the CLI."""
from __future__ import annotations

import platform
import sys
import tempfile
from pathlib import Path
from typing import List, Optional, Tuple

from . import __version__
from .config import AppConfig
from .excel_writer import validate_template
from .ollama_client import OllamaClient, OllamaError
from .qpn_renderer import renderer_status


def run_diagnostics(cfg: AppConfig, template: Optional[str] = None,
                    output_folder: Optional[str] = None, check_ollama: bool = True) -> List[Tuple[str, str]]:
    rows: List[Tuple[str, str]] = []
    rows.append(("Report Extractor", f"v{__version__} ({'portable' if getattr(sys, 'frozen', False) else 'dev'})"))
    rows.append(("Python/runtime", f"OK – {platform.python_version()} / {platform.system()} {platform.release()}"))
    try:
        import pptx, openpyxl, PIL, requests  # noqa: F401
        rows.append(("Thư viện", "OK (python-pptx, openpyxl, Pillow, requests)"))
    except Exception as e:  # noqa: BLE001
        rows.append(("Thư viện", f"LỖI: {e}"))
    rs = renderer_status()
    rows.append(("PowerPoint renderer", rs["powerpoint"]))
    rows.append(("LibreOffice renderer", rs["libreoffice"]))
    rows.append(("Built-in renderer", rs["builtin"]))

    if check_ollama:
        try:
            info = OllamaClient(cfg.ollama_server).test_connection()
            rows.append(("Ollama server", f"Đã kết nối – {info['server']} ({len(info['models'])} model)"))
            chosen = cfg.model or info.get("preferred") or ""
            ok = chosen in info["models"]
            rows.append(("Model Qwen đã chọn", f"{chosen or '(chưa chọn)'}" + ("" if ok else " – KHÔNG có trên server")))
        except OllamaError as e:
            rows.append(("Ollama server", f"Chưa kết nối – {e}"))
            rows.append(("Model Qwen đã chọn", cfg.model or "(chưa chọn)"))
    tpl = template or cfg.last_template
    if tpl and Path(tpl).exists():
        ok, msg = validate_template(Path(tpl))
        rows.append(("Form Excel", ("Hợp lệ – " if ok else "Không hợp lệ – ") + msg))
    else:
        rows.append(("Form Excel", "Chưa chọn"))
    out = output_folder or cfg.last_output_folder
    if out:
        try:
            Path(out).mkdir(parents=True, exist_ok=True)
            with tempfile.NamedTemporaryFile(dir=out, prefix=".re_write_test", delete=True):
                pass
            rows.append(("Thư mục kết quả", f"Ghi được – {out}"))
        except Exception as e:  # noqa: BLE001
            rows.append(("Thư mục kết quả", f"Không ghi được – {e}"))
    else:
        rows.append(("Thư mục kết quả", "Chưa chọn"))
    return rows


def format_diagnostics(rows: List[Tuple[str, str]]) -> str:
    w = max(len(k) for k, _ in rows) if rows else 20
    return "\n".join(f"{k.ljust(w)} : {v}" for k, v in rows)
