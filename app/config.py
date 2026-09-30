"""Configuration persistence (config.json next to the executable) and helpers."""
from __future__ import annotations

import json
import os
import sys
from dataclasses import dataclass, asdict, field
from pathlib import Path
from typing import Any, Dict, List, Optional


DEFAULT_OLLAMA = "http://127.0.0.1:11434"


def app_base_dir() -> Path:
    """Directory that holds config.json.

    - Frozen (PyInstaller onedir): folder containing ReportExtractor.exe.
    - Development: project root (parent of the ``app`` package).
    """
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent.parent


def normalize_ollama_url(address: str) -> str:
    """Normalize a user-typed Ollama address to ``http://host:port``.

    Accepts ``192.168.1.50:11434``, ``http://192.168.1.50:11434/``,
    ``192.168.1.50`` (default port 11434), ``https://host:port/api/tags``.
    """
    addr = (address or "").strip()
    if not addr:
        return DEFAULT_OLLAMA
    scheme = "http"
    if "://" in addr:
        scheme, addr = addr.split("://", 1)
        scheme = scheme.lower() or "http"
    # drop any path component
    addr = addr.split("/", 1)[0]
    # strip trailing spaces / slashes
    addr = addr.strip().rstrip("/")
    if not addr:
        return DEFAULT_OLLAMA
    host, sep, port = addr.rpartition(":")
    if not sep or not port.isdigit():
        # No port supplied (or IPv6 literal without port)
        host = addr
        port = "11434"
    return f"{scheme}://{host}:{port}"


@dataclass
class AppConfig:
    ollama_server: str = DEFAULT_OLLAMA
    model: str = ""
    last_report_folder: str = ""
    last_template: str = ""
    last_output_folder: str = ""
    last_output_file: str = ""
    request_timeout: int = 180
    render_dpi: int = 150
    force_reprocess: bool = False
    fill_temporary_column: bool = False
    row_mode: str = "match"               # "match" (Management Number -> existing row) | "append"
    vendors: List[str] = field(default_factory=list)   # controlled Vendor list; empty -> built-in CANONICAL_VENDORS
    extra: Dict[str, Any] = field(default_factory=dict)

    # ------------------------------------------------------------------
    @classmethod
    def path(cls) -> Path:
        return app_base_dir() / "config.json"

    @classmethod
    def load(cls, path: Optional[os.PathLike | str] = None) -> "AppConfig":
        p = Path(path) if path else cls.path()
        cfg = cls()
        if p.exists():
            try:
                data = json.loads(p.read_text(encoding="utf-8"))
            except Exception:
                data = {}
            known = {f for f in cfg.__dataclass_fields__ if f != "extra"}
            for k, v in (data or {}).items():
                if k in known:
                    setattr(cfg, k, v)
                else:
                    cfg.extra[k] = v
        cfg.ollama_server = normalize_ollama_url(cfg.ollama_server)
        return cfg

    def save(self, path: Optional[os.PathLike | str] = None) -> Path:
        p = Path(path) if path else self.path()
        data = asdict(self)
        extra = data.pop("extra", {}) or {}
        data.update(extra)
        data["ollama_server"] = normalize_ollama_url(self.ollama_server)
        p.parent.mkdir(parents=True, exist_ok=True)
        tmp = p.with_suffix(".json.tmp")
        tmp.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        os.replace(tmp, p)
        return p
