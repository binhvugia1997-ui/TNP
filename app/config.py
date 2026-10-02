"""Configuration persistence (<portable>/config/config.json) and helpers."""
from __future__ import annotations

import re

import json
import os
from dataclasses import dataclass, asdict, field
from pathlib import Path
from typing import Any, Dict, List, Optional


DEFAULT_OLLAMA = "http://127.0.0.1:11434"
DEFAULT_MODEL = "qwen3:4b"
LOCAL_OLLAMA = DEFAULT_OLLAMA                   # http://127.0.0.1:11434 – always probed first
LOCAL_HOSTS = ("127.0.0.1", "localhost", "::1", "0.0.0.0")
DEFAULT_PROBE_TIMEOUT = 3


def is_local_host(host: str) -> bool:
    return str(host or "").strip().strip("[]").lower() in LOCAL_HOSTS



def app_base_dir() -> Path:
    """Writable portable application directory (see :mod:`app.runtime_paths`)."""
    from .runtime_paths import portable_root
    return portable_root()


def normalize_ollama_url(address: str) -> str:
    """Normalize a user-typed Ollama address to ``http://host:port``.

    Accepts ``192.168.1.50:11434``, ``http://192.168.1.50:11434/``,
    ``192.168.1.50`` (default port 11434), ``https://host:port/api/tags``.
    """
    addr = (address or "").strip()
    if not addr:
        return DEFAULT_OLLAMA
    scheme = "http"
    while "://" in addr:                      # tolerate pasted "http://http://host" – never emit a double scheme
        sch, addr = addr.split("://", 1)
        sch = sch.strip().lower()
        if sch in ("http", "https"):
            scheme = sch
    # drop any path component
    addr = addr.split("/", 1)[0]
    # strip trailing spaces / slashes
    addr = addr.strip().rstrip("/")
    if not addr:
        return DEFAULT_OLLAMA
    if addr.startswith("["):                                  # IPv6 literal "[::1]:11434"
        host, _, rest = addr[1:].partition("]")
        port = rest.lstrip(":").split(":")[0]
        return f"{scheme}://[{host}]:{int(port) if port.isdigit() else 11434}"
    parts = addr.split(":")
    if len(parts) > 1 and all(p.strip().isdigit() for p in parts[1:]) and parts[0].strip():
        host, port = parts[0].strip(), parts[-1].strip()      # "host:11434" (a duplicated ":11434" collapses)
    else:
        host, port = addr, "11434"                            # no port (or IPv6 literal without port)
    return f"{scheme}://{host}:{int(port)}"


def split_endpoint(address: str, default_port: int = 11434) -> tuple:
    """(host, port) from any accepted spelling: '192.168.1.50', 'host:1234', 'http://host:1234/'."""
    url = normalize_ollama_url(address)
    rest = url.split("://", 1)[1]
    host, _, port = rest.rpartition(":")
    try:
        return host, int(port)
    except ValueError:
        return host, default_port


_HOST_RE = re.compile(r"^(?:[A-Za-z0-9](?:[A-Za-z0-9\-]{0,61}[A-Za-z0-9])?)(?:\.[A-Za-z0-9](?:[A-Za-z0-9\-]{0,61}[A-Za-z0-9])?)*$")


def endpoint_problem(host: str, port) -> str:
    """Vietnamese message when the host/port pair is unusable, '' when fine."""
    h = (host or "").strip()
    if "://" in h or "/" in h or ":" in h:
        h, _ = split_endpoint(h)
    if not h:
        return "Chưa nhập IP / Server của Ollama."
    if not _HOST_RE.match(h):
        return f"Địa chỉ Ollama không hợp lệ: {host!r}"
    if h.count(".") == 3 and all(p.isdigit() for p in h.split(".")):
        if any(int(p) > 255 for p in h.split(".")):
            return f"Địa chỉ IP không hợp lệ: {host!r}"
    try:
        p = int(str(port).strip())
    except (TypeError, ValueError):
        return f"Port không hợp lệ: {port!r} (ví dụ 11434)"
    if not 1 <= p <= 65535:
        return f"Port không hợp lệ: {port!r} (1–65535)"
    return ""


@dataclass
class AppConfig:
    ollama_server: str = DEFAULT_OLLAMA
    model: str = DEFAULT_MODEL
    last_report_folder: str = ""
    last_template: str = ""
    last_output_folder: str = ""
    last_output_file: str = ""
    request_timeout: int = 180            # Qwen inference (/api/generate) – NOT used for connection probes
    probe_timeout: int = 3                # connection / model-list probe (/api/tags) – must fail fast
    render_dpi: int = 150
    force_reprocess: bool = False
    fill_temporary_column: bool = False
    row_mode: str = "match"               # "match" (Management Number -> existing row) | "append"
    vendors: List[str] = field(default_factory=list)   # controlled Vendor list; empty -> built-in CANONICAL_VENDORS
    # processing period (fast pre-scan): auto (from the Excel file name) | month | range | all
    period_mode: str = "auto"
    period_month: int = 0                 # last manual month/year (convenience only)
    period_year: int = 0
    period_from: str = ""                 # last manual range dd/mm/yyyy
    period_to: str = ""
    extra: Dict[str, Any] = field(default_factory=dict)

    # ------------------------------------------------------------------
    @classmethod
    def path(cls) -> Path:
        """``<portable>/config/config.json``; a pre-1.0.3 ``config.json`` next to the executable is still read
        (migration) until the new file exists."""
        from .runtime_paths import config_file, legacy_config_file
        p = config_file()
        if not p.exists():
            legacy = legacy_config_file()
            if legacy is not None:
                return legacy
        return p

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
        if path:
            p = Path(path)
        else:
            from .runtime_paths import config_file, legacy_config_file
            p = self.path()
            if p == legacy_config_file():
                p = config_file()                  # migrate: always write the portable config/ folder
        data = asdict(self)
        extra = data.pop("extra", {}) or {}
        data.update(extra)
        data["ollama_server"] = normalize_ollama_url(self.ollama_server)
        p.parent.mkdir(parents=True, exist_ok=True)
        tmp = p.with_suffix(".json.tmp")
        tmp.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        os.replace(tmp, p)
        return p
