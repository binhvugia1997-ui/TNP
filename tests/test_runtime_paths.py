"""PROMPT-003 §63 – runtime path resolution in source mode and simulated PyInstaller (frozen) mode."""
import json
import os
import sys
from pathlib import Path

import pytest

from app import runtime_paths as rp
from app.config import AppConfig, app_base_dir

ROOT = Path(__file__).resolve().parent.parent


def _freeze(monkeypatch, exe_dir: Path):
    exe_dir.mkdir(parents=True, exist_ok=True)
    internal = exe_dir / "_internal"
    internal.mkdir(exist_ok=True)
    monkeypatch.setattr(sys, "frozen", True, raising=False)
    monkeypatch.setattr(sys, "_MEIPASS", str(internal), raising=False)
    monkeypatch.setattr(sys, "executable", str(exe_dir / "ReportExtractor.exe"))
    return internal


def test_source_mode_roots_are_project_root(monkeypatch):
    monkeypatch.delattr(sys, "frozen", raising=False)
    assert rp.is_packaged() is False
    assert rp.resource_root() == ROOT and rp.portable_root() == ROOT and app_base_dir() == ROOT
    assert (rp.resource_root() / "assets" / "DejaVuSans.ttf").exists()
    assert rp.config_dir(create=False) == ROOT / "config"
    assert rp.logs_dir(create=False) == ROOT / "logs" and rp.output_dir(create=False) == ROOT / "Output"


def test_frozen_mode_separates_resource_and_portable_root(monkeypatch, tmp_path):
    exe_dir = tmp_path / "ReportExtractor_v1.0.3_Portable"
    internal = _freeze(monkeypatch, exe_dir)
    assert rp.is_packaged() is True
    assert rp.resource_root() == internal                     # bundled assets
    assert rp.portable_root() == exe_dir                      # writable state next to the exe
    assert rp.config_file() == exe_dir / "config" / "config.json"
    assert rp.logs_dir() == exe_dir / "logs" and rp.output_dir() == exe_dir / "Output"
    for d in ("config", "logs", "Output"):
        assert (exe_dir / d).is_dir()                         # auto-created, no admin
    # nothing writable is ever resolved inside _internal
    for p in (rp.config_dir(), rp.logs_dir(), rp.output_dir()):
        assert "_internal" not in p.parts


def test_frozen_mode_never_uses_cwd(monkeypatch, tmp_path):
    exe_dir = tmp_path / "App Dir"
    _freeze(monkeypatch, exe_dir)
    other = tmp_path / "somewhere else"
    other.mkdir()
    monkeypatch.chdir(other)
    assert rp.portable_root() == exe_dir
    assert rp.config_file().parent.parent == exe_dir
    assert not (other / "config").exists() and not (other / "logs").exists()


@pytest.mark.parametrize("folder", ["Báo cáo Tuần 38", "Thư mục có dấu cách và tiếng Việt", "Users/Nguyễn Văn A/Desktop"])
def test_paths_with_spaces_and_vietnamese(monkeypatch, tmp_path, folder):
    exe_dir = tmp_path / folder / "ReportExtractor_Portable"
    _freeze(monkeypatch, exe_dir)
    rp.ensure_portable_dirs()
    cfg = AppConfig()
    saved = cfg.save()
    assert saved == exe_dir / "config" / "config.json" and saved.exists()
    data = json.loads(saved.read_text(encoding="utf-8"))
    assert data["ollama_server"] == "http://127.0.0.1:11434" and data["model"] == "qwen3:4b"
    assert AppConfig.load().ollama_server == "http://127.0.0.1:11434"


def test_config_path_uses_config_subdir_and_migrates_legacy(monkeypatch, tmp_path):
    exe_dir = tmp_path / "portable"
    _freeze(monkeypatch, exe_dir)
    assert AppConfig.path() == exe_dir / "config" / "config.json"
    legacy = exe_dir / "config.json"                           # pre-1.0.3 layout next to the exe
    legacy.write_text(json.dumps({"ollama_server": "http://192.168.9.9:11434", "model": "qwen3:1.7b"}), encoding="utf-8")
    assert AppConfig.path() == legacy                          # read for migration …
    cfg = AppConfig.load()
    assert cfg.ollama_server == "http://192.168.9.9:11434"
    saved = cfg.save()                                         # … but always written to config/
    assert saved == exe_dir / "config" / "config.json"
    assert AppConfig.path() == saved                           # new file now takes precedence


def test_meipass_missing_falls_back_to_internal_dir(monkeypatch, tmp_path):
    exe_dir = tmp_path / "p"
    _freeze(monkeypatch, exe_dir)
    monkeypatch.delattr(sys, "_MEIPASS", raising=False)
    assert rp.resource_root() == exe_dir / "_internal"
    assert rp.portable_root() == exe_dir


def test_portable_root_never_resolves_to_internal(monkeypatch, tmp_path):
    exe_dir = tmp_path / "p"
    _freeze(monkeypatch, exe_dir)
    monkeypatch.setattr(sys, "executable", str(exe_dir / "_internal" / "python.exe"))   # defensive case
    assert rp.config_dir(create=False) == exe_dir / "config"


def test_describe_line_has_required_fields(monkeypatch, tmp_path):
    exe_dir = tmp_path / "p"
    _freeze(monkeypatch, exe_dir)
    line = rp.describe()
    assert "packaged=true" in line and f"executable={exe_dir / 'ReportExtractor.exe'}" in line
    assert f"portable_root={exe_dir}" in line
    monkeypatch.delattr(sys, "frozen", raising=False)
    assert "packaged=false" in rp.describe()


def test_font_candidates_use_resource_root(monkeypatch, tmp_path):
    from app import qpn_renderer
    exe_dir = tmp_path / "p"
    internal = _freeze(monkeypatch, exe_dir)
    assert qpn_renderer._font_candidates()[0] == str(internal / "assets" / "DejaVuSans.ttf")


def test_startup_error_log_written_next_to_exe(monkeypatch, tmp_path):
    from app import main as m
    exe_dir = tmp_path / "p"
    _freeze(monkeypatch, exe_dir)
    try:
        raise RuntimeError("boom at startup")
    except RuntimeError as e:
        target = m.record_startup_error(e)
    assert target == exe_dir / "logs" / "startup_error.log"
    text = target.read_text(encoding="utf-8")
    assert "boom at startup" in text and "version=1.1.0-beta build=009" in text
    assert "startup_error.log" in m.STARTUP_ERROR_VI and "Không thể khởi động Report Extractor" in m.STARTUP_ERROR_VI


def test_launch_gui_fatal_path_returns_nonzero_and_logs(monkeypatch, tmp_path):
    from app import main as m
    exe_dir = tmp_path / "p"
    _freeze(monkeypatch, exe_dir)
    shown = []
    monkeypatch.setattr(m, "_show_startup_error_dialog", lambda: shown.append(True))

    def broken_logging():
        raise OSError("disk full")
    monkeypatch.setattr(m, "_startup_logging", broken_logging)
    assert m.launch_gui() == 4
    assert shown == [True]
    assert "disk full" in (exe_dir / "logs" / "startup_error.log").read_text(encoding="utf-8")


def test_startup_logging_writes_startup_env_line(monkeypatch, tmp_path):
    import logging
    from app import main as m
    exe_dir = tmp_path / "p"
    _freeze(monkeypatch, exe_dir)
    m._startup_logging()
    for h in logging.getLogger("report_extractor").handlers:
        try:
            h.flush()
        except Exception:  # noqa: BLE001
            pass
    text = (exe_dir / "logs" / "app.log").read_text(encoding="utf-8")
    assert "STARTUP version=1.1.0-beta build=009 packaged=true" in text and f"portable_root={exe_dir}" in text
    assert os.sep in text
