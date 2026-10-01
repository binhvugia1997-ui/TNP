"""Setup/bootstrap support (tools/setup_support.py, setup.bat, run.bat) – static + logic checks."""
import re
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "tools"))
import setup_support as ss  # noqa: E402

from app.config import DEFAULT_MODEL, AppConfig, normalize_ollama_url  # noqa: E402
from app.ollama_client import preferred_model  # noqa: E402

ROOT = Path(__file__).resolve().parent.parent


def test_supported_python_range_is_consistent():                                   # 1
    assert ss.PYTHON_MIN < ss.PYTHON_MAX_EXCLUSIVE
    assert ss.python_version_ok((3, 10)) and ss.python_version_ok((3, 12)) and ss.python_version_ok((3, 13))
    assert not ss.python_version_ok((3, 9)) and not ss.python_version_ok((3, 14))
    assert ss.python_version_ok(sys.version_info)                                   # the suite itself runs on it
    readme = (ROOT / "README.md").read_text(encoding="utf-8")
    assert f"{ss.PYTHON_MIN[0]}.{ss.PYTHON_MIN[1]}" in readme and ss.PYTHON_RECOMMENDED in readme
    setup = (ROOT / "setup.bat").read_text(encoding="utf-8")
    assert ss.WINGET_PYTHON_ID in setup and f"py -{ss.PYTHON_RECOMMENDED}" in setup.replace("py %%V", "py -3.12")


def test_runtime_imports_and_requirements_cover_each_other():                      # 2
    assert ss.missing_runtime_imports(extra_windows=False) == []
    req = (ROOT / "requirements.txt").read_text(encoding="utf-8").lower()
    for _, dist in ss.RUNTIME_IMPORTS + ss.WINDOWS_ONLY_IMPORTS:
        assert dist.lower() in req, dist
    dev = (ROOT / "requirements-dev.txt").read_text(encoding="utf-8")
    assert "-r requirements.txt" in dev and "pytest" in dev and "pytest" not in req


def test_existing_config_compatible_and_set_model_uses_appconfig(tmp_path, monkeypatch):   # 3, 7
    assert DEFAULT_MODEL == "qwen3:4b" and ss.SETUP_MODELS == ("qwen3:1.7b", "qwen3:4b")
    cfg_path = tmp_path / "config.json"
    monkeypatch.setattr(AppConfig, "path", classmethod(lambda cls: cfg_path))
    AppConfig(ollama_server="192.168.1.50:11434").save()
    assert ss.cmd_set_model("qwen3:1.7b") == 0
    cfg = AppConfig.load()
    assert cfg.model == "qwen3:1.7b" and cfg.ollama_server == "http://192.168.1.50:11434"   # other settings kept
    assert AppConfig().model == DEFAULT_MODEL                                      # default itself untouched


def test_model_choice_does_not_touch_extraction_logic():                           # 4
    src = "".join((ROOT / "app" / f).read_text(encoding="utf-8") for f in
                  ("extractor.py", "content_region.py", "improvement_pictures.py", "excel_writer.py"))
    assert "setup_support" not in src and "SETUP_MODELS" not in src
    assert preferred_model(["qwen3:1.7b", "qwen3:4b"]) == "qwen3:4b"               # 4b still preferred


def test_no_model_and_remote_ollama_configs_remain_valid():                        # 5, 6
    c = AppConfig(model="")
    assert c.model == "" and normalize_ollama_url(c.ollama_server) == "http://127.0.0.1:11434"
    r = AppConfig(ollama_server="http://192.168.10.20:11434", model="qwen3:4b")
    assert normalize_ollama_url(r.ollama_server) == "http://192.168.10.20:11434"
    from app.gui_controller import GuiController
    ctl = GuiController(AppConfig(period_mode="all"))
    assert ctl.set_endpoint("192.168.10.20", 11434, "qwen3:4b") == ""
    assert "192.168.10.20:11434" in ctl.ai_status_text()


def test_ollama_list_parsing_and_missing_models():
    out = "NAME            ID      SIZE    MODIFIED\nqwen3:4b        abc     2.6 GB  2 days ago\nllama3:latest   def     4 GB    x\n"
    assert ss.parse_ollama_list(out) == ["qwen3:4b", "llama3:latest"]
    assert ss.models_missing(["qwen3:1.7b", "qwen3:4b"], ["qwen3:4b"]) == ["qwen3:1.7b"]
    assert ss.models_missing(["llama3"], ["llama3:latest"]) == []
    assert ss.parse_ollama_list("") == []


@pytest.mark.parametrize("name", ["setup.bat", "run.bat", "run_cli.bat"])
def test_batch_files_static_validation(name):
    raw = (ROOT / name).read_bytes()
    assert raw.startswith(b"@echo off") and b"\r\n" in raw and not raw.startswith(b"\xef\xbb\xbf")   # CRLF, no BOM
    text = raw.decode("utf-8")
    assert "chcp 65001" in text and 'cd /d "%~dp0"' in text                        # UTF-8 + repo-relative
    assert "D:\\arena" not in text and ".venv\\Scripts\\python.exe" in text
    assert "pause" in text or name == "run_cli.bat"                                  # terminal stays visible (CLI runs in a console)
    labels = set(re.findall(r"^:([A-Za-z_][\w]*)", text, flags=re.M))
    for target in re.findall(r"goto :([A-Za-z_][\w]*)", text):
        assert target in labels, target
    for target in re.findall(r"call :([A-Za-z_][\w]*)", text):
        assert target in labels, target
    forbidden = ("Set-ExecutionPolicy", "netsh advfirewall", "Set-MpPreference", "Invoke-Expression", "iex ", "curl ")
    assert not any(f in text for f in forbidden)


def test_setup_bat_flow_contents():
    text = (ROOT / "setup.bat").read_text(encoding="utf-8")
    assert "qwen3:4b ^(mặc định/khuyến nghị^)" in text and "[4] Không tải model" in text
    assert "Cài thêm công cụ phát triển/test? [y/N]" in text
    assert "Bạn có thể cấu hình Ollama trên máy khác" in text
    assert "ollama pull %%M" in text and "models-missing" in text                   # existing models skipped
    assert "Model AI có thể cần tải dữ liệu dung lượng lớn" in text
    assert "Ollama.Ollama" in text and "127.0.0.1:11434" not in text                # no running server required
    gi = (ROOT / ".gitignore").read_text(encoding="utf-8")
    assert ".venv/" in gi and "config.json" in gi
