"""PROMPT-016 regressions for cp1252-safe build and release consoles."""
from __future__ import annotations

import io
import os
import subprocess
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools"))

from console_safe import child_env, configure_console, safe_text  # noqa: E402
import build_and_publish  # noqa: E402


class StrictCp1252:
    """A minimal legacy console stream without ``reconfigure``."""
    encoding = "cp1252"
    errors = "strict"

    def __init__(self):
        self.parts = []

    def write(self, text):
        text.encode(self.encoding, errors=self.errors)  # reproduce TextIOWrapper's failure
        self.parts.append(text)
        return len(text)

    def flush(self):
        return None

    def getvalue(self):
        return "".join(self.parts)


class ReconfigurableCp1252(StrictCp1252):
    def __init__(self):
        super().__init__()
        self.calls = 0

    def reconfigure(self, *, encoding=None, errors=None):
        self.calls += 1
        self.encoding = encoding or self.encoding
        self.errors = errors or self.errors


def test_raw_strict_cp1252_reproduces_original_unicode_failure():
    stream = StrictCp1252()
    with pytest.raises(UnicodeEncodeError):
        stream.write("BUILD OK ✓ — Việt Nam")


def test_safe_text_has_ascii_fallback_for_common_status_glyphs():
    source = "✓ ✔ ✗ ❌ ⚠ • ○ – … ‘dữ liệu’ “tốt”"
    result = safe_text(source, encoding="cp1252")
    result.encode("cp1252", errors="strict")
    for fallback in ("[OK]", "[FAIL]", "[WARN]", "*", "o", "...", "'", '"'):
        assert fallback in result
    assert "du lieu" in result


def test_configure_console_makes_stdout_and_stderr_safe_without_reconfigure(monkeypatch):
    out, err = StrictCp1252(), StrictCp1252()
    monkeypatch.setattr(sys, "stdout", out)
    monkeypatch.setattr(sys, "stderr", err)
    configure_console()
    print("✓ Hoàn thành ở Việt Nam")
    print("⚠ Cần kiểm tra", file=sys.stderr)
    assert "[OK]" in sys.stdout.getvalue() and "Hoan thanh o Viet Nam" in sys.stdout.getvalue()
    assert "[WARN]" in sys.stderr.getvalue() and "Can kiem tra" in sys.stderr.getvalue()


def test_configure_console_handles_reconfigurable_stream_and_is_idempotent(monkeypatch):
    out, err = ReconfigurableCp1252(), ReconfigurableCp1252()
    monkeypatch.setattr(sys, "stdout", out)
    monkeypatch.setattr(sys, "stderr", err)
    configure_console()
    assert sys.stdout is out and sys.stderr is err
    first_out, first_err = sys.stdout, sys.stderr
    configure_console()
    print("✓ UTF-8-safe")
    print("✓ UTF-8-safe", file=sys.stderr)
    assert sys.stdout is first_out and sys.stderr is first_err
    assert out.calls == 1 and err.calls == 1
    assert out.encoding == "utf-8" and out.errors == "replace"
    assert "✓ UTF-8-safe" in out.getvalue() and "✓ UTF-8-safe" in err.getvalue()


def test_build_and_publish_status_is_safe_but_log_keeps_unicode(monkeypatch, tmp_path):
    out = StrictCp1252()
    monkeypatch.setattr(sys, "stdout", out)
    configure_console()
    tee = build_and_publish.Tee(tmp_path / "build.log")
    tee("✓ Xử lý hoàn tất — dữ liệu đã sẵn sàng")
    tee.close()
    displayed = sys.stdout.getvalue()
    assert "[OK]" in displayed and "Xu ly hoan tat" in displayed
    assert "✓ Xử lý hoàn tất" in (tmp_path / "build.log").read_text(encoding="utf-8")


def test_child_env_forces_utf8_safe_python_io_and_process_smoke():
    base = dict(os.environ)
    base.pop("PYTHONUTF8", None)
    base["PYTHONIOENCODING"] = "cp1252:strict"
    code = "import sys; sys.stdout.write('Việt ✓')"
    bad = subprocess.run([sys.executable, "-c", code], env=base, capture_output=True)
    assert bad.returncode != 0
    good_env = child_env(base)
    assert good_env["PYTHONUTF8"] == "1"
    assert good_env["PYTHONIOENCODING"] == "utf-8:replace"
    good = subprocess.run([sys.executable, "-c", code], env=good_env, capture_output=True,
                          encoding="utf-8", errors="replace")
    assert good.returncode == 0 and good.stdout == "Việt ✓"


def test_streamed_build_child_receives_utf8_environment(monkeypatch):
    captured = {}

    class Child:
        stdout = io.StringIO("✓ child output\n")

        def wait(self):
            return 0

    def fake_popen(command, **kwargs):
        captured["command"] = command
        captured.update(kwargs)
        return Child()

    monkeypatch.setattr(build_and_publish.subprocess, "Popen", fake_popen)
    output = []
    assert build_and_publish.run_streamed(["fake-child"], output.append) == 0
    assert output == ["✓ child output"]
    assert captured["env"]["PYTHONUTF8"] == "1"
    assert captured["env"]["PYTHONIOENCODING"] == "utf-8:replace"


def test_console_stream_fallback_preserves_plain_text():
    stream = StrictCp1252()
    assert safe_text("Build 015", stream=stream) == "Build 015"
    assert safe_text("✓", stream=io.StringIO()) == "✓"
