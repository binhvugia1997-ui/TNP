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

import console_safe  # noqa: E402
from console_safe import child_env, configure_console, safe_text  # noqa: E402
import build_and_publish  # noqa: E402

CHECK = "\u2713"        # ✓  – has an entry in the glyph fallback table
SQRT = "\u221a"         # √  – has NO fallback entry: only a Unicode-native stream can hold it
VIET = "Việt Nam"


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


# ================================================== PROMPT-029W: the fallback must be STREAM-AWARE
def _host_locale(monkeypatch, name: str):
    """Pin ``locale.getpreferredencoding()`` so no expectation depends on the machine running pytest (§7)."""
    monkeypatch.setattr(console_safe.locale, "getpreferredencoding", lambda *a, **k: name)
    assert console_safe._encoding_for(None) == name


def test_unicode_native_stream_keeps_unicode_even_when_the_host_locale_is_cp1252(monkeypatch):
    """THE reported Windows regression, reproduced deterministically on any host.

    ``io.StringIO`` declares ``encoding is None`` because it stores ``str``, not bytes.  Reading that as
    "ask the locale" made ``safe_text()`` fall back to ``"[OK]"`` on Windows — whose preferred encoding is
    cp1252 — while the very same call passed on Linux.  The supplied stream is authoritative (§2) and a
    Unicode text stream keeps Unicode (§3).
    """
    _host_locale(monkeypatch, "cp1252")
    assert safe_text(CHECK, stream=io.StringIO()) == CHECK
    assert safe_text(SQRT, stream=io.StringIO()) == SQRT
    assert safe_text("Build 017", stream=io.StringIO()) == "Build 017", "plain ASCII is never rewritten"
    assert safe_text(f"{CHECK} {VIET}", stream=io.StringIO()) == f"{CHECK} {VIET}"
    _host_locale(monkeypatch, "utf-8")
    assert safe_text(CHECK, stream=io.StringIO()) == CHECK, "and identical on a UTF-8 host"


def test_explicit_stream_wins_over_stdout_and_the_locale(monkeypatch):
    """§2: safety is decided from the supplied stream, never from sys.stdout or a global code page."""
    _host_locale(monkeypatch, "cp1252")
    monkeypatch.setattr(sys, "stdout", StrictCp1252())
    assert safe_text(CHECK, stream=io.StringIO()) == CHECK, "the explicit stream is authoritative"
    assert safe_text(CHECK) == "[OK]", "with no explicit stream the console/locale answer still applies"


def test_strict_cp1252_stream_never_raises_and_falls_back_to_ascii():
    """§4/§5: the legacy-console guarantee is untouched by the stream-aware contract."""
    stream = StrictCp1252()
    result = safe_text(f"{CHECK} Hoàn thành — {VIET}", stream=stream)
    stream.write(result)                                     # raises UnicodeEncodeError if it were unsafe
    assert result.encode("ascii", "strict"), "the fallback must be pure ASCII"
    assert "[OK]" in result and "Hoan thanh" in result and "Viet Nam" in result
    assert CHECK not in result and "—" not in result and "ệ" not in result


def test_utf8_declared_stream_preserves_unicode(monkeypatch):
    """§4: a stream that DECLARES an encoding able to hold the text keeps it exactly."""
    _host_locale(monkeypatch, "cp1252")                       # must not leak into this decision
    stream = io.TextIOWrapper(io.BytesIO(), encoding="utf-8", errors="strict")
    text = f"{CHECK} {SQRT} {VIET}"
    assert stream.encoding == "utf-8"
    assert safe_text(text, stream=stream) == text
    stream.write(text)
    stream.flush()


def test_stream_none_console_path_is_unchanged(monkeypatch):
    """§6: the stream-less / current-console behaviour must be exactly what it was before."""
    _host_locale(monkeypatch, "cp1252")
    assert console_safe._encoding_for(None) == "cp1252"
    assert safe_text(f"{CHECK} {VIET}") == "[OK] Viet Nam"
    _host_locale(monkeypatch, "utf-8")
    assert safe_text(f"{CHECK} {VIET}") == f"{CHECK} {VIET}"


def test_safe_text_is_not_an_unconditional_passthrough():
    """§5: the semantic fallback table is intact — a restricted encoding still maps glyphs."""
    assert safe_text(CHECK, encoding="cp1252") == "[OK]"
    assert safe_text("✔", encoding="cp1252") == "[OK]"
    assert safe_text("✗", encoding="cp1252") == "[FAIL]"
    assert safe_text("❌", encoding="cp1252") == "[FAIL]"
    assert safe_text("⚠", encoding="cp1252") == "[WARN]"
    assert safe_text("○", encoding="cp1252") == "o"
    assert safe_text(SQRT, encoding="ascii") == "?", "no table entry -> replaced, never raised"
    assert safe_text("Build 017", encoding="cp1252") == "Build 017"


def test_glyphs_a_legacy_console_can_already_show_are_not_needlessly_replaced():
    """The fallback fires only when the destination genuinely cannot represent the text.

    cp1252 natively encodes the bullet, en/em dash, ellipsis and curly quotes, so on a legacy console those
    must survive unchanged — replacing them would lose information for no safety gain.  This is also why the
    stream-aware fix cannot be "always transliterate": the decision belongs to the destination.
    """
    for ch in ("•", "–", "—", "…", "‘", "’", "‚", "“", "”", "„"):
        ch.encode("cp1252", "strict")                       # precondition: representable in cp1252
        assert safe_text(ch, encoding="cp1252") == ch, ch
    assert safe_text("• – — …", stream=StrictCp1252()) == "• – — …"


def test_console_proxy_keeps_legacy_safety_and_unicode_fidelity(monkeypatch):
    """§5 end-to-end through configure_console(): cp1252 stays safe, Unicode-native stays faithful."""
    legacy = StrictCp1252()
    monkeypatch.setattr(sys, "stdout", legacy)
    configure_console()
    print(f"{CHECK} Hoàn thành ở {VIET}")
    assert "[OK]" in legacy.getvalue() and "Hoan thanh o" in legacy.getvalue()

    native = io.StringIO()
    monkeypatch.setattr(sys, "stdout", native)
    configure_console()
    print(f"{CHECK} Hoàn thành ở {VIET}")
    assert f"{CHECK} Hoàn thành ở {VIET}" in native.getvalue()
    assert sys.stdout.encoding is None, "the proxy must stay faithful to a Unicode-native stream"
