"""Shared console and child-Python encoding safeguards for Windows build tooling.

Windows consoles may expose a strict cp1252 stream even when build tools emit UTF-8.  The
helpers in this module make the parent process safe and pass an explicit UTF-8 policy to
Python subprocesses; they do not rely on ``chcp 65001``.
"""
from __future__ import annotations

import locale
import os
import sys
import unicodedata
from typing import Mapping, Optional, TextIO

_GLYPH_FALLBACKS = str.maketrans({
    "✓": "[OK]", "✔": "[OK]", "✗": "[FAIL]", "❌": "[FAIL]", "⚠": "[WARN]",
    "•": "*", "○": "o", "–": "-", "—": "-", "…": "...",
    "‘": "'", "’": "'", "‚": "'", "“": '"', "”": '"', "„": '"',
})


def _stream_encoding(stream) -> Optional[str]:
    """The byte encoding ``stream`` declares, or ``None`` when it declares none.

    ``None`` is a MEANINGFUL answer, not missing information: ``io.StringIO`` -- and any ``TextIOBase`` that
    stores ``str`` instead of bytes -- reports ``encoding is None`` precisely because it has no byte encoding
    at all.  Such a stream is Unicode-native and can hold every character, so it must never be judged by an
    unrelated global setting (PROMPT-029W §3).  Reading ``encoding is None`` as "ask the locale" is what made
    ``safe_text("✓", stream=StringIO())`` return ``"[OK]"`` on a Windows host, whose
    ``locale.getpreferredencoding()`` is cp1252.
    """
    if stream is None:
        return None
    declared = getattr(stream, "encoding", None)
    return str(declared) if declared else None


def _encoding_for(stream=None, encoding: Optional[str] = None) -> str:
    """Best-effort encoding NAME, for the current-console path where a ``str`` is required.

    Behaviour is unchanged when no stream is supplied (§6): that call means "the console I am writing to
    now", so the locale's preferred encoding is still the right answer there.
    """
    value = encoding or _stream_encoding(stream)
    if value:
        return value
    try:
        return locale.getpreferredencoding(False) or "utf-8"
    except Exception:  # noqa: BLE001 -- console configuration must be best effort
        return "utf-8"


def safe_text(text, stream=None, encoding: Optional[str] = None) -> str:
    """Return text that can be written to ``stream`` without encoding errors.

    STREAM-AWARE contract (PROMPT-029W §2): an explicitly supplied ``stream`` is authoritative for that call
    -- safety is decided from THAT stream, never from ``sys.stdout``, the Windows console code page, the
    locale's preferred encoding or any other global.

      * the stream declares no byte encoding (``io.StringIO`` and friends) -> it is Unicode-native, so the
        original text is returned EXACTLY (§3);
      * the stream declares an encoding (a legacy cp1252 console) and the text is not representable in it ->
        the existing fallback applies: common UI glyphs become readable ASCII tokens, then accents are
        transliterated and any remaining unrepresentable characters are replaced rather than raising (§4/§5).

    Unicode-capable streams retain the original text. For restricted streams, common UI glyphs become readable
    ASCII tokens, then accents are transliterated and any remaining unrepresentable characters are replaced
    rather than raising.
    """
    value = str(text)
    if encoding is None and stream is not None and _stream_encoding(stream) is None:
        # Unicode-native destination: nothing to fall back from.  This is NOT an unconditional pass-through --
        # every stream that declares an encoding still goes through the representability check below, so
        # cp1252 console output stays safe and status glyphs keep their ASCII fallbacks.
        return value
    codec = _encoding_for(stream, encoding)
    try:
        value.encode(codec, errors="strict")
        return value
    except (LookupError, UnicodeEncodeError):
        value = value.translate(_GLYPH_FALLBACKS).replace("đ", "d").replace("Đ", "D")
        decomposed = unicodedata.normalize("NFKD", value)
        value = "".join(ch for ch in decomposed if not unicodedata.combining(ch))
        try:
            return value.encode(codec, errors="replace").decode(codec, errors="replace")
        except LookupError:  # invalid/custom encoding name
            return value.encode("ascii", errors="replace").decode("ascii")


class _SafeTextStream:
    """Fallback proxy for text streams without a usable ``reconfigure`` method."""
    _arena_console_safe = True

    def __init__(self, stream: TextIO):
        self._stream = stream
        # Faithful to the wrapped stream: None for a Unicode-native one, so the proxy does not re-inject the
        # locale's encoding and mangle text the destination could have held (PROMPT-029W §2/§3).
        self.encoding = _stream_encoding(stream)
        self.errors = "replace"

    def write(self, text):
        value = safe_text(text, stream=self._stream, encoding=self.encoding)
        try:
            return self._stream.write(value)
        except UnicodeEncodeError:  # unusual stream whose declared encoding is inaccurate
            fallback = safe_text(value, encoding="ascii")
            return self._stream.write(fallback)

    def flush(self):
        return self._stream.flush()

    def isatty(self):
        try:
            return self._stream.isatty()
        except Exception:  # noqa: BLE001
            return False

    def fileno(self):
        return self._stream.fileno()

    def writable(self):
        try:
            return self._stream.writable()
        except Exception:  # noqa: BLE001
            return True

    def __getattr__(self, name):
        return getattr(self._stream, name)


def _configure_stream(stream):
    if stream is None or getattr(stream, "_arena_console_safe", False):
        return stream
    reconfigure = getattr(stream, "reconfigure", None)
    if callable(reconfigure):
        try:
            reconfigure(encoding="utf-8", errors="replace")
            try:
                stream._arena_console_safe = True
            except Exception:  # noqa: BLE001 -- some TextIOWrapper implementations are immutable
                pass
            return stream
        except Exception:  # noqa: BLE001 -- fall back to a safe proxy
            pass
    return _SafeTextStream(stream)


def configure_console() -> None:
    """Make ``sys.stdout`` and ``sys.stderr`` safe; idempotent and best effort."""
    try:
        sys.stdout = _configure_stream(sys.stdout)
    except Exception:  # noqa: BLE001
        pass
    try:
        sys.stderr = _configure_stream(sys.stderr)
    except Exception:  # noqa: BLE001
        pass


def child_env(env: Optional[Mapping[str, str]] = None) -> dict:
    """Copy an environment and force UTF-8-with-replacement for child Python I/O."""
    child = dict(os.environ if env is None else env)
    child["PYTHONUTF8"] = "1"
    child["PYTHONIOENCODING"] = "utf-8:replace"
    return child


__all__ = ["configure_console", "child_env", "safe_text"]
