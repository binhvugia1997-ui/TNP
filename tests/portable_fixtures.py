"""Canonical synthetic Portable runtime layout for updater tests.

These deliberately non-loadable DLL payloads exercise the updater's structural path checks only. They are not
runtime binaries and must never be used for a build or a real Portable installation; PE/assembly authenticity is
covered by the packaging tests in ``test_prompt028_portable.py``.
"""
from pathlib import Path, PurePosixPath
from typing import Callable


PYTHON_RUNTIME_PATH = "_internal/pythonnet/runtime/Python.Runtime.dll"
CLRLOADER_PATHS = (
    "_internal/clr_loader/ffi/dlls/amd64/ClrLoader.dll",
    "_internal/clr_loader/ffi/dlls/x86/ClrLoader.dll",
)
WEBVIEW_CORE_PATH = "_internal/webview/lib/Microsoft.Web.WebView2.Core.dll"
WEBVIEW_FORMS_PATH = "_internal/webview/lib/Microsoft.Web.WebView2.WinForms.dll"
WEBVIEW_LOADER_PATH = "_internal/webview/lib/runtimes/win-x64/native/WebView2Loader.dll"

SYNTHETIC_WINDOWS_RUNTIME_FILES = {
    PYTHON_RUNTIME_PATH: b"synthetic Python.Runtime.dll placeholder; structural tests only",
    # The real hook packages both architecture variants; the updater's layout check requires at least one.
    CLRLOADER_PATHS[0]: b"synthetic amd64 ClrLoader.dll placeholder; structural tests only",
    CLRLOADER_PATHS[1]: b"synthetic x86 ClrLoader.dll placeholder; structural tests only",
    WEBVIEW_CORE_PATH: b"synthetic WebView2 Core placeholder; structural tests only",
    WEBVIEW_FORMS_PATH: b"synthetic WebView2 WinForms placeholder; structural tests only",
    WEBVIEW_LOADER_PATH: b"synthetic WebView2Loader placeholder; structural tests only",
}


def _populate_windows_runtime(write_file: Callable[[str, bytes], None]) -> None:
    for relative_path, payload in SYNTHETIC_WINDOWS_RUNTIME_FILES.items():
        write_file(relative_path, payload)


def add_windows_runtime_to_directory(root: Path) -> None:
    """Write the same structural Windows runtime paths into a synthetic Portable directory."""
    root = Path(root)

    def write_file(relative_path: str, payload: bytes) -> None:
        target = root.joinpath(*PurePosixPath(relative_path).parts)
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_bytes(payload)

    _populate_windows_runtime(write_file)


def add_windows_runtime_to_zip(archive, prefix: str = "") -> None:
    """Write the same structural Windows runtime paths into a synthetic Portable ZIP."""
    _populate_windows_runtime(lambda relative_path, payload: archive.writestr(prefix + relative_path, payload))
