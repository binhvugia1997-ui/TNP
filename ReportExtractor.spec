# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller spec – Windows Portable build (PROMPT-003, adapted to React + pywebview by PROMPT-028).

    onedir + windowed  (no console window, no onefile, no UPX, no packers)
    dist/ReportExtractor_v<version>_Portable/ReportExtractor.exe  (+ _internal/)

The executable launches the SAME React application as ``python -m app.desktop`` (§1): ``run.py`` →
``app.main.main()`` → ``launch_ui()`` → ``app.desktop.main()`` (pywebview + EdgeChromium/WebView2).

Bundled on purpose:
  * ``frontend/dist``          – the built React UI (index.html + assets/*), placed at ``_internal/frontend/dist``
                                 so ``app.desktop.frontend_url()`` resolves it without any repository path (§3/§16).
                                 The builder ALSO copies it next to ``_internal`` for the documented Portable
                                 layout; ``tools/build_portable.py`` verifies the two copies are identical.
  * ``webview``                – pywebview itself PLUS its data files: ``webview/js/*`` (the JS↔Python bridge
                                 injection – without it §18 fails) and ``webview/lib/*.dll`` +
                                 ``lib/runtimes/win-x64/native/WebView2Loader.dll`` (the WebView2 interop
                                 assemblies that keep Microsoft Edge WebView2 as the UI backend, §6).
  * ``pythonnet`` / ``clr`` / ``clr_loader`` – the .NET bridge the EdgeChromium backend imports.

Optional by design:
  * PowerPoint (``win32com``/``pywin32``) – collected only when importable on the build machine.  A target PC
    WITHOUT PowerPoint still starts: ``app.qpn_renderer`` probes COM lazily and falls back to LibreOffice and
    then the built-in renderer (§7).
  * ``tkinterdnd2`` / ``fitz`` – drag & drop and PDF rasterisation extras, only when installed.

Version and folder name come from app/__init__.py (single source of truth).  Runtime state (config/, logs/,
Output/) is created next to the executable by app/runtime_paths.py – nothing is written into _internal/.
Run through tools/build_portable.py (build_portable.bat) which also builds the frontend, assembles docs,
ZIP and SHA256SUMS.
"""
import importlib.util
import os
import sys

from PyInstaller.utils.hooks import collect_data_files, collect_dynamic_libs, collect_submodules

ROOT = os.path.abspath(os.path.dirname(SPEC))
sys.path.insert(0, ROOT)

_spec = importlib.util.spec_from_file_location("app_version", os.path.join(ROOT, "app", "__init__.py"))
_mod = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_mod)
VERSION = _mod.__version__
PORTABLE_NAME = f"ReportExtractor_v{VERSION}_Portable"

FRONTEND_DIST = os.path.join(ROOT, "frontend", "dist")
FRONTEND_INDEX = os.path.join(FRONTEND_DIST, "index.html")


def _module_available(name: str) -> bool:
    try:
        return importlib.util.find_spec(name) is not None
    except Exception:                                  # noqa: BLE001 – an optional extra must never fail the build
        return False


# --------------------------------------------------------------------------- frontend (React) bundle
if not os.path.isfile(FRONTEND_INDEX):
    raise SystemExit(
        "\n[spec] LỖI: thiếu frontend/dist/index.html – không thể đóng gói UI React.\n"
        "       Chạy TRƯỚC khi build:  cd frontend && npm run build\n")
if not os.path.isdir(os.path.join(FRONTEND_DIST, "assets")):
    raise SystemExit("\n[spec] LỖI: thiếu frontend/dist/assets – bản build React không đầy đủ.\n")

datas = [(os.path.join(ROOT, "assets"), "assets"),
         (FRONTEND_DIST, os.path.join("frontend", "dist"))]      # → _internal/frontend/dist (§3)
datas += collect_data_files("pptx")          # default template parts used by python-pptx

hiddenimports = [
    "app", "app.main", "app.desktop", "app.gui", "app.gui_controller", "app.config", "app.runtime_paths",
    "app.pptx_parser", "app.ollama_client", "app.ollama_discovery", "app.classifier", "app.extractor",
    "app.qpn_renderer", "app.image_extractor", "app.excel_writer", "app.batch_processor", "app.logger",
    "app.scanner", "app.history", "app.diagnostics", "app.prescan", "app.application_service", "app.bridge",
    "app.improvement_visual", "app.improvement_pictures", "app.image_learning", "app.preview_geometry",
    "app.report_identity", "app.updater",
    # pywebview runtime (§5/§6): the loopback HTTP server, WSGI layer and JS bridge injection
    "webview", "bottle", "proxy_tools", "typing_extensions", "webview.platforms.edgechromium",
    "openpyxl.cell._writer", "lxml._elementpath", "requests", "urllib3",
    "charset_normalizer", "idna", "certifi",
]
hiddenimports += collect_submodules("pptx")
hiddenimports += collect_submodules("openpyxl")
hiddenimports += collect_submodules("app")

# pywebview ships its JS bridge + the WebView2 interop DLLs as DATA files; without them the packaged app
# opens but window.pywebview.api never appears (§18).  Microsoft Edge WebView2 stays the Windows backend (§6).
for _name in ("webview", "pythonnet", "clr", "clr_loader"):
    if _module_available(_name):
        datas += collect_data_files(_name)
        try:
            for _lib in collect_dynamic_libs(_name):
                datas += [(_lib[0], _lib[1])]
        except Exception as _exc:                      # noqa: BLE001
            print(f"[spec] WARNING: dynamic libs for {_name}: {_exc}")
        hiddenimports.append(_name)
        hiddenimports += collect_submodules(_name)
        print(f"[spec] bundled pywebview runtime component: {_name}")
    else:
        print(f"[spec] pywebview runtime component NOT installed on this build machine: {_name}")
if sys.platform == "win32" and not _module_available("webview"):
    raise SystemExit("\n[spec] LỖI: pywebview chưa được cài trong môi trường build.\n"
                     "       Cài:  pip install -r requirements-webview.txt\n")

# --------------------------------------------------------------------------- optional extras
# Classic Tk GUI pieces – still bundled so the fallback UI and the startup error dialog work.
if _module_available("tkinter"):
    hiddenimports += ["PIL._tkinter_finder", "PIL.ImageTk"]

# PowerPoint / pywin32 – OPTIONAL (§7).  Present on a build machine with pywin32; a target PC without
# PowerPoint still launches and uses LibreOffice / the built-in renderer.
if sys.platform == "win32":
    # pywin32's own DLLs (pythoncom*.dll, pywintypes*.dll) are picked up by PyInstaller's bundled hooks;
    # only the import names need declaring, and only when pywin32 is actually installed.
    for _name in ("win32com", "win32com.client", "pythoncom", "pywintypes", "win32api", "winreg"):
        if _module_available(_name.split(".")[0]):
            hiddenimports.append(_name)

for mod in ("tkinterdnd2", "fitz", "pymupdf"):
    try:
        __import__(mod)
        datas += collect_data_files(mod)
        hiddenimports.append(mod)
    except Exception:
        pass

ICON = os.path.join(ROOT, "assets", "app.ico")

a = Analysis(
    [os.path.join(ROOT, "run.py")],
    pathex=[ROOT],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    # Alternative pywebview backends are deliberately excluded so the frozen app always takes the
    # Microsoft Edge WebView2 path on Windows (§6); numpy/pandas/etc. keep the bundle small.
    excludes=["matplotlib", "numpy", "scipy", "pandas", "IPython", "notebook", "pytest", "tests",
              "PyQt5", "PySide2", "PyQt6", "PySide6", "gi", "gtk", "cefpython3"],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure, a.zipped_data)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,            # onedir
    name="ReportExtractor",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,                        # no packers (anti-virus / SmartScreen friendliness)
    console=False,                    # windowed GUI app
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=ICON if os.path.exists(ICON) else None,
    version=None,
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name=PORTABLE_NAME,
)
