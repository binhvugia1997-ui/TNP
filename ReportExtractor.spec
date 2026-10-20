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
  * ``webview/js``             – the JS↔Python bridge injection (api.js / dom_json.js / state.js).  This is the
                                 ONE pywebview resource collected manually, because the official hook-webview
                                 uses ``subdir='lib'`` and therefore misses it; without it §18 fails.
  * ``clr`` / ``clr_loader`` / ``pythonnet`` – registered as hiddenimports ONLY, so the official
                                 pyinstaller-hooks-contrib hooks are their single owner (PROMPT-028R).
                                 hook-clr collects ``pythonnet/runtime/Python.Runtime.dll`` (as a BINARY on
                                 Windows, as DATA elsewhere) and hook-clr_loader collects
                                 ``clr_loader/ffi/dlls/*/ClrLoader.dll``.  Microsoft Edge WebView2 stays the
                                 Windows backend (§6): webview.platforms.winforms hosts it via edgechromium +
                                 the WebView2 assemblies that hook-webview collects.

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

# collect_dynamic_libs is deliberately NOT imported: PROMPT-028R leaves every dynamic library
# (Python.Runtime.dll, ClrLoader.dll, the WebView2 assemblies) to the official PyInstaller hooks.
from PyInstaller.utils.hooks import collect_data_files, collect_submodules

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
    "openpyxl.cell._writer", "lxml._elementpath", "requests", "urllib3",
    "charset_normalizer", "idna", "certifi",
]
hiddenimports += collect_submodules("pptx")
hiddenimports += collect_submodules("openpyxl")
hiddenimports += collect_submodules("app")

# --------------------------------------------------------------------------- pywebview / .NET runtime
# PROMPT-028R — the .NET side is owned ENTIRELY by the official pyinstaller-hooks-contrib hooks.
#
#   hook-clr        →  Windows: `binaries = [(…/pythonnet/runtime/Python.Runtime.dll, 'pythonnet/runtime')]`
#                      Its own comment: "On Windows, collect runtime DLL file(s) as binaries; on other OSes,
#                      collect them as data files, to prevent fatal errors in binary dependency analysis."
#                      It also adds hiddenimports ["platform", "warnings"], which Python.Runtime.dll imports.
#   hook-clr_loader →  Windows: `binaries = collect_dynamic_libs("clr_loader")`, i.e.
#                      clr_loader/ffi/dlls/{amd64,x86}/ClrLoader.dll — the native netfx host that
#                      clr_loader.ffi.load_netfx() dlopen()s from `Path(__file__).parent / "dlls" / arch`.
#
# `clr` is listed explicitly so the hook is guaranteed to run even if a future pywebview release moves the
# `import clr` in webview/platforms/winforms.py behind importlib.
#
# DO NOT add collect_data_files()/collect_dynamic_libs() for pythonnet, clr or clr_loader here.  Measured
# against PyInstaller 6.22.3 / hooks-contrib 2026.8, such a call is NOT what breaks the CLR — Analysis ends
# with `normalize_toc(self.datas + self.binaries)` (building/build_main.py) and normalize_toc's
# _TOC_TYPE_PRIORITIES gives BINARY/EXTENSION priority 1 over DATA priority 0, so a same-destination `datas`
# entry is simply discarded in favour of hook-clr's `binaries` entry.  It is removed anyway, because a single
# owner is what makes tools/build_portable.py's validate_pythonnet_runtime() check meaningful: exactly ONE
# Python.Runtime.dll, at exactly the path pythonnet computes, byte-identical to the build venv's copy.
#
# The failure this file must not reproduce is
#     RuntimeError: Failed to resolve Python.Runtime.Loader.Initialize
#                   from _internal\pythonnet\runtime\Python.Runtime.dll
# raised by clr_loader/netfx.py:50 whenever pyclr_get_function() returns NULL.  That message reports the path
# pythonnet ASKED for, not a path that exists, so it is equally consistent with the DLL being absent, being a
# different build, or being unloadable (Python.Runtime.dll targets .NETStandard 2.0 and therefore needs the
# netstandard facade of .NET Framework 4.7.2+).  hook-clr's own legacy fallback is the concrete way the path
# ends up empty: when importlib.metadata.files('pythonnet') does not yield exactly one match it calls
# ctypes.util.find_library('Python.Runtime') and collects the result with destination '.', i.e.
# _internal/Python.Runtime.dll — one directory away from where pythonnet looks.  The builder gate catches it.
for _net_module in ("clr", "clr_loader", "pythonnet"):
    if _module_available(_net_module):
        hiddenimports.append(_net_module)
        print(f"[spec] .NET runtime module registered for the official hook: {_net_module}")
    else:
        print(f"[spec] WARNING: {_net_module} not installed – its PyInstaller hook will NOT run")
if sys.platform == "win32":
    # pywebview on Windows needs pythonnet; without it the app cannot create a window at all.
    for _required in ("clr", "pythonnet", "clr_loader"):
        if not _module_available(_required):
            raise SystemExit(f"\n[spec] LỖI: thiếu {_required} trong môi trường build (pip install "
                             f"-r requirements-webview.txt). Không có nó hook-clr không chạy và "
                             f"Python.Runtime.dll sẽ không được đóng gói đúng cách.\n")
hiddenimports += ["cffi"]                      # clr_loader.ffi dlopen()s ClrLoader.dll through cffi

# clr_loader/__init__.py does NOT import its runtime backends at module level: get_netfx() body starts with
# `from .netfx import NetFx` (line 183 of clr_loader 0.2.10).  Bytecode analysis does follow imports inside
# function bodies, but this is the exact code path that failed on Windows, so the whole netfx chain is
# declared explicitly instead of relying on that.  These are MODULE registrations only — they cannot collide
# with hook-clr_loader's `binaries`, which is what the removed collect_*() calls were.
if _module_available("clr_loader"):
    hiddenimports += ["clr_loader.ffi", "clr_loader.netfx", "clr_loader.types", "clr_loader.util"]
    print("[spec] clr_loader netfx chain registered explicitly (get_netfx imports it lazily)")

# pywebview's OWN hook (hook-webview) collects, on Windows only:
#     datas    = collect_data_files('webview', subdir='lib')    → the WebView2 interop assemblies
#     binaries = collect_dynamic_libs('webview')                → the same DLLs, as binaries
# It uses subdir='lib', so it does NOT collect webview/js/* — and those files ARE the JS↔Python bridge
# injection (api.js / dom_json.js / state.js).  Without them the window opens but window.pywebview.api
# never appears.  This is the ONE pywebview collection that is proven necessary to add manually, and it is
# deliberately restricted to `js` so it can never collide with hook-webview's handling of webview/lib.
if _module_available("webview"):
    _js_files = collect_data_files("webview", subdir="js")
    if not _js_files:
        raise SystemExit("\n[spec] LỖI: không thu thập được webview/js – cầu nối JS↔Python sẽ không hoạt động.\n")
    datas += _js_files
    print(f"[spec] webview JS bridge files collected: {len(_js_files)}")
    hiddenimports += [
        "webview", "bottle", "proxy_tools", "typing_extensions",
        # On Windows the backend is winforms; winforms.py then does `from . import edgechromium` and sets
        # renderer='edgechromium', which hosts Microsoft Edge WebView2 (CoreWebView2 + WebView2Loader.dll).
        # webview/guilib.py imports both from inside nested functions, which PyInstaller's bytecode analysis
        # does follow, so these are explicit-by-intent rather than strictly required; they keep the backend
        # selection pinned even if guilib.py switches to importlib.import_module().
        "webview.platforms.winforms", "webview.platforms.edgechromium",
    ]
else:
    if sys.platform == "win32":
        raise SystemExit("\n[spec] LỖI: pywebview chưa được cài trong môi trường build.\n"
                         "       Cài:  pip install -r requirements-webview.txt\n")
    print("[spec] pywebview not installed – skipping its collection (non-Windows experiment)")

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
    # webview/guilib.py references every backend inside nested try/except importers, so analysis would
    # otherwise pull in Qt/GTK/Cocoa/CEF.  Windows only ever uses winforms -> edgechromium (WebView2), and
    # each excluded backend degrades gracefully through pywebview's own ImportError handling.
    excludes=["matplotlib", "numpy", "scipy", "pandas", "IPython", "notebook", "pytest", "tests",
              "PyQt5", "PySide2", "PyQt6", "PySide6", "gi", "gtk", "cefpython3",
              "webview.platforms.qt", "webview.platforms.gtk", "webview.platforms.cocoa",
              "webview.platforms.android", "webview.platforms.mshtml"],
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
