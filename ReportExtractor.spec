# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller spec – Windows Portable build (PROMPT-003).

    onedir + windowed  (no console window, no onefile, no UPX, no packers)
    dist/ReportExtractor_v<version>_Portable/ReportExtractor.exe  (+ _internal/)

Version and folder name come from app/__init__.py (single source of truth).  Runtime state (config/, logs/,
Output/) is created next to the executable by app/runtime_paths.py – nothing is written into _internal/.
Run through tools/build_portable.py (build_portable.bat) which also assembles docs, ZIP and SHA256SUMS.
"""
import importlib.util
import os
import sys

from PyInstaller.utils.hooks import collect_data_files, collect_submodules

ROOT = os.path.abspath(os.path.dirname(SPEC))
sys.path.insert(0, ROOT)

_spec = importlib.util.spec_from_file_location("app_version", os.path.join(ROOT, "app", "__init__.py"))
_mod = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_mod)
VERSION = _mod.__version__
PORTABLE_NAME = f"ReportExtractor_v{VERSION}_Portable"

datas = [(os.path.join(ROOT, "assets"), "assets")]
datas += collect_data_files("pptx")          # default template parts used by python-pptx
hiddenimports = [
    "app", "app.main", "app.gui", "app.gui_controller", "app.config", "app.runtime_paths", "app.pptx_parser",
    "app.ollama_client", "app.ollama_discovery", "app.classifier", "app.extractor", "app.qpn_renderer",
    "app.image_extractor", "app.excel_writer", "app.batch_processor", "app.logger", "app.scanner", "app.history",
    "app.diagnostics", "app.prescan",
    "PIL._tkinter_finder", "PIL.ImageTk", "openpyxl.cell._writer", "lxml._elementpath", "requests", "urllib3",
    "charset_normalizer", "idna", "certifi",
]
hiddenimports += collect_submodules("pptx")
hiddenimports += collect_submodules("openpyxl")
hiddenimports += collect_submodules("app")

# optional extras – only when installed on the build machine (drag & drop, PDF rasterizer)
for mod in ("tkinterdnd2", "fitz", "pymupdf"):
    try:
        __import__(mod)
        datas += collect_data_files(mod)
        hiddenimports.append(mod)
    except Exception:
        pass
if sys.platform == "win32":
    hiddenimports += ["win32com", "win32com.client", "pythoncom", "pywintypes", "win32api", "winreg"]

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
    excludes=["matplotlib", "numpy", "scipy", "pandas", "IPython", "notebook", "PyQt5", "PySide2", "pytest",
              "tests"],
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
