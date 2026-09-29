# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller spec – onedir portable build.

Result:  dist/ReportExtractor_Portable/ReportExtractor.exe  (+ _internal/)
"""
import os
import sys
from PyInstaller.utils.hooks import collect_data_files, collect_submodules

block_cipher = None
ROOT = os.path.abspath(os.path.dirname(SPEC))

datas = [(os.path.join(ROOT, "assets"), "assets")]
hiddenimports = ["app", "app.main", "app.gui", "app.config", "app.pptx_parser", "app.ollama_client",
                 "app.classifier", "app.extractor", "app.qpn_renderer", "app.image_extractor",
                 "app.excel_writer", "app.batch_processor", "app.logger", "app.scanner",
                 "app.history", "app.diagnostics",
                 "PIL._tkinter_finder", "openpyxl.cell._writer", "lxml._elementpath"]
hiddenimports += collect_submodules("pptx")
hiddenimports += collect_submodules("openpyxl")

# optional extras – only when installed on the build machine
for mod in ("tkinterdnd2", "fitz", "pymupdf"):
    try:
        __import__(mod)
        datas += collect_data_files(mod)
        hiddenimports.append(mod)
    except Exception:
        pass
if sys.platform == "win32":
    hiddenimports += ["win32com", "win32com.client", "pythoncom", "pywintypes", "win32api", "winreg"]
datas += collect_data_files("pptx")      # default template parts used by python-pptx

a = Analysis(
    [os.path.join(ROOT, "run.py")],
    pathex=[ROOT],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=["matplotlib", "numpy", "scipy", "pandas", "IPython", "notebook", "PyQt5", "PySide2"],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="ReportExtractor",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,          # GUI app – no console window
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=os.path.join(ROOT, "assets", "app.ico") if os.path.exists(os.path.join(ROOT, "assets", "app.ico")) else None,
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name="ReportExtractor_Portable",
)
