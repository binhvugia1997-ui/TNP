"""Launcher used by RUN_DEV.bat and by PyInstaller (ReportExtractor.spec)."""
import multiprocessing
import sys
import os

if getattr(sys, "frozen", False):
    # make sure Windows console code page issues never break Vietnamese output
    os.environ.setdefault("PYTHONIOENCODING", "utf-8")

from app.main import main  # noqa: E402

if __name__ == "__main__":
    multiprocessing.freeze_support()
    sys.exit(main())
