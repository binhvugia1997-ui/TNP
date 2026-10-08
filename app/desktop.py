"""Run the React frontend inside a local pywebview desktop window.

This is the entry point the Windows Portable ``ReportExtractor.exe`` launches (PROMPT-028): the frozen
executable must open the SAME React application as ``python -m app.desktop``.

Frontend location is resolved for BOTH runtimes, because pywebview resolves a *relative* window URL
against its own application root, which differs between the two:

* source mode  – ``webview.util.get_app_root()`` == ``dirname(realpath(sys.argv[0]))`` == ``<repo>/app``,
  so the bundle is addressed as ``../frontend/dist/index.html``.
* frozen mode  – ``get_app_root()`` == ``sys._MEIPASS`` == ``<portable>/_internal``, so the very same
  file is addressed as ``frontend/dist/index.html`` (or ``../frontend/dist/index.html`` when the builder
  placed the bundle next to ``_internal``).

The relative form is required on purpose: pywebview serves it over a loopback HTTP server rooted at the
dist folder.  An absolute ``file://`` URL would break Chromium ES-module loading of ``assets/*``.
Nothing here depends on the current working directory or on the repository path (§16).
"""
from __future__ import annotations

import argparse
import logging
import os
import sys
from pathlib import Path
from typing import List, Optional

from . import APP_TITLE
from .application_service import ApplicationService
from .bridge import BridgeService
from .runtime_paths import is_packaged, portable_root, resource_root

LOG = logging.getLogger("report_extractor.webview.desktop")

#: Where the Vite bundle lives, relative to a root.  Kept as ONE definition so the packaged layout and
#: the source layout can never drift apart.
FRONTEND_DIST_PARTS = ("frontend", "dist")
FRONTEND_INDEX_NAME = "index.html"
#: pywebview resolves a relative window URL against its own app root; see the module docstring.
SOURCE_URL_RELATIVE_TO_APP_DIR = "../frontend/dist/index.html"


def assumed_app_root() -> Path:
    """The directory pywebview resolves relative window URLs against, computed deterministically.

    pywebview's own ``get_app_root()`` returns ``sys._MEIPASS`` when frozen and
    ``dirname(realpath(sys.argv[0]))`` otherwise.  For the two launch modes that matter those are:

    * ``ReportExtractor.exe``            → ``<portable>/_internal``   (``sys._MEIPASS``)
    * ``python -m app.desktop``          → ``<repo>/app``             (``sys.argv[0]`` is that module)

    ``sys.argv[0]`` is deliberately NOT read here: under pytest, an IDE runner or ``python -c`` it points at
    the launcher rather than the app, which would make this value (and the module constant below) depend on
    how the process happened to start.  :func:`runtime_app_root` re-checks against pywebview itself.
    """
    if is_packaged():
        meipass = getattr(sys, "_MEIPASS", None)
        if meipass:
            return Path(meipass)
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def runtime_app_root() -> Path:
    """pywebview's REAL app root, used when actually creating the window.

    Prefers ``webview.util.get_app_root()`` so the URL handed to ``create_window`` can never disagree with
    the one pywebview resolves; falls back to :func:`assumed_app_root` when pywebview is not importable.
    """
    try:
        from webview.util import get_app_root          # noqa: PLC0415 – optional dependency
        return Path(get_app_root())
    except Exception:                                  # noqa: BLE001 – pywebview may be absent (tests, CLI)
        return assumed_app_root()


def frontend_dist_candidates() -> List[Path]:
    """Every location the built React bundle may legitimately occupy, most specific first.

    Packaged builds accept BOTH supported layouts: the bundle next to ``_internal`` (the documented
    Portable layout, easy for a tester to inspect) and the bundle inside ``_internal`` (PyInstaller
    ``datas``).  Source builds resolve to ``<repo>/frontend/dist``.
    """
    seen: List[Path] = []

    def add(path: Optional[Path]) -> None:
        if path is None:
            return
        try:
            resolved = Path(path).resolve()
        except OSError:
            return
        if resolved not in seen:
            seen.append(resolved)

    dist = Path(*FRONTEND_DIST_PARTS)
    if is_packaged():
        # A packaged exe resolves ONLY inside its own folder.  Falling back to a source checkout would
        # hide a broken package behind a developer machine that happens to have `frontend/dist` (§16).
        add(portable_root() / dist)                     # <portable>/frontend/dist  (documented layout)
        add(resource_root() / dist)                     # <portable>/_internal/frontend/dist (spec datas)
        add(resource_root().parent / dist)              # defensive: sibling of the bundle root
        return seen
    add(resource_root() / dist)                         # source mode: <repo>/frontend/dist
    add(Path(__file__).resolve().parent.parent / dist)  # defensive: next to the app/ package
    return seen


def frontend_dist_dir(require: bool = False) -> Optional[Path]:
    """The first candidate that actually contains ``index.html`` (None when the bundle is missing)."""
    for candidate in frontend_dist_candidates():
        if (candidate / FRONTEND_INDEX_NAME).is_file():
            return candidate
    if require:
        searched = "\n  ".join(str(c) for c in frontend_dist_candidates()) or "(none)"
        raise FileNotFoundError(
            "React bundle not found. Run `cd frontend && npm run build` first.\nSearched:\n  " + searched)
    return None


def frontend_url(dist_dir: Optional[Path] = None, root: Optional[Path] = None) -> str:
    """Relative window URL for pywebview, derived from the RESOLVED bundle directory.

    Never hardcoded per runtime: it is ``relpath(dist_dir, root)/index.html``, which yields
    ``../frontend/dist/index.html`` in source mode and ``frontend/dist/index.html`` when the bundle is
    inside ``_internal`` — so assets keep loading over pywebview's loopback HTTP server in both cases.
    ``root`` defaults to :func:`runtime_app_root` — pywebview's OWN root — because that is the directory
    pywebview will resolve the URL against.  Callers that need a value independent of how the process was
    launched (the module constant below, tests) pass :func:`assumed_app_root` explicitly.
    """
    dist_dir = Path(dist_dir) if dist_dir else frontend_dist_dir()
    if dist_dir is None:
        return SOURCE_URL_RELATIVE_TO_APP_DIR            # keep the historical, source-mode-correct value
    try:
        relative = os.path.relpath(dist_dir, Path(root) if root else runtime_app_root())
    except ValueError:                                   # different drives on Windows
        relative = str(dist_dir)
    url = Path(relative).as_posix() + "/" + FRONTEND_INDEX_NAME
    resolved = (Path(root) if root else runtime_app_root()) / url
    if Path(os.path.normpath(str(resolved))) != dist_dir / FRONTEND_INDEX_NAME:
        LOG.warning("WEBVIEW_FRONTEND_URL_MISMATCH root=%s url=%s resolved=%s expected=%s",
                    root, url, resolved, dist_dir / FRONTEND_INDEX_NAME)
    return url


def frontend_index(dist_dir: Optional[Path] = None) -> Optional[Path]:
    resolved = Path(dist_dir) if dist_dir else frontend_dist_dir()
    return (resolved / FRONTEND_INDEX_NAME) if resolved else None


# Import-time convenience values for the SOURCE runtime.  They stay module-level constants because other
# modules/tests reference them, but ``main()`` always re-resolves through the helpers above so a frozen
# executable is never tied to the repository layout.
FRONTEND_DIST = frontend_dist_dir() or (resource_root() / Path(*FRONTEND_DIST_PARTS))
FRONTEND_INDEX = FRONTEND_DIST / FRONTEND_INDEX_NAME
# Computed against assumed_app_root() so this constant is stable no matter how the process was launched
# (pytest, an IDE, `python -c`).  main() re-resolves against pywebview's real root before creating a window.
FRONTEND_URL = frontend_url(FRONTEND_DIST if FRONTEND_INDEX.is_file() else None, assumed_app_root())


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description="Run the local React + pywebview desktop application.")
    parser.add_argument("--debug", action="store_true", help="Enable pywebview developer tools and verbose logs.")
    args = parser.parse_args(argv)
    logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(name)s: %(message)s")

    dist_dir = frontend_dist_dir()
    index = frontend_index(dist_dir)
    if index is None or not index.is_file():
        try:
            frontend_dist_dir(require=True)
        except FileNotFoundError as exc:
            parser.error(str(exc))
        parser.error(f"React bundle not found: {index}")
    try:
        import webview
    except ImportError:
        parser.error("pywebview is missing. Install `requirements-webview.txt` into the development environment.")

    # Resolve against pywebview's OWN app root (sys._MEIPASS when frozen) and log both, so a packaged app
    # that cannot find its bundle explains itself in logs/ instead of opening a blank window.
    app_root = runtime_app_root()
    url = frontend_url(dist_dir, app_root)
    LOG.info("WEBVIEW_FRONTEND packaged=%s app_root=%s dist=%s url=%s",
             is_packaged(), app_root, dist_dir, url)
    if not (Path(app_root) / url).resolve() == Path(dist_dir) / FRONTEND_INDEX_NAME:
        LOG.error("WEBVIEW_FRONTEND_MISMATCH pywebview would resolve %r to %s, expected %s",
                  url, (Path(app_root) / url), Path(dist_dir) / FRONTEND_INDEX_NAME)

    service = ApplicationService()
    bridge = BridgeService(service)
    window = webview.create_window(
        APP_TITLE,
        url=url,
        js_api=bridge,
        width=1440,
        height=940,
        min_size=(1000, 680),
        confirm_close=True,
        text_select=True,
    )
    bridge.bind_window(window)
    LOG.info("WEBVIEW_BRIDGE bound")

    def on_loaded(*_args):
        LOG.info("WEBVIEW_BRIDGE ready")

    window.events.loaded += on_loaded

    def on_closing(*_args):
        allowed, message = service.close_requested()
        if allowed:
            return True
        try:
            window.evaluate_js("window.dispatchEvent(new CustomEvent('tnp-close-requested', {detail: "
                               + __import__("json").dumps(message, ensure_ascii=False) + "}))")
        except Exception:
            LOG.debug("Could not notify React about a deferred close", exc_info=True)
        return False

    window.events.closing += on_closing
    webview.start(debug=args.debug, gui="edgechromium" if sys.platform == "win32" else None)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
