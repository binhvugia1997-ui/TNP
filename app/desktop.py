"""Run the React frontend inside a local pywebview desktop window (development only)."""
from __future__ import annotations

import argparse
import logging
from pathlib import Path

from . import APP_TITLE
from .application_service import ApplicationService
from .bridge import BridgeService

LOG = logging.getLogger("report_extractor.webview.desktop")
REPO_ROOT = Path(__file__).resolve().parent.parent
FRONTEND_DIST = REPO_ROOT / "frontend" / "dist"
FRONTEND_INDEX = FRONTEND_DIST / "index.html"
# pywebview resolves this relative to app/ and serves only frontend/dist over its loopback HTTP server.
# An absolute file:// URL breaks Chromium ES-module asset loading and is intentionally not used.
FRONTEND_URL = "../frontend/dist/index.html"


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description="Run the local React + pywebview development application.")
    parser.add_argument("--debug", action="store_true", help="Enable pywebview developer tools and verbose logs.")
    args = parser.parse_args(argv)
    logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(name)s: %(message)s")

    if not FRONTEND_INDEX.is_file():
        parser.error(f"React bundle not found: {FRONTEND_INDEX}. Run `cd frontend && npm run build` first.")
    try:
        import webview
    except ImportError:
        parser.error("pywebview is missing. Install `requirements-webview.txt` into the development environment.")

    service = ApplicationService()
    bridge = BridgeService(service)
    window = webview.create_window(
        APP_TITLE,
        url=FRONTEND_URL,
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
    webview.start(debug=args.debug, gui="edgechromium" if __import__("sys").platform == "win32" else None)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
