"""PROMPT-036R: desktop hang after pywebviewready / BRIDGE_UNAVAILABLE / Learning UI unresponsive.

ROOT CAUSE (proven below, not assumed): ``BridgeService`` is the exact object pywebview hands to its own
``webview.util.get_functions()`` reflection walker to build the JS method manifest right after
``before_load`` fires (the debug line "_pywebviewready event fired" that the bug report captured is printed
by pywebview BEFORE that walk runs, not after it succeeds — see the comment in ``app/bridge.py``). That
walker recursively inspects every public, non-callable attribute with a ``__module__`` and only guards
against cycles by Python object ``id()``. ``BridgeService`` used to publish the live, bound ``pywebview``
window under a PUBLIC ``self.window`` attribute. Once the native window exists, ``window.gui``/``window.native``
resolve (on Windows) to pythonnet-wrapped WinForms/WebView2 CLR objects. Struct-typed CLR properties such as
``System.Drawing.Size``/``SizeF``/``Point``/``Padding`` carry a static ``Empty`` field that returns a BRAND
NEW wrapper object every time it is read, so two reads of the same logical value never share an ``id()`` —
pywebview's cycle guard cannot see the loop and the walker recurses down ``...Empty.Empty.Empty...`` until
CPython's recursion limit is hit, raising ``RecursionError``.

Reproducing pywebview's actual ``get_functions()`` verbatim (see ``pywebview_get_functions`` below) shows
that this ``RecursionError`` is NOT fatal to the whole call: pywebview wraps every single attribute visit in
its own ``try/except Exception: continue``, so each of the (up to) ~1000 nested recursive calls catches its
*own* ``RecursionError`` the instant it is raised and the walk unwinds normally. The practical effect is
therefore not a crash but a CPU/stack-burning storm: for every struct-typed CLR property reachable from the
bound window (a real WinForms ``Form``/WebView2 control has DOZENS — ``Size``, ``ClientSize``,
``MinimumSize``, ``Padding``, ``Margin``, ``Font``, ``Bounds``, …), the walker drives Python to the very
edge of its recursion limit and back, each time emitting pywebview's own
``"Error while processing <name>: maximum recursion depth exceeded"`` log line whose ``<name>`` is the
dotted attribute path built up during the descent — reproduced byte-for-byte below as
``...window.gui.Empty.Empty.Empty...Empty`` — which is exactly the "Empty.Empty.Empty...maximum recursion
depth exceeded" evidence in the bug report. Multiplied across every struct-valued property of a real native
window, this is tens of thousands of wasted recursive Python/CLR interop calls competing for the GIL on
pywebview's own startup thread — enough, on a real Windows WinForms message loop, to starve UI-thread
responsiveness long enough for Windows to mark the window "(Not Responding)" and for the frontend's own
bounded ping/method wait to lapse into "Python bridge chưa sẵn sàng · BRIDGE_UNAVAILABLE" even though
"before_load"/"_pywebviewready" had already been logged (those lines are printed by pywebview BEFORE this
reflection walk runs, not after it finishes cleanly).

``tests/test_prompt036r_bridge_hang.py`` reproduces this mechanism deterministically without needing real
pythonnet/Windows (not available in CI): a small stand-in object simulates the ".Empty returns a new
instance every time" CLR shape and is bound exactly the way ``app/desktop.py`` binds the real pywebview
window. The fix is the smallest possible change: the live window reference is now private
(``self._window``), which pywebview's own walker already skips (it explicitly ignores any attribute name
starting with ``_``), so NONE of this recursive reflection work ever happens, regardless of what the native
window resolves to on Windows.
"""
from __future__ import annotations

import inspect
import threading
import time
from typing import Any, Dict, List, Optional

import pytest

from app.application_service import ApplicationService
from app.bridge import BridgeService
from tests.test_prompt032_learning_state import _finished_production_service

REQUIRED_BRIDGE_METHODS = {
    "ping", "get_app_version", "get_current_config", "get_dashboard_state",
    "get_learning_state", "set_learning_label", "save_learning_labels",
    "scan_reports", "start_processing", "choose_pptx_file",
}


# ---------------------------------------------------------------------------------------- reflection harness
def pywebview_get_functions(obj: object, base_name: str = "", functions: Optional[Dict[str, Any]] = None,
                            exposed_objects: Optional[List[int]] = None,
                            errors: Optional[List[str]] = None) -> Dict[str, Any]:
    """A faithful port of ``webview.util.create_window``'s nested ``get_functions`` (pywebview 6.2.1).

    Ported (not imported) because pywebview is a native-GUI-only dependency not installable in this
    sandbox/CI; the algorithm below is copied verbatim from ``webview/util.py`` so this test exercises the
    REAL reflection logic pywebview runs at startup, including its ``id()``-based "already visited" guard,
    its leading-underscore skip, and its per-attribute ``except Exception: continue`` — all three are
    load-bearing for this bug. ``errors`` collects the attribute-path strings pywebview would have logged
    via ``logger.error('Error while processing %s: %s', full_name, e)`` for each caught exception, so a test
    can assert on whether (and how badly) the "Empty.Empty.Empty..." storm actually occurred.
    """
    if exposed_objects is None:
        exposed_objects = []
    if functions is None:
        functions = {}
    if errors is None:
        errors = []

    obj_id = id(obj)
    if obj_id in exposed_objects:
        return functions
    exposed_objects.append(obj_id)

    for name in dir(obj):
        full_name = f"{base_name}.{name}" if base_name else name
        try:
            if name.startswith("_"):
                continue
            attr = getattr(obj, name)
            if not getattr(attr, "_serializable", True):
                continue
            if inspect.ismethod(attr) or inspect.isfunction(attr):
                try:
                    functions[full_name] = list(inspect.getfullargspec(attr).args)
                except TypeError:
                    functions[full_name] = []
            elif inspect.isclass(attr) or (
                isinstance(attr, object) and not callable(attr) and hasattr(attr, "__module__")
            ):
                pywebview_get_functions(attr, full_name, functions, exposed_objects, errors)
        except Exception as exc:  # mirrors pywebview's own `except Exception as e: logger.error(...); continue`
            errors.append(f"{full_name}: {type(exc).__name__}")
            continue
    return functions


class _PoisonClrValue:
    """Stands in for a pythonnet-wrapped .NET struct (``System.Drawing.Size`` et al.).

    Its ``Empty`` property mirrors the real static CLR field: reading it returns a NEW instance every
    time, so no two reads ever share a Python ``id()``. That is precisely what defeats pywebview's
    ``id()``-based reflection cycle guard and is the mechanism behind the observed
    "Empty.Empty.Empty...maximum recursion depth exceeded" failure.
    """

    visit_count = 0

    @property
    def Empty(self):  # noqa: N802 - matches the real CLR member's PascalCase name on purpose
        _PoisonClrValue.visit_count += 1
        return _PoisonClrValue()


class _PoisonNativeWindow:
    """Stands in for the live pywebview ``Window`` once its native backend has attached CLR handles."""

    def __init__(self) -> None:
        self.gui = _PoisonClrValue()
        self.native = _PoisonClrValue()

    def create_file_dialog(self, *_args, **_kwargs):  # pragma: no cover - not exercised here
        return None

    def destroy(self) -> None:  # pragma: no cover - not exercised here
        pass


@pytest.fixture
def minimal_service(tmp_path) -> ApplicationService:
    return ApplicationService(config_path=tmp_path / "config" / "settings.json",
                              manual_fields_path=tmp_path / "config" / "manual_fields.json")


# ---------------------------------------------------------------------------------------- 1) the fix itself
def test_bound_window_is_private_and_not_a_public_bridge_attribute(minimal_service):
    """The exact shape of the fix: pywebview's reflection only skips names starting with ``_``."""
    bridge = BridgeService(minimal_service)
    poison = _PoisonNativeWindow()
    bridge.bind_window(poison)

    assert not hasattr(bridge, "window"), "a public `window` attribute is reachable by pywebview reflection again"
    assert bridge._window is poison


# ------------------------------------------------------------- 2) no recursion through the real bridge object
def test_pywebview_reflection_does_not_recurse_through_bound_bridge(minimal_service):
    """Regression test #2 and #6: startup reflection cannot recurse into a bound CLR-like window.

    Runs the REAL pywebview reflection algorithm over the REAL ``BridgeService`` object, exactly as
    ``generate_js_object()`` does the moment ``before_load`` fires, with the native window already bound
    (as ``app/desktop.py`` does, before ``webview.start()``).
    """
    bridge = BridgeService(minimal_service)
    poison = _PoisonNativeWindow()
    bridge.bind_window(poison)
    _PoisonClrValue.visit_count = 0
    errors: List[str] = []

    functions = pywebview_get_functions(bridge, errors=errors)

    assert _PoisonClrValue.visit_count == 0, "the poisoned native window must never be reflected into"
    assert errors == [], f"reflection must not even attempt the poisoned window, got: {errors}"
    missing = REQUIRED_BRIDGE_METHODS - set(functions)
    assert not missing, f"required bridge methods missing from the generated manifest: {missing}"


def test_pywebview_reflection_survives_even_without_any_recursion_limit_margin(minimal_service):
    """Same as above, but with a tight recursion limit so a reintroduced leak fails loudly and immediately."""
    bridge = BridgeService(minimal_service)
    bridge.bind_window(_PoisonNativeWindow())

    import sys
    previous_limit = sys.getrecursionlimit()
    sys.setrecursionlimit(200)
    try:
        errors: List[str] = []
        functions = pywebview_get_functions(bridge, errors=errors)
        assert errors == [], f"a tight recursion limit must never even be approached: {errors}"
    finally:
        sys.setrecursionlimit(previous_limit)
    assert REQUIRED_BRIDGE_METHODS <= set(functions)


# ------------------------------------------------- 3) prove the mechanism: the pre-fix shape DOES recurse
def test_public_window_attribute_reproduces_the_empty_empty_recursion():
    """Historical-evidence test: proves WHY the fix works by reproducing the bug in isolation.

    This intentionally rebuilds the bridge's PRE-FIX shape (a public, non-underscore ``window`` attribute
    bound to a live CLR-like object) in a throwaway stand-in class — never by mutating the real
    ``BridgeService`` — and feeds it through the SAME reflection algorithm used above. Because pywebview
    wraps every attribute visit in its own ``try/except Exception: continue``, the resulting
    ``RecursionError`` is caught one level at a time and never escapes ``get_functions`` itself (proven by
    the first assertion below); what DOES happen is pywebview logging the exact
    "Error while processing <path>: maximum recursion depth exceeded" shape the bug report captured, with
    ``<path>`` literally built out of hundreds of repeated ``.Empty`` segments (proven by the second
    assertion) — while burning a full stack's worth of recursive Python/CLR calls to get there. This is the
    reproduced, concrete explanation for the "Empty.Empty.Empty...maximum recursion depth exceeded" evidence.
    """

    class _PreFixBridgeShape:
        """Mirrors `BridgeService.__init__` / `bind_window` BEFORE PROMPT-036R (public `self.window`)."""

        def __init__(self) -> None:
            self.window = None

        def bind_window(self, window) -> None:
            self.window = window

        def ping(self) -> dict:  # pragma: no cover - only here to resemble the real contract shape
            return {"ok": True, "data": "Python bridge OK"}

    pre_fix_bridge = _PreFixBridgeShape()
    pre_fix_bridge.bind_window(_PoisonNativeWindow())

    import sys
    previous_limit = sys.getrecursionlimit()
    sys.setrecursionlimit(300)  # bound the proof so the test is fast instead of eating a full real stack
    try:
        errors: List[str] = []
        functions = pywebview_get_functions(pre_fix_bridge, errors=errors)  # does NOT raise - see docstring
    finally:
        sys.setrecursionlimit(previous_limit)

    assert errors, "the poisoned public window must trigger at least one caught RecursionError"
    assert any(error.endswith(": RecursionError") for error in errors)
    storm_path = next(error for error in errors if error.endswith(": RecursionError"))
    assert ".Empty.Empty.Empty.Empty.Empty" in storm_path, (
        f"expected the reported repeated-'.Empty' recursion shape, got: {storm_path[:120]}..."
    )
    # The contract methods living beside `window` on the SAME object are completely unaffected; only the
    # poisoned branch under `window.*` is ever contaminated by the caught recursion.
    assert "ping" in functions


# -------------------------------------------------------------------------- 4) contract intact after the fix
def test_ping_and_startup_contract_calls_still_succeed_after_binding_the_window(minimal_service):
    """Regression test #1 (Python side): ping/app-version/config/dashboard all still succeed once bound."""
    bridge = BridgeService(minimal_service)
    bridge.bind_window(_PoisonNativeWindow())

    for method in ("ping", "get_app_version", "get_current_config", "get_dashboard_state"):
        envelope = getattr(bridge, method)()
        assert envelope["ok"] is True, f"{method} unexpectedly failed: {envelope}"


# ------------------------------------------------------------------ 5) slow backend work is awaited, not failed
def test_slow_dashboard_work_still_returns_a_healthy_envelope(minimal_service, monkeypatch):
    """Regression test #3 (Python side): the backend never turns slow-but-working into a failure signal.

    ``BRIDGE_UNAVAILABLE`` is purely a frontend transport classification (see
    ``frontend/src/lib/bridge.ts``); the Python contract must keep returning the SAME ``{"ok": True, ...}``
    envelope no matter how long ``dashboard_state`` legitimately takes, so a slow dashboard refresh can
    never be misread as a dead bridge.
    """
    bridge = BridgeService(minimal_service)
    original = minimal_service._dashboard_state_unlocked

    def slow(*args, **kwargs):
        time.sleep(0.2)
        return original(*args, **kwargs)

    monkeypatch.setattr(minimal_service, "_dashboard_state_unlocked", slow)

    started = time.perf_counter()
    envelope = bridge.get_dashboard_state()
    elapsed = time.perf_counter() - started

    assert envelope["ok"] is True
    assert elapsed >= 0.2


# --------------------------------------------------------- 6) Learning stays responsive during normal polling
def test_learning_interaction_stays_responsive_during_concurrent_dashboard_polling(tmp_path):
    """Regression test #5: a Learning click must not stall behind ordinary concurrent dashboard polling."""
    service, controller, processor, _deck = _finished_production_service(tmp_path)
    monkeypatch_preview = pytest.MonkeyPatch()
    monkeypatch_preview.setattr(service, "_slide_preview_for",
                               lambda candidate, learning=None: {
                                   "src": "", "width": 0, "height": 0, "backend": "", "faithful": False,
                                   "regions": [],
                               })
    try:
        state = service.learning_state()  # applies the terminal worker event once, like the real poll would
        image_candidate = next(c for c in state["images"])
        candidate_id = image_candidate["id"]

        stop = threading.Event()
        poll_count = 0
        poll_errors: List[BaseException] = []

        def poll_loop() -> None:
            nonlocal poll_count
            while not stop.is_set():
                try:
                    service.dashboard_state()
                    poll_count += 1
                except BaseException as exc:  # noqa: BLE001 - captured for the assertion below
                    poll_errors.append(exc)
                time.sleep(0.02)  # much tighter than the real 750ms poll, to guarantee overlap in this test

        poller = threading.Thread(target=poll_loop, daemon=True)
        poller.start()
        time.sleep(0.05)  # let a few polls land first, as they would before any user click in production

        started = time.perf_counter()
        result = service.set_learning_label("image", candidate_id, "AFTER", "responsive click")
        elapsed = time.perf_counter() - started

        stop.set()
        poller.join(timeout=2)

        assert not poller.is_alive()
        assert not poll_errors
        assert poll_count > 0, "the dashboard poll never actually ran concurrently with the click"
        assert result["images"][0]["labelPending"] is True
        assert elapsed < 0.3, f"Learning click took {elapsed:.3f}s while normal dashboard polling ran concurrently"
    finally:
        monkeypatch_preview.undo()


# --------------------------------------------------------------------- 7) no bridge-side duplicate sessions
def test_rebinding_the_window_never_accumulates_background_state(minimal_service):
    """Regression test #4 (Python side complement): BridgeService itself owns no timers/threads/sessions.

    Reconnect/duplicate-polling-session protection is an explicitly tested frontend property
    (``frontend/tests/bridge.test.mjs``: "poll failure pauses the interval, reconnects once, and resumes
    without timer leaks" asserts exactly one active interval survives a reconnect). On the Python side
    there is no session object to duplicate; this test documents and locks in that invariant so a future
    change cannot quietly add one to ``BridgeService`` without a test noticing.
    """
    bridge = BridgeService(minimal_service)
    threads_before = threading.active_count()

    for _ in range(5):
        bridge.bind_window(_PoisonNativeWindow())
        assert bridge.ping()["ok"] is True

    assert threading.active_count() == threads_before
