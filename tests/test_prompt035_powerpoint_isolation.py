"""PROMPT-035 — PowerPoint session isolation: a user's PowerPoint must never be closed by Report Extractor.

Root cause being pinned (see ``.specify/bugs/prompt035-powerpoint-session-isolation/assessment.md``):
PowerPoint is a *single-instance* automation server.  ``DispatchEx("PowerPoint.Application")`` therefore hands back
the user's already-running instance, yet the old renderer treated the returned object as "ours" and called
``Application.Quit()`` on it.  These tests drive ``render_with_powerpoint`` against a fake COM machine that models that
behaviour (and records every Close / Quit together with *who owns* the thing being closed), so the unsafe lifecycle
fails here and the ownership-aware one passes.

NOTHING HERE REQUIRES REAL POWERPOINT.  The fake models documented single-instance semantics; it is development
evidence only and is never Windows / real-PowerPoint acceptance (constitution §VII).
"""
from __future__ import annotations

import itertools
import os
import re
import subprocess
import sys
from pathlib import Path
from types import ModuleType

import pytest
from PIL import Image

import app.qpn_renderer as qrn
from app.cancellation import CancellationRequested
from app.qpn_renderer import PowerPointStageError, render_with_powerpoint

MK_E_UNAVAILABLE = -2147221021  # 0x800401E3 – GetActiveObject when no instance is registered


class _ComError(Exception):
    def __init__(self, hresult: int, text: str = "com error"):
        super().__init__(hresult, text)
        self.hresult = hresult


# ---------------------------------------------------------------------------------------------------------- fakes
class FakePresentation:
    def __init__(self, app: "FakeApp", name: str, owner: str):
        self.app, self.name, self.owner, self.closed = app, name, owner, False
        self.PageSetup = type("PS", (), {"SlideWidth": 960, "SlideHeight": 540})()

    def Slides(self, n):
        pres = self

        class Slide:
            def Export(self_inner, target, fmt, w, h):
                pres.app.machine.log("Export", pres.owner, pres.name)
                hook = pres.app.machine.on_export
                if hook:
                    hook(pres)
                Image.new("RGB", (int(w), int(h)), "white").save(target)
        return Slide()

    def Close(self):
        self.app.machine.log("Presentation.Close", self.owner, self.name)
        if self.owner == "user":
            self.app.machine.user_presentations_closed.append(self.name)
        self.closed = True
        if self in self.app._presentations:
            self.app._presentations.remove(self)


class FakePresentations:
    def __init__(self, app: "FakeApp"):
        self.app = app

    @property
    def Count(self):
        return len(self.app._presentations)

    def Open(self, path, read_only, untitled, with_window):
        m = self.app.machine
        m.log("Open", self.app.owner, Path(path).name)
        if self.app.owner == "user":
            m.touched_user_app = True
        if m.fail == "presentation-open":
            raise RuntimeError("cannot open presentation")
        pres = FakePresentation(self.app, Path(path).name, "app")
        self.app._presentations.append(pres)
        return pres


class FakeApp:
    def __init__(self, machine: "FakeMachine", pid: int, owner: str):
        self.machine, self.pid, self.owner = machine, pid, owner
        self._presentations: list = []
        self.Presentations = FakePresentations(self)
        self.Windows = type("W", (), {"Count": 0})()
        self.quit_calls = 0
        self.alive = True

    def add_user_presentation(self, name: str = "user_unsaved.pptx") -> FakePresentation:
        pres = FakePresentation(self, name, "user")
        self._presentations.append(pres)
        self.Windows = type("W", (), {"Count": len(self._presentations)})()
        return pres

    def Quit(self):
        self.machine.log("App.Quit", self.owner, f"pid={self.pid}")
        self.quit_calls += 1
        if self.machine.fail == "quit":
            raise RuntimeError("PowerPoint is busy")
        # real Quit closes every presentation the process holds – including a user's unsaved work
        for pres in list(self._presentations):
            if pres.owner == "user":
                self.machine.user_presentations_closed.append(pres.name)
        self._presentations.clear()
        self.alive = False
        self.machine.pending_exit[self.pid] = self.machine.quit_lingers_polls


class FakeMachine:
    """A Windows box with single-instance PowerPoint semantics and a process table."""

    def __init__(self):
        self.events: list = []
        self.running: FakeApp | None = None
        self.pending_exit: dict = {}
        self.quit_lingers_polls = 0
        self._pids = itertools.count(5000)
        self.fail = ""
        self.on_export = None
        self.on_dispatch = None
        self.user_presentations_closed: list = []
        self.touched_user_app = False
        self.snapshot_unreadable = False
        self.snapshot_blind = False          # the process table cannot see a running PowerPoint (ROT still can)
        self.dispatches = 0
        self.co_init = 0
        self.co_uninit = 0

    def log(self, name, owner="-", detail=""):
        self.events.append((name, owner, detail))

    def names(self):
        return [e[0] for e in self.events]

    def calls(self, name, owner=None):
        return [e for e in self.events if e[0] == name and (owner is None or e[1] == owner)]

    # -- user behaviour
    def start_user_powerpoint(self, *presentations: str) -> FakeApp:
        app = FakeApp(self, 4000, "user")
        for name in presentations:
            app.add_user_presentation(name)
        self.running = app
        return app

    # -- COM surface
    def dispatch_ex(self, progid):
        self.dispatches += 1
        self.log("DispatchEx", "-", progid)
        if self.on_dispatch:
            self.on_dispatch(self)
        if self.running is not None and self.running.alive:
            return self.running                      # SINGLE-INSTANCE: the running (user's) PowerPoint is returned
        app = FakeApp(self, next(self._pids), "app")
        self.running = app
        return app

    def get_active_object(self, progid):
        self.log("GetActiveObject", "-", progid)
        if self.running is not None and self.running.alive:
            return self.running
        raise _ComError(MK_E_UNAVAILABLE, "Operation unavailable")

    def process_ids(self):
        if self.snapshot_unreadable:
            return None
        if self.snapshot_blind:
            return frozenset()
        pids = set()
        if self.running is not None and self.running.alive:
            pids.add(self.running.pid)
        for pid in list(self.pending_exit):                      # a quitting process lingers for a few polls
            if self.pending_exit[pid] > 0:
                self.pending_exit[pid] -= 1
                pids.add(pid)
        return frozenset(pids)

    def install(self, monkeypatch):
        machine = self
        pythoncom = ModuleType("pythoncom")

        def co_init():
            machine.co_init += 1
        pythoncom.CoInitialize = co_init
        pythoncom.CoUninitialize = lambda: setattr(machine, "co_uninit", machine.co_uninit + 1)
        client = ModuleType("win32com.client")
        client.DispatchEx = self.dispatch_ex
        client.GetActiveObject = self.get_active_object
        client.Dispatch = lambda progid: pytest.fail("Dispatch may never be used: it attaches to a running instance")
        client.GetObject = lambda *a, **k: pytest.fail("GetObject may never be used")
        win32com = ModuleType("win32com")
        win32com.client = client
        monkeypatch.setitem(sys.modules, "pythoncom", pythoncom)
        monkeypatch.setitem(sys.modules, "win32com", win32com)
        monkeypatch.setitem(sys.modules, "win32com.client", client)
        # seams introduced by the fix (raising=False: they do not exist in the unsafe implementation)
        monkeypatch.setattr(qrn, "_powerpoint_process_ids", self.process_ids, raising=False)
        # The fake box's process table cannot vouch for itself: pin the fail-closed default so the suite
        # models the same machine on every host.  (Left unpinned, a real Windows host vouches for real —
        # its Toolhelp snapshot works and contains the test process — and the ROT scenarios below take the
        # activation path instead of the pre-dispatch refusal.)  Suites that model a vouching table
        # override this AFTER install().
        monkeypatch.setattr(qrn, "_process_table_vouched", lambda: False, raising=False)
        monkeypatch.setattr(qrn, "_PROCESS_EXIT_TIMEOUT_S", 2.0, raising=False)
        monkeypatch.setattr(qrn, "_PROCESS_POLL_S", 0.0, raising=False)
        # ANY attempt to end a process from the renderer is a test failure
        def forbidden(*a, **k):
            raise AssertionError(f"process termination attempted: {a!r}")
        monkeypatch.setattr(os, "kill", forbidden)
        monkeypatch.setattr(subprocess, "run", forbidden)
        monkeypatch.setattr(subprocess, "Popen", forbidden)
        return self


@pytest.fixture()
def machine(monkeypatch):
    return FakeMachine().install(monkeypatch)


@pytest.fixture()
def deck(tmp_path):
    path = tmp_path / "job.pptx"
    path.write_bytes(b"PK-not-a-real-deck")           # the fake COM layer never parses it
    return path


def _render(deck, tmp_path):
    return render_with_powerpoint(deck, [1], tmp_path, width_px=960)


def _render_tolerating_refusal(deck, tmp_path):
    """Run the renderer; a fail-closed refusal is acceptable, a crash of another kind is not."""
    try:
        return _render(deck, tmp_path)
    except PowerPointStageError as exc:
        return exc


# =============================================================================================== TEST 1
def test_1_existing_user_powerpoint_session_is_never_quit(machine, deck, tmp_path):
    user = machine.start_user_powerpoint()                       # PowerPoint open, start screen only
    _render_tolerating_refusal(deck, tmp_path)
    assert user.quit_calls == 0, "Application.Quit() was sent to the user's PowerPoint"
    assert user.alive and machine.running is user
    assert machine.calls("App.Quit", "user") == []
    assert machine.touched_user_app is False, "the user's instance must not be used at all"


def test_1b_refusal_is_fail_closed_and_names_the_ownership_stage(machine, deck, tmp_path):
    machine.start_user_powerpoint()
    with pytest.raises(PowerPointStageError) as info:
        _render(deck, tmp_path)
    assert info.value.stage == "ownership"
    assert machine.dispatches == 0, "must not even touch COM activation while a PowerPoint session exists"
    assert machine.co_init == machine.co_uninit == 1, "CoInitialize/CoUninitialize stay balanced on refusal"


# =============================================================================================== TEST 2
def test_2_existing_user_presentation_is_never_closed(machine, deck, tmp_path):
    user = machine.start_user_powerpoint("user_unsaved.pptx")
    _render_tolerating_refusal(deck, tmp_path)
    assert machine.user_presentations_closed == [], "the user's open presentation was closed"
    assert [p.name for p in user._presentations] == ["user_unsaved.pptx"]
    assert machine.calls("Presentation.Close", "user") == []
    assert machine.calls("Open", "user") == [], "no document may be opened inside the user's session"


# =============================================================================================== TEST 3
def test_3_app_opened_presentation_is_closed_after_processing(machine, deck, tmp_path):
    out = _render(deck, tmp_path)
    assert out[1].exists()
    closes = machine.calls("Presentation.Close")
    assert len(closes) == 1 and closes[0][1] == "app" and closes[0][2] == "job.pptx"


# =============================================================================================== TEST 4
def test_4_app_owned_application_is_quit_after_processing_and_waited_for(machine, deck, tmp_path):
    machine.quit_lingers_polls = 3
    _render(deck, tmp_path)
    quits = machine.calls("App.Quit")
    assert len(quits) == 1 and quits[0][1] == "app"
    names = machine.names()
    assert names.index("Presentation.Close") < names.index("App.Quit")
    assert machine.co_init == machine.co_uninit == 1
    assert machine.process_ids() == frozenset(), "the app-owned PowerPoint process is gone after cleanup"
    assert machine.pending_exit[5000] == 0, "cleanup waited (bounded) for the owned process to exit"


def test_4b_second_render_after_cleanup_can_use_powerpoint_again(machine, deck, tmp_path):
    machine.quit_lingers_polls = 2
    _render(deck, tmp_path)
    _render(deck, tmp_path)
    assert len(machine.calls("App.Quit", "app")) == 2


# =============================================================================================== TEST 5
@pytest.mark.parametrize("stage", ["presentation-open", "slide-export"])
def test_5_exception_path_closes_only_app_owned_resources(machine, deck, tmp_path, stage):
    if stage == "presentation-open":
        machine.fail = "presentation-open"
    else:
        machine.on_export = lambda pres: (_ for _ in ()).throw(RuntimeError("export refused"))
    with pytest.raises(PowerPointStageError) as info:
        _render(deck, tmp_path)
    assert info.value.stage == stage
    assert machine.calls("Presentation.Close", "user") == [] and machine.calls("App.Quit", "user") == []
    assert len(machine.calls("App.Quit", "app")) == 1
    assert len(machine.calls("Presentation.Close", "app")) == (0 if stage == "presentation-open" else 1)
    assert machine.co_init == machine.co_uninit == 1


def test_5b_exception_with_user_session_present_touches_nothing(machine, deck, tmp_path):
    user = machine.start_user_powerpoint("a.pptx", "b.pptx")
    machine.on_export = lambda pres: (_ for _ in ()).throw(RuntimeError("export refused"))
    _render_tolerating_refusal(deck, tmp_path)
    assert user.quit_calls == 0 and machine.user_presentations_closed == []


def test_5c_cleanup_failure_never_escalates_to_process_termination(machine, deck, tmp_path):
    machine.fail = "quit"                                         # Quit raises (hung / busy) – os.kill/subprocess are armed
    out = _render(deck, tmp_path)                                 # exported slide is kept, no kill, no crash
    assert out[1].exists()
    assert len(machine.calls("App.Quit", "app")) == 1


# =============================================================================================== TEST 6
def test_6_cancellation_inside_a_running_export_is_not_a_failure_and_cleans_only_owned(machine, deck, tmp_path):
    def cancel(pres):
        raise CancellationRequested("cancel")
    machine.on_export = cancel
    with pytest.raises(CancellationRequested):                    # cooperative cancel is never reclassified (§VI)
        _render(deck, tmp_path)
    assert machine.calls("Presentation.Close", "user") == [] and machine.calls("App.Quit", "user") == []
    assert len(machine.calls("Presentation.Close", "app")) == 1 and len(machine.calls("App.Quit", "app")) == 1


def test_6b_timeout_in_com_call_cleans_only_owned_and_never_kills(machine, deck, tmp_path):
    def hang(pres):
        raise TimeoutError("RPC timeout")
    machine.on_export = hang
    with pytest.raises(PowerPointStageError) as info:
        _render(deck, tmp_path)                                   # os.kill / subprocess are armed in the fixture
    assert info.value.stage == "slide-export"
    assert len(machine.calls("App.Quit", "app")) == 1 and machine.calls("App.Quit", "user") == []


def test_6c_cancel_before_backend_start_never_touches_com_with_user_session(machine, tmp_path):
    from app.pptx_parser import parse_pptx
    from tests.test_prompt027_rendering import _simple_slide
    from tests.test_after_evidence import _pics
    report = parse_pptx(str(_simple_slide(tmp_path / "c.pptx", lambda s: _pics(s, ["#c8e6c9"], 4, 3, w=2, h=1.4))))
    user = machine.start_user_powerpoint("keep_open.pptx")
    renderer = qrn.SlideRenderer(prefer=("powerpoint",))
    with pytest.raises(CancellationRequested):
        renderer.render(report, [1], tmp_path / "o", should_cancel=lambda: True)
    assert machine.events == [] and user.quit_calls == 0


# =============================================================================================== TEST 7
def test_7_user_presentation_arriving_in_our_instance_is_never_closed_or_quit(machine, deck, tmp_path):
    """The user double-clicks a deck while our hidden instance is rendering: Windows routes it to the one running
    PowerPoint, i.e. OUR instance.  We may close only our own presentation, and must NOT Quit the shared process."""
    holder = {}

    def user_opens_a_file(pres):
        holder["user_pres"] = pres.app.add_user_presentation("opened_by_user_midway.pptx")

    machine.on_export = user_opens_a_file
    out = _render(deck, tmp_path)
    assert out[1].exists()
    app = machine.running
    assert machine.calls("Presentation.Close", "app") != [] and len(machine.calls("Presentation.Close", "app")) == 1
    assert machine.calls("Presentation.Close", "user") == []
    assert machine.calls("App.Quit") == [], "a PowerPoint that now holds user work is not exclusively ours"
    assert app.alive and [p.name for p in app._presentations] == ["opened_by_user_midway.pptx"]
    assert machine.user_presentations_closed == []


def test_7b_user_and_app_presentations_coexist_only_the_app_one_is_closed(machine, deck, tmp_path):
    user = machine.start_user_powerpoint("user_unsaved.pptx")
    _render_tolerating_refusal(deck, tmp_path)
    closed = [e for e in machine.events if e[0] == "Presentation.Close"]
    assert all(e[1] == "app" for e in closed)
    assert machine.user_presentations_closed == [] and user.alive


# ================================================================================ ownership cannot be proven
def test_race_user_launches_powerpoint_during_dispatch_is_not_adopted_as_ours(machine, deck, tmp_path):
    """TOCTOU: no PowerPoint at the pre-check, the user starts one before DispatchEx returns – COM hands us THEIR
    instance (it already holds a document).  It must be released, never opened into, closed or quit."""
    holder = {}

    def user_starts(m):
        holder["user"] = m.start_user_powerpoint("race_unsaved.pptx")

    machine.on_dispatch = user_starts
    with pytest.raises(PowerPointStageError) as info:
        _render(deck, tmp_path)
    assert info.value.stage == "ownership"
    user = holder["user"]
    assert user.quit_calls == 0 and machine.user_presentations_closed == [] and user.alive
    assert machine.calls("Open") == [] and machine.calls("Presentation.Close") == []


def test_running_instance_invisible_to_process_table_is_caught_by_the_running_object_table(machine, deck, tmp_path):
    user = machine.start_user_powerpoint("hidden_from_snapshot.pptx")
    machine.snapshot_blind = True
    with pytest.raises(PowerPointStageError) as info:
        _render(deck, tmp_path)
    assert info.value.stage == "ownership" and machine.dispatches == 0
    assert user.quit_calls == 0 and machine.user_presentations_closed == []


def test_unreadable_process_table_fails_closed(machine, deck, tmp_path):
    machine.snapshot_unreadable = True
    with pytest.raises(PowerPointStageError) as info:
        _render(deck, tmp_path)
    assert info.value.stage == "ownership" and machine.dispatches == 0
    assert machine.calls("App.Quit") == [] and machine.calls("Presentation.Close") == []


def test_ownership_refusal_degrades_through_the_renderer_chain(machine, tmp_path, monkeypatch):
    """Fail-closed must not break extraction: the chain falls through to the next backend and records a
    *non-availability* failure (so the degraded preview is retried, not pinned)."""
    from app.pptx_parser import parse_pptx
    from tests.test_prompt027_rendering import _simple_slide
    from tests.test_after_evidence import _pics
    report = parse_pptx(str(_simple_slide(tmp_path / "d.pptx", lambda s: _pics(s, ["#c8e6c9"], 4, 3, w=2, h=1.4))))
    user = machine.start_user_powerpoint("keep.pptx")
    monkeypatch.setattr(qrn, "powerpoint_available", lambda: True)
    renderer = qrn.SlideRenderer(prefer=("powerpoint", "builtin"))
    out = renderer.render(report, [1], tmp_path / "o2")
    assert out[1].exists() and renderer.last_backend == "builtin"
    assert renderer.last_failures[0]["stage"] == "ownership" and renderer.degraded is True
    assert user.quit_calls == 0 and machine.user_presentations_closed == []


# ===================================================================================== concurrency / seams
def test_concurrent_jobs_cannot_adopt_or_quit_each_others_instance(machine, deck, tmp_path):
    """Batch worker + learning preview can render at once.  The lifecycle is serialised, so job B never receives
    job A's running instance (which it would otherwise 'own' and quit)."""
    import threading
    a_in_export = threading.Event()
    release_a = threading.Event()
    first = {"done": False}

    def hold_first(pres):
        if not first["done"]:
            first["done"] = True
            a_in_export.set()
            release_a.wait(2.0)

    machine.on_export = hold_first
    results = {}
    out_a, out_b = tmp_path / "a", tmp_path / "b"
    out_a.mkdir(), out_b.mkdir()

    def job(name, out_dir):
        try:
            results[name] = render_with_powerpoint(deck, [1], out_dir, width_px=960)
        except PowerPointStageError as exc:
            results[name] = exc

    ta = threading.Thread(target=job, args=("a", out_a))
    tb = threading.Thread(target=job, args=("b", out_b))
    ta.start()
    assert a_in_export.wait(2.0)
    tb.start()
    import time as _t
    _t.sleep(0.1)
    assert machine.dispatches == 1, "job B must wait for job A's lifecycle to finish"
    release_a.set()
    ta.join(5), tb.join(5)
    names = machine.names()
    assert names.count("DispatchEx") == 2 and names.count("App.Quit") == 2
    assert names.index("App.Quit") < len(names) - 1 - names[::-1].index("DispatchEx"), "lifecycles interleaved"
    assert all(isinstance(r, dict) for r in results.values()), results


def test_real_process_table_reader_reports_unknown_off_windows(monkeypatch):
    monkeypatch.setattr(sys, "platform", "linux")
    assert qrn._powerpoint_process_ids() is None            # unknown => callers fail closed


# =============================================================================================== TEST 8
APP_DIR = Path(qrn.__file__).resolve().parent


def _code_without_comments_and_docstrings(text: str) -> str:
    import ast
    tree = ast.parse(text)
    for node in ast.walk(tree):
        if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef, ast.ClassDef, ast.Module)):
            body = node.body
            if body and isinstance(body[0], ast.Expr) and isinstance(getattr(body[0], "value", None), ast.Constant) \
                    and isinstance(body[0].value.value, str):
                body[0] = ast.Pass()
    return ast.unparse(tree)


def test_8_no_global_powerpoint_termination_anywhere_in_the_application():
    forbidden_everywhere = ("taskkill", "TerminateProcess", "psutil", "pkill", "killall", "wmic process")
    for path in sorted(APP_DIR.glob("*.py")):
        code = _code_without_comments_and_docstrings(path.read_text(encoding="utf-8")).lower()
        for needle in forbidden_everywhere:
            assert needle.lower() not in code, f"{path.name} contains {needle!r}"


def test_8b_renderer_has_no_process_kill_primitive_at_all():
    code = _code_without_comments_and_docstrings(Path(qrn.__file__).read_text(encoding="utf-8"))
    for pattern in (r"\bos\.kill\b", r"\.kill\(", r"\.terminate\(", r"\bOpenProcess\b", r"\bkillpg\b",
                    r"\bsignal\.", r"\bsubprocess\.(?:Popen|run|call|check_call|check_output)\([^)]*(?:taskkill|powerpnt)"):
        assert not re.search(pattern, code, re.IGNORECASE), f"renderer contains forbidden primitive {pattern!r}"


def test_8c_cleanup_never_iterates_or_closes_every_presentation_and_never_attaches():
    code = _code_without_comments_and_docstrings(Path(qrn.__file__).read_text(encoding="utf-8"))
    assert not re.search(r"for\s+\w+\s+in\s+\w*\.?Presentations\b", code), "must not enumerate/close all presentations"
    assert "ActivePresentation" not in code and "GetObject" not in code
    assert not re.search(r"win32com\.client\.Dispatch\(", code), "attaching Dispatch() is forbidden; only DispatchEx"
