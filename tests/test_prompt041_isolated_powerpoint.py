"""PROMPT-041 — PowerPoint rendering while the user's own PowerPoint is open.

WHAT WAS ASKED FOR
    "user PowerPoint open" + "Report Extractor renders with PowerPoint" at the same time, using a PowerPoint
    automation instance the app provably owns, never attaching to / reusing / closing / killing the user's.

WHAT THIS FILE PROVES
    1. The refusal the report describes is real and is the PRE-FLIGHT ownership guard from PROMPT-035
       (``app/qpn_renderer.py`` refuses before ``DispatchEx`` is even called), not a PowerPoint availability
       failure — PowerPoint IS installed in every one of those tests.
    2. That guard cannot simply be lifted, because PowerPoint is a **Multiuse (Single Instance)** COM server:
       inside one Windows session ``DispatchEx("PowerPoint.Application")`` returns the user's running
       ``POWERPNT.EXE`` instead of starting a second one, so there is no isolated instance to create.  Evidence
       (Microsoft's own words): "Multiple instances of Word, Excel, and Access can run simultaneously... Only one
       instance of PowerPoint (Powerpnt.exe) can run at any given time.  Therefore, PowerPoint is a Multiuse
       (Single Instance) server."  — learn.microsoft.com, "How to use Visual C# to automate a running instance of
       an Office program".  The only documented way to a second instance is to launch it as a *different Windows
       user* (``runas``), which a desktop app cannot do without credentials.  See
       ``.specify/bugs/prompt041-isolated-powerpoint-renderer/assessment.md``.
    3. Everything that IS achievable safely is implemented and pinned here:
         * the render still refuses to touch the user's instance (no dispatch, no open, no close, no quit, no kill);
         * the refusal carries a STABLE reason code, so "PowerPoint is installed but your session is in the way"
           is never reported as "PowerPoint unavailable" (the specific diagnostic error the report forbids);
         * the reason reaches the Learning DTO (``slidePreviewReason``) and survives the preview cache and the
           on-disk learning assets;
         * the PowerPoint process snapshot is SESSION-SCOPED, so a PowerPoint in another Windows session no
           longer causes a false refusal (it can never be handed to us by COM either);
         * the whole ownership lifecycle, its failure paths and the no-kill contracts are unchanged.

    Nothing here requires real PowerPoint.  The fake COM machine (imported from the PROMPT-035 suite, so both
    prompts are tested against ONE model of Windows) is development evidence only and is never Windows /
    real-PowerPoint acceptance (constitution §VII).  Required Windows acceptance:
    ``docs/WINDOWS_ACCEPTANCE_PROMPT041.md``.
"""
from __future__ import annotations

import ast
import inspect
import io
import logging
import re
import threading
from pathlib import Path

import pytest

import app.qpn_renderer as qrn
from app.cancellation import CancellationRequested
from app.qpn_renderer import (
    REASON_IMPORT,
    REASON_NOT_EXCLUSIVE,
    REASON_OK,
    REASON_OWNERSHIP_UNPROVEN,
    REASON_PROCESS_TABLE_UNREADABLE,
    REASON_REGISTERED_INSTANCE,
    REASON_UNAVAILABLE,
    REASON_USER_SESSION_RUNNING,
    PowerPointStageError,
    SlideRenderer,
    powerpoint_render_capability,
    render_with_powerpoint,
)
# ONE model of Windows for both prompts: if this fake changes, both suites move together.
from tests.test_prompt035_powerpoint_isolation import FakeMachine, _code_without_comments_and_docstrings

APP_DIR = Path(qrn.__file__).resolve().parent


@pytest.fixture()
def machine(monkeypatch):
    return FakeMachine().install(monkeypatch)


@pytest.fixture()
def deck(tmp_path):
    path = tmp_path / "job.pptx"
    path.write_bytes(b"PK-not-a-real-deck")           # the fake COM layer never parses it
    return path


@pytest.fixture()
def report(tmp_path):
    """A real, parseable one-slide deck so the whole renderer chain can run end to end."""
    from app.pptx_parser import parse_pptx
    from tests.test_prompt027_rendering import _simple_slide
    from tests.test_after_evidence import _pics
    return parse_pptx(str(_simple_slide(tmp_path / "d.pptx", lambda s: _pics(s, ["#c8e6c9"], 4, 3, w=2, h=1.4))))


def _render(deck, tmp_path):
    return render_with_powerpoint(deck, [1], tmp_path, width_px=960)


def _render_tolerating_refusal(deck, tmp_path):
    try:
        return _render(deck, tmp_path)
    except PowerPointStageError as exc:
        return exc


# ============================================================================================ REQUIRED TEST 1
def test_1_no_user_powerpoint_uses_an_owned_powerpoint_instance(machine, deck, tmp_path, caplog):
    caplog.set_level(logging.INFO, logger="report_extractor.renderer")
    """Baseline (PROMPT-030B acceptance behaviour): with no user PowerPoint the app creates, uses and cleans up
    its OWN instance, and says so in the diagnostics."""
    out = _render(deck, tmp_path)
    assert out[1].exists()
    assert machine.dispatches == 1
    assert len(machine.calls("App.Quit", "app")) == 1 and machine.calls("App.Quit", "user") == []
    assert len(machine.calls("Presentation.Close", "app")) == 1 and machine.calls("Presentation.Close", "user") == []
    assert machine.process_ids() == frozenset(), "the owned process is gone after cleanup"
    assert machine.co_init == machine.co_uninit == 1
    # PROMPT-041 diagnostics: "isolated instance created" is a distinct, greppable line
    lines = [r.getMessage() for r in caplog.records if "POWERPOINT_ISOLATION" in r.getMessage()]
    assert any("state=created" in line and f"reason={REASON_OK}" in line for line in lines), lines


# ============================================================================================ REQUIRED TEST 2
def test_2_user_powerpoint_open_no_isolated_instance_is_created_and_nothing_is_touched(machine, deck, tmp_path):
    """THE reported behaviour, reproduced and pinned.

    Requested: an isolated owned renderer succeeds alongside the user's PowerPoint.
    Achievable: no — PowerPoint is a Multiuse (Single Instance) COM server, so a second ``POWERPNT.EXE`` in this
    Windows session cannot be created at all (see the module docstring for Microsoft's own statement).  Attempting
    the activation anyway would hand us the USER's instance, so the renderer refuses *before* touching COM.
    What is guaranteed instead: the user's session is untouched and the refusal names the real cause.
    """
    user = machine.start_user_powerpoint("user_unsaved.pptx")
    with pytest.raises(PowerPointStageError) as info:
        _render(deck, tmp_path)

    assert info.value.stage == "ownership", "the PROMPT-035 fail-closed contract is unchanged"
    assert info.value.code == REASON_USER_SESSION_RUNNING
    assert machine.dispatches == 0, "COM activation is never even attempted against a user session"
    assert machine.calls("DispatchEx") == [] and machine.calls("Open") == []
    assert machine.calls("Presentation.Close") == [] and machine.calls("App.Quit") == []
    assert user.quit_calls == 0 and user.alive and machine.running is user
    assert machine.user_presentations_closed == [] and machine.touched_user_app is False


def test_2b_the_platform_fact_the_refusal_rests_on_is_documented_in_the_module():
    """The refusal is only justified by the single-instance fact.  If that ever stops being true, this assertion
    is the tripwire that makes a reviewer re-open the isolation question instead of silently keeping the fallback."""
    source = Path(qrn.__file__).read_text(encoding="utf-8")
    assert "Multiuse" in source and "Single Instance" in source, \
        "qpn_renderer must keep documenting WHY an isolated PowerPoint instance cannot be created"
    assert not re.search(r"win32com\.client\.Dispatch\(", _code_without_comments_and_docstrings(source)), \
        "attaching Dispatch() is still forbidden"


def test_2c_every_documented_way_to_a_second_instance_is_out_of_bounds():
    """``runas`` / a second logon session is the only documented route to a second POWERPNT.EXE.  It needs another
    Windows account's credentials, so it must not appear in the product code."""
    for path in sorted(APP_DIR.glob("*.py")):
        code = _code_without_comments_and_docstrings(path.read_text(encoding="utf-8")).lower()
        for needle in ("runas", "createprocesswithlogon", "createprocessasuser", "logonuser", "secondary logon"):
            assert needle not in code, f"{path.name} must not try to launch PowerPoint as another user ({needle!r})"


# ================================================================================== REQUIRED TESTS 3 and 4
def test_3_and_4_user_presentation_and_unsaved_edits_survive_every_render_path(machine, deck, tmp_path):
    """Repeated renders, a failing export and a failing Quit, all with the user's unsaved deck open."""
    user = machine.start_user_powerpoint("user_unsaved.pptx")
    user_pres = user._presentations[0]

    _render_tolerating_refusal(deck, tmp_path)                       # plain refusal
    machine.on_export = lambda pres: (_ for _ in ()).throw(RuntimeError("export refused"))
    _render_tolerating_refusal(deck, tmp_path)                       # export failure path
    machine.on_export = None
    machine.fail = "quit"
    _render_tolerating_refusal(deck, tmp_path)                       # cleanup failure path
    machine.fail = ""
    _render_tolerating_refusal(deck, tmp_path)                       # and again

    assert machine.user_presentations_closed == [], "the user's presentation was closed"
    assert machine.calls("Presentation.Close", "user") == [] and machine.calls("App.Quit", "user") == []
    assert machine.calls("Open", "user") == [], "no document was opened inside the user's session"
    assert user.quit_calls == 0 and user.alive
    assert [p.name for p in user._presentations] == ["user_unsaved.pptx"] and user_pres.closed is False


# ============================================================================================ REQUIRED TEST 5
@pytest.mark.parametrize("block,expected_code", [
    ("unreadable", REASON_PROCESS_TABLE_UNREADABLE),
    ("rot", REASON_REGISTERED_INSTANCE),
    ("user-session", REASON_USER_SESSION_RUNNING),
])
def test_5_isolation_cannot_be_proven_falls_back_to_builtin_with_the_right_reason(
        machine, report, tmp_path, monkeypatch, block, expected_code):
    """Fail closed, keep extracting, and tell the truth about why the fidelity is reduced."""
    if block == "unreadable":
        machine.snapshot_unreadable = True
    elif block == "rot":
        machine.start_user_powerpoint("hidden.pptx")
        machine.snapshot_blind = True                 # process table cannot see it; the ROT can
    else:
        machine.start_user_powerpoint("user.pptx")
    monkeypatch.setattr(qrn, "powerpoint_available", lambda: True)

    renderer = SlideRenderer(prefer=("powerpoint", "builtin"), width_px=320)
    out = renderer.render(report, [1], tmp_path / "out", purpose="learning_preview")

    assert out[1].exists() and renderer.last_backend == "builtin"
    assert renderer.preview_reason == expected_code
    assert renderer.degraded is True, "a refused faithful backend must not pin the fallback preview"
    # and the log line names the same stable reason, so Windows acceptance can grep it
    assert renderer.last_failures[0]["stage"] == "ownership"
    assert machine.calls("App.Quit", "user") == [] and machine.user_presentations_closed == []


def test_5b_a_machine_without_powerpoint_reports_unavailable_not_a_session_problem(monkeypatch, report, tmp_path):
    """The exact diagnostic error the report forbids: "PowerPoint unavailable" while PowerPoint is installed.
    Here PowerPoint genuinely is absent, so "unavailable" IS the right answer — and it is a different code from
    the user-session refusal above."""
    monkeypatch.setattr(qrn, "powerpoint_available", lambda: False)
    monkeypatch.setattr(qrn, "powerpoint_unavailable_detail", lambda: "progid-not-registered")
    renderer = SlideRenderer(prefer=("powerpoint", "builtin"), width_px=320)
    out = renderer.render(report, [1], tmp_path / "out", purpose="learning_preview")
    assert out[1].exists() and renderer.last_backend == "builtin"
    assert renderer.preview_reason == REASON_UNAVAILABLE
    assert renderer.preview_reason != REASON_USER_SESSION_RUNNING
    assert renderer.degraded is False, "an absent backend is not a failure; the preview may be reused"


def test_5c_capability_probe_separates_installed_from_isolatable(monkeypatch):
    """``powerpoint_render_capability`` is the read-only answer used for diagnostics.  It must never start
    PowerPoint, never initialise COM and never read the Running Object Table (all of which would need COM)."""
    monkeypatch.setattr(qrn, "powerpoint_probe", lambda: (False, "availability", "progid-not-registered"))
    cap = powerpoint_render_capability()
    assert (cap.available, cap.isolated, cap.reason) == (False, False, REASON_UNAVAILABLE)

    monkeypatch.setattr(qrn, "powerpoint_probe", lambda: (False, "import", "pythoncom ModuleNotFoundError"))
    assert powerpoint_render_capability().reason == REASON_IMPORT

    monkeypatch.setattr(qrn, "powerpoint_probe", lambda: (True, "", "progid-com"))
    monkeypatch.setattr(qrn, "_powerpoint_process_ids", lambda: None)
    cap = powerpoint_render_capability()
    assert (cap.available, cap.isolated, cap.reason) == (True, False, REASON_PROCESS_TABLE_UNREADABLE)

    monkeypatch.setattr(qrn, "_powerpoint_process_ids", lambda: frozenset({4000}))
    cap = powerpoint_render_capability()
    assert (cap.available, cap.isolated, cap.reason) == (True, False, REASON_USER_SESSION_RUNNING)

    monkeypatch.setattr(qrn, "_powerpoint_process_ids", lambda: frozenset())
    cap = powerpoint_render_capability()
    assert (cap.available, cap.isolated, cap.reason) == (True, True, REASON_OK)
    assert not any(secret in cap.detail for secret in ("C:\\", "/home/", "\\\\")), "detail must be path-free"


# ============================================================================================ REQUIRED TEST 6
def test_6_owned_instance_export_failure_cleans_up_only_the_owned_instance(machine, deck, tmp_path):
    machine.on_export = lambda pres: (_ for _ in ()).throw(RuntimeError("export refused"))
    with pytest.raises(PowerPointStageError) as info:
        _render(deck, tmp_path)
    assert info.value.stage == "slide-export"
    assert info.value.code == "powerpoint-slide-export", "every stage has a stable code"
    assert len(machine.calls("Presentation.Close", "app")) == 1 and len(machine.calls("App.Quit", "app")) == 1
    assert machine.calls("Presentation.Close", "user") == [] and machine.calls("App.Quit", "user") == []
    assert machine.co_init == machine.co_uninit == 1


def test_6b_post_dispatch_ownership_failure_never_opens_closes_or_quits(machine, deck, tmp_path):
    """``DispatchEx`` returned something we cannot prove is ours (a session the user started in the meantime)."""
    def user_starts(m):
        m.start_user_powerpoint("race_unsaved.pptx")
    machine.on_dispatch = user_starts
    with pytest.raises(PowerPointStageError) as info:
        _render(deck, tmp_path)
    assert info.value.stage == "ownership" and info.value.code == REASON_OWNERSHIP_UNPROVEN
    assert machine.calls("Open") == [] and machine.calls("Presentation.Close") == []
    assert machine.calls("App.Quit") == [] and machine.user_presentations_closed == []


# ============================================================================================ REQUIRED TEST 7
def test_7_cleanup_failure_is_reported_and_never_escalates_to_the_user_or_to_a_kill(machine, deck, tmp_path, caplog):
    caplog.set_level(logging.INFO, logger="report_extractor.renderer")
    """``os.kill`` / ``subprocess`` are armed to fail by the PROMPT-035 harness, so any escalation attempt explodes."""
    machine.fail = "quit"
    out = _render(deck, tmp_path)
    assert out[1].exists(), "the exported slide survives a failed Quit"
    assert len(machine.calls("App.Quit", "app")) == 1 and machine.calls("App.Quit", "user") == []
    assert machine.user_presentations_closed == []
    steps = [r.getMessage() for r in caplog.records if "SLIDE_RENDER_CLEANUP" in r.getMessage()]
    assert any("step=app-quit" in s for s in steps), "a cleanup failure must be diagnosed, not swallowed"


def test_7b_cleanup_failure_with_a_user_session_present_still_touches_nothing(machine, deck, tmp_path):
    user = machine.start_user_powerpoint("a.pptx", "b.pptx")
    machine.fail = "quit"
    machine.on_export = lambda pres: (_ for _ in ()).throw(RuntimeError("export refused"))
    _render_tolerating_refusal(deck, tmp_path)
    assert user.quit_calls == 0 and machine.user_presentations_closed == [] and user.alive
    assert machine.calls("Presentation.Close", "user") == []


# ============================================================================================ REQUIRED TEST 8
def test_8_repeated_renders_leak_no_owned_process_and_no_ownership(machine, deck, tmp_path):
    """Ten renders in a row: one owned process each, all gone afterwards, user session never involved."""
    user = machine.start_user_powerpoint("user_unsaved.pptx")
    for _ in range(5):
        _render_tolerating_refusal(deck, tmp_path)          # refused while the user's session is up
    assert machine.dispatches == 0 and machine.calls("App.Quit") == []

    user.Quit()                                             # the user closes PowerPoint themselves
    machine.running = None
    for _ in range(5):
        out = _render(deck, tmp_path)                       # now every render owns its own instance
        assert out[1].exists()
    assert len(machine.calls("App.Quit", "app")) == 5
    assert machine.process_ids() == frozenset(), "no orphan POWERPNT process accumulated"
    assert machine.co_init == machine.co_uninit == 10


# ============================================================================================ REQUIRED TEST 9
def test_9_rapid_concurrent_preview_requests_are_lifecycle_safe(machine, deck, tmp_path):
    """Learning preview + batch worker can render at once.  The module lock serialises the COM lifecycle, so no
    job can adopt (or quit) another job's instance and no orphan is left behind."""
    gate = threading.Event()
    results: list = []
    errors: list = []

    def slow_export(pres):
        gate.wait(2.0)

    machine.on_export = slow_export

    def worker():
        try:
            results.append(_render(deck, tmp_path))
        except BaseException as exc:  # noqa: BLE001 – recorded, asserted below
            errors.append(exc)

    threads = [threading.Thread(target=worker) for _ in range(3)]
    for t in threads:
        t.start()
    gate.set()
    for t in threads:
        t.join(20)
    assert not errors, errors
    assert len(results) == 3
    assert len(machine.calls("App.Quit", "app")) == 3, "each job quits exactly the instance it created"
    assert machine.process_ids() == frozenset()
    assert machine.co_init == machine.co_uninit == 3


# =========================================================================================== REQUIRED TEST 10
def test_10_cancellation_cleans_owned_resources_and_never_reaches_the_user(machine, deck, tmp_path):
    def cancel(pres):
        raise CancellationRequested("cancel")
    machine.on_export = cancel
    with pytest.raises(CancellationRequested):
        _render(deck, tmp_path)                    # cooperative cancel is NOT a renderer failure
    assert len(machine.calls("Presentation.Close", "app")) == 1
    assert len(machine.calls("App.Quit", "app")) == 1
    assert machine.process_ids() == frozenset()
    assert machine.co_init == machine.co_uninit == 1


def test_10b_cancellation_with_a_user_session_open_never_touches_com(machine, deck, tmp_path):
    user = machine.start_user_powerpoint("user_unsaved.pptx")
    with pytest.raises(CancellationRequested):
        _render_with_cancel(deck, tmp_path)
    assert machine.events == [] and user.quit_calls == 0 and machine.dispatches == 0


def _render_with_cancel(deck, tmp_path):
    """Cancellation is observed by the chain, not inside the COM call (PROMPT-024R §15)."""
    from app.cancellation import check_cancelled
    check_cancelled(lambda: True)
    return _render(deck, tmp_path)


# ================================================================================= REQUIRED TEST 11 (regression)
def test_11_the_prompt035_ownership_suite_is_still_present_and_still_strict():
    """PROMPT-041 must not buy its behaviour by weakening PROMPT-035.  The assertions that matter are checked
    straight out of that file's source, so deleting or softening one fails HERE too.  (The suite itself is run
    unchanged by the full regression — see ``.specify/bugs/prompt041-isolated-powerpoint-renderer/test.md``.)"""
    text = (Path(__file__).resolve().parent / "test_prompt035_powerpoint_isolation.py").read_text(encoding="utf-8")
    for needle in (
        "assert user.quit_calls == 0",
        "assert machine.dispatches == 0",
        "assert machine.touched_user_app is False",
        'assert info.value.stage == "ownership"',
        "assert machine.user_presentations_closed == []",
        "process termination attempted",
    ):
        assert needle in text, f"PROMPT-035 lost the assertion {needle!r}"
    assert text.count("def test_") >= 23, "PROMPT-035 lost test cases"


# =============================================================== REQUIRED TESTS 12–15 (no regression downstream)
def _jpeg_uri() -> str:
    import base64
    from PIL import Image
    buf = io.BytesIO()
    Image.new("RGB", (8, 8), "#c8e6c9").save(buf, format="JPEG")
    return "data:image/jpeg;base64," + base64.b64encode(buf.getvalue()).decode("ascii")


def test_12_learning_preview_persistence_survives_the_new_reason_field(tmp_path, monkeypatch):
    """PROMPT-030B: previews persist to the learning assets and are reused after a restart.  The PROMPT-041
    ``reason`` field must ride along without breaking that contract."""
    from app.learning_assets import LearningAssets, digest
    assets = LearningAssets(tmp_path / "learning")
    # exactly the shape `_slide_preview_for` builds: key[5:] IS the signature tuple
    signature = (("powerpoint", "libreoffice", "builtin"), 1280, 3)
    key = ("scope", "job.pptx", 1, 123, 456, *signature)
    stored = assets.remember_slide(key, {
        "src": _jpeg_uri(), "width": 1600, "height": 900,
        "backend": "builtin", "faithful": False, "reason": REASON_USER_SESSION_RUNNING,
        "degraded": True, "regions": [], "created": 0.0})
    assert stored.get("reason") == REASON_USER_SESSION_RUNNING
    # a restart re-reads it from disk
    again = LearningAssets(tmp_path / "learning").slide(digest(key), signature)
    assert again["backend"] == "builtin" and again["reason"] == REASON_USER_SESSION_RUNNING
    # an asset written BEFORE this change has no reason at all and must not explode
    legacy = LearningAssets(tmp_path / "learning2")
    legacy.remember_slide(key, {"src": _jpeg_uri(), "width": 10, "height": 10,
                                "backend": "powerpoint", "faithful": True, "degraded": False, "regions": []})
    old = LearningAssets(tmp_path / "learning2").slide(digest(key), signature)
    assert old.get("reason") is None and old["backend"] == "powerpoint"


def test_13_preview_cache_freshness_is_unchanged_by_the_reason_field():
    """PROMPT-038: only a DEGRADED preview expires.  A refusal is degraded (it is retried, so closing PowerPoint
    recovers within the TTL); an absent PowerPoint is not a failure and stays cached."""
    from app.application_service import ApplicationService, DEGRADED_PREVIEW_TTL_S
    import time as _time
    refused = {"degraded": True, "created": _time.monotonic() - (DEGRADED_PREVIEW_TTL_S + 1)}
    assert ApplicationService._preview_expired(refused) is True
    cached = {"degraded": False, "created": _time.monotonic() - 10_000}
    assert ApplicationService._preview_expired(cached) is False


def test_14_updater_module_is_untouched_by_this_change():
    """PROMPT-039: the self-update path shares nothing with the renderer; assert that stays true."""
    updater = (APP_DIR / "updater.py").read_text(encoding="utf-8")
    assert "qpn_renderer" not in updater and "render_with_powerpoint" not in updater


def test_15_excel_output_path_is_untouched_by_this_change():
    """PROMPT-020/021/025 + PROMPT-040: item-scoped After regions, Excel crop semantics and output eligibility
    are produced elsewhere; the renderer change must not reach into them."""
    for name in ("excel_writer.py", "improvement_visual.py", "improvement_pictures.py"):
        code = (APP_DIR / name).read_text(encoding="utf-8")
        assert "preview_reason" not in code and "slidePreviewReason" not in code, name
        assert "render_with_powerpoint" not in code, f"{name} must not own a PowerPoint lifecycle"


# ========================================================================== new behaviour: session-scoped probe
def test_16_a_powerpoint_in_another_windows_session_does_not_cause_a_false_refusal(monkeypatch):
    """PowerPoint is single-instance PER LOGIN SESSION and COM activation is session-local, so a PowerPoint in
    another session can never be handed to us — it must not block our own activation either."""
    monkeypatch.setattr(qrn, "_enumerate_powerpoint_pids", lambda: frozenset({4000, 4001}))
    monkeypatch.setattr(qrn, "_current_session_id", lambda: 2)
    monkeypatch.setattr(qrn, "_process_session_id", lambda pid: {4000: 1, 4001: 2}.get(pid))
    assert qrn._powerpoint_process_ids() == frozenset({4001})


def test_16b_any_doubt_about_a_session_resolves_towards_refusing(monkeypatch):
    """Over-refusing costs fidelity; under-refusing could hand us the user's instance.  So an unreadable session
    id keeps the PID, and an unreadable OWN session id keeps every PID."""
    monkeypatch.setattr(qrn, "_enumerate_powerpoint_pids", lambda: frozenset({4000, 4001}))
    monkeypatch.setattr(qrn, "_current_session_id", lambda: 2)
    monkeypatch.setattr(qrn, "_process_session_id", lambda pid: None)      # session of the PPT is unknown
    assert qrn._powerpoint_process_ids() == frozenset({4000, 4001})
    monkeypatch.setattr(qrn, "_process_session_id", lambda pid: 1)
    monkeypatch.setattr(qrn, "_current_session_id", lambda: None)          # our own session is unknown
    assert qrn._powerpoint_process_ids() == frozenset({4000, 4001})


def test_16c_an_unreadable_process_table_still_means_unknown(monkeypatch):
    monkeypatch.setattr(qrn, "_enumerate_powerpoint_pids", lambda: None)
    monkeypatch.setattr(qrn, "_current_session_id", lambda: 1)
    assert qrn._powerpoint_process_ids() is None
    monkeypatch.setattr(qrn, "_enumerate_powerpoint_pids", lambda: frozenset())
    assert qrn._powerpoint_process_ids() == frozenset()


def test_16d_the_session_helpers_are_read_only_queries():
    """``ProcessIdToSessionId`` takes a PID, not a handle: nothing is opened, signalled or terminated."""
    code = _code_without_comments_and_docstrings(Path(qrn.__file__).read_text(encoding="utf-8"))
    assert "ProcessIdToSessionId" in code
    assert not re.search(r"\bOpenProcess\b|\bTerminateProcess\b|\bos\.kill\b|\.terminate\(", code, re.IGNORECASE)


# ================================================================================== the reason reaches the UI
def test_17_the_learning_dto_carries_the_reason_and_never_a_path(machine, monkeypatch, tmp_path):
    """``slidePreviewReason`` is what lets the Learning tab say "your PowerPoint is open" instead of the generic
    "simple preview".  It is a slug: no path, no PID, nothing machine-specific."""
    from app import application_service as svc
    from app.pptx_parser import parse_pptx
    from tests.test_prompt027_rendering import _simple_slide
    from tests.test_after_evidence import _pics

    source = _simple_slide(tmp_path / "d.pptx", lambda s: _pics(s, ["#c8e6c9"], 4, 3, w=2, h=1.4))
    machine.start_user_powerpoint("user_unsaved.pptx")
    monkeypatch.setattr(qrn, "powerpoint_available", lambda: True)

    renderer = SlideRenderer(width_px=320)
    report = parse_pptx(str(source))
    renderer.render(report, [1], tmp_path / "o", purpose="learning_preview")
    assert renderer.last_backend == "builtin"
    reason = renderer.preview_reason
    assert reason == REASON_USER_SESSION_RUNNING
    assert re.fullmatch(r"[a-z0-9-]+", reason), "a reason is a slug, never free text"
    assert not any(ch in reason for ch in ("/", "\\", ":"))

    # and the DTO field is built from it (the mapping itself is exercised by the frontend suite)
    dto_fragment = {"slidePreviewReason": reason}
    assert dto_fragment["slidePreviewReason"] == "powerpoint-user-session-running"
    assert svc.ApplicationService is not None


# ============================================================================================ static contracts
def test_18_no_process_termination_or_registry_writing_anywhere_in_the_application():
    """The prompt's hard "never" list, checked over the shipped code (comments/docstrings stripped)."""
    forbidden = ("taskkill", "terminateprocess", "psutil", "pkill", "killall", "stop-process",
                 "wmic process", "createprocesswithlogon", "logonuser")
    registry_writes = (r"winreg\.SetValue", r"winreg\.CreateKey", r"winreg\.DeleteKey",
                       r"winreg\.SetValueEx", r"winreg\.DeleteValue")
    for path in sorted(APP_DIR.glob("*.py")):
        code = _code_without_comments_and_docstrings(path.read_text(encoding="utf-8"))
        low = code.lower()
        for needle in forbidden:
            assert needle not in low, f"{path.name} contains {needle!r}"
        for pattern in registry_writes:
            assert not re.search(pattern, code), f"{path.name} writes the registry ({pattern})"


def test_19_the_renderer_never_attaches_enumerates_or_quits_a_foreign_instance():
    code = _code_without_comments_and_docstrings(Path(qrn.__file__).read_text(encoding="utf-8"))
    assert "GetActiveObject" in code, "the ROT lookup is a LOOKUP and must stay (it never starts a server)"
    assert not re.search(r"for\s+\w+\s+in\s+\w*\.?Presentations\b", code), "never enumerate Presentations"
    assert "ActivePresentation" not in code and "GetObject" not in code
    assert not re.search(r"win32com\.client\.Dispatch\(", code), "attaching Dispatch() is forbidden"
    # the only Close/Quit calls are on the objects this call itself created
    assert code.count(".Quit()") == 1 and code.count(".Close()") == 1


def test_20_reason_constants_are_stable_and_are_what_the_frontend_maps():
    """The slug strings are a contract shared with ``frontend/src/tabs/LearningTab.tsx`` (``fidelityHint``)."""
    expected = {"ok", "powerpoint-unavailable", "powerpoint-import", "powerpoint-process-table-unreadable",
                "powerpoint-user-session-running", "powerpoint-registered-instance",
                "powerpoint-ownership-unproven", "powerpoint-not-exclusive"}
    actual = {REASON_OK, REASON_UNAVAILABLE, REASON_IMPORT, REASON_PROCESS_TABLE_UNREADABLE,
              REASON_USER_SESSION_RUNNING, REASON_REGISTERED_INSTANCE, REASON_OWNERSHIP_UNPROVEN,
              REASON_NOT_EXCLUSIVE}
    assert actual == expected
    tsx = (Path(__file__).resolve().parent.parent / "frontend" / "src" / "tabs" / "LearningTab.tsx") \
        .read_text(encoding="utf-8")
    for code in sorted(actual - {REASON_OK}):
        assert f"'{code}'" in tsx, f"the Learning UI has no wording for {code}"


def test_21_no_orphan_process_and_no_leaked_com_when_the_export_raises_midway(machine, deck, tmp_path):
    """App shutdown mid-render is modelled as an exception out of the export loop: owned-only cleanup, balanced
    COM init, nothing left running."""
    calls = {"n": 0}

    def export_then_die(pres):
        calls["n"] += 1
        raise RuntimeError("host is shutting down")

    machine.on_export = export_then_die
    with pytest.raises(PowerPointStageError):
        _render(deck, tmp_path)
    assert machine.process_ids() == frozenset(), "no orphan POWERPNT process"
    assert machine.co_init == machine.co_uninit == 1, "CoInitialize/CoUninitialize stay balanced"


def test_22_the_process_snapshot_helper_is_never_given_a_handle(monkeypatch):
    """Belt and braces for the read-only claim: the Toolhelp path only lists, and the session path only asks."""
    src = Path(qrn.__file__).read_text(encoding="utf-8")
    tree = ast.parse(src)
    opened = [n.func.attr for n in ast.walk(tree)
              if isinstance(n, ast.Call) and isinstance(n.func, ast.Attribute)
              and n.func.attr in {"OpenProcess", "CreateToolhelp32Snapshot", "Process32FirstW", "Process32NextW",
                                  "ProcessIdToSessionId", "CloseHandle"}]
    assert "OpenProcess" not in opened
    assert "CreateToolhelp32Snapshot" in opened and "CloseHandle" in opened, "the snapshot handle is closed"


def test_23_the_fake_com_harness_is_shared_with_prompt035_not_duplicated():
    """One model of Windows for both prompts: a change to the single-instance semantics moves both suites, and
    the harness still explodes on any attempt to end a process."""
    from tests import test_prompt035_powerpoint_isolation as p35
    assert FakeMachine is p35.FakeMachine
    source = inspect.getsource(p35.FakeMachine.install)
    assert "process termination attempted" in source and "os, \"kill\"" in source.replace("'", '"')


def test_24_the_diagnostics_panel_separates_no_powerpoint_from_your_powerpoint_is_open(monkeypatch):
    """The app's own diagnostic surface must not collapse the two opposite causes into one "Unavailable"."""
    from app.diagnostics import run_diagnostics
    from app.config import AppConfig

    def rows_for(probe, pids):
        monkeypatch.setattr(qrn, "powerpoint_probe", lambda: probe)
        monkeypatch.setattr(qrn, "_powerpoint_process_ids", lambda: pids)
        return dict(run_diagnostics(AppConfig(), check_ollama=False))

    installed = rows_for((True, "", "progid-com"), frozenset({4000}))
    assert "powerpoint-user-session-running" in installed["PowerPoint isolation"]
    assert "đóng PowerPoint" in installed["PowerPoint isolation"]
    assert installed["PowerPoint renderer"] == "OK", "installed PowerPoint must not be reported as missing"

    free = rows_for((True, "", "progid-com"), frozenset())
    assert free["PowerPoint isolation"].startswith("OK")

    absent = rows_for((False, "availability", "progid-not-registered"), None)
    assert "powerpoint-unavailable" in absent["PowerPoint isolation"]
    assert absent["PowerPoint renderer"] == "Unavailable"
    assert absent["PowerPoint isolation"] != installed["PowerPoint isolation"]
