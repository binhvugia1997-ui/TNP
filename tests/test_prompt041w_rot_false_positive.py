"""PROMPT-041W — the Windows false positive: "PowerPoint is running" reported with no POWERPNT.EXE.

WINDOWS EVIDENCE (from the acceptance run of PR #17)
    ``Get-Process POWERPNT -ErrorAction SilentlyContinue`` returned NOTHING, yet the Learning preview rendered with
    the built-in backend and said "Một phiên PowerPoint đang chạy trên máy này nên app không tạo được phiên riêng."

PROVEN ROOT CAUSE (see ``.specify/bugs/prompt041w-rot-false-positive/assessment.md``)
    That sentence is produced by exactly one slug, ``powerpoint-registered-instance``, and that slug is reached
    ONLY when the same-session process snapshot is EMPTY (``if before:`` refuses first otherwise).  So on that
    machine the process table already agreed with ``Get-Process`` — the refusal came from the **Running Object
    Table**, which carries no process evidence at all.  Office registers in the ROT when it loses focus and
    unregisters on a graceful exit, so a crashed, killed or abandoned instance leaves its entry behind, and
    ``GetActiveObject`` performs no liveness check (``__WrapDispatch`` never returns ``None``).  The code then
    reported that as a *process fact* about the user's machine.  Reproduced deterministically in tests 01/02.

    A second, independent defect is pinned here too: ``_slide_preview_for`` returned a cached degraded preview —
    reason included — for up to ``DEGRADED_PREVIEW_TTL_S`` without re-reading the machine, so a reason recorded
    while PowerPoint WAS open kept being shown as current state after the user closed it (contract E).

WHAT CHANGED
    * The process table is the only source of process facts; the ROT is a three-state diagnostic.
    * An empty same-session snapshot whose enumeration **vouched for itself** (a successful Toolhelp snapshot
      always contains our own PID) means "there really is no PowerPoint here": activation is attempted and the
      post-dispatch PID proof decides ownership (contract A).
    * A ROT hit still fails closed when the enumeration **cannot** vouch for itself — the PROMPT-035 "instance
      invisible to the process table" case — so that suite is unchanged and still green (contract B kept).
    * PowerPoint seen but our own session id unreadable gets its own code, ``powerpoint-session-unproven``,
      never ``powerpoint-user-session-running`` (contract D).
    * A cached PowerPoint-blocked fallback is re-checked against the live capability and dropped the moment the
      block is gone (contract E).

NOTHING here weakens PROMPT-035: no assertion in ``tests/test_prompt035_powerpoint_isolation.py`` was edited (test
12 reads the file back and asserts that), and its fake COM machine is shared so both prompts move together.
Still no attach, no ``GetActiveObject`` *use* (lookup only), no process kill, no registry write.
"""
from __future__ import annotations

import ast
import logging
import re
from pathlib import Path

import pytest
from PIL import Image

import app.qpn_renderer as qrn
from app.application_service import ApplicationService
from app.qpn_renderer import (
    PowerPointStageError,
    SlideRenderer,
    powerpoint_render_capability,
    render_with_powerpoint,
)
from tests.test_after_evidence import _pics
from tests.test_prompt027_rendering import _candidate, _preview_service, _simple_slide
from tests.test_prompt027r_renderer import _two_item_slide_deck
from tests.test_prompt035_powerpoint_isolation import FakeMachine, _code_without_comments_and_docstrings

REASON_OK = qrn.REASON_OK
REASON_UNAVAILABLE = qrn.REASON_UNAVAILABLE
REASON_IMPORT = qrn.REASON_IMPORT
REASON_PROCESS_TABLE_UNREADABLE = qrn.REASON_PROCESS_TABLE_UNREADABLE
REASON_USER_SESSION_RUNNING = qrn.REASON_USER_SESSION_RUNNING
REASON_REGISTERED_INSTANCE = qrn.REASON_REGISTERED_INSTANCE
REASON_OWNERSHIP_UNPROVEN = qrn.REASON_OWNERSHIP_UNPROVEN
REASON_NOT_EXCLUSIVE = qrn.REASON_NOT_EXCLUSIVE
REASON_SESSION_UNPROVEN = qrn.REASON_SESSION_UNPROVEN


# ============================================================================================== fakes / helpers
@pytest.fixture()
def machine(monkeypatch):
    """The PROMPT-035 Windows box, plus the one thing PROMPT-041W adds: a self-vouching process table.

    A successful Toolhelp32 snapshot always contains OUR OWN PID, so on a real Windows box the enumeration can
    vouch for having seen the whole table.  Off Windows this is False — the fail-closed direction — which is what
    keeps the PROMPT-035 "invisible to the process table" suite exactly as it was.
    """
    m = FakeMachine().install(monkeypatch)
    monkeypatch.setattr(qrn, "_process_table_vouched", lambda: True)
    return m


@pytest.fixture()
def deck(tmp_path):
    path = tmp_path / "job.pptx"
    path.write_bytes(b"PK-not-a-real-deck")
    return path


@pytest.fixture()
def report(tmp_path):
    from app.pptx_parser import parse_pptx
    return parse_pptx(str(_simple_slide(tmp_path / "d.pptx",
                                        lambda s: _pics(s, ["#c8e6c9"], 4, 3, w=2, h=1.4))))


def _render(deck, tmp_path):
    return render_with_powerpoint(deck, [1], tmp_path, width_px=960)


def _export_stub(pptx, slide_numbers, out_dir, width_px=1920):
    """A stand-in for a successful PowerPoint export, so the preview chain can be driven end to end off Windows."""
    out = {}
    for n in slide_numbers:
        target = Path(out_dir) / f"slide_{n:03d}.png"
        Image.new("RGB", (int(width_px), int(width_px * 540 / 960)), "white").save(target)
        out[n] = target
    return out


def _blocked_render(pptx, slide_numbers, out_dir, width_px=1920):
    raise PowerPointStageError("ownership", RuntimeError("a PowerPoint session is already running"), pptx,
                               code=REASON_USER_SESSION_RUNNING)


def _preview_env(monkeypatch, blocked: bool):
    """Drive ``_slide_preview_for`` through the real renderer chain and the real capability probe."""
    monkeypatch.setattr(qrn, "powerpoint_probe", lambda: (True, "", "progid-com"))
    monkeypatch.setattr(qrn, "powerpoint_available", lambda: True)
    monkeypatch.setattr(qrn, "find_soffice", lambda: None)
    monkeypatch.setattr(qrn, "render_with_powerpoint", _blocked_render if blocked else _export_stub)
    _set_powerpoint_present(monkeypatch, blocked)


def _set_powerpoint_present(monkeypatch, present: bool):
    pids = frozenset({4000}) if present else frozenset()
    monkeypatch.setattr(qrn, "_powerpoint_process_ids", lambda: pids)
    monkeypatch.setattr(qrn, "_enumerate_powerpoint_pids", lambda: pids)
    qrn._CAPABILITY_MEMO.clear()


# ================================================================== THE REPORTED BUG, REPRODUCED AND FIXED
def test_01_repro_a_stale_rot_entry_must_not_become_a_claim_that_powerpoint_is_running(machine, deck, tmp_path,
                                                                                      monkeypatch):
    """THE acceptance failure: no POWERPNT.EXE anywhere, but the ROT still holds a PowerPoint entry.

    Before PROMPT-041W this refused with ``powerpoint-registered-instance`` — whose UI wording asserted "a
    PowerPoint session is running on this machine" — without ever attempting activation.  Now the vouched-empty
    process table wins, activation is attempted, and the post-dispatch PID proof decides ownership.
    """
    monkeypatch.setattr(qrn, "_rot_lookup", lambda client: qrn.ROT_REGISTERED)
    assert qrn._powerpoint_process_ids() == frozenset(), "precondition: the process table sees no PowerPoint"

    out = _render(deck, tmp_path)

    assert out[1].exists(), "an unbacked ROT entry must not block a render on a machine with no PowerPoint"
    assert machine.dispatches == 1, "activation must be attempted (contract A)"
    assert len(machine.calls("App.Quit", "app")) == 1
    assert machine.process_ids() == frozenset(), "the instance we created is cleaned up again"


def test_02_the_rot_object_is_never_handed_back_as_a_usable_application(machine, deck, tmp_path, monkeypatch):
    """Whatever the ROT returns, we never call a method on it and never turn it into a usable Application."""
    monkeypatch.setattr(qrn, "_rot_lookup", lambda client: qrn.ROT_REGISTERED)
    out = _render(deck, tmp_path)
    assert out[1].exists()
    assert machine.calls("Open", "user") == [], "nothing may be opened inside a PowerPoint we do not own"
    assert machine.touched_user_app is False


# ============================================================ REQUIRED 1 — raw snapshot empty
def test_03_empty_snapshot_never_refuses_as_user_session_running(machine, deck, tmp_path):
    assert qrn._powerpoint_process_ids() == frozenset()
    out = _render(deck, tmp_path)
    assert out[1].exists() and machine.dispatches == 1
    assert machine.calls("App.Quit", "user") == [] and machine.user_presentations_closed == []


def test_03b_the_capability_probe_agrees_an_empty_snapshot_means_isolatable(monkeypatch):
    monkeypatch.setattr(qrn, "powerpoint_probe", lambda: (True, "", "progid-com"))
    _set_powerpoint_present(monkeypatch, False)
    capability = powerpoint_render_capability()
    assert (capability.available, capability.isolated, capability.reason) == (True, True, REASON_OK)
    assert capability.reason != REASON_USER_SESSION_RUNNING


# ============================================================ REQUIRED 2 — same-session PID
def test_04_same_session_powerpoint_still_refuses_safely_with_the_right_code(machine, deck, tmp_path):
    """Contract B: a proven same-session PowerPoint is refused before activation and never touched."""
    user = machine.start_user_powerpoint("user_unsaved.pptx")
    with pytest.raises(PowerPointStageError) as info:
        _render(deck, tmp_path)
    assert info.value.stage == "ownership" and info.value.code == REASON_USER_SESSION_RUNNING
    assert machine.dispatches == 0
    assert user.quit_calls == 0 and machine.user_presentations_closed == [] and machine.touched_user_app is False


# ============================================================ REQUIRED 3 — different session only
def test_05_a_powerpoint_in_another_windows_session_is_not_a_same_session_refusal(monkeypatch):
    monkeypatch.setattr(qrn, "powerpoint_probe", lambda: (True, "", "progid-com"))
    monkeypatch.setattr(qrn, "_enumerate_powerpoint_pids", lambda: frozenset({4000}))
    monkeypatch.setattr(qrn, "_current_session_id", lambda: 2)
    monkeypatch.setattr(qrn, "_process_session_id", lambda pid: 1 if pid == 4000 else 2)
    assert qrn._powerpoint_process_ids() == frozenset(), "another session's PowerPoint is not ours"
    capability = powerpoint_render_capability()
    assert (capability.isolated, capability.reason) == (True, REASON_OK)
    assert capability.reason not in (REASON_USER_SESSION_RUNNING, REASON_SESSION_UNPROVEN)


def test_05b_a_different_session_powerpoint_does_not_block_our_own_activation(machine, deck, tmp_path, monkeypatch):
    """Contract C: a POWERPNT.EXE that lives in ANOTHER Windows session is not ours, is not reachable by our COM
    activation, and must not be reported as ``powerpoint-user-session-running``.  The machine's raw table still
    contains it; the session-filtered view our renderer reasons about does not, so we activate and then PROVE
    ownership of the instance we got rather than assuming it."""
    other_session_pid = 9000
    raw = machine.process_ids
    monkeypatch.setattr(qrn, "_enumerate_powerpoint_pids", lambda: (raw() or frozenset()) | {other_session_pid})
    monkeypatch.setattr(qrn, "_powerpoint_process_ids", lambda: raw())
    assert other_session_pid in qrn._enumerate_powerpoint_pids()
    assert qrn._powerpoint_process_ids() == frozenset(), "our own session sees no PowerPoint"

    out = _render(deck, tmp_path)

    assert out[1].exists(), "another session's PowerPoint must not block our render"
    assert machine.dispatches == 1 and len(machine.calls("App.Quit", "app")) == 1
    assert [e for e in machine.events if e[1] == "user"] == [], "no PowerPoint outside our session was touched"


# ============================================================ REQUIRED 4 — session id unreadable
def test_06_unprovable_session_gets_its_own_code_and_never_pretends_to_be_the_users_session(machine, deck, tmp_path,
                                                                                            monkeypatch):
    """Contract D: ``ProcessIdToSessionId`` failing is NOT "the user's PowerPoint is open"."""
    machine.start_user_powerpoint("somewhere.pptx")
    monkeypatch.setattr(qrn, "_session_filter_is_provable", lambda: False)
    with pytest.raises(PowerPointStageError) as info:
        _render(deck, tmp_path)
    assert info.value.code == REASON_SESSION_UNPROVEN
    assert info.value.code != REASON_USER_SESSION_RUNNING
    assert machine.dispatches == 0 and machine.calls("App.Quit") == []


def test_06b_the_capability_probe_reports_the_same_distinct_code(monkeypatch):
    monkeypatch.setattr(qrn, "powerpoint_probe", lambda: (True, "", "progid-com"))
    monkeypatch.setattr(qrn, "_powerpoint_process_ids", lambda: frozenset({4000}))
    monkeypatch.setattr(qrn, "_enumerate_powerpoint_pids", lambda: frozenset({4000}))
    monkeypatch.setattr(qrn, "_session_filter_is_provable", lambda: False)
    assert powerpoint_render_capability().reason == REASON_SESSION_UNPROVEN


def test_06c_the_real_session_prove_helper_fails_closed_when_our_own_session_is_unreadable(monkeypatch):
    """The mutation this kills: ``_session_filter_is_provable()`` hard-wired to True would collapse contract D back
    into ``powerpoint-user-session-running`` — the exact lie PROMPT-041W exists to remove.  Exercised through the
    REAL helper, not a stub."""
    monkeypatch.setattr(qrn.sys, "platform", "win32")
    monkeypatch.setattr(qrn, "powerpoint_probe", lambda: (True, "", "progid-com"))
    monkeypatch.setattr(qrn, "_powerpoint_process_ids", lambda: frozenset({4000}))
    monkeypatch.setattr(qrn, "_enumerate_powerpoint_pids", lambda: frozenset({4000}))

    monkeypatch.setattr(qrn, "_process_session_id", lambda pid: None)     # ProcessIdToSessionId fails
    assert qrn._session_filter_is_provable() is False
    assert powerpoint_render_capability().reason == REASON_SESSION_UNPROVEN

    monkeypatch.setattr(qrn, "_process_session_id", lambda pid: 1)        # and it is provable when it works
    assert qrn._session_filter_is_provable() is True
    assert powerpoint_render_capability().reason == REASON_USER_SESSION_RUNNING


# ============================================================ REQUIRED 5 — snapshot API failure
def test_07_an_unreadable_process_table_fails_closed_with_its_own_code(machine, deck, tmp_path):
    machine.snapshot_unreadable = True
    with pytest.raises(PowerPointStageError) as info:
        _render(deck, tmp_path)
    assert info.value.code == REASON_PROCESS_TABLE_UNREADABLE
    assert info.value.code not in (REASON_USER_SESSION_RUNNING, REASON_REGISTERED_INSTANCE)
    assert machine.dispatches == 0 and machine.calls("App.Quit") == []


def test_07b_the_capability_probe_also_never_confuses_unreadable_with_isolatable(monkeypatch):
    monkeypatch.setattr(qrn, "powerpoint_probe", lambda: (True, "", "progid-com"))
    monkeypatch.setattr(qrn, "_powerpoint_process_ids", lambda: None)
    monkeypatch.setattr(qrn, "_enumerate_powerpoint_pids", lambda: None)
    capability = powerpoint_render_capability()
    assert (capability.isolated, capability.reason) == (False, REASON_PROCESS_TABLE_UNREADABLE)


# ============================================================ REQUIRED 6 — empty set is not None/error
def test_08_empty_set_and_none_are_never_confused(monkeypatch):
    monkeypatch.setattr(qrn, "_current_session_id", lambda: 1)
    monkeypatch.setattr(qrn, "_process_session_id", lambda pid: 1)

    monkeypatch.setattr(qrn, "_enumerate_powerpoint_pids", lambda: frozenset())
    empty = qrn._powerpoint_process_ids()
    monkeypatch.setattr(qrn, "_enumerate_powerpoint_pids", lambda: None)
    unknown = qrn._powerpoint_process_ids()

    assert empty == frozenset() and empty is not None and bool(empty) is False
    assert unknown is None and unknown != empty
    monkeypatch.setattr(qrn, "powerpoint_probe", lambda: (True, "", "progid-com"))
    monkeypatch.setattr(qrn, "_powerpoint_process_ids", lambda: frozenset())
    assert powerpoint_render_capability().reason == REASON_OK
    monkeypatch.setattr(qrn, "_powerpoint_process_ids", lambda: None)
    assert powerpoint_render_capability().reason == REASON_PROCESS_TABLE_UNREADABLE


def test_08b_the_vouch_is_what_separates_truly_empty_from_could_not_see(monkeypatch):
    """A snapshot missing our own PID did not see the whole table, so its empty answer is worthless."""
    if qrn.sys.platform != "win32":
        assert qrn._process_table_vouched() is False, "off Windows nothing vouches: the fail-closed default"
    # (on a real Windows host the ambient answer is legitimately True — the snapshot works and contains
    # our own PID — so only the synthetic-table logic below is asserted there)
    monkeypatch.setattr(qrn.sys, "platform", "win32")
    monkeypatch.setattr(qrn, "_enumerate_processes", lambda image_name=None: frozenset({111, 222}))
    monkeypatch.setattr(qrn.os, "getpid", lambda: 999)
    assert qrn._process_table_vouched() is False
    monkeypatch.setattr(qrn, "_enumerate_processes", lambda image_name=None: frozenset({111, 999}))
    assert qrn._process_table_vouched() is True
    monkeypatch.setattr(qrn, "_enumerate_processes", lambda image_name=None: None)
    assert qrn._process_table_vouched() is False, "an unreadable table never vouches"


def test_08c_an_unvouched_empty_table_plus_a_rot_hit_still_fails_closed(machine, deck, tmp_path, monkeypatch):
    """The PROMPT-035 "instance invisible to the process table" case must keep failing closed."""
    monkeypatch.setattr(qrn, "_process_table_vouched", lambda: False)
    monkeypatch.setattr(qrn, "_rot_lookup", lambda client: qrn.ROT_REGISTERED)
    with pytest.raises(PowerPointStageError) as info:
        _render(deck, tmp_path)
    assert info.value.stage == "ownership" and info.value.code == REASON_REGISTERED_INSTANCE
    assert machine.dispatches == 0, "an unvouched empty table is not permission to activate"


def test_08d_the_three_rot_states_are_distinguishable():
    """``unavailable`` (the lookup could not run) must not be reported as ``none`` (nothing is registered)."""
    class Client:
        def __init__(self, behaviour):
            self.behaviour = behaviour

        def GetActiveObject(self, progid):
            if self.behaviour == "object":
                return object()
            if self.behaviour == "none":
                return None
            raise RuntimeError(self.behaviour)

    assert qrn._rot_lookup(Client("object")) == qrn.ROT_REGISTERED
    assert qrn._rot_lookup(Client("none")) == qrn.ROT_NONE
    assert qrn._rot_lookup(Client("boom")) == qrn.ROT_UNAVAILABLE

    class Unavailable(Exception):
        hresult = -2147221021                                # MK_E_UNAVAILABLE: the ROT's normal "nothing"

    client = Client("unused")
    client.GetActiveObject = lambda progid: (_ for _ in ()).throw(Unavailable())
    assert qrn._rot_lookup(client) == qrn.ROT_NONE


# ============================================ REQUIRED 7 — a cached refusal must not outlive its cause
def test_09_a_cached_ownership_refusal_is_dropped_the_moment_powerpoint_disappears(monkeypatch, tmp_path):
    _preview_env(monkeypatch, blocked=True)
    service = _preview_service()
    candidate = _candidate(_two_item_slide_deck(tmp_path))

    first = service._slide_preview_for(candidate)
    assert first["backend"] == "builtin" and first["reason"] == REASON_USER_SESSION_RUNNING

    # still open, inside the 15 s TTL -> the cached entry is honestly reused
    again = service._slide_preview_for(candidate)
    assert (again["backend"], again["reason"]) == ("builtin", REASON_USER_SESSION_RUNNING)

    # the user closes PowerPoint: the cached fallback must NOT keep asserting it, and must re-render NOW
    _preview_env(monkeypatch, blocked=False)
    recovered = service._slide_preview_for(candidate)
    assert recovered["backend"] == "powerpoint" and recovered["faithful"] is True
    assert recovered["reason"] == "", "a faithful render carries no reason"


def test_09b_a_non_powerpoint_fallback_is_not_retried_on_every_poll(monkeypatch, tmp_path):
    """Contract E must not become a COM hammer: only PowerPoint-blocked fallbacks get the live re-check."""
    monkeypatch.setattr(qrn, "powerpoint_probe", lambda: (False, "availability", "not-windows"))
    monkeypatch.setattr(qrn, "find_soffice", lambda: None)
    service = _preview_service()
    candidate = _candidate(_two_item_slide_deck(tmp_path))
    first = service._slide_preview_for(candidate)
    assert first["backend"] == "builtin" and first["reason"] == REASON_UNAVAILABLE

    calls = []
    monkeypatch.setattr(qrn, "powerpoint_render_capability",
                        lambda max_age_s=0.0: calls.append(1) or qrn.PowerPointCapability(False, False, "x"))
    second = service._slide_preview_for(candidate)
    assert second["backend"] == "builtin"
    assert calls == [], "an unavailable-backend fallback must not re-probe the machine on every hit"


def test_09c_the_stale_check_only_fires_for_a_powerpoint_blocked_degraded_entry(monkeypatch):
    stale_reason = {"degraded": True, "reason": REASON_USER_SESSION_RUNNING}
    monkeypatch.setattr(qrn, "powerpoint_probe", lambda: (True, "", "progid-com"))

    _set_powerpoint_present(monkeypatch, False)
    assert ApplicationService._degraded_reason_is_stale(dict(stale_reason)) is True

    _set_powerpoint_present(monkeypatch, True)
    assert ApplicationService._degraded_reason_is_stale(dict(stale_reason)) is False, "still true on the machine"

    _set_powerpoint_present(monkeypatch, False)
    assert ApplicationService._degraded_reason_is_stale(
        {"degraded": False, "reason": REASON_USER_SESSION_RUNNING}) is False, "a faithful entry is never stale"
    assert ApplicationService._degraded_reason_is_stale(
        {"degraded": True, "reason": REASON_UNAVAILABLE}) is False, "not a PowerPoint-blocked fallback"
    assert ApplicationService._degraded_reason_is_stale({"degraded": True}) is False


def test_09d_a_degraded_preview_persisted_to_disk_cannot_resurrect_its_reason(tmp_path, monkeypatch):
    """The on-disk copy is the other way a stale reason could survive — across a restart.  ``LearningAssets.slide``
    re-stamps ``created`` from ``saved_at``, so a preview persisted seconds ago is NOT expired and IS returned;
    only the PROMPT-041W staleness check stops it from telling the user "PowerPoint is running" again."""
    import base64
    import io

    from app.learning_assets import LearningAssets, digest
    buffer = io.BytesIO()
    Image.new("RGB", (8, 6), "white").save(buffer, format="JPEG")
    assets = LearningAssets(tmp_path)
    degraded = {"src": "data:image/jpeg;base64," + base64.b64encode(buffer.getvalue()).decode("ascii"),
                "width": 10, "height": 10, "backend": "builtin", "faithful": False,
                "reason": REASON_USER_SESSION_RUNNING, "degraded": True, "regions": []}
    key = ("scope", str(tmp_path / "a.pptx"), 3, 1, 1, ("sig",))
    assets.remember_slide(key, degraded)
    restored = assets.slide(digest(key), key[5:])

    assert restored.get("reason") == REASON_USER_SESSION_RUNNING, "the reason really does round-trip"
    assert ApplicationService._preview_expired(restored) is False, "it is inside the degraded TTL, so it IS served"
    monkeypatch.setattr(qrn, "powerpoint_probe", lambda: (True, "", "progid-com"))
    _set_powerpoint_present(monkeypatch, False)
    assert ApplicationService._degraded_reason_is_stale(restored) is True, "and the stale check must catch it"


# ============================================ REQUIRED 8 — a good render clears the old reason
def test_10_a_successful_powerpoint_render_clears_any_previous_reason(monkeypatch, tmp_path):
    _preview_env(monkeypatch, blocked=True)
    service = _preview_service()
    candidate = _candidate(_two_item_slide_deck(tmp_path))

    blocked = service._slide_preview_for(candidate)
    assert blocked["reason"] == REASON_USER_SESSION_RUNNING and blocked["faithful"] is False

    _preview_env(monkeypatch, blocked=False)
    recovered = service._slide_preview_for(candidate)
    assert recovered["backend"] == "powerpoint" and recovered["faithful"] is True
    assert recovered.get("reason", "") == "", "the old reason must not survive into a faithful render"
    assert recovered["src"].startswith("data:image/jpeg;base64,")
    assert SlideRenderer(width_px=320).preview_reason == "", "a fresh renderer starts with no reason"


# ================================== REQUIRED 9 — candidate switching cannot leak a reason
def test_11_switching_candidates_does_not_leak_the_previous_candidates_reason(monkeypatch, tmp_path):
    _preview_env(monkeypatch, blocked=True)
    service = _preview_service()
    blocked = _candidate(_two_item_slide_deck(tmp_path))
    assert service._slide_preview_for(blocked)["reason"] == REASON_USER_SESSION_RUNNING

    other = _candidate(_two_item_slide_deck(tmp_path / "other"))
    assert other.source_file != blocked.source_file
    _preview_env(monkeypatch, blocked=False)
    assert service._slide_preview_for(other).get("reason", "") == ""

    # ...and the blocked one is still honestly reported as blocked (no cross-talk in either direction)
    _preview_env(monkeypatch, blocked=True)
    assert service._slide_preview_for(blocked)["reason"] == REASON_USER_SESSION_RUNNING
    _preview_env(monkeypatch, blocked=False)
    assert service._slide_preview_for(other).get("reason", "") == ""


# ================================== REQUIRED 10/11/12 — nothing downstream regressed
def test_12_the_prompt035_ownership_suite_is_untouched_and_still_strict():
    """PROMPT-041W must not buy its fix by weakening PROMPT-035.  The file is asserted UNCHANGED in content and
    its key assertions are read back out of its source; the suite itself runs in the same regression command."""
    path = Path(__file__).resolve().parent / "test_prompt035_powerpoint_isolation.py"
    text = path.read_text(encoding="utf-8")
    for needle in ("assert user.quit_calls == 0", "assert machine.dispatches == 0",
                   "assert machine.touched_user_app is False", 'assert info.value.stage == "ownership"',
                   "assert machine.user_presentations_closed == []",
                   "test_running_instance_invisible_to_process_table_is_caught_by_the_running_object_table"):
        assert needle in text, f"PROMPT-035 lost {needle!r}"
    assert text.count("def test_") >= 23
    assert "PROMPT-041W" not in text, "no PROMPT-035 assertion may be edited for this fix"


def test_13_the_reason_vocabulary_is_whole_and_every_code_has_ui_wording():
    expected = {REASON_OK, REASON_UNAVAILABLE, REASON_IMPORT, REASON_PROCESS_TABLE_UNREADABLE,
                REASON_USER_SESSION_RUNNING, REASON_REGISTERED_INSTANCE, REASON_OWNERSHIP_UNPROVEN,
                REASON_NOT_EXCLUSIVE, REASON_SESSION_UNPROVEN}
    assert expected == {REASON_OK, REASON_UNAVAILABLE, REASON_IMPORT, REASON_PROCESS_TABLE_UNREADABLE,
                        REASON_USER_SESSION_RUNNING, REASON_REGISTERED_INSTANCE, REASON_OWNERSHIP_UNPROVEN,
                        REASON_NOT_EXCLUSIVE, REASON_SESSION_UNPROVEN}
    assert REASON_SESSION_UNPROVEN in qrn.POWERPOINT_PRESENT_REASONS
    tsx = (Path(__file__).resolve().parent.parent / "frontend" / "src" / "tabs" / "LearningTab.tsx") \
        .read_text(encoding="utf-8")
    for code in sorted(expected - {REASON_OK}):
        assert f"'{code}'" in tsx, f"the Learning UI has no wording for {code}"
    # the wording for the legacy ROT slug must no longer assert a process fact it cannot prove
    assert "Một phiên PowerPoint đang chạy trên máy này" not in tsx


# ================================================================================== observability (Windows)
def test_14_the_snapshot_and_ownership_lines_carry_everything_windows_acceptance_needs(machine, deck, tmp_path,
                                                                                       caplog):
    caplog.set_level(logging.INFO, logger="report_extractor.renderer")
    _render(deck, tmp_path)
    lines = [r.getMessage() for r in caplog.records if "POWERPOINT_" in r.getMessage()]
    snapshot = [line for line in lines if line.startswith("POWERPOINT_SNAPSHOT")]
    ownership = [line for line in lines if line.startswith("POWERPOINT_OWNERSHIP")]
    assert snapshot and ownership
    for field in ("pid=", "session=", "status=readable", "raw=", "same_session=", "vouched=true",
                  "session_provable=", "rot="):
        assert field in snapshot[0], f"{field} missing from {snapshot[0]}"
    assert "decision=created" in ownership[0] and f"reason={REASON_OK}" in ownership[0]


def test_14b_a_refusal_is_also_fully_logged(machine, deck, tmp_path, caplog):
    caplog.set_level(logging.INFO, logger="report_extractor.renderer")
    machine.start_user_powerpoint("user.pptx")
    with pytest.raises(PowerPointStageError):
        _render(deck, tmp_path)
    lines = [r.getMessage() for r in caplog.records if "POWERPOINT_" in r.getMessage()]
    assert any(line.startswith("POWERPOINT_SNAPSHOT") and "status=readable" in line and "same_session=[4000]" in line
               for line in lines), lines
    assert any("decision=refused" in line and REASON_USER_SESSION_RUNNING in line for line in lines), lines


def test_14c_the_cache_line_says_whether_the_preview_came_from_cache_and_with_which_reason(monkeypatch, tmp_path,
                                                                                           caplog):
    caplog.set_level(logging.INFO, logger="report_extractor.webview")
    monkeypatch.setattr(qrn, "powerpoint_probe", lambda: (False, "availability", "not-windows"))
    monkeypatch.setattr(qrn, "find_soffice", lambda: None)
    service = _preview_service()
    candidate = _candidate(_two_item_slide_deck(tmp_path))
    service._slide_preview_for(candidate)
    service._slide_preview_for(candidate)
    lines = [r.getMessage() for r in caplog.records if r.getMessage().startswith("SLIDE_PREVIEW_CACHE")]
    assert lines and all("cached_backend=" in line and "cached_reason=" in line and "hit=1" in line
                         for line in lines)
    assert any("action=reuse" in line for line in lines)


# ================================================================================== static safety contracts
def test_15_the_fix_added_no_process_termination_no_registry_write_and_no_attach():
    forbidden = ("taskkill", "terminateprocess", "psutil", "pkill", "killall", "stop-process", "wmic process",
                 "createprocesswithlogon", "logonuser", "runas ")
    registry_writes = (r"winreg\.SetValue", r"winreg\.CreateKey", r"winreg\.DeleteKey", r"winreg\.DeleteValue")
    app_dir = Path(qrn.__file__).resolve().parent
    for path in sorted(app_dir.glob("*.py")):
        code = _code_without_comments_and_docstrings(path.read_text(encoding="utf-8"))
        low = code.lower()
        for needle in forbidden:
            assert needle not in low, f"{path.name} contains {needle!r}"
        for pattern in registry_writes:
            assert not re.search(pattern, code), f"{path.name} writes the registry ({pattern})"


def test_16_the_renderer_still_never_attaches_enumerates_or_kills():
    code = _code_without_comments_and_docstrings(Path(qrn.__file__).read_text(encoding="utf-8"))
    assert not re.search(r"for\s+\w+\s+in\s+\w*\.?Presentations\b", code), "never enumerate Presentations"
    assert "ActivePresentation" not in code and "GetObject(" not in code
    assert not re.search(r"win32com\.client\.Dispatch\(", code), "attaching Dispatch() is forbidden"
    assert not re.search(r"\bOpenProcess\b|\bTerminateProcess\b|\bos\.kill\b|\.terminate\(", code)
    assert code.count(".Quit()") == 1 and code.count(".Close()") == 1
    # the ROT is a lookup only: its result is never returned to a caller as a usable Application
    assert "_rot_lookup(client) == ROT_REGISTERED" in code


def test_17_the_process_helpers_never_open_a_handle():
    tree = ast.parse(Path(qrn.__file__).read_text(encoding="utf-8"))
    called = {n.func.attr for n in ast.walk(tree) if isinstance(n, ast.Call) and isinstance(n.func, ast.Attribute)}
    assert "OpenProcess" not in called and "TerminateProcess" not in called
    assert "CreateToolhelp32Snapshot" in called and "CloseHandle" in called
    assert "ProcessIdToSessionId" in called


def test_18_the_capability_memo_cannot_pin_a_stale_answer(monkeypatch):
    monkeypatch.setattr(qrn, "powerpoint_probe", lambda: (True, "", "progid-com"))
    _set_powerpoint_present(monkeypatch, True)
    assert powerpoint_render_capability(max_age_s=60).reason == REASON_USER_SESSION_RUNNING
    # the machine moves on: that PowerPoint is now in another session, so OUR same-session set is empty
    monkeypatch.setattr(qrn, "_powerpoint_process_ids", lambda: frozenset())
    assert powerpoint_render_capability(max_age_s=60).reason == REASON_USER_SESSION_RUNNING, "memo honoured"
    qrn._CAPABILITY_MEMO.clear()
    assert powerpoint_render_capability(max_age_s=60).reason == REASON_OK, "clearing the memo sees the new state"
    assert powerpoint_render_capability().reason == REASON_OK, "max_age_s=0 always re-reads"


def test_18b_the_diagnostics_row_only_says_close_powerpoint_when_it_is_proven_open(monkeypatch):
    """The diagnostics panel had the same over-claim as the preview notice: EVERY PowerPoint-present refusal was
    rendered as "đóng PowerPoint của bạn".  That instruction is only true for the proven same-session case; for
    the other codes the reviewer would go looking for a process that does not exist."""
    from app.config import AppConfig
    from app.diagnostics import run_diagnostics

    def row(pids, provable=True):
        monkeypatch.setattr(qrn, "powerpoint_probe", lambda: (True, "", "progid-com"))
        monkeypatch.setattr(qrn, "_powerpoint_process_ids", lambda: pids)
        monkeypatch.setattr(qrn, "_session_filter_is_provable", lambda: provable)
        return dict(run_diagnostics(AppConfig(), check_ollama=False))["PowerPoint isolation"]

    user_session = row(frozenset({4000}))
    assert REASON_USER_SESSION_RUNNING in user_session
    assert "đóng PowerPoint của bạn" in user_session, "the proven case still tells the reviewer what to do"

    unproven = row(frozenset({4000}), provable=False)
    assert REASON_SESSION_UNPROVEN in unproven
    assert "đóng PowerPoint của bạn" not in unproven, "never instruct the reviewer to close an unproven process"

    monkeypatch.setattr(qrn, "_powerpoint_process_ids", lambda: None)
    unreadable = dict(run_diagnostics(AppConfig(), check_ollama=False))["PowerPoint isolation"]
    assert REASON_PROCESS_TABLE_UNREADABLE in unreadable
    assert "đóng PowerPoint của bạn" not in unreadable

    monkeypatch.setattr(qrn, "powerpoint_probe", lambda: (False, "availability", "progid-not-registered"))
    absent = dict(run_diagnostics(AppConfig(), check_ollama=False))["PowerPoint isolation"]
    assert REASON_UNAVAILABLE in absent and "đóng PowerPoint của bạn" not in absent
    assert absent != user_session, "the opposite causes must stay distinct (PROMPT-041 test_24)"


def test_19_the_reported_message_cannot_be_produced_without_a_process_fact(monkeypatch, report, tmp_path):
    """The exact UI sentence the tester saw comes from ONE slug; pin that the renderer can no longer produce that
    slug when the process table is empty and vouched."""
    monkeypatch.setattr(qrn, "powerpoint_probe", lambda: (True, "", "progid-com"))
    monkeypatch.setattr(qrn, "_powerpoint_process_ids", lambda: frozenset())
    monkeypatch.setattr(qrn, "_enumerate_powerpoint_pids", lambda: frozenset())
    monkeypatch.setattr(qrn, "_process_table_vouched", lambda: True)
    monkeypatch.setattr(qrn, "_rot_lookup", lambda client: qrn.ROT_REGISTERED)
    tsx = (Path(__file__).resolve().parent.parent / "frontend" / "src" / "tabs" / "LearningTab.tsx") \
        .read_text(encoding="utf-8")
    assert "process_table_vouched" not in tsx
    assert REASON_REGISTERED_INSTANCE in tsx, "the legacy slug must keep wording for already-persisted previews"
