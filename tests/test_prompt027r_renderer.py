"""PROMPT-027R renderer: PowerPoint detection, COM instance safety, staged diagnostics and cache identity.

No test requires Microsoft PowerPoint.  The COM layer (pythoncom / win32com.client) and the Windows registry are
replaced by small, explicit fakes, so every branch of the backend chain is exercised headlessly.  Real PowerPoint
fidelity remains a Windows acceptance step (PROMPT-027R §34) and is never claimed by these tests.
"""
from __future__ import annotations

import logging
import sys
from pathlib import Path
from types import ModuleType, SimpleNamespace

import pytest
from PIL import Image

import app.application_service as svc
import app.qpn_renderer as qrn
from app.pptx_parser import parse_pptx
from app.qpn_renderer import (FAITHFUL_BACKENDS, PowerPointStageError, SlideRenderer, is_faithful_backend,
                              powerpoint_probe, render_with_powerpoint)
from tests.test_after_evidence import _deck, _pics
from tests.test_prompt027_rendering import _candidate, _preview_service, _simple_slide
from tests.test_prompt025_multi_item import NAME, _two_item_slide


def _report(tmp_path: Path, name: str = "render.pptx"):
    deck = _simple_slide(tmp_path / name, lambda s: _pics(s, ["#c8e6c9"], 4, 3, w=2, h=1.4))
    return parse_pptx(str(deck))


def _messages(caplog, prefix: str):
    return [r.getMessage() for r in caplog.records if r.getMessage().startswith(prefix)]


# ================================================================== backend selection, faithfulness and logs
def test_powerpoint_is_selected_and_reported_faithful_when_available(tmp_path, monkeypatch, caplog):
    report = _report(tmp_path)

    def export(pptx, slide_numbers, out_dir, width_px=1920):
        out = {}
        for n in slide_numbers:
            target = Path(out_dir) / f"slide_{n:03d}.png"
            Image.new("RGB", (width_px, int(report.slide_height * width_px / report.slide_width)), "white").save(target)
            out[n] = target
        return out

    monkeypatch.setattr(qrn, "powerpoint_available", lambda: True)
    monkeypatch.setattr(qrn, "render_with_powerpoint", export)
    renderer = SlideRenderer(width_px=1600)
    with caplog.at_level(logging.INFO, logger="report_extractor.renderer"):
        renderer.render(report, [1], tmp_path / "o", purpose="learning_preview")
    assert renderer.last_backend == "powerpoint" and is_faithful_backend(renderer.last_backend)
    assert renderer.degraded is False
    line = _messages(caplog, "SLIDE_RENDER ")[0]
    assert "purpose=learning_preview backend=powerpoint slide=1" in line
    assert "attempted_backend=powerpoint" in line and "selected_backend=powerpoint" in line
    assert "faithful=true" in line


def test_unavailable_powerpoint_is_reported_by_stage_and_builtin_is_selected(tmp_path, monkeypatch, caplog):
    report = _report(tmp_path, "unavail.pptx")
    monkeypatch.setattr(qrn, "powerpoint_available", lambda: False)
    monkeypatch.setattr(qrn, "find_soffice", lambda: None)
    renderer = SlideRenderer(width_px=1200)
    with caplog.at_level(logging.WARNING, logger="report_extractor.renderer"):
        renderer.render(report, [1], tmp_path / "o", purpose="after_evidence")
    assert renderer.last_backend == "builtin" and renderer.degraded is False     # absent != degraded
    failure = _messages(caplog, "SLIDE_RENDER_BACKEND_FAILED")[0]
    assert "purpose=after_evidence backend=powerpoint stage=availability" in failure
    assert "reason=unavailable" in failure


def test_powerpoint_failure_is_logged_with_its_stage_then_falls_back(tmp_path, monkeypatch, caplog):
    report = _report(tmp_path, "stage.pptx")
    hres_error = type("ComError", (Exception,), {})("(-2147221005, 'Invalid class string')")
    hres_error.hresult = -2147221005

    def broken(pptx, slide_numbers, out_dir, width_px=1920):
        raise PowerPointStageError("dispatch", hres_error, pptx)

    monkeypatch.setattr(qrn, "powerpoint_available", lambda: True)
    monkeypatch.setattr(qrn, "render_with_powerpoint", broken)
    renderer = SlideRenderer(prefer=("powerpoint", "builtin"), width_px=1200)
    with caplog.at_level(logging.WARNING, logger="report_extractor.renderer"):
        out = renderer.render(report, [1], tmp_path / "o", purpose="learning_preview")
    assert out[1].exists() and renderer.last_backend == "builtin"
    assert renderer.degraded is True, "a faithful backend that FAILED marks the result degraded"
    failure = _messages(caplog, "SLIDE_RENDER_BACKEND_FAILED")[0]
    assert "purpose=learning_preview backend=powerpoint stage=dispatch" in failure
    assert "exception_type=ComError" in failure and "hresult=0x800401F3" in failure
    assert "reason=" in failure and "Invalid class string" in failure


def test_com_failure_reason_never_contains_a_filesystem_path(tmp_path, monkeypatch):
    report = _report(tmp_path, "secret.pptx")
    leaked = f"cannot open {report.path}"
    error = PowerPointStageError("presentation-open", RuntimeError(leaked), report.path)
    assert str(report.path) not in error.reason and "secret.pptx" not in error.reason
    assert "<report>" in error.reason and error.stage == "presentation-open"


def test_only_powerpoint_is_faithful_and_the_chain_is_unchanged():
    assert FAITHFUL_BACKENDS == frozenset({"powerpoint"})
    assert SlideRenderer().prefer == ["powerpoint", "libreoffice", "builtin"]


def test_availability_signature_reports_each_backend_in_the_chain(monkeypatch):
    monkeypatch.setattr(qrn, "powerpoint_available", lambda: True)
    monkeypatch.setattr(qrn, "find_soffice", lambda: None)
    assert SlideRenderer().availability() == {"powerpoint": True, "libreoffice": False, "builtin": True}


# ================================================================== probe (§15): stages and registry views
def _force_win32(monkeypatch):
    monkeypatch.setattr(qrn, "sys", SimpleNamespace(platform="win32", version_info=sys.version_info))


def _force_non_win32(monkeypatch, host: str = "linux"):
    """Simulate a non-Windows host explicitly.

    PROMPT-029: this suite now also runs on REAL Windows during build validation, so a test of the
    off-Windows branch must not depend on the machine it happens to run on.  ``powerpoint_probe()`` reads
    ``qrn.sys.platform``, so replacing that namespace is enough — production behaviour is untouched.
    """
    assert host != "win32"
    monkeypatch.setattr(qrn, "sys", SimpleNamespace(platform=host, version_info=sys.version_info))


def test_probe_is_negative_off_windows_on_every_host(monkeypatch):
    """The off-Windows branch, exercised deterministically no matter where the suite runs."""
    for host in ("linux", "darwin", "freebsd"):
        _force_non_win32(monkeypatch, host)
        assert powerpoint_probe() == (False, "availability", "not-windows"), host
        assert qrn.powerpoint_available() is False and qrn.powerpoint_unavailable_detail() == "not-windows"


def test_probe_contract_holds_on_the_real_host_whatever_it_is():
    """Host-independent: the probe always returns the documented 3-tuple with a safe, path-free detail.

    On real Windows a machine without PowerPoint (or with a partial pywin32 whose ``pythoncom`` has no
    ``CLSIDFromProgID``) legitimately answers ``progid-not-registered (<ExceptionType>)`` — that is a valid
    availability answer, not a bug, so this test accepts it instead of hardcoding one host's outcome.
    """
    available, stage, detail = powerpoint_probe()
    assert isinstance(available, bool) and isinstance(stage, str) and isinstance(detail, str)
    assert stage in ("", "availability", "import"), stage
    assert available is (stage == ""), "an available backend has nothing to explain"
    if sys.platform == "win32":
        assert detail != "not-windows", "production must NOT claim not-windows on Windows"
        assert stage != "availability" or detail.startswith("progid-"), detail
    else:
        assert (available, stage, detail) == (False, "availability", "not-windows")
    # the detail is a diagnostic shown in the UI/logs: no drive letters, no user paths
    for needle in (":\\", ":/", "\\Users\\", "/home/"):
        assert needle not in detail, f"{needle!r} leaked into {detail!r}"


def test_a_partial_pywin32_is_reported_as_unavailable_with_the_reason_named(monkeypatch):
    """Regression for the value seen on the real Windows build machine:
    ``pythoncom`` imports but has no ``CLSIDFromProgID`` → ``progid-not-registered (AttributeError)``.

    A broken/partial pywin32 must never be mistaken for "PowerPoint is installed", and must never crash
    the probe or launch PowerPoint.
    """
    _force_win32(monkeypatch)
    pythoncom = ModuleType("pythoncom")                       # deliberately WITHOUT CLSIDFromProgID
    win32com = ModuleType("win32com")
    client = ModuleType("win32com.client")
    client.DispatchEx = lambda progid: pytest.fail("the probe must not create a COM instance")
    win32com.client = client
    monkeypatch.setitem(sys.modules, "pythoncom", pythoncom)
    monkeypatch.setitem(sys.modules, "win32com", win32com)
    monkeypatch.setitem(sys.modules, "win32com.client", client)
    winreg = ModuleType("winreg")
    winreg.HKEY_CLASSES_ROOT = "HKCR"
    winreg.KEY_READ = winreg.KEY_WOW64_64KEY = winreg.KEY_WOW64_32KEY = 1
    winreg.OpenKey = lambda *a, **k: (_ for _ in ()).throw(OSError("missing"))
    monkeypatch.setitem(sys.modules, "winreg", winreg)
    assert powerpoint_probe() == (False, "availability", "progid-not-registered (AttributeError)")


def test_probe_never_launches_powerpoint_even_when_it_is_registered(monkeypatch):
    """The other half of "side-effect free": availability must not create a COM instance."""
    _force_win32(monkeypatch)
    launched = []

    def boom(*args, **kwargs):
        launched.append(args or kwargs)
        raise AssertionError("powerpoint_probe() must not touch a live COM object")

    pythoncom = ModuleType("pythoncom")
    pythoncom.CLSIDFromProgID = lambda progid: "{91493441-5A91-11CF-8700-00AA0060263B}"
    pythoncom.CoInitialize = boom
    pythoncom.CoUninitialize = boom
    win32com = ModuleType("win32com")
    client = ModuleType("win32com.client")
    client.DispatchEx, client.Dispatch, client.GetObject = boom, boom, boom
    win32com.client = client
    monkeypatch.setitem(sys.modules, "pythoncom", pythoncom)
    monkeypatch.setitem(sys.modules, "win32com", win32com)
    monkeypatch.setitem(sys.modules, "win32com.client", client)
    assert powerpoint_probe() == (True, "", "progid-com")
    assert qrn.powerpoint_available() is True
    assert launched == [], f"the probe started PowerPoint: {launched}"


def test_missing_pywin32_is_classified_as_an_import_stage(monkeypatch):
    _force_win32(monkeypatch)
    monkeypatch.setitem(sys.modules, "pythoncom", None)           # import raises ImportError
    available, stage, detail = powerpoint_probe()
    assert (available, stage) == (False, "import") and detail.startswith("pythoncom")


def _fake_pywin32(monkeypatch, progid_ok: bool):
    pythoncom = ModuleType("pythoncom")

    def clsid(progid):
        if progid_ok:
            return "{clsid}"
        raise OSError("progid not registered")

    pythoncom.CLSIDFromProgID = clsid
    pythoncom.CoInitialize = lambda: None
    pythoncom.CoUninitialize = lambda: None
    win32com = ModuleType("win32com")
    client = ModuleType("win32com.client")
    client.DispatchEx = lambda progid: None
    client.Dispatch = lambda progid: None
    win32com.client = client
    monkeypatch.setitem(sys.modules, "pythoncom", pythoncom)
    monkeypatch.setitem(sys.modules, "win32com", win32com)
    monkeypatch.setitem(sys.modules, "win32com.client", client)


def test_com_progid_lookup_is_the_first_answer(monkeypatch):
    _force_win32(monkeypatch)
    _fake_pywin32(monkeypatch, progid_ok=True)
    assert powerpoint_probe() == (True, "", "progid-com")


def test_registry_fallback_tries_default_64bit_and_32bit_views(monkeypatch):
    _force_win32(monkeypatch)
    _fake_pywin32(monkeypatch, progid_ok=False)
    tried = []
    winreg = ModuleType("winreg")
    winreg.HKEY_CLASSES_ROOT = "HKCR"
    winreg.KEY_READ = 0x20019
    winreg.KEY_WOW64_64KEY = 0x100
    winreg.KEY_WOW64_32KEY = 0x200

    class Handle:
        def __enter__(self):
            return self

        def __exit__(self, *exc):
            return False

    def open_key(root, path, reserved, access):
        tried.append(access & (winreg.KEY_WOW64_64KEY | winreg.KEY_WOW64_32KEY))
        if access & winreg.KEY_WOW64_32KEY:                       # only the 32-bit view has the ProgID
            return Handle()
        raise OSError("not in this view")

    winreg.OpenKey = open_key
    monkeypatch.setitem(sys.modules, "winreg", winreg)
    assert powerpoint_probe() == (True, "", "progid-registry")
    assert tried[0] == 0 and winreg.KEY_WOW64_64KEY in tried and winreg.KEY_WOW64_32KEY in tried


def test_not_registered_is_an_availability_stage_with_a_safe_detail(monkeypatch):
    _force_win32(monkeypatch)
    _fake_pywin32(monkeypatch, progid_ok=False)
    winreg = ModuleType("winreg")
    winreg.HKEY_CLASSES_ROOT = "HKCR"
    winreg.KEY_READ = 1
    winreg.KEY_WOW64_64KEY = 2
    winreg.KEY_WOW64_32KEY = 4

    def never(*_a, **_k):
        raise OSError("missing")

    winreg.OpenKey = never
    monkeypatch.setitem(sys.modules, "winreg", winreg)
    available, stage, detail = powerpoint_probe()
    assert (available, stage) == (False, "availability") and "progid-not-registered" in detail


def test_availability_is_never_cached_between_probes(monkeypatch):
    _force_win32(monkeypatch)
    _fake_pywin32(monkeypatch, progid_ok=False)
    winreg = ModuleType("winreg")
    winreg.HKEY_CLASSES_ROOT = "HKCR"
    winreg.KEY_READ = 1
    winreg.KEY_WOW64_64KEY = 2
    winreg.KEY_WOW64_32KEY = 4
    def missing(*_a, **_k):
        raise OSError("missing")

    winreg.OpenKey = missing
    monkeypatch.setitem(sys.modules, "winreg", winreg)
    assert qrn.powerpoint_available() is False
    _fake_pywin32(monkeypatch, progid_ok=True)                      # PowerPoint installed in the meantime
    assert qrn.powerpoint_available() is True


# ================================================================== COM instance safety (§16) with fake COM
class _Calls:
    """Ordered record of every COM call the fake PowerPoint receives."""

    def __init__(self):
        self.events = []


def _fake_com(monkeypatch, calls: _Calls, fail: str = "", com_init_fails: bool = False,
              close_fails: bool = False):
    """Install fake pythoncom + win32com.client.  ``fail`` names the stage that raises."""

    class Presentation:
        PageSetup = SimpleNamespace(SlideWidth=960, SlideHeight=540)

        def __init__(self, path):
            self.path = path

        def Slides(self, n):
            calls.events.append(("Slides", n))
            if fail == "slide-export":
                raise RuntimeError("export refused")

            class Slide:
                def Export(self_inner, target, fmt, w, h):
                    calls.events.append(("Export", fmt))
                    Image.new("RGB", (int(w), int(h)), "white").save(target)
            return Slide()

        def Close(self):
            calls.events.append(("Presentation.Close", None))
            if close_fails:
                raise RuntimeError("close failed")

    class PresentationCollection:
        def Open(self, path, read_only, untitled, with_window):
            calls.events.append(("Open", (read_only, untitled, with_window)))
            if fail == "presentation-open":
                raise RuntimeError("cannot open presentation")
            return Presentation(path)

    class App:
        Presentations = PresentationCollection()

        def Quit(self):
            calls.events.append(("App.Quit", None))

    def dispatch_ex(progid):
        calls.events.append(("DispatchEx", progid))
        if fail == "dispatch":
            error = RuntimeError("Class not registered")
            error.hresult = -2147221005
            raise error
        return App()

    def dispatch(progid):
        calls.events.append(("Dispatch", progid))
        raise AssertionError("the application must never attach to a user's instance with Dispatch")

    pythoncom = ModuleType("pythoncom")

    def co_init():
        calls.events.append(("CoInitialize", None))
        if com_init_fails:
            raise RuntimeError("RPC_E_CHANGED_MODE")

    pythoncom.CoInitialize = co_init
    pythoncom.CoUninitialize = lambda: calls.events.append(("CoUninitialize", None))
    client = ModuleType("win32com.client")
    client.DispatchEx = dispatch_ex
    client.Dispatch = dispatch
    win32com = ModuleType("win32com")
    win32com.client = client
    monkeypatch.setitem(sys.modules, "pythoncom", pythoncom)
    monkeypatch.setitem(sys.modules, "win32com", win32com)
    monkeypatch.setitem(sys.modules, "win32com.client", client)


def _names(calls: _Calls):
    return [name for name, _ in calls.events]


def test_powerpoint_export_uses_an_app_owned_instance_and_balances_com(tmp_path, monkeypatch):
    report = _report(tmp_path, "com_ok.pptx")
    calls = _Calls()
    _fake_com(monkeypatch, calls)
    out = render_with_powerpoint(report.path, [1], tmp_path, width_px=960)
    assert out[1].exists()
    names = _names(calls)
    assert "DispatchEx" in names and "Dispatch" not in names
    assert names.count("CoInitialize") == 1 and names.count("CoUninitialize") == 1
    # only the presentation WE opened is closed, and only the instance WE created is quit (§16)
    assert names.index("Presentation.Close") < names.index("App.Quit") < names.index("CoUninitialize")
    open_args = [args for name, args in calls.events if name == "Open"][0]
    assert open_args == (True, False, False), "ReadOnly, no untitled, no window"


def test_presentation_open_failure_still_quits_only_our_instance(tmp_path, monkeypatch):
    report = _report(tmp_path, "com_open.pptx")
    calls = _Calls()
    _fake_com(monkeypatch, calls, fail="presentation-open")
    with pytest.raises(PowerPointStageError) as info:
        render_with_powerpoint(report.path, [1], tmp_path, width_px=960)
    assert info.value.stage == "presentation-open"
    names = _names(calls)
    assert "Presentation.Close" not in names                       # nothing was opened, nothing is closed
    assert "App.Quit" in names and names.count("CoUninitialize") == 1


def test_dispatch_failure_names_the_stage_and_hresult(tmp_path, monkeypatch):
    report = _report(tmp_path, "com_dispatch.pptx")
    calls = _Calls()
    _fake_com(monkeypatch, calls, fail="dispatch")
    with pytest.raises(PowerPointStageError) as info:
        render_with_powerpoint(report.path, [1], tmp_path, width_px=960)
    assert info.value.stage == "dispatch" and info.value.hresult == "0x800401F3"
    assert info.value.exception_type == "RuntimeError"
    assert "App.Quit" not in _names(calls), "no instance was created, so none is quit"
    assert _names(calls).count("CoUninitialize") == 1


def test_com_initialisation_failure_never_uninitialises_what_it_did_not_start(tmp_path, monkeypatch):
    report = _report(tmp_path, "com_init.pptx")
    calls = _Calls()
    _fake_com(monkeypatch, calls, com_init_fails=True)
    with pytest.raises(PowerPointStageError) as info:
        render_with_powerpoint(report.path, [1], tmp_path, width_px=960)
    assert info.value.stage == "com-init"
    assert "CoUninitialize" not in _names(calls)


def test_cleanup_failure_is_logged_and_the_exported_slide_is_kept(tmp_path, monkeypatch, caplog):
    report = _report(tmp_path, "com_cleanup.pptx")
    calls = _Calls()
    _fake_com(monkeypatch, calls, close_fails=True)
    with caplog.at_level(logging.INFO, logger="report_extractor.renderer"):
        out = render_with_powerpoint(report.path, [1], tmp_path, width_px=960)
    assert out[1].exists()
    cleanup = _messages(caplog, "SLIDE_RENDER_CLEANUP")
    assert cleanup and "stage=cleanup" in cleanup[0] and "step=presentation-close" in cleanup[0]


def test_export_failure_is_tagged_slide_export(tmp_path, monkeypatch):
    report = _report(tmp_path, "com_export.pptx")
    calls = _Calls()
    _fake_com(monkeypatch, calls, fail="slide-export")
    with pytest.raises(PowerPointStageError) as info:
        render_with_powerpoint(report.path, [1], tmp_path, width_px=960)
    assert info.value.stage == "slide-export"
    assert "App.Quit" in _names(calls) and "Presentation.Close" in _names(calls)


# ================================================================== §30 cache identity and recovery
def test_preview_identity_includes_backend_availability(tmp_path, monkeypatch):
    deck = _two_item_slide_deck(tmp_path)
    service = _preview_service()
    candidate = _candidate(deck)
    monkeypatch.setattr(qrn, "powerpoint_available", lambda: False)
    monkeypatch.setattr(qrn, "find_soffice", lambda: None)
    fallback = service._slide_preview_for(candidate)
    assert fallback["backend"] == "builtin" and fallback["faithful"] is False

    def export(pptx, slide_numbers, out_dir, width_px=1920):
        out = {}
        for n in slide_numbers:
            target = Path(out_dir) / f"slide_{n:03d}.png"
            Image.new("RGB", (int(width_px), int(width_px * 540 / 960)), "white").save(target)
            out[n] = target
        return out

    monkeypatch.setattr(qrn, "powerpoint_available", lambda: True)
    monkeypatch.setattr(qrn, "render_with_powerpoint", export)
    recovered = service._slide_preview_for(candidate)
    assert len(service._slide_preview_cache) == 2, "a newly available backend is a new identity, not a stale hit"
    assert recovered["backend"] == "powerpoint" and recovered["faithful"] is True


def test_degraded_fallback_is_retried_after_the_ttl_and_then_recovers(tmp_path, monkeypatch):
    deck = _two_item_slide_deck(tmp_path)
    service = _preview_service()
    candidate = _candidate(deck)
    clock = {"now": 1000.0}
    monkeypatch.setattr(svc, "time", SimpleNamespace(monotonic=lambda: clock["now"]))
    state = {"renders": 0, "broken": True}

    def flaky(pptx, slide_numbers, out_dir, width_px=1920):
        state["renders"] += 1
        if state["broken"]:
            raise PowerPointStageError("slide-export", RuntimeError("busy"), pptx)
        out = {}
        for n in slide_numbers:
            target = Path(out_dir) / f"slide_{n:03d}.png"
            Image.new("RGB", (int(width_px), int(width_px * 540 / 960)), "white").save(target)
            out[n] = target
        return out

    monkeypatch.setattr(qrn, "powerpoint_available", lambda: True)
    monkeypatch.setattr(qrn, "render_with_powerpoint", flaky)
    first = service._slide_preview_for(candidate)
    assert first["backend"] == "builtin" and first["degraded"] is True and state["renders"] == 1
    service._slide_preview_for(candidate)
    assert state["renders"] == 1, "within the TTL a degraded preview is reused (no re-render per candidate)"
    clock["now"] += svc.DEGRADED_PREVIEW_TTL_S + 1
    state["broken"] = False                                   # COM recovered
    recovered = service._slide_preview_for(candidate)
    assert state["renders"] == 2
    assert recovered["backend"] == "powerpoint" and recovered["degraded"] is False


def test_faithful_preview_is_reused_for_the_whole_session(tmp_path, monkeypatch):
    deck = _two_item_slide_deck(tmp_path)
    service = _preview_service()
    candidate = _candidate(deck)
    clock = {"now": 5.0}
    monkeypatch.setattr(svc, "time", SimpleNamespace(monotonic=lambda: clock["now"]))
    renders = {"n": 0}

    def export(pptx, slide_numbers, out_dir, width_px=1920):
        renders["n"] += 1
        out = {}
        for n in slide_numbers:
            target = Path(out_dir) / f"slide_{n:03d}.png"
            Image.new("RGB", (int(width_px), int(width_px * 540 / 960)), "white").save(target)
            out[n] = target
        return out

    monkeypatch.setattr(qrn, "powerpoint_available", lambda: True)
    monkeypatch.setattr(qrn, "render_with_powerpoint", export)
    service._slide_preview_for(candidate)
    clock["now"] += 10 * svc.DEGRADED_PREVIEW_TTL_S
    service._slide_preview_for(candidate)
    assert renders["n"] == 1


def _two_item_slide_deck(tmp_path: Path) -> Path:
    return _deck(tmp_path / "in" / NAME, [_two_item_slide()])


def test_renderer_source_keeps_the_structural_checks_of_prompt027():
    source = Path(qrn.__file__).read_text(encoding="utf-8")
    assert "b.inset_top" in source and "b.vertical_anchor" in source
    assert "DispatchEx" in source and "Dispatch(" not in source.replace("DispatchEx(", "")
    for stage in ("import", "com-init", "dispatch", "presentation-open", "slide-export"):
        assert f'PowerPointStageError("{stage}"' in source, stage
