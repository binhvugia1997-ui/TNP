"""PROMPT-030 — PowerPoint preview renderer regression (probe vs. render, COM, fallbacks, diagnostics).

Covers §2–3: staged diagnostics, side-effect-free probe, COM/thread handling, fallback reasons.
All tests are deterministic on Linux via monkeypatching — no real PowerPoint required.
Real PowerPoint validation is Windows acceptance and is NOT claimed here.
"""
from __future__ import annotations

import sys
from pathlib import Path
from types import ModuleType, SimpleNamespace

import pytest

from app.qpn_renderer import (
    PowerPointProbeResult,
    PowerPointRenderDiagnostics,
    SlideRenderer,
    last_powerpoint_diagnostics,
    powerpoint_available,
    powerpoint_probe,
    renderer_status,
    renderer_status_detailed,
    render_with_powerpoint,
)
from tests.test_application_service import _learning_candidates  # helpers: real PPT parse not needed, use tiny deck
from tests.test_prompt025_multi_item import _deck
import app.qpn_renderer as qr


# ================================================================= helpers
def _fake_report(tmp_path: Path, slides=1):
    """Minimal valid PPTX deck for SlideRenderer.render (needs at least one slide)."""
    from pptx import Presentation
    prs = Presentation()
    prs.slide_width = 12192000
    prs.slide_height = 6858000
    for _ in range(slides):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        from pptx.util import Inches
        title = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(4), Inches(0.8))
        title.text_frame.text = "Test slide"
    path = tmp_path / "test.pptx"
    prs.save(str(path))
    from app.pptx_parser import parse_pptx
    return parse_pptx(path), path


# ================================================================= § probe (side-effect-free)
# PROMPT-030C: probe side-effect-free trên MỌI nền tảng — không bao giờ đụng COM entry point có
# thể khởi động/attach PowerPoint. Availability trên Windows là chuyện CỦA MÁY (PowerPoint có thể
# đã cài) nên test không bao giờ ép ``available is False`` trên Windows; ma trận registry/progid
# giả lập bên dưới bao phủ cả trạng thái có và không có PowerPoint trên mọi host.

_PROBE_LAUNCH_METHODS = (
    "Dispatch", "DispatchEx", "GetActiveObject", "GetObject", "EnsureDispatch",
    "CoInitialize", "CoInitializeEx", "CoUninitialize", "Open", "Quit", "Visible",
)


def _recorded_spy(log, name):
    """Callable ghi log khi bị gọi rồi nỗng — kể cả khi caller nuốt exception, log vẫn fail test."""
    def _spy(*args, **kwargs):
        log.append(name)
        raise AssertionError(f"probe phải không được gọi {name}")
    return _spy


class _SpyModule(ModuleType):
    """Module giả cho pythoncom / win32com(.client): mọi tên trong nhóm khởi động PowerPoint được
    ghi log khi TRUY CẬP và nỗng AssertionError khi GỌI. Dunder bị AttributeError như module thường."""

    def __init__(self, name, log):
        super().__init__(name)
        self._launch_log = log

    def __getattr__(self, item):
        if item.startswith("__"):
            raise AttributeError(item)
        if item in _PROBE_LAUNCH_METHODS:
            log = object.__getattribute__(self, "_launch_log")
            log.append(item)
            return _recorded_spy(log, item)
        raise AttributeError(item)


class com_error(Exception):  # noqa: N801 – cố ý trùng tên pythoncom.com_error (type name đi vào detail)
    """pythoncom.com_error giả (mô phỏng REGDB_E_CLASSNOTREG khi thiếu ProgID)."""

    def __init__(self, hresult=-2147221167, strerror="REGDB_E_CLASSNOTREG"):
        super().__init__(hresult, strerror)
        self.hresult = hresult


class _FakePythoncom(_SpyModule):
    def __init__(self, log, clsid_ok: bool):
        super().__init__("pythoncom", log)
        self.com_error = com_error
        if clsid_ok:
            self.CLSIDFromProgID = lambda progid: "{000209FF-0000-0000-C000-000000000046}"
        else:
            def _fail(progid):
                raise com_error()
            self.CLSIDFromProgID = _fail


class _FakeWinregKey:
    def __enter__(self):
        return self

    def __exit__(self, *exc):
        return False


class _FakeWinreg(ModuleType):
    HKEY_CLASSES_ROOT = 0
    KEY_READ = 0x20019
    KEY_WOW64_64KEY = 0x0100
    KEY_WOW64_32KEY = 0x0200

    def __init__(self, key_present: bool):
        super().__init__("winreg")
        self._key_present = key_present

    def OpenKey(self, *args, **kwargs):
        if not self._key_present:
            raise FileNotFoundError("no PowerPoint ProgID")
        return _FakeWinregKey()


def _probe_with_fakes(monkeypatch, *, platform="win32", clsid_ok=False, registry_key=True,
                      pythoncom_present=True, win32com_present=True):
    """Cài đủ pythoncom/win32com/winreg giả cho MỘT trạng thái máy; trả về (probe, launch_log).

    ``registry_key``: True = ProgID có trong HKCR, False = không có. Không có đường nào chạm COM thật.
    """
    log = []
    monkeypatch.setattr(qr, "sys", SimpleNamespace(platform=platform, version_info=sys.version_info))
    if pythoncom_present:
        monkeypatch.setitem(sys.modules, "pythoncom", _FakePythoncom(log, clsid_ok))
    else:
        monkeypatch.setitem(sys.modules, "pythoncom", None)   # ImportError khi import
    if win32com_present:
        fake_client = _SpyModule("win32com.client", log)
        fake_win32 = _SpyModule("win32com", log)
        fake_win32.client = fake_client
        monkeypatch.setitem(sys.modules, "win32com", fake_win32)
        monkeypatch.setitem(sys.modules, "win32com.client", fake_client)
    else:
        monkeypatch.setitem(sys.modules, "win32com", None)
        monkeypatch.setitem(sys.modules, "win32com.client", None)
    monkeypatch.setitem(sys.modules, "winreg", _FakeWinreg(registry_key))
    return powerpoint_probe(), log


def _install_real_launch_spies(monkeypatch, log):
    """Gắn spy lên COM entry point THẬT nếu pywin32 thật có mặt (máy Windows) để chứng minh probe
    thật không sờ tới chúng. Không có module thật (Linux) thì bỏ qua — probe cũng không import."""
    for name in ("win32com.client", "win32com", "pythoncom"):
        try:
            mod = __import__(name, fromlist=["*"])
        except Exception:  # noqa: BLE001 – host không có pywin32: không có gì để spy
            continue
        for attr in _PROBE_LAUNCH_METHODS:
            if hasattr(mod, attr):
                monkeypatch.setattr(mod, attr, _recorded_spy(log, attr))


def test_probe_side_effect_free_does_not_launch_powerpoint(monkeypatch):
    """PROMPT-030C: probe side-effect-free trên CẢ HAI nền tảng — không Dispatch/DispatchEx/Open/COM init.

    Trước đây test giả định probe luôn unavailable (môi trường Linux) nên fail trên Windows thật có
    PowerPoint — ở đó ``powerpoint_probe()`` trả ``(True, '', 'progid-registry')`` đúng hợp đồng
    (ProgID thấy qua HKCR). Test hiện tại: (1) máy thật — đúng hợp đồng host, KHÔNG ép availability
    trên Windows; (2) ma trận giả lập đủ trạng thái (COM ProgID, registry-only ProgID, không cài
    PowerPoint, thiếu pywin32, không-Windows); (3) mọi trạng thái đều chứng minh probe không gọi
    bất kỳ hàm khởi động PowerPoint nào — kể cả khi available=True."""
    # ---- 1) máy thật: hợp đồng của host, không suy đoán availability trên Windows
    real_log = []
    _install_real_launch_spies(monkeypatch, real_log)
    real = powerpoint_probe()
    assert real_log == [], f"probe thật đã gọi hàm khởi động PowerPoint: {real_log}"
    assert isinstance(real.available, bool) and real.available is tuple(real)[0]
    if sys.platform != "win32":
        assert real.available is False
        assert real.stage == "not_windows"
        assert tuple(real) == (False, "availability", "not-windows")
    else:
        # Windows: PowerPoint có/không là chuyện của máy — chỉ khẳng định hợp đồng, không ép False.
        assert real.stage in ("ok", "registry_missing", "win32com_missing")
        assert real.available is (real.stage == "ok")
        if real.available:
            assert real.details in ("progid-com", "progid-registry")
            assert tuple(real) == (True, "", real.details)
            assert real.reason == f"PowerPoint có sẵn ({real.details})"
        else:
            assert tuple(real) != (False, "availability", "not-windows"), "Windows phải không bao giờ trả not-windows"

    # ---- 2) ma trận giả lập registry/progid — deterministic trên mọi host
    # (a) Windows + ProgID thấy qua COM (progid-com)
    probe, log = _probe_with_fakes(monkeypatch, clsid_ok=True, registry_key=True)
    assert probe.available is True and probe.stage == "ok"
    assert tuple(probe) == (True, "", "progid-com")
    assert probe.reason == "PowerPoint có sẵn (progid-com)" and probe.details == "progid-com"
    assert log == [], f"probe đã gọi {log} khi báo available"

    # (b) Windows + ProgID chỉ thấy qua registry (progid-registry — đúng dạng máy Windows đã tái lập)
    probe, log = _probe_with_fakes(monkeypatch, clsid_ok=False, registry_key=True)
    assert probe.available is True and probe.stage == "ok"
    assert tuple(probe) == (True, "", "progid-registry")
    assert probe.reason == "PowerPoint có sẵn (progid-registry)" and probe.details == "progid-registry"
    assert log == [], f"probe đã gọi {log} khi báo available"

    # (c) Windows + không cài PowerPoint (CLSID fail, HKCR trống)
    probe, log = _probe_with_fakes(monkeypatch, clsid_ok=False, registry_key=False)
    assert probe.available is False and probe.stage == "registry_missing"
    assert tuple(probe) == (False, "availability", "progid-not-registered (com_error)")
    assert log == [], f"probe đã gọi {log}"

    # (d) Windows + thiếu pythoncom
    probe, log = _probe_with_fakes(monkeypatch, pythoncom_present=False)
    assert probe.available is False and probe.stage == "win32com_missing"
    assert tuple(probe)[1] == "import" and tuple(probe)[2].startswith("pythoncom ")
    assert log == [], f"probe đã gọi {log}"

    # (e) Windows + thiếu win32com.client
    probe, log = _probe_with_fakes(monkeypatch, clsid_ok=True, win32com_present=False)
    assert probe.available is False and probe.stage == "win32com_missing"
    assert tuple(probe)[1] == "import" and tuple(probe)[2].startswith("win32com.client ")
    assert log == [], f"probe đã gọi {log}"

    # (f) ép không-Windows — hợp đồng not-windows giữ nguyên trên mọi host
    for host in ("linux", "darwin"):
        probe, log = _probe_with_fakes(monkeypatch, platform=host)
        assert probe.available is False and probe.stage == "not_windows"
        assert tuple(probe) == (False, "availability", "not-windows"), host
        assert log == [], f"probe đã gọi {log} trên {host}"


def test_powerpoint_available_delegates_to_probe():
    probe = powerpoint_probe()
    assert powerpoint_available() is probe.available


def test_renderer_status_detailed_has_stages():
    d = renderer_status_detailed()
    assert "powerpoint" in d and "libreoffice" in d and "builtin" in d
    assert d["powerpoint"]["stage"] in ("not_windows", "ok", "registry_missing", "win32com_missing")


# ================================================================= § SlideRenderer fallback stages
def test_powerpoint_renderer_available_and_successful(tmp_path, monkeypatch):
    report, path = _fake_report(tmp_path, slides=1)
    # Monkeypatch: pretend PowerPoint is available and export succeeds
    monkeypatch.setattr(qr, "powerpoint_probe", lambda: PowerPointProbeResult(True, "ok", "PowerPoint có sẵn", ""))
    def fake_render(pptx, slides_, out_dir, width_px=1920, diagnostics=None):
        # Simulate PowerPoint export: write a tiny PNG for each slide
        from PIL import Image
        out = {}
        for n in slides_:
            img = Image.new("RGB", (100, 60), "white")
            target = Path(out_dir) / f"slide_{n:03d}.png"
            img.save(target, "PNG")
            out[n] = target
        if diagnostics is not None:
            diagnostics.availability = "ok"
            diagnostics.com_init = "ok"
            diagnostics.open_result = "ok"
            diagnostics.export_result = "ok"
            diagnostics.cleanup = "ok"
        return out
    monkeypatch.setattr(qr, "render_with_powerpoint", fake_render)

    r = SlideRenderer(prefer=("powerpoint", "builtin"), width_px=400)
    out_dir = tmp_path / "out1"
    res = r.render(report, [1], out_dir)
    assert 1 in res and Path(res[1]).is_file()
    assert r.last_backend == "powerpoint"
    assert r.last_fallback_reason == ""  # no fallback when PowerPoint succeeded
    # No yellow warning state: fallback reason empty
    assert not r.last_fallback_reason


def test_availability_negative_fallback_to_builtin_with_reason(tmp_path, monkeypatch):
    report, _ = _fake_report(tmp_path, slides=1)
    monkeypatch.setattr(qr, "powerpoint_probe",
                        lambda: PowerPointProbeResult(False, "registry_missing", "PowerPoint chưa được cài đặt (không có registry PowerPoint.Application)", ""))

    r = SlideRenderer(prefer=("powerpoint", "builtin"), width_px=400)
    out_dir = tmp_path / "out2"
    res = r.render(report, [1], out_dir)
    assert 1 in res
    assert r.last_backend == "builtin"
    assert r.last_fallback_stage == "probe"
    assert "registry" in r.last_fallback_reason.lower() or "PowerPoint" in r.last_fallback_reason
    assert r.last_diagnostics is not None and r.last_diagnostics.fallback_stage == "probe"


def test_availability_positive_but_export_fails_fallback_with_exact_stage(tmp_path, monkeypatch):
    report, _ = _fake_report(tmp_path, slides=1)
    monkeypatch.setattr(qr, "powerpoint_probe", lambda: PowerPointProbeResult(True, "ok", "PowerPoint có sẵn", ""))

    def failing_render(pptx, slides_, out_dir, width_px=1920, diagnostics=None):
        if diagnostics is not None:
            diagnostics.availability = "ok"
            diagnostics.com_init = "ok"
            diagnostics.open_result = "ok"
            diagnostics.export_result = "failed"
            diagnostics.export_detail = "Export failed: mock error"
            diagnostics.fallback_stage = "export"
            diagnostics.fallback_reason = "PowerPoint export thất bại: mock error"
            # mimic that render_with_powerpoint sets global diagnostics before raising
            qr._last_powerpoint_diagnostics = diagnostics  # type: ignore[attr-defined]
        raise RuntimeError("mock PowerPoint export failure")

    monkeypatch.setattr(qr, "render_with_powerpoint", failing_render)

    r = SlideRenderer(prefer=("powerpoint", "builtin"), width_px=400)
    out_dir = tmp_path / "out3"
    res = r.render(report, [1], out_dir)
    # Fallback still produced a file via builtin
    assert 1 in res and Path(res[1]).is_file()
    assert r.last_backend == "builtin"
    assert r.last_fallback_stage == "export"
    assert "export" in r.last_fallback_reason.lower() or "mock" in r.last_fallback_reason


def test_com_thread_init_already_initialized_is_diagnostic_not_fatal(tmp_path, monkeypatch):
    """PROMPT-030: RPC_E_CHANGED_MODE on CoInitialize is 'already_initialized' diagnostic.

    Lập luận theo PROMPT-032: COM per-thread state; PRD 4.4 cho phép COM error → fallback sau khi cleanup.
    RPC_E_CHANGED_MODE nghĩa là thread đã init COM với apartment mode khác → COM API vẫn dùng được
    (ĐÂY KHÔNG PHẢI LỖI KHỞI TẠO THIẾU) → chỉ ghi diagnostic com_init=already_initialized, KHÔNG fatal,
    KHÔNG close app của user, và cleanup không bao giờ CoUninitialize khi không init được.

    Harness adjusts to the merged contract: the renderer only ever creates app-owned instances via
    ``DispatchEx`` (PROMPT-035) and proves ownership first, so the fakes model an empty process table
    and a pristine instance — the export path is reachable and asserted."""
    report, _ = _fake_report(tmp_path, slides=1)
    close_called = []
    quit_called = []
    co_uninit = []
    pids = {"set": frozenset()}

    class FakePres:
        class PageSetupClass:
            SlideWidth = 12192000
            SlideHeight = 6858000
        PageSetup = PageSetupClass()
        def Close(self): close_called.append(True)
        def Slides(self, n):
            class S:
                def Export(self, path, *a, **kw):
                    import pathlib
                    pathlib.Path(path).write_bytes(b"")
            return S()

    class FakeApp:
        class PresentationsClass:
            Count = 0
            def Open(self, *a, **kw):
                return FakePres()
        Presentations = PresentationsClass()
        Windows = type("W", (), {"Count": 0})()
        def Quit(self):
            quit_called.append(True)
            pids["set"] = frozenset()

    class FakeComError(Exception):
        def __init__(self):
            super().__init__("CoInitialize has already been called.", -2147417833)
            self.hresult = -2147417833

    import types
    fake_pc = types.ModuleType("pythoncom")
    fake_pc.CoInitialize = lambda: (_ for _ in ()).throw(FakeComError())
    fake_pc.CoUninitialize = lambda: co_uninit.append(True)
    monkeypatch.setitem(sys.modules, "pythoncom", fake_pc)
    fake_win32 = types.ModuleType("win32com")
    fake_client = types.ModuleType("win32com.client")

    def _dispatch_ex(prog):
        pids["set"] = frozenset({4242})
        return FakeApp()

    fake_client.DispatchEx = _dispatch_ex
    fake_client.Dispatch = lambda prog: (_ for _ in ()).throw(
        AssertionError("Dispatch must never attach to a user instance"))
    fake_win32.client = fake_client
    monkeypatch.setitem(sys.modules, "win32com", fake_win32)
    monkeypatch.setitem(sys.modules, "win32com.client", fake_client)

    monkeypatch.setattr(qr, "powerpoint_probe", lambda: PowerPointProbeResult(True, "ok", "PowerPoint có sẵn", ""))
    # Ownership seams (same shape as PROMPT-035's own harness): empty table before, one new pid after.
    monkeypatch.setattr(qr, "_powerpoint_process_ids", lambda: pids["set"])
    monkeypatch.setattr(qr, "_running_powerpoint_registered", lambda client: False)
    monkeypatch.setattr(qr, "_PROCESS_EXIT_TIMEOUT_S", 0.1)

    diag = PowerPointRenderDiagnostics()
    out = tmp_path / "out_com"
    out.mkdir()
    try:
        render_with_powerpoint(report.path, [1], out, diagnostics=diag)
    except RuntimeError:
        # Chấp nhận nếu COM layer vẫn fail — QUAN TRỌNG: chỉ chẩn đoán, không nuốt lỗi
        assert diag.com_init == "already_initialized"
    assert diag.com_init == "already_initialized"
    # Already-initialized COM must NEVER be CoUninitialized by this module (no balanced pair to uninit)
    assert not co_uninit, "cleanup must not CoUninitialize COM this module did not initialise"


def test_cleanup_always_runs_even_on_export_failure(tmp_path, monkeypatch):
    """Export failure must still Close OUR presentation and Quit OUR app-owned instance (§16)."""
    report, _ = _fake_report(tmp_path, slides=1)
    close_called = []
    quit_called = []
    co_uninit = []
    pids = {"set": frozenset()}

    class TrackingPres:
        class PageSetupClass:
            SlideWidth = 12192000
            SlideHeight = 6858000
        PageSetup = PageSetupClass()
        def Close(self): close_called.append(True)
        def Slides(self, n):
            class S:
                def Export(self, *a, **kw):
                    raise RuntimeError("export boom")
            return S()

    class TrackingApp:
        class PresentationsClass:
            Count = 0
            def Open(self, *a, **kw):
                return TrackingPres()
        Presentations = PresentationsClass()
        Windows = type("W", (), {"Count": 0})()  # _is_pristine requires empty Presentations AND Windows
        def Quit(self):
            quit_called.append(True)
            pids["set"] = frozenset()

    monkeypatch.setattr(qr, "powerpoint_probe", lambda: PowerPointProbeResult(True, "ok", "PowerPoint có sẵn", ""))
    # Need Windows platform for render_with_powerpoint to proceed to COM stage
    monkeypatch.setattr(qr.sys, "platform", "win32", raising=False)
    monkeypatch.setattr(sys, "platform", "win32", raising=False)
    # Ownership seams (PROMPT-035): app-owned instance only — empty table before, one new pid after.
    monkeypatch.setattr(qr, "_powerpoint_process_ids", lambda: pids["set"])
    monkeypatch.setattr(qr, "_running_powerpoint_registered", lambda client: False)
    monkeypatch.setattr(qr, "_PROCESS_EXIT_TIMEOUT_S", 0.1)

    import types
    fake_pc = types.ModuleType("pythoncom")
    fake_pc.CoInitialize = lambda: 0
    fake_pc.CoUninitialize = lambda: co_uninit.append(True)
    monkeypatch.setitem(sys.modules, "pythoncom", fake_pc)
    fake_win32 = types.ModuleType("win32com")
    fake_client = types.ModuleType("win32com.client")

    def _dispatch_ex(prog):
        pids["set"] = frozenset({4242})  # this DispatchEx created exactly one new PowerPoint process
        return TrackingApp()

    fake_client.DispatchEx = _dispatch_ex
    fake_client.Dispatch = lambda prog: (_ for _ in ()).throw(
        AssertionError("Dispatch must never attach to a user instance"))
    fake_win32.client = fake_client
    monkeypatch.setitem(sys.modules, "win32com", fake_win32)
    monkeypatch.setitem(sys.modules, "win32com.client", fake_client)

    diag = PowerPointRenderDiagnostics()
    with pytest.raises(RuntimeError, match="Export failed|export boom|PowerPoint Export"):
        render_with_powerpoint(report.path, [1], tmp_path / "out_cleanup", diagnostics=diag)
    assert close_called  # presentation was closed even though export failed
    # Quit is only called when Count == 0 — our mock has Count == 0, so it should be called
    assert quit_called
    assert co_uninit  # cleanup always runs, including on export failure
    assert diag.fallback_stage == "export"


def test_non_windows_monkeypatched_deterministic_fallback(tmp_path, monkeypatch):
    # Non-Windows: powerpoint never attempted, builtin is used deterministically
    report, _ = _fake_report(tmp_path, slides=1)
    monkeypatch.setattr(qr.sys, "platform", "linux", raising=False)
    r = SlideRenderer(prefer=("powerpoint", "builtin"), width_px=400)
    res = r.render(report, [1], tmp_path / "out_linux")
    assert 1 in res
    assert r.last_backend == "builtin"
    assert r.last_fallback_stage == "probe"
