"""PowerPoint-only rendering (Bug A): fail closed, never silently degrade to builtin.

When ``powerpoint_only=True`` (the production entry points), any PowerPoint failure must raise
:class:`PowerPointUnavailableError` carrying the stable reason code + remedy, and no
LibreOffice/built-in output may be produced.  The default chain is preserved for direct/test use.
"""
import pytest
from pathlib import Path

from PIL import Image

from app.pptx_parser import Block, ReportData, SlideData
from app.qpn_renderer import (
    REASON_IMPORT,
    REASON_RENDER_FAILED,
    REASON_UNAVAILABLE,
    REASON_USER_SESSION_RUNNING,
    PowerPointStageError,
    PowerPointUnavailableError,
    SlideRenderer,
    powerpoint_remedy,
)


def _report(tmp_path):
    slide = SlideData(number=1, blocks=[
        Block(kind="title", text="Hieu qua cai tien", top=0, height=100000, shape_id=1, order=0),
        Block(kind="paragraph", text="Sau cai tien: khong phat sinh loi.", top=200000,
              height=100000, shape_id=2, order=1),
    ], width=9144000, height=5143500)
    return ReportData(path=tmp_path / "deck.pptx", slides=[slide],
                      slide_width=9144000, slide_height=5143500)


def _write_png(path: Path):
    Image.new("RGB", (8, 8), "white").save(path, "PNG")


def _no_builtin(monkeypatch):
    monkeypatch.setattr("app.qpn_renderer.render_builtin", lambda *a, **k: (_ for _ in ()).throw(
        AssertionError("builtin renderer must not run in powerpoint_only mode")))


def test_unavailable_powerpoint_raises_with_code_and_remedy(tmp_path, monkeypatch):
    monkeypatch.setattr("app.qpn_renderer.powerpoint_available", lambda: False)
    monkeypatch.setattr("app.qpn_renderer.powerpoint_probe",
                        lambda: (False, "availability", "progid-not-registered"))
    _no_builtin(monkeypatch)
    renderer = SlideRenderer(powerpoint_only=True)
    with pytest.raises(PowerPointUnavailableError) as excinfo:
        renderer.render(_report(tmp_path), [1], tmp_path / "out", purpose="learning_preview")
    assert excinfo.value.code == REASON_UNAVAILABLE
    assert "PowerPoint" in excinfo.value.remedy
    assert "PowerPoint" in str(excinfo.value)
    assert renderer.last_codes == [REASON_UNAVAILABLE]
    assert not list((tmp_path / "out").glob("*.png"))


def test_missing_pywin32_reports_import_code(tmp_path, monkeypatch):
    monkeypatch.setattr("app.qpn_renderer.powerpoint_available", lambda: False)
    monkeypatch.setattr("app.qpn_renderer.powerpoint_probe",
                        lambda: (False, "import", "pythoncom ImportError"))
    _no_builtin(monkeypatch)
    with pytest.raises(PowerPointUnavailableError) as excinfo:
        SlideRenderer(powerpoint_only=True).render(_report(tmp_path), [1], tmp_path / "out",
                                                    purpose="after_evidence")
    assert excinfo.value.code == REASON_IMPORT
    assert "pywin32" in excinfo.value.remedy


def test_user_session_refusal_raises_with_session_code_and_remedy(tmp_path, monkeypatch):
    monkeypatch.setattr("app.qpn_renderer.powerpoint_available", lambda: True)

    def _refuse(pptx, slides, out_dir, width_px):
        raise PowerPointStageError("ownership", RuntimeError("user session running"),
                                   code=REASON_USER_SESSION_RUNNING)

    monkeypatch.setattr("app.qpn_renderer.render_with_powerpoint", _refuse)
    _no_builtin(monkeypatch)
    renderer = SlideRenderer(powerpoint_only=True)
    with pytest.raises(PowerPointUnavailableError) as excinfo:
        renderer.render(_report(tmp_path), [1], tmp_path / "out", purpose="learning_preview")
    assert excinfo.value.code == REASON_USER_SESSION_RUNNING
    assert "Đóng tất cả cửa sổ PowerPoint" in str(excinfo.value)
    assert renderer.last_codes == [REASON_USER_SESSION_RUNNING]
    assert not list((tmp_path / "out").glob("*.png"))


def test_generic_powerpoint_render_failure_gets_render_failed_code(tmp_path, monkeypatch):
    monkeypatch.setattr("app.qpn_renderer.powerpoint_available", lambda: True)
    monkeypatch.setattr("app.qpn_renderer.render_with_powerpoint",
                        lambda *a, **k: (_ for _ in ()).throw(RuntimeError("COM exploded")))
    _no_builtin(monkeypatch)
    with pytest.raises(PowerPointUnavailableError) as excinfo:
        SlideRenderer(powerpoint_only=True).render(_report(tmp_path), [1], tmp_path / "out",
                                                    purpose="qpn_panel")
    assert excinfo.value.code == REASON_RENDER_FAILED


def test_powerpoint_success_still_renders(tmp_path, monkeypatch):
    monkeypatch.setattr("app.qpn_renderer.powerpoint_available", lambda: True)

    def _ok(pptx, slides, out_dir, width_px):
        target = Path(out_dir) / "slide_001.png"
        _write_png(target)
        return {1: target}

    monkeypatch.setattr("app.qpn_renderer.render_with_powerpoint", _ok)
    renderer = SlideRenderer(powerpoint_only=True)
    res = renderer.render(_report(tmp_path), [1], tmp_path / "out", purpose="learning_preview")
    assert res[1].exists()
    assert renderer.last_backend == "powerpoint"


def test_default_chain_still_falls_back_to_builtin(tmp_path, monkeypatch):
    """The fallback implementations are retained: the default (non-gated) chain degrades as before."""
    monkeypatch.setattr("app.qpn_renderer.powerpoint_available", lambda: False)
    monkeypatch.setattr("app.qpn_renderer.powerpoint_probe",
                        lambda: (False, "availability", "progid-not-registered"))
    monkeypatch.setattr("app.qpn_renderer.find_soffice", lambda: None)
    renderer = SlideRenderer()
    res = renderer.render(_report(tmp_path), [1], tmp_path / "out", purpose="learning_preview")
    assert res[1].exists()
    assert renderer.last_backend == "builtin"


def test_remedy_covers_every_stable_code():
    codes = [REASON_UNAVAILABLE, REASON_IMPORT, REASON_USER_SESSION_RUNNING,
             "powerpoint-registered-instance", "powerpoint-session-unproven",
             "powerpoint-process-table-unreadable", "powerpoint-ownership-unproven",
             "powerpoint-not-exclusive", REASON_RENDER_FAILED, "something-unknown"]
    for code in codes:
        remedy = powerpoint_remedy(code)
        assert remedy, code
        err = PowerPointUnavailableError(code)
        assert err.code == code and err.remedy == remedy and code in str(err)
    # the two answers that must never be confused keep distinct remedies
    assert powerpoint_remedy(REASON_USER_SESSION_RUNNING) != powerpoint_remedy(REASON_UNAVAILABLE)
