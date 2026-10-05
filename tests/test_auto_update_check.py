"""PROMPT-009: automatic LAN update check on start-up + one-shot notification (Build 009).

Uses the recording fake tkinter of test_gui_redesign; the update folder helpers come from test_updater.
"""
import json
import threading
import time
from pathlib import Path

import app
import app.updater as up
from app.config import AppConfig
from app.gui_controller import AUTO_CHECK_FAILED_VI, AUTO_UPDATE_CHECK_KEY, GuiController
from tests.test_gui_redesign import _make_app
from tests.test_updater import make_update_folder

SRC = Path(app.__file__).with_name("gui.py").read_text(encoding="utf-8")
CUR = app.BUILD_NUMBER                      # 11


def _ctl(tmp_path, path="", extra=None):
    return GuiController(AppConfig(update_path=path, extra=dict(extra or {})), config_path=tmp_path / "cfg" / "config.json")


def _join(c, timeout=30):
    """Deterministic completion wait: the controller's done-event is set only after the result is published, so a
    None result can never be mistaken for a finished check."""
    if c._update_thread:
        c._update_thread.join(timeout)
    assert c.update_check_done.wait(timeout), "update worker did not complete"
    assert not c.update_busy


def _startup(a):
    fn = next(fn for ms, fn in a.root.scheduled if fn == a._startup_update_check)
    fn()
    _join(a.ctl)
    a._poll()


# ------------------------------------------------------------------ 1-3 setting
def test_auto_check_enabled_by_default_and_persists(tmp_path):
    c = _ctl(tmp_path)
    assert c.auto_update_check is True
    c.set_auto_update_check(False)
    data = json.loads((tmp_path / "cfg" / "config.json").read_text(encoding="utf-8"))
    assert data[AUTO_UPDATE_CHECK_KEY] is False
    again = _ctl(tmp_path, extra={AUTO_UPDATE_CHECK_KEY: False})
    assert again.auto_update_check is False
    assert _ctl(tmp_path, extra={AUTO_UPDATE_CHECK_KEY: "false"}).auto_update_check is False
    assert _ctl(tmp_path, extra={AUTO_UPDATE_CHECK_KEY: "true"}).auto_update_check is True


def test_disabled_setting_prevents_startup_check_but_not_manual(tmp_path):
    folder = make_update_folder(tmp_path / "Update", version="1.2.2", build=CUR + 1)
    c = _ctl(tmp_path, str(folder), extra={AUTO_UPDATE_CHECK_KEY: False})
    assert not c.check_update_async(startup=True) and c.update_check is None and not c.update_busy
    assert c.check_update_async()                                       # manual still works
    _join(c)
    assert c.update_check.status == "available"


def test_gui_checkbox_toggles_and_persists(monkeypatch, tmp_path):
    gui, a, reg, _ = _make_app(monkeypatch, tmp_path)
    assert a.chk_auto_update.cfg["text"] == "Tự động kiểm tra cập nhật khi khởi động"
    assert a.var_auto_update.get() is True and a.chk_auto_update.master is a.cfg_cards["update"]
    assert a.chk_auto_update in a._config_widgets                       # locked while a batch runs
    a.var_auto_update.set(False)
    a._on_auto_update_toggled()
    assert a.ctl.auto_update_check is False and a.ctl.cfg.extra[AUTO_UPDATE_CHECK_KEY] is False


# ------------------------------------------------------------------ 4, 15 non-blocking / Tk thread
def test_startup_gui_not_blocked_by_slow_unc_check(monkeypatch, tmp_path):
    gate = threading.Event()

    def slow(path, *a, **k):
        gate.wait(5)
        return up.UpdateCheck("inaccessible", up.MSG["inaccessible"], path)
    monkeypatch.setattr(up, "check_for_update", slow)
    cfg = AppConfig(update_path=r"\\BUILD-PC\ReportExtractor_Update")
    monkeypatch.setattr("tests.test_gui_redesign.AppConfig", lambda **k: cfg)
    t0 = time.monotonic()
    gui, a, reg, _ = _make_app(monkeypatch, tmp_path)
    assert time.monotonic() - t0 < 2.0                                  # GUI constructed; check only scheduled
    assert a.ctl._update_thread is None
    fn = next(fn for ms, fn in a.root.scheduled if fn == a._startup_update_check)
    fn()
    assert time.monotonic() - t0 < 3.0 and a.ctl.update_busy           # worker started, Tk thread free
    assert not a.ctl.update_check_done.is_set() and a.ctl.update_check is None   # "pending", not "failed"
    assert a.lbl_update.cfg["text"] == "Đang kiểm tra cập nhật…"
    gate.set()
    _join(a.ctl)
    a._poll()                                                           # widgets updated by the poll loop only
    assert a.lbl_update.cfg["text"] == AUTO_CHECK_FAILED_VI


def test_worker_never_touches_widgets():
    worker = SRC[SRC.index("def _startup_update_check("):SRC.index("def pick_update_path(")]
    assert "lbl_update" not in worker                                   # only schedules the controller worker
    ctl_src = Path(app.__file__).with_name("gui_controller.py").read_text(encoding="utf-8")
    body = ctl_src[ctl_src.index("def check_update("):ctl_src.index("def check_update_async(")]
    assert "update_dirty = True" in body and "lbl_" not in body         # flag consumed by _poll on the Tk thread
    assert "if self.ctl.update_dirty:\n                self._render_update()" in SRC


# ------------------------------------------------------------------ 5-8 comparison + manifest-only read
def test_build_comparison_008_009_scenarios(tmp_path):
    folder = make_update_folder(tmp_path / "Update", version="1.1.0-beta", build=9)
    assert up.check_for_update(str(folder), current_build=8).status == "available"   # 008 -> 009
    assert up.check_for_update(str(folder), current_build=9).status == "latest"      # 009 -> 009
    assert up.check_for_update(str(folder), current_build=10).status == "older"      # 010 -> 009: no downgrade
    r = up.check_for_update(str(folder), current_build=8)
    assert r.info.label() == "1.1.0-beta — Build 009" and r.message == "Có phiên bản mới: 1.1.0-beta — Build 009"


def test_version_check_reads_only_manifest_and_never_copies_zip(tmp_path, monkeypatch):
    folder = make_update_folder(tmp_path / "Update", version="1.2.2", build=CUR + 1)
    zip_path = next(folder.glob("*.zip"))
    reads = []
    real_open = Path.open

    def spy_open(self, *a, **k):
        reads.append(self.name)
        return real_open(self, *a, **k)
    monkeypatch.setattr(Path, "open", spy_open)
    monkeypatch.setattr(up, "sha256_file", lambda *a, **k: (_ for _ in ()).throw(AssertionError("sha over LAN")))
    monkeypatch.setattr(up, "stage_update", lambda *a, **k: (_ for _ in ()).throw(AssertionError("staged")))
    c = _ctl(tmp_path, str(folder))
    assert c.check_update_async(startup=True)
    _join(c)
    assert c.update_check.status == "available"
    assert zip_path.name not in reads and not (tmp_path / "update_staging").exists()
    assert not list(Path(tmp_path).rglob("update_staging"))


# ------------------------------------------------------------------ 9-10 unavailable build PC
def test_unavailable_path_does_not_crash_and_shows_soft_status(monkeypatch, tmp_path):
    """Build PC off / LAN down.  Real UNC name resolution on Windows can take tens of seconds (that is exactly why
    the check runs in a worker), so the two UNC cases use a deterministic offline stand-in for the read-only
    check; the local missing folder exercises the real updater code path."""
    real_check = up.check_for_update
    seen = []

    def offline_check(path, *a, **k):
        seen.append(path)
        if str(path).startswith("\\\\"):
            return up.UpdateCheck("inaccessible", up.MSG["inaccessible"], path)
        return real_check(path, *a, **k)
    monkeypatch.setattr(up, "check_for_update", offline_check)
    for path in (r"\\BUILD-PC\ReportExtractor_Update", r"\\192.168.1.50\ReportExtractor_Update",
                 str(tmp_path / "missing_share")):
        cfg = AppConfig(update_path=path)
        monkeypatch.setattr("tests.test_gui_redesign.AppConfig", lambda **k: cfg)
        popups = []
        gui, a, reg, _ = _make_app(monkeypatch, tmp_path)
        monkeypatch.setattr(gui.messagebox, "showerror", lambda *x, **k: popups.append(x))
        monkeypatch.setattr(gui.messagebox, "showwarning", lambda *x, **k: popups.append(x))
        assert a.ctl.update_check is None and a.ctl.update_check_done.is_set()      # not started yet (not "pending")
        _startup(a)
        assert a.ctl.update_check is not None, "worker finished but published no result"
        assert a.ctl.update_check.status in ("inaccessible", "no_permission", "no_manifest")
        assert seen[-1] == path
        assert a.lbl_update.cfg["text"] == AUTO_CHECK_FAILED_VI and a.lbl_update.cfg["style"] == "Secondary.TLabel"
        assert popups == [] and a.btn_check_update.cfg["state"] == "normal"
        assert a.btn_install_update.cfg["state"] == "disabled" and a.ctl.update_path == path    # UNC kept


def test_manual_check_still_reports_explicit_errors(monkeypatch, tmp_path):
    gui, a, reg, _ = _make_app(monkeypatch, tmp_path)
    a.var_update_path.set(str(tmp_path / "nope"))
    a.check_update()
    _join(a.ctl)
    a._poll()
    assert a.lbl_update.cfg["text"] == up.MSG["inaccessible"] and a.lbl_update.cfg["style"] == "Warning.TLabel"


# ------------------------------------------------------------------ 12-14 notification once / Để sau / Cập nhật ngay
def _gui_with_update(monkeypatch, tmp_path, build=CUR + 1):
    folder = make_update_folder(tmp_path / "Update", version="1.2.2", build=build)
    cfg = AppConfig(update_path=str(folder))
    monkeypatch.setattr("tests.test_gui_redesign.AppConfig", lambda **k: cfg)
    gui, a, reg, _ = _make_app(monkeypatch, tmp_path)
    return gui, a, reg, folder


def test_startup_detects_newer_build_and_offers_once(monkeypatch, tmp_path):
    gui, a, reg, folder = _gui_with_update(monkeypatch, tmp_path)
    offers = []
    monkeypatch.setattr(a, "_confirm_update", lambda label: offers.append(label) or False)   # user: Để sau
    _startup(a)
    assert offers == ["1.2.2 — Build 013"]
    assert a.lbl_update.cfg["text"] == "Có phiên bản mới: 1.2.2 — Build 013"
    assert a.btn_install_update.cfg["state"] == "normal"
    for _ in range(3):                                                  # polls / re-renders: no repeated popup
        a._poll()
        a._render_update()
    a.check_update()                                                    # manual check of the same build
    _join(a.ctl)
    a._poll()
    assert offers == ["1.2.2 — Build 013"] and a.lbl_update.cfg["text"].startswith("Có phiên bản mới")
    # a NEWER remote build published later in the same session is announced again (once)
    make_update_folder(folder, version="1.2.2", build=CUR + 2)
    a.check_update()
    _join(a.ctl)
    a._poll()
    assert offers == ["1.2.2 — Build 013", "1.2.2 — Build 014"]


def test_same_or_older_build_is_not_offered(monkeypatch, tmp_path):
    for build, status in ((CUR, "latest"), (CUR - 1, "older")):
        gui, a, reg, folder = _gui_with_update(monkeypatch, tmp_path / str(build), build=build)
        offers = []
        monkeypatch.setattr(a, "_confirm_update", lambda label: offers.append(label) or False)
        _startup(a)
        assert a.ctl.update_check.status == status and offers == [] and a.btn_install_update.cfg["state"] == "disabled"


def test_update_now_routes_into_existing_installer(monkeypatch, tmp_path):
    gui, a, reg, folder = _gui_with_update(monkeypatch, tmp_path)
    monkeypatch.setattr(a, "_confirm_update", lambda label: True)       # user: Cập nhật ngay
    calls = []
    monkeypatch.setattr(a.ctl, "install_update_async", lambda *x, **k: calls.append(1) or True)
    _startup(a)
    assert calls == [1]                                                 # existing updater path (stage/verify/apply)
    assert a.update_progress_win is not None                            # PROMPT-011 progress window opened
    confirm_src = SRC[SRC.index("def _confirm_update("):SRC.index("def install_update(")]
    assert '"Cập nhật ngay"' in confirm_src and '"Để sau"' in confirm_src and "Phiên bản hiện tại" in confirm_src
    inst = SRC[SRC.index("def install_update("):SRC.index("def _open_update_progress(")]
    assert "c.install_update_async()" in inst
    done = SRC[SRC.index("def _on_update_done("):SRC.index("def _render_event(")]
    assert "self.root.after(300, self.root.destroy)" in done


def test_notification_never_auto_installs(monkeypatch, tmp_path):
    gui, a, reg, folder = _gui_with_update(monkeypatch, tmp_path)
    monkeypatch.setattr(a, "_confirm_update", lambda label: False)
    staged = []
    monkeypatch.setattr(up, "stage_update", lambda *x, **k: staged.append(1))
    _startup(a)
    assert staged == [] and a.root.scheduled[-1][1] != a.root.destroy


# ------------------------------------------------------------------ 16-17 path kinds
def test_unc_and_local_paths_accepted(tmp_path):
    c = _ctl(tmp_path)
    for p in (r"\\BUILD-PC\ReportExtractor_Update", r"\\192.168.1.50\ReportExtractor_Update", r"I:\QPn"):
        c.set_update_path(p)
        assert c.update_path == p and c.check_update().status in ("inaccessible", "no_permission", "no_manifest")
    folder = make_update_folder(tmp_path / "local", version="1.2.2", build=CUR + 1)
    c.set_update_path(str(folder))
    assert c.check_update().status == "available"
    c.save_settings()
    saved = json.loads((tmp_path / "cfg" / "config.json").read_text(encoding="utf-8"))
    assert saved["update_path"] == str(folder)


# ------------------------------------------------------------------ 18 layout + version
def test_update_card_layout_and_build_009(monkeypatch, tmp_path):
    gui, a, reg, _ = _make_app(monkeypatch, tmp_path)
    body = a.cfg_cards["update"]
    rows = {w.grid_info_.get("row") for w in reg["widgets"] if getattr(w, "master", None) is body and w.grid_info_}
    assert rows == {0, 1, 2, 3, 4, 5} and a.chk_auto_update.grid_info_["row"] == 2 < a.lbl_update.grid_info_["row"]
    assert a.lbl_cur_version.cfg["text"] == "Phiên bản hiện tại: 1.2.1 — Build 012"
    assert app.__version__ == "1.2.1" and app.BUILD_NUMBER == 12 and app.BUILD_LABEL == "Build 012"
    assert a.cfg_page.row_weights == {4: 1}                             # responsive settings page unchanged


def test_pending_vs_finished_state_contract(monkeypatch, tmp_path):
    """None result + done-event cleared = still running; done-event set => update_check is a completed result."""
    gate = threading.Event()

    def slow(path, *a, **k):
        gate.wait(5)
        return up.UpdateCheck("inaccessible", up.MSG["inaccessible"], path)
    monkeypatch.setattr(up, "check_for_update", slow)
    c = _ctl(tmp_path, r"\\BUILD-PC\ReportExtractor_Update")
    assert c.update_check_done.is_set() and c.update_check is None             # idle, never checked
    assert c.check_update_async(startup=True)
    assert not c.update_check_done.is_set() and c.update_busy and c.update_check is None   # pending
    assert not c.check_update_async()                                           # one worker at a time
    gate.set()
    assert c.update_check_done.wait(10)
    assert c.update_check is not None and c.update_check.status == "inaccessible" and not c.update_busy
    assert c.update_dirty and c.update_status_text() == AUTO_CHECK_FAILED_VI


def test_legacy_beta_client_sees_1_2_0_build_011_as_normal_update(tmp_path):
    """Compatibility: a 1.1.0-beta / Build 010 client must treat 1.2.0 / Build 011 as an ordinary update
    (numeric build authoritative; the version-string change is irrelevant for ordering)."""
    folder = make_update_folder(tmp_path / "Update", version="1.2.0", build=11)
    manifest = json.loads((folder / "version.json").read_text(encoding="utf-8"))
    assert manifest["version"] == "1.2.0" and manifest["build"] == 11 and manifest["package"] == "ReportExtractor_1.2.0.zip"
    r = up.check_for_update(str(folder), current_build=10, current_version="1.1.0-beta")
    assert r.status == "available" and r.message == "Có phiên bản mới: 1.2.0 — Build 011"
    assert up.check_for_update(str(folder), current_build=11, current_version="1.2.0").status == "latest"
    assert up.check_for_update(str(folder), current_build=12, current_version="1.2.2").status == "older"
    assert up.check_for_update(str(folder), current_build=12, current_version="1.2.1").status == "older"
    assert up.version_label() == "1.2.1 — Build 012"
    import build_portable as bp
    assert bp.package_name(app.__version__) == "ReportExtractor_1.2.1.zip" and bp._version() == ("1.2.1", "012")
