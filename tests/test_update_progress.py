"""PROMPT-011: real byte progress for the LAN update (copy -> verify -> handoff) and the progress window (Build 011)."""
import hashlib
import json
import os
import threading
from pathlib import Path

import app
import app.gui_controller as gc
import app.updater as up
from app.config import AppConfig
from app.gui_controller import UPDATE_STAGES_VI, GuiController, UpdateProgress, format_bytes
from tests.test_gui_redesign import _make_app
from tests.test_updater import make_portable, make_release_zip

SRC = Path(app.__file__).with_name("gui.py").read_text(encoding="utf-8")
CTL_SRC = Path(app.__file__).with_name("gui_controller.py").read_text(encoding="utf-8")
NEW_V, NEW_B = "1.2.2", app.BUILD_NUMBER + 1


class _TestClock:
    """Deterministic monotonic clock for progress-speed tests (no wall-clock sleeps)."""
    def __init__(self, start=0.0, step=0.0):
        self.value = float(start)
        self.step = float(step)
        self.lock = threading.Lock()

    def __call__(self):
        with self.lock:
            now = self.value
            self.value += self.step
            return now


def _big_update_folder(tmp_path, size_mb=3, sha=True):
    """Release ZIP with a few MB of payload so the chunked copy produces many progress callbacks."""
    folder = tmp_path / "Update"
    pkg = folder / f"ReportExtractor_{NEW_V}.zip"
    make_release_zip(pkg, version=NEW_V, build=NEW_B, extra={"_internal/blob.bin": os.urandom(size_mb * 1024 * 1024)})   # incompressible -> real MBs
    data = {"version": NEW_V, "build": NEW_B, "package": pkg.name}
    if sha:
        data["sha256"] = hashlib.sha256(pkg.read_bytes()).hexdigest()
    (folder / "version.json").write_text(json.dumps(data), encoding="utf-8")
    return folder, pkg


def _ctl(tmp_path, folder):
    return GuiController(AppConfig(update_path=str(folder)), config_path=tmp_path / "cfg" / "config.json")


# ------------------------------------------------------------------ 1-3 real byte progress (updater layer)
def test_copy_with_progress_reports_real_monotonic_bytes(tmp_path):
    src = tmp_path / "pkg.bin"
    src.write_bytes(b"x" * (5 * 1024 * 1024 + 123))
    dst = tmp_path / "out" / "pkg.bin"
    dst.parent.mkdir()
    events = []
    copied = up.copy_with_progress(src, dst, lambda d, t, p: events.append((d, t, p)), chunk=1 << 20)
    assert copied == src.stat().st_size == dst.stat().st_size and dst.read_bytes() == src.read_bytes()
    done = [e[0] for e in events]
    assert done == sorted(done) and done[0] == 0 and done[-1] == src.stat().st_size   # monotonic, real bytes
    assert len(events) >= 7                                                            # 0 + 6 chunks
    assert all(t == src.stat().st_size for _, t, _ in events)
    assert events[-1][2] == 100.0 and all(0.0 <= p <= 100.0 for _, _, p in events)
    assert all(abs(p - d * 100.0 / t) < 1e-6 for d, t, p in events)                    # percent derived from bytes


def test_stage_update_emits_prepare_copy_verify_ready_in_order(tmp_path):
    folder, pkg = _big_update_folder(tmp_path)
    portable = make_portable(tmp_path / "app")
    chk = up.check_for_update(str(folder), current_build=app.BUILD_NUMBER)
    assert chk.available
    stages = []
    staged = up.stage_update(chk, portable, progress=lambda st, d, t, p: stages.append((st, d, t, p)))
    names = [s[0] for s in stages]
    assert names[0] == "PREPARING" and names[-1] == "READY"
    assert names.index("COPYING") < names.index("VERIFYING") < names.index("READY")      # verify AFTER copy
    copying = [s for s in stages if s[0] == "COPYING"]
    assert copying[-1][3] == 100.0 and copying[-1][1] == pkg.stat().st_size            # final copy = 100 %
    assert Path(staged.zip_path).stat().st_size == pkg.stat().st_size
    verifying = next(s for s in stages if s[0] == "VERIFYING")
    assert verifying[1] == verifying[2] == pkg.stat().st_size and verifying[3] == 100.0


# ------------------------------------------------------------------ 4, 6-8 controller worker + events
def test_install_async_publishes_snapshots_from_worker_without_tk(tmp_path, monkeypatch):
    folder, pkg = _big_update_folder(tmp_path)
    portable = make_portable(tmp_path / "app")
    monkeypatch.setattr(gc, "portable_root", lambda: portable)
    c = _ctl(tmp_path, folder)
    c.check_update()
    clock = _TestClock(start=100.0, step=1.0)  # deterministic, measurable intervals in the worker
    c._clock = clock
    spawned, threads = [], []

    def spawn(cmd, **kw):
        spawned.append(cmd)
        threads.append(threading.current_thread().name)
    assert c.install_update_async(spawn=spawn)
    assert c.update_installing and not c.update_available()                           # button state source
    assert not c.install_update_async(spawn=spawn)                                     # duplicate start rejected
    assert c.install_update(spawn=spawn)[0] is False                                   # sync entry also refused
    c._install_thread.join(30)
    events = c.pump()
    kinds = [e.kind for e in events]
    assert kinds[-1] == "update_done" and events[-1].payload[0] is True
    assert threads == ["update-install"] and len(spawned) == 1                          # handoff once, off Tk thread
    snaps = [e.payload for e in events if e.kind == "update_progress"]
    assert all(isinstance(s, UpdateProgress) for s in snaps)                           # plain data, no widgets
    stages = [s.stage for s in snaps]
    assert stages[0] == "PREPARING" and stages[-1] == "HANDOFF"
    assert stages.index("COPYING") < stages.index("VERIFYING") < stages.index("READY") < stages.index("HANDOFF")
    copying = [s for s in snaps if s.stage == "COPYING"]
    pcts = [s.percent for s in copying]
    assert pcts == sorted(pcts) and pcts[-1] == 100.0
    assert copying[-1].percent == 100.0
    assert copying[-1].bytes_copied == copying[-1].total_bytes == pkg.stat().st_size
    assert snaps[-1].percent == 100.0 and snaps[-1].bytes_copied == pkg.stat().st_size
    assert any(s.speed_bps > 0 for s in copying)
    for i, snapshot in enumerate(copying[1:], start=1):
        expected_speed = snapshot.bytes_copied / (i * clock.step)
        assert abs(snapshot.speed_bps - expected_speed) < 1e-6  # bytes / measured time, not a fabricated rate
    # handoff happened strictly after READY (verified) – the spawn call is the last thing the worker did
    assert events[-2].payload.stage == "HANDOFF"


def test_copy_speed_handles_a_zero_monotonic_origin(tmp_path):
    c = _ctl(tmp_path, tmp_path)
    clock = _TestClock(start=0.0, step=0.25)
    c._clock = clock
    start = [None]
    c._publish_update_progress("COPYING", 0, 100, 0.0, emit=False, t0=start)
    assert start == [0.0] and c.update_progress.speed_bps == 0.0
    c._publish_update_progress("COPYING", 100, 100, 100.0, emit=False, t0=start)
    assert c.update_progress.speed_bps == 400.0  # 100 measured bytes / 0.25 measured seconds


def test_fast_small_update_has_no_invented_speed_when_clock_does_not_advance(tmp_path, monkeypatch):
    """A tiny copy may finish within one clock tick; zero speed is honest when elapsed time is unmeasurable."""
    folder, pkg = _big_update_folder(tmp_path, size_mb=0)
    assert pkg.stat().st_size < 1 << 20
    portable = make_portable(tmp_path / "app")
    monkeypatch.setattr(gc, "portable_root", lambda: portable)
    c = _ctl(tmp_path, folder)
    c.check_update()
    c._clock = _TestClock(start=50.0, step=0.0)
    spawned = []

    assert c.install_update_async(spawn=lambda cmd, **kw: spawned.append((cmd, threading.current_thread().name)))
    c._install_thread.join(30)
    events = c.pump()
    assert events[-1].kind == "update_done" and events[-1].payload[0] is True
    assert len(spawned) == 1 and spawned[0][1] == "update-install"
    snaps = [e.payload for e in events if e.kind == "update_progress"]
    stages = [s.stage for s in snaps]
    assert stages[0] == "PREPARING" and stages[-1] == "HANDOFF"
    assert stages.index("COPYING") < stages.index("VERIFYING") < stages.index("READY") < stages.index("HANDOFF")
    copying = [s for s in snaps if s.stage == "COPYING"]
    assert copying[0].bytes_copied == 0 and copying[0].percent == 0.0
    assert copying[-1].bytes_copied == pkg.stat().st_size and copying[-1].percent == 100.0
    assert all(s.speed_bps == 0.0 for s in copying)  # no elapsed tick -> unavailable, not a guessed rate
    assert snaps[-1].bytes_copied == pkg.stat().st_size and snaps[-1].percent == 100.0


def test_copy_failure_shows_error_and_never_hands_off(tmp_path, monkeypatch):
    folder, pkg = _big_update_folder(tmp_path)
    portable = make_portable(tmp_path / "app")
    monkeypatch.setattr(gc, "portable_root", lambda: portable)
    c = _ctl(tmp_path, folder)
    c.check_update()
    real = up.copy_with_progress

    def flaky(src, dst, progress=None, chunk=1 << 20):
        calls = {"n": 0}

        def p(d, t, pct):
            calls["n"] += 1
            if progress:
                progress(d, t, pct)
            if calls["n"] == 3:
                raise OSError(64, "The specified network name is no longer available")
        return real(src, dst, p, chunk)
    monkeypatch.setattr(up, "copy_with_progress", flaky)
    spawned = []
    assert c.install_update_async(spawn=lambda cmd, **k: spawned.append(cmd))
    c._install_thread.join(30)
    events = c.pump()
    assert events[-1].kind == "update_done" and events[-1].payload[0] is False
    assert "Không sao chép được" in events[-1].payload[1]
    assert events[-2].payload.stage == "ERROR" and events[-2].payload.error
    assert spawned == [] and not c.update_installing and c.update_available()          # can retry later
    assert not (portable / up.STAGING_DIR).exists()                                     # staging cleaned


def test_sha_mismatch_never_proceeds_to_apply(tmp_path, monkeypatch):
    folder, pkg = _big_update_folder(tmp_path)
    m = json.loads((folder / "version.json").read_text(encoding="utf-8"))
    m["sha256"] = "0" * 64
    (folder / "version.json").write_text(json.dumps(m), encoding="utf-8")
    portable = make_portable(tmp_path / "app")
    monkeypatch.setattr(gc, "portable_root", lambda: portable)
    c = _ctl(tmp_path, folder)
    c.check_update()
    spawned = []
    assert c.install_update_async(spawn=lambda cmd, **k: spawned.append(cmd))
    c._install_thread.join(30)
    events = c.pump()
    stages = [e.payload.stage for e in events if e.kind == "update_progress"]
    assert "VERIFYING" in stages and "READY" not in stages and "HANDOFF" not in stages and stages[-1] == "ERROR"
    assert spawned == [] and events[-1].payload == (False, up.MSG["bad_package"])
    assert (portable / up.APP_EXE).read_bytes() == b"MZ-old"


def test_sync_install_update_still_works_for_cli_and_tests(tmp_path, monkeypatch):
    folder, pkg = _big_update_folder(tmp_path, size_mb=1)
    portable = make_portable(tmp_path / "app")
    monkeypatch.setattr(gc, "portable_root", lambda: portable)
    c = _ctl(tmp_path, folder)
    c.check_update()
    spawned = []
    ok, msg = c.install_update(spawn=lambda cmd, **k: spawned.append(cmd))
    assert ok and spawned and c.update_progress.stage == "HANDOFF" and c.update_progress.percent == 100.0


# ------------------------------------------------------------------ 5, 9-10 GUI: window, buttons, close behaviour
def _gui(monkeypatch, tmp_path):
    folder, pkg = _big_update_folder(tmp_path)
    portable = make_portable(tmp_path / "app")
    monkeypatch.setattr(gc, "portable_root", lambda: portable)
    cfg = AppConfig(update_path=str(folder))
    monkeypatch.setattr("tests.test_gui_redesign.AppConfig", lambda **k: cfg)
    gui, a, reg, _ = _make_app(monkeypatch, tmp_path)
    a.ctl.check_update()
    a._poll()
    return gui, a, reg, pkg, portable


def test_gui_progress_window_and_buttons_during_update(monkeypatch, tmp_path):
    gui, a, reg, pkg, portable = _gui(monkeypatch, tmp_path)
    a.ctl.offered_update_builds.add(NEW_B)                       # no startup offer dialog in this test
    monkeypatch.setattr(a, "_confirm_update", lambda label: True)
    gate = threading.Event()
    real = up.copy_with_progress
    monkeypatch.setattr(up, "copy_with_progress", lambda s, d, progress=None, chunk=1 << 20:
                        (gate.wait(10), real(s, d, progress, chunk))[1])
    spawned = []
    monkeypatch.setattr(up, "launch_updater", lambda staged, spawn=None: spawned.append(staged) or ["cmd"])
    assert a.btn_install_update.cfg["state"] == "normal"
    a.install_update()
    assert a.update_progress_win is not None and a.ctl.update_installing
    assert a.btn_install_update.cfg["state"] == "disabled" and a.btn_check_update.cfg["state"] == "disabled"
    assert a.btn_upd_close.cfg["state"] == "disabled"
    n_tops = len(reg["toplevels"])
    a.install_update()                                            # second click: nothing new
    a.install_update(confirmed=True)
    assert len(reg["toplevels"]) == n_tops and not a.ctl.install_update_async()
    # closing the window while copying only hides it – the worker keeps staging
    a._on_update_progress_close()
    assert a.update_progress_win is not None and a.ctl.update_installing
    gate.set()
    a.ctl._install_thread.join(30)
    a._poll()                                                     # Tk thread renders the queued snapshots
    assert a.pb_update["value"] == 100 and a.pb_update.cfg["mode"] == "determinate"
    assert a.lbl_upd_stage.cfg["text"] == UPDATE_STAGES_VI["HANDOFF"]
    assert format_bytes(pkg.stat().st_size) in a.lbl_upd_bytes.cfg["text"] and "100%" in a.lbl_upd_bytes.cfg["text"]
    assert spawned and (300, a.root.destroy) in a.root.scheduled     # handoff only after verified staging
    assert (portable / up.APP_EXE).read_bytes() == b"MZ-old"           # running EXE never touched by the GUI


def test_gui_error_path_reenables_buttons(monkeypatch, tmp_path):
    gui, a, reg, pkg, portable = _gui(monkeypatch, tmp_path)
    a.ctl.offered_update_builds.add(NEW_B)
    monkeypatch.setattr(a, "_confirm_update", lambda label: True)
    monkeypatch.setattr(up, "copy_with_progress", lambda *x, **k: (_ for _ in ()).throw(OSError("LAN lost")))
    errors = []
    monkeypatch.setattr(gui.messagebox, "showerror", lambda *x, **k: errors.append(x))
    a.install_update()
    a.ctl._install_thread.join(30)
    a._poll()
    assert errors and "Cập nhật thất bại" in errors[-1][1] and "LAN lost" in errors[-1][1]
    assert a.lbl_upd_stage.cfg["style"] == "Error.TLabel" and a.btn_upd_close.cfg["state"] == "normal"
    assert a.btn_install_update.cfg["state"] == "normal" and not a.ctl.update_installing
    assert (300, a.root.destroy) not in a.root.scheduled               # application stays open
    a._on_update_progress_close()
    assert a.update_progress_win is None


def test_progress_rendering_is_confined_to_tk_thread_code():
    worker = CTL_SRC[CTL_SRC.index("def _publish_update_progress("):CTL_SRC.index("def save_settings(")]
    assert "lbl_" not in worker and "configure(" not in worker and '"update_progress"' in worker
    render = SRC[SRC.index("def _render_event("):]
    assert 'ev.kind == "update_progress"' in render and 'ev.kind == "update_done"' in render
    assert "time.sleep" not in SRC[SRC.index("def _open_update_progress("):SRC.index("def _render_event(")]
    stage_src = up.__file__ and Path(up.__file__).read_text(encoding="utf-8")
    copy_src = stage_src[stage_src.index("def copy_with_progress("):stage_src.index("def stage_update(")]
    assert "progress(copied, total" in copy_src and "time." not in copy_src and "sleep" not in copy_src


def test_labels_and_build_011():
    for key in ("PREPARING", "COPYING", "VERIFYING", "READY", "APPLYING", "BACKUP", "INSTALLING", "COMPLETE", "ERROR"):
        assert UPDATE_STAGES_VI[key]
    assert UPDATE_STAGES_VI["COPYING"] == "Đang sao chép bản cập nhật..." and UPDATE_STAGES_VI["ERROR"] == "Cập nhật thất bại"
    assert format_bytes(87.1 * (1 << 20)) == "87.1 MB" and format_bytes(11.2 * (1 << 20)) == "11.2 MB"
    p = UpdateProgress(stage="COPYING", bytes_copied=int(58.4 * (1 << 20)), total_bytes=int(87.1 * (1 << 20)),
                       percent=67.0, speed_bps=11.2 * (1 << 20))
    assert p.bytes_text == "58.4 MB / 87.1 MB" and p.speed_text == "Tốc độ: 11.2 MB/s"
    assert app.__version__ == "1.3.3" and app.BUILD_NUMBER == 16 and app.BUILD_LABEL == "Build 016"
