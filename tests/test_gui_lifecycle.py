"""GUI lifecycle regressions: KeyError 'master_not_found' killed _poll(), batch finished but GUI stayed running.

The real view (app.gui) is built on a fake tkinter so the REAL _poll/_render_event/_on_done path runs headless.
"""
import importlib
import sys
import threading
import types
from types import SimpleNamespace

import pytest

from app.batch_processor import BatchSummary
from app.config import AppConfig
from app.gui_controller import STAGE_WEIGHTS, Progress, UiEvent
from app.prescan import PRESCAN_COUNT_KEYS, PreScanResult

from tests.test_gui_progress_ollama import FakeClock, _begin, _ctl_with_rows


# ------------------------------------------------------------------ fake tkinter view harness
def _make_view(monkeypatch, tmp_path, n_files=1):
    class Widget:
        def __init__(self, *a, **k): self.k = k; self.cfg = {}
        def __getattr__(self, name): return lambda *a, **k: None
        def __setitem__(self, k, v): self.cfg[k] = v
        def __getitem__(self, k): return self.cfg.get(k, "")
        def configure(self, *a, **k): self.cfg.update(k)
        def get_children(self): return []
        def insert(self, *a, **k): return f"I{len(self.cfg)}"
        def selection(self): return ()
        def get(self): return self.k.get("value", "")
    class Tree(Widget):
        def __init__(self, *a, **k): super().__init__(*a, **k); self.items = {}; self._n = 0
        def insert(self, *a, **k): self._n += 1; iid = f"I{self._n}"; self.items[iid] = k; return iid
        def item(self, iid, **k): self.items[iid].update(k)
        def get_children(self): return list(self.items)
        def delete(self, *iids): [self.items.pop(i, None) for i in iids]
    class Var(Widget):
        def __init__(self, value="", **k): self.v = value; self.cfg = {}
        def set(self, v): self.v = v
        def get(self): return self.v
    class Root(Widget):
        def __init__(self, *a, **k): self.cfg = {}; self.scheduled = []
        def after(self, ms, fn=None): self.scheduled.append((ms, fn))
    class Top(Widget):
        created = []
        def __init__(self, *a, **k): super().__init__(*a, **k); Top.created.append(self)
    tkmod = types.ModuleType("tkinter")
    for n in ("Text", "Menu", "Canvas"):
        setattr(tkmod, n, Widget)
    tkmod.Toplevel = Top
    tkmod.Tk, tkmod.StringVar, tkmod.BooleanVar, tkmod.TclError = Root, Var, Var, Exception
    ttkmod = types.ModuleType("tkinter.ttk")
    for n in ("Style", "Label", "LabelFrame", "Frame", "Entry", "Button", "Combobox", "Checkbutton",
              "Scrollbar", "Progressbar", "Radiobutton", "Spinbox", "Notebook", "PanedWindow", "Separator"):
        setattr(ttkmod, n, Widget)
    ttkmod.Treeview = Tree
    fd, mb = types.ModuleType("tkinter.filedialog"), types.ModuleType("tkinter.messagebox")
    for m in (fd, mb):
        m.__getattr__ = lambda name: (lambda *a, **k: True)
    tkmod.ttk, tkmod.filedialog, tkmod.messagebox = ttkmod, fd, mb
    for name, mod in (("tkinter", tkmod), ("tkinter.ttk", ttkmod), ("tkinter.filedialog", fd), ("tkinter.messagebox", mb)):
        monkeypatch.setitem(sys.modules, name, mod)
    monkeypatch.delitem(sys.modules, "app.gui", raising=False)
    gui = importlib.import_module("app.gui")
    folder = tmp_path / "reports"
    folder.mkdir(exist_ok=True)
    for i in range(n_files):
        (folder / f"26092{i:04d}-VOC_r{i}.pptx").write_bytes(b"x")
    app = gui.ReportExtractorApp(AppConfig(last_report_folder=str(folder), period_mode="all"))
    for t in threading.enumerate():                 # the view kicked off the background pre-scan: let it finish
        if t.name == "prescan":
            t.join(timeout=10)
    app.ctl.pump()
    clock = FakeClock()
    app.ctl._clock = clock
    _begin(app.ctl, clock)                       # controller 'running' exactly like start(), without a worker
    app._done_handled = False
    app._render_rows()
    app._set_running(True)
    app.root.scheduled.clear()
    Top.created.clear()
    return gui, app, clock, Top


def _polls(app, n=3):
    for _ in range(n):
        app._poll()


def _row_status(app, i):
    return app.tree.items[app.row_items[i]]["values"][6]


# ------------------------------------------------------------------ #1 / #14 canonical schema + exact KeyError
def test_prescan_counts_schema_is_canonical(tmp_path):
    res = PreScanResult(period=SimpleNamespace(label_vi=lambda: "all"))
    assert tuple(res.counts()) == PRESCAN_COUNT_KEYS and "master_not_found" in PRESCAN_COUNT_KEYS
    import pathlib
    gui_src = pathlib.Path("app/gui.py").read_text(encoding="utf-8")
    assert "c['master_not_found']" not in gui_src and "prescan_stage_text(" in gui_src   # default-safe access only


def test_prescan_event_without_master_not_found_does_not_kill_poll(monkeypatch, tmp_path):
    gui, app, clock, Top = _make_view(monkeypatch, tmp_path, 1)
    counts = {k: 0 for k in PRESCAN_COUNT_KEYS}
    counts.update(discovered=1, candidates=1, incomplete=1)
    fake = SimpleNamespace(counts=lambda: counts, candidates=[SimpleNamespace(index=0)],
                           summary_lines_vi=lambda: ["x"])
    app.ctl._queue.put(UiEvent("prescan", fake))
    app._poll()                                              # old code: KeyError here, loop dead
    assert "KeyError" not in app.txt_log.cfg.get("text", "") and app.ctl.is_running()
    assert "cần bổ sung 1" in app.lbl_stage.cfg["text"] and "Đang quét" not in app.lbl_stage.cfg["text"]
    assert app.root.scheduled and app.root.scheduled[-1][0] == 100   # re-armed
    app.ctl._queue.put(UiEvent("row", (0, "reading", "")))
    app.ctl._queue.put(UiEvent("row", (0, "completed", "")))
    app.ctl._queue.put(UiEvent("done", BatchSummary(total=1, completed=1)))
    _polls(app)
    assert not app.ctl.is_running() and app._done_handled
    assert app.ctl.progress.percent == 100.0 and app.lbl_progress.cfg["text"] == "Đã xử lý: 1 / 1 — 100%"


# ------------------------------------------------------------------ #2 / #15 malformed display event
def test_malformed_display_event_is_logged_and_done_still_processed(monkeypatch, tmp_path):
    gui, app, clock, Top = _make_view(monkeypatch, tmp_path, 1)
    logged = []
    app.log = lambda m: logged.append(m)
    app.ctl._queue.put(UiEvent("row", (0, "reading", "")))
    app.ctl._queue.put(UiEvent("ollama", "not-a-pair"))          # malformed non-critical display event
    app.ctl._queue.put(UiEvent("row", (0, "completed", "")))
    app.ctl._queue.put(UiEvent("done", BatchSummary(total=1, completed=1)))
    clock.advance(12)
    _polls(app)
    assert any(m.startswith("GUI_EVENT_ERROR event=ollama") for m in logged + app.ctl.log_lines)
    assert app.ctl.progress.percent == 100.0 and app.pb["value"] == 100.0 and app.lbl_percent.cfg["text"] == "100%"
    assert not app.ctl.is_running() and app.ctl.finished_at is not None
    assert app.lbl_elapsed.cfg["text"] == "Tổng thời gian: 00:12"
    clock.advance(600)
    app._tick()                                                   # timer no longer updates after finalization
    assert app.lbl_elapsed.cfg["text"] == "Tổng thời gian: 00:12"
    assert app.btn_start.cfg["state"] == "normal" and app.btn_stop.cfg["state"] == "disabled"


# ------------------------------------------------------------------ #16 one file
def test_single_file_stage_events_then_done_reach_100_and_final_row_status(monkeypatch, tmp_path):
    gui, app, clock, Top = _make_view(monkeypatch, tmp_path, 1)
    seen = []
    for ev in (UiEvent("row", (0, "reading", "")), UiEvent("row", (0, "extracting", ""))):
        app.ctl._queue.put(ev); app._poll(); seen.append(app.ctl.progress.percent)
    assert seen == [STAGE_WEIGHTS["reading"] * 100, STAGE_WEIGHTS["extracting"] * 100]
    assert _row_status(app, 0).endswith("Đang trích xuất nội dung") or "Đang" in _row_status(app, 0)
    app.ctl._queue.put(UiEvent("row", (0, "completed", ""))); app._poll(); seen.append(app.ctl.progress.percent)
    app.ctl._queue.put(UiEvent("done", BatchSummary(total=1, completed=1))); app._poll(); seen.append(app.ctl.progress.percent)
    assert seen == sorted(seen) and seen[-1] == 100.0
    assert app.lbl_progress.cfg["text"] == "Đã xử lý: 1 / 1 — 100%"
    assert _row_status(app, 0) == "✓ Hoàn thành"
    assert app.lbl_eta.cfg["text"] == "Đã hoàn thành" and app.lbl_stage.cfg["text"] == "Hoàn thành."
    assert len(Top.created) == 1                                   # completion dialog shown once, after cleanup
    assert app.ctl.elapsed_text().startswith("Tổng thời gian:")


# ------------------------------------------------------------------ #17 five files, monotonic
def test_five_files_progress_monotonic_and_stale_stage_never_lowers(tmp_path):
    ctl, clock = _ctl_with_rows(tmp_path, 5)
    _begin(ctl, clock)
    history = []
    for i in range(5):
        for stage in ("reading", "extracting", "writing_excel"):
            ctl.apply_event(UiEvent("row", (i, stage, ""))); history.append(ctl.progress.percent)
        ctl.apply_event(UiEvent("row", (i, "completed", ""))); history.append(ctl.progress.percent)
        assert ctl.progress.percent == pytest.approx((i + 1) * 20.0)
        if i < 4:
            ctl.apply_event(UiEvent("row", (i, "reading", "")))     # stale event for a finished file
            history.append(ctl.progress.percent)
    ctl.apply_event(UiEvent("done", BatchSummary(total=5, completed=5))); history.append(ctl.progress.percent)
    assert history == sorted(history) and history[-1] == 100.0 and ctl.progress.text == "Đã xử lý: 5 / 5 — 100%"
    # after finalization a late stage event cannot re-open the batch
    ctl.apply_event(UiEvent("row", (2, "reading", "")))
    assert ctl.progress.percent == 100.0 and not ctl.is_running() and ctl.rows[2].is_final and ctl.stale_events == 5


# ------------------------------------------------------------------ #18 dead worker
class _DeadWorker:
    def __init__(self, summary=None, thread=True):
        self.summary = summary or BatchSummary()
        self.results = []
        self._thread = object() if thread else None
    def is_running(self): return False


def test_dead_worker_without_done_is_reconciled(monkeypatch, tmp_path):
    gui, app, clock, Top = _make_view(monkeypatch, tmp_path, 2)
    # (a) worker finished normally (summary complete) but its done event was lost
    app.ctl.processor = _DeadWorker(BatchSummary(total=2, completed=2))
    app.ctl._queue.put(UiEvent("row", (0, "completed", "")))
    app.ctl._queue.put(UiEvent("row", (1, "completed", "")))
    _polls(app)
    assert not app.ctl.is_running() and app._done_handled and app.ctl.progress.percent == 100.0
    assert any("WORKER_RECONCILE" in ln for ln in app.ctl.log_lines) and not app.ctl.worker_failure
    assert app.btn_start.cfg["state"] == "normal"
    # (b) worker crashed mid-way: failure surfaced, GUI usable again, not hidden
    gui, app, clock, Top = _make_view(monkeypatch, tmp_path, 2)
    app.ctl.processor = _DeadWorker(BatchSummary(total=0))
    app.ctl._queue.put(UiEvent("row", (0, "reading", "")))
    _polls(app)
    assert not app.ctl.is_running() and app.ctl.worker_failure.startswith("Lỗi tiến trình xử lý")
    assert app.lbl_stage.cfg["text"].startswith("Lỗi tiến trình xử lý")
    assert app.ctl.rows[0].stage == "error" and any("WORKER_DIED" in ln for ln in app.ctl.log_lines)
    assert app.btn_start.cfg["state"] == "normal" and app.btn_stop.cfg["state"] == "disabled"


def test_reconcile_is_noop_while_worker_alive_or_queue_pending(tmp_path):
    ctl, clock = _ctl_with_rows(tmp_path, 1)
    _begin(ctl, clock)
    class Alive(_DeadWorker):
        def is_running(self): return True
    ctl.processor = Alive()
    assert ctl.reconcile() == "" and ctl.is_running()
    ctl.processor = _DeadWorker(BatchSummary(total=1, completed=1))
    ctl._queue.put(UiEvent("log", "pending"))
    assert ctl.reconcile() == "" and ctl.is_running()             # queue must be drained first
    ctl.pump()
    assert ctl.reconcile().startswith("WORKER_RECONCILE") and not ctl.is_running()


# ------------------------------------------------------------------ #19 duplicate done
def test_duplicate_done_is_idempotent(monkeypatch, tmp_path):
    gui, app, clock, Top = _make_view(monkeypatch, tmp_path, 1)
    app.ctl._queue.put(UiEvent("row", (0, "completed", "")))
    app.ctl._queue.put(UiEvent("done", BatchSummary(total=1, completed=1)))
    _polls(app)
    finished_at, summary = app.ctl.finished_at, app.ctl.summary
    clock.advance(30)
    app.ctl._queue.put(UiEvent("done", BatchSummary(total=1, completed=0, failed=1)))
    app.ctl._queue.put(UiEvent("done", BatchSummary(total=1, completed=1)))
    _polls(app)
    assert app.ctl.finished_at == finished_at and app.ctl.summary is summary and app.ctl.done_events == 3
    assert len(Top.created) == 1 and app.ctl.progress.percent == 100.0 and not app.ctl.is_running()
    assert app.lbl_elapsed.cfg["text"] == "Tổng thời gian: 00:00"


def test_new_batch_resets_progress_and_done_flag(tmp_path):
    ctl, clock = _ctl_with_rows(tmp_path, 2)
    _begin(ctl, clock)
    ctl.apply_event(UiEvent("row", (0, "completed", "")))
    ctl.apply_event(UiEvent("done", BatchSummary(total=2, completed=1, stopped=True)))
    assert ctl.progress.percent == 50.0 and ctl.done_events == 1
    _begin(ctl, clock)
    ctl.progress = Progress(total=2)
    assert ctl.progress.percent == 0.0 and ctl.is_running()
