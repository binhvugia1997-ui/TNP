"""PROMPT-002 GUI regression suite (headless: the real view is built on a recording fake tkinter)."""
import importlib
import inspect
import re
import sys
import threading
import types
from pathlib import Path
from types import SimpleNamespace

import pytest

import app
from app.batch_processor import BatchSummary
from app.config import AppConfig
from app.gui_controller import SCAN_FILTERS_VI, UiEvent

from tests.test_gui_progress_ollama import FakeClock, _begin

SRC = Path(app.__file__).with_name("gui.py").read_text(encoding="utf-8")


# ------------------------------------------------------------------ recording fake tkinter
def _install_fake_tk(monkeypatch):
    registry = {"widgets": [], "toplevels": []}

    class Widget:
        def __init__(self, master=None, *a, **k):
            self.master = master
            self.k = dict(k)
            self.cfg = dict(k)
            self.grid_info_ = {}
            self.pack_info_ = {}
            self.bindings = {}
            self.col_weights, self.row_weights = {}, {}
            registry["widgets"].append(self)

        def __getattr__(self, name):
            if name.startswith("__"):
                raise AttributeError(name)
            fn = lambda *a, **k: None  # noqa: E731
            self.__dict__[name] = fn                      # stable identity (tree.yscrollcommand == vsb.set)
            return fn

        def __setitem__(self, k, v): self.cfg[k] = v
        def __getitem__(self, k): return self.cfg.get(k, "")
        def configure(self, *a, **k): self.cfg.update(k)
        config = configure
        def grid(self, **k): self.grid_info_ = k
        def pack(self, **k): self.pack_info_ = k
        def grid_forget(self): self.grid_info_ = {}
        def columnconfigure(self, i, **k): self.col_weights[i] = k.get("weight", 0)
        def rowconfigure(self, i, **k): self.row_weights[i] = k.get("weight", 0)
        def bind(self, seq, fn=None, add=None): self.bindings[seq] = fn
        bind_all = bind
        def get_children(self): return []
        def insert(self, *a, **k): return f"I{len(self.cfg)}"
        def selection(self): return ()
        def get(self): return self.k.get("value", "")
        def winfo_reqheight(self): return 900
        def winfo_screenwidth(self): return 1920
        def winfo_screenheight(self): return 1080
        def winfo_width(self): return 1400
        def winfo_height(self): return 850
        def state(self, *a): return "normal"

    class Tree(Widget):
        def __init__(self, *a, **k):
            super().__init__(*a, **k)
            self.items, self._n, self.sel, self.tags = {}, 0, (), {}
            self.columns_, self.headings_ = {}, {}
        def insert(self, *a, **k): self._n += 1; iid = f"I{self._n}"; self.items[iid] = k; return iid
        def item(self, iid, **k): self.items[iid].update(k)
        def get_children(self): return list(self.items)
        def delete(self, *iids): [self.items.pop(i, None) for i in iids]
        def selection(self): return self.sel
        def selection_set(self, *iids): self.sel = tuple(iids)
        def tag_configure(self, tag, **k): self.tags[tag] = k
        def column(self, key, **k): self.columns_[key] = k
        def heading(self, key, **k): self.headings_[key] = k

    class Text(Widget):
        def __init__(self, *a, **k): super().__init__(*a, **k); self.lines = []
        def insert(self, _idx, text, *tags): self.lines.append(text)
        def delete(self, *a): self.lines.clear()

    class Var(Widget):
        def __init__(self, value="", **k):
            self.v = value; self.cfg = {}; self.traces = []
        def set(self, v):
            self.v = v
            for fn in self.traces:
                fn()
        def get(self): return self.v
        def trace_add(self, _mode, fn): self.traces.append(fn)

    class Root(Widget):
        def __init__(self, *a, **k):
            super().__init__(None)
            self.scheduled, self.geometry_, self.minsize_, self.title_ = [], "", (), ""
        def after(self, ms, fn=None, *a): self.scheduled.append((ms, fn))
        def geometry(self, g=None): self.geometry_ = g
        def minsize(self, w, h): self.minsize_ = (w, h)
        def title(self, t): self.title_ = t

    class Top(Widget):
        def __init__(self, *a, **k): super().__init__(*a, **k); registry["toplevels"].append(self)

    class Style(Widget):
        def __init__(self, *a, **k): super().__init__(None); self.styles, self.theme = {}, ""
        def theme_names(self): return ("clam", "alt", "default", "vista")
        def theme_use(self, name=None): self.theme = name
        def configure(self, name, **k): self.styles.setdefault(name, {}).update(k)
        def map(self, name, **k): self.styles.setdefault(name, {}).update({"map": k})

    tkmod = types.ModuleType("tkinter")
    tkmod.Text, tkmod.Toplevel, tkmod.Menu, tkmod.Canvas = Text, Top, Widget, Widget
    tkmod.Tk, tkmod.StringVar, tkmod.BooleanVar, tkmod.TclError = Root, Var, Var, Exception
    ttkmod = types.ModuleType("tkinter.ttk")
    for n in ("Label", "LabelFrame", "Frame", "Entry", "Button", "Combobox", "Checkbutton", "Scrollbar", "Progressbar",
              "Radiobutton", "Spinbox", "Notebook", "PanedWindow", "Separator"):
        setattr(ttkmod, n, type(n, (Widget,), {}))
    ttkmod.Treeview, ttkmod.Style = Tree, Style
    fd, mb = types.ModuleType("tkinter.filedialog"), types.ModuleType("tkinter.messagebox")
    for m in (fd, mb):
        m.__getattr__ = lambda name: (lambda *a, **k: True)
    tkmod.ttk, tkmod.filedialog, tkmod.messagebox = ttkmod, fd, mb
    for name, mod in (("tkinter", tkmod), ("tkinter.ttk", ttkmod), ("tkinter.filedialog", fd), ("tkinter.messagebox", mb)):
        monkeypatch.setitem(sys.modules, name, mod)
    monkeypatch.delitem(sys.modules, "app.gui", raising=False)
    registry["classes"] = {"Widget": Widget, "Tree": Tree, "Scrollbar": ttkmod.Scrollbar, "Notebook": ttkmod.Notebook,
                           "Entry": ttkmod.Entry, "Button": ttkmod.Button, "Radiobutton": ttkmod.Radiobutton,
                           "Label": ttkmod.Label}
    return registry


def _make_app(monkeypatch, tmp_path, n_files=3, names=None, cfg_extra=None):
    reg = _install_fake_tk(monkeypatch)
    gui = importlib.import_module("app.gui")
    folder = tmp_path / "reports"
    folder.mkdir(exist_ok=True)
    names = names or [f"26092{i:04d}-VOC_report{i}.pptx" for i in range(n_files)]
    for n in names:
        (folder / n).write_bytes(b"x")
    cfg = AppConfig(last_report_folder=str(folder), period_mode="all", extra=dict(cfg_extra or {}))
    app_ = gui.ReportExtractorApp(cfg)
    for t in threading.enumerate():
        if t.name == "prescan":
            t.join(timeout=10)
    app_.ctl.pump()
    app_._render_scan()
    return gui, app_, reg, folder


def _running(app_):
    clock = FakeClock()
    app_.ctl._clock = clock
    _begin(app_.ctl, clock)
    app_._done_handled = False
    app_._render_rows()
    app_._set_running(True)
    return clock


def _scan_tree_names(app_):
    return [v["values"][4] for v in app_.scan_tree.items.values()]


# ------------------------------------------------------------------ 1 version
def test_canonical_version_prompt002():
    assert app.__version__ == "1.2.0" and app.BUILD_ID == "011" and app.BUILD_LABEL == "Build 011"
    assert app.APP_TITLE == "Report Extractor v1.2.0" and app.VERSION_LINE == "version=1.2.0 build=011"
    assert "1.2.0" not in SRC and "Build 011" not in SRC.split('"""', 2)[2]   # header imports the constants


def test_window_title_and_header_use_constants(monkeypatch, tmp_path):
    gui, a, reg, _ = _make_app(monkeypatch, tmp_path)
    assert a.root.title_.startswith("Report Extractor v1.2.0")
    assert a.lbl_version.cfg["text"] == "v1.2.0" and a.lbl_build.cfg["text"] == "Build 011"   # numeric build only
    assert a.root.geometry_ == "1400x850" and a.root.minsize_ == (880, 540)


# ------------------------------------------------------------------ 2 tabs, 3 source, 4 period
def test_main_tabs_exist(monkeypatch, tmp_path):
    gui, a, reg, _ = _make_app(monkeypatch, tmp_path)
    nb = a.nb
    assert isinstance(nb, reg["classes"]["Notebook"])
    assert a.tab_run.master is nb and a.tab_cfg.master is nb
    tab_texts = re.findall(r'self\.nb\.add\(self\.tab_\w+, text="([^"]+)"\)', SRC)
    assert [t.strip(" ▶⚙") for t in tab_texts] == ["Xử lý báo cáo", "Cấu hình & Ollama"]


def test_source_controls_exist(monkeypatch, tmp_path):
    gui, a, reg, folder = _make_app(monkeypatch, tmp_path)
    assert a.var_folder.get() == str(folder) and hasattr(a, "var_template") and hasattr(a, "var_output")
    entries = [w for w in reg["widgets"] if isinstance(w, reg["classes"]["Entry"])]
    assert all(e.grid_info_.get("sticky") == "ew" for e in entries if e.k.get("textvariable") in (a.var_folder, a.var_template, a.var_output))
    assert a.lbl_found.cfg["text"] == "Đã tìm thấy 3 báo cáo" and a.lbl_found.k.get("style") == "Secondary.TLabel"
    assert a.btn_scan.k["text"] == "Quét"


def test_period_controls_exist(monkeypatch, tmp_path):
    gui, a, reg, _ = _make_app(monkeypatch, tmp_path)
    assert [rb.k["text"] for rb in a._period_radios] == ["Tự động theo file Excel", "Chọn tháng", "Khoảng thời gian", "Tất cả"]
    assert a.lbl_period.cfg["text"] == "Đang áp dụng: Tất cả"
    a.var_period_mode.set("month")
    a._on_period_changed()
    assert a.frm_month.grid_info_ and not a.frm_range.grid_info_
    assert a.lbl_period.cfg["text"].startswith("Đang áp dụng: Tháng")
    a.var_period_mode.set("range")
    a._on_period_changed()
    assert a.frm_range.grid_info_ and not a.frm_month.grid_info_
    assert a.lbl_period_warn.k.get("style") == "Warning.TLabel"


# ------------------------------------------------------------------ 5, 6 scrollbars; 24 grid weights
def test_tables_have_both_scrollbars(monkeypatch, tmp_path):
    gui, a, reg, _ = _make_app(monkeypatch, tmp_path)
    sb = reg["classes"]["Scrollbar"]
    for tree, vsb, hsb in ((a.scan_tree, a.scan_vsb, a.scan_hsb), (a.tree, a.result_vsb, a.result_hsb)):
        assert isinstance(vsb, sb) and vsb.k["orient"] == "vertical" and vsb.grid_info_["sticky"] == "ns"
        assert isinstance(hsb, sb) and hsb.k["orient"] == "horizontal" and hsb.grid_info_["sticky"] == "ew"
        assert tree.cfg["yscrollcommand"] == vsb.set and tree.cfg["xscrollcommand"] == hsb.set
        assert tree.grid_info_["sticky"] == "nsew"


def test_resize_grid_weights_make_tables_expandable(monkeypatch, tmp_path):
    gui, a, reg, _ = _make_app(monkeypatch, tmp_path)
    assert a.root.col_weights == {0: 1} and a.root.row_weights.get(1) == 1          # notebook grows
    assert a.tab_run.row_weights.get(0) == 1 and a.tab_run.col_weights.get(0) == 1
    assert a.page.col_weights.get(0) == 1
    assert a.page.row_weights.get(2, 0) >= a.page.row_weights.get(4, 0) > 0        # file list, then results
    for tree in (a.scan_tree, a.tree):
        holder = tree.master
        assert holder.col_weights.get(0) == 1 and holder.row_weights.get(0) == 1
    assert a.scan_tree.columns_["file"]["stretch"] and not a.scan_tree.columns_["mgmt"]["stretch"]
    assert a.tree.columns_["note"]["stretch"] and a.tree.columns_["file"]["stretch"]
    assert a.txt_log.grid_info_["sticky"] == "nsew" and a.cfg_page.row_weights.get(4) == 1 and a.cfg_page.row_weights.get(2) in (None, 0)   # only the log row grows


# ------------------------------------------------------------------ 7–11 running-state controls
def test_start_enabled_stop_disabled_when_idle(monkeypatch, tmp_path):
    gui, a, reg, _ = _make_app(monkeypatch, tmp_path)
    assert a.btn_start.cfg.get("state", "normal") == "normal" and a.btn_stop.k["state"] == "disabled"
    assert a.btn_start.k["style"] == "Primary.TButton" and a.btn_stop.k["style"] == "Danger.TButton"


def test_running_state_locks_configuration_and_enables_stop(monkeypatch, tmp_path):
    gui, a, reg, _ = _make_app(monkeypatch, tmp_path)
    _running(a)
    assert a.btn_start.cfg["state"] == "disabled" and a.btn_stop.cfg["state"] == "normal"
    locked = [w.cfg.get("state") for w in a._config_widgets]
    assert locked and all(s == "disabled" for s in locked)
    assert a.btn_rescan.cfg["state"] == "disabled" and a.cb_model.cfg["state"] == "disabled"
    assert a.e_search.cfg.get("state", "normal") == "normal"                     # GUI-only search stays usable
    a._rescan_if_idle()                                                          # F5 ignored while running
    assert a.ctl.is_running()


def test_done_restores_controls(monkeypatch, tmp_path):
    gui, a, reg, _ = _make_app(monkeypatch, tmp_path, 1)
    _running(a)
    a.ctl._queue.put(UiEvent("row", (0, "completed", "")))
    a.ctl._queue.put(UiEvent("done", BatchSummary(total=1, completed=1)))
    for _ in range(3):
        a._poll()
    assert not a.ctl.is_running() and a._done_handled
    assert a.btn_start.cfg["state"] == "normal" and a.btn_stop.cfg["state"] == "disabled"
    assert all(w.cfg.get("state") == "normal" for w in a._config_widgets)
    assert len(reg["toplevels"]) == 1                                            # summary dialog once, no Excel auto-open
    assert "os.startfile" not in SRC[SRC.index("def _on_done("):SRC.index("def _show_summary_dialog(")]


# ------------------------------------------------------------------ 12–14 progress from the controller only
def test_progress_renders_controller_percentage(monkeypatch, tmp_path):
    gui, a, reg, _ = _make_app(monkeypatch, tmp_path, 4)
    clock = _running(a)
    for i, stage in ((0, "reading"), (0, "completed"), (1, "extracting")):
        a.ctl._queue.put(UiEvent("row", (i, stage, "")))
    a._poll()
    p = a.ctl.progress
    assert a.pb["value"] == p.percent and a.lbl_percent.cfg["text"] == f"{p.percent_int}%"
    assert a.lbl_progress.cfg["text"] == p.text and p.text.startswith("Đang xử lý: 2 / 4")
    assert a.lbl_file.cfg["text"].endswith("report1.pptx")
    render = SRC[SRC.index("def _render_progress("):SRC.index("def _render_counts(")]
    assert "p.percent" in render and "/" not in render.replace("//", "").split("p = self.ctl.progress")[1].split("self.pb")[0]
    clock.advance(65)
    a._tick()
    assert a.lbl_elapsed.cfg["text"] == a.ctl.elapsed_text() == "Đã chạy: 01:05"


def test_hundred_percent_and_elapsed_freeze_after_done(monkeypatch, tmp_path):
    gui, a, reg, _ = _make_app(monkeypatch, tmp_path, 2)
    clock = _running(a)
    a.ctl._queue.put(UiEvent("row", (0, "completed", "")))
    a.ctl._queue.put(UiEvent("row", (1, "completed", "")))
    clock.advance(42)
    a.ctl._queue.put(UiEvent("done", BatchSummary(total=2, completed=2)))
    for _ in range(3):
        a._poll()
    assert a.pb["value"] == 100.0 and a.lbl_percent.cfg["text"] == "100%"
    assert a.lbl_progress.cfg["text"] == "Đã xử lý: 2 / 2 — 100%" and a.lbl_stage.cfg["text"] == "Hoàn thành."
    assert a.lbl_elapsed.cfg["text"] == "Tổng thời gian: 00:42" and a.lbl_eta.cfg["text"] == "Đã hoàn thành"
    clock.advance(500)
    a._tick()
    assert a.lbl_elapsed.cfg["text"] == "Tổng thời gian: 00:42"
    assert a.card_values["completed"].cfg["text"] == "2" and a.card_values["failed"].cfg["text"] == "0"


# ------------------------------------------------------------------ 15, 16 search
def test_search_filters_gui_rows_only_and_clear_restores(monkeypatch, tmp_path):
    names = ["260901001-VOC_alpha.pptx", "260901002-VOC_beta.pptx", "260901003-VOC_gamma.pptx"]
    gui, a, reg, folder = _make_app(monkeypatch, tmp_path, names=names)
    a.var_scan_filter.set(SCAN_FILTERS_VI[1])
    a._render_scan()
    assert sorted(_scan_tree_names(a)) == sorted(names) and a.lbl_list_count.cfg["text"] == "3 file"
    before_rows = [(r.key, r.excluded) for r in a.ctl.scan_rows()]
    scan_calls = []
    monkeypatch.setattr(a.ctl, "scan", lambda *a_, **k: scan_calls.append(1))
    monkeypatch.setattr(a.ctl, "scan_async", lambda *a_, **k: scan_calls.append(1))
    a.set_search("260901002")
    assert _scan_tree_names(a) == ["260901002-VOC_beta.pptx"] and a.lbl_list_count.cfg["text"] == "1 / 3 file"
    a.set_search("GAMMA")
    assert _scan_tree_names(a) == ["260901003-VOC_gamma.pptx"]
    a.set_search("")
    assert sorted(_scan_tree_names(a)) == sorted(names) and a.lbl_list_count.cfg["text"] == "3 file"
    assert scan_calls == [] and [(r.key, r.excluded) for r in a.ctl.scan_rows()] == before_rows   # GUI-only
    assert all((folder / n).exists() for n in names)
    assert a.var_search.get() == gui.SEARCH_PLACEHOLDER and a.search_text() == ""


# ------------------------------------------------------------------ 17, 18 manual remove / restore
def test_manual_remove_keeps_source_file_and_restore_returns_it(monkeypatch, tmp_path):
    names = ["260901001-VOC_alpha.pptx", "260901002-VOC_beta.pptx"]
    gui, a, reg, folder = _make_app(monkeypatch, tmp_path, names=names)
    a.var_scan_filter.set(SCAN_FILTERS_VI[1])
    a._render_scan()
    iid = next(i for i, v in a.scan_tree.items.items() if v["values"][4] == "260901002-VOC_beta.pptx")
    a.scan_tree.selection_set(iid)
    a.exclude_selected()
    row = next(r for r in a.ctl.scan_rows() if r.path.name == "260901002-VOC_beta.pptx")
    assert row.excluded and (folder / "260901002-VOC_beta.pptx").exists()
    assert "Đã loại thủ công" in a.scan_tree.items[next(i for i, v in a.scan_tree.items.items() if v["values"][4] == row.path.name)]["values"][5]
    assert "excluded" in a.scan_tree.items[next(i for i, v in a.scan_tree.items.items() if v["values"][4] == row.path.name)]["tags"]
    iid2 = next(i for i, v in a.scan_tree.items.items() if v["values"][4] == row.path.name)
    a.scan_tree.selection_set(iid2)
    a.restore_selected()
    assert not next(r for r in a.ctl.scan_rows() if r.path.name == row.path.name).excluded
    assert SRC.count("self.ctl.exclude(") == 1 and "os.remove" not in SRC and "send2trash" not in SRC


# ------------------------------------------------------------------ 19 Ollama check asynchronous
def test_ollama_check_is_asynchronous(monkeypatch, tmp_path):
    gui, a, reg, _ = _make_app(monkeypatch, tmp_path)
    called = []
    monkeypatch.setattr(a.ctl, "check_ollama_async", lambda **k: called.append("async"))
    monkeypatch.setattr(a.ctl, "check_ollama", lambda *a_, **k: called.append("SYNC"))
    a.var_host.set("192.168.1.50")
    a.var_port.set("11434")
    a.var_model.set("qwen3:4b")
    a.check_ollama()
    assert called == ["async"] and a.lbl_conn.cfg["text"].startswith("● Đang kiểm tra")
    a.ctl._queue.put(UiEvent("ollama", (True, "Ollama: Sẵn sàng — qwen3:4b")))
    a.ctl.ollama_ok = True
    a._poll()
    assert a.lbl_conn.cfg["style"] == "Success.TLabel" and a.lbl_ai_run.cfg["text"] == "● Ollama: qwen3:4b — Sẵn sàng"
    assert "self.ctl.check_ollama_async(" in SRC[SRC.index("def check_ollama("):SRC.index("def refresh_models(")]


# ------------------------------------------------------------------ 20 log viewer
def test_log_viewer_receives_log_events_and_clear_only_clears_widget(monkeypatch, tmp_path):
    gui, a, reg, _ = _make_app(monkeypatch, tmp_path)
    a.ctl._queue.put(UiEvent("log", "dòng nhật ký 1"))
    a._poll()
    assert "dòng nhật ký 1\n" in a.txt_log.lines
    log_file = tmp_path / "app.log"
    log_file.write_text("keep", encoding="utf-8")
    monkeypatch.setattr(a.ctl, "log_file", lambda: log_file)
    a.clear_log_view()
    assert a.txt_log.lines == [] and log_file.read_text(encoding="utf-8") == "keep"
    assert a.txt_log.master.master.master is not a.tab_run                       # lives on the configuration tab


# ------------------------------------------------------------------ 21, 22 lifecycle protections
def test_malformed_display_event_does_not_kill_polling(monkeypatch, tmp_path):
    gui, a, reg, _ = _make_app(monkeypatch, tmp_path, 1)
    _running(a)
    a.ctl._queue.put(UiEvent("ollama", "not-a-pair"))
    a.ctl._queue.put(UiEvent("row", ("bad-index",)))
    a._poll()
    assert any("GUI_EVENT_ERROR event=ollama" in ln for ln in a.txt_log.lines)
    assert a.root.scheduled and a.root.scheduled[-1][0] == 100 and a.ctl.is_running()
    a.ctl._queue.put(UiEvent("row", (0, "completed", "")))
    a.ctl._queue.put(UiEvent("done", BatchSummary(total=1, completed=1)))
    for _ in range(3):
        a._poll()
    assert not a.ctl.is_running() and a.lbl_percent.cfg["text"] == "100%"


def test_dead_worker_reconciliation_still_works(monkeypatch, tmp_path):
    gui, a, reg, _ = _make_app(monkeypatch, tmp_path, 2)
    _running(a)
    a.ctl.processor = SimpleNamespace(summary=BatchSummary(total=2, completed=2), results=[], _thread=object(),
                                      is_running=lambda: False)
    a.ctl._queue.put(UiEvent("row", (0, "completed", "")))
    a.ctl._queue.put(UiEvent("row", (1, "completed", "")))
    for _ in range(3):
        a._poll()
    assert not a.ctl.is_running() and a._done_handled and a.btn_start.cfg["state"] == "normal"
    assert any("WORKER_RECONCILE" in ln for ln in a.ctl.log_lines)


# ------------------------------------------------------------------ 23 diagnostics
def test_double_click_result_opens_diagnostics(monkeypatch, tmp_path):
    gui, a, reg, _ = _make_app(monkeypatch, tmp_path, 1)
    a._render_rows()
    diag = {"Tên file": "r.pptx", "Trạng thái": "Hoàn thành", "Vùng QPN": "dominant_picture", "Loại khỏi QPN": "title: 1. HIỆN TRẠNG",
            "Mục loại khỏi cải tiến": "S7 inspection", "Ảnh Sau cải tiến": "S5#37", "Quyết định từng ảnh": "…",
            "Trường đã điền": "improvement", "Classifier": "qwen", "Confidence": 0.9, "Review reasons": "[]", "_hidden": "x"}
    monkeypatch.setattr(a.ctl, "diagnostics_for", lambda i: diag)
    a.tree.selection_set(a.row_items[0])
    a._on_row_double_click()
    assert len(reg["toplevels"]) == 1
    text = "".join(w.lines[0] if False else "".join(w.lines) for w in reg["widgets"] if isinstance(w, type(a.txt_log)) and w is not a.txt_log)
    for key in ("Vùng QPN", "Loại khỏi QPN", "Mục loại khỏi cải tiến", "Ảnh Sau cải tiến", "Quyết định từng ảnh",
                "Trường đã điền", "Classifier", "Confidence", "Review reasons"):
        assert f"{key}: " in text
    assert "_hidden" not in text


# ------------------------------------------------------------------ 25, 26 architecture guards
def test_no_obsolete_master_not_found_access():
    assert "c['master_not_found']" not in SRC and 'counts["master_not_found"]' not in SRC
    assert re.search(r"\[\s*['\"]master_not_found['\"]\s*\]", SRC) is None
    counts_fn = SRC[SRC.index("def summary_counts("):SRC.index("def ollama_indicator_text(")]
    assert "getattr(summary, key, 0)" in counts_fn


def test_no_extraction_or_classifier_code_in_gui():
    for forbidden in ("from .extractor", "from .classifier", "from .qpn_region", "from .qpn_renderer", "from .improvement_pictures",
                      "from .excel_writer", "from .pptx_parser", "BatchProcessor(", "heuristic_classify", "extract_record("):
        assert forbidden not in SRC, forbidden
    import app.gui_controller as gc
    assert "BatchProcessor" in inspect.getsource(gc)                      # processing stays behind the controller


def test_summary_cards_and_status_tags(monkeypatch, tmp_path):
    gui, a, reg, _ = _make_app(monkeypatch, tmp_path)
    assert [t for t in a.card_values] == ["completed", "needs_review", "failed", "skipped"]
    for tag in ("candidate", "excluded", "outside", "duplicate", "fast_skip", "complete", "invalid"):
        assert tag in a.scan_tree.tags
    for tag in ("completed", "needs_review", "error", "skipped", "working"):
        assert tag in a.tree.tags
    assert gui.summary_counts(SimpleNamespace(completed="3", failed=None)) == {
        "total": 0, "completed": 3, "needs_review": 0, "failed": 0, "skipped": 0, "not_written": 0}


def test_window_geometry_persistence_is_validated(monkeypatch, tmp_path):
    gui, a, reg, _ = _make_app(monkeypatch, tmp_path, cfg_extra={"window_width": 1250, "window_height": 760})
    assert a.root.geometry_ == "1250x760"
    gui2, b, reg2, _ = _make_app(monkeypatch, tmp_path, cfg_extra={"window_width": 99999, "window_height": 10})
    assert b.root.geometry_ == "1400x850"                                    # invalid persisted size ignored
    assert gui.valid_geometry(880, 540, 1920, 1080) and not gui.valid_geometry(800, 540, 1920, 1080)
    b._remember_window_geometry()
    assert (b.ctl.cfg.extra["window_width"], b.ctl.cfg.extra["window_height"]) == (1400, 850)


@pytest.mark.parametrize("ok,server,expect", [(None, "http://h:1", "Chưa kiểm tra"), (True, "http://h:1", "Sẵn sàng"),
                                               (False, "http://h:1", "Không kết nối được"), (None, "", "heuristic fallback")])
def test_ollama_indicator_text(ok, server, expect):
    from app.gui import ollama_indicator_text
    ctl = SimpleNamespace(server=server, model="qwen3:4b", ollama_ok=ok)
    assert expect in ollama_indicator_text(ctl) and ollama_indicator_text(ctl).startswith("● Ollama:")


def test_saved_period_mode_survives_view_load(monkeypatch, tmp_path):
    """Regression: traces on the template/output entries used to push the period while the period widgets still
    held the default 'auto', silently resetting a saved 'all' / 'month' mode on every start."""
    gui, a, reg, _ = _make_app(monkeypatch, tmp_path)
    assert a.ctl.cfg.period_mode == "all" and a.ctl.period_mode == "all" and a.var_period_mode.get() == "all"
