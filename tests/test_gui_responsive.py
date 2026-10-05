"""PROMPT-007A: responsive GUI + scrollable settings tab (Build 008).

Real Tk is not available in the CI sandbox, so these tests drive the recording fake tkinter of
test_gui_redesign with a small recording Canvas/Toplevel and feed synthetic <Configure> events to the layout
handlers; geometry rules are additionally checked on the source (no fixed card heights, no pack/grid mixing).
"""
import importlib
import re
import sys
import threading
import types
from pathlib import Path

import app
from app.config import AppConfig
from tests.test_gui_redesign import _install_fake_tk

SRC = Path(app.__file__).with_name("gui.py").read_text(encoding="utf-8")
CFG_TAB = SRC[SRC.index("def _build_cfg_tab("):SRC.index("def _build_learning_tab(")]
LEARNING_TAB = SRC[SRC.index("def _build_learning_tab("):SRC.index("def _bind_shortcuts(")]
RUN_TAB = SRC[SRC.index("def _build_run_tab("):SRC.index("def _build_cfg_tab(")]


class Ev:
    def __init__(self, width, height, widget=None):
        self.width, self.height, self.widget = width, height, widget


def _app(monkeypatch, tmp_path, inner_req_height=900):
    reg = _install_fake_tk(monkeypatch)
    tkmod = sys.modules["tkinter"]
    Widget = reg["classes"]["Widget"]

    class Canvas(Widget):
        def __init__(self, *a, **k):
            super().__init__(*a, **k)
            self.windows, self.itemcfg, self.scrolls, self.bbox_ = {}, [], [], (0, 0, 700, 1500)

        def create_window(self, pos, window=None, anchor="nw"):
            wid = f"W{len(self.windows) + 1}"
            self.windows[wid] = window
            return wid

        def itemconfigure(self, wid, **k):
            self.itemcfg.append((wid, dict(k)))

        def bbox(self, _tag):
            return self.bbox_

        def yview_scroll(self, n, unit):
            self.scrolls.append(n)

    class Top(Widget):
        def __init__(self, *a, **k):
            super().__init__(*a, **k)
            reg["toplevels"].append(self)
            self.calls = []

        def resizable(self, w, h): self.calls.append(("resizable", w, h))
        def minsize(self, w, h): self.calls.append(("minsize", w, h))
        def geometry(self, g): self.calls.append(("geometry", g))

    tkmod.Canvas, tkmod.Toplevel = Canvas, Top
    Widget.winfo_reqheight = lambda self: inner_req_height
    gui = importlib.import_module("app.gui")
    folder = tmp_path / "reports"
    folder.mkdir(exist_ok=True)
    (folder / "260901001-VOC_a.pptx").write_bytes(b"x")
    a = gui.ReportExtractorApp(AppConfig(last_report_folder=str(folder), period_mode="all"))
    for t in threading.enumerate():
        if t.name == "prescan":
            t.join(timeout=10)
    a.ctl.pump()
    return gui, a, reg


def _children(reg, parent):
    return [w for w in reg["widgets"] if getattr(w, "master", None) is parent]


def _descendants(reg, parent):
    out, stack = [], [parent]
    while stack:
        p = stack.pop()
        for w in _children(reg, p):
            out.append(w)
            stack.append(w)
    return out


def _page(a, name):
    return next(sp for sp in a._scroll_pages if sp["name"] == name)


def _last_item(canvas, win):
    return next(k for w, k in reversed(canvas.itemcfg) if w == win)


# ------------------------------------------------------------------ 1-4 cards render their complete body
def test_update_card_body_exists_and_is_mapped(monkeypatch, tmp_path):
    gui, a, reg = _app(monkeypatch, tmp_path)
    body = a.cfg_cards["update"]
    assert body.grid_info_.get("sticky") == "nsew" and body.grid_info_.get("row") == 1
    kids = _children(reg, body)
    assert a.lbl_cur_version in kids and a.lbl_update in kids
    assert all(k.grid_info_ or k.pack_info_ for k in kids)                      # every child is managed
    assert a.btn_check_update.pack_info_ and a.btn_install_update.pack_info_


def test_learning_cards_are_on_their_own_tab_and_mapped(monkeypatch, tmp_path):
    gui, a, reg = _app(monkeypatch, tmp_path)
    assert a.learning_canvas is _page(a, "learning")["canvas"]
    assert "learning" not in a.cfg_cards
    data = a.learning_cards["data"]
    review = a.learning_cards["review"]
    assert data.grid_info_.get("sticky") == "nsew" and review.grid_info_.get("sticky") == "nsew"
    assert a.lbl_learning in _children(reg, data) and a.lbl_learning_content in _children(reg, data)
    assert a.lbl_review_summary in _children(reg, review) and a.lbl_review_content_summary in _children(reg, review)
    assert a.lbl_learning_model in _children(reg, a.learning_cards["models"])
    assert a.lbl_reapply_pending in _children(reg, a.learning_cards["apply"])
    assert all(k.grid_info_ or k.pack_info_ for card in a.learning_cards.values() for k in _children(reg, card))
    settings = _descendants(reg, a.cfg_page)
    assert all(w not in settings for w in (a.btn_review_content, a.btn_review_images, a.btn_train_images,
                                           a.btn_reapply_saved))


def test_all_learning_buttons_exist_only_on_learning_tab(monkeypatch, tmp_path):
    gui, a, reg = _app(monkeypatch, tmp_path)
    Button = reg["classes"]["Button"]
    texts = {w.cfg.get("text") for w in _descendants(reg, a.learning_page) if isinstance(w, Button)}
    assert texts == {"Kiểm tra nội dung cải tiến", "Kiểm tra hình ảnh cải tiến", "Cập nhật mô hình học",
                     "Mở thư mục dữ liệu học", "Xuất dữ liệu học", "Cập nhật Excel từ nhãn đã lưu"}
    assert a.btn_review_content.pack_info_ and a.btn_review_images.pack_info_ and a.btn_train_images.grid_info_


def test_cards_take_their_height_from_children(monkeypatch, tmp_path):
    gui, a, reg = _app(monkeypatch, tmp_path)
    for name in ("ollama", "options", "update", "log"):
        assert len(_children(reg, a.cfg_cards[name])) >= 1                      # a settings body with real content
    assert all(len(_children(reg, a.learning_cards[name])) >= 1 for name in a.learning_cards)
    card_src = SRC[SRC.index("def _card("):SRC.index("def _make_table(")]
    assert "height=" not in card_src and "propagate" not in card_src              # no fixed card heights
    assert "grid_propagate(False)" not in SRC and "pack_propagate(False)" not in SRC
    # settings page: only the log row is weighted (Build 007 weighted the update card -> header-only collapse)
    assert a.cfg_page.row_weights == {3: 1}
    assert a.cfg_cards["log"].row_weights.get(0) == 1 and a.txt_log.grid_info_["sticky"] == "nsew"


# ------------------------------------------------------------------ 5-7 canvas / inner frame geometry
def test_settings_inner_frame_may_exceed_viewport_height(monkeypatch, tmp_path):
    gui, a, reg = _app(monkeypatch, tmp_path, inner_req_height=1400)
    sp = _page(a, "cfg")
    assert sp["canvas"] is a.cfg_canvas and sp["canvas"].windows[sp["win"]] is a.cfg_page
    sp["canvas"].bindings["<Configure>"](Ev(700, 400))
    assert _last_item(sp["canvas"], sp["win"]) == {"width": 700, "height": 1400}     # taller than the viewport
    assert a.cfg_vsb.grid_info_["sticky"] == "ns" and a.cfg_canvas.cfg["yscrollcommand"] == a.cfg_vsb.set


def test_canvas_scrollregion_follows_content(monkeypatch, tmp_path):
    gui, a, reg = _app(monkeypatch, tmp_path)
    for name in ("cfg", "run", "learning"):
        sp = _page(a, name)
        sp["canvas"].bbox_ = (0, 0, 700, 2300)
        sp["inner"].bindings["<Configure>"](Ev(700, 2300))
        assert sp["canvas"].cfg["scrollregion"] == (0, 0, 700, 2300)
    inner_src = SRC[SRC.index("def _on_scroll_inner_configure("):SRC.index("def _on_scroll_canvas_configure(")]
    assert 'canvas.bbox("all")' in inner_src and "scrollregion=bbox" in inner_src


def test_inner_frame_width_follows_viewport_width(monkeypatch, tmp_path):
    gui, a, reg = _app(monkeypatch, tmp_path, inner_req_height=900)
    sp = _page(a, "cfg")
    for w in (640, 1024, 1900):
        sp["canvas"].bindings["<Configure>"](Ev(w, 700))
        assert _last_item(sp["canvas"], sp["win"])["width"] == w
    assert _last_item(sp["canvas"], sp["win"])["height"] == 900                 # content taller than 700 viewport
    sp["canvas"].bindings["<Configure>"](Ev(1900, 1200))
    assert _last_item(sp["canvas"], sp["win"])["height"] == 1200                # fills a tall viewport


def test_learning_page_resizes_and_pending_status_refreshes(monkeypatch, tmp_path):
    gui, a, reg = _app(monkeypatch, tmp_path, inner_req_height=1100)
    sp = _page(a, "learning")
    for width, height in ((520, 420), (980, 700), (1600, 1200)):
        sp["canvas"].bindings["<Configure>"](Ev(width, height))
        assert _last_item(sp["canvas"], sp["win"])["width"] == width
        assert a.lbl_learning.cfg["wraplength"] == max(gui.WRAP_MIN, width - 6 * gui.L)
    a.ctl.pending_content_reapply = ["row-a", "row-b"]
    a._render_learning()
    assert a.btn_reapply_saved.cfg["state"] == "normal"
    assert a.lbl_reapply_pending.cfg["text"].startswith("Đang chờ cập nhật Excel: 2 mục")
    a.ctl.pending_content_reapply = []
    a._render_learning()
    assert a.btn_reapply_saved.cfg["state"] == "disabled"
    assert a.lbl_reapply_pending.cfg["text"] == "Không có lần cập nhật Excel nào đang chờ."


# ------------------------------------------------------------------ 8-9 shrinking / enlarging
def test_reducing_window_height_keeps_every_control(monkeypatch, tmp_path):
    gui, a, reg = _app(monkeypatch, tmp_path, inner_req_height=1600)
    before_n = len(reg["widgets"])
    managed = {id(w): (dict(w.grid_info_), dict(w.pack_info_)) for w in _descendants(reg, a.cfg_page)}
    for h in (900, 600, 400, 250):
        _page(a, "cfg")["canvas"].bindings["<Configure>"](Ev(800, h))
        _page(a, "run")["canvas"].bindings["<Configure>"](Ev(800, h))
    assert len(reg["widgets"]) == before_n                                      # nothing destroyed / recreated
    assert {id(w): (dict(w.grid_info_), dict(w.pack_info_)) for w in _descendants(reg, a.cfg_page)} == managed
    assert _last_item(a.cfg_canvas, _page(a, "cfg")["win"])["height"] == 1600  # never clipped to 250
    assert a.root.minsize_ == gui.MIN_GEOMETRY and gui.MIN_GEOMETRY[1] <= 600   # scrolling stays the fallback


def test_enlarging_window_expands_entries_and_wrap_labels(monkeypatch, tmp_path):
    gui, a, reg = _app(monkeypatch, tmp_path)
    Entry = reg["classes"]["Entry"]
    ew = [w for w in reg["widgets"] if isinstance(w, Entry) and w.grid_info_.get("sticky") == "ew"]
    assert len(ew) >= 5                                                         # folder / Excel / output / server / update
    assert any(w.cfg.get("textvariable") is a.var_host for w in ew)
    assert any(w.cfg.get("textvariable") is a.var_update_path for w in ew)
    assert any(w.cfg.get("textvariable") is a.var_folder for w in ew)
    for body in a.cfg_cards.values():
        assert body.col_weights.get(0) == 1 or body.col_weights.get(1) == 1
    _page(a, "cfg")["canvas"].bindings["<Configure>"](Ev(1800, 1000))
    assert all(lbl.cfg["wraplength"] == 1800 - 6 * gui.L for lbl in a._wrap_labels)
    _page(a, "cfg")["canvas"].bindings["<Configure>"](Ev(500, 1000))
    assert all(lbl.cfg["wraplength"] == max(gui.WRAP_MIN, 500 - 6 * gui.L) for lbl in a._wrap_labels)
    Button = reg["classes"]["Button"]
    stretched = [w for w in reg["widgets"] if isinstance(w, Button) and w.grid_info_.get("sticky") in ("ew", "nsew")]
    assert stretched == []                                                      # buttons keep natural widths


# ------------------------------------------------------------------ 10-12 report list
def test_report_treeview_expands_with_window(monkeypatch, tmp_path):
    gui, a, reg = _app(monkeypatch, tmp_path)
    assert a.scan_tree.grid_info_["sticky"] == "nsew"
    holder = a.scan_tree.master
    assert holder.col_weights.get(0) == 1 and holder.row_weights.get(0) == 1
    assert holder.master.row_weights.get(3) == 1 and a.page.row_weights.get(2) == 3   # card row grows most
    assert a.scan_tree.columns_["file"]["stretch"] and not a.scan_tree.columns_["mgmt"]["stretch"]
    assert list(a.scan_tree.columns_) == ["stt", "mgmt", "date", "vendor", "file", "status"]   # no full path column
    sp = _page(a, "run")
    sp["canvas"].bindings["<Configure>"](Ev(1600, 1400))
    assert _last_item(sp["canvas"], sp["win"]) == {"width": 1600, "height": 1400}


def test_treeview_vertical_and_horizontal_scrollbars(monkeypatch, tmp_path):
    gui, a, reg = _app(monkeypatch, tmp_path)
    assert a.scan_vsb.grid_info_["sticky"] == "ns" and a.scan_hsb.grid_info_["sticky"] == "ew"
    assert a.scan_tree.cfg["yscrollcommand"] == a.scan_vsb.set and a.scan_tree.cfg["xscrollcommand"] == a.scan_hsb.set
    assert all(c.get("minwidth") for c in a.scan_tree.columns_.values())        # columns keep a minimum -> hsb


# ------------------------------------------------------------------ 13-14 learning review windows
def _fake_candidates():
    return [types.SimpleNamespace(decision="review", source_name="r.pptx", slide=2, order=0, candidate_id="c1",
                                  user_label="", management_number="260901001-VOC", picture_id=1, shape_id=1,
                                  text="t", thumbnail="", slide_size=(100, 100), bounds=(0, 0, 10, 10),
                                  nearest_title="", nearest_heading="")]


def test_review_toplevels_are_resizable(monkeypatch, tmp_path):
    gui, a, reg = _app(monkeypatch, tmp_path)
    monkeypatch.setattr(a.ctl, "review_candidates", _fake_candidates)
    monkeypatch.setattr(a.ctl, "review_content_candidates", _fake_candidates)
    monkeypatch.setattr(a, "_review_render", lambda: None)
    monkeypatch.setattr(a, "_creview_render", lambda: None)
    a.open_image_review()
    a.open_content_review()
    tops = reg["toplevels"][-2:]
    for win in tops:
        assert ("resizable", True, True) in win.calls
        assert ("minsize", *gui.REVIEW_MIN_GEOMETRY) in win.calls and gui.REVIEW_MIN_GEOMETRY <= (700, 480)


def test_review_controls_remain_reachable_when_small(monkeypatch, tmp_path):
    gui, a, reg = _app(monkeypatch, tmp_path)
    monkeypatch.setattr(a.ctl, "review_candidates", _fake_candidates)
    monkeypatch.setattr(a.ctl, "review_content_candidates", _fake_candidates)
    monkeypatch.setattr(a, "_review_render", lambda: None)
    monkeypatch.setattr(a, "_creview_render", lambda: None)
    a.open_image_review()
    img = reg["toplevels"][-1]
    assert img.row_weights == {0: 1}                                            # only preview/details row grows
    for b in (a.btn_review_prev, a.btn_review_next, a.btn_review_save, *a.review_label_buttons.values()):
        assert b.pack_info_ and b.master.grid_info_.get("row") in (1, 2)        # fixed rows below the weighted one
    assert a.review_text.cfg["yscrollcommand"]                                  # details scroll on their own
    a.open_content_review()
    txt = reg["toplevels"][-1]
    assert txt.row_weights == {1: 1}
    for b in (a.btn_creview_prev, a.btn_creview_next, a.btn_creview_save, *a.creview_label_buttons.values()):
        assert b.pack_info_ and b.master.grid_info_.get("row") in (2, 3)
    assert a.creview_text.cfg["yscrollcommand"] and a.creview_info.cfg["yscrollcommand"]


# ------------------------------------------------------------------ 15 repeated resize = no duplicates
def test_repeated_resize_events_create_no_widgets(monkeypatch, tmp_path):
    gui, a, reg = _app(monkeypatch, tmp_path)
    n = len(reg["widgets"])
    for i in range(40):
        for name in ("cfg", "run", "learning"):
            sp = _page(a, name)
            sp["canvas"].bindings["<Configure>"](Ev(600 + i * 20, 300 + i * 15))
            sp["inner"].bindings["<Configure>"](Ev(600 + i * 20, 1500))
    assert len(reg["widgets"]) == n and len(a._scroll_pages) == 3


# ------------------------------------------------------------------ mouse wheel routing / geometry audit
def test_mouse_wheel_scrolls_only_the_page_under_the_pointer(monkeypatch, tmp_path):
    gui, a, reg = _app(monkeypatch, tmp_path)
    cfg, run = _page(a, "cfg"), _page(a, "run")
    a._on_page_wheel(types.SimpleNamespace(widget=a.lbl_update, delta=-120))
    assert cfg["canvas"].scrolls == [1] and run["canvas"].scrolls == []
    a._on_page_wheel(types.SimpleNamespace(widget=a.lbl_found, delta=120))
    assert run["canvas"].scrolls == [-1]
    a._on_page_wheel(types.SimpleNamespace(widget=a.txt_log, delta=-120))     # log scrolls itself
    a._on_page_wheel(types.SimpleNamespace(widget=a.scan_tree, delta=-120))   # table scrolls itself
    a._on_page_wheel(types.SimpleNamespace(widget=a.root, delta=-120))        # unrelated widget: nothing
    assert cfg["canvas"].scrolls == [1] and run["canvas"].scrolls == [-1]


def test_no_pack_grid_mixing_in_same_parent(monkeypatch, tmp_path):
    gui, a, reg = _app(monkeypatch, tmp_path)
    by_parent = {}
    for w in reg["widgets"]:
        if w.grid_info_:
            by_parent.setdefault(id(w.master), set()).add("grid")
        if w.pack_info_:
            by_parent.setdefault(id(w.master), set()).add("pack")
    assert all(len(v) == 1 for v in by_parent.values())


def test_version_and_build_prompt017():
    assert app.__version__ == "1.3.2" and app.BUILD_NUMBER == 15 and app.BUILD_LABEL == "Build 015"
    assert not re.search(r"wraplength=900\)", CFG_TAB + LEARNING_TAB)              # no fixed 900px wrap on either tab
    assert "self._scroll_page(self.tab_cfg" in CFG_TAB and "self._scroll_page(self.tab_run" in RUN_TAB
    assert "self._scroll_page(self.tab_learning" in LEARNING_TAB
