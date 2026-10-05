"""PROMPT-006B addendum: compact "Danh sách báo cáo" + manually removed rows auto-moved to the bottom.

Controller-layer grouping (GuiController.scan_rows) + GUI rendering through the fake tk harness of test_gui_redesign.
"""
import re
from pathlib import Path

from app.gui_controller import (GROUP_COMPLETED, GROUP_EXCLUDED, GROUP_PROCESSING, GROUP_REVIEW, GROUP_SKIPPED,
                                GROUP_WAITING, SCAN_FILTERS_VI, USER_EXCLUDED, GuiController, ScanRow)
from tests.test_gui_redesign import _make_app, _scan_tree_names
from tests.test_scan_list import world  # noqa: F401  (fixture reuse)

GUI_SRC = Path("app/gui.py").read_text(encoding="utf-8")
NAMES = ["260901001-VOC_A.pptx", "260901002-VOC_B.pptx", "260901003-VOC_C.pptx", "260901004-VOC_D.pptx"]


def _app(monkeypatch, tmp_path, names=NAMES):
    gui, a, reg, folder = _make_app(monkeypatch, tmp_path, names=names)
    a.var_scan_filter.set(SCAN_FILTERS_VI[1])
    a._render_scan()
    return gui, a, folder


def _tree_rows(a):
    return [v["values"] for v in a.scan_tree.items.values()]


def _select_name(a, name):
    iid = next(i for i, v in a.scan_tree.items.items() if v["values"][4] == name)
    a.scan_tree.selection_set(iid)
    return iid


def _row(ctl, name):
    return next(r for r in ctl.scan_rows() if r.path.name == name)


# ------------------------------------------------------------------ 1. compact columns, path internal
def test_compact_columns_and_internal_path(monkeypatch, tmp_path):
    gui, a, folder = _app(monkeypatch, tmp_path)
    assert [c[1] for c in gui.SCAN_COLUMNS] == ["STT", "Management Number", "Ngày phát sinh", "Vendor", "Tên file",
                                                "Trạng thái"]
    assert [c[3] for c in gui.SCAN_COLUMNS] == [False, False, False, False, True, False]   # only Tên file stretches
    for v in _tree_rows(a):
        assert len(v) == 6
        assert not any(str(folder) in str(x) for x in v)                        # no full path on the main table
    row = _row(a.ctl, "260901001-VOC_A.pptx")
    assert str(row.path) == str(folder / "260901001-VOC_A.pptx")              # kept internally
    assert row.as_values(1) == (1, "260901001-VOC", "01/09/2026", "", "260901001-VOC_A.pptx", row.status_vi)
    assert str(row.path) in row.details() and "Đường dẫn:" in row.details()   # reachable through diagnostics
    assert "Trạng thái quét" not in [c[1] for c in gui.SCAN_COLUMNS] and '"Đường dẫn"' not in GUI_SRC.split("SCAN_COLUMNS")[1].split("\n\n")[0]


# ------------------------------------------------------------------ 2-5. remove middle / multiple / restore / STT
def test_remove_middle_row_moves_it_to_bottom_without_rescan(monkeypatch, tmp_path):
    gui, a, folder = _app(monkeypatch, tmp_path)
    assert _scan_tree_names(a) == NAMES
    scans = []
    monkeypatch.setattr(a.ctl, "scan", lambda *x, **k: scans.append(1))
    monkeypatch.setattr(a.ctl, "scan_async", lambda *x, **k: scans.append(1))
    _select_name(a, "260901002-VOC_B.pptx")
    a.exclude_selected()
    assert _scan_tree_names(a) == ["260901001-VOC_A.pptx", "260901003-VOC_C.pptx", "260901004-VOC_D.pptx",
                                   "260901002-VOC_B.pptx"]
    assert _tree_rows(a)[-1][5] == "Đã loại thủ công" and _tree_rows(a)[-1][0] == 4
    assert [v[0] for v in _tree_rows(a)] == [1, 2, 3, 4]                       # STT = display order
    assert scans == [] and (folder / "260901002-VOC_B.pptx").exists()           # immediate, GUI/controller only


def test_multiple_removed_rows_keep_original_scan_order_at_bottom(monkeypatch, tmp_path):
    gui, a, folder = _app(monkeypatch, tmp_path)
    for name in ("260901003-VOC_C.pptx", "260901001-VOC_A.pptx"):              # remove C first, then A
        _select_name(a, name)
        a.exclude_selected()
    assert _scan_tree_names(a) == ["260901002-VOC_B.pptx", "260901004-VOC_D.pptx",
                                   "260901001-VOC_A.pptx", "260901003-VOC_C.pptx"]   # A before C (scan order)
    assert [v[5] for v in _tree_rows(a)][2:] == ["Đã loại thủ công"] * 2
    assert [v[0] for v in _tree_rows(a)] == [1, 2, 3, 4]


def test_restore_returns_row_to_its_group_and_queue(monkeypatch, tmp_path):
    gui, a, folder = _app(monkeypatch, tmp_path)
    _select_name(a, "260901002-VOC_B.pptx")
    a.exclude_selected()
    assert [it.path.name for it in a.ctl.final_queue()] == ["260901001-VOC_A.pptx", "260901003-VOC_C.pptx",
                                                            "260901004-VOC_D.pptx"]
    _select_name(a, "260901002-VOC_B.pptx")
    a.restore_selected()
    assert _scan_tree_names(a) == NAMES                                         # back to A B C D
    assert _tree_rows(a)[1][5] != "Đã loại thủ công" and [v[0] for v in _tree_rows(a)] == [1, 2, 3, 4]
    assert [it.path.name for it in a.ctl.final_queue()] == NAMES


def test_stt_is_display_order_and_identity_is_unchanged(monkeypatch, tmp_path):
    gui, a, folder = _app(monkeypatch, tmp_path)
    before = {r.path.name: (r.index, r.key) for r in a.ctl.scan_rows()}
    _select_name(a, "260901001-VOC_A.pptx")
    a.exclude_selected()
    after = {r.path.name: (r.index, r.key) for r in a.ctl.scan_rows()}
    assert before == after                                                      # index/key never renumbered
    assert [v[0] for v in _tree_rows(a)] == [1, 2, 3, 4]                       # STT recalculated
    assert _tree_rows(a)[-1][4] == "260901001-VOC_A.pptx"
    # the GUI maps tree items to the controller index (not to STT)
    iid = next(i for i, v in a.scan_tree.items.items() if v["values"][4] == "260901001-VOC_A.pptx")
    assert a.scan_items[iid] == before["260901001-VOC_A.pptx"][0]


# ------------------------------------------------------------------ 6-7. batch safety / rescan safety
def test_removed_rows_never_enter_the_batch_queue(monkeypatch, tmp_path):
    gui, a, folder = _app(monkeypatch, tmp_path)
    _select_name(a, "260901004-VOC_D.pptx")
    a.exclude_selected()
    q = a.ctl.final_queue()
    assert all(it.path.name != "260901004-VOC_D.pptx" for it in q) and len(q) == 3
    assert a.ctl.queue_counts()["excluded"] == 1


def test_rescan_and_refresh_do_not_reactivate_removed_rows(monkeypatch, tmp_path):
    gui, a, folder = _app(monkeypatch, tmp_path)
    _select_name(a, "260901002-VOC_B.pptx")
    a.exclude_selected()
    a.ctl.scan()                                                                 # same folder / same inputs
    a._render_scan()
    assert _row(a.ctl, "260901002-VOC_B.pptx").excluded
    assert _scan_tree_names(a)[-1] == "260901002-VOC_B.pptx" and _tree_rows(a)[-1][5] == "Đã loại thủ công"
    for _ in range(3):                                                           # repeated refreshes
        a._render_scan()
        a._render_scan_state()
    assert sorted(_scan_tree_names(a)) == sorted(NAMES) and len(_scan_tree_names(a)) == 4   # no dup / no loss
    assert _row(a.ctl, "260901002-VOC_B.pptx").excluded
    assert "260901002-VOC_B.pptx" not in [it.path.name for it in a.ctl.final_queue()]


# ------------------------------------------------------------------ 8-9. search / filter
def test_search_finds_removed_rows_without_reactivating(monkeypatch, tmp_path):
    gui, a, folder = _app(monkeypatch, tmp_path)
    _select_name(a, "260901003-VOC_C.pptx")
    a.exclude_selected()
    a.set_search("260901003")
    assert _scan_tree_names(a) == ["260901003-VOC_C.pptx"] and _tree_rows(a)[0][5] == "Đã loại thủ công"
    assert _tree_rows(a)[0][0] == 1                                              # STT within the current view
    a.set_search("_c")
    assert _scan_tree_names(a) == ["260901003-VOC_C.pptx"]
    a.set_search("")
    assert _row(a.ctl, "260901003-VOC_C.pptx").excluded and _scan_tree_names(a)[-1] == "260901003-VOC_C.pptx"


def test_removed_filter_and_presentation_only_filtering(monkeypatch, tmp_path):
    gui, a, folder = _app(monkeypatch, tmp_path)
    _select_name(a, "260901001-VOC_A.pptx")
    a.exclude_selected()
    assert SCAN_FILTERS_VI[3] == "File đã loại thủ công"
    a.var_scan_filter.set(SCAN_FILTERS_VI[3])
    a._render_scan()
    assert _scan_tree_names(a) == ["260901001-VOC_A.pptx"]
    a.var_scan_filter.set(SCAN_FILTERS_VI[2])                                   # skipped view still contains it
    a._render_scan()
    assert "260901001-VOC_A.pptx" in _scan_tree_names(a)
    a.var_scan_filter.set(SCAN_FILTERS_VI[0])                                   # "cần xử lý" hides it
    a._render_scan()
    assert "260901001-VOC_A.pptx" not in _scan_tree_names(a)
    assert a.ctl.queue_counts()["excluded"] == 1                                 # state untouched by filtering


# ------------------------------------------------------------------ 10. diagnostics path
def test_double_click_diagnostics_shows_full_path(monkeypatch, tmp_path):
    gui, a, folder = _app(monkeypatch, tmp_path)
    shown = []
    monkeypatch.setattr(gui.messagebox, "showinfo", lambda title, msg: shown.append(msg))
    _select_name(a, "260901002-VOC_B.pptx")
    a.exclude_selected()
    iid = _select_name(a, "260901002-VOC_B.pptx")
    a.scan_tree._focus = iid
    monkeypatch.setattr(a.scan_tree, "focus", lambda *x: iid, raising=False)
    a._on_scan_double_click(None)
    assert shown and str(folder / "260901002-VOC_B.pptx") in shown[-1] and "Đã loại thủ công" in shown[-1]
    assert "nhấp đúp" in GUI_SRC.lower() or "double" in GUI_SRC.lower()


# ------------------------------------------------------------------ 11. grouping order (controller layer, pure)
def _mk(i, **kw):
    base = dict(index=i, key=f"k{i}", path=Path(f"/r/{i}.pptx"), management_number="260901001-VOC",
                occurrence_date="01/09/2026", action="PROCESS", excluded=False, reason="")
    base.update(kw)
    return ScanRow(**base)


def test_group_ranks_are_deterministic_and_removed_is_always_last():
    assert _mk(0, run_stage="extracting").group == GROUP_PROCESSING
    assert _mk(1).group == GROUP_WAITING and _mk(1, run_stage="waiting").group == GROUP_WAITING
    assert _mk(2, action="PROCESS_NEW_ROW").group == GROUP_WAITING             # new row = will process
    assert _mk(3, run_stage="needs_review").group == GROUP_REVIEW
    assert _mk(3, run_stage="error").group == GROUP_REVIEW
    assert _mk(4, action="INVALID_MANAGEMENT_NUMBER").group == GROUP_REVIEW
    assert _mk(5, run_stage="completed").group == GROUP_COMPLETED
    assert _mk(5, run_stage="completed_new").group == GROUP_COMPLETED
    assert _mk(6, action="FAST_SKIP").group == GROUP_SKIPPED
    assert _mk(7, action="OUTSIDE_PERIOD").group == GROUP_SKIPPED
    for st in ("", "extracting", "completed", "needs_review", "error"):
        r = _mk(8, run_stage=st, excluded=True)
        assert r.group == GROUP_EXCLUDED and r.status_vi == "Đã loại thủ công"
    assert GROUP_PROCESSING < GROUP_WAITING < GROUP_REVIEW < GROUP_COMPLETED < GROUP_SKIPPED < GROUP_EXCLUDED
    assert USER_EXCLUDED == "USER_EXCLUDED"


def test_scan_rows_sorted_by_group_then_scan_order(world):
    ctl: GuiController = world.ctl
    ctl.scan()
    rows = ctl.scan_rows()
    groups = [r.group for r in rows]
    assert groups == sorted(groups)
    for g in set(groups):
        idx = [r.index for r in rows if r.group == g]
        assert idx == sorted(idx)                                               # stable inside each group
    ctl.exclude([rows[0].index])
    rows2 = ctl.scan_rows()
    assert rows2[-1].index == rows[0].index and rows2[-1].group == GROUP_EXCLUDED
    assert [r.index for r in rows2 if r.group != GROUP_EXCLUDED] == [r.index for r in rows if r.index != rows[0].index]


# ------------------------------------------------------------------ 12. no regrouping churn during processing
def test_gui_regroups_scan_list_only_on_final_row_status():
    body = GUI_SRC[GUI_SRC.index("if ev.kind == \"row\":"):GUI_SRC.index("elif ev.kind == \"progress\":")]
    assert re.search(r"is_final:\s*\n\s*self\._render_scan\(\)", body)
    done = GUI_SRC[GUI_SRC.rindex("self._done_handled = True"):][:600]
    assert "self._render_scan()" in done


