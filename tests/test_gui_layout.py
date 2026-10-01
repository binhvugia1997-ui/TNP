"""Layout rules #63-#70 for tab "Xử lý báo cáo" (tkinter is unavailable headless -> verified on the view source).

The view is a thin layer; these tests pin the SECTION ORDER, that every workflow button is built ABOVE the file
list, and that both tables own a vertical + horizontal scrollbar while the page has its own scroll layer.
"""
import re
from pathlib import Path

SRC = (Path(__file__).resolve().parent.parent / "app" / "gui.py").read_text(encoding="utf-8")
BUILD = SRC[SRC.index("def _build("):SRC.index("def _make_table(")]


def _pos(snippet: str) -> int:
    assert snippet in BUILD, snippet
    return BUILD.index(snippet)


def test_section_order_rule_63():
    order = ['text="1. Nguồn dữ liệu"', 'text="2. Thời gian xử lý"', 'text="3. AI"', 'text="4. Các nút chức năng"',
             'text="5. Thống kê quét"', 'text="6. Danh sách file', 'text="7. Tiến trình"', 'text="8. Kết quả xử lý']
    positions = [_pos(s) for s in order]
    assert positions == sorted(positions)


def test_action_buttons_above_file_list_rules_64_65():
    list_pos = _pos('text="6. Danh sách file')
    for btn in ('text="Quét file"', 'text="Quét lại"', 'text="Xóa khỏi danh sách"', 'text="Khôi phục"',
                'text="Bắt đầu xử lý"', 'text="Dừng sau báo cáo hiện tại"', "variable=self.var_force"):
        assert _pos(btn) < list_pos, btn
    # grouped: list actions row, separator, processing row; Start keeps its emphasised style
    assert _pos('text="Danh sách:"') < _pos("ttk.Separator(xf") < _pos('text="Xử lý:"')
    assert 'style="Start.TButton"' in BUILD
    # all three entry points still call the single controller-backed handlers
    assert BUILD.count("command=self.exclude_selected") >= 2 and "lambda _e: self.exclude_selected()" in BUILD


def test_tables_have_own_scrollbars_rules_66_69():
    helper = SRC[SRC.index("def _make_table("):SRC.index("def _on_page_configure(")]
    assert 'orient="vertical", command=tree.yview' in helper and 'orient="horizontal", command=tree.xview' in helper
    assert "yscrollcommand=vsb.set, xscrollcommand=hsb.set" in helper
    assert "stretch=False" in helper                                  # long names/paths scroll, not squeezed
    assert re.search(r"height=\d+", helper) or "height=height" in helper
    assert "SCAN_TREE_ROWS" in BUILD and "RESULT_TREE_ROWS" in BUILD  # fixed visible rows, rows scroll internally
    assert "ttk.PanedWindow" not in BUILD                             # replaced by page scroll + fixed-height tables


def test_file_list_columns_rule_68():
    import importlib.util
    cols = re.search(r"SCAN_COLUMNS = \((.*?)\)\)\n", SRC, re.S).group(0)
    titles = re.findall(r'\("(\w+)", "([^"]+)", (\d+)\)', cols)
    names = [t[1] for t in titles]
    for want in ("STT", "Management Number", "Ngày phát sinh", "Tên file", "Trạng thái quét", "Đường dẫn"):
        assert want in names
    widths = {t[1]: int(t[2]) for t in titles}
    assert widths["Tên file"] >= 300 and widths["Đường dẫn"] >= 250   # readable, reached via horizontal scroll
    assert importlib.util.find_spec("app.gui_controller") is not None


def test_whole_page_scroll_layer_rule_70():
    assert "self.page_canvas = tk.Canvas(self.tab_run" in BUILD
    assert 'command=self.page_canvas.yview' in BUILD and "yscrollcommand=self.page_vsb.set" in BUILD
    assert "create_window((0, 0), window=self.page" in BUILD
    # wheel over a table/log scrolls that widget, not the page
    wheel = SRC[SRC.index("def _on_page_wheel("):SRC.index("# ------------------------------------------------------------------ controller <-> widgets")]
    assert "ttk.Treeview" in wheel and "tk.Text" in wheel
