"""PROMPT-002 layout rules verified on the view source (tkinter is unavailable headless).

The view is a thin layer; these tests pin the SECTION ORDER of the redesigned processing page, that every
workflow control sits in the toolbar ABOVE the file list, and that both tables own a vertical + horizontal
scrollbar while the page has its own scroll layer.  (The former numbered LabelFrames "1. … 8." and the
standalone "3. AI" / "4. Các nút chức năng" blocks were removed by PROMPT-002.)
"""
import re
from pathlib import Path

SRC = (Path(__file__).resolve().parent.parent / "app" / "gui.py").read_text(encoding="utf-8")
RUN_TAB = SRC[SRC.index("def _build_run_tab("):SRC.index("def _build_cfg_tab(")]
CFG_TAB = SRC[SRC.index("def _build_cfg_tab("):SRC.index("def _bind_shortcuts(")]


def _pos(snippet: str, src: str = RUN_TAB) -> int:
    assert snippet in src, snippet
    return src.index(snippet)


def test_section_order_processing_page():
    order = ['"Nguồn dữ liệu"', '"Thời gian xử lý"', '"Danh sách báo cáo"', '"Xử lý"', '"Kết quả xử lý"']
    positions = [_pos(s) for s in order]
    assert positions == sorted(positions)
    for gone in ('"1. Nguồn dữ liệu"', '"3. AI"', '"4. Các nút chức năng"', '"5. Thống kê quét"', "ttk.LabelFrame("):
        assert gone not in SRC, gone


def test_toolbar_controls_above_file_list():
    list_pos = _pos("self.scan_tree, self.scan_vsb, self.scan_hsb = self._make_table(")
    for ctrl in ('text="Quét lại"', 'text="Khôi phục"', 'text="Xóa khỏi danh sách"', "self.e_search = ttk.Entry(",
                 "self.cb_scan_filter = ttk.Combobox("):
        assert _pos(ctrl) < list_pos, ctrl
    # Start / Stop are the primary actions of the "Xử lý" card, below the list and above the result table
    assert list_pos < _pos('text="▶  BẮT ĐẦU XỬ LÝ", style="Primary.TButton"') < _pos("self.tree, self.result_vsb")
    assert 'style="Danger.TButton", command=self.stop' in RUN_TAB and "variable=self.var_force" in RUN_TAB
    # all three entry points still call the single controller-backed handler
    assert RUN_TAB.count("command=self.exclude_selected") >= 2 and "lambda _e: self.exclude_selected()" in RUN_TAB


def test_no_standalone_ai_block_on_processing_page():
    assert "Cấu hình Ollama…" not in RUN_TAB and "self.var_host" not in RUN_TAB and "self.cb_model" not in RUN_TAB
    assert "ollama_indicator_text(self.ctl)" in RUN_TAB                 # compact status indicator only
    assert "self.var_host = tk.StringVar()" in CFG_TAB and "self.cb_model = ttk.Combobox(" in CFG_TAB


def test_tables_have_own_scrollbars():
    helper = SRC[SRC.index("def _make_table("):SRC.index("# ------------------------------------------------------------------ layout")]
    assert 'orient="vertical", command=tree.yview' in helper and 'orient="horizontal", command=tree.xview' in helper
    assert "yscrollcommand=vsb.set, xscrollcommand=hsb.set" in helper
    assert "SCAN_TREE_ROWS" in RUN_TAB and "RESULT_TREE_ROWS" in RUN_TAB
    assert "ttk.PanedWindow" not in SRC
    assert "self.txt_log.configure(yscrollcommand=log_vsb.set)" in CFG_TAB


def test_file_list_columns():
    cols = re.search(r"SCAN_COLUMNS = \((.*?)\)\)\n", SRC, re.S).group(0)
    titles = re.findall(r'\("(\w+)", "([^"]+)", (\d+), (True|False)\)', cols)
    names = [t[1] for t in titles]
    assert names == ["STT", "Management Number", "Ngày phát sinh", "Vendor", "Tên file", "Trạng thái quét", "Đường dẫn"]
    widths = {t[1]: int(t[2]) for t in titles}
    stretch = {t[1]: t[3] == "True" for t in titles}
    assert widths["Tên file"] >= 300 and widths["Đường dẫn"] >= 250 and widths["Management Number"] >= 150
    assert stretch["Tên file"] and not stretch["Management Number"] and not stretch["STT"]


def test_whole_page_scroll_layer():
    assert "self.page_canvas = tk.Canvas(self.tab_run" in RUN_TAB
    assert "command=self.page_canvas.yview" in RUN_TAB and "yscrollcommand=self.page_vsb.set" in RUN_TAB
    assert "create_window((0, 0), window=self.page" in RUN_TAB
    wheel = SRC[SRC.index("def _on_page_wheel("):SRC.index("# ------------------------------------------------------------------ controller <-> widgets")]
    assert "ttk.Treeview" in wheel and "tk.Text" in wheel and "self.tab_cfg" in wheel
    resize = SRC[SRC.index("def _on_canvas_configure("):SRC.index("def _on_page_wheel(")]
    assert "max(event.height, req_h)" in resize                          # tables grow with the window


def test_styling_is_centralised():
    styles = SRC[SRC.index("def build_style_spec("):SRC.index("def apply_dpi_awareness(")]
    for name in ("App.TFrame", "Card.TFrame", "Header.TLabel", "Section.TLabel", "Secondary.TLabel", "Primary.TButton",
                 "Danger.TButton", "Success.TLabel", "Warning.TLabel", "Error.TLabel", "Status.Treeview"):
        assert f'"{name}"' in styles, name
    build = SRC[SRC.index("def _build("):SRC.index("def _bind_shortcuts(")]
    assert "font=(" not in build.replace("font=(MONO_FAMILY, 9)", "")   # fonts come from styles, not widgets
    assert "ttk.Style(" not in build and "theme_use(" not in build
    assert "XS, S, M, L, XL = 4, 6, 10, 16, 24" in SRC
