"""PROMPT-004B – explicit, readable ttk styles in every widget state (no reliance on platform ttk defaults)."""
import pytest

from tests.test_gui_redesign import _make_app, _running

CLAM, VISTA = "clam", "vista"


def _gui(monkeypatch, tmp_path):
    gui, a, reg, _ = _make_app(monkeypatch, tmp_path, 2)
    return gui, a


def _fg_bg(spec, name, state="normal"):
    e = spec[name]
    fg = e["configure"].get("foreground")
    bg = e["configure"].get("background") or e["configure"].get("fieldbackground")
    if state != "normal":
        m = e.get("map", {})
        fg = dict(m.get("foreground", [])).get(state, fg)
        bg = dict(m.get("background", [])).get(state, dict(m.get("fieldbackground", [])).get(state, bg))
    return fg, bg


# ------------------------------------------------------------------ pure colour maths
def test_contrast_ratio_reference_values(monkeypatch, tmp_path):
    gui, _ = _gui(monkeypatch, tmp_path)
    assert gui.contrast_ratio("#000000", "#ffffff") == pytest.approx(21.0)
    assert gui.contrast_ratio("#ffffff", "#ffffff") == pytest.approx(1.0)
    assert gui.contrast_ratio("#ffffff", "#f0f0f0") < 1.2          # the old invisible disabled-state combination


@pytest.mark.parametrize("theme", [CLAM, VISTA, "winnative", ""])
def test_every_text_style_readable_in_every_state(monkeypatch, tmp_path, theme):
    gui, _ = _gui(monkeypatch, tmp_path)
    spec = gui.build_style_spec(theme)
    assert gui.audit_style_contrast(spec) == []
    for name, e in spec.items():                                      # explicit fg+bg wherever text is drawn
        if name.endswith(("TLabel", "TButton", "Treeview", "Tab", "Heading", "TEntry", "TCombobox", "TSpinbox",
                          "TCheckbutton", "TRadiobutton")) or name == ".":
            assert e["configure"].get("foreground"), name
            assert e["configure"].get("background") or e["configure"].get("fieldbackground"), name


# ------------------------------------------------------------------ Start / Stop buttons
@pytest.mark.parametrize("theme", [CLAM, VISTA])
@pytest.mark.parametrize("state", ["normal", "active", "pressed", "disabled"])
def test_start_button_text_visible(monkeypatch, tmp_path, theme, state):
    gui, _ = _gui(monkeypatch, tmp_path)
    fg, bg = _fg_bg(gui.build_style_spec(theme), "Primary.TButton", state)
    assert fg and bg and fg.lower() != bg.lower()
    assert gui.contrast_ratio(fg, bg) >= (gui.MIN_CONTRAST_SOFT if state == "disabled" else gui.MIN_CONTRAST)


def test_native_theme_never_uses_white_text_on_unpaintable_button(monkeypatch, tmp_path):
    """Windows 'vista' ignores TButton background -> white text would vanish on the light native face."""
    gui, _ = _gui(monkeypatch, tmp_path)
    for theme in gui.NATIVE_THEMES:
        for state in ("normal", "active", "disabled"):
            fg, _bg = _fg_bg(gui.build_style_spec(theme), "Primary.TButton", state)
            assert gui.contrast_ratio(fg, gui.PALETTE["button_face"]) >= gui.MIN_CONTRAST_SOFT
            assert gui.contrast_ratio(fg, "#f0f0f0") >= gui.MIN_CONTRAST_SOFT          # classic Windows face
    assert "#f0f0f0" not in str(gui.build_style_spec(CLAM)["Primary.TButton"])         # old mapping removed


@pytest.mark.parametrize("theme", [CLAM, VISTA])
@pytest.mark.parametrize("state", ["normal", "active", "disabled"])
def test_stop_button_and_plain_buttons_visible(monkeypatch, tmp_path, theme, state):
    gui, _ = _gui(monkeypatch, tmp_path)
    spec = gui.build_style_spec(theme)
    for name in ("Danger.TButton", "TButton"):
        fg, bg = _fg_bg(spec, name, state)
        assert fg.lower() != bg.lower()
        assert gui.contrast_ratio(fg, bg) >= (gui.MIN_CONTRAST_SOFT if state == "disabled" else gui.MIN_CONTRAST)


def test_start_stop_widgets_use_the_explicit_styles_and_states(monkeypatch, tmp_path):
    gui, a = _gui(monkeypatch, tmp_path)
    assert a.btn_start.k["style"] == "Primary.TButton" and a.btn_stop.k["style"] == "Danger.TButton"
    assert a.btn_start.cfg.get("state", "normal") == "normal" and a.btn_stop.cfg["state"] == "disabled"
    _running(a)
    assert a.btn_start.cfg["state"] == "disabled" and a.btn_stop.cfg["state"] == "normal"
    a._set_running(False)
    assert a.btn_start.cfg["state"] == "normal" and a.btn_stop.cfg["state"] == "disabled"


# ------------------------------------------------------------------ labels: percent / elapsed / ETA / stage / status
def test_progress_elapsed_eta_stage_labels_have_readable_styles(monkeypatch, tmp_path):
    gui, a = _gui(monkeypatch, tmp_path)
    spec = gui.build_style_spec(CLAM)
    styled = {"lbl_percent": "Percent.TLabel", "lbl_progress": "Progress.TLabel"}
    for attr, style in styled.items():
        if hasattr(a, attr):
            assert a.__dict__[attr].k.get("style") == style
    for name in ("Percent.TLabel", "Progress.TLabel", "Field.TLabel", "Secondary.TLabel", "CardValue.TLabel",
                 "Section.TLabel", "Header.TLabel", "SubHeader.TLabel", "Version.TLabel", "Card.TLabel", "TLabel"):
        fg, bg = _fg_bg(spec, name)
        assert fg and bg and gui.contrast_ratio(fg, bg) >= gui.MIN_CONTRAST, name
    # the time / ETA / stage / file labels exist and carry one of the explicit card styles
    for attr in ("lbl_stage", "lbl_file", "lbl_elapsed", "lbl_eta", "lbl_scan", "lbl_queue"):
        w = getattr(a, attr, None)
        if w is not None:
            assert w.k.get("style", "TLabel") in spec, attr


@pytest.mark.parametrize("style", ["Success.TLabel", "Warning.TLabel", "Error.TLabel", "Secondary.TLabel"])
def test_status_and_summary_label_styles_readable(monkeypatch, tmp_path, style):
    gui, _ = _gui(monkeypatch, tmp_path)
    fg, bg = _fg_bg(gui.build_style_spec(CLAM), style)
    assert bg == gui.PALETTE["card"] and gui.contrast_ratio(fg, bg) >= gui.MIN_CONTRAST
    assert fg.lower() != bg.lower()


# ------------------------------------------------------------------ Treeview rows
def test_treeview_normal_and_selected_rows_readable(monkeypatch, tmp_path):
    gui, a = _gui(monkeypatch, tmp_path)
    spec = gui.build_style_spec(CLAM)
    fg, bg = _fg_bg(spec, "Status.Treeview")
    assert gui.contrast_ratio(fg, bg) >= gui.MIN_CONTRAST
    sfg, sbg = _fg_bg(spec, "Status.Treeview", "selected")
    assert sfg.lower() != sbg.lower() and gui.contrast_ratio(sfg, sbg) >= gui.MIN_CONTRAST
    # coloured status tags stay readable on white, on the stripe and on the selection colour
    for color in set(gui.SCAN_TAG_COLORS.values()) | set(gui.RESULT_TAG_COLORS.values()):
        assert gui.contrast_ratio(color, gui.PALETTE["card"]) >= gui.MIN_CONTRAST
        assert gui.contrast_ratio(color, gui.PALETTE["stripe"]) >= gui.MIN_CONTRAST
        assert gui.contrast_ratio(color, gui.PALETTE["selection"]) >= gui.MIN_CONTRAST_SOFT
    hfg, hbg = _fg_bg(spec, "Status.Treeview.Heading")
    assert gui.contrast_ratio(hfg, hbg) >= gui.MIN_CONTRAST
    assert a.tree.k.get("style") == "Status.Treeview" and a.scan_tree.k.get("style") == "Status.Treeview"


# ------------------------------------------------------------------ Entry / Combobox / Notebook tabs
@pytest.mark.parametrize("name,state", [("TEntry", "normal"), ("TEntry", "disabled"), ("TCombobox", "normal"),
                                        ("TCombobox", "readonly"), ("TCombobox", "disabled"), ("TSpinbox", "normal"),
                                        ("TNotebook.Tab", "normal"), ("TNotebook.Tab", "selected"),
                                        ("TNotebook.Tab", "active"), ("TLabelframe.Label", "normal")])
def test_inputs_and_tabs_readable(monkeypatch, tmp_path, name, state):
    gui, _ = _gui(monkeypatch, tmp_path)
    fg, bg = _fg_bg(gui.build_style_spec(CLAM), name, state)
    assert fg and bg and fg.lower() != bg.lower()
    assert gui.contrast_ratio(fg, bg) >= (gui.MIN_CONTRAST_SOFT if state == "disabled" else gui.MIN_CONTRAST)


# ------------------------------------------------------------------ Ollama ready / warning / error
def test_ollama_states_use_readable_styles(monkeypatch, tmp_path):
    from app.gui_controller import UiEvent
    gui, a = _gui(monkeypatch, tmp_path)
    spec = gui.build_style_spec(CLAM)
    for ok, source, expected in ((True, "local", "Success.TLabel"), (False, "local", "Warning.TLabel"),
                                 (False, "none", "Error.TLabel")):
        a.ctl.ollama_source = source
        a.ctl.ollama_ok = ok
        a.ctl._queue.put(UiEvent("autoconnect", (ok, "msg")))
        a._poll()
        assert a.lbl_conn.cfg["style"] == expected
        fg, bg = _fg_bg(spec, expected)
        assert gui.contrast_ratio(fg, bg) >= gui.MIN_CONTRAST and fg.lower() != bg.lower()


# ------------------------------------------------------------------ theme selection + application
def test_theme_prefers_clam_and_applies_every_style(monkeypatch, tmp_path):
    gui, a = _gui(monkeypatch, tmp_path)
    assert a.theme == "clam"                                           # honours explicit colours in all states
    spec = gui.build_style_spec("clam")
    recorded = _recorded_styles(a)
    for name in spec:
        assert name in recorded, name
        for key in ("foreground", "background", "fieldbackground"):
            if key in spec[name].get("configure", {}):
                assert recorded[name].get(key) == spec[name]["configure"][key], (name, key)
        if spec[name].get("map"):
            assert recorded[name]["map"] == spec[name]["map"]


def _recorded_styles(a):
    import gc
    for obj in gc.get_objects():
        if type(obj).__name__ == "Style" and hasattr(obj, "styles") and "Primary.TButton" in getattr(obj, "styles", {}):
            return obj.styles
    raise AssertionError("fake ttk.Style not found")
