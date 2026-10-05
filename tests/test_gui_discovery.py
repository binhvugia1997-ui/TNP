"""PROMPT-003 – LAN discovery wiring in the controller and the (fake-tkinter) view.

Rules under test: user-initiated only, separate progress channel, disabled while a batch runs, explicit
confirmation before a server is applied (even for one result), persistence through the existing config path,
configured model kept if present / warning otherwise, never auto-download.
"""
import json
import sys
import threading
from pathlib import Path

import app
from app.config import AppConfig
from app.gui_controller import GuiController, UiEvent
from app.ollama_discovery import OllamaDiscoveryResult

from tests.test_gui_redesign import _make_app, _running

SRC = Path(app.__file__).with_name("gui.py").read_text(encoding="utf-8")
CTL_SRC = Path(app.__file__).with_name("gui_controller.py").read_text(encoding="utf-8")


class FakeDiscovery:
    """Deterministic stand-in for OllamaDiscovery.run()."""
    instances = []

    def __init__(self, results=None, total=10, block=None):
        self.port, self.checked, self.total = 11434, 0, total
        self._results = results or []
        self.results, self.networks = [], []
        self.cancel_event = threading.Event()
        self._block = block
        FakeDiscovery.instances.append(self)

    def cancel(self):
        self.cancel_event.set()

    @property
    def cancelled(self):
        return self.cancel_event.is_set()

    def run(self, progress=None, found=None):
        progress and progress(0, self.total)
        for i in range(1, self.total + 1):
            if self._block is not None and i == 3:
                self._block.wait(5)
            if self.cancel_event.is_set():
                break
            self.checked = i
            progress and progress(i, self.total)
        if not self.cancel_event.is_set():
            for r in self._results:
                self.results.append(r)
                found and found(r)
        return list(self.results)


class FakeClient:
    models_by_server = {}

    def __init__(self, server, timeout=None):
        self.server = server

    def test_connection(self):
        return {"ok": True, "models": list(FakeClient.models_by_server.get(self.server, []))}

    def list_models(self, timeout=None):
        return list(FakeClient.models_by_server.get(self.server, []))


def _ctl(tmp_path, results=None, **kw):
    cfg_path = tmp_path / "config.json"
    ctl = GuiController(AppConfig(period_mode="all", model="qwen3:4b"), client_factory=FakeClient, config_path=cfg_path)
    ctl._discovery_factory = lambda: FakeDiscovery(results, **kw)
    return ctl, cfg_path


def _drain(ctl, timeout=5.0):
    for t in threading.enumerate():
        if t.name == "ollama-discovery":
            t.join(timeout)
    return ctl.pump()


# ------------------------------------------------------------------ controller
def test_controller_discovery_lifecycle_and_separate_progress(tmp_path):
    res = OllamaDiscoveryResult("192.168.1.50", 11434, ["qwen3:4b"], 7)
    ctl, _ = _ctl(tmp_path, [res])
    assert ctl.can_discover() and ctl.discover_ollama_async() is True
    assert ctl.discovery_running and ctl.discover_ollama_async() is False      # no second concurrent scan
    events = _drain(ctl)
    kinds = [e.kind for e in events]
    assert "discovery_progress" in kinds and "discovery_found" in kinds and kinds[-1] == "discovery_done"
    assert not any(k in ("progress", "row", "done") for k in kinds)              # never touches batch progress
    assert ctl.progress.done == 0 and not ctl.progress.finished and ctl.state == "idle"
    assert ctl.discovery_results == [res] and ctl.discovery_state == "idle"
    assert "Tìm thấy 1 server" in ctl.discovery_message
    assert ctl.server == "http://127.0.0.1:11434"                                  # NOT auto-switched


def test_controller_progress_text_format(tmp_path):
    ctl, _ = _ctl(tmp_path)
    ctl.discovery_state = "running"
    ctl.apply_event(UiEvent("discovery_progress", (37, 254)))
    assert ctl.discovery_progress_text() == "Đang tìm Ollama trong mạng LAN... Đã kiểm tra 37 / 254 địa chỉ"


def test_controller_discovery_blocked_while_batch_runs(tmp_path):
    ctl, _ = _ctl(tmp_path)
    ctl.state = "running"
    assert ctl.can_discover() is False and ctl.discover_ollama_async() is False
    ctl.state = "stopping"
    assert ctl.discover_ollama_async() is False
    ctl.state = "idle"
    assert ctl.can_discover() is True


def test_controller_cancel_discovery(tmp_path):
    gate = threading.Event()
    ctl, _ = _ctl(tmp_path, [OllamaDiscoveryResult("192.168.1.50")], total=10, block=gate)
    assert ctl.cancel_discovery() is False                                        # nothing running
    ctl.discover_ollama_async()
    for _ in range(200):
        if FakeDiscovery.instances and FakeDiscovery.instances[-1].checked >= 2:
            break
        threading.Event().wait(0.01)
    assert ctl.cancel_discovery() is True
    gate.set()
    _drain(ctl)
    assert ctl.discovery_results == [] and "(đã dừng)" in ctl.discovery_message
    assert ctl.discovery_checked < ctl.discovery_total


def test_controller_no_results_message(tmp_path):
    ctl, _ = _ctl(tmp_path, [])
    ctl.discover_ollama_async()
    _drain(ctl)
    assert "Không tìm thấy Ollama trong mạng LAN." in ctl.discovery_message
    assert "nhập địa chỉ server thủ công" in ctl.discovery_message


def test_apply_discovered_server_persists_and_keeps_model(tmp_path):
    FakeClient.models_by_server = {"http://192.168.1.50:11434": ["llama3:8b", "qwen3:4b"]}
    ctl, cfg_path = _ctl(tmp_path)
    res = OllamaDiscoveryResult("192.168.1.50", 11434, ["llama3:8b", "qwen3:4b"], 7)
    ok, msg = ctl.apply_discovered_server(res)
    assert ok is True and "192.168.1.50:11434" in msg
    assert ctl.host == "192.168.1.50" and ctl.port == 11434 and ctl.model == "qwen3:4b"
    assert ctl.available_models == ["llama3:8b", "qwen3:4b"]
    data = json.loads(cfg_path.read_text(encoding="utf-8"))
    assert data["ollama_server"] == "http://192.168.1.50:11434" and data["model"] == "qwen3:4b"
    assert GuiController(AppConfig.load(cfg_path), client_factory=FakeClient, config_path=cfg_path).host == "192.168.1.50"


def test_apply_discovered_server_warns_when_model_missing_and_never_switches_model(tmp_path):
    FakeClient.models_by_server = {"http://192.168.1.60:11434": ["llama3:8b"]}
    ctl, cfg_path = _ctl(tmp_path)
    ok, msg = ctl.apply_discovered_server(OllamaDiscoveryResult("192.168.1.60", 11434, ["llama3:8b"], 3))
    assert ok is False and msg == "Model qwen3:4b chưa có trên server này.\nVui lòng chọn model có sẵn."
    assert ctl.model == "qwen3:4b"                                               # kept, user decides
    assert ctl.host == "192.168.1.60" and json.loads(cfg_path.read_text(encoding="utf-8"))["model"] == "qwen3:4b"
    assert "pull" not in CTL_SRC[CTL_SRC.index("def apply_discovered_server"):CTL_SRC.index("# ------------------------------------------------------------------ run control")]


def test_apply_uses_scan_models_when_refresh_fails(tmp_path):
    class Dead(FakeClient):
        def test_connection(self):
            raise RuntimeError("gone")
    cfg_path = tmp_path / "config.json"
    ctl = GuiController(AppConfig(period_mode="all", model="qwen3:4b"), client_factory=Dead, config_path=cfg_path)
    ok, _ = ctl.apply_discovered_server(OllamaDiscoveryResult("192.168.1.70", 11434, ["qwen3:4b"], 3))
    assert ok is True and ctl.available_models == ["qwen3:4b"]


def test_controller_never_scans_at_startup(tmp_path):
    FakeDiscovery.instances.clear()
    ctl, _ = _ctl(tmp_path)
    assert FakeDiscovery.instances == [] and ctl.discovery_runs == 0 and ctl.discovery_state == "idle"
    assert "discover_ollama" not in CTL_SRC[CTL_SRC.index("def __init__"):CTL_SRC.index("# ------------------------------------------------------------------ Ollama endpoint")]


# ------------------------------------------------------------------ view (fake tkinter)
def _app(monkeypatch, tmp_path, results=None, **kw):
    gui, a, reg, folder = _make_app(monkeypatch, tmp_path, 1)
    a.ctl._client_factory = FakeClient
    a.ctl._discovery_factory = lambda: FakeDiscovery(results, **kw)
    return gui, a, reg


def _finish(a):
    for t in threading.enumerate():
        if t.name == "ollama-discovery":
            t.join(5)
    a._poll()                                           # the real poll loop renders the queued discovery events


def test_view_has_button_in_config_tab_and_no_auto_scan(monkeypatch, tmp_path):
    gui, a, reg = _app(monkeypatch, tmp_path)
    assert a.btn_discover.cfg["text"] == "Tìm Ollama trong mạng LAN"
    assert a.btn_discover_stop.cfg["text"] == "Dừng tìm" and a.btn_discover_stop.cfg["state"] == "disabled"
    assert a.btn_discover in a._config_widgets
    cfg_block = SRC[SRC.index("def _build_cfg_tab"):SRC.index("def _bind_shortcuts")]
    assert "Tìm Ollama trong mạng LAN" in cfg_block
    assert "Tìm Ollama" not in SRC[SRC.index("def _build_run_tab"):SRC.index("def _build_cfg_tab")]
    assert a.ctl.discovery_runs == 0
    # only the button command triggers a scan
    assert "self.discover_ollama()" not in SRC[SRC.index("def __init__"):SRC.index("def _build_cfg_tab")]


def test_view_discovery_progress_results_and_confirmation(monkeypatch, tmp_path):
    res = OllamaDiscoveryResult("192.168.1.50", 11434, ["qwen3:4b"], 9)
    gui, a, reg = _app(monkeypatch, tmp_path, [res])
    asked = []
    monkeypatch.setattr(gui.messagebox, "askyesno", lambda *args, **k: asked.append(args) or True)
    FakeClient.models_by_server = {"http://192.168.1.50:11434": ["qwen3:4b"]}
    a.discover_ollama()
    assert a.btn_discover.cfg["state"] == "disabled" and a.btn_discover_stop.cfg["state"] == "normal"
    assert a.lbl_discovery.cfg["text"].startswith("Đang tìm Ollama trong mạng LAN... Đã kiểm tra")
    _finish(a)
    assert a.btn_discover.cfg["state"] == "normal" and a.btn_discover_stop.cfg["state"] == "disabled"
    assert "tìm thấy 1 server" in a.lbl_discovery.cfg["text"]
    # results window opened, server NOT applied yet (even though there is exactly one result)
    assert a.discovery_window is not None and a.discovery_tree is not None
    assert a.ctl.host == "127.0.0.1" and a.var_host.get() == "127.0.0.1"
    values = [v["values"] for v in a.discovery_tree.items.values()]
    assert values == [("192.168.1.50:11434", "qwen3:4b", "Sẵn sàng (9 ms)")]
    # nothing selected → nothing happens
    a.use_discovered_server(a.discovery_window)
    assert a.ctl.host == "127.0.0.1" and asked == []
    # select + confirm → applied + persisted + widgets refreshed
    a.discovery_tree.selection_set(next(iter(a.discovery_tree.items)))
    a.use_discovered_server(a.discovery_window)
    assert len(asked) == 1 and "192.168.1.50:11434" in asked[0][1]
    assert a.ctl.host == "192.168.1.50" and a.var_host.get() == "192.168.1.50" and a.var_port.get() == "11434"
    assert a.var_model.get() == "qwen3:4b" and a.cb_model["values"] == ["qwen3:4b"]
    assert "192.168.1.50:11434" in a.lbl_conn.cfg["text"]
    assert json.loads(AppConfig.path().read_text(encoding="utf-8"))["ollama_server"] == "http://192.168.1.50:11434" \
        if AppConfig.path().exists() else True


def test_view_declined_confirmation_keeps_server(monkeypatch, tmp_path):
    res = OllamaDiscoveryResult("192.168.1.50", 11434, ["qwen3:4b"], 9)
    gui, a, reg = _app(monkeypatch, tmp_path, [res])
    monkeypatch.setattr(gui.messagebox, "askyesno", lambda *args, **k: False)
    a.discover_ollama()
    _finish(a)
    a.discovery_tree.selection_set(next(iter(a.discovery_tree.items)))
    a.use_discovered_server(a.discovery_window)
    assert a.ctl.host == "127.0.0.1"


def test_view_missing_model_warning(monkeypatch, tmp_path):
    res = OllamaDiscoveryResult("192.168.1.60", 11434, ["llama3:8b"], 9)
    gui, a, reg = _app(monkeypatch, tmp_path, [res])
    warnings = []
    monkeypatch.setattr(gui.messagebox, "askyesno", lambda *args, **k: True)
    monkeypatch.setattr(gui.messagebox, "showwarning", lambda *args, **k: warnings.append(args[1]))
    FakeClient.models_by_server = {"http://192.168.1.60:11434": ["llama3:8b"]}
    a.discover_ollama()
    _finish(a)
    a.discovery_tree.selection_set(next(iter(a.discovery_tree.items)))
    a.use_discovered_server(a.discovery_window)
    assert warnings == ["Model qwen3:4b chưa có trên server này.\nVui lòng chọn model có sẵn."]
    assert a.ctl.model == "qwen3:4b" and a.var_model.get() == "qwen3:4b" and a.ctl.host == "192.168.1.60"


def test_view_no_results_shows_spec_message(monkeypatch, tmp_path):
    gui, a, reg = _app(monkeypatch, tmp_path, [])
    infos = []
    monkeypatch.setattr(gui.messagebox, "showinfo", lambda *args, **k: infos.append(args[1]))
    a.discover_ollama()
    _finish(a)
    assert infos and "Không tìm thấy Ollama trong mạng LAN." in infos[-1]
    assert a.discovery_window is None and a.btn_discover_results.cfg["state"] == "disabled"
    assert "không tìm thấy" in a.lbl_discovery.cfg["text"]


def test_view_discovery_disabled_during_batch(monkeypatch, tmp_path):
    gui, a, reg = _app(monkeypatch, tmp_path, [OllamaDiscoveryResult("192.168.1.50")])
    infos = []
    monkeypatch.setattr(gui.messagebox, "showinfo", lambda *args, **k: infos.append(args[1]))
    _running(a)
    assert a.btn_discover.cfg["state"] == "disabled"
    a.discover_ollama()
    assert a.ctl.discovery_runs == 0 and infos and "Đang xử lý" in infos[-1]


def test_view_cancel_button(monkeypatch, tmp_path):
    gate = threading.Event()
    gui, a, reg = _app(monkeypatch, tmp_path, [OllamaDiscoveryResult("192.168.1.50")], total=10, block=gate)
    monkeypatch.setattr(gui.messagebox, "showinfo", lambda *args, **k: None)
    a.discover_ollama()
    for _ in range(200):
        if FakeDiscovery.instances and FakeDiscovery.instances[-1].checked >= 2:
            break
        threading.Event().wait(0.01)
    a.cancel_discovery()
    assert a.btn_discover_stop.cfg["state"] == "disabled" and "dừng" in a.lbl_discovery.cfg["text"].lower()
    gate.set()
    _finish(a)
    assert a.ctl.discovery_results == [] and a.btn_discover.cfg["state"] == "normal"


def test_view_rescan_button_and_labels_present():
    block = SRC[SRC.index("def show_discovery_results"):SRC.index("def _selected_discovery_result")]
    assert 'text="Quét lại"' in block and 'text="Sử dụng server này"' in block
    assert "askyesno" in SRC[SRC.index("def use_discovered_server"):SRC.index("# ------------------------------------------------------------------ run")]


def test_build_artifacts_declare_portable_policy():
    root = Path(app.__file__).resolve().parent.parent
    spec = (root / "ReportExtractor.spec").read_text(encoding="utf-8")
    assert "console=False" in spec and "upx=False" in spec and "exclude_binaries=True" in spec
    assert "app.runtime_paths" in spec and "app.ollama_discovery" in spec and "onefile" not in spec.lower().replace("no onefile", "")
    assert 'PORTABLE_NAME = f"ReportExtractor_v{VERSION}_Portable"' in spec
    assert (root / "build_portable.bat").exists() and (root / "requirements-build.txt").exists()
    assert not (root / "BUILD_PORTABLE.bat").exists() or sys.platform == "win32"
    ignore = (root / ".gitignore").read_text(encoding="utf-8").split()
    assert {"build/", "dist/", "release/", "config/"} <= set(ignore)
    for doc in ("README.txt", "FIRST_RUN.txt", "Install_Ollama_Optional.bat"):
        assert (root / "release_docs" / doc).exists()
    readme = (root / "release_docs" / "README.txt").read_text(encoding="utf-8")
    assert "OLLAMA_HOST=0.0.0.0" in readme and "Program Files" in readme and "192.168.1.50" not in readme
    assert "192.168.1.50" not in (root / "tools" / "build_portable.py").read_text(encoding="utf-8")


def test_build_validator_rejects_dev_files(tmp_path):
    sys.path.insert(0, str(Path(app.__file__).resolve().parent.parent / "tools"))
    import build_portable as bp
    folder = tmp_path / "ReportExtractor_v1.0.3_Portable"
    (folder / "_internal").mkdir(parents=True)
    (folder / "ReportExtractor.exe").write_bytes(b"MZ")
    for d in ("Output", "logs", "config"):
        (folder / d).mkdir()
    for f in ("README.txt", "FIRST_RUN.txt", "VERSION.txt", "Install_Ollama_Optional.bat"):
        (folder / f).write_text("x", encoding="utf-8")
    assert bp.validate_artifact(folder) == []
    (folder / "config" / "config.json").write_text("{}", encoding="utf-8")
    (folder / "tests").mkdir()
    (folder / "sample.pptx").write_bytes(b"x")
    (folder / "_internal" / "ollama.exe").write_bytes(b"x")
    problems = bp.validate_artifact(folder)
    assert any("config.json" in p for p in problems) and any("tests" in p for p in problems)
    assert any("sample.pptx" in p for p in problems) and any("ollama.exe" in p for p in problems)


# ------------------------------------------------------------------ Build 008 blocker: stale autoconnect vs discovery
class GatedLocalClient(FakeClient):
    """Local probe blocks until ``release`` is set – deterministic 'autoconnect finishes AFTER discovery' ordering."""
    release = threading.Event()

    def test_connection(self):
        if self.server.startswith("http://127.0.0.1"):
            GatedLocalClient.release.wait(5)
        return super().test_connection()


def _autoconnect_thread():
    return next((t for t in threading.enumerate() if t.name == "ollama-autoconnect"), None)


def test_stale_autoconnect_does_not_overwrite_confirmed_discovery(tmp_path):
    GatedLocalClient.release = threading.Event()
    FakeClient.models_by_server = {"http://127.0.0.1:11434": ["qwen3:4b", "llama3:8b", "x"],
                                   "http://192.168.1.50:11434": ["qwen3:4b"]}
    ctl, _ = _ctl(tmp_path)
    ctl._client_factory = GatedLocalClient
    ctl.auto_connect_async()                                          # 1. autoconnect starts first (blocked)
    t = _autoconnect_thread()
    assert t is not None and t.is_alive()
    res = OllamaDiscoveryResult("192.168.1.50", 11434, ["qwen3:4b"], 9)
    ok, msg = ctl.apply_discovered_server(res)                        # 2. user confirms the LAN server
    assert ok and ctl.host == "192.168.1.50"
    GatedLocalClient.release.set()                                    # 3. the old autoconnect answers afterwards
    t.join(5)
    assert ctl.host == "192.168.1.50" and ctl.port == 11434           # 4. LAN server remains selected
    assert ctl.server == "http://192.168.1.50:11434" and ctl.ollama_source != "local"
    assert ctl.autoconnect_stale == 1 and ctl.autoconnect_runs == 1
    assert [e.kind for e in ctl.pump()] == []                         # no stale 'autoconnect' event reaches the GUI


def test_typed_server_outranks_pending_autoconnect(tmp_path):
    GatedLocalClient.release = threading.Event()
    FakeClient.models_by_server = {"http://127.0.0.1:11434": ["qwen3:4b"], "http://10.0.0.7:11434": ["qwen3:4b"]}
    ctl, _ = _ctl(tmp_path)
    ctl._client_factory = GatedLocalClient
    ctl.auto_connect_async()
    t = _autoconnect_thread()
    assert ctl.set_endpoint("10.0.0.7", 11434) == ""                   # explicit manual choice while probing
    GatedLocalClient.release.set()
    t.join(5)
    assert ctl.host == "10.0.0.7" and ctl.autoconnect_stale == 1 and ctl.pump() == []


def test_startup_autoconnect_still_selects_local_without_user_choice(tmp_path):
    FakeClient.models_by_server = {"http://127.0.0.1:11434": ["qwen3:4b"]}
    ctl, _ = _ctl(tmp_path)
    ctl.auto_connect_async()
    _autoconnect_thread() and _autoconnect_thread().join(5)
    assert ctl.host == "127.0.0.1" and ctl.ollama_source == "local" and ctl.ollama_ok
    assert ctl.autoconnect_stale == 0 and [e.kind for e in ctl.pump()] == ["autoconnect"]
    # a later, unrelated model edit does not make the epoch move (only host/port changes count)
    e = ctl.endpoint_epoch
    assert ctl.set_endpoint("127.0.0.1", 11434, "qwen3:4b") == "" and ctl.endpoint_epoch == e


def test_view_confirmed_discovery_survives_late_autoconnect(monkeypatch, tmp_path):
    """GUI-level reproduction of the Windows failure: the start-up autoconnect thread finishes after the user
    confirmed the discovered LAN server; the widgets must keep 192.168.1.50."""
    GatedLocalClient.release = threading.Event()
    res = OllamaDiscoveryResult("192.168.1.50", 11434, ["qwen3:4b"], 9)
    gui, a, reg = _app(monkeypatch, tmp_path, [res])
    t0 = _autoconnect_thread()
    if t0 is not None:
        t0.join(5)                                                    # settle the constructor's own autoconnect
    FakeClient.models_by_server = {"http://127.0.0.1:11434": ["qwen3:4b"], "http://192.168.1.50:11434": ["qwen3:4b"]}
    a.ctl._client_factory = GatedLocalClient
    a.ctl.auto_connect_async()                                        # a slow local autoconnect is in flight
    t = _autoconnect_thread()
    monkeypatch.setattr(gui.messagebox, "askyesno", lambda *args, **k: True)
    a.discover_ollama()
    _finish(a)
    a.discovery_tree.selection_set(next(iter(a.discovery_tree.items)))
    a.use_discovered_server(a.discovery_window)
    assert a.ctl.host == "192.168.1.50" and a.var_host.get() == "192.168.1.50"
    GatedLocalClient.release.set()
    t.join(5)
    a._poll()                                                         # drain whatever the worker queued
    assert a.ctl.host == "192.168.1.50" and a.var_host.get() == "192.168.1.50" and a.var_port.get() == "11434"
    assert "192.168.1.50:11434" in a.lbl_conn.cfg["text"] and a.ctl.autoconnect_stale == 1
