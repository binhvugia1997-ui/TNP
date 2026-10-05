"""PROMPT-004A – Ollama local-first connection: 127.0.0.1 → saved server → manual / LAN discovery (never automatic)."""
import json
import threading
from pathlib import Path

import app
from app.config import DEFAULT_MODEL, AppConfig
from app.gui_controller import GuiController
from app.ollama_client import OllamaClient, OllamaError
from tests.test_gui_discovery import FakeDiscovery
from tests.test_gui_redesign import _make_app

LOCAL = "http://127.0.0.1:11434"
LAN = "http://192.168.1.60:11434"
GUI_SRC = Path(app.__file__).with_name("gui.py").read_text(encoding="utf-8")


class ProbeClient:
    """Records every construction; servers not in `up` raise like an unreachable host."""
    up = {}
    calls = []

    def __init__(self, server, timeout=180, probe_timeout=None):
        self.server, self.timeout, self.probe_timeout = server, timeout, probe_timeout
        ProbeClient.calls.append(self)

    def test_connection(self):
        if self.server not in ProbeClient.up:
            raise OllamaError(f"Không kết nối được Ollama tại {self.server}: timed out") from TimeoutError("timed out")
        return {"ok": True, "models": list(ProbeClient.up[self.server])}


def _ctl(tmp_path, server=LAN, model=DEFAULT_MODEL, up=None, timeout=180):
    ProbeClient.up = dict(up or {})
    ProbeClient.calls = []
    FakeDiscovery.instances.clear()
    cfg = AppConfig(ollama_server=server, model=model, period_mode="all", request_timeout=timeout)
    ctl = GuiController(cfg, client_factory=ProbeClient, config_path=tmp_path / "config.json")
    ctl._discovery_factory = lambda: FakeDiscovery([])
    return ctl


# 1. local available + stale saved LAN IP -> local selected and displayed as 127.0.0.1
def test_local_wins_over_stale_saved_lan_ip(tmp_path):
    ctl = _ctl(tmp_path, server=LAN, up={LOCAL: ["qwen3:4b", "llama3"]})
    assert ctl.host == "192.168.1.60"                                  # restored from config
    ok, msg = ctl.auto_connect()
    assert ok and msg == "Ollama local: Sẵn sàng — qwen3:4b"
    assert (ctl.host, ctl.port, ctl.model) == ("127.0.0.1", 11434, "qwen3:4b")
    assert ctl.ollama_source == "local" and ctl.ollama_ok is True and ctl.available_models == ["qwen3:4b", "llama3"]
    assert [c.server for c in ProbeClient.calls] == [LOCAL]           # LAN IP never contacted


def test_localhost_saved_is_normalised_to_127(tmp_path):
    ctl = _ctl(tmp_path, server="localhost:11434", up={LOCAL: ["qwen3:4b"]})
    assert ctl.host == "localhost"
    assert ctl.auto_connect()[0] and ctl.host == "127.0.0.1" and ctl.endpoint_label == "127.0.0.1:11434"


# 2. local available -> no LAN discovery
def test_local_available_never_starts_lan_discovery(tmp_path):
    ctl = _ctl(tmp_path, up={LOCAL: ["qwen3:4b"]})
    ctl.auto_connect()
    assert FakeDiscovery.instances == [] and ctl.discovery_runs == 0 and ctl.discovery_state == "idle"
    src = Path(app.__file__).with_name("gui_controller.py").read_text(encoding="utf-8")
    body = src[src.index("def auto_connect(self)"):src.index("def auto_connect_async")]
    assert "discover" not in body


# 3. local down + saved server up -> saved server kept
def test_saved_server_used_when_local_unavailable(tmp_path):
    ctl = _ctl(tmp_path, server=LAN, up={LAN: ["qwen3:4b"]})
    ok, msg = ctl.auto_connect()
    assert ok and ctl.ollama_source == "saved" and (ctl.host, ctl.port) == ("192.168.1.60", 11434)
    assert "192.168.1.60:11434" in msg and "qwen3:4b" in msg
    assert [c.server for c in ProbeClient.calls] == [LOCAL, LAN]       # local probed FIRST


# 4. both down -> spec message, no automatic LAN scan, controls still editable
def test_neither_available_no_auto_scan(tmp_path):
    ctl = _ctl(tmp_path, server=LAN, up={})
    ok, msg = ctl.auto_connect()
    assert not ok and msg == "Không kết nối được Ollama local hoặc server đã lưu."
    assert ctl.ollama_source == "none" and FakeDiscovery.instances == [] and ctl.discovery_runs == 0
    assert (ctl.host, ctl.port) == ("192.168.1.60", 11434)             # saved value kept for manual editing
    assert ctl.can_discover()                                           # LAN discovery still available on demand
    assert ctl.set_endpoint("192.168.1.61", "11434") == ""             # manual edit still works


# 5. local running without qwen3:4b -> warning, installed models exposed, model NOT switched
def test_local_without_model_warns_and_lists_installed(tmp_path):
    ctl = _ctl(tmp_path, up={LOCAL: ["llama3:8b", "gemma:2b"]})
    ok, msg = ctl.auto_connect()
    assert not ok and msg == "Ollama local đang chạy nhưng không tìm thấy model qwen3:4b"
    assert ctl.model == "qwen3:4b" and ctl.available_models == ["llama3:8b", "gemma:2b"]
    assert (ctl.host, ctl.port) == ("127.0.0.1", 11434) and ctl.ollama_source == "local"
    ctl.set_endpoint("127.0.0.1", 11434, "gemma:2b")                   # user picks an installed model
    assert ctl.check_ollama()[0]


# 6. manual Server/Port still works ("Kiểm tra kết nối" checks the entered endpoint first)
def test_manual_server_check_still_works(tmp_path):
    ctl = _ctl(tmp_path, up={"http://10.0.0.5:11434": ["qwen3:4b"]})
    assert ctl.set_endpoint("10.0.0.5", "11434", "qwen3:4b") == ""
    ok, msg = ctl.check_ollama_local_first()
    assert ok and ctl.host == "10.0.0.5" and "Đã kết nối" in msg
    # dead manual address + local alive -> local detected instead of a timeout dead end
    ProbeClient.up = {LOCAL: ["qwen3:4b"]}
    ctl.set_endpoint("10.0.0.99", "11434", "qwen3:4b")
    ok, msg = ctl.check_ollama_local_first()
    assert ok and ctl.host == "127.0.0.1" and msg == "Ollama local: Sẵn sàng — qwen3:4b"


# 7. saved config still works (Lưu cấu hình persists what is in the controls)
def test_save_config_round_trip(tmp_path):
    ctl = _ctl(tmp_path, up={LOCAL: ["qwen3:4b"]})
    ctl.auto_connect()                                                  # selected local, not yet saved
    assert json.loads((tmp_path / "config.json").read_text()) if (tmp_path / "config.json").exists() else True
    assert ctl.save_ollama_settings() == ""
    data = json.loads((tmp_path / "config.json").read_text(encoding="utf-8"))
    assert data["ollama_server"] == LOCAL and data["model"] == "qwen3:4b" and data["request_timeout"] == 180
    ctl2 = GuiController(AppConfig.load(tmp_path / "config.json"), client_factory=ProbeClient,
                         config_path=tmp_path / "config.json")
    assert (ctl2.host, ctl2.port, ctl2.model) == ("127.0.0.1", 11434, "qwen3:4b")
    # saved LAN server survives an auto-connect that picked local (fallback for later), until the user saves
    ctl3 = _ctl(tmp_path / "b", server=LAN, up={LOCAL: ["qwen3:4b"]})
    ctl3.auto_connect()
    assert ctl3.cfg.ollama_server == LAN


# 8. probes use the short timeout, 9. inference timeout unchanged
def test_probe_uses_short_timeout_and_inference_timeout_unchanged(tmp_path):
    ctl = _ctl(tmp_path, up={}, timeout=180)
    ctl.auto_connect()
    assert ProbeClient.calls and all(c.probe_timeout == 3.0 for c in ProbeClient.calls)
    assert all(c.timeout == 180 for c in ProbeClient.calls)            # inference timeout passed through unchanged
    assert ctl.probe_timeout == 3.0 and ctl.cfg.request_timeout == 180
    ctl.check_ollama(timeout=15)
    assert ProbeClient.calls[-1].probe_timeout == 3.0 and ProbeClient.calls[-1].timeout == 180
    ctl.cfg.probe_timeout = 0                                           # bad config value -> default, never 180
    assert ctl.probe_timeout == 3.0
    ctl.cfg.probe_timeout = 999
    assert ctl.probe_timeout == 15.0


def test_client_separates_probe_and_inference_timeouts(monkeypatch):
    seen = {}

    class R:
        def raise_for_status(self): pass
        def json(self): return {"models": [{"name": "qwen3:4b"}]}

    monkeypatch.setattr("app.ollama_client.requests.get", lambda url, timeout=None: (seen.setdefault("get", timeout), R())[1])
    c = OllamaClient(LOCAL, timeout=180, probe_timeout=3)
    assert c.test_connection()["models"] == ["qwen3:4b"] and seen["get"] == 3.0 and c.timeout == 180
    seen.clear()
    c2 = OllamaClient(LOCAL, timeout=180)                               # legacy ctor: probe capped, inference 180
    c2.list_models()
    assert seen["get"] == OllamaClient.PROBE_TIMEOUT_CAP and c2.timeout == 180
    assert c2.build_generate_payload("qwen3:4b", "p")["think"] is False  # generation payload untouched


def test_config_has_separate_probe_timeout_field():
    cfg = AppConfig()
    assert cfg.request_timeout == 180 and cfg.probe_timeout == 3
    assert "probe_timeout" in AppConfig.__dataclass_fields__


# ------------------------------------------------------------------ GUI (fake tkinter)
def _wait_autoconnect(a):
    for t in threading.enumerate():
        if t.name == "ollama-autoconnect":
            t.join(5)
    a._poll()


def test_gui_startup_selects_local_and_updates_controls(monkeypatch, tmp_path):
    ProbeClient.up = {LOCAL: ["qwen3:4b"]}
    ProbeClient.calls = []
    FakeDiscovery.instances.clear()
    gui, a, reg, _ = _make_app(monkeypatch, tmp_path, 1, cfg_extra=None)
    _wait_autoconnect(a)                                                # drain the start-up probe (real client)
    a.ctl._client_factory = ProbeClient
    a.ctl.cfg.ollama_server = LAN
    a.ctl.server = LAN
    a.var_host.set("192.168.1.60")
    a.ctl.auto_connect_async()
    _wait_autoconnect(a)
    assert a.var_host.get() == "127.0.0.1" and a.var_port.get() == "11434" and a.var_model.get() == "qwen3:4b"
    assert a.lbl_conn.cfg["text"] == "Ollama local: Sẵn sàng — qwen3:4b" and a.lbl_conn.cfg["style"] == "Success.TLabel"
    assert a.lbl_ai_run.cfg["text"] == "● Ollama: qwen3:4b — Sẵn sàng"
    assert a.cb_model.cfg["values"] == ["qwen3:4b"]
    assert FakeDiscovery.instances == [] and a.btn_discover.cfg["text"] == "Tìm Ollama trong mạng LAN"
    assert "self.ctl.auto_connect_async()" in GUI_SRC[GUI_SRC.index("class ReportExtractorApp"):GUI_SRC.index("def _build")]
    # user may still change the server afterwards
    a.var_host.set("192.168.1.70")
    assert a.ctl.ollama_ok is None and a.lbl_conn.cfg["text"].startswith("● Chưa kiểm tra kết nối")


def test_gui_startup_local_without_model_shows_warning_and_models(monkeypatch, tmp_path):
    ProbeClient.up = {LOCAL: ["llama3:8b"]}
    gui, a, reg, _ = _make_app(monkeypatch, tmp_path, 1)
    _wait_autoconnect(a)                                                # drain the start-up probe (real client)
    a.ctl._client_factory = ProbeClient
    a.ctl.auto_connect_async()
    _wait_autoconnect(a)
    assert a.lbl_conn.cfg["text"] == "Ollama local đang chạy nhưng không tìm thấy model qwen3:4b"
    assert a.lbl_conn.cfg["style"] == "Warning.TLabel" and a.cb_model.cfg["values"] == ["llama3:8b"]
    assert a.var_model.get() == "qwen3:4b"                              # not silently switched
    assert "llama3:8b" in "".join(a.txt_log.lines)


def test_gui_startup_nothing_reachable_message(monkeypatch, tmp_path):
    ProbeClient.up = {}
    gui, a, reg, _ = _make_app(monkeypatch, tmp_path, 1)
    _wait_autoconnect(a)                                                # drain the start-up probe (real client)
    a.ctl._client_factory = ProbeClient
    a.ctl.auto_connect_async()
    _wait_autoconnect(a)
    assert a.lbl_conn.cfg["text"] == "Không kết nối được Ollama local hoặc server đã lưu."
    assert a.lbl_conn.cfg["style"] == "Error.TLabel" and a.btn_discover.cfg.get("state") != "disabled"
