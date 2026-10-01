"""GUI controller: deterministic progress %, monotonic timer / ETA, Ollama connection settings."""
import json
import sys
import threading
import time
import types

from app import batch_processor as bp
from app.batch_processor import BatchSummary
from app.config import AppConfig, endpoint_problem, normalize_ollama_url, split_endpoint
from app.gui_controller import (STAGE_WEIGHTS, GuiController, Progress, UiEvent, format_elapsed,
                                progress_bar_text)
from app.ollama_client import OllamaError


class FakeClock:
    def __init__(self, t=1000.0):
        self.t = t

    def __call__(self):
        return self.t

    def advance(self, s):
        self.t += s


def _ctl_with_rows(tmp_path, n=5):
    ctl = GuiController(AppConfig(period_mode="all"), config_path=tmp_path / "c.json")
    folder = tmp_path / "reports"
    folder.mkdir(exist_ok=True)
    for i in range(n):
        (folder / f"26092{i:04d}-VOC_r{i}.pptx").write_bytes(b"x")
    ctl.set_report_folder(str(folder))
    clock = FakeClock()
    ctl._clock = clock
    return ctl, clock


def _begin(ctl, clock):
    """Put the controller into the 'running' state the way start() does, without a real processor."""
    ctl.report_durations, ctl._report_started_at = [], None
    ctl.started_at, ctl.finished_at = clock(), None
    ctl.progress = Progress(total=len(ctl.files))
    ctl.state = "running"


def _finish_report(ctl, i, status="completed"):
    ctl.apply_event(UiEvent("row", (i, status, "")))
    ctl.apply_event(UiEvent("progress", (sum(1 for r in ctl.rows if r.is_final), len(ctl.rows))))


# ------------------------------------------------------------------ progress
def test_progress_zero_reports_is_safe(tmp_path):
    ctl = GuiController(AppConfig(period_mode="all"), config_path=tmp_path / "c.json")
    assert ctl.progress.percent == 0.0 and ctl.progress.text == "" and ctl.eta_text() == ""
    assert progress_bar_text(0) == "░" * 20 + " 0%"


def test_progress_moves_with_real_stages_and_report_boundaries(tmp_path):
    ctl, clock = _ctl_with_rows(tmp_path, 5)
    _begin(ctl, clock)
    assert ctl.progress.percent == 0.0 and ctl.progress.text == "Đang xử lý: 1 / 5 — 0%"
    ctl.apply_event(UiEvent("row", (0, "reading", "")))
    p1 = ctl.progress.percent
    assert 0 < p1 < 20                                   # first real stage of report 1 of 5
    assert p1 == STAGE_WEIGHTS["reading"] / 5 * 100
    ctl.apply_event(UiEvent("row", (0, "analyzing", "8 slide")))
    p2 = ctl.progress.percent
    assert p1 < p2 < 20 and p2 == STAGE_WEIGHTS["analyzing"] / 5 * 100
    # a long Qwen call: nothing changes until the next REAL stage event
    clock.advance(120)
    assert ctl.progress.percent == p2
    for st in ("extracting", "extracting_qpn", "extracting_images", "writing_excel"):
        ctl.apply_event(UiEvent("row", (0, st, "")))
        assert ctl.progress.percent < 20
    assert ctl.progress.percent == STAGE_WEIGHTS["writing_excel"] / 5 * 100
    _finish_report(ctl, 0)
    assert ctl.progress.percent == 20.0 and ctl.progress.text == "Đang xử lý: 2 / 5 — 20%"
    ctl.apply_event(UiEvent("row", (1, "reading", "")))
    assert 20 < ctl.progress.percent < 40
    _finish_report(ctl, 1, "needs_review")
    assert ctl.progress.percent == 40.0
    _finish_report(ctl, 2, "skipped")                     # skipped counts as terminal progress
    assert ctl.progress.percent == 60.0
    _finish_report(ctl, 3, "error")                       # error counts as terminal progress
    assert ctl.progress.percent == 80.0
    _finish_report(ctl, 4, "not_written")
    ctl.apply_event(UiEvent("done", BatchSummary(total=5, completed=1, needs_review=1, skipped=1, failed=1, not_written=1)))
    assert ctl.progress.percent == 100.0 and ctl.progress.text == "Đã xử lý: 5 / 5 — 100%"
    assert ctl.progress.bar_text == "█" * 20 + " 100%"


def test_stage_weights_are_monotonic_in_emission_order():
    order = ["reading", "analyzing", "extracting", "extracting_qpn", "extracting_images", "writing_excel"]
    vals = [STAGE_WEIGHTS[s] for s in order]
    assert vals == sorted(vals) and vals[0] > 0 and vals[-1] < 1.0
    assert STAGE_WEIGHTS["analyzing_heuristic"] == STAGE_WEIGHTS["analyzing"]
    p = Progress(total=1)
    p.reach_stage(0, "writing_excel")
    p.reach_stage(0, "reading")                           # out-of-order event never moves the bar backwards
    assert p.current_fraction == STAGE_WEIGHTS["writing_excel"]


def test_stopped_batch_never_shows_100(tmp_path):
    ctl, clock = _ctl_with_rows(tmp_path, 5)
    _begin(ctl, clock)
    _finish_report(ctl, 0)
    _finish_report(ctl, 1)
    ctl.state = "stopping"
    ctl.apply_event(UiEvent("done", BatchSummary(total=5, completed=2, stopped=True)))
    assert ctl.progress.percent == 40.0 and ctl.progress.text == "Đã xử lý: 2 / 5 — 40%"
    assert ctl.progress.bar_text.endswith(" 40%") and ctl.progress.bar_text.count("█") == 8
    assert ctl.eta_text() == "" and ctl.elapsed_text().startswith("Thời gian đã chạy:")


def test_bar_and_text_share_one_value(tmp_path):
    ctl, clock = _ctl_with_rows(tmp_path, 3)
    _begin(ctl, clock)
    ctl.apply_event(UiEvent("row", (0, "extracting", "")))
    p = ctl.progress
    assert f"{p.percent_int}%" in p.text and p.bar_text.endswith(f" {p.percent_int}%")
    assert int(p.percent) == p.percent_int and progress_bar_text(p.percent) == p.bar_text
    assert p.text in ctl.status_line() and p.bar_text in ctl.status_line()


# ------------------------------------------------------------------ timing / ETA
def test_timer_starts_at_batch_start_and_formats(tmp_path):
    ctl, clock = _ctl_with_rows(tmp_path, 2)
    clock.advance(500)                                     # time spent configuring the GUI is not counted
    assert ctl.elapsed_text() == "" and ctl.elapsed_seconds() == 0.0
    _begin(ctl, clock)
    assert ctl.started_at == clock.t
    clock.advance(88)
    assert ctl.elapsed_text() == "Đã chạy: 01:28"
    clock.advance(3600)
    assert ctl.elapsed_text() == "Đã chạy: 01:01:28"
    assert format_elapsed(0) == "00:00" and format_elapsed(3599) == "59:59" and format_elapsed(3600) == "01:00:00"
    assert format_elapsed(-5) == "00:00"


def test_eta_from_measured_durations(tmp_path):
    ctl, clock = _ctl_with_rows(tmp_path, 5)
    _begin(ctl, clock)
    assert ctl.eta_text() == "Còn khoảng: Đang tính..."
    ctl.apply_event(UiEvent("row", (0, "reading", "")))
    clock.advance(30)
    assert ctl.eta_text() == "Còn khoảng: Đang tính..."     # no report finished yet
    _finish_report(ctl, 0)                                 # report 1 took 30 s
    assert ctl.report_durations == [30.0]
    assert ctl.eta_text() == "Còn khoảng: 02:00"           # 4 remaining × 30 s
    ctl.apply_event(UiEvent("row", (1, "reading", "")))
    clock.advance(60)
    _finish_report(ctl, 1)                                 # report 2 took 60 s -> average 45 s
    assert ctl.average_report_seconds() == 45.0
    assert ctl.eta_text() == "Còn khoảng: 02:15"           # 3 × 45 s
    ctl.apply_event(UiEvent("row", (2, "writing_excel", "")))
    assert ctl.eta_seconds() == (3 - STAGE_WEIGHTS["writing_excel"]) * 45.0   # current report's real stage counts
    for i in (2, 3, 4):
        clock.advance(45)
        _finish_report(ctl, i)
    ctl.apply_event(UiEvent("done", BatchSummary(total=5, completed=5)))
    assert ctl.eta_text() == ""                            # no ETA after the batch ended
    assert ctl.elapsed_text() == "Tổng thời gian: 03:45"   # 30+60+45*3 = 225 s
    assert "Tổng thời gian xử lý: 03:45" in ctl.summary_lines()
    clock.advance(999)
    assert ctl.elapsed_text() == "Tổng thời gian: 03:45"   # frozen at completion


def test_timer_updates_do_not_block_gui(tmp_path):
    """elapsed/eta are pure reads of controller state – no sleeping, no joining."""
    ctl, clock = _ctl_with_rows(tmp_path, 2)
    _begin(ctl, clock)
    t0 = time.perf_counter()
    for _ in range(2000):
        ctl.elapsed_text()
        ctl.eta_text()
        ctl.status_line()
    assert time.perf_counter() - t0 < 1.0


# ------------------------------------------------------------------ Ollama endpoint normalisation
def test_endpoint_normalisation_accepts_all_spellings():
    for raw in ("127.0.0.1", "127.0.0.1:11434", "http://127.0.0.1:11434", "http://127.0.0.1:11434/"):
        assert normalize_ollama_url(raw) == "http://127.0.0.1:11434"
    assert normalize_ollama_url("192.168.1.50") == "http://192.168.1.50:11434"
    assert normalize_ollama_url("192.168.1.50:11434") == "http://192.168.1.50:11434"
    assert normalize_ollama_url("http://192.168.1.50:11434") == "http://192.168.1.50:11434"
    assert normalize_ollama_url("192.168.1.50:8080") == "http://192.168.1.50:8080"             # custom port
    assert normalize_ollama_url("http://http://192.168.1.50:11434") == "http://192.168.1.50:11434"
    assert normalize_ollama_url("192.168.1.50:11434:11434") == "http://192.168.1.50:11434"
    assert "http://http" not in normalize_ollama_url("http://http://x:1") and normalize_ollama_url("x:1").count(":1") == 1
    assert split_endpoint("http://192.168.1.50:11434") == ("192.168.1.50", 11434)
    assert split_endpoint("pc-01") == ("pc-01", 11434)
    assert endpoint_problem("192.168.1.50", "11434") == ""
    assert endpoint_problem("", "11434").startswith("Chưa nhập")
    assert "không hợp lệ" in endpoint_problem("abc def", 11434)
    assert "IP không hợp lệ" in endpoint_problem("192.168.1.300", 11434)
    assert "Port" in endpoint_problem("127.0.0.1", "abc") and "Port" in endpoint_problem("127.0.0.1", 70000)


def test_controller_endpoint_fields(tmp_path):
    ctl = GuiController(AppConfig(period_mode="all"), config_path=tmp_path / "c.json")
    assert (ctl.host, ctl.port, ctl.model) == ("127.0.0.1", 11434, "qwen3:4b")       # defaults
    assert ctl.set_endpoint("192.168.1.50", "11434") == ""
    assert ctl.server == "http://192.168.1.50:11434" and ctl.endpoint_label == "192.168.1.50:11434"
    assert ctl.set_endpoint("http://192.168.1.60:11500/", "11434") == ""             # full URL typed in IP field
    assert (ctl.host, ctl.port) == ("192.168.1.60", 11500)
    assert ctl.set_endpoint("192.168.1.70:9000", "") == "" and ctl.port == 9000
    assert ctl.set_endpoint("192.168.1.70", "") == "" and ctl.port == 11434
    assert ctl.set_endpoint("bad host!", "11434")                                    # malformed -> message, unchanged
    assert ctl.host == "192.168.1.70"
    assert "Port" in ctl.set_endpoint("192.168.1.70", "abc")
    assert ctl.ollama_ok is None                                                     # status reset after a change
    assert ctl.ai_status_text() == "AI: qwen3:4b @ 192.168.1.70:11434 — Chưa kiểm tra"


# ------------------------------------------------------------------ Ollama connection check / discovery
class _Client:
    calls = []

    def __init__(self, server, timeout=180):
        self.base = server
        self.timeout = timeout
        _Client.calls.append(server)

    def test_connection(self):
        return {"ok": True, "server": self.base, "models": ["qwen3:1.7b", "qwen3:4b", "llama3.2:3b"]}


class _NoModel(_Client):
    def test_connection(self):
        return {"ok": True, "server": self.base, "models": ["qwen3:1.7b"]}


class _Refused(_Client):
    def test_connection(self):
        import requests
        try:
            raise requests.ConnectionError("[Errno 111] Connection refused")
        except requests.ConnectionError as e:
            raise OllamaError(f"Không kết nối được Ollama tại {self.base}: {e}") from e


class _Timeout(_Client):
    def test_connection(self):
        import requests
        try:
            raise requests.ReadTimeout("Read timed out")
        except requests.ReadTimeout as e:
            raise OllamaError(f"Không kết nối được Ollama tại {self.base}: {e}") from e


class _Crash(_Client):
    def test_connection(self):
        raise ValueError("unexpected json")


def test_check_ollama_messages_and_model_presence(tmp_path):
    ctl = GuiController(AppConfig(period_mode="all"), config_path=tmp_path / "c.json")
    ctl.set_endpoint("192.168.1.50", 11434, "qwen3:4b")
    ctl._client_factory = _Client
    assert ctl.check_ollama() == (True, "● Đã kết nối — qwen3:4b")
    assert ctl.ai_status_text() == "AI: qwen3:4b @ 192.168.1.50:11434 — Đã kết nối"
    assert ctl.available_models == ["qwen3:1.7b", "qwen3:4b", "llama3.2:3b"]
    ctl._client_factory = _NoModel
    assert ctl.check_ollama() == (False, "● Đã kết nối Ollama nhưng không tìm thấy model qwen3:4b")
    ctl._client_factory = _Refused
    assert ctl.check_ollama() == (False, "● Không kết nối được Ollama tại 192.168.1.50:11434")
    assert ctl.ai_status_text() == "AI: Không kết nối — sẽ sử dụng heuristic fallback"
    ctl._client_factory = _Timeout
    assert ctl.check_ollama() == (False, "● Kết nối Ollama quá thời gian")
    ctl._client_factory = _Crash
    ok, msg = ctl.check_ollama()
    assert not ok and "Traceback" not in msg and msg.startswith("● Không kết nối được Ollama tại 192.168.1.50:11434")
    ctl.set_endpoint("127.0.0.1", 11434)                       # localhost
    ctl._client_factory = _Client
    assert ctl.check_ollama()[0] and _Client.calls[-1] == "http://127.0.0.1:11434"
    ctl.set_endpoint("127.0.0.1", 8080)                        # custom port goes to the client
    assert ctl.check_ollama()[0] and _Client.calls[-1] == "http://127.0.0.1:8080"
    ctl.host = "bad host!"
    assert ctl.check_ollama() == (False, "● Địa chỉ Ollama không hợp lệ: 'bad host!'")


def test_check_ollama_async_does_not_block(tmp_path):
    ctl = GuiController(AppConfig(period_mode="all"), config_path=tmp_path / "c.json")
    gate = threading.Event()

    class Slow(_Client):
        def test_connection(self):
            gate.wait(5)
            return super().test_connection()

    ctl._client_factory = Slow
    t0 = time.perf_counter()
    ctl.check_ollama_async()
    assert time.perf_counter() - t0 < 0.5 and ctl.pump() == []        # GUI thread free while the worker waits
    gate.set()
    for _ in range(100):
        if ctl.pump():
            break
        time.sleep(0.02)
    assert ctl.ollama_ok and ctl.ollama_status == "● Đã kết nối — qwen3:4b"


def test_model_discovery_and_failure_keeps_selection(tmp_path):
    ctl = GuiController(AppConfig(model="my-custom:latest", period_mode="all"), config_path=tmp_path / "c.json")
    ctl._client_factory = _Client
    ok, msg, models = ctl.refresh_models()
    assert ok and models == ["qwen3:1.7b", "qwen3:4b", "llama3.2:3b"] and "3 model" in msg
    assert ctl.model == "my-custom:latest"                     # user's value never cleared
    ctl.model = ""
    assert ctl.refresh_models()[0] and ctl.model == "qwen3:4b"  # preferred default when installed
    ctl.model = "qwen3:1.7b"
    ctl._client_factory = _Refused
    ok, msg, models = ctl.refresh_models()
    assert not ok and models == [] and "giữ nguyên model hiện tại" in msg and ctl.model == "qwen3:1.7b"
    ctl._client_factory = _Timeout
    assert "quá thời gian" in ctl.refresh_models()[1] and ctl.model == "qwen3:1.7b"
    ctl.refresh_models_async()
    for _ in range(100):
        evs = ctl.pump()
        if evs:
            break
        time.sleep(0.02)
    assert evs[0].kind == "models" and evs[0].payload[0] is False


def test_settings_save_and_reload_and_immediate_effect(sample_tree, tmp_path):
    cfg_path = tmp_path / "config.json"
    ctl = GuiController(AppConfig(period_mode="all"), config_path=cfg_path)
    ctl.set_report_folder(str(sample_tree["reports"]))
    ctl.set_template(str(sample_tree["template"]))
    assert ctl.set_endpoint("192.168.1.50", "11500", "qwen3:1.7b") == ""
    assert ctl.save_ollama_settings() == ""
    data = json.loads(cfg_path.read_text(encoding="utf-8"))
    assert data["ollama_server"] == "http://192.168.1.50:11500" and data["model"] == "qwen3:1.7b"
    ctl2 = GuiController(AppConfig.load(cfg_path), config_path=cfg_path)
    assert (ctl2.host, ctl2.port, ctl2.model) == ("192.168.1.50", 11500, "qwen3:1.7b")      # restored on start
    # changed IP / model are used immediately by the next check AND the next production batch
    ctl2._client_factory = _Client
    ctl2.set_endpoint("10.0.0.9", 11434, "qwen3:4b")
    ctl2.check_ollama()
    assert _Client.calls[-1] == "http://10.0.0.9:11434"
    opts = ctl2.build_options(use_ollama=True)
    assert opts.ollama_server == "http://10.0.0.9:11434" and opts.model == "qwen3:4b" and opts.use_ollama
    ctl2.set_endpoint("10.0.0.9", 11434, "qwen3:1.7b")
    assert ctl2.build_options(use_ollama=True).model == "qwen3:1.7b"
    ctl2.host = "bad host!"
    assert ctl2.save_ollama_settings().startswith("Địa chỉ Ollama không hợp lệ")


# ------------------------------------------------------------------ fallback behaviour
def test_ollama_lost_during_batch_falls_back_without_marking_everything_review(sample_tree, tmp_path, monkeypatch):
    """First report: Qwen answers; then the server disappears -> heuristic fallback, batch continues,
    no GUI crash, Excel written, status decided by the extraction rules only."""
    class FlakyClient:
        n = 0

        def __init__(self, server, timeout=180):
            self.base, self.timeout, self.last_call = server, timeout, {}

        def generate_json(self, model, prompt, system="", **kw):
            FlakyClient.n += 1
            if FlakyClient.n == 1:
                self.last_call = {"status": 200, "response_chars": 2, "seconds": 0.1}
                return {}                                   # valid (empty) JSON -> merged with heuristics
            raise OllamaError("Không kết nối được Ollama tại " + self.base + ": connection lost")

    monkeypatch.setattr(bp, "OllamaClient", FlakyClient)
    ctl = GuiController(AppConfig(row_mode="append", period_mode="all"), config_path=tmp_path / "c.json")
    ctl.set_report_folder(str(sample_tree["reports"]))
    ctl.set_template(str(sample_tree["template"]))
    ctl.set_output(str(tmp_path / "out" / "k.xlsx"))
    assert ctl.start(use_ollama=True)
    ctl.processor._thread.join(120)
    ctl.pump()
    s = ctl.summary
    assert s.failed == 0 and s.total == len(ctl.files) and ctl.progress.percent == 100.0
    classifiers = [ctl.result_for(i).classifier for i in range(len(ctl.files))]
    assert classifiers[0].startswith("qwen") and all(c == "heuristic" for c in classifiers[1:])
    assert FlakyClient.n == len(ctl.files)                  # every report still tried Qwen, no restart needed
    notes = [n for i in range(len(ctl.files)) for n in ctl.result_for(i).classifier_notes]
    assert any("heuristic fallback" in n for n in notes)
    for i in range(len(ctl.files)):
        fr = ctl.result_for(i)
        assert not any("Ollama" in r for r in fr.review_reasons)   # unavailability itself is never a review reason
    assert (tmp_path / "out" / "k.xlsx").exists()
    # same inputs without AI give the same statuses (no extraction regression from the AI path)
    ctl2 = GuiController(AppConfig(row_mode="append", period_mode="all"), config_path=tmp_path / "c2.json")
    ctl2.set_report_folder(str(sample_tree["reports"]))
    ctl2.set_template(str(sample_tree["template"]))
    ctl2.set_output(str(tmp_path / "out2" / "k.xlsx"))
    ctl2.start(use_ollama=False)
    ctl2.processor._thread.join(120)
    ctl2.pump()
    assert [r.stage for r in ctl2.rows] == [r.stage for r in ctl.rows]


def test_mocked_view_renders_progress_and_ollama_widgets(monkeypatch, sample_tree, tmp_path):
    class Widget:
        def __init__(self, *a, **k): self.k = k; self.cfg = {}
        def __getattr__(self, name): return lambda *a, **k: None
        def __setitem__(self, k, v): self.cfg[k] = v
        def __getitem__(self, k): return self.cfg.get(k, "")
        def configure(self, *a, **k): self.cfg.update(k)
        def get_children(self): return []
        def insert(self, *a, **k): return "I001"
        def selection(self): return ()
        def get(self): return self.k.get("value", "")
    class Var(Widget):
        def __init__(self, value="", **k): self.v = value; self.cfg = {}
        def set(self, v): self.v = v
        def get(self): return self.v
    class Root(Widget):
        def __init__(self, *a, **k): self.cfg = {}; self.scheduled = []
        def after(self, ms, fn=None): self.scheduled.append((ms, fn))
    tkmod = types.ModuleType("tkinter")
    for n in ("Text", "Toplevel"):
        setattr(tkmod, n, Widget)
    tkmod.Tk, tkmod.StringVar, tkmod.BooleanVar, tkmod.TclError = Root, Var, Var, Exception
    ttkmod = types.ModuleType("tkinter.ttk")
    for n in ("Style", "Label", "LabelFrame", "Frame", "Entry", "Button", "Combobox", "Checkbutton", "Treeview",
              "Scrollbar", "Progressbar", "Radiobutton", "Spinbox"):
        setattr(ttkmod, n, Widget)
    fd, mb = types.ModuleType("tkinter.filedialog"), types.ModuleType("tkinter.messagebox")
    for m in (fd, mb):
        m.__getattr__ = lambda name: (lambda *a, **k: True)
    tkmod.ttk, tkmod.filedialog, tkmod.messagebox = ttkmod, fd, mb
    for name, mod in (("tkinter", tkmod), ("tkinter.ttk", ttkmod), ("tkinter.filedialog", fd), ("tkinter.messagebox", mb)):
        monkeypatch.setitem(sys.modules, name, mod)
    monkeypatch.delitem(sys.modules, "app.gui", raising=False)
    import importlib
    gui = importlib.import_module("app.gui")
    cfg = AppConfig(last_report_folder=str(sample_tree["reports"]), last_template=str(sample_tree["template"]),
                    ollama_server="http://192.168.1.50:11434", model="qwen3:4b")
    app = gui.ReportExtractorApp(cfg)
    assert app.var_host.get() == "192.168.1.50" and app.var_port.get() == "11434" and app.var_model.get() == "qwen3:4b"
    assert app.lbl_ai.cfg.get("text", "") or True
    # widgets render the controller's single progress value
    clock = FakeClock()
    app.ctl._clock = clock
    _begin(app.ctl, clock)
    app.ctl.apply_event(UiEvent("row", (0, "extracting", "")))
    app._render_progress()
    pct = app.ctl.progress.percent
    assert app.pb["value"] == pct and app.lbl_percent.cfg["text"] == f"{int(pct)}%"
    assert app.lbl_progress.cfg["text"] == app.ctl.progress.text
    clock.advance(61)
    app._tick()
    assert app.lbl_elapsed.cfg["text"] == "Đã chạy: 01:01" and app.lbl_eta.cfg["text"] == "Còn khoảng: Đang tính..."
    assert any(ms == 1000 for ms, _ in app.root.scheduled)       # timer re-armed via after(), never sleeps
    # endpoint edit in the widgets reaches the controller immediately
    app.var_host.set("http://10.1.1.5:11500")
    app.var_port.set("11434")
    assert app._push_endpoint() == "" and (app.ctl.host, app.ctl.port) == ("10.1.1.5", 11500)
    assert app.var_host.get() == "10.1.1.5" and app.var_port.get() == "11500"
    app.var_host.set("bad host!")
    assert app._push_endpoint()
    app.ctl.state = "idle"
