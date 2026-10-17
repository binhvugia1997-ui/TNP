"""PROMPT-040 — the Excel “File kiểm chứng / File kết quả” selection must be persisted once and reused.

Regression for the repeated Excel picker: picking a file returned the chosen path to the browser but
nothing wrote it into the existing user config, so a restart started with empty Excel fields and the
user had to open the native picker again.  These tests drive the production objects
(``BridgeService`` → ``ApplicationService`` → ``GuiController`` → ``AppConfig``) with a scripted native
dialog window and the real ``config.json`` persistence.  They are Linux development evidence only: no
Windows dialog, WebView2, PowerPoint COM or Portable acceptance is claimed here.
"""
from __future__ import annotations

import json
import os
import shutil
import time
from pathlib import Path

from app.application_service import ApplicationService
from app.batch_processor import BatchSummary
from app.bridge import BridgeService
from app.config import AppConfig
from app.gui_controller import GuiController
from app.runtime_paths import learning_dir

EXCEL_STATES = {"unconfigured", "checking", "ok", "missing", "inaccessible", "invalid"}


class DialogWindow:
    """Scripted stand-in for the pywebview window: each call consumes one prepared native choice."""

    def __init__(self, *choices):
        self.choices = list(choices)
        self.calls = []

    def create_file_dialog(self, kind, **kwargs):
        self.calls.append((kind, kwargs))
        if not self.choices:
            raise AssertionError(f"unexpected native dialog call: {kind} {kwargs}")
        return self.choices.pop(0)


def make_service(tmp_path: Path, *, config_path: Path | None = None, processor_factory=None,
                 template: str = "", output: str = "") -> ApplicationService:
    config_path = Path(config_path) if config_path else tmp_path / "config" / "config.json"
    cfg = AppConfig(last_template=template, last_output_file=output)
    extra = {"processor_factory": processor_factory} if processor_factory else {}
    controller = GuiController(cfg, config_path=config_path, learning_dir_override=tmp_path / "learning_data",
                               **extra)
    return ApplicationService(controller, config_path=config_path,
                              manual_fields_path=config_path.parent / "manual_fields.json")


def restarted_service(tmp_path: Path, config_path: Path, processor_factory=None) -> ApplicationService:
    """A fresh process over the same config file — exactly what an application restart does."""
    config_path = Path(config_path)
    extra = {"processor_factory": processor_factory} if processor_factory else {}
    controller = GuiController(config_path=config_path, learning_dir_override=tmp_path / "learning_data",
                               **extra)
    return ApplicationService(controller, config_path=config_path,
                              manual_fields_path=config_path.parent / "manual_fields.json")


def bind_bridge(monkeypatch, service: ApplicationService, window: DialogWindow | None) -> BridgeService:
    bridge = BridgeService(service)
    if window is not None:
        bridge.bind_window(window)
    monkeypatch.setattr(bridge, "_webview_constant", lambda name: name)
    return bridge


def excel_state(service: ApplicationService) -> dict:
    """State DTO with the one-shot deep template check finished (a join, never a sleep)."""
    service.excel_selection()                                   # schedules the check for the current selection
    thread = getattr(service, "_excel_check_thread", None)
    if thread is not None:
        thread.join(30)
    return service.excel_selection()


def config_dto(folder: Path, template: Path, output: Path) -> dict:
    return {
        "paths": {"reportFolder": str(folder), "template": str(template), "output": str(output)},
        "period": {"mode": "all", "month": "", "year": "", "from": "", "to": ""},
        "forceReprocess": False,
        "ollama": {"host": "127.0.0.1", "port": "11434", "model": "qwen3:4b"},
        "update": {"path": "", "autoCheck": False},
    }


class _RecordingProcessor:
    """Synchronous worker stand-in: records the BatchOptions the controller really built."""

    instances: list = []

    def __init__(self, options, on_prescan=None, on_file=None, on_batch=None, on_done=None, on_log=None):
        self.options = options
        self.on_done = on_done
        self.summary = None
        self.__class__.instances.append(self)

    def start(self) -> None:
        self.summary = BatchSummary(total=len(self.options.files), completed=len(self.options.files))
        if self.on_done:
            self.on_done(self.summary)

    def is_running(self) -> bool:
        return False

    def request_stop(self) -> bool:
        return True


def pick_both(monkeypatch, tmp_path: Path, *, template: Path, output: Path,
              config_path: Path | None = None) -> tuple[ApplicationService, BridgeService, DialogWindow]:
    service = make_service(tmp_path, config_path=config_path)
    window = DialogWindow([str(template)], [str(output)])
    bridge = bind_bridge(monkeypatch, service, window)
    assert bridge.choose_template_file()["ok"] is True
    assert bridge.choose_output_file()["ok"] is True
    return service, bridge, window


# --------------------------------------------------------------------------- 1. pick → config persist
def test_excel_pick_is_written_into_the_existing_config_file(sample_tree, tmp_path, monkeypatch):
    config_path = tmp_path / "config" / "config.json"
    template = Path(sample_tree["template"]).resolve()
    output = (tmp_path / "Output" / "Kiem chung 09-2026.xlsx").resolve()
    service, _bridge, _window = pick_both(monkeypatch, tmp_path, template=template, output=output,
                                          config_path=config_path)

    assert config_path.is_file()
    stored = json.loads(config_path.read_text(encoding="utf-8"))
    assert stored["last_template"] == str(template)
    assert stored["last_output_file"] == str(output)
    assert service.controller.template == str(template)
    assert service.controller.output == str(output)
    assert not list(config_path.parent.glob("*.tmp"))


def test_pick_response_shape_is_unchanged(sample_tree, tmp_path, monkeypatch):
    """The bridge contract the frontend already consumes must not change shape."""
    template = Path(sample_tree["template"]).resolve()
    output = (tmp_path / "out.xlsx").resolve()
    service = make_service(tmp_path)
    window = DialogWindow([str(template)], [str(output)])
    bridge = bind_bridge(monkeypatch, service, window)
    assert bridge.choose_template_file() == {"ok": True, "data": {"cancelled": False,
                                                                  "path": str(template), "kind": "template"}}
    assert bridge.choose_output_file() == {"ok": True, "data": {"cancelled": False, "path": str(output)}}
    assert window.calls == [
        ("OPEN_DIALOG", {"file_types": ("Excel workbook (*.xlsx)",)}),
        ("SAVE_DIALOG", {"save_filename": "Kiem_chung_Ket_qua.xlsx",
                         "file_types": ("Excel workbook (*.xlsx)",)}),
    ]


# --------------------------------------------------------------------------- 2. restart restores it
def test_restart_restores_both_excel_paths_without_any_dialog(sample_tree, tmp_path, monkeypatch):
    config_path = tmp_path / "config" / "config.json"
    template = Path(sample_tree["template"]).resolve()
    output = (tmp_path / "Output" / "Kiem chung 10-2026.xlsx").resolve()
    pick_both(monkeypatch, tmp_path, template=template, output=output, config_path=config_path)

    restarted = restarted_service(tmp_path, config_path)          # no native window at all
    paths = restarted.current_config()["paths"]
    assert paths["template"] == str(template)
    assert paths["output"] == str(output)
    state = excel_state(restarted)
    assert state["template"]["path"] == str(template)
    assert state["template"]["state"] == "ok"
    assert state["template"]["persisted"] is True
    assert state["output"]["state"] == "ok" and state["output"]["persisted"] is True
    # A restart never needs a dialog: the bridge answers without any window bound.
    assert bind_bridge(monkeypatch, restarted, None).get_excel_selection()["ok"] is True


def test_restarted_default_is_exposed_in_the_dashboard(sample_tree, tmp_path, monkeypatch):
    config_path = tmp_path / "config" / "config.json"
    template = Path(sample_tree["template"]).resolve()
    output = (tmp_path / "Output" / "ket_qua.xlsx").resolve()
    pick_both(monkeypatch, tmp_path, template=template, output=output, config_path=config_path)

    dashboard = restarted_service(tmp_path, config_path).dashboard_state()
    assert dashboard["excel"]["template"]["name"] == template.name
    assert dashboard["excel"]["template"]["path"] == str(template)
    assert dashboard["excel"]["output"]["state"] == "ok"


# --------------------------------------------------------------------------- 3. processing reuses it
def test_processing_uses_the_stored_excel_files_without_a_picker(sample_tree, tmp_path, monkeypatch):
    _RecordingProcessor.instances.clear()
    config_path = tmp_path / "config" / "config.json"
    template = Path(sample_tree["template"]).resolve()
    output = (tmp_path / "Output" / "Kiem chung 09-2026.xlsx").resolve()
    output.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(template, output)                                # an existing result workbook
    pick_both(monkeypatch, tmp_path, template=template, output=output, config_path=config_path)

    service = restarted_service(tmp_path, config_path, processor_factory=_RecordingProcessor)
    window = DialogWindow()                                       # any dialog call fails loudly
    bridge = bind_bridge(monkeypatch, service, window)
    dto = config_dto(Path(sample_tree["reports"]), template, output)
    assert bridge.scan_reports(dto, "")["ok"] is True
    assert bridge.start_processing(dto, False)["ok"] is True

    worker = _RecordingProcessor.instances[-1]
    assert worker.options.template == template
    assert worker.options.output_file == output
    assert window.calls == []                                     # processing never opened a picker


# --------------------------------------------------------------------------- 4. “Đổi file”
def test_second_pick_replaces_the_persisted_default(sample_tree, tmp_path, monkeypatch):
    config_path = tmp_path / "config" / "config.json"
    first = Path(sample_tree["template"]).resolve()
    second = (tmp_path / "Kiểm chứng 2026 (bản mới).xlsx")
    shutil.copy2(first, second)
    output = (tmp_path / "out.xlsx").resolve()
    service, bridge, _window = pick_both(monkeypatch, tmp_path, template=first, output=output,
                                         config_path=config_path)

    window = DialogWindow([str(second)])
    bridge = bind_bridge(monkeypatch, service, window)
    assert bridge.choose_template_file()["ok"] is True

    stored = AppConfig.load(config_path)
    assert stored.last_template == str(second)
    assert stored.last_output_file == str(output)                 # the other kind is untouched
    assert excel_state(service)["template"]["name"] == second.name
    assert window.calls == [("OPEN_DIALOG", {"file_types": ("Excel workbook (*.xlsx)",)})]


# --------------------------------------------------------------------------- 5. missing / invalid
def test_deleted_stored_template_is_reported_and_never_silently_replaced(sample_tree, tmp_path, monkeypatch):
    _RecordingProcessor.instances.clear()
    config_path = tmp_path / "config" / "config.json"
    template = (tmp_path / "Kiểm chứng sẽ bị xóa.xlsx")
    shutil.copy2(Path(sample_tree["template"]), template)
    output = (tmp_path / "Output" / "ket_qua.xlsx").resolve()
    pick_both(monkeypatch, tmp_path, template=template, output=output, config_path=config_path)

    template.unlink()
    service = restarted_service(tmp_path, config_path, processor_factory=_RecordingProcessor)
    window = DialogWindow()
    bridge = bind_bridge(monkeypatch, service, window)
    state = bridge.get_excel_selection()["data"]
    assert state["template"]["state"] == "missing"
    assert state["template"]["path"] == str(template)             # reported, not replaced
    assert "không" in state["template"]["message"].lower()
    assert state["output"]["state"] == "ok"

    assert service.current_config()["paths"]["template"] == str(template)
    dto = config_dto(Path(sample_tree["reports"]), template, output)
    assert bridge.scan_reports(dto, "")["ok"] is True
    failed = bridge.start_processing(dto, False)
    assert failed["ok"] is False and failed["error"]["code"] == "VALIDATION_FAILED"
    assert "File Kiểm chứng" in failed["error"]["message"]
    assert window.calls == []                                     # no automatic picker, no substitution


def test_corrupt_template_is_reported_as_invalid(sample_tree, tmp_path, monkeypatch):
    config_path = tmp_path / "config" / "config.json"
    template = tmp_path / "Khong phai bảng tính.xlsx"
    template.write_bytes(b"this is not a workbook")
    output = (tmp_path / "out.xlsx").resolve()
    service, _bridge, _window = pick_both(monkeypatch, tmp_path, template=template, output=output,
                                          config_path=config_path)
    state = excel_state(service)
    assert state["template"]["state"] == "invalid"
    assert state["template"]["message"]
    assert state["output"]["state"] == "ok"


def test_unconfigured_state_before_the_first_pick(tmp_path):
    service = make_service(tmp_path)
    state = excel_state(service)
    assert set(state) >= {"template", "output"}
    for kind in ("template", "output"):
        assert state[kind]["state"] == "unconfigured"
        assert state[kind]["path"] == "" and state[kind]["name"] == ""
        assert state[kind]["state"] in EXCEL_STATES


# --------------------------------------------------------------------------- 6. Unicode / spaces
def test_unicode_space_and_relative_pick_paths_round_trip(tmp_path, monkeypatch):
    config_path = tmp_path / "config" / "config.json"
    folder = tmp_path / "Báo cáo kiểm chứng 2026"
    nested = folder / "thư mục con"
    nested.mkdir(parents=True)
    template = folder / "Biểu mẫu kiểm chứng (bản 2).xlsx"
    shutil.copy2(_template_bytes(tmp_path), template)
    output = tmp_path / "Kết quả kiểm chứng 2026.xlsx"
    relative_template = nested / ".." / template.name
    relative_output = nested / ".." / ".." / output.name

    service, _bridge, _window = pick_both(monkeypatch, tmp_path, template=relative_template,
                                          output=relative_output, config_path=config_path)
    stored = AppConfig.load(config_path)
    assert stored.last_template == str(template.resolve())
    assert stored.last_output_file == str(output.resolve())

    restarted = restarted_service(tmp_path, config_path)
    assert restarted.controller.template == str(template.resolve())
    assert restarted.controller.output == str(output.resolve())
    state = excel_state(restarted)
    assert state["template"]["name"] == template.name
    assert state["template"]["folder"] == str(folder.resolve())
    assert state["template"]["state"] == "ok"


# --------------------------------------------------------------------------- 7/8. config compatibility
def test_legacy_config_without_excel_selection_loads_normally(tmp_path, monkeypatch):
    config_path = tmp_path / "config" / "config.json"
    config_path.parent.mkdir(parents=True)
    config_path.write_text(json.dumps({"ollama_server": "127.0.0.1:11434", "model": "qwen3:4b"}),
                           encoding="utf-8")
    service = restarted_service(tmp_path, config_path)
    assert service.controller.template == "" and service.controller.output == ""
    assert service.current_config()["paths"] == {"reportFolder": "", "template": "", "output": ""}
    assert excel_state(service)["template"]["state"] == "unconfigured"

    template = tmp_path / "Bieu mau.xlsx"
    shutil.copy2(_template_bytes(tmp_path), template)
    output = tmp_path / "ket_qua.xlsx"
    window = DialogWindow([str(template)], [str(output)])
    bridge = bind_bridge(monkeypatch, service, window)
    assert bridge.choose_template_file()["ok"] and bridge.choose_output_file()["ok"] is True
    stored = AppConfig.load(config_path)
    assert stored.last_template == str(template.resolve())
    assert stored.model == "qwen3:4b"                             # legacy values kept
    assert stored.ollama_server == "http://127.0.0.1:11434"


def test_corrupt_config_file_keeps_the_existing_load_contract(tmp_path, monkeypatch):
    config_path = tmp_path / "config" / "config.json"
    config_path.parent.mkdir(parents=True)
    config_path.write_text('{"ollama_server": "127.0.0.1:11434", ', encoding="utf-8")
    service = restarted_service(tmp_path, config_path)             # must not raise
    assert excel_state(service)["template"]["state"] == "unconfigured"

    template = tmp_path / "Bieu mau.xlsx"
    shutil.copy2(_template_bytes(tmp_path), template)
    window = DialogWindow([str(template)])
    bridge = bind_bridge(monkeypatch, service, window)
    assert bridge.choose_template_file()["ok"] is True
    stored = json.loads(config_path.read_text(encoding="utf-8"))   # rewritten as valid JSON
    assert stored["last_template"] == str(template.resolve())
    assert not list(config_path.parent.glob("*.tmp"))


def test_unknown_legacy_keys_survive_a_pick(tmp_path, monkeypatch):
    config_path = tmp_path / "config" / "config.json"
    config_path.parent.mkdir(parents=True)
    config_path.write_text(json.dumps({"last_template": "", "future_key": {"x": 1}}), encoding="utf-8")
    service = restarted_service(tmp_path, config_path)
    template = tmp_path / "Bieu mau.xlsx"
    shutil.copy2(_template_bytes(tmp_path), template)
    bridge = bind_bridge(monkeypatch, service, DialogWindow([str(template)]))
    assert bridge.choose_template_file()["ok"] is True
    stored = json.loads(config_path.read_text(encoding="utf-8"))
    assert stored["future_key"] == {"x": 1}
    assert stored["last_template"] == str(template.resolve())


# --------------------------------------------------------------------------- 9. updater preservation
def test_updater_preserves_the_saved_excel_selection(tmp_path, monkeypatch):
    import app.updater as up
    from tests.test_updater import (V_NEWER, B_NEWER, make_portable, make_update_folder,
                                    set_helper_runtime, set_packaged_runtime)

    portable = make_portable(tmp_path / "app")
    config_file = portable / "config" / "config.json"
    saved = {"last_template": "D:\\Kiểm chứng\\Biểu mẫu.xlsx",
             "last_output_file": "D:\\Kết quả\\Kiem chung 2026.xlsx", "model": "qwen3:4b"}
    config_file.write_text(json.dumps(saved, ensure_ascii=False), encoding="utf-8")
    learning = portable / "learning_data"
    learning.mkdir()
    (learning / "labels.json").write_text('{"image": {}}', encoding="utf-8")

    folder = make_update_folder(tmp_path / "Update", version=V_NEWER, build=B_NEWER)
    set_packaged_runtime(monkeypatch, portable)
    check = up.check_for_update(str(folder))
    lock = up.acquire_update_lock(portable, timeout=0)
    assert lock is not None
    try:
        staged = up.stage_update(check, portable, lock=lock)
    finally:
        lock.release()
    manifest = json.loads(Path(staged.manifest_path).read_text(encoding="utf-8"))
    manifest["old_pid"] = os.getpid() + 12345
    Path(staged.manifest_path).write_text(json.dumps(manifest), encoding="utf-8")
    set_helper_runtime(monkeypatch, staged)

    rc = up.apply_update(Path(staged.manifest_path), alive=lambda pid: False, sleep=lambda s: None,
                         spawn=lambda *a, **k: None)
    assert rc == 0
    assert (portable / up.APP_EXE).read_bytes() == b"MZ-new-" + V_NEWER.encode()
    assert json.loads(config_file.read_text(encoding="utf-8")) == saved        # Excel default preserved
    assert (learning / "labels.json").read_text(encoding="utf-8") == '{"image": {}}'


# --------------------------------------------------------------------------- 10. Excel safety
def test_persisting_a_pick_does_not_open_or_modify_the_workbooks(tmp_path, monkeypatch):
    config_path = tmp_path / "config" / "config.json"
    template = tmp_path / "Kiem chung goc.xlsx"
    shutil.copy2(_template_bytes(tmp_path), template)
    output = tmp_path / "Ket qua.xlsx"
    shutil.copy2(template, output)
    before = {path: (path.read_bytes(), path.stat().st_mtime_ns) for path in (template, output)}

    service, bridge, _window = pick_both(monkeypatch, tmp_path, template=template, output=output,
                                         config_path=config_path)
    excel_state(service)
    assert bridge.clear_excel_selection("template")["ok"] is True

    for path, (payload, mtime) in before.items():
        assert path.read_bytes() == payload
        assert path.stat().st_mtime_ns == mtime
        assert not (path.parent / f"~${path.name}").exists()
        assert not Path(f"{path}.saving.xlsx").exists()
    assert not list(tmp_path.glob("*.tmp"))


def test_pick_reports_when_the_selection_could_not_be_persisted(tmp_path, monkeypatch):
    """A pick whose config write genuinely fails stays usable and is reported honestly.

    The failure is injected at the real disk boundary — ``AppConfig.save`` raising ``OSError``, exactly what an
    unwritable or locked ``config/`` folder produces — instead of relying on POSIX directory permissions.
    That keeps this test identical on Windows, macOS and Linux: Windows has no ``os.geteuid`` at all, and
    ``chmod`` does not stop file creation inside a directory there.  The production path still runs unmodified:
    the real ``GuiController.save_settings()`` body, its ``except OSError`` handler, the re-read of
    ``AppConfig.load(...)`` and the resulting ``persisted: false`` in the DTO.
    """
    config_path = tmp_path / "config" / "config.json"
    template = tmp_path / "Bieu mau.xlsx"
    shutil.copy2(_template_bytes(tmp_path), template)
    service = make_service(tmp_path, config_path=config_path)
    bridge = bind_bridge(monkeypatch, service, DialogWindow([str(template)]))

    real_save = AppConfig.save

    def unwritable_save(_self, _path=None):
        raise PermissionError(13, "config.json không ghi được (mô phỏng thư mục chỉ đọc)")

    monkeypatch.setattr(AppConfig, "save", unwritable_save)

    picked = bridge.choose_template_file()
    assert picked["ok"] is True and picked["data"]["path"] == str(template.resolve())
    assert not config_path.exists()                                # the write really did not happen
    assert not config_path.with_suffix(".json.tmp").exists()       # …not even a leftover temp file
    assert any("Không lưu được config.json" in line for line in service.controller.log_lines)
    state = bridge.get_excel_selection()["data"]
    assert state["template"]["path"] == str(template)              # used for this session
    assert state["template"]["state"] in EXCEL_STATES
    assert state["template"]["persisted"] is False                 # and reported honestly

    # The failure is not sticky: once writes succeed again the same selection is persisted for real.
    monkeypatch.setattr(AppConfig, "save", real_save)
    repaired = service.persist_excel_selection("template", str(template))
    assert repaired["template"]["persisted"] is True
    assert AppConfig.load(config_path).last_template == str(template)


# --------------------------------------------------------------------------- 11. other pickers unchanged
def test_folder_pptx_and_update_pickers_do_not_write_an_excel_default(sample_tree, tmp_path, monkeypatch):
    service = make_service(tmp_path)
    template = Path(sample_tree["template"]).resolve()
    window = DialogWindow([str(sample_tree["reports"])], [str(sample_tree["files"][0])],
                          [str(sample_tree["reports"])], [str(template)])
    bridge = bind_bridge(monkeypatch, service, window)

    folder = bridge.choose_report_folder()
    assert folder["ok"] and Path(folder["data"]["path"]) == Path(sample_tree["reports"]).resolve()
    pptx = bridge.choose_pptx_file()
    assert pptx["ok"] and pptx["data"]["name"] == Path(sample_tree["files"][0]).name
    assert service.controller.template == "" and service.controller.output == ""
    update_folder = bridge.choose_update_folder()
    assert update_folder["ok"] and Path(update_folder["data"]["path"]) == Path(sample_tree["reports"]).resolve()
    assert service.controller.template == ""

    assert bridge.choose_template_file()["ok"] is True
    assert AppConfig.load(tmp_path / "config" / "config.json").last_template == str(template)
    assert window.calls == [
        ("FOLDER_DIALOG", {}),
        ("OPEN_DIALOG", {"file_types": ("PowerPoint presentation (*.pptx)",)}),
        ("FOLDER_DIALOG", {}),
        ("OPEN_DIALOG", {"file_types": ("Excel workbook (*.xlsx)",)}),
    ]


def test_cancelled_excel_pick_is_not_persisted_and_changes_nothing(tmp_path, monkeypatch):
    config_path = tmp_path / "config" / "config.json"
    service = make_service(tmp_path, config_path=config_path)
    bridge = bind_bridge(monkeypatch, service, DialogWindow(None))
    assert bridge.choose_template_file() == {"ok": True, "data": {"cancelled": True}}
    assert not config_path.exists()
    assert service.controller.template == ""
    assert excel_state(service)["template"]["state"] == "unconfigured"


# --------------------------------------------------------------------------- 12/13. clear + validation
def test_clear_excel_selection_is_persisted_and_idempotent(sample_tree, tmp_path, monkeypatch):
    config_path = tmp_path / "config" / "config.json"
    template = Path(sample_tree["template"]).resolve()
    output = (tmp_path / "out.xlsx").resolve()
    service, bridge, _window = pick_both(monkeypatch, tmp_path, template=template, output=output,
                                         config_path=config_path)

    cleared = bridge.clear_excel_selection("template")
    assert cleared["ok"] is True and cleared["data"]["template"]["state"] == "unconfigured"
    assert service.controller.template == ""
    assert AppConfig.load(config_path).last_template == ""
    assert AppConfig.load(config_path).last_output_file == str(output)

    again = bridge.clear_excel_selection("template")
    assert again["ok"] is True and again["data"]["template"]["state"] == "unconfigured"
    assert bridge.clear_excel_selection("nonsense")["ok"] is False

    bridge.clear_excel_selection("output")
    stored = AppConfig.load(config_path)
    assert stored.last_output_file == "" and stored.last_template == ""


def test_excel_states_cover_every_documented_state(sample_tree, tmp_path, monkeypatch):
    config_path = tmp_path / "config" / "config.json"
    template = tmp_path / "Bieu mau.xlsx"
    shutil.copy2(Path(sample_tree["template"]), template)
    service, bridge, _window = pick_both(monkeypatch, tmp_path, template=template,
                                         output=(tmp_path / "out.xlsx").resolve(),
                                         config_path=config_path)
    assert bridge.get_excel_selection()["data"]["template"]["state"] in EXCEL_STATES
    template.write_bytes(b"corrupt")
    time.sleep(0.01)
    assert excel_state(service)["template"]["state"] == "invalid"
    template.unlink()
    assert excel_state(service)["template"]["state"] == "missing"
    template.write_bytes(Path(sample_tree["template"]).read_bytes())
    assert excel_state(service)["template"]["state"] == "ok"


def test_excel_selection_keeps_learning_storage_untouched(sample_tree, tmp_path, monkeypatch):
    config_path = tmp_path / "config" / "config.json"
    pick_both(monkeypatch, tmp_path, template=Path(sample_tree["template"]),
              output=(tmp_path / "out.xlsx").resolve(), config_path=config_path)
    saved = learning_dir(create=False)
    assert not saved.exists() or not any(saved.iterdir())
    assert AppConfig.load(config_path).last_template


# --------------------------------------------------------------------------- helpers
def _template_bytes(tmp_path: Path) -> Path:
    """A real, valid check workbook created once per test (never a hand-written fake)."""
    from make_samples import make_template

    cached = tmp_path / "fixture-template.xlsx"
    if not cached.exists():
        make_template(cached)
    return cached
