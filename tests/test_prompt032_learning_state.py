"""PROMPT-032: current Learning lifecycle, candidate availability, bridge and preview-failure contracts.

The production fixture deliberately runs the real parser/classifier/extractor/writer path with Ollama disabled.  It
proves that candidate objects retained by a completed worker are available to Learning without relying on an unrelated
dashboard poll.  Linux does not prove Windows pywebview timing or PowerPoint rendering fidelity.
"""
from __future__ import annotations

import io
import json
import logging
from pathlib import Path
from types import SimpleNamespace

from PIL import Image

from app.application_service import ApplicationService
from app.bridge import BridgeService
from app.config import AppConfig
from app.gui_controller import GuiController
from tests.test_after_evidence import _deck
from tests.test_application_service import _learning_candidates, _make_service
from tests.test_content_learning import slide_items
from tests.test_content_region import _prefill
from tests.test_prompt004 import MGMT
from make_samples import make_template


def _config(reports: Path, template: Path, output: Path) -> dict:
    return {
        "paths": {"reportFolder": str(reports), "template": str(template), "output": str(output)},
        "period": {"mode": "all", "month": "9", "year": "2026", "from": "", "to": ""},
        "forceReprocess": False,
        "ollama": {"host": "127.0.0.1", "port": "11434", "model": "qwen3:4b"},
        "update": {"path": "", "autoCheck": False},
    }


def _finished_production_service(tmp_path: Path):
    reports = tmp_path / "reports"
    reports.mkdir()
    deck = _deck(
        reports / f"(CTMS)_{MGMT}_ĐỐI SÁCH LỖI SƠN 24.09.2026.pptx",
        [slide_items],
    )
    template = make_template(tmp_path / "Verification.xlsx")
    _prefill(template, [{"mgmt": MGMT}])
    output = tmp_path / "output" / "result.xlsx"
    config = _config(reports, template, output)
    cfg = AppConfig(
        last_report_folder=str(reports), last_template=str(template), last_output_file=str(output), period_mode="all"
    )
    controller = GuiController(
        cfg, config_path=tmp_path / "config" / "settings.json", learning_dir_override=tmp_path / "learning_data"
    )
    service = ApplicationService(
        controller,
        config_path=tmp_path / "config" / "settings.json",
        manual_fields_path=tmp_path / "config" / "manual_fields.json",
    )
    token = service.register_pptx_selection(str(deck))["token"]
    scanned = service.scan_reports(config, token)
    assert scanned["scan"]["scanned"] is True and len(scanned["reports"]) == 1
    service.start_processing(config, use_ollama=False)
    processor = controller.processor
    processor._thread.join(timeout=60)
    assert not processor._thread.is_alive()
    return service, controller, processor, deck


def _empty_preview(**overrides):
    value = {"src": "", "width": 0, "height": 0, "backend": "", "faithful": False, "regions": []}
    value.update(overrides)
    return value


def _png_bytes() -> bytes:
    stream = io.BytesIO()
    Image.new("RGB", (24, 16), "#20a080").save(stream, format="PNG")
    return stream.getvalue()


def test_completed_production_batch_learning_read_reconciles_and_returns_both_candidate_kinds(tmp_path, monkeypatch):
    service, controller, processor, deck = _finished_production_service(tmp_path)
    result = processor.results[0]

    # Production creation and retention: one eligible reviewable image plus verbatim content are already present.
    assert result.status in {"completed", "needs_review"}
    assert result.image_candidates and any(
        not candidate.hard_excluded and candidate.excel_output_eligible for candidate in result.image_candidates
    )
    assert result.content_candidates and any(not candidate.hard_excluded for candidate in result.content_candidates)
    assert controller.state == "running" and not controller._queue.empty()  # terminal callback not applied yet

    # Keep this lifecycle regression focused; preview-failure behavior has a dedicated assertion below.
    monkeypatch.setattr(service, "_slide_preview_for", lambda candidate, learning=None: _empty_preview())
    envelope = BridgeService(service).get_learning_state()
    assert envelope["ok"] is True
    state = envelope["data"]

    # No dashboard read or restart was needed: get_learning_state itself applied the terminal event.
    assert controller.state == "idle"
    assert len(state["images"]) == sum(not candidate.hard_excluded for candidate in result.image_candidates) >= 1
    assert len(state["contents"]) == sum(not candidate.hard_excluded for candidate in result.content_candidates) >= 1
    assert state["candidateStatus"]["image"] == {
        "state": "ready", "reason": "available", "count": len(state["images"]),
        "message": f"Đã tải {len(state['images'])} ảnh ứng viên từ lượt xử lý hiện tại.",
    }
    assert state["candidateStatus"]["content"]["state"] == "ready"
    assert state["candidateStatus"]["content"]["reason"] == "available"
    assert state["candidateStatus"]["content"]["count"] == len(state["contents"])
    assert any("Cải tiến lỗi" in candidate["text"] for candidate in state["contents"])
    assert str(tmp_path) not in json.dumps(state, ensure_ascii=False)

    # A second manual-style call really re-reads current backend state; it is not a frozen frontend snapshot.
    processor.results = []
    refreshed = BridgeService(service).get_learning_state()["data"]
    assert refreshed["images"] == [] and refreshed["contents"] == []
    assert refreshed["candidateStatus"]["image"]["reason"] == "empty"
    assert refreshed["candidateStatus"]["content"]["reason"] == "empty"
    assert deck.name not in json.dumps(refreshed, ensure_ascii=False)


def test_learning_availability_distinguishes_no_run_processing_valid_empty_and_unavailable(
        sample_tree, tmp_path, monkeypatch):
    service, _ = _make_service(sample_tree, tmp_path)

    no_run = service.learning_state()
    assert no_run["candidateStatus"]["image"]["state"] == "no_run"
    assert no_run["candidateStatus"]["content"]["state"] == "no_run"

    service.controller.processor = SimpleNamespace(results=[])
    empty = service.learning_state()
    assert empty["candidateStatus"]["image"]["state"] == "ready"
    assert empty["candidateStatus"]["image"]["reason"] == "empty"
    assert empty["candidateStatus"]["content"]["state"] == "ready"
    assert empty["candidateStatus"]["content"]["reason"] == "empty"

    image, content = _learning_candidates(sample_tree)
    service.controller.processor = SimpleNamespace(
        results=[SimpleNamespace(image_candidates=[image], content_candidates=[content])],
        is_running=lambda: True,
    )
    service.controller.state = "running"
    processing = service.learning_state()
    assert processing["images"] == [] and processing["contents"] == []
    assert processing["candidateStatus"]["image"]["state"] == "processing"
    assert processing["candidateStatus"]["content"]["state"] == "processing"

    service.controller.state = "idle"
    monkeypatch.setattr(GuiController, "learning", property(lambda self: None))
    unavailable = service.learning_state()
    assert unavailable["candidateStatus"]["image"]["state"] == "unavailable"
    assert unavailable["candidateStatus"]["content"]["state"] == "unavailable"


def test_image_preview_exception_keeps_truthful_candidate_metadata(sample_tree, tmp_path, monkeypatch):
    service, _ = _make_service(sample_tree, tmp_path)
    image, _content = _learning_candidates(sample_tree)
    service.controller.processor = SimpleNamespace(
        results=[SimpleNamespace(image_candidates=[image], content_candidates=[])]
    )
    monkeypatch.setattr(service.controller, "candidate_blob", lambda candidate: _png_bytes())

    def fail_preview(candidate, learning=None):
        raise RuntimeError("display-only synthetic renderer failure")

    monkeypatch.setattr(service, "_slide_preview_for", fail_preview)
    state = service.learning_state()
    dto = state["images"][0]
    assert dto["id"] == image.candidate_id
    assert dto["decision"] == image.decision and dto["excelEligible"] == image.excel_output_eligible
    assert dto["targetBbox"] == {"x": 100, "y": 100, "width": 300, "height": 200}
    assert dto["evidence"] == image.evidence and dto["eligibilityReason"] == image.eligibility_reason
    assert dto["slidePreview"] == "" and dto["src"].startswith("data:image/jpeg;base64,")
    assert state["candidateStatus"]["image"]["state"] == "ready"
    assert state["candidateStatus"]["image"]["count"] == 1


def test_learning_diagnostics_are_transition_bounded_and_do_not_leak_payloads(
        sample_tree, tmp_path, monkeypatch, caplog):
    service, _ = _make_service(sample_tree, tmp_path)
    image, content = _learning_candidates(sample_tree)
    service.controller.processor = SimpleNamespace(
        results=[SimpleNamespace(image_candidates=[image], content_candidates=[content])]
    )
    monkeypatch.setattr(service, "_slide_preview_for", lambda candidate, learning=None: _empty_preview())
    caplog.set_level(logging.INFO, logger="report_extractor.webview")

    service.learning_state()
    service.learning_state()
    service.learning_state()
    lines = [record.getMessage() for record in caplog.records if record.getMessage().startswith("LEARNING_")]
    state_lines = [line for line in lines if line.startswith("LEARNING_STATE")]
    assert len(state_lines) == 1
    assert "image_raw=1" in state_lines[0] and "content_raw=1" in state_lines[0]

    service.controller.processor.results = []
    service.learning_state()
    state_lines = [record.getMessage() for record in caplog.records if record.getMessage().startswith("LEARNING_STATE")]
    assert len(state_lines) == 2  # one more transition, still not one line per refresh
    dump = "\n".join(state_lines)
    assert str(tmp_path) not in dump and str(sample_tree["reports"]) not in dump
    assert content.text not in dump and "data:image" not in dump
