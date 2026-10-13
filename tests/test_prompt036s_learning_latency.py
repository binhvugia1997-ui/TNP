"""PROMPT-036S: Learning label clicks must not pay for full-state bridge round trips.

Proven cause (see .specify/bugs/prompt036s-learning-latency/assessment.md): the old
``set_learning_label`` rebuilt and returned the ENTIRE LearningState per click (3.13 MB /
694 ms cold at 24 image candidates on the bench fixture), the Save path chained
``set_learning_label`` + ``save_learning_labels`` (two full-state round trips, the second
holding the global service lock across the Excel reapply), and the frontend applied no
local state until that round trip resolved. These tests pin the additive fast contract:
a minimal-ack pending write that never rebuilds the state, and one atomic
set+save call with identical persistence/Excel semantics.
"""
from __future__ import annotations

import shutil
from pathlib import Path
from types import SimpleNamespace

import pytest

from app.bridge import BridgeService
from tests.test_application_service import _learning_candidates, _make_service, _png_bytes


def _service_with_candidates(sample_tree, tmp_path, monkeypatch):
    service, config = _make_service(sample_tree, tmp_path)
    image, content = _learning_candidates(sample_tree)
    service.controller.processor = SimpleNamespace(
        results=[SimpleNamespace(image_candidates=[image], content_candidates=[content])])
    monkeypatch.setattr(service.controller, "candidate_blob", lambda candidate: _png_bytes(), raising=False)
    output = Path(config["paths"]["output"])
    output.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(sample_tree["template"], output)
    service.controller.template = str(sample_tree["template"])
    service.controller.output = str(output)
    return service, image, content


def test_fast_label_returns_minimal_ack_without_rebuilding_state(sample_tree, tmp_path, monkeypatch):
    service, image, _content = _service_with_candidates(sample_tree, tmp_path, monkeypatch)
    rebuilds = []
    original = service.learning_state

    def spy():
        rebuilds.append(1)
        return original()

    monkeypatch.setattr(service, "learning_state", spy)

    ack = service.set_learning_label_fast("image", image.candidate_id, "AFTER", "quick note")

    assert ack == {"kind": "image", "candidateId": image.candidate_id,
                   "label": "AFTER", "labelPending": True}
    assert rebuilds == [], "a single label click must not rebuild the full LearningState"
    # The pending write itself is the real controller operation, visible to later reads.
    assert service.controller.pending_labels[image.candidate_id] == "AFTER"
    assert service.controller.pending_label_notes[image.candidate_id] == "quick note"
    state = original()
    assert state["images"][0]["userLabel"] == "AFTER"
    assert state["images"][0]["labelPending"] is True


def test_fast_label_last_click_wins_and_unknown_candidate_is_an_error(sample_tree, tmp_path, monkeypatch):
    service, image, content = _service_with_candidates(sample_tree, tmp_path, monkeypatch)

    service.set_learning_label_fast("image", image.candidate_id, "BEFORE", "")
    service.set_learning_label_fast("image", image.candidate_id, "AFTER", "")
    service.set_learning_label_fast("image", image.candidate_id, "BEFORE", "")
    assert service.controller.pending_labels[image.candidate_id] == "BEFORE"

    service.set_learning_label_fast("content", content.candidate_id, "IMPROVEMENT_CONTENT", "")
    assert service.controller.pending_content_labels[content.candidate_id] == "IMPROVEMENT_CONTENT"

    from app.application_service import ServiceError
    with pytest.raises(ServiceError):
        service.set_learning_label_fast("image", "missing-candidate", "AFTER", "")


def test_atomic_save_candidate_persists_label_and_note_in_one_call(sample_tree, tmp_path, monkeypatch):
    service, image, _content = _service_with_candidates(sample_tree, tmp_path, monkeypatch)
    from app.image_review import ReapplyResult
    import app.image_review as image_review

    reapply_calls = []

    def reapply(*args, **kwargs):
        reapply_calls.append(1)
        return ReapplyResult(prepared_rows=[4], updated_rows=[4], pictures=1)

    monkeypatch.setattr(image_review, "reapply_labels", reapply)
    rebuilds = []
    original = service.learning_state
    monkeypatch.setattr(service, "learning_state", lambda: (rebuilds.append(1), original())[1])

    result = service.save_learning_candidate("image", image.candidate_id, "AFTER", "saved atomically")

    assert result["labelPending"] is False
    assert result["excelPending"] == 0
    assert result["excelLastResult"]["kind"] == "ok"
    assert "Đã lưu 1 nhãn" in result["message"]
    assert rebuilds == [], "the atomic save must not rebuild the full LearningState either"
    assert reapply_calls, "Excel reapply semantics must be preserved"
    saved = service.controller.learning.store.latest()
    assert saved[image.candidate_id]["label"] == "AFTER"
    assert saved[image.candidate_id]["note"] == "saved atomically"
    assert service.controller.pending_labels == {}


def test_atomic_save_candidate_preserves_locked_excel_retry_semantics(sample_tree, tmp_path, monkeypatch):
    service, image, _content = _service_with_candidates(sample_tree, tmp_path, monkeypatch)
    from app.image_review import ReapplyResult
    import app.image_review as image_review

    output = service.controller.output
    results = [
        ReapplyResult(prepared_rows=[4], errors=["Excel locked"], locked_path=output),
        ReapplyResult(prepared_rows=[4], updated_rows=[4], pictures=1),
    ]
    monkeypatch.setattr(image_review, "reapply_labels", lambda *a, **k: results.pop(0))

    locked = service.save_learning_candidate("image", image.candidate_id, "AFTER", "")
    assert locked["excelLastResult"]["kind"] == "locked"
    assert locked["excelPending"] == 1
    # Labels are committed even while Excel is locked; only the Excel step is owed.
    assert service.controller.learning.store.latest()[image.candidate_id]["label"] == "AFTER"

    retried = service.retry_learning_excel()
    assert retried["state"]["excelLastResult"]["kind"] == "ok"
    assert retried["state"]["excelPending"] == 0
    assert len(service.controller.learning.store.latest()) == 1  # retry never duplicates labels


def test_bridge_fast_paths_use_the_same_envelope(sample_tree, tmp_path, monkeypatch):
    service, image, _content = _service_with_candidates(sample_tree, tmp_path, monkeypatch)
    bridge = BridgeService(service)

    ack = bridge.set_learning_label_fast("image", image.candidate_id, "BEFORE", "")
    assert ack["ok"] is True
    assert ack["data"]["label"] == "BEFORE"
    assert "images" not in ack["data"], "bridge ack must stay minimal"

    saved = bridge.save_learning_candidate("image", image.candidate_id, "AFTER", "via bridge")
    assert saved["ok"] is True
    assert saved["data"]["labelPending"] is False
    assert service.controller.learning.store.latest()[image.candidate_id]["note"] == "via bridge"

    bad = bridge.set_learning_label_fast("image", "missing", "AFTER", "")
    assert bad["ok"] is False and bad["error"]["code"] == "CANDIDATE_NOT_FOUND"
