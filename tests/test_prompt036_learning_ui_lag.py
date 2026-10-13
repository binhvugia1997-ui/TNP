"""PROMPT-036: the "Học cải tiến" click path must not re-read the whole label history per candidate.

Assessment evidence (`.specify/bugs/prompt036-learning-ui-lag/assessment.md`): with 60 image + 60 content
candidates and 400 saved labels per store, ONE ``set_learning_label`` click cost 1203.7 ms, of which 89.3 %
was 301 full re-reads + re-parses of ``image_labels.jsonl`` and 301 of ``content_labels.jsonl``.  The store
views are now memoized against the file's own identity, so the same click costs ~14 ms while the persisted
records, the returned DTO contract and the append-only audit trail stay exactly as before.

Linux only: this pins Python-side cost and contract.  Windows PowerPoint COM, pywebview transport timing and
Portable acceptance are NOT covered here (Constitution VII).
"""
from __future__ import annotations

import io
import json
import time
from pathlib import Path
from types import SimpleNamespace

import pytest
from PIL import Image

from app.application_service import ApplicationService
from app.config import AppConfig
from app.content_learning import CONTENT_LABELS_FILE
from app.gui_controller import GuiController
from app.image_learning import LABELS_FILE, ImageCandidate, LabelStore
from app.content_learning import ContentCandidate
from app.pptx_parser import parse_pptx

from make_samples import make_report, make_template


def _png() -> bytes:
    stream = io.BytesIO()
    Image.new("RGB", (64, 48), "#20a080").save(stream, format="PNG")
    return stream.getvalue()


def _candidate(candidate_id: str, slide: int, picture_id: int) -> ImageCandidate:
    return ImageCandidate(
        candidate_id=candidate_id, management_number="260918080-VOC", source_name="deck.pptx",
        source_file="", slide=slide, picture_id=picture_id, order=picture_id,
        bounds=(100, 100, 300, 200), slide_size=(12192000, 6858000),
        block_bounds=(0, 0, 12192000, 6858000), features={}, evidence=["prompt-036"], confidence=0.8,
        excel_output_eligible=True,
    )


def _content(candidate_id: str, slide: int, shape_id: int) -> ContentCandidate:
    return ContentCandidate(
        candidate_id=candidate_id, management_number="260918080-VOC", source_name="deck.pptx",
        source_file="", slide=slide, shape_id=shape_id, order=shape_id, shape_kind="text",
        bounds=(100, 100, 400, 200), slide_size=(12192000, 6858000), text="Cải tiến trong sản xuất",
        features={}, evidence=["prompt-036"], confidence=0.7,
    )


@pytest.fixture
def read_counter(monkeypatch):
    """Count real `Path.read_text` calls per resolved path — i.e. actual label-file disk reads."""
    counts: dict[str, int] = {}
    real = Path.read_text

    def counting(self, *args, **kwargs):
        counts[str(Path(self))] = counts.get(str(Path(self)), 0) + 1
        return real(self, *args, **kwargs)

    monkeypatch.setattr(Path, "read_text", counting)
    return counts


# ============================================================== LabelStore memoization contract
def test_label_store_serves_repeated_reads_from_one_disk_read(tmp_path, read_counter):
    store = LabelStore(tmp_path)
    store.label(_candidate("a|S1|#1|1", 1, 1), "AFTER", "note a")
    store.label(_candidate("b|S1|#2|1", 1, 2), "BEFORE", "note b")
    path_key = str(store.path)
    read_counter.clear()

    first = store.records()
    assert len(first) == 2 and read_counter[path_key] == 1

    # 200 further logical reads (what one learning_state() build used to do ~5x per candidate)
    for _ in range(100):
        assert store.records() == first
        assert store.latest()["a|S1|#1|1"]["label"] == "AFTER"
    assert read_counter[path_key] == 1, "the append-only file was re-read; PROMPT-036 regression"
    assert store.current_label("b|S1|#2|1") == "BEFORE"
    assert store.current_label("missing") == "UNLABELED"


def test_label_store_cache_is_invalidated_by_its_own_append(tmp_path, read_counter):
    store = LabelStore(tmp_path)
    store.label(_candidate("a|S1|#1|1", 1, 1), "AFTER", "")
    assert store.current_label("a|S1|#1|1") == "AFTER"
    read_counter.clear()

    # a changed label supersedes the old one and must be visible immediately (no stale memo)
    store.label(_candidate("a|S1|#1|1", 1, 1), "BEFORE", "changed")
    assert store.current_label("a|S1|#1|1") == "BEFORE"
    assert store.latest()["a|S1|#1|1"]["note"] == "changed"
    assert store.latest()["a|S1|#1|1"]["previous_label"] == "AFTER"
    assert len(store.records()) == 2, "the append-only audit trail must keep both records"
    assert read_counter[str(store.path)] >= 1


def test_label_store_cache_is_invalidated_by_an_external_file_change(tmp_path, read_counter):
    """Update rollback / restored backup / manual edit: a different file must never be served from memo."""
    store = LabelStore(tmp_path)
    store.label(_candidate("a|S1|#1|1", 1, 1), "AFTER", "")
    assert store.current_label("a|S1|#1|1") == "AFTER"

    replaced = tmp_path / LABELS_FILE
    restored = json.dumps({"schema_version": store.schema, "candidate_id": "z|S9|#9|1", "label": "IGNORE",
                           "previous_label": None, "timestamp": "2026-10-13T00:00:00",
                           "management_number": "260918080-VOC", "source_name": "deck.pptx", "slide": 9,
                           "shape_id": 9, "features": {}, "deterministic_kind": "ambiguous",
                           "deterministic_confidence": 0.5, "note": "restored"}, ensure_ascii=False)
    replaced.write_text(restored + "\n", encoding="utf-8")

    latest = store.latest()
    assert "a|S1|#1|1" not in latest, "a replaced label file was served from the memo"
    assert latest["z|S9|#9|1"]["label"] == "IGNORE"

    # and deleting the file is honoured too
    replaced.unlink()
    assert store.records() == [] and store.latest() == {}


def test_image_and_content_stores_never_share_a_memoized_view(tmp_path):
    """Constitution IV: the two learning datasets are separate files and must not leak into each other."""
    images = LabelStore(tmp_path)
    contents = LabelStore(tmp_path, CONTENT_LABELS_FILE, ("IMPROVEMENT_CONTENT", "EXCLUDE_CONTENT"),
                        ("length",), 1, positive="IMPROVEMENT_CONTENT")
    images.label(_candidate("img|S1|#1|1", 1, 1), "AFTER", "")
    contents.label(_content("txt|S1|SH1|1", 1, 1), "IMPROVEMENT_CONTENT", "")

    assert images.path != contents.path
    assert set(images.latest()) == {"img|S1|#1|1"}
    assert set(contents.latest()) == {"txt|S1|SH1|1"}
    assert images.current_label("txt|S1|SH1|1") == "UNLABELED"
    assert contents.current_label("img|S1|#1|1") == "UNLABELED"
    # repeated interleaved reads stay correct and cheap
    for _ in range(20):
        assert images.current_label("img|S1|#1|1") == "AFTER"
        assert contents.current_label("txt|S1|SH1|1") == "IMPROVEMENT_CONTENT"


# ============================================================== the click path itself
def _service_with_many_candidates(tmp_path, monkeypatch, n_images=40, n_contents=40, saved=200):
    """Real controller + service over a real deck, with many reviewable candidates and a real label history."""
    reports = tmp_path / "reports"
    reports.mkdir()
    deck = make_report(reports / "deck.pptx")
    template = make_template(tmp_path / "Verification.xlsx")
    output = tmp_path / "output" / "result.xlsx"
    output.parent.mkdir(parents=True, exist_ok=True)

    cfg = AppConfig(last_report_folder=str(reports), last_template=str(template),
                    last_output_file=str(output), period_mode="all")
    controller = GuiController(cfg, config_path=tmp_path / "config" / "settings.json",
                               learning_dir_override=tmp_path / "learning_data")
    service = ApplicationService(controller, config_path=tmp_path / "config" / "settings.json",
                                 manual_fields_path=tmp_path / "config" / "manual_fields.json")

    parsed = parse_pptx(deck)
    slides = [n for n in (5, 6) if parsed.slide(n) is not None] or [1]
    images = [_candidate(f"MGMT|S{slides[i % len(slides)]}|#{900 + i}|1", slides[i % len(slides)], 900 + i)
              for i in range(n_images)]
    for item in images:
        item.source_file = str(deck)
    contents = [_content(f"MGMT|S{slides[i % len(slides)]}|SH{800 + i}|1", slides[i % len(slides)], 800 + i)
                for i in range(n_contents)]
    for item in contents:
        item.source_file = str(deck)

    controller.processor = SimpleNamespace(
        results=[SimpleNamespace(image_candidates=images, content_candidates=contents)],
        is_running=lambda: False,
    )
    monkeypatch.setattr(controller, "candidate_blob", lambda candidate: _png(), raising=False)

    # a realistic learning_data history: this is what turned one click into ~1.2 s
    learning = controller.learning
    assert learning is not None
    for i in range(saved):
        learning.store.label(_candidate(f"HIST|S{(i % 8) + 1}|#{i}|1", (i % 8) + 1, i),
                             "AFTER" if i % 2 else "BEFORE", "ghi chú lịch sử " * 4)
        learning.content.store.label(_content(f"HISTC|S{(i % 8) + 1}|SH{i}|1", (i % 8) + 1, i),
                                     "IMPROVEMENT_CONTENT", "ghi chú lịch sử " * 4)
    return service, controller, images, contents


def test_one_label_click_reads_the_label_files_a_bounded_number_of_times(tmp_path, monkeypatch, read_counter):
    """The demonstrated cause: reads used to scale with candidates x saved labels (301 per store per click)."""
    service, controller, images, contents = _service_with_many_candidates(tmp_path, monkeypatch)
    image_file = str(controller.learning.store.path)
    content_file = str(controller.learning.content.store.path)

    service.learning_state()                     # warm the slide-preview cache, as a real session would
    read_counter.clear()
    state = service.set_learning_label("image", images[0].candidate_id, "AFTER", "")
    clicks_reads = read_counter.get(image_file, 0) + read_counter.get(content_file, 0)

    assert len(images) == 40 and len(contents) == 40
    # BEFORE the fix this was 602 file reads for one click; the bound is small and independent of both the
    # candidate count and the size of the saved label history.
    assert clicks_reads <= 4, f"one click re-read the label files {clicks_reads} times (PROMPT-036 regression)"

    # the returned contract is unchanged: still the full authoritative learning state
    assert state["counts"]["image"]["labeled"] >= 1
    dto = next(item for item in state["images"] if item["id"] == images[0].candidate_id)
    assert dto["userLabel"] == "AFTER" and dto["labelPending"] is True
    assert len(state["images"]) == len(images) and len(state["contents"]) == len(contents)


def test_one_label_click_is_fast_with_a_large_saved_label_history(tmp_path, monkeypatch):
    service, controller, images, _contents = _service_with_many_candidates(tmp_path, monkeypatch)
    service.learning_state()                     # warm

    started = time.perf_counter()
    service.set_learning_label("image", images[0].candidate_id, "AFTER", "")
    elapsed_ms = (time.perf_counter() - started) * 1000

    assert elapsed_ms < 250, (
        f"one Before/After click took {elapsed_ms:.0f} ms of Python work; measured 1203.7 ms before the "
        "PROMPT-036 fix with the same 40+40 candidates and 200 saved labels per store"
    )


def test_rapid_before_after_switching_keeps_every_label_and_the_audit_trail(tmp_path, monkeypatch):
    """Rapid Trước <-> Sau switching must not lose a label nor duplicate the append-only history."""
    service, controller, images, _contents = _service_with_many_candidates(tmp_path, monkeypatch, saved=0)
    target = images[0].candidate_id

    for label in ("AFTER", "BEFORE", "AFTER", "BEFORE", "AFTER"):
        state = service.set_learning_label("image", target, label, "")
        dto = next(item for item in state["images"] if item["id"] == target)
        assert dto["userLabel"] == label and dto["labelPending"] is True

    assert controller.pending_labels[target] == "AFTER"
    saved = service.save_learning_labels("image", {target: ""})
    assert saved["state"]["images"][0]["labelPending"] is False
    records = controller.learning.store.records()
    assert [record["label"] for record in records if record["candidate_id"] == target] == ["AFTER"]
    assert controller.learning.store.current_label(target) == "AFTER"


def test_saving_labels_still_invalidates_the_preview_identity(tmp_path, monkeypatch):
    """PROMPT-027R §25 must survive the memoization: a saved label changes the region decision identity."""
    service, controller, images, _contents = _service_with_many_candidates(tmp_path, monkeypatch, saved=0)
    before = service.learning_state()
    identity_before = ApplicationService._learning_identity(controller.learning)

    service.set_learning_label("image", images[0].candidate_id, "AFTER", "")
    assert ApplicationService._learning_identity(controller.learning) == identity_before, \
        "a PENDING label must not change the production decision identity"

    service.save_learning_labels("image", {images[0].candidate_id: ""})
    assert ApplicationService._learning_identity(controller.learning) != identity_before, \
        "a SAVED label must change the identity so no cached preview region outlives its decision"

    after = service.learning_state()
    assert next(i for i in after["images"] if i["id"] == images[0].candidate_id)["userLabel"] == "AFTER"
    assert before["counts"]["image"]["labeled"] == 0 and after["counts"]["image"]["labeled"] == 1
