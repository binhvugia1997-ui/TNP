"""PROMPT-024R part B: Học cải tiến review-workspace DTO, geometry and preview-cache contracts.

Headless coverage of §51–§55: the learning DTO carries enough authoritative data (report / slide /
item / target bbox / slide size / prediction / confidence / eligibility) for the compact three-region
desktop layout; slide→preview geometry stays aligned under fit/resize/zoom/letterbox; PROMPT-025 item
identity drives the highlight; previews are cached per report+slide and never cross reports. The visual
highlight itself still needs the Windows manual check against a real slide.
"""
from __future__ import annotations

import json
from pathlib import Path
from types import SimpleNamespace

import pytest

from app.image_learning import ImageCandidate
from app.preview_geometry import (fit_slide_box, map_fraction_box_to_px, map_slide_box_to_preview,
                                  slide_fraction_box)
from tests.test_application_service import _learning_candidates, _make_service, _scan
from tests.test_prompt025_multi_item import NAME, _deck, _two_item_slide

EMU_PER_INCH = 914400


# ================================================================== §52 slide → preview geometry
def test_fraction_box_maps_emu_to_slide_fractions():
    slide_w, slide_h = 13.333 * EMU_PER_INCH, 7.5 * EMU_PER_INCH
    fx, fy, fw, fh = slide_fraction_box((1 * EMU_PER_INCH, 2 * EMU_PER_INCH,
                                         4 * EMU_PER_INCH, 3 * EMU_PER_INCH), slide_w, slide_h)
    assert fx == pytest.approx(1 / 13.333) and fy == pytest.approx(2 / 7.5)
    assert fw == pytest.approx(4 / 13.333) and fh == pytest.approx(3 / 7.5)
    # out-of-slide geometry is clamped – the highlight can never leave the preview
    fx, fy, fw, fh = slide_fraction_box((-1 * EMU_PER_INCH, -1 * EMU_PER_INCH,
                                          100 * EMU_PER_INCH, 100 * EMU_PER_INCH), slide_w, slide_h)
    assert (fx, fy, fw, fh) == (0.0, 0.0, 1.0, 1.0)
    # degenerate slide size never divides by zero and stays inside [0, 1]
    deg = slide_fraction_box((10, 10, 10, 10), 0, 0)
    assert all(0.0 <= v <= 1.0 for v in deg)


def test_fit_and_mapping_stay_aligned_through_resize_letterbox_zoom_and_multiple_boxes():
    slide_w, slide_h = 12192000, 6858000                     # 16:9 EMU
    box1 = (1219200, 685800, 2438400, 1371600)               # 10%,10% @ 20%,20%
    box2 = (6096000, 3429000, 1219200, 685800)               # 50%,50% @ 10%,10%

    def mapped(display_w, display_h, zoom=1.0):
        fit = fit_slide_box(display_w, display_h, slide_w, slide_h, zoom=zoom)
        f1 = slide_fraction_box(box1, slide_w, slide_h)
        f2 = slide_fraction_box(box2, slide_w, slide_h)
        return fit, map_fraction_box_to_px(f1, fit["width"], fit["height"]), \
            map_fraction_box_to_px(f2, fit["width"], fit["height"])

    # normal fit: 16:9 slide in a 16:9 container fills it without letterboxing
    fit, m1, m2 = mapped(1600, 900)
    assert abs(fit["width"] - 1600) < 1e-6 and abs(fit["height"] - 900) < 1e-6
    assert abs(m1["x"] - 160) < 1e-6 and abs(m1["y"] - 90) < 1e-6
    assert abs(m1["width"] - 320) < 1e-6 and abs(m1["height"] - 180) < 1e-6
    assert abs(m2["x"] - 800) < 1e-6 and abs(m2["y"] - 450) < 1e-6

    # resized center panel: fractions of the slide are invariant, so the highlight follows
    fit_r, r1, r2 = mapped(800, 450)
    assert abs(r1["x"] / fit_r["width"] - m1["x"] / fit["width"]) < 1e-9
    assert abs(r2["width"] / fit_r["width"] - m2["width"] / fit["width"]) < 1e-9

    # letterboxing: a taller container centres the slide vertically (offset > 0), width still drives x
    fit_l, l1, _ = mapped(800, 800)
    assert abs(fit_l["width"] - 800) < 1e-6 and fit_l["height"] < 800
    assert fit_l["offset_y"] > 0 and fit_l["offset_x"] == 0.0
    assert abs(l1["x"] - 80) < 1e-6                            # x mapping unaffected by vertical bars

    # zoom multiplies the fitted size uniformly; the relative highlight position is unchanged
    fit_z, z1, _ = mapped(1600, 900, zoom=2.0)
    assert abs(fit_z["width"] - 3200) < 1e-6 and abs(fit_z["height"] - 1800) < 1e-6
    assert abs(z1["x"] - 320) < 1e-6 and abs(z1["width"] - 640) < 1e-6

    # one-shot mapping used by tooling agrees with the two-step form
    direct = map_slide_box_to_preview(box2, slide_w, slide_h, fit["width"], fit["height"])
    assert abs(direct["x"] - m2["x"]) < 1e-9 and abs(direct["height"] - m2["height"]) < 1e-9


# ================================================================== §51/§27/§38 learning DTO
def _service_with_candidates(sample_tree, tmp_path, candidates):
    service, config = _make_service(sample_tree, tmp_path)
    _scan(service, config)
    service.controller.processor = SimpleNamespace(
        results=[SimpleNamespace(image_candidates=candidates, content_candidates=[])])
    return service, config


def test_learning_dto_carries_full_slide_context_and_authoritative_target(sample_tree, tmp_path):
    service, _ = _service_with_candidates(sample_tree, tmp_path, [_learning_candidates(sample_tree)[0]])
    state = service.learning_state()
    assert state["images"], "expected one reviewable image candidate"
    dto = state["images"][0]

    # authoritative geometry: slide size + EMU target box + fractions + kind (§27/§38)
    assert dto["slideWidth"] == 1000 and dto["slideHeight"] == 800
    assert dto["targetBbox"] == {"x": 100, "y": 100, "width": 300, "height": 200}
    assert dto["targetKind"] == "picture"
    assert dto["targetBboxPct"] == {"x": 10.0, "y": 12.5, "w": 30.0, "h": 25.0}
    # prediction / confidence / eligibility remain Python-owned business facts
    assert dto["decision"] and dto["confidenceBand"] in ("high", "medium", "low")
    assert dto["excelEligible"] in (True, False) and "eligibilityReason" in dto
    # full authored slide rendered as a bridge-safe data URI (§23/§39), never a filesystem path
    assert dto["slidePreview"].startswith("data:image/jpeg;base64,")
    assert dto["slidePreviewWidth"] > 0 and dto["slidePreviewHeight"] > 0
    # §39: no raw source path leaks anywhere in the learning DTO
    dump = json.dumps(state, ensure_ascii=False)
    assert str(Path(sample_tree["files"][0]).resolve()) not in dump
    assert "sample" not in dump.lower() or str(sample_tree["reports"]) not in dump


def test_learning_dto_multi_item_same_slide_keeps_item_identity_and_target(sample_tree, tmp_path):
    deck = _deck(tmp_path / "in" / NAME, [_two_item_slide()])
    W, H = int(13.333 * EMU_PER_INCH), int(7.5 * EMU_PER_INCH)
    item1 = ImageCandidate(
        candidate_id="260920045-VOC|S3|#11|1", management_number="260920045-VOC",
        source_name=deck.name, source_file=str(deck), slide=3, picture_id=11, order=1,
        bounds=(1554480, 2514600, 1188720, 777240), slide_size=(W, H), block_bounds=(0, 0, W, H),
        features={}, evidence=["item one"], confidence=0.9,
        logical_item_owner="item:one", owner_heading="Mục một", item_index=0)
    item2 = ImageCandidate(
        candidate_id="260920045-VOC|S3|#16|2", management_number="260920045-VOC",
        source_name=deck.name, source_file=str(deck), slide=3, picture_id=16, order=2,
        bounds=(8503920, 2514600, 1188720, 777240), slide_size=(W, H), block_bounds=(0, 0, W, H),
        features={}, evidence=["item two"], confidence=0.9,
        logical_item_owner="item:two", owner_heading="Mục hai", item_index=1)

    service, _ = _service_with_candidates(sample_tree, tmp_path, [item1, item2])
    state = service.learning_state()
    d1, d2 = state["images"]

    # §53: same report + same slide, different PROMPT-025 items -> distinct identity and geometry
    assert d1["itemId"] != d2["itemId"] and d1["itemIndex"] == 0 and d2["itemIndex"] == 1
    assert d1["itemHeading"] == "Mục một" and d2["itemHeading"] == "Mục hai"
    assert d1["targetBbox"]["x"] != d2["targetBbox"]["x"]
    assert d1["slide"] == d2["slide"] == 3
    # §40/§41: ONE rendered slide shared by both items – switching candidates only moves the highlight
    assert d1["slidePreview"] and d1["slidePreview"] == d2["slidePreview"]
    assert service._slide_preview_cache and len(service._slide_preview_cache) == 1


# ================================================================== §54 report/cache isolation
def test_slide_preview_cache_never_shared_between_reports(sample_tree, tmp_path):
    deck_a = _deck(tmp_path / "a" / NAME, [_two_item_slide()])
    deck_b = _deck(tmp_path / "b" / NAME, [_two_item_slide()])   # structurally identical layout
    W, H = int(13.333 * EMU_PER_INCH), int(7.5 * EMU_PER_INCH)

    def cand(deck, picture_id):
        return ImageCandidate(
            candidate_id=f"260920045-VOC|S3|#{picture_id}|1", management_number="260920045-VOC",
            source_name=deck.name, source_file=str(deck), slide=3, picture_id=picture_id, order=1,
            bounds=(1554480, 2514600, 1188720, 777240), slide_size=(W, H), block_bounds=(0, 0, W, H),
            features={}, evidence=[], confidence=0.9)

    service, _ = _service_with_candidates(sample_tree, tmp_path, [cand(deck_a, 11), cand(deck_b, 11)])
    state = service.learning_state()
    d_a, d_b = state["images"]
    # identical layouts still produce TWO cache entries keyed by report path – no cross-report reuse
    assert len(service._slide_preview_cache) == 2
    assert d_a["slidePreview"].startswith("data:image/") and d_b["slidePreview"].startswith("data:image/")


# ================================================================== §55 persistence with the new DTO
def test_labels_still_persist_and_reload_with_extended_dto(sample_tree, tmp_path):
    service, _ = _service_with_candidates(sample_tree, tmp_path, [_learning_candidates(sample_tree)[0]])
    candidate_id = "260918080-VOC|S5|#21|1"
    service.set_learning_label("image", candidate_id, "AFTER", "review note 024R")
    saved = service.save_learning_labels("image", {candidate_id: "review note 024R"})
    assert saved["state"]["images"][0]["userLabel"] == "AFTER"
    assert saved["state"]["images"][0]["slidePreview"].startswith("data:image/")
    # reload from the persisted JSONL store: label + note survive, extended geometry stays present
    reloaded = service.learning_state()
    dto = reloaded["images"][0]
    assert dto["userLabel"] == "AFTER" and dto["note"] == "review note 024R"
    assert dto["targetBbox"] == {"x": 100, "y": 100, "width": 300, "height": 200}
