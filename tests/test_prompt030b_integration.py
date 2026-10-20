"""PR #14 behavior reimplemented on 9faba26. Linux/mocked COM is NOT Windows acceptance."""
from __future__ import annotations

import io
import json
import os
from dataclasses import replace
from pathlib import Path
from types import SimpleNamespace

import pytest
from PIL import Image

import app.application_service as service_module
import app.learning_assets as assets_module
import app.qpn_renderer as renderer_module
from app.application_service import ApplicationService, ServiceError
from app.bridge import BridgeService
from app.config import AppConfig
from app.gui_controller import GuiController
from app.image_learning import ImageCandidate
from app.learning_assets import LearningAssets, asset_key, review_id
from app import updater
from tests.test_prompt032_learning_state import _finished_production_service
from tests.test_prompt025_multi_item import _deck, _two_item_slide, NAME
from tests.test_updater import _staged, set_helper_runtime


def _service(root):
    controller = GuiController(AppConfig(), config_path=root / "config.json",
                               learning_dir_override=root / "learning_data")
    return ApplicationService(controller, config_path=root / "config.json", manual_fields_path=root / "fields.json")


def _candidate(source, *, slide=3, shape=11):
    return ImageCandidate(f"M|S{slide}|#{shape}|1", "M", source.name, str(source), slide, shape, 1,
                          (100, 200, 300, 400), (1000, 1000), (100, 200, 300, 400), {},
                          logical_item_owner="item:one", owner_heading="Mục một", item_index=0)


def _install(service, *candidates):
    service.controller.processor = SimpleNamespace(results=[SimpleNamespace(image_candidates=list(candidates),
                                                                          content_candidates=[])])


def _uri(color):
    buf = io.BytesIO()
    Image.new("RGB", (80, 60), color).save(buf, format="JPEG")
    import base64
    return "data:image/jpeg;base64," + base64.b64encode(buf.getvalue()).decode("ascii")


@pytest.fixture(autouse=True)
def _deterministic_backend(monkeypatch):
    # Explicit test seam, never a claim about Office availability on the real host.
    monkeypatch.setattr(renderer_module, "powerpoint_available", lambda: False)
    monkeypatch.setattr(renderer_module, "find_soffice", lambda: None)


def test_real_production_flow_restores_unlabeled_and_confirmed_offline_with_geometry(tmp_path):
    service, controller, _processor, deck = _finished_production_service(tmp_path)
    first = service.learning_state()
    assert first["images"]
    row = next(r for r in first["images"] if r["excelEligible"])
    assert row["src"].startswith("data:image/jpeg;base64,")
    assert row["slidePreview"].startswith("data:image/jpeg;base64,")
    assert row["slidePreviewWidth"] == 1600 and row["slidePreviewHeight"] > 100
    assert row["slidePreviewBackend"] == "builtin" and not row["slidePreviewFaithful"]
    assert row["evidenceRegionBbox"] and row["itemIndex"] == 0

    # No label required: actual candidates are durable too, not just their eventual training rows.
    fresh = _service(tmp_path)
    assert fresh.controller.processor is None
    # The production fixture's learning directory is also tmp_path/learning_data.
    before_label = fresh.learning_state()
    restored = next(r for r in before_label["images"] if r["reviewId"] == row["reviewId"])
    assert restored["userLabel"] == "UNLABELED"
    assert restored["src"] == row["src"] and restored["slidePreview"] == row["slidePreview"]
    assert restored["evidenceRegionBbox"] == row["evidenceRegionBbox"]
    assert before_label["candidateStatus"]["image"]["state"] == "ready"

    ack = fresh.save_learning_candidate("image", row["reviewId"], "AFTER", "durable note")
    assert not ack["labelPending"] and "state" not in ack
    # Render under the final saved learning identity before going offline.
    saved = next(r for r in fresh.learning_state()["images"] if r["reviewId"] == row["reviewId"])
    deck.unlink()  # only this test's generated PPTX; production data is never touched
    final = _service(tmp_path).learning_state()
    actual = next(r for r in final["images"] if r["reviewId"] == row["reviewId"])
    for field in ("src", "slidePreview", "slidePreviewWidth", "slidePreviewHeight", "targetBboxPct",
                  "evidenceRegionBbox", "evidenceRegionPictureIds", "itemIndex", "slidePreviewBackend"):
        assert actual[field] == saved[field], field
    assert actual["userLabel"] == "AFTER" and actual["note"] == "durable note"
    assert actual["persisted"] is True and actual["itemIndex"] == 0
    assert str(tmp_path) not in json.dumps(final, ensure_ascii=False)
    assert controller.learning.store.counts()["total"] == 1

    # Relabel after restart without parsing missing PPTX or fabricating eligible evidence.
    restarted = _service(tmp_path)
    restarted.save_learning_candidate("image", row["reviewId"], "BEFORE", "corrected offline")
    offline = next(r for r in restarted.learning_state()["images"] if r["reviewId"] == row["reviewId"])
    assert offline["userLabel"] == "BEFORE" and offline["note"] == "corrected offline"
    assert offline["slidePreview"] == actual["slidePreview"]
    assert offline["evidenceRegionBbox"] is None, "a different learning decision cannot reuse old regions offline"


def test_scoped_pixels_review_identity_and_restart_dedup_do_not_change_classifier_ids(tmp_path, monkeypatch):
    a = _candidate(tmp_path / "a" / "same.pptx")
    b = replace(a, source_file=str(tmp_path / "b" / "same.pptx"))
    assert a.candidate_id == b.candidate_id   # deliberately shared classification ID remains compatible
    assert asset_key(a) != asset_key(b) and review_id(a) != review_id(b)
    monkeypatch.chdir(tmp_path)
    assert asset_key(replace(a, source_file="a/same.pptx")) == asset_key(a)
    service = _service(tmp_path)
    _install(service, a, b, a)  # repeated candidates dedup by evidence scope, not just shared classifier ID
    srcs = {str(a.source_file): _uri("red"), str(b.source_file): _uri("blue")}
    monkeypatch.setattr(service.controller, "candidate_blob", lambda c: b"")
    assets = LearningAssets(service.controller.learning.dir)
    for c in (a, b):
        assets.remember_candidate(c, srcs[c.source_file])
    first = service.learning_state()["images"]
    assert len(first) == 2 and first[0]["id"] == first[1]["id"]
    assert first[0]["src"] != first[1]["src"]
    files = {p.name: p.stat().st_mtime_ns for p in assets.root.iterdir()}
    service.learning_state()
    assert files == {p.name: p.stat().st_mtime_ns for p in assets.root.iterdir()}
    fresh = _service(tmp_path)
    restored = fresh.learning_state()["images"]
    assert {r["reviewId"]: r["src"] for r in first} == {r["reviewId"]: r["src"] for r in restored}
    with pytest.raises(ServiceError, match="Có nhiều nguồn"):
        fresh.set_learning_label_fast("image", a.candidate_id, "BEFORE")
    fresh.save_learning_candidate("image", review_id(a), "BEFORE", "selected a")
    assert fresh.controller.pending_image_reapply[0].source_file == a.source_file
    assert fresh.controller.learning.store.counts()["total"] == 1  # dataset IDs/dedup intentionally unchanged


def test_multiple_candidates_share_one_slide_asset_and_polls_do_not_rewrite(tmp_path, monkeypatch):
    deck = _deck(tmp_path / "in" / NAME, [_two_item_slide()])
    service = _service(tmp_path)
    a, b = _candidate(deck), _candidate(deck, shape=12)
    _install(service, a, b)
    monkeypatch.setattr(service.controller, "candidate_blob", lambda c: b"")
    rows = service.learning_state()["images"]
    assert rows[0]["slidePreview"] and rows[0]["slidePreview"] == rows[1]["slidePreview"]
    root = service.controller.learning.dir / "previews"
    assert len(list(root.glob("*_candidate.json"))) == 2
    assert len(list(root.glob("*_slide.json"))) == 1
    before = {p.name: (p.read_bytes(), p.stat().st_mtime_ns) for p in root.iterdir()}
    service.learning_state()
    assert before == {p.name: (p.read_bytes(), p.stat().st_mtime_ns) for p in root.iterdir()}
    assert not list(root.glob("*.tmp"))


@pytest.mark.parametrize("change", ["mtime", "size", "schema", "width", "availability", "label", "model"])
def test_persistent_cache_uses_accepted_identity_across_service_restart(tmp_path, monkeypatch, change):
    deck = _deck(tmp_path / "in" / NAME, [_two_item_slide()])
    candidate = _candidate(deck)
    original = renderer_module.SlideRenderer.render
    calls = []

    def counted(self, *args, **kwargs):
        calls.append(1)
        return original(self, *args, **kwargs)
    monkeypatch.setattr(renderer_module.SlideRenderer, "render", counted)
    first = _service(tmp_path)
    first._slide_preview_for(candidate, first.controller.learning)
    second = _service(tmp_path)
    assert second._slide_preview_for(candidate, second.controller.learning)["src"]
    assert len(calls) == 1, "fresh service must reuse valid on-disk identity"
    if change == "mtime":
        stat = deck.stat()
        os.utime(deck, ns=(stat.st_atime_ns, stat.st_mtime_ns + 1000000))
    elif change == "size":
        deck.write_bytes(deck.read_bytes() + b"zip trailing bytes")
    elif change == "schema":
        monkeypatch.setattr(renderer_module, "RENDER_SCHEMA_VERSION", renderer_module.RENDER_SCHEMA_VERSION + 1)
    elif change == "width":
        monkeypatch.setattr(service_module, "LEARNING_PREVIEW_WIDTH_PX", 1400)
    elif change == "availability":
        monkeypatch.setattr(renderer_module, "powerpoint_available", lambda: True)
        # A simulated failure (not a fake real-host availability claim) must cause a truthful fallback.
        def refuse(*args, **kwargs):
            raise renderer_module.PowerPointStageError("ownership", RuntimeError("test refusal"))
        monkeypatch.setattr(renderer_module, "render_with_powerpoint", refuse)
    elif change == "label":
        second.controller.learning.store.label(candidate, "BEFORE")
    else:
        second.controller.learning.model = SimpleNamespace(trained_at="2026-10-16T01:00:00", n_examples=12)
    result = second._slide_preview_for(candidate, second.controller.learning)
    assert result["src"] and len(calls) == 2, "a changed accepted identity must not reuse disk pixels/regions"
    if change == "availability":
        assert result["backend"] == "builtin" and not result["faithful"] and result["degraded"]


def test_ownership_refusal_diagnostic_and_degraded_ttl_survive_restart_then_recover(tmp_path, monkeypatch, caplog):
    deck = _deck(tmp_path / "in" / NAME, [_two_item_slide()])
    candidate = _candidate(deck)
    clock = {"now": 1000.0, "wall": 1000000.0}
    monkeypatch.setattr(assets_module, "time", SimpleNamespace(time=lambda: clock["wall"], monotonic=lambda: clock["now"]))
    monkeypatch.setattr(service_module, "time", SimpleNamespace(monotonic=lambda: clock["now"]))
    calls = []
    broken = {"value": True}
    def attempt(source, numbers, out_dir, width):
        calls.append(1)
        if broken["value"]:
            raise renderer_module.PowerPointStageError("ownership", RuntimeError("user instance; refused"))
        path = out_dir / "real-test-export.png"
        Image.new("RGB", (width, width // 2), "green").save(path)
        return {n: path for n in numbers}
    monkeypatch.setattr(renderer_module, "powerpoint_available", lambda: True)
    monkeypatch.setattr(renderer_module, "render_with_powerpoint", attempt)
    caplog.set_level("INFO")
    first = _service(tmp_path)
    result = first._slide_preview_for(candidate, first.controller.learning)
    assert result["backend"] == "builtin" and result["degraded"] and not result["faithful"]
    assert "stage=ownership" in caplog.text and "user instance; refused" in caplog.text
    second = _service(tmp_path)
    second._slide_preview_for(candidate, second.controller.learning)
    assert len(calls) == 1
    clock["now"] += service_module.DEGRADED_PREVIEW_TTL_S + 1
    clock["wall"] += service_module.DEGRADED_PREVIEW_TTL_S + 1
    broken["value"] = False
    third = _service(tmp_path)
    recovered = third._slide_preview_for(candidate, third.controller.learning)
    assert len(calls) == 2 and recovered["backend"] == "powerpoint" and recovered["faithful"]
    assert not recovered["degraded"]


@pytest.mark.parametrize("corruption", ["image", "snapshot", "geometry", "slide_metadata"])
def test_corrupt_or_missing_assets_keep_label_and_other_rows(tmp_path, corruption):
    service = _service(tmp_path)
    candidate = _candidate(tmp_path / "missing.pptx")
    other = _candidate(tmp_path / "other.pptx", shape=12)
    assets = LearningAssets(service.controller.learning.dir)
    assets.remember_candidate(candidate, _uri("red"))
    assets.remember_candidate(other, _uri("blue"))
    service.controller.learning.store.label(candidate, "AFTER", "keep even if corrupt")
    path = assets.root / (asset_key(candidate) + "_candidate.json")
    snapshot = json.loads(path.read_text())
    if corruption == "image":
        (assets.root / snapshot["image"]).write_bytes(b"corrupt")
    elif corruption == "snapshot":
        path.write_text("{broken", encoding="utf-8")
    elif corruption == "geometry":
        snapshot["candidate"]["bounds"] = [1]
        path.write_text(json.dumps(snapshot), encoding="utf-8")
    else:
        snapshot["slide"] = "../../private-path"
        path.write_text(json.dumps(snapshot), encoding="utf-8")
    rows = _service(tmp_path).learning_state()["images"]
    assert len(rows) == 2
    bad = next(r for r in rows if r["id"] == candidate.candidate_id)
    good = next(r for r in rows if r["id"] == other.candidate_id)
    assert bad["userLabel"] == "AFTER" and bad["note"] == "keep even if corrupt"
    assert bad["slidePreview"] == ""
    if corruption != "slide_metadata":
        assert bad["src"] == ""
    assert good["src"].startswith("data:image/jpeg;")
    assert service.controller.learning.store.counts()["total"] == 1
    assert str(tmp_path) not in json.dumps(rows)


def test_legacy_source_path_is_not_trusted_and_missing_preview_does_not_fabricate_geometry(tmp_path):
    service = _service(tmp_path)
    store = service.controller.learning.store
    store.dir.mkdir(parents=True)
    record = {"candidate_id": "LEGACY|S1|#7|1", "label": "BEFORE", "schema_version": 1,
              "source_file": "/tmp/old_session/preview.png", "source_name": "C:\\private\\old.pptx"}
    store.path.write_text(json.dumps(record) + "\n", encoding="utf-8")
    row = service.learning_state()["images"][0]
    assert row["sourceFile"] == "old.pptx" and row["src"] == "" and row["slidePreview"] == ""
    assert row["evidenceRegionBbox"] is None and not row["excelEligible"]
    assert row["targetBboxPct"] == {"x": 0.0, "y": 0.0, "w": 0.0, "h": 0.0}
    assert "private" not in json.dumps(row) and "old_session" not in json.dumps(row)
    service.save_learning_candidate("image", row["reviewId"], "CONTROL", "legacy correction")
    assert store.latest()[record["candidate_id"]]["label"] == "CONTROL"


def test_atomic_metadata_failure_keeps_previous_pair_and_no_temp_files(tmp_path, monkeypatch):
    candidate = _candidate(tmp_path / "r.pptx")
    assets = LearningAssets(tmp_path / "learning_data")
    assets.remember_candidate(candidate, _uri("red"))
    previous = assets.candidate_src(candidate)
    original = Path.replace
    def fail_metadata(self, target):
        if str(target).endswith("_candidate.json"):
            raise PermissionError("simulated disk failure")
        return original(self, target)
    monkeypatch.setattr(Path, "replace", fail_metadata)
    assets.remember_candidate(candidate, _uri("blue"))
    assert assets.candidate_src(candidate) == previous
    assert not list(assets.root.glob("*.tmp"))


def test_current_empty_run_and_processing_do_not_resurrect_archived_candidates(tmp_path):
    service = _service(tmp_path)
    candidate = _candidate(tmp_path / "r.pptx")
    LearningAssets(service.controller.learning.dir).remember_candidate(candidate, _uri("red"))
    assert service.learning_state()["images"]
    service.controller.processor = SimpleNamespace(results=[])
    assert service.learning_state()["images"] == []
    service.controller.processor.is_running = lambda: True
    service.controller.state = "running"
    state = service.learning_state()
    assert state["images"] == [] and state["candidateStatus"]["image"]["state"] == "processing"


def test_changed_live_source_requires_new_run_not_mismatched_snapshot_overlays(tmp_path):
    deck = _deck(tmp_path / "in" / NAME, [_two_item_slide()])
    service = _service(tmp_path)
    candidate = _candidate(deck)
    _install(service, candidate)
    service.learning_state()
    stat = deck.stat()
    os.utime(deck, ns=(stat.st_atime_ns, stat.st_mtime_ns + 1000000))
    row = _service(tmp_path).learning_state()["images"][0]
    assert not row["src"] and not row["slidePreview"] and not row["excelEligible"]
    assert row["evidenceRegionBbox"] is None


def test_bulk_ack_prevalidates_scope_and_preserves_labels_first_contract(tmp_path, monkeypatch):
    service = _service(tmp_path)
    a, b = _candidate(tmp_path / "r.pptx"), _candidate(tmp_path / "r.pptx", shape=12)
    other = replace(b, source_file=str(tmp_path / "elsewhere" / "r.pptx"))
    _install(service, a, b, other)
    bridge = BridgeService(service)
    bad = bridge.save_learning_images([review_id(a), review_id(other)], "AFTER")
    assert bad["ok"] is False and bad["error"]["code"] == "CROSS_SLIDE_SELECTION"
    assert service.controller.pending_labels == {} and service.controller.learning.store.records() == []
    missing = bridge.save_learning_images([review_id(a), "missing"], "AFTER")
    assert missing["ok"] is False and service.controller.pending_labels == {}
    invalid_note = bridge.save_learning_images([review_id(a), review_id(b)], "AFTER", {review_id(b): "x" * 1001})
    assert not invalid_note["ok"] and not service.controller.pending_labels
    monkeypatch.setattr(service, "learning_state", lambda: pytest.fail("batch save must not rebuild full state"))
    saved = bridge.save_learning_images([review_id(a), review_id(b), review_id(a)], "AFTER",
                                        {review_id(a): "one", review_id(b): "two"})
    assert saved["ok"] and len(saved["data"]["candidateIds"]) == 2 and "state" not in saved["data"]
    assert saved["data"]["excelPending"] == 2  # labels persisted; no Excel in this temporary fixture
    assert {cid: r["note"] for cid, r in service.controller.learning.store.latest().items()} == {
        a.candidate_id: "one", b.candidate_id: "two"}
    assert len(json.dumps(saved)) < 2000
    bridge.save_learning_images([review_id(a), review_id(b)], "AFTER")
    assert len(service.controller.learning.store.records()) == 2  # duplicate-label protection retained


def test_bulk_disk_failure_is_reported_with_pending_choices_retained_for_retry(tmp_path, monkeypatch):
    service = _service(tmp_path)
    a, b = _candidate(tmp_path / "r.pptx"), _candidate(tmp_path / "r.pptx", shape=12)
    _install(service, a, b)
    store = service.controller.learning.store
    original = store.label
    def fail_second(candidate, *args):
        if candidate.picture_id == b.picture_id:
            raise OSError("simulated disk full")
        return original(candidate, *args)
    monkeypatch.setattr(store, "label", fail_second)
    bridge = BridgeService(service)
    result = bridge.save_learning_images([review_id(a), review_id(b)], "BEFORE")
    assert not result["ok"]
    assert store.current_label(a.candidate_id) == "BEFORE"  # partial write is not mislabeled as all-or-none
    assert set(service.controller.pending_labels) == {a.candidate_id, b.candidate_id}
    monkeypatch.setattr(store, "label", original)
    assert bridge.save_learning_images([review_id(a), review_id(b)], "BEFORE")["ok"]
    assert not service.controller.pending_labels and len(store.records()) == 2


@pytest.mark.parametrize("rollback", [False, True])
def test_actual_validated_updater_apply_and_rollback_preserve_all_new_assets(tmp_path, monkeypatch, rollback):
    portable, staged = _staged(tmp_path, monkeypatch, top=False)
    directory = portable / "learning_data"
    assets = LearningAssets(directory)
    candidate = _candidate(tmp_path / "r.pptx")
    slide = assets.remember_slide(("scope", "source", 3, 0, 0, ("builtin",), 1600, 3, (), ()),
                                  {"src": _uri("green"), "backend": "builtin", "faithful": False,
                                   "regions": [], "degraded": False})
    assets.remember_candidate(candidate, _uri("red"), slide)
    assert "learning_data" not in updater.release_entries(Path(staged.source_root))
    service = _service(tmp_path)
    service.controller.learning.store.label(candidate, "AFTER")
    (directory / "image_labels.jsonl").write_bytes(service.controller.learning.store.path.read_bytes())
    before = {str(p.relative_to(directory)): p.read_bytes() for p in directory.rglob("*") if p.is_file()}
    set_helper_runtime(monkeypatch, staged)
    if rollback:
        original = updater.os.replace
        def fail_runtime(src, dst):
            if Path(dst) == portable / "_internal":
                raise OSError(5, "simulated locked runtime file")
            return original(src, dst)
        monkeypatch.setattr(updater.os, "replace", fail_runtime)
    rc = updater.apply_update(Path(staged.manifest_path), alive=lambda pid: False,
                              sleep=lambda seconds: None, spawn=lambda cmd, **kwargs: None)
    assert rc == (1 if rollback else 0)
    after = {str(p.relative_to(directory)): p.read_bytes() for p in directory.rglob("*") if p.is_file()}
    assert before == after and before
    assert LearningAssets(directory).candidate_src(candidate) == _uri("red")
    assert updater.is_preserved("Learning_Data")
    assert LearningAssets(directory).offline_slide(candidate, ())["src"] == _uri("green")
    assert (portable / updater.APP_EXE).read_bytes().startswith(b"MZ-old" if rollback else b"MZ-new-")
