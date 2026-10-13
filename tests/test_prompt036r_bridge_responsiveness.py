"""PROMPT-036R regression: desktop hangs after pywebviewready / BRIDGE_UNAVAILABLE / Learning UI unresponsive.

Root cause: ``learning_state()`` rendered slide previews synchronously (PPTX parsing + image encoding)
on the pywebview bridge thread, blocking the WinForms UI thread on Windows EdgeChromium.

These tests prove:
1. learning_state(render_previews=False) does NOT call expensive rendering.
2. learning_state(render_previews=True) still renders (backward compat).
3. set_learning_label uses the fast path.
4. Background prerender is scheduled after render_previews=False.
5. Bridge _respond validates JSON safety.
6. No unhandled recursion with many candidates.
7. dashboard_state does not call learning_state.
"""
from __future__ import annotations

import json
import sys
import time
from pathlib import Path
from unittest.mock import MagicMock, PropertyMock, patch

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))


class MockCandidate:
    """Minimal mock for an image candidate."""
    def __init__(self, candidate_id: str):
        self.candidate_id = candidate_id
        self.source_file = ""
        self.slide = 1
        self.picture_id = 1
        self.bounds = (100, 100, 300, 200)
        self.slide_size = (9144000, 6858000)
        self.management_number = "12345"
        self.source_name = "test.pptx"
        self.decision = "Sau cải tiến"
        self.confidence_band = "high"
        self.confidence = 0.9
        self.evidence = ["test evidence"]
        self.excel_output_eligible = True
        self.eligibility_reason = ""
        self.nearby_text = "test"
        self.logical_item_owner = ""
        self.item_index = 0
        self.owner_heading = ""
        self.hard_excluded = False


def _make_mock_learning():
    lrn = MagicMock()
    lrn.store.latest.return_value = {}
    lrn.content.store.latest.return_value = {}
    lrn.model = None
    lrn.model_status = ""
    lrn.content.model = None
    lrn.content.model_status = ""
    return lrn


class TestRenderPreviewsParameter:
    """Prove that render_previews=False skips expensive work."""

    def test_render_previews_false_skips_rendering(self, tmp_path):
        from app.application_service import ApplicationService
        from app.gui_controller import GuiController

        config_path = tmp_path / "config.json"
        config_path.write_text("{}", encoding="utf-8")
        controller = GuiController(config_path=config_path)
        service = ApplicationService(controller=controller, config_path=config_path)

        render_calls = []

        def mock_preview(candidate, learning=None, render_if_missing=True):
            render_calls.append(render_if_missing)
            return {"src": "", "width": 0, "height": 0, "backend": "", "faithful": False, "regions": []}

        service._slide_preview_for = mock_preview
        service._schedule_preview_prerender = lambda c, l: None  # prevent background thread

        mock_candidate = MockCandidate("test-1")
        mock_learning = _make_mock_learning()

        with patch.object(type(controller), 'learning', new_callable=PropertyMock, return_value=mock_learning):
            controller.review_candidates = lambda: [mock_candidate]
            controller.review_content_candidates = lambda: []
            controller.processor = None

            result = service.learning_state(render_previews=False)

        assert all(not call for call in render_calls), "render_if_missing=True must not be called"
        assert "images" in result

    def test_render_previews_true_renders(self, tmp_path):
        from app.application_service import ApplicationService
        from app.gui_controller import GuiController

        config_path = tmp_path / "config.json"
        config_path.write_text("{}", encoding="utf-8")
        controller = GuiController(config_path=config_path)
        service = ApplicationService(controller=controller, config_path=config_path)

        render_calls = []

        def mock_preview(candidate, learning=None, render_if_missing=True):
            render_calls.append(render_if_missing)
            return {"src": "data:image/jpeg;base64,test", "width": 100, "height": 100,
                    "backend": "test", "faithful": True, "regions": []}

        service._slide_preview_for = mock_preview

        mock_candidate = MockCandidate("test-1")
        mock_learning = _make_mock_learning()

        with patch.object(type(controller), 'learning', new_callable=PropertyMock, return_value=mock_learning):
            controller.review_candidates = lambda: [mock_candidate]
            controller.review_content_candidates = lambda: []
            controller.processor = None

            result = service.learning_state(render_previews=True)

        assert all(call for call in render_calls), "render_if_missing=True must be called"
        assert result["images"][0]["slidePreview"].startswith("data:image/")

    def test_set_learning_label_uses_fast_path(self, tmp_path):
        from app.application_service import ApplicationService
        from app.gui_controller import GuiController

        config_path = tmp_path / "config.json"
        config_path.write_text("{}", encoding="utf-8")
        controller = GuiController(config_path=config_path)
        service = ApplicationService(controller=controller, config_path=config_path)

        learning_state_calls = []
        original = service.learning_state

        def mock_learning_state(render_previews=True):
            learning_state_calls.append(render_previews)
            return {"images": [], "contents": [], "candidateStatus": {}, "counts": {},
                    "modelStatus": {}, "excelPending": 0, "excelLastResult": {}}

        service.learning_state = mock_learning_state

        mock_candidate = MockCandidate("test-1")
        controller.review_candidates = lambda: [mock_candidate]

        service.set_learning_label("image", "test-1", "AFTER")

        assert len(learning_state_calls) == 1
        assert learning_state_calls[0] is False


class TestBackgroundPrerender:
    """Prove that background prerender is scheduled correctly."""

    def test_prerender_scheduled_when_render_previews_false(self, tmp_path):
        from app.application_service import ApplicationService
        from app.gui_controller import GuiController

        config_path = tmp_path / "config.json"
        config_path.write_text("{}", encoding="utf-8")
        controller = GuiController(config_path=config_path)
        service = ApplicationService(controller=controller, config_path=config_path)

        prerender_calls = []

        def mock_schedule(candidates, learning):
            prerender_calls.append((len(candidates), learning))

        service._schedule_preview_prerender = mock_schedule
        service._slide_preview_for = lambda c, l=None, render_if_missing=True: {
            "src": "", "width": 0, "height": 0, "backend": "", "faithful": False, "regions": []
        }

        mock_candidate = MockCandidate("test-1")
        mock_learning = _make_mock_learning()

        with patch.object(type(controller), 'learning', new_callable=PropertyMock, return_value=mock_learning):
            controller.review_candidates = lambda: [mock_candidate]
            controller.review_content_candidates = lambda: []
            controller.processor = None

            service.learning_state(render_previews=False)

        assert len(prerender_calls) == 1
        assert prerender_calls[0][0] == 1

    def test_prerender_not_scheduled_when_render_previews_true(self, tmp_path):
        from app.application_service import ApplicationService
        from app.gui_controller import GuiController

        config_path = tmp_path / "config.json"
        config_path.write_text("{}", encoding="utf-8")
        controller = GuiController(config_path=config_path)
        service = ApplicationService(controller=controller, config_path=config_path)

        prerender_calls = []

        def mock_schedule(candidates, learning):
            prerender_calls.append(len(candidates))

        service._schedule_preview_prerender = mock_schedule
        service._slide_preview_for = lambda c, l=None, render_if_missing=True: {
            "src": "", "width": 0, "height": 0, "backend": "", "faithful": False, "regions": []
        }

        mock_candidate = MockCandidate("test-1")
        mock_learning = _make_mock_learning()

        with patch.object(type(controller), 'learning', new_callable=PropertyMock, return_value=mock_learning):
            controller.review_candidates = lambda: [mock_candidate]
            controller.review_content_candidates = lambda: []
            controller.processor = None

            service.learning_state(render_previews=True)

        assert len(prerender_calls) == 0


class TestJSONSafety:
    """Prove that bridge _respond validates JSON safety."""

    def test_non_json_safe_object_returns_error(self, tmp_path):
        from app.bridge import BridgeService
        from app.application_service import ApplicationService
        from app.gui_controller import GuiController

        config_path = tmp_path / "config.json"
        config_path.write_text("{}", encoding="utf-8")
        controller = GuiController(config_path=config_path)
        service = ApplicationService(controller=controller, config_path=config_path)
        bridge = BridgeService(service)

        class NonJSONSafe:
            pass

        result = bridge._respond(lambda: {"unsafe": NonJSONSafe()})

        assert result["ok"] is False
        assert result["error"]["code"] == "INTERNAL_ERROR"

    def test_cyclic_reference_returns_error(self, tmp_path):
        from app.bridge import BridgeService
        from app.application_service import ApplicationService
        from app.gui_controller import GuiController

        config_path = tmp_path / "config.json"
        config_path.write_text("{}", encoding="utf-8")
        controller = GuiController(config_path=config_path)
        service = ApplicationService(controller=controller, config_path=config_path)
        bridge = BridgeService(service)

        cyclic = {"a": 1}
        cyclic["self"] = cyclic

        result = bridge._respond(lambda: cyclic)

        assert result["ok"] is False
        assert result["error"]["code"] == "INTERNAL_ERROR"

    def test_normal_json_passes(self, tmp_path):
        from app.bridge import BridgeService
        from app.application_service import ApplicationService
        from app.gui_controller import GuiController

        config_path = tmp_path / "config.json"
        config_path.write_text("{}", encoding="utf-8")
        controller = GuiController(config_path=config_path)
        service = ApplicationService(controller=controller, config_path=config_path)
        bridge = BridgeService(service)

        result = bridge._respond(lambda: {"safe": "data", "number": 42, "list": [1, 2, 3]})

        assert result["ok"] is True
        assert result["data"]["safe"] == "data"


class TestNoRecursion:
    """Prove that bridge calls do not cause unbounded recursion."""

    def test_many_candidates_no_recursion(self, tmp_path):
        from app.application_service import ApplicationService
        from app.gui_controller import GuiController

        config_path = tmp_path / "config.json"
        config_path.write_text("{}", encoding="utf-8")
        controller = GuiController(config_path=config_path)
        service = ApplicationService(controller=controller, config_path=config_path)

        candidates = [MockCandidate(f"test-{i}") for i in range(100)]
        mock_learning = _make_mock_learning()

        service._slide_preview_for = lambda c, l=None, render_if_missing=True: {
            "src": "", "width": 0, "height": 0, "backend": "", "faithful": False, "regions": []
        }
        service._schedule_preview_prerender = lambda c, l: None

        with patch.object(type(controller), 'learning', new_callable=PropertyMock, return_value=mock_learning):
            controller.review_candidates = lambda: candidates
            controller.review_content_candidates = lambda: []
            controller.processor = None

            old_limit = sys.getrecursionlimit()
            sys.setrecursionlimit(200)
            try:
                result = service.learning_state(render_previews=False)
                assert len(result["images"]) == 100
            finally:
                sys.setrecursionlimit(old_limit)


class TestDashboardIsolation:
    """Prove that dashboard_state does not call learning_state."""

    def test_dashboard_does_not_call_learning(self, tmp_path):
        from app.application_service import ApplicationService
        from app.gui_controller import GuiController

        config_path = tmp_path / "config.json"
        config_path.write_text("{}", encoding="utf-8")
        controller = GuiController(config_path=config_path)
        service = ApplicationService(controller=controller, config_path=config_path)

        learning_calls = []
        original = service.learning_state

        def mock_learning_state(render_previews=True):
            learning_calls.append(render_previews)
            return {}

        service.learning_state = mock_learning_state

        service.dashboard_state()

        assert len(learning_calls) == 0


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
