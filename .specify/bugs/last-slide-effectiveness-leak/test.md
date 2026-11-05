# Bug Verification: Last slide "Hiệu quả cải tiến" still contributes its content (last-slide-effectiveness-leak)

- **Slug**: last-slide-effectiveness-leak
- **Tested**: 2026-11-04
- **Assessment**: ./assessment.md
- **Fix**: ./fix.md
- **Result**: **partial** — automated verification passes on Linux with simulated slide fixtures (no real PPTX deck with this shape was available; fixtures are simulated `SlideData`/`Block` objects shaped like real decks, and the A1285 real-deck regression test passes). Windows acceptance with a real deck still required.
- **Spec Kit execution mode**: Executed from the checked-in Spec Kit skill instructions; native Arena skill invocation was not available.
- **Environment**: Linux sandbox, `.venv-build` in the snapshot tree, pytest.

## Checks performed

| # | Check | Result |
|---|---|---|
| 1 | Reproduction probe (6 simulated shapes, pre-fix) | S3 (title placeholder + improvement-like first body heading) leaked the whole effectiveness body as `improvement`; S1/S2/S4/S5/S6 excluded correctly. The exclusion vocabulary DID exist — this was a segmentation bypass, not a missing feature. |
| 2 | Probe re-run (post-fix) | all 6 shapes excluded |
| 3 | New regressions `tests/test_effectiveness_slide_exclusion.py` | **8 passed** (end-to-end heuristic; LLM-style dual-listed classification; body-before-title; trailing block splits; picture bands; mid-slide block closes; `first_heading_kind`; header-band A1285 shape) |
| 4 | Correction regression | The first fix over-excluded the real A1285 deck's slide 8 (header-band effectiveness heading + later long-term heading): `test_real_a1285.py::test_a1285_record_contents` caught it. Restricted title-led stickiness to real title placeholders; A1285 passes again and the new header-band test pins the shape. See the fix record's Correction section. |
| 5 | Segmentation contracts | `tests/test_prompt001.py` + `tests/test_real_a1285.py` pass unchanged (39 passed across the three files in the final run) |
| 6 | Full Python suite | **1404 passed, 1 failed, 13 skipped** — the single failure is the pre-existing `pythonnet` environment probe, unrelated to this change |

## Not verified here

A real countermeasure deck whose last slide is "Hiệu quả cải tiến" (title placeholder + countermeasure-restating body, S3 shape) through the cloud Windows build: the Excel improvement field must contain zero effectiveness lines and no After picture may come from that slide.
