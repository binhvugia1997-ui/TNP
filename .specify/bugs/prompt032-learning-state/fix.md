# Bug Fix: Current Learning lifecycle and explicit candidate availability

- **Slug**: prompt032-learning-state
- **Date**: 2026-10-10
- **Starting SHA**: `2ca7dbfeb12a414933f0efc2b4f3f8fe6a501361`
- **Working branch**: `arena/2e088bee-tnp`
- **Assessment**: `.specify/bugs/prompt032-learning-state/assessment.md`
- **Spec Kit execution**: Executed from checked-in Spec Kit skill instructions; native Arena skill invocation was not available.

## Implemented fix

The fix follows the two independently proven failure classes and does not alter image/content eligibility.

### Backend lifecycle

`ApplicationService.learning_state()` now performs the existing bounded controller lifecycle sequence under the service `RLock`:

```text
pump → reconcile → pump
```

This is the same sequence already used by dashboard reads. A completed worker's queued terminal event is therefore applied by the Learning read itself. `review_candidates()` and `review_content_candidates()` no longer see a stale `running` state and suppress already-retained candidates while waiting for an unrelated dashboard poll.

No reload, processor recreation, persistence reconstruction, new polling interval, sleep, or fabricated candidate was added.

### Explicit image/content availability

Every successful Learning DTO now includes independent `candidateStatus.image` and `candidateStatus.content` records:

- `state`: `ready`, `processing`, `no_run`, or `unavailable`;
- `reason`: `available`, `empty`, `processing`, `no_run`, or `unavailable`;
- current DTO candidate `count`;
- a bounded visible Vietnamese message.

A successful `ready/empty` snapshot is now distinct from no completed run, active processing, local learning unavailability, a malformed bridge response, and a rejected bridge/backend request.

The frontend validates that both arrays and both status records are present and internally count-consistent. It no longer fills a missing `images` or `contents` field from an empty default. Refresh state is tracked separately as `idle/loading/ready/error`. On refresh failure, the previous successful data is retained, but both review paths show that the refresh failed so stale data cannot masquerade as current. Entering Learning and the existing manual button continue to issue the real `get_learning_state` bridge call; the backend method is now authoritative by itself.

### Preview failure

The existing production preview path already catches renderer failures and retains candidate/region truth where available. `_image_candidate_dto()` now also has an outer display-only guard so an unexpected preview helper exception cannot remove the candidate DTO. Identity, source display name, management number, decision, confidence, eligibility, target geometry, evidence, notes, and cropped-image fallback remain present; only full-slide preview fields become unavailable.

No PROMPT-027 geometry, item ownership, renderer order, fidelity flag, cache identity, report isolation, bitmap treatment, or Excel evidence selection was changed.

### Bounded diagnostics

- `LEARNING_CANDIDATES_RETAINED`: exactly once per processed report, with opaque report scope and raw/reviewable image/content counts.
- `LEARNING_STATE`: emitted only when lifecycle/count/availability signature changes, not on every active-tab poll.
- `LEARNING_IMAGE_PREVIEW_FAILED`: unexpected display-only preview exception with a hashed candidate identity and exception class.

These diagnostics contain no absolute source path, content/candidate payload, data URI, candidate text, evidence body, or image bytes.

## Changed files

| Path | Change |
|---|---|
| `app/batch_processor.py` | Added one report-scoped candidate-retention diagnostic after production candidate objects are copied to `FileResult`. |
| `app/application_service.py` | Reconciles lifecycle on Learning reads; adds independent availability DTOs and transition-bounded diagnostics; guards display-only image preview DTO generation. |
| `frontend/src/state/store.tsx` | Adds strict Learning DTO normalization plus independent load state; preserves last successful data while exposing failures. |
| `frontend/src/tabs/LearningTab.tsx` | Renders processing/no-run/valid-empty/unavailable/failure feedback independently for image and content; shows failed refreshes even when prior data remains visible. |
| `tests/test_prompt032_learning_state.py` | Adds production-lifecycle, availability, preview-failure, bridge/redaction, and diagnostic regressions. |
| `frontend/tests/learning-hooks.test.mjs` | Extends the mounted pywebview/store/Learning flow for entry refresh, manual refresh, image/content data, valid empty, missing fields, backend failure, and retry. |
| `frontend/tests/prompt027-overlays.test.mjs` | Supplies the new required availability contract and preserves existing hook/overlay regressions. |
| `.specify/bugs/prompt032-learning-state/assessment.md` | Pre-fix reproduction and independent root-cause assessment. |
| `.specify/bugs/prompt032-learning-state/fix.md` | This implementation record. |

## Scope deliberately unchanged

- Image and content candidate generation, deterministic decisions, hard exclusions, labels, and model training.
- Ollama use and fallback behavior.
- Report terminal statuses and Excel row/write semantics.
- Verbatim content rebuilding and temporary/inspection exclusion semantics.
- PROMPT-027/027R improvement regions, item-scoped ownership, overlays, renderer diagnostics/fallbacks, preview cache, unmodified bitmap display, and report isolation.
- PROMPT-031/031R scan locks, generation/revision snapshots, incremental leaf delivery, frontend merge behavior, polling cadence, and scan diagnostics.
- Version `1.3.4` / Build `017`.

## Linux limitation

The deterministic production fixture executes the actual PPTX parser/classifier/extractor/Excel pipeline and bridge/service/frontend contracts on Linux. It does not prove Windows WebView transport timing, PowerPoint COM rendering, or acceptance against the user's real source files. Those remain explicit Windows retest items in `test.md`.
