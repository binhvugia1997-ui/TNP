# Bug Assessment: Learning state can suppress retained candidates and cannot represent empty versus failed loads

- **Slug**: prompt032-learning-state
- **Created**: 2026-10-10
- **Source**: pasted text (PROMPT-032)
- **Verdict**: valid
- **Severity**: high
- **Starting SHA**: `2ca7dbfeb12a414933f0efc2b4f3f8fe6a501361`
- **Working branch**: `arena/2e088bee-tnp`
- **Implementation status**: assessment complete; no production or test source modified before this artifact
- **Spec Kit execution**: Executed from checked-in Spec Kit skill instructions; native Arena skill invocation was not available.

## Report (summarized)

After a real completed Windows processing run, the image-learning view can say that Python returned no candidate and the content-learning view provides no truthful visible response. PROMPT-032 requires the actual candidate lifecycle to be traced and reproduced before a minimal fix, with image and content failures classified independently. Entering Learning and manual refresh must read current backend state without restarting; valid empty state, missing fields, and backend failure must remain distinct. Image preview failure must not erase candidate metadata. PROMPT-027 preview/evidence semantics, report isolation, and PROMPT-031/031R scan behavior must remain unchanged.

## End-to-end production paths

### Image candidates

`extract_record(..., learning=...)` calls `select_with_learning()` → `ExtractedRecord.image_candidates` → `BatchProcessor._process_one()` copies the objects to `FileResult.image_candidates` → the completed processor retains `FileResult` in `processor.results` → `GuiController.review_candidates()` filters only hard exclusions → `ApplicationService.learning_state()` creates DTOs with `_image_candidate_dto()` → `BridgeService.get_learning_state()` returns the pywebview envelope → `store.refreshLearning()` assigns `learning.images` → `LearningTab.ImageReview` renders the candidate and PROMPT-027 preview/evidence overlays.

### Content candidates

`extract_record(..., learning=...)` calls `select_content_with_learning()` → `ExtractedRecord.content_candidates` → `BatchProcessor._process_one()` copies the objects to `FileResult.content_candidates` → the completed processor retains `FileResult` in `processor.results` → `GuiController.review_content_candidates()` filters only hard exclusions → `ApplicationService.learning_state()` creates DTOs with `_content_candidate_dto()` → the same bridge/store call assigns `learning.contents` → `LearningTab.ContentReview` renders verbatim text and label controls.

Both candidate sets are generated without requiring Ollama: the local deterministic decision paths remain the baseline. Neither candidate set is persisted as a reconstructable review snapshot; the authoritative current review objects are the latest retained in-memory `processor.results`.

## Deterministic pre-fix reproduction

A temporary fixture used the production deck builder, parser/classifier/extractor, `BatchProcessor`, `GuiController`, `ApplicationService`, bridge-shaped state, and a real XLSX writer. The deck contained a production improvement section with eligible After picture evidence and reviewable improvement-content blocks. Ollama was disabled. The resulting report was `needs_review`, which is an explicitly required terminal state for this reproduction.

The first control run proved candidate generation and retention themselves are working:

```text
status needs_review
image all 2 image reviewable 1 image eligible 1
content all 8 content reviewable 3
dto 1 3
```

The complete asynchronous service lifecycle then reproduced the defect before any fix. After joining the real worker, but before any dashboard request happened to drain the controller event queue:

```text
worker_done True
controller_state_before_pump running
queue_empty False
raw 2 8
learning_before_pump 0 0 {'image': {'total': 0, 'labeled': 0}, 'content': {'total': 0, 'labeled': 0}}
dashboard done
controller_state_after_pump idle
learning_after_pump 1 3 {'image': {'total': 1, 'labeled': 0}, 'content': {'total': 3, 'labeled': 0}}
```

This is not fabricated candidate loss. The same processor retained two raw image candidates and eight raw content candidates. A Learning read returned zero only because the controller still had an unapplied terminal event. An unrelated dashboard read drained that event and made one image plus three content candidates immediately visible.

No repository source was modified by either reproduction.

## Root cause findings

**Confidence: high.** The two visible symptoms have separate proven failure classes even though both can occur in the same user flow.

### 1. Image symptom: stale lifecycle gating hides retained candidates

`BatchProcessor` queues its terminal callback. Worker completion therefore precedes GUI/controller finalization. `ApplicationService.dashboard_state()` explicitly performs `pump()` → `reconcile()` → `pump()` before reading state, but `ApplicationService.learning_state()` performs none of those lifecycle steps. Both controller review methods return an empty list whenever `controller.is_running()` is true.

Consequently a direct Learning read (including its manual refresh) can observe a dead/completed worker plus retained candidates while the controller still says `running`, and report zero candidates. The candidates are present, not absent, filtered, discarded, or omitted by extraction. The exact failure class is **stale controller lifecycle state suppressing retained image candidates at the review gate**.

This also proves why depending on a later periodic dashboard request is not a valid Learning refresh contract. Manual refresh currently calls only `get_learning_state`; it does not itself establish current backend lifecycle state.

### 2. Content symptom: the frontend has no content availability/failure contract

The content render path has only `learning.contents: []` versus a non-empty array. `refreshLearning()` spreads any response over `EMPTY_LEARNING`; a missing `contents` field silently becomes an empty array. A bridge/service rejection records a global error but does not update any Learning load/availability state. `ContentReview` then renders the same generic “Python backend chưa trả về…” branch for all of these materially different states:

- a valid successful response with zero reviewable content candidates;
- an unapplied/stale processing lifecycle response;
- a malformed successful response missing `contents`;
- a backend/bridge failure after the Learning store was still empty.

The image empty branch has the same UI-contract defect, but the content symptom is independently classified as **missing DTO/store/render state for successful-empty versus missing-field versus failed content loads**. It is not evidence that content eligibility or Ollama behavior is wrong, and no new content-learning semantics are justified.

### 3. Preview failure is already mostly fail-soft, but the contract is implicit

`_slide_preview_for()` catches production rendering failures and returns empty bitmap metadata while retaining resolved evidence regions where possible. `_image_candidate_dto()` still emits the candidate identity, bounds, decision, eligibility, and cropped-image fallback. That is the desired behavior. PROMPT-032 should pin it with a regression and make the candidate-state outcome independent of preview availability; it must not alter authoritative region construction, renderer order, cache identity, unmodified bitmap display, or report-scoped caching.

## Current refresh and rendering behavior

- Entering the top-level Learning tab already calls `refreshLearning()`.
- A bounded 750 ms session poll always reads dashboard state; while Learning is active, it also requests Learning state when the last Learning request is older than 1400 ms.
- The manual Learning button calls the same `refreshLearning()` method.
- Those hooks exist, but the backend Learning method is not lifecycle-current by itself and the frontend does not model load outcomes.
- The bridge envelope correctly turns Python exceptions into `{ok: false, error}`; the store currently collapses the visual outcome by leaving/defaulting candidate arrays.

## Proposed minimal remediation

1. At the beginning of the locked `learning_state()` read, apply the same bounded controller lifecycle reconciliation sequence already used by dashboard state: `pump()` → `reconcile()` → `pump()`. Do not reload the app, recreate the processor, poll more aggressively, sleep, or reconstruct candidates from persisted labels.
2. Return explicit, independent image/content candidate availability metadata in every successful Learning DTO. It must identify a current successful candidate snapshot (including valid zero), a still-processing state, no completed run, and learning-subsystem unavailability without inventing candidates.
3. Make the React store validate the required candidate arrays and availability fields instead of filling missing arrays from `EMPTY_LEARNING`. Track Learning load state (`idle/loading/ready/error`) and a bounded safe failure message. A manual refresh must invoke a fresh bridge call even after an earlier failure.
4. Render non-empty data normally; render a truthful valid-empty message for a successful zero-candidate snapshot; render processing/no-run/unavailable feedback when reported by Python; and render explicit backend/malformed-response failure feedback for both image and content views.
5. Preserve image candidate DTO metadata when preview generation returns no bitmap or raises inside display-only preview work. Continue to use the cropped candidate image fallback when available.
6. Add transition-bounded `LEARNING_*` diagnostics with lifecycle and image/content raw/reviewable/DTO counts plus safe opaque report identity where useful. Never log absolute paths, candidate text/evidence payloads, data URIs, or image bytes.
7. Do not change candidate eligibility, report status, Excel selection, content rebuilding, model/Ollama behavior, scan polling, scan DTOs, or renderer/evidence semantics.

## Regression plan (23 required contracts)

The implementation must pin all 23 contracts below, split across focused backend/bridge/frontend coverage:

1. Production completed/needs-review batch creates at least one eligible image candidate.
2. The same production batch creates at least one reviewable content candidate.
3. The completed processor retains both candidate object sets in `FileResult`.
4. A direct post-worker Learning read reconciles the queued terminal lifecycle without a prior dashboard read.
5. That direct read returns current image candidates.
6. That direct read returns current content candidates.
7. A subsequent manual-style backend read is fresh and reflects changed current processor state.
8. Image hard exclusions remain filtered without changing extraction/eligibility.
9. Content hard exclusions remain filtered without changing verbatim content semantics.
10. A successful current snapshot with zero image candidates is explicitly valid-empty.
11. A successful current snapshot with zero content candidates is explicitly valid-empty.
12. No completed processor is distinguishable from a current successful empty batch.
13. Learning-subsystem unavailability is distinguishable from both empty and backend failure.
14. Image preview failure preserves candidate identity, decision, eligibility, target geometry, and evidence metadata.
15. The bridge returns the explicit availability contract and candidate counts without source-path leakage.
16. `LEARNING_*` diagnostics are transition-bounded and contain no absolute paths, payload text, data URI, or bytes.
17. Entering the top-level Learning tab performs a real backend call.
18. Manual refresh performs another real backend call and can populate the same mounted view.
19. A populated image response renders image candidate feedback.
20. A populated content response renders verbatim content feedback.
21. Valid-empty image and valid-empty content responses each render truthful visible feedback.
22. A response missing required image/content fields is rendered as a contract failure, not valid empty.
23. A rejected backend/bridge Learning call renders visible failure feedback in both review paths and remains retryable.

Existing PROMPT-024R/027/027R evidence, overlay, preview-cache, renderer, report-isolation, and label/reapply tests remain required. PROMPT-031/031R scan tests must prove this change did not alter scan locks, incremental snapshots, polling, or diagnostics.

## Risks and constraints

- Calling lifecycle reconciliation from Learning must remain under the service `RLock` and use the existing controller methods; duplicating done-state logic would create divergence.
- A real running worker must still report processing rather than exposing mutable in-flight candidate lists.
- Diagnostic signatures must be bounded; Learning is periodically refreshed while active.
- DTO validation must not discard the previous successful candidates silently on a transient failure. The UI may retain prior data, but it must visibly identify the failed refresh so stale data cannot masquerade as current.
- Preview failure handling must not change PROMPT-027’s authoritative improvement regions, item ownership, renderer preference/fallback diagnostics, cache identity, report isolation, or displayed bitmap.
- Linux can prove deterministic extraction, lifecycle, DTO/bridge, store, and rendering contracts. It cannot establish Windows pywebview serialization timing, PowerPoint COM fidelity, real customer deck eligibility, or Windows acceptance.

## Windows evidence still required

After the Linux fix is green, Windows must run a real source processing batch with a report that visibly produces at least one image and one content candidate. The operator must capture bounded `LEARNING_*` lines, enter Learning without restarting, check both review views, use manual refresh, and verify PowerPoint preview fidelity/evidence overlays. Linux results must not be presented as Windows or real-PowerPoint acceptance.
