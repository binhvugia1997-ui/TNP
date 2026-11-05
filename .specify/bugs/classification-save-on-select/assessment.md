# Bug Assessment: Classification save logic — selecting an item saves immediately and each save loads slowly

- **Slug**: classification-save-on-select
- **Created**: 2026-11-04
- **Source**: pasted text (user report in chat, 2026-11-04)
- **Verdict**: valid
- **Severity**: high (the normal review flow — classify Before/After first, then press Save — is broken for multi-select/double-click, and every label interaction triggers a slow full reload)

## Report (verbatim or summarized)

User (translated): "The normal flow is: classify Before/After improvements FIRST, then press Save. Now, merely SELECTING an item saves immediately, and each save loads very slowly."

## Symptom

Two linked symptoms in the Learning tab (image + content classification):
1. Selecting/classifying an item persists immediately in some paths instead of staging in memory until the explicit Save button is pressed.
2. Each interaction waits a long time (full reload / loading spinner).

## Reproduction

Code-level reproduction (no live UI run was performed on Linux; the wiring below is the reproduction evidence):

1. Open the Learning tab, select 2+ image candidates, click a label button ("Trước cải tiến" / "Sau cải tiến" / "Không lấy").
   → `frontend/src/tabs/LearningTab.tsx:726-735`: multi-select label button calls `confirm()` → `saveImageCandidates` (real save: JSONL labels + Excel re-apply) with NO Save press.
2. Double-click any image card → `LearningTab.tsx:753-754` `handleDoubleClick` → `confirm()` → real save. The UI hint text itself advertises this: "Bấm đúp để chọn + xác nhận nhanh (double-click)".
3. Click a label button for ONE candidate → stages only (correct), BUT `frontend/src/state/store.tsx:731-770` `persistLabel` → `drainLearningRefresh()` → full `get_learning_state` reload (`refreshLearning`, store.tsx:1058-1135) on EVERY click: multi-MB payload (PROMPT-036S measured 3.13 MB at 24 candidates, 694 ms cold construction) plus slide-preview renders with a 15 s degraded-preview TTL (`app/application_service.py:1599`), and a loading state flash (`setLearningLoad`).
4. Every real save additionally runs `gui_controller.py:696 save_confirmations` → labels JSONL + `_image_reapply` → `app/image_review.py:90 reapply_labels`: re-parses the PPTX, re-runs `select_with_learning`, re-renders After pictures via `export_after_pictures` (renderer chain → PowerPoint COM) and rewrites Excel image cells — on EVERY save — followed by another full `refreshLearning`.

[NEEDS CLARIFICATION: the live UI save flow was not run on Linux. Whether the user's installed build (v1.3.7/Build 020 per his log) predates the PROMPT-036S fast-path staging is unconfirmed; the wiring above is the current snapshot's behavior. Windows/UI acceptance remains pending.]

## Suspected Code Paths

- `frontend/src/tabs/LearningTab.tsx:726-735` — label buttons: `const multiple = s.imageSelected.length > 1; ... multiple ? confirm() : s.labelImage(...)`.
- `frontend/src/tabs/LearningTab.tsx:753-754` — `handleDoubleClick` → `confirm()` → `saveImageCandidate`.
- `frontend/src/tabs/LearningTab.tsx:761-771` — explicit Save button ("Lưu nhãn & ghi chú") → `confirm()` → `saveImageCandidate` (the only path that should persist for images).
- `frontend/src/tabs/LearningTab.tsx:988,1012` — content tab: label click stages (`s.labelContent`, pending-only); Save button persists (`saveContentCandidate`). The content side ALREADY matches the desired flow — the image side must match it.
- `frontend/src/state/store.tsx:1209-1235` `applyLocalLabel` (optimistic local update) + `731-770` `persistLabel` coalescing worker → bridge `set_learning_label_fast` (`app/application_service.py:1233`, pending-only ack; `gui_controller.py:688 set_pending_label` writes in-memory dicts only — confirmed memory-only).
- `frontend/src/state/store.tsx:1265-1313` `saveLearningCandidate` (atomic set+save) + `finally { void refreshLearning() }`; `1320-1378` `saveImageCandidates` (bulk save + refresh).
- `frontend/src/state/store.tsx:1058-1135` `refreshLearning` → `get_learning_state` full-state reload; `741-749` `drainLearningRefresh` runs it after every persist-worker drain.
- `app/application_service.py:1095` `learning_state()` — `pump()/reconcile()` + full DTO rebuild; `:1787` `_image_candidate_dto` → `:1582` `_slide_preview_for` (cached per report+slide+identity; degraded previews expire after `DEGRADED_PREVIEW_TTL_S = 15.0`).
- Prior art: `.specify/bugs/prompt036s-learning-latency/assessment.md` (2026-10-13) already diagnosed click latency and landed the fast-path/optimistic/coalescing fix in this snapshot. The remaining slowness is the per-click full refresh and the per-save Excel re-apply (by design for saves — see Risks).

## Root Cause Hypothesis

Confidence: high for the wiring; medium for the exact per-click cost split (refresh vs. preview re-render) without a live profile.
1. Save-on-select: the multi-select and double-click handlers were deliberately wired to "quick confirm" (save immediately) — a design that contradicts the explicit-save flow the user describes. Single-candidate label clicks stage correctly, so the bug is path-specific, not a total absence of staging.
2. Slowness: two distinct costs — (a) every real save runs the full Excel re-apply (PPTX re-parse + evidence re-render + workbook rewrite) plus a full `refreshLearning` reload; (b) the bridge-session background poll (`frontend/src/state/store.tsx:796`, throttled to 1400 ms) re-fetches the full `get_learning_state` on the Learning tab, and a poll refresh requested during label writes defers and is then drained by `drainLearningRefresh` (`store.tsx:725-729`) immediately after the write worker drains — so full reloads land on top of labeling bursts. CORRECTION (fix phase, 2026-11-04): the drain itself only fires when a refresh was actually requested during writes (flag `learningRefreshQueuedRef`), and the deferred-refresh contract is pinned by `frontend/tests/learning-labels.test.mjs` step 3 — so the drain is KEPT. The targeted change is a quiet-period gate on the background poll only: no poll refresh while the user labeled within the last 1400 ms. Explicit refreshes (Save finally, refresh button, tab switch) and the drain are untouched.

## Proposed Remediation

**Preferred**:
1. Label-button clicks stage only (always `applyLocalLabel` + pending write) for single AND multi-select; never call `confirm()`/save. Double-click stages the same way (no save); update the hint text. Only the explicit Save buttons persist (image Save → `saveImageCandidate`/`saveImageCandidates`; content Save already correct).
2. Gate the background bridge poll: skip the poll-triggered `refreshLearning` while the user labeled within the last 1400 ms (same throttle as the refresh itself). The deferred-refresh drain (`drainLearningRefresh`) and all explicit refreshes stay exactly as they are — the drain only fires for refreshes requested during writes and is pinned by the 036S tests.
3. Do NOT remove the Excel re-apply from the save path — it is required save semantics; report its cost precisely instead.

**Alternatives**:
- Keep double-click quick-confirm as an explicit gesture. Rejected: the user's flow is "classify first, then press Save"; double-click save violates it literally, and the task direction is explicit-save only.
- Debounce the per-click refresh instead of removing it. Weaker: still pays a full reload per interaction burst for no new information.

**Files likely to change**:
- `frontend/src/tabs/LearningTab.tsx` (label buttons, double-click handler, hint text)
- `frontend/src/state/store.tsx` (persistLabel refresh drain)
- `frontend/tests/learning-labels.test.mjs` — read-only reference for the 036S contract; new regression tests in a NEW file so pinned tests are untouched (or extended only additively)
- Python: no change expected unless a bridge contract changes (none planned)

**Tests to add or update**:
- Selecting/classifying (single, multi-select, double-click) without Save persists nothing: no `save_learning_candidate` / `save_learning_images` / `save_learning_labels` bridge calls; only the pending `set_learning_label_fast` writes.
- Save persists the staged classification (staged label + note reach the save call).
- A staged label click does not trigger a full `get_learning_state` reload (no refresh call in the bridge log after a click, before Save).
- Existing 036S contract tests keep passing unchanged.

## Risks & Considerations

- Removing the post-click refresh means the UI trusts optimistic state between clicks and the next save/refresh. If a `set_learning_label_fast` ack fails, the existing error path (`recordError`, `labelPending` + error surfacing) must still fire — verify in tests.
- The save path remains slow by design (Excel re-apply with evidence re-render). If Windows timing shows saves are still unacceptable, that is a separate change (scoping the re-apply), not to be smuggled into this fix.
- The multi-select label flow previously gave immediate feedback that labels were saved; after the fix the pending indicator (`labelPending`) is the feedback until Save. Confirm the pending indicator is visible for multi-select in the UI (it is: `LearningTab.tsx` renders pending state per candidate).

## Open Questions

- [NEEDS CLARIFICATION: none blocking. Whether double-click quick-confirm should survive as an opt-in is settled by the task direction (explicit-save only); if the user objects, restore it as a separate decision.]
