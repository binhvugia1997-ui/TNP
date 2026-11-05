# Fix: Classification save-on-select (classification-save-on-select)

- **Slug**: classification-save-on-select
- **Assessment**: `.specify/bugs/classification-save-on-select/assessment.md`
- **Status**: fixed (Linux dev evidence only; the real UI save flow still needs Windows verification)
- **Mode B disclosure**: Executed from checked-in Spec Kit skill instructions; native Arena skill invocation was not available.

## Root cause (confirmed by code trace; no live UI run on Linux)

Save-on-select wiring in `frontend/src/tabs/LearningTab.tsx`: the label buttons called `confirm()` — a REAL save via `saveImageCandidates` — when `selectedOnSlide.length > 1`, and `handleDoubleClick` called `confirm()` too. Single-candidate clicks already staged correctly. The Content tab already matched the desired flow (`labelContent` stages, Save persists). `confirm()` → `saveImageCandidate`/`saveImageCandidates` → bridge `save_learning_candidate`/`save_learning_images` → JSONL + `gui_controller.py:696 save_confirmations` → `app/image_review.py:90 reapply_labels` (re-parse the PPTX, re-run selection, re-render After pictures, rewrite Excel cells) + a full `refreshLearning`. Pending writes are memory-only (confirmed: `gui_controller.py:688 set_pending_label` writes only in-memory dicts; `app/application_service.py:1233 set_learning_label_fast` is a pending-only ack; the store's `applyLocalLabel` is optimistic with a coalescing `persistLabel` worker).

Why each save felt slow: (a) every real save pays the Excel re-apply + full reload (by design, kept); (b) the bridge-session background poll (`frontend/src/state/store.tsx`, 1400 ms throttle) re-fetched the full `get_learning_state` on the Learning tab, and poll refreshes requested during label writes were deferred (`learningRefreshQueuedRef`) and drained by `drainLearningRefresh` immediately after the write worker drained — full reloads landing on labeling bursts.

Assessment correction (recorded in the assessment file): the original assessment proposed removing the `drainLearningRefresh` drain; deeper tracing showed the drain only fires when a refresh was actually requested during writes and is pinned by `frontend/tests/learning-labels.test.mjs` step 3, so the drain was KEPT. The targeted slowness change is a quiet-period gate on the background poll only.

## Changes

- `frontend/src/tabs/LearningTab.tsx`:
  - Label-button `onClick` now always stages (`for (const c of targets) s.labelImage(c.id, l.key)`, targets = multi-selection or the current candidate) and never calls `confirm()`.
  - `handleDoubleClick` now stages the active/existing label on the clicked/selected candidates and never saves, with the hint "Vui lòng chọn nhãn trước khi gắn nhãn nhanh (double-click)." when no label is chosen. (Double-click quick-confirm is REMOVED by this change, per the explicit-save direction; restoring it is a separate decision if the user objects.)
  - `confirm()` now serves ONLY the Save button ("Lưu nhãn & ghi chú", the sole remaining caller); its empty-label hint is "Vui lòng chọn nhãn trước khi lưu."
  - The Content tab is untouched (already correct).
- `frontend/src/state/store.tsx`:
  - New `lastLabelActivityRef`, updated in `applyLocalLabel`. The background poll callback now skips the Learning refresh while the user labeled within the last 1400 ms (`Date.now() - lastLabelActivityRef.current > 1400`, added to the existing 1400 ms throttle). Explicit refreshes (Save `finally`, refresh button, tab switch) and the deferred-refresh drain are untouched. The post-save full refresh is kept (the save ack does not return full state; re-apply can change decisions/evidence that only a reload surfaces).
- `frontend/tests/learning-explicit-save.test.mjs` (new, 3 tests, real LearningTab + fake bridge, modeled on the 030B harness): single click stages only + no full reload + Save persists; multi-select stages both + Save makes one bulk save; double-click stages only + Save persists.
- `frontend/tests/prompt030b-learning.test.mjs` (3 tests updated, intent preserved): the native double-click test renamed/rewritten to "double-click stages without saving; Save confirms once"; the bulk-save test now stages via label click (2 fast calls, 0 saves) then saves via the Save button (deferred bulk save, poll-defer, failure-retain, retry via Save); the "double-click preserves … one bulk save" test now stages both via double-click and saves once via Save. Other 030B tests (including "Save keeps the current candidate label", the no-label double-click hint) are unchanged.

## Verification (Linux)

- Targeted subset (`learning-explicit-save` + `prompt030b-learning` + `learning-labels`): 11 passed, 0 failed.
- `npx tsc --noEmit`: exit 0.
- Full frontend suite (14 `test:frontend` files + the new file, 92 tests): 91 passed, 1 failed — `PROMPT-034R · a 1000-report list adds no per-row native bridge traffic`. Attribution was settled definitively: with ALL Bug B edits reverted (backed-up files restored afterwards, markers verified), the test fails identically (1000 rows: `get_dashboard_state: 5, get_report_details: 0` vs 6 rows: `4 / 8`). It is a pre-existing failure on a different tab (report list), unrelated to this change; a companion height assertion in the same file is flaky across runs (fails solo, passes in the full run). Both are left untouched.
- The Python tests covering the touched bridge/service (`test_prompt030b_integration.py`, shared with Bug A) pass: 23/23.

## Still needs Windows verification (via the cloud workflow build, real UI)

- Stage several labels including a multi-select, confirm nothing persists until Save; Save persists the staged labels + notes.
- Save timing with PowerPoint open/closed.
- The background poll no longer reloads the Learning state during active labeling.
