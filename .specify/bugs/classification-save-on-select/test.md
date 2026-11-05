# Bug Verification: Classification save-on-select (classification-save-on-select)

- **Slug**: classification-save-on-select
- **Tested**: 2026-11-04
- **Assessment**: ./assessment.md
- **Fix**: ./fix.md
- **Result**: **partial** — automated verification passes on Linux (code-level trace + frontend tests with a fake bridge; no live UI run was performed and none is claimed). The real UI save flow still needs verification on a real Windows machine via the cloud workflow build.
- **Spec Kit execution mode**: Executed from the checked-in Spec Kit skill instructions; native Arena skill invocation was not available.
- **Environment**: Linux sandbox, Node + `node --test` for the frontend, `npx tsc --noEmit`, pytest for the Python bridge/service tests.

## Checks performed

| # | Check | Result |
|---|---|---|
| 1 | New regressions `frontend/tests/learning-explicit-save.test.mjs` | **3 passed** (single click stages only + Save persists; multi-select stages both + one bulk save; double-click stages only + Save persists) |
| 2 | Targeted subset (`learning-explicit-save` + `prompt030b-learning` + `learning-labels`) | **11 passed, 0 failed** |
| 3 | Typecheck | `npx tsc --noEmit` exit 0 |
| 4 | Full frontend suite (14 `test:frontend` files + the new file, 92 tests) | **90 passed, 2 failed** — both in `tests/prompt034r-report-list-layout.test.mjs`, a different tab (report list), pre-existing (see #5) |
| 5 | Attribution of the 2 failures | The bridge-traffic failure was reproduced IDENTICALLY with all Bug B edits reverted (backed-up files restored afterwards, markers verified) — pre-existing, unrelated. The viewport-height failure is flaky across runs (fails solo and in this run, passed in an earlier full run). Both left untouched. |
| 6 | Python bridge/service tests (`test_prompt030b_integration.py`, shared with Bug A) | **23 passed** |

## Contract tests updated (not weakened)

Three tests in `frontend/tests/prompt030b-learning.test.mjs` pinning the old double-click/quick-confirm contract were rewritten to the explicit-save contract with intent preserved (staging + single Save persistence, deferred bulk save, failure-retain + retry). Documented in the fix record.

## Not verified here

The real UI flow on Windows: staging several labels incl. multi-select, nothing persisting until Save, Save persisting staged labels + notes, save timing, and the poll quiet-period gate during active labeling.
