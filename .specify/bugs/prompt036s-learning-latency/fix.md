# Bug Fix: Learning label clicks wait on congested full-state bridge round trips

- **Slug**: prompt036s-learning-latency
- **Fixed**: 2026-10-13
- **Assessment**: ./assessment.md
- **Status**: applied

## Summary

Label clicks are now optimistic (local visual update in the same React tick), persist
asynchronously through a coalescing per-candidate worker over a new minimal-ack bridge
method, and Save is one atomic backend call instead of the sequential
`set_learning_label` + `save_learning_labels` pair. Polling defers to user writes and
stale responses can no longer revert a choice. Existing full-state methods, Learning
semantics, Excel behaviour, and poll intervals are unchanged.

## Changes

| File | Change | Notes |
|------|--------|-------|
| `app/application_service.py` | modified | Extracted `_set_pending_label_unlocked` / `_save_learning_labels_unlocked` (exact former logic); added `set_learning_label_fast` (minimal ack, never rebuilds state) and `save_learning_candidate` (atomic set+save, minimal ack). Old methods keep their full-state returns. |
| `app/bridge.py` | modified | Added `set_learning_label_fast` / `save_learning_candidate` bridge methods with the same bounded-string validation. |
| `frontend/src/state/store.tsx` | modified | Optimistic label update + labeled-count delta; per-candidate coalescing persist worker (≤1 in flight + latest, last choice wins); local-label overlay merged over incoming full states (entries drop on backend confirmation or save ack); refresh defers while writes are in flight and epoch-discards responses that overlapped a click, then re-queues; Save candidate = 1 `save_learning_candidate` call, failure keeps the label visible marked unsaved (`labelPending`) with global error, retry via Save; save/train/retry/dismiss full states routed through the overlay merge. |
| `frontend/package.json` | modified | New `test:learning-labels` script; included in `test:frontend`. No version change. |
| `tests/test_prompt036s_learning_latency.py` | added test | Fast ack shape + no state rebuild (spy), last-click-wins, unknown candidate error, atomic save persistence + locked-Excel retry semantics, bridge envelopes. |
| `frontend/tests/learning-labels.test.mjs` | added test | Immediate visual update with hung backend, rapid-click coalescing (3 clicks → 2 calls, last wins), poll deferral during writes, stale-response overlay + epoch discard, failure stays visible unsaved, Save recovery with exactly 1 atomic call and 0 legacy calls, content path. |

## Diff Highlights (optional)

- `set_learning_label` previously ended in `return self.learning_state()`; the fast path
  returns `{"kind", "candidateId", "label", "labelPending": true}` (97 bytes measured)
  after the same O(1) pending write.
- Frontend `labelImage`/`labelContent` no longer `await` anything before `setLearningState`.

## Tests Added or Updated

- `tests/test_prompt036s_learning_latency.py` (5 tests) — pins the fast/atomic backend contract.
- `frontend/tests/learning-labels.test.mjs` (1 scenario test, 6 phases) — pins the UX contract.

## Local Verification

- New Python tests: 5 passed. New frontend test: passed.
- Existing Learning suites (prompt032, prompt036r, prompt024r, application_service): pass
  except one pre-existing HEAD failure (stale `1.3.5` expectation vs backend `1.3.6`).
- Full Python suite: 1197 passed; failure set identical to pristine HEAD (36 pre-existing).
- Frontend suite: 38/39 (same pre-existing version failure); `tsc --noEmit` PASS;
  `pyflakes` PASS; `git diff --check` PASS. Details in ./test.md.

## Deviations from Assessment

None. Poll intervals were investigated and deliberately left unchanged (not the proven
cause); reconciliation still happens via the existing poll plus post-write refreshes.

## Follow-ups

- Windows/pywebview acceptance of real click latency remains for the Windows check
  (Linux numbers are development evidence only, per constitution VII).
