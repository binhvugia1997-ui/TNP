# Bug Verification: Learning label clicks wait on congested full-state bridge round trips

- **Slug**: prompt036s-learning-latency
- **Tested**: 2026-10-13
- **Assessment**: ./assessment.md
- **Fix**: ./fix.md
- **Result**: partial (all automatable checks pass; Windows pywebview timing acceptance not performable on Linux — constitution VII)

## Summary

The original symptom path was reproduced by measurement before the fix (full-state
rebuild + 3.13 MB payload per click at 24 candidates, 694 ms cold; 2 sequential
full-state calls per Save; 292 ms lock wait for a click during a save) and the fixed
paths are pinned by new backend/frontend tests. No regressions: Python failure set is
byte-identical to pristine HEAD; frontend has only the same pre-existing failure.

## Checks Performed

| Check | Command / Action | Result | Notes |
|-------|------------------|--------|-------|
| Reproduction (pre-fix) | `/tmp/bench036s.py`, `/tmp/bench036s_scale.py` (production fixture) | pass (bug proven) | learning_state cold 694 ms / 3.13 MB @24 images; set per click returns full state; save = set+save 166 ms; click during save waited 292 ms on the single service lock |
| Post-fix backend numbers | `/tmp/bench036s_after.py` | pass | `set_learning_label_fast` 0.01 ms, 97-byte ack; `save_learning_candidate` 1 call, 481-byte ack, label+note persisted |
| One click changes UI immediately | `frontend/tests/learning-labels.test.mjs` phase 1 | pass | UI shows BEFORE\|pending while the bridge promise is still unresolved |
| Rapid Before→After→Before (last wins, bounded) | phase 2 (BEFORE→AFTER→CONTROL) | pass | 3 rapid clicks → 2 bridge calls; final persisted = last click |
| Slow backend does not freeze selection | phases 1–2 (hung deferred bridge) | pass | Selection renders/updates during an indefinitely pending call |
| Stale get_learning_state cannot revert | phases 3–4 | pass | Poll deferred during writes; stale response overlaid; overlapping response epoch-discarded + re-fetched |
| Background polling cannot starve user actions | phase 3 + code path | pass | `refreshLearning` issues no bridge call while a label write is in flight; user writes use O(1) fast ack so lock hold collapses |
| Persistence failure visible + recoverable | phase 5 | pass | Rejected fast-set keeps CONTROL\|pending (chưa lưu) + error message; Save retry succeeds, label saved |
| Save candidate = 1 bridge call | phase 5 | pass | `save_learning_candidate` ×1; legacy `set_learning_label`/`save_learning_labels` ×0 |
| Content label path | phase 6 | pass | Same optimistic/coalesced behaviour |
| New Python tests | `pytest tests/test_prompt036s_learning_latency.py` | pass | 5 passed |
| Existing Learning tests | `pytest tests/test_prompt032_learning_state.py tests/test_prompt036r_bridge_hang.py tests/test_prompt024r_learning_workspace.py tests/test_application_service.py` | partial | All pass except pre-existing `test_configuration_ollama_update_state_and_disabled_install` (expects `1.3.5`, backend is `1.3.6` at HEAD 131dd37; verified failing on pristine HEAD worktree) |
| Full Python regression | `pytest tests` | partial | 1197 passed, 36 failed, 15 skipped — failure list identical (`diff` empty) to pristine HEAD run: zero new regressions |
| Frontend suite | `npm run test:frontend` (39 tests incl. new) | partial | 38 pass; 1 pre-existing failure `the frontend package metadata agrees with the authoritative backend version` (test constant `1.3.5` vs `1.3.6` already at HEAD; not touched — version bumps forbidden by PROMPT-036S) |
| Frontend typecheck | `npm run typecheck` (`tsc --noEmit`) | pass | |
| Lint / diff hygiene | `pyflakes` on changed Python; `git diff --check` | pass | |

## Output Excerpts

```text
BEFORE: learning_state COLD(24 img) ms=694  payload=3.13 MB
BEFORE: save-candidate today: set 2.5ms + save 163.4ms = 165.9ms, 2 sequential bridge calls
BEFORE: label click while save in flight: 291.9ms
AFTER:  set_learning_label_fast ms=[0.02, 0.01, ...] ack_bytes=97
AFTER:  save_learning_candidate atomic: 50.3ms, 1 call, ack_bytes=481, labelPending=False
pytest tests/test_prompt036s_learning_latency.py → 5 passed
node --test tests/learning-labels.test.mjs → pass 1 / fail 0
```

Bridge calls per click — BEFORE: label = 1 (full-state response), Save = 2 sequential
(full-state responses). AFTER: label = 1 minimal ack (coalesced under rapid clicking),
Save = 1 atomic minimal ack.
Click visual latency — BEFORE: full bridge round trip (backend rebuild + multi-MB
transport + normalize/render; seconds when cold/congested). AFTER: same React tick
(local `setState` before any bridge traffic; pinned with a hung bridge in tests).
Backend/persistence latency — fast pending write ~0.01 ms warm (vs 2.6 ms warm /
694 ms cold full-state rebuild at bench scale, plus 31 ms JSON serialization of the
3.13 MB payload that no longer happens per click); atomic save 50 ms at fixture scale
(Excel reapply, semantics unchanged, now paid once instead of after a duplicate set).

## Residual Risks

- Real Windows pywebview transport/render timing (PowerPoint COM slide renders, WebView2
  main-thread cost) cannot be measured in this Linux sandbox; the fix removes the
  per-click payload/render work those costs multiplied, but Windows acceptance of the
  felt latency remains outstanding.
- The two `1.3.5`/`1.3.6` stale-expectation failures pre-date this change and are out of
  scope (PROMPT-036S forbids version bumps); they need a separate authorized prompt.

## Recommendation

Accept as PARTIAL: fix verified end-to-end on Linux with deterministic tests and
zero regressions; hold final PASS for the Windows click-latency acceptance check.
