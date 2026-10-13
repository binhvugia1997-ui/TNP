# Bug Fix: "Học cải tiến" click/input lag — selections now answer locally and the label store stops being re-read per candidate

- **Slug**: prompt036-learning-ui-lag
- **Fixed**: 2026-10-13
- **Assessment**: ./assessment.md
- **Status**: applied
- **Spec Kit execution**: Executed from checked-in Spec Kit skill instructions; native Arena skill invocation was not available.
- **Extension hooks**: `.specify/extensions.yml` parsed successfully; `hooks: {}` — no `before_bug_fix` / `after_bug_fix` hooks registered, none executed.

## Summary

Two changes address the two demonstrated dominant causes. Backend: `LabelStore.records()`/`latest()` now
memoize their parsed view against the label file's own identity `(size, mtime_ns)`, removing the ~301 full
re-reads + re-parses of `image_labels.jsonl` and `content_labels.jsonl` that made up 89.3 % of one click
(RC-1). Frontend: a Before/After (or content) selection is written into local state **inside the click
event** and reconciled by the reply, guarded by a monotonic learning epoch so a snapshot that was already in
flight when the reviewer clicked can never revert it (RC-2, RC-3), and the expensive Learning poll is held
off while a selection is being acknowledged (RC-4).

## Plan confirmed before editing (per the fix skill)

- Memoize the append-only label-store views by file identity — `app/image_learning.py` only.
- Make the selection optimistic and provisional (`labelPending`), reconciled or rolled back by Python's reply.
- Add a monotonic epoch so stale full-state snapshots are discarded instead of reverting newer local edits.
- Keep unsaved local notes alive across authoritative snapshots (they reach Python only on Save).
- Hold the poll-driven Learning snapshot off the click's critical path; explicit refresh stays unconditional.
- Add regression tests on both sides that fail without the fix and pass with it.

## Changes

| File | Change | Notes |
|------|--------|-------|
| `app/image_learning.py` | modified | `LabelStore`: added `_file_identity()`, `_invalidate_cache()`, memoized `records()` and `latest()`, and an explicit invalidation after the `label()` append. Shared by the image and content stores (`content_learning.py` imports `LabelStore`), so both datasets are fixed by one change. Parsing, filtering, "last record wins", the append-only audit trail and `current_label`'s `UNLABELED` default are byte-for-byte unchanged. |
| `frontend/src/state/store.tsx` | modified | Added the PROMPT-036 learning click path: `learningStateRef`, `learningEpoch`, `learningMutationInFlight`, `lastLearningMutation`, `dirtyNotes`; helpers `publishLearning`, `mergeDirtyNotes`, `applyLearningSnapshot`, `clearDirtyNote`, `applyLocalLabel`, `restoreLocalLabel`, `nextLearningEpoch`, `beginLearningMutation`, `endLearningMutation`. Rewrote `labelImage`/`labelContent` (now one shared `labelCandidate`), `setNote`, `saveLearningLabels`, `saveLearningCandidate`, `trainModels`, `applyLabelsToExcel`, `dismissExcelResult` and `refreshLearning` to route through them. Poll gating now uses `LEARNING_POLL_MIN_INTERVAL_MS` (1400, unchanged cadence) plus the mutation hold-off. |
| `frontend/package.json` | modified | Added `test:prompt036` and appended the new test file to `test:frontend`. Version/build untouched (1.3.6). |
| `tests/test_prompt036_learning_ui_lag.py` | added test | 8 tests pinning the store memoization contract, cache invalidation, image/content isolation, bounded reads per click, click latency, rapid Before↔After switching, and PROMPT-027R §25 preview-identity semantics. |
| `frontend/tests/prompt036-learning-click.test.mjs` | added test | 7 tests pinning immediate visible selection, no lost clicks under rapid switching, stale-snapshot immunity, rollback + error surfacing on rejection, unsaved-note survival, heavy-preview responsiveness, and the PROMPT-032 refresh/content contracts. |
| `.specify/bugs/prompt036-learning-ui-lag/*` | added | Spec Kit assessment / fix / test artifacts. |

## Diff highlights

Backend — one memoized read instead of ~5 per candidate:

```python
def records(self) -> List[dict]:
    identity = self._file_identity()
    if self._records_cache is not None and self._cache_identity == identity:
        return list(self._records_cache)
    ...unchanged parse/filter...
    self._records_cache = out
    self._latest_cache = None
    self._cache_identity = identity
    return list(out)
```

Frontend — the selection lands inside the click event (everything before the first `await` runs in the
handler), and Python reconciles afterwards:

```ts
const labelCandidate = useCallback(async (kind, id, label) => {
  const note = (...current candidate...)?.note ?? ''
  const epoch = beginLearningMutation()          // in-flight snapshots are now stale
  const previous = applyLocalLabel(kind, id, label, note)   // visible THIS frame, labelPending: true
  try {
    const next = await callNative<LearningState>('set_learning_label', kind, id, label, note)
    clearDirtyNote(kind, id)                     // Python now owns this note
    applyLearningSnapshot(normalizeLearning(next), epoch)
  } catch (error) {
    if (previous) restoreLocalLabel(kind, id, previous, epoch)   // never leave a rejected label on screen
    recordError(error)
  } finally { endLearningMutation() }
}, [...])
```

The guard that kills the lost-click race (M5):

```ts
const applyLearningSnapshot = useCallback((next, requestedEpoch) => {
  if (!alive.current) return false
  if (requestedEpoch !== learningEpoch.current) return false   // a newer local edit supersedes this snapshot
  publishLearning(mergeDirtyNotes(next))
  setTraining(Boolean(next.training))
  return true
}, [mergeDirtyNotes, publishLearning])
```

## Measured effect (same harnesses as the assessment)

Backend, 60 image + 60 content candidates, 400 saved labels per store
(`image_labels.jsonl` 385 KiB, `content_labels.jsonl` 419 KiB):

| Metric | Before | After |
|---|---|---|
| `set_learning_label` — ONE click | 1203.7 ms | **13.4 ms** (90×) |
| label-file reads per click (each store) | 301 | **0** (memo hit) |
| label-store I/O share of a click | 89.3 % | **0.0 %** |
| `learning_state()` warm, labels populated | 2058.0 ms | **770.7 ms** |

The remaining 770.7 ms is the one-time PROMPT-027R §25 slide re-render after labels are *saved*
(`parse_pptx` 8, `render` 8) — mandated semantics, deliberately kept.

Frontend, real `StoreProvider` + `LearningTab` in jsdom against an instrumented bridge:

| Probe | Before | After |
|---|---|---|
| A. single click → selection visible (400 ms bridge) | 418.6 ms | **22.3 ms** |
| B. 5 rapid Trước↔Sau clicks not shown immediately | **5 / 5** | **0 / 5** (4.9–7.0 ms each) |
| B. final UI vs backend truth vs expected | agreed | agreed (`AFTER`) |
| D. click while a poll is in flight | shown = **false** (507 ms ack) | shown = **true** in **5.2 ms** |
| E. stale in-flight poll reverts an accepted click | **true** (UI ended unlabeled) | **false** (UI ended `BEFORE`) |
| Learning polls during a 3 s settle after rapid clicks | 2 | 1 |

## Tests added or updated

- `tests/test_prompt036_learning_ui_lag.py::test_label_store_serves_repeated_reads_from_one_disk_read` — 200 logical reads = 1 disk read; pins RC-1.
- `::test_label_store_cache_is_invalidated_by_its_own_append` — a changed label is visible immediately; both records survive in the audit trail.
- `::test_label_store_cache_is_invalidated_by_an_external_file_change` — replaced/deleted file never served from memo (update rollback, Constitution IV/X).
- `::test_image_and_content_stores_never_share_a_memoized_view` — the two datasets stay isolated.
- `::test_one_label_click_reads_the_label_files_a_bounded_number_of_times` — ≤ 4 reads per click (was 602); the returned DTO contract is unchanged.
- `::test_one_label_click_is_fast_with_a_large_saved_label_history` — < 250 ms (measured 398 ms unfixed, 1203.7 ms in the assessment harness).
- `::test_rapid_before_after_switching_keeps_every_label_and_the_audit_trail` — 5 rapid switches, no lost or duplicated label.
- `::test_saving_labels_still_invalidates_the_preview_identity` — PROMPT-027R §25: pending label does not change the identity, saved label does.
- `frontend/tests/prompt036-learning-click.test.mjs` — 7 tests: immediate visible selection; rapid switching loses nothing and converges; stale snapshot cannot revert; rejection rolls back and surfaces `CANDIDATE_NOT_FOUND`; unsaved note survives a snapshot and is what reaches Python; heavy preview stays responsive and switching candidate makes **zero** bridge calls; explicit refresh still performs a real read and the content path keeps its own label semantics.

## Local verification

- `.venv/bin/python -m pytest tests/test_prompt036_learning_ui_lag.py -q` → **8 passed**.
- Same file with `app/image_learning.py` stashed → **3 failed, 5 passed** (the 5 are contract-preservation tests that must pass either way). Proves the tests pin the cause.
- `cd frontend && node --test tests/prompt036-learning-click.test.mjs` → **7 passed**.
- Same file with `frontend/src/state/store.tsx` stashed → **0 passed, 7 failed**, each with the intended assertion message.
- `cd frontend && npm run test:frontend` → **43 tests, 42 pass, 1 fail**; the single failure is the pre-existing `the frontend package metadata agrees with the authoritative backend version` (baseline was 36/35/1 — same failure).
- `cd frontend && npm run typecheck` → **clean**.
- `.venv/bin/python -m pytest tests/ -q -p no:randomly` → **36 failed, 1189 passed, 18 skipped**; baseline was **36 failed, 1181 passed, 18 skipped** and `diff` of the sorted FAILED lists is **empty** (identical failure sets; +8 = the new tests).
- `.venv/bin/python -m pyflakes app/image_learning.py app/application_service.py tests/test_prompt036_learning_ui_lag.py` → clean.
- `git diff --check` → clean.

## Deviations from Assessment

1. **Proposed fix 2 (compute `SlideRenderer.availability()` once per `learning_state()` build) was NOT implemented.** Two reasons, both discovered while implementing:
   - Measurement: the assessment's own decomposition put `availability()` at **4.3 ms of 1319.7 ms (0.3 %)** — it is not a demonstrated dominant cause, so the constitution's "fix the demonstrated cause without unrelated changes" argues against touching it.
   - Contract: `tests/test_prompt027r_renderer.py::test_preview_identity_includes_backend_availability` calls `_slide_preview_for` twice in a row (no `learning_state()` between) and requires a newly available PowerPoint to be seen immediately, and `::test_availability_is_never_cached_between_probes` pins that the probe is never cached. A per-build memo would break both. Weakening authoritative tests is forbidden (Constitution IV/VIII, AGENTS.md §4).
2. **Scope addition (same root cause, small):** `setNote` now participates in the epoch guard and unsaved notes are re-applied over authoritative snapshots. The assessment listed RC-3 only for labels; while implementing, the identical mechanism was shown to erase in-flight typing on the next poll (the request explicitly covers "click/input lag" and "Không làm mất learning feedback"). It is pinned by `an unsaved note survives an authoritative snapshot`, which fails without the fix.
3. **`clearDirtyNote` is per-candidate, not per-kind.** `save_learning_labels` only persists notes of candidates that already hold a pending label, so clearing every note of a kind would have discarded notes Python never received. Clearing is now restricted to the candidate whose note Python actually took ownership of.
4. **Store-object memoization was NOT done.** The assessment listed unnecessary rerenders as a suspect, but measurement did not demonstrate it (M6/C: `store renders +1` over 4 s idle). Wrapping the ~80-field store object in `useMemo` would be a large, risk-carrying refactor without evidence, so it was left alone.

## Follow-ups

- The full Learning payload still carries every candidate's `src` and `slidePreview` data URI (4.6 MiB for 60 synthetic candidates; ~18 MiB with realistic bitmaps). Changing that is a DTO contract change explicitly rejected in the assessment; if Learning datasets keep growing, a lazy per-candidate preview endpoint would be the next lever.
- `save_learning_labels` still triggers the PROMPT-027R §25 full slide re-render (~770 ms measured). Correct by spec, but a targeted per-slide invalidation (only slides whose region decision actually changed) would cut it further without weakening §25.
- Windows acceptance remains outstanding: real pywebview WebView2 transport timing, PowerPoint COM preview rendering, and Portable behaviour were not exercised (Constitution VII).
