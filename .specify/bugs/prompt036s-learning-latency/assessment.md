# Bug Assessment: Learning label clicks wait on congested full-state bridge round trips

- **Slug**: prompt036s-learning-latency
- **Created**: 2026-10-13
- **Source**: pasted text (PROMPT-036S)
- **Verdict**: valid
- **Severity**: high (core review interaction; no data loss — labels do eventually persist)
- **Starting SHA**: `131dd37` (PROMPT-036R fixed commit, base branch `arena/6a73ef6f-tnp`)
- **Working branch**: `arena/2507b6dc-tnp`
- **Spec Kit execution**: Executed from checked-in Spec Kit skill instructions; native Arena skill invocation was not available.

## Report (summarized)

After PROMPT-036R removed the pywebview reflection recursion, Learning interactions
(Trước cải tiến / Sau cải tiến and candidate clicks) still respond very slowly, as if
bridge/backend commands are congested. PROMPT-036S requires instrumentation first —
not assuming dashboard polling (~750 ms), Learning refresh (~1400 ms), full
`LearningState` returns from `set_learning_label`, the `set_learning_label` +
`save_learning_labels` save sequence, or overlapping poll/user calls are the cause —
then a fix whose UX target is: a label click updates visually immediately, rapid
clicks never wait on backend round trips, last user choice wins, failures are visible
and recoverable, and persistence/Excel semantics are preserved.

## Measurement method

Deterministic bench (throwaway, `/tmp/bench036s.py`, `/tmp/bench036s_scale.py`) built on
the real production fixture `tests/test_prompt032_learning_state._finished_production_service`
(real parser/classifier/extractor/writer, Ollama disabled), plus a scaled variant that
duplicates retained candidates with unique ids (24 image candidates — conservative:
duplicates share one slide render; distinct slides cost more). Frontend timings are
code-proven (no `setState` exists before the bridge `await` in the click handlers) and
are pinned by the new frontend tests in the fix. pywebview transport time itself is
Windows-only and is honestly reported as not measurable on Linux.

## Proven root cause (measured, not assumed)

**Confidence: high.** One user-visible click pays for a full LearningState rebuild +
multi-megabyte transport, behind a single global lock, with no optimistic UI,
no coalescing, and no stale-response protection:

1. **Click → visual change waits for the entire backend round trip (frontend).**
   `labelImage`/`labelContent` (`frontend/src/state/store.tsx`) `await
   callNative('set_learning_label', …)` and only then `setLearningState`. There is no
   optimistic update, so click→visual latency == bridge queue + backend + transport +
   normalize/render of the returned state.
2. **A single label change rebuilds and returns the ENTIRE LearningState.**
   `ApplicationService.set_learning_label` ends with `return self.learning_state()`.
   Measured at 24 image candidates: `learning_state` payload **3.13 MB** (each image
   candidate carries `src` + 1600px `slidePreview` data URIs); cold construction
   **694 ms** (slide-preview rendering dominates: 101 ms for the first slide in the
   1-candidate fixture, ~25 ms/candidate blob→JPEG); `json.dumps` of the warm state
   alone is **31 ms**. Warm DTO rebuild is only ~2.6 ms — the cost is payload
   construction/serialization/transport and cold renders, repeated **per click**.
   Degraded slide previews expire after 15 s (`DEGRADED_PREVIEW_TTL_S`), so polls
   periodically re-pay cold render cost.
3. **One Save click = 2 sequential bridge calls, each returning full state.**
   Frontend `saveLearningCandidate` calls `set_learning_label` then
   `save_learning_labels`; measured 166 ms backend at bench scale, the second call
   including the Excel reapply. Both responses are full states.
4. **Everything serializes on one lock; the slowest holder stalls all interactions.**
   Every `ApplicationService` method takes the same `self._lock`, and
   `save_learning_labels` holds it across the Excel reapply. Measured: a label click
   issued while a save was in flight waited **292 ms**; cold `learning_state` hogging
   produced 44 ms lock waits even at toy scale. `dashboard_state` itself is cheap
   (0.6 ms, 7.4 KB) — dashboard polling is **not** intrinsically the bottleneck; it
   (and the ~1400 ms Learning poll, which re-downloads the multi-MB state and
   re-renders it) merely queue on the same lock and the same JS main thread.
   **The polling intervals are therefore NOT the root cause and are not changed.**
5. **No coalescing / priority / stale guard.** Each rapid click fires its own
   un-coalesced bridge call (unbounded backlog under a slow backend); a
   `get_learning_state` poll that started before a click can resolve after it and
   wholesale-overwrite the store with the older label; nothing gives user actions
   priority over background refresh.

Bridge calls per user action BEFORE fix: label click = 1 (full-state response);
Save candidate = 2 sequential (2 full-state responses); plus Learning poll every
~1400 ms while the tab is active (full state).

## Proposed minimal remediation

Backend (additive only — `set_learning_label`/`save_learning_labels` keep their
existing full-state contracts, which existing tests pin):

1. `ApplicationService.set_learning_label_fast(...)`: same validation + pending-label
   write as `set_learning_label`, returns a minimal ack
   `{kind, candidateId, label, labelPending: true}` — never calls `learning_state()`.
2. `ApplicationService.save_learning_candidate(...)`: atomic set-pending + save (the
   exact `save_learning_labels` flow for that kind, all pending labels, same Excel
   semantics) under one lock hold; returns `{message, kind, candidateId, label,
   labelPending: false, excelPending, excelLastResult}` — no full state.
3. Bridge methods `set_learning_label_fast` / `save_learning_candidate` with the same
   bounded-string validation as the existing learning methods.

Frontend (`frontend/src/state/store.tsx`):

4. Optimistic label update synchronously on click (userLabel + labelPending + labeled
   count delta), persisted asynchronously.
5. Per-candidate coalescing persist worker: at most one in-flight
   `set_learning_label_fast` per candidate; superseded values collapse to the latest;
   last user choice wins; backlog is bounded (≤ 1 in flight + latest per candidate).
6. Save candidate = exactly one `save_learning_candidate` bridge call (supersedes any
   queued fast-set for that candidate). On failure the label stays visible marked
   unsaved (`labelPending`) with the global error surfaced; clicking Save again
   retries. Never silently dropped.
7. User-action priority + stale protection: Learning polling defers while label
   writes are in flight; a refresh whose in-flight window overlapped a local label
   change is discarded and re-queued; a local-label overlay is merged over any
   incoming full state so a stale response can never revert an unconfirmed choice.
8. Poll intervals unchanged (750 ms dashboard / 1400 ms Learning): they were not the
   proven cause, and polling remains the reconciliation path.

## What is explicitly NOT changed

Learning semantics, label sets, Excel reapply behaviour, persistence files,
candidate eligibility, PROMPT-035, `TAT_QPN_PHASE6`, version/build, polling
intervals, existing bridge method signatures/returns.

## Tests to add

- `tests/test_prompt036s_learning_latency.py`: fast ack shape + pending write +
  `learning_state` not rebuilt (spy); atomic save persists label+note with Excel
  semantics in one call; bridge envelopes; last-click-wins pending value.
- `frontend/tests/learning-labels.test.mjs`: click updates UI before backend
  resolves; rapid Before→After→Before coalesces to bounded calls with last value
  persisted; slow backend does not freeze selection; stale `get_learning_state`
  cannot revert optimistic label; no Learning poll issued while a write is in
  flight; persistence failure stays visible as unsaved and Save retry recovers;
  Save candidate = 1 bridge call.
- Regression: existing Learning/backend tests, frontend suite, `tsc --noEmit`,
  `git diff --check`.

## Risks

- Overlay/epoch logic must not hide genuine backend state forever: overlay entries
  are removed on save ack and only override candidates still present; counts delta
  mirrors the backend's `saved ∪ pending` counting.
- The atomic save preserves the current save-all-pending behaviour of
  `save_learning_labels` (saving one candidate also commits other pending labels of
  that kind) — identical to today's two-call sequence, minus the duplicate round trip.
