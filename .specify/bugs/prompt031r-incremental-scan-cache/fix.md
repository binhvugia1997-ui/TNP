# Bug Fix: Safe intrinsic scan cache and incremental leaf-folder delivery

- **Slug**: prompt031r-incremental-scan-cache
- **Fixed**: 2026-10-10
- **Assessment**: ./assessment.md
- **Status**: applied
- **Starting assessment commit**: `df8e42353d97c91cb745737b0bd517179bc62580`
- **Spec Kit execution**: Executed from checked-in Spec Kit skill instructions; native Arena skill invocation was not available.

## Confirmed Plan

The user accepted the assessment STOP and rejected complete-row caching by Manager+mtime. The authorized continuation is:

- identify reusable source facts by normalized source path + Management Number + file size + mtime;
- cache only context-free intrinsic facts in bounded session memory;
- always recompute period, duplicates, current recent-success state, current master state, force behavior, and final row/action;
- establish global discovery/metadata/duplicate context before publishing any leaf, then complete and publish current status work by deterministic final containing folder;
- reuse the existing bounded React dashboard poll, with generation/revision isolation and no Python-to-JavaScript worker callback or per-file bridge call;
- preserve the PROMPT-031 one-redactor context, one traversal, zero PPTX parse, and zero PowerPoint COM list-discovery contracts.

The original assessment was not rewritten. Its unsafe Manager+mtime evidence remains intact; this fix report records the subsequent authorized identity decision.

## Summary

Quét now owns a bounded, session-memory intrinsic cache with the exact authorized source identity. A hit reuses only Management-Number-derived occurrence-date facts; every dynamic status decision is rerun against the current scan. The benefit is deliberately modest because stat and Manager extraction remain necessary and PROMPT-031 had already made discovery cheap.

After global candidate/period/duplicate context is known, status work proceeds by deterministic final containing-folder groups. Each completed group publishes one generation-tagged immutable delta snapshot. Existing 750 ms polling coalesces pending leaf deltas, React merges them by stable opaque row ID, delayed revisions are rejected, and the final non-incremental response replaces the list with the canonical complete result.

## Changes

| File | Change | Notes |
|------|--------|-------|
| `app/prescan.py` | modified | Added bounded `IntrinsicScanCache`, exact `SourceIdentity`, intrinsic facts, full-identity hit/invalidation, current-context recomputation, global duplicate resolution, and containing-folder status callbacks. |
| `app/gui_controller.py` | modified | Owns one cache for the running application session, passes it into Quét, publishes one partial result per completed containing folder, and reuses the canonical normalized path. |
| `app/application_service.py` | modified | Shortens service-lock scope, publishes immutable scan generations/revisions, coalesces leaf DTO deltas for the existing poll, blocks conflicting mutations during scan, builds each incremental row once, and returns one final canonical aggregate. |
| `app/scan_diagnostics.py` | modified | Adds bounded aggregate cache/leaf counters and time-to-first-leaf measurement without sensitive absolute folder paths. |
| `frontend/src/state/store.tsx` | modified | Adds generation/revision stale-response rejection, deterministic delta merging, final replacement semantics, old-list clearing at scan start, and detail-fetch suppression while Quét is active. |
| `frontend/package.json` | modified | Includes the PROMPT-031R frontend regression in `test:frontend`; no version change. |
| `frontend/tests/prompt031r-incremental.test.mjs` | added | Verifies generation ordering and that the store renders leaf deltas while the original `scan_reports` Promise remains pending. |
| `tests/test_incremental_scan_cache.py` | added | Covers exact cache identity, every required dynamic status context, bounding, cold/warm equivalence, cross-leaf duplicates, slow-later-leaf delivery, mixed hit/miss/invalidation, and 30/120/300 scaling. |
| `tests/test_scan_performance.py` | modified | Extends PROMPT-031 gates with cold/warm cache counts, leaf delivery counts, one-row-build-per-incremental-row, stat-error isolation, and new bounded markers. |

## Exact Cache Contract

### Source identity

```text
normalized source path
+ Management Number
+ current file size
+ rounded current file mtime
```

The cache is an in-memory LRU-like ordered map bounded to 20,000 entries. It is owned by `GuiController`, so it survives repeated Quét operations in one running application and is not persisted to a database or file.

### Cached value

Only:

- derived occurrence date from the Management Number;
- bounded derivation reason when the date is invalid.

The cache does **not** contain a final action, status, duplicate winner, reason string containing another source, Excel/master row, recent-success result, force decision, manual exclusion, display order, job state, or bridge DTO.

### Work always recomputed

Every Quét still performs current source stat and filename Management Number extraction to construct/validate the authorized identity. It then recomputes:

- selected month/range/all-period decision;
- complete-tree source duplicate groups and newest/path tie-break winner;
- current persisted recent-success cache lookup, including its path+size+mtime source-change semantics;
- current master workbook rows and managed-field completeness;
- force mode;
- final action/status/reason/Excel row;
- current counters, deterministic display ordering, DTOs, and path redaction.

A same-path size change with preserved mtime replaces that path entry and records an invalidation. A changed mtime does the same. The same Manager+size+mtime under a different normalized path is a miss and a separate entry, never an alias.

## Incremental Architecture

1. The original synchronous `scan_reports` bridge call remains one batched call.
2. The service applies/validates configuration under a short lock, starts a new generation, clears the old list, and releases the lock.
3. One deterministic recursive discovery establishes all candidates. Current stat/Manager/date/period data is established globally, then duplicate winners are finalized across the complete tree.
4. Survivors' recent-success/master/force status work runs by final containing folder in deterministic discovery order.
5. At each folder boundary, the scan thread builds one leaf-only DTO delta under the service lock and appends it to a bounded pending-delivery list. It never calls JavaScript.
6. Existing `get_dashboard_state` polling returns and clears all completed leaf deltas since the previous poll. Multiple fast leaves may be safely coalesced into one poll; each row DTO was still built once.
7. React merges incremental rows by the backend opaque ID, sorts them deterministically by source index, rejects older generation/revision responses, and never removes already accepted rows because a delayed poll arrived.
8. The final response is marked non-incremental and replaces the list with the canonical complete result and existing status-group ordering.

The service reuses one scan-wide PROMPT-031 lexical redactor for all leaf deltas and the final response. `leaf_rows_delivered == candidates` is pinned, preventing cumulative per-leaf rebuilding from introducing another O(N²) row path.

## Diagnostics

Normal scans add bounded aggregate/safe-label markers:

- `SCAN_CACHE_SUMMARY hits=... misses=... invalidations=... unavailable=... stores=... entries=... evictions=... intrinsic_derivations=...`
- `SCAN_LEAF_START index=... total=... folder=<bounded basename+hash> files=...`
- `SCAN_LEAF_END index=... total=... folder=<bounded basename+hash> files=... elapsed_ms=...`
- `SCAN_LEAF_DELIVER index=... total=... rows=... elapsed_ms=... generation=... revision=...`
- `SCAN_INCREMENTAL_SUMMARY leaf_folders=... leaf_completed=... leaf_deliveries=... rows_delivered=... first_leaf_ms=... leaf_ms=...`

All PROMPT-031 `SCAN_*` markers remain. New folder labels use a bounded basename plus hash and never include an absolute directory.

## Tests Added or Updated

- `test_intrinsic_cache_identity_invalidates_path_size_and_mtime_and_never_caches_dynamic_status`
  - pins exact hit identity, different-path miss, same-mtime size invalidation, mtime invalidation, period recomputation, master missing/complete recomputation, force recomputation, and current recent-success recomputation.
- `test_intrinsic_cache_is_session_local_and_bounded`
  - pins the configured entry bound and deterministic eviction/miss.
- `test_cold_and_warm_duplicate_rows_are_equivalent_with_distinct_path_identities`
  - pins `PROCESS`/`SOURCE_DUPLICATE`, separate path identities, and complete semantic cold/warm equivalence.
- `test_incremental_leaf_snapshots_mix_hits_misses_invalidation_and_global_duplicates`
  - pins three containing folders, four hits plus one size invalidation, global later-folder duplicate winner known before early delivery, first two leaves visible while the final leaf is blocked, no path leak/duplicate ID, exactly three deliveries and five delivered row DTOs, final cold/warm equivalence, and safe close handling during an active scan.
- `test_cold_warm_scaling_uses_work_counts_not_wall_clock[30/120/300]`
  - pins one traversal, one current metadata read per candidate on both cold and warm scans, N cold misses/N warm hits, zero warm intrinsic derivations, one delivery per leaf, one incremental row build per candidate, zero PPTX parses, and zero COM starts.
- `test_large_quiet_scan_is_one_bounded_batch_with_identical_rescan`
  - now pins 200 cold misses, 200 warm hits, 200 current metadata reads, one scan-wide redactor, one leaf delivery, 200 incremental rows, one traversal, zero parser/COM work, and identical final results.
- `prompt031r-incremental.test.mjs`
  - pins stale-generation/revision rejection and renders leaf one then leaf two before the still-pending batched `scan_reports` response resolves.

## Local Verification

- Focused PROMPT-031R/performance/application-service: **34 passed in 4.99 s**.
- Scan/controller/master/new-row regressions: **120 passed in 35.45 s**.
- PROMPT-030 runtime regression: **103 passed, 1 skipped in 1.49 s** (platform-specific skip).
- Full Linux Python suite, run once after targeted checks were green: **1218 passed, 3 skipped in 110.63 s**.
- Frontend `npm run test:frontend`: **32 passed**.
- Frontend `npm run typecheck`: pass.
- Frontend `npm run build`: pass; no version change.
- `.venv/bin/python -m compileall -q app tests`: pass.
- Changed-path pyflakes: pass.
- `git diff --check`: pass at this stage.

## Benchmark Evidence

The benchmark used corrupt synthetic `.pptx` bytes so any content parse would fail, and used five runs with median timings. The PROMPT-031 comparison mode used the current canonical production service with the new intrinsic cache and leaf callback disabled; it is an unprofiled behavioral comparison, not a checkout rewrite. Timing is supporting evidence only; operation counts are the gate.

### Realistic 610-entry / 200-candidate one-folder fixture

| Mode | Entries | Candidates | Hits | Misses | Metadata reads | Leaves / deliveries | Intrinsic derivations | First leaf | Total | PPTX / COM |
|---|---:|---:|---:|---:|---:|---:|---:|---:|---:|---:|
| PROMPT-031 behavior | 610 | 200 | 0 | 0 | 200 | 1 / 0 | 200 | n/a | 39.330 ms | 0 / 0 |
| PROMPT-031R cold | 610 | 200 | 0 | 200 | 200 | 1 / 1 | 200 | 35.716 ms | 46.045 ms | 0 / 0 |
| PROMPT-031R warm | 610 | 200 | 200 | 0 | 200 | 1 / 1 | 0 | 38.413 ms | 48.790 ms | 0 / 0 |

### Multi-leaf scaling

| Candidates / leaf folders | Mode | Entries | Hits / misses | Metadata reads | Deliveries | First leaf | Total |
|---:|---|---:|---:|---:|---:|---:|---:|
| 30 / 3 | PROMPT-031 behavior | 33 | 0 / 0 | 30 | 0 | n/a | 8.167 ms |
| 30 / 3 | 031R cold | 33 | 0 / 30 | 30 | 3 | 5.917 ms | 11.021 ms |
| 30 / 3 | 031R warm | 33 | 30 / 0 | 30 | 3 | 6.684 ms | 11.917 ms |
| 120 / 6 | PROMPT-031 behavior | 126 | 0 / 0 | 120 | 0 | n/a | 26.382 ms |
| 120 / 6 | 031R cold | 126 | 0 / 120 | 120 | 6 | 20.650 ms | 47.155 ms |
| 120 / 6 | 031R warm | 126 | 120 / 0 | 120 | 6 | 22.373 ms | 48.875 ms |
| 300 / 10 | PROMPT-031 behavior | 310 | 0 / 0 | 300 | 0 | n/a | 64.172 ms |
| 300 / 10 | 031R cold | 310 | 0 / 300 | 300 | 10 | 48.195 ms | 149.796 ms |
| 300 / 10 | 031R warm | 310 | 300 / 0 | 300 | 10 | 53.923 ms | 151.580 ms |

The cache eliminates exactly one cheap date derivation per warm hit, but cannot eliminate current stat metadata or dynamic classification. In these short Linux runs that did **not** create a total-time improvement: warm was about 1–8% slower than cold, within small-operation/locking noise. The larger 031R total versus non-incremental PROMPT-031 behavior is mainly the required incremental DTO/delivery work plus the final canonical response. The implementation therefore makes no broad speedup claim. Its measured value is safe reuse and explicit hit accounting; the user-visible value is that a slow later leaf no longer withholds already completed leaves.

## Deviations from Assessment

The assessment correctly stopped the unsafe complete-row design. The user subsequently authorized the preferred normalized-path + Manager + size + mtime identity, but restricted values to intrinsic facts and required dynamic recomputation. No assessment evidence was removed or rewritten.

The implementation uses coalesced leaf **deltas**, not cumulative full snapshots. This was necessary to preserve the PROMPT-031 linear response-work fix when a tree contains many one-file leaf folders. React merges monotonic deltas and the final canonical result replaces them.

## Follow-ups

- Run the required Windows pywebview retest from source and capture the complete cold and warm `SCAN_*` blocks.
- Verify visually on a genuinely slow multi-leaf source that rows from earlier folders appear while a later folder is still running.
- Do not claim a cache speedup unless real Windows measurements demonstrate one; current Linux evidence does not.
- No database, version bump, Portable build, LAN publish, or release action is required.
