# Bug Verification: Safe cache semantics and incremental folder-level Quét results

- **Slug**: prompt031r-incremental-scan-cache
- **Tested**: 2026-10-10
- **Assessment**: ./assessment.md
- **Fix**: ./fix.md
- **Result**: partial
- **Environment**: Linux agent, Python 3.11.2, Node/npm frontend toolchain; no real Windows pywebview/UNC/removable-drive acceptance
- **Spec Kit execution**: Executed from checked-in Spec Kit skill instructions; native Arena skill invocation was not available.

## Summary

The Linux implementation and regressions verify the authorized source identity, safe intrinsic-only values, mandatory dynamic recomputation, path/size/mtime invalidation, deterministic global duplicates, one publication per completed containing folder, slow-later-leaf visibility, monotonic frontend merging, and final cold/warm equivalence. PROMPT-031 remains intact: one traversal, one scan-wide redactor, zero PPTX parsing, and zero PowerPoint COM during Quét.

The result remains **partial** because the acceptance target includes real Windows pywebview concurrency and visible incremental rendering on the user's actual multi-folder source. Linux cannot substitute for that retest.

## Checks Performed

| Check | Command / Action | Result | Notes |
|-------|------------------|--------|-------|
| Focused cache/incremental/performance/service | `.venv/bin/pytest -q tests/test_incremental_scan_cache.py tests/test_scan_performance.py tests/test_application_service.py` | pass | 34 passed in 4.99 s. |
| Scan/controller/master/new-row regressions | `.venv/bin/pytest -q tests/test_prescan.py tests/test_scanner_visibility.py tests/test_scan_list.py tests/test_scan_list_compact.py tests/test_gui_controller.py tests/test_master_not_found.py tests/test_new_master_row.py tests/test_new_row.py` | pass | 120 passed in 35.45 s. |
| PROMPT-030 runtime regression | `.venv/bin/pytest -q tests/test_prompt030_runtime.py` | pass | 103 passed, 1 platform-specific skip in 1.49 s. |
| Full Linux Python suite (once) | `.venv/bin/pytest -q` | pass | 1218 passed, 3 skipped in 110.63 s. |
| Frontend regressions | `cd frontend && npm run test:frontend` | pass | 32 passed, including 3 PROMPT-031R generation/render tests. |
| Frontend typecheck | `cd frontend && npm run typecheck` | pass | `tsc --noEmit` clean. |
| Frontend production build | `cd frontend && npm run build` | pass | 1588 modules transformed; no version bump. |
| Compile | `.venv/bin/python -m compileall -q app tests` | pass | No compilation errors. |
| Changed-path lint | `.venv/bin/pyflakes app/prescan.py app/gui_controller.py app/application_service.py app/scan_diagnostics.py tests/test_incremental_scan_cache.py tests/test_scan_performance.py` | pass | No warnings. |
| Whitespace | `git diff --check` | pass | Clean at report-writing stage. |
| Windows pywebview incremental UI | manual retest below | not-run | Required for final verified verdict. |
| Portable/LAN publishing | prohibited / not applicable | skipped | No version bump, build, or publish action was run. |

## Required Cache Regression Matrix

| Requirement | Evidence | Result |
|---|---|---|
| Same path + Manager + size + mtime hits | Warm direct and 30/120/300 scans report N hits, zero misses/derivations | pass |
| Same Manager+mtime, different path does not collide | Distinct-path direct test reports miss and separate entry | pass |
| Changed size with same path/Manager/mtime invalidates | Size is changed while mtime is restored; invalidations=1, hit=0 | pass |
| Changed mtime invalidates | Direct test changes only mtime after the size replacement | pass |
| Period changes recompute status | Same intrinsic hit changes `PROCESS` to `OUTSIDE_PERIOD` | pass |
| Master state changes recompute status | Same hit produces `PROCESS_NEW_ROW` then `MASTER_COMPLETE` | pass |
| Force changes recompute status | Same hit changes `MASTER_COMPLETE` to `PROCESS` | pass |
| Recent-success state is current | Same intrinsic hit changes `PROCESS` to `FAST_SKIP` only after current recent cache is recorded | pass |
| Duplicate winner remains path/mtime-correct | Global newer duplicate in later folder is selected before early leaf publication | pass |
| `PROCESS` / `SOURCE_DUPLICATE` remain correct | Cold and warm duplicate result tuples are identical | pass |
| `OUTSIDE_PERIOD` remains correct | Period-context test and 200-candidate month fixture | pass |
| `PROCESS_NEW_ROW` remains correct | Current empty master lookup on a warm hit | pass |
| `MASTER_COMPLETE` remains correct | Current complete master lookup on a warm hit | pass |
| No stale final rows | All dynamic fields/actions are absent from cached values; cold/warm semantic tuples match | pass |
| Cache is bounded/session-only | `max_entries=2` eviction test plus controller-owned in-memory implementation | pass |

## Incremental Findings

- Five files across three containing folders produce exactly three leaf publications and five incremental row DTOs—not five bridge callbacks and not cumulative row rebuilding.
- The service continues one original `scan_reports` call and exposes immutable deltas through the already-existing bounded dashboard poll.
- A synthetic later-folder recent-cache lookup is blocked. While it is blocked, the first two folders are returned by `dashboard_state`, proving a later slow folder does not withhold completed earlier folders.
- A duplicate source in the early folder already has final `source_duplicate` status because all candidate metadata/period/duplicate context is completed globally before the first leaf publication.
- No row ID is duplicated. The final canonical DTO tuple is equal between cold and warm scans even when one warm source is size-invalidated.
- React renders leaf one and then merges leaf two while the original `scan_reports` Promise remains pending. A delayed lower revision cannot remove either row. The final response replaces deltas with canonical status-group ordering.
- Conflicting operations are rejected while scan state is mutable; close is safely refused during the deliberately blocked later leaf. Existing cooperative batch cancellation regressions remain green.

## Output Excerpts

```text
$ .venv/bin/pytest -q
1218 passed, 3 skipped in 110.63s (0:01:50)
```

```text
$ cd frontend && npm run test:frontend
# tests 32
# pass 32
# fail 0
```

```text
Focused incremental/cache/performance/application-service:
34 passed in 4.99s

Scan/controller/master/new-row regressions:
120 passed in 35.45s

PROMPT-030 runtime:
103 passed, 1 skipped in 1.49s
```

## Benchmark and Actual Cache Benefit

Five-run unprofiled medians; synthetic PPTX bytes are intentionally invalid. No wall-clock threshold is a test gate.

### 610 entries / 200 candidates / one containing folder

| Mode | Hits / misses / invalidations | Metadata reads | Leaf deliveries / rows | First leaf | Total | PPTX / COM |
|---|---:|---:|---:|---:|---:|---:|
| PROMPT-031 behavior | 0 / 0 / 0 | 200 | 0 / 0 | n/a | 39.330 ms | 0 / 0 |
| 031R cold | 0 / 200 / 0 | 200 | 1 / 200 | 35.716 ms | 46.045 ms | 0 / 0 |
| 031R warm | 200 / 0 / 0 | 200 | 1 / 200 | 38.413 ms | 48.790 ms | 0 / 0 |

### 30 / 120 / 300 candidate multi-leaf scaling

| Candidates / leaves | Mode | Hits / misses | Metadata | Deliveries | First leaf | Total |
|---:|---|---:|---:|---:|---:|---:|
| 30 / 3 | PROMPT-031 behavior | 0 / 0 | 30 | 0 | n/a | 8.167 ms |
| 30 / 3 | 031R cold | 0 / 30 | 30 | 3 | 5.917 ms | 11.021 ms |
| 30 / 3 | 031R warm | 30 / 0 | 30 | 3 | 6.684 ms | 11.917 ms |
| 120 / 6 | PROMPT-031 behavior | 0 / 0 | 120 | 0 | n/a | 26.382 ms |
| 120 / 6 | 031R cold | 0 / 120 | 120 | 6 | 20.650 ms | 47.155 ms |
| 120 / 6 | 031R warm | 120 / 0 | 120 | 6 | 22.373 ms | 48.875 ms |
| 300 / 10 | PROMPT-031 behavior | 0 / 0 | 300 | 0 | n/a | 64.172 ms |
| 300 / 10 | 031R cold | 0 / 300 | 300 | 10 | 48.195 ms | 149.796 ms |
| 300 / 10 | 031R warm | 300 / 0 | 300 | 10 | 53.923 ms | 151.580 ms |

Warm hits eliminate N occurrence-date derivations but do not eliminate N current metadata reads or any dynamic status work. On these already-fast Linux fixtures, warm total time is about 1–8% slower than cold rather than faster. This is reported as **no demonstrated total-time cache speedup**. The required incremental list construction also adds work relative to PROMPT-031's final-only behavior, although each incremental row is built exactly once and the final result remains linear. The meaningful measured UX property is time-to-first-leaf before total completion when later leaf work is slow.

## Windows Source Retest Procedure

Use a disposable report copy and the React + pywebview development launcher; do not build or publish Portable/LAN artifacts.

### 1. Verify and launch the exact source

From PowerShell in the repository:

```powershell
Set-Location 'C:\path\to\TNP'
git branch --show-current
git rev-parse HEAD
# Confirm branch arena/2e088bee-tnp and the final SHA reported with this change.

Push-Location frontend
npm.cmd ci --ignore-scripts
npm.cmd run typecheck
npm.cmd run build
Pop-Location

& .\.venv-webview\Scripts\python.exe -m app.desktop --debug
```

Do not run `BUILD_AND_PUBLISH.bat`, `build_portable.bat`, or any update publishing command.

### 2. Cold multi-leaf Quét

Prepare one root with at least three sorted final containing folders, for example `01_early`, `02_middle`, and `99_late_many`. Put enough disposable report copies in `99_late_many` that its status work is visibly later/slower.

1. Chọn tháng matching the test Management Numbers.
2. Chọn the multi-leaf root folder.
3. Bấm **Quét** exactly once.
4. Confirm rows from `01_early` become visible while the app still shows **Đang quét…** and before `99_late_many` finishes.
5. Confirm rows never duplicate or disappear and the final ordering/statuses settle deterministically.
6. Record the contiguous `logs/app.log` block from `SCAN_START` through `SCAN_RETURN` as the **cold** block.

### 3. Warm Quét in the same running app

Without restarting, changing month, changing folder, or changing files:

1. Bấm **Quét** again.
2. Confirm the same final rows, order, statuses, duplicate winner, and counts.
3. Record the second contiguous `SCAN_START` through `SCAN_RETURN` block as the **warm** block.
4. Cold should show intrinsic misses; warm should show hits. Both must still show one metadata read per candidate and zero PPTX/COM work.

### 4. Size and path invalidation on disposable copies

To prove same-mtime size invalidation, select one disposable PPTX copy and run:

```powershell
$probe = Get-Item 'C:\disposable\reports\02_middle\261010102-VOC_middle.pptx'
$mtime = $probe.LastWriteTimeUtc
[System.IO.File]::AppendAllText($probe.FullName, ' ')
$probe.LastWriteTimeUtc = $mtime
```

Run Quét again. The file must be an intrinsic miss/invalidation, not a hit, while current final status remains semantically correct. Do not process this deliberately modified disposable PPTX.

Then move/rename a fresh disposable copy while preserving its timestamp and Quét again. The new normalized path must miss as a distinct source; it must not reuse the old path's intrinsic entry or duplicate row ID.

### 5. Required log markers

Capture these markers for cold, warm, and invalidation runs:

```text
SCAN_START
SCAN_ENUMERATE_START
SCAN_ENUMERATE_END elapsed_ms=... entries=... candidates=... traversals=...
SCAN_FILTER_START
SCAN_LEAF_START index=... total=... folder=<safe-basename+hash> files=...
SCAN_LEAF_END index=... total=... folder=<safe-basename+hash> files=... elapsed_ms=...
SCAN_LEAF_DELIVER index=... total=... rows=... elapsed_ms=... generation=... revision=...
SCAN_FILTER_END elapsed_ms=... candidates=... accepted=... skipped=... sequential=true
SCAN_BUILD_RESPONSE_START
SCAN_BUILD_RESPONSE_END elapsed_ms=... files=... redaction_contexts=1 redaction_sources=...
SCAN_CACHE_SUMMARY hits=... misses=... invalidations=... unavailable=... stores=... entries=... evictions=... intrinsic_derivations=...
SCAN_INCREMENTAL_SUMMARY leaf_folders=... leaf_completed=... leaf_deliveries=... rows_delivered=... first_leaf_ms=... leaf_ms=...
SCAN_METADATA elapsed_ms=... reads=... errors=... master_open_ms=...
SCAN_PARSE_PPTX elapsed_ms=... opens=0
SCAN_POWERPOINT_COM elapsed_ms=... starts=0
SCAN_TOTAL elapsed_ms=... finished=True ... traversals=1 ...
SCAN_SERIALIZE elapsed_ms=unavailable boundary=pywebview_after_python_return
SCAN_RETURN elapsed_ms=... ok=True bridge_calls=1
```

Acceptance expectations:

- `leaf_folders == leaf_completed == leaf_deliveries`;
- `rows_delivered == candidates`/returned rows for the tested eligible fixture;
- warm `hits` equals unchanged cache-eligible candidates and `intrinsic_derivations=0`;
- current `metadata reads` remain equal to candidates on both cold and warm scans;
- invalidated same-path size/mtime entries count as misses and increment `invalidations`;
- moved paths count as misses, never unsafe hits;
- `first_leaf_ms < SCAN_TOTAL elapsed_ms` when a later leaf is genuinely slow;
- only safe folder basenames/hashes appear in new `SCAN_LEAF_*` lines;
- `opens=0`, COM `starts=0`, `traversals=1`, and `redaction_contexts=1`.

If `SCAN_LEAF_DELIVER` appears but no rows become visible before `SCAN_RETURN`, capture the browser/runtime console plus the full block; that isolates pywebview/frontend delivery from backend leaf publication.

## Residual Risks

- Real Windows pywebview must actually execute the long `scan_reports` call and dashboard polls concurrently for visual incremental delivery. Linux thread tests prove service locking/snapshots, not the Windows host scheduler.
- A very fast scan can finish before the 750 ms polling interval; in that case users correctly receive the final aggregate directly rather than observing a transient partial list.
- Global discovery/metadata/period/duplicate work intentionally precedes first leaf publication. This is required so an early duplicate status never becomes wrong later.
- The cache's demonstrated work saving is tiny and does not justify widening its value or skipping current dynamic checks.
- Real master workbook latency, UNC/removable behavior, and a deliberately slow later folder remain Windows acceptance items.

## Recommendation

Proceed with the exact Windows source retest above. Promote the result from partial to verified only after earlier leaf rows are visibly present before a slow later leaf completes, cold/warm final states match, path/size invalidation logs are correct, and all required zero-parse/zero-COM/one-redactor/one-traversal markers are captured.
