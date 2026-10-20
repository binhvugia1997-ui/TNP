# Bug Assessment: Incremental scan cache identity is insufficient for complete list rows

- **Slug**: prompt031r-incremental-scan-cache
- **Created**: 2026-10-10
- **Source**: pasted text (PROMPT-031R)
- **Verdict**: valid
- **Severity**: high
- **Starting SHA**: `adf61f6f4d9dee24322a5ef07a50cdcf98825365`
- **Working branch**: `arena/2e088bee-tnp`
- **Implementation status**: stopped before production edits because the required Manager+mtime-only identity is proven unsafe
- **Spec Kit execution**: Executed from checked-in Spec Kit skill instructions; native Arena skill invocation was not available.

## Report (summarized)

PROMPT-031R requests two related scan improvements while preserving all current list semantics and the PROMPT-031 linear response fix:

1. Cache complete scan/list-row results in application-session memory under exactly `(Management Number, file mtime)`. The same pair must hit; a new Manager or changed mtime must miss and replace. Path, size, and content hash must not be added silently. If correctness requires another identity component, stop and report.
2. Publish incremental results after each completed final containing/leaf folder, never after each file, and finish with an aggregate result equivalent to a correct complete scan.

The cache must not produce stale month/status/master/force/duplicate results or cross-scan contamination. Discovery must continue to perform zero PPTX parsing and zero PowerPoint COM work. Incremental delivery must use a bounded pywebview-safe mechanism rather than per-file callbacks or polling.

## Symptom and Expected Behavior

A cold recursive Quét waits for all discovery, classification, and response construction before the React list receives any rows. Repeated Quét operations also redo classification work even when sources are unchanged. The expected behavior is session-local reuse for unchanged rows and visible progress at leaf-folder boundaries, while the final result remains byte-for-byte equivalent in meaningful row fields and deterministic ordering to a correct complete scan.

The performance/usability request is valid. However, the mandated cache identity is not a one-to-one identity for the complete row result it is required to cache. Implementing that identity as specified would make stale or misassigned rows possible under existing supported behavior.

## Reproduction

### Read-only deterministic identity proof

A temporary no-PPTX-parse script called the existing `prescan()` function with unchanged source identity values. It produced:

```text
same_key= ('260915001-VOC', 1700000000.0)
period_actions= PROCESS OUTSIDE_PERIOD
master_actions= PROCESS_NEW_ROW MASTER_COMPLETE PROCESS
duplicate_keys_equal= True [('260915002-VOC', 1700000500.0), ('260915002-VOC', 1700000500.0)]
duplicate_actions= ['PROCESS', 'SOURCE_DUPLICATE']
```

This proves all of the following:

1. One unchanged `(Manager, mtime)` pair is `PROCESS` for September and `OUTSIDE_PERIOD` for October.
2. The same pair is `PROCESS_NEW_ROW`, `MASTER_COMPLETE`, or `PROCESS` depending on current master rows and force mode.
3. Two simultaneously present source files can have exactly the same approved key while requiring different complete row statuses: one `PROCESS`, one `SOURCE_DUPLICATE`.
4. The duplicate winner is selected by normalized path when mtimes tie. A complete cached row therefore cannot be assigned to the correct current file without using path (or an equivalent occurrence discriminator) during identity/matching.

The existing cache regression was also executed:

```text
$ .venv/bin/pytest -q \
    tests/test_prescan.py::test_changed_mtime_or_size_is_cache_miss \
    tests/test_prescan.py::test_source_duplicates_newest_wins_and_tie_is_deterministic
..                                                                       [100%]
2 passed in 0.45s
```

The first test proves current semantics intentionally treat a same-Manager, same-mtime source whose size changed as `source_changed`. The second proves deterministic path-based duplicate selection for equal mtimes. A new complete-row cache that hits solely because Manager+mtime is unchanged would bypass or overwrite both decisions.

No production source was modified during this reproduction.

## Suspected Code Paths

- `app/prescan.py:437-564` — a row action depends on period, the complete source duplicate set, current recent-success cache state, current master workbook rows, and force mode; it is not a pure function of Manager+mtime.
- `app/prescan.py:288-307` — the existing `FastScanCache.lookup()` rejects source changes using normalized path, size, and mtime. It is a separate persisted recent-success optimization, but its source-change rule remains part of current list semantics.
- `app/gui_controller.py:873-943` — Quét discovers the whole tree, opens the current master in probe mode, instantiates the recent-success cache, then performs one global `prescan()`.
- `app/gui_controller.py:959-985` — list rows preserve source path identity and sort stably by status group; path is also used for manual exclusion identity.
- `app/application_service.py:172-252` — the service holds its `RLock` through the entire synchronous scan and response build.
- `app/application_service.py:270-327` — dashboard polling takes the same lock and builds rows from the current complete `scan_result`, so current polling cannot observe leaf completion while Quét owns the lock.
- `frontend/src/lib/startup.ts:createBridgeSessionController()` — React already has one bounded 750 ms dashboard poll with one request in flight; this is the preferred transport rather than adding bridge calls per file.
- `frontend/src/state/store.tsx:acceptDashboard()` — each poll replaces the report array wholesale and has no scan-generation guard, so incremental snapshots would require generation-aware acceptance to prevent an older scan from replacing a newer one.
- `app/scanner.py:46-102` — discovery is one deterministic recursive `os.walk`; final containing folders can be derived from candidates without another traversal and without opening PPTX contents.

## Root Cause Findings

**Confidence: high.** There are two separate architectural findings.

### 1. Complete list rows have a compound identity and context

`(Management Number, mtime)` is not unique within one scan. Equal-Manager/equal-mtime duplicate sources are legal and are differentiated by path. It is also not sufficient across scans: the selected period, force flag, current master workbook state, and existing path+size+mtime recent-success decision all alter the final row while the requested pair remains unchanged.

A map keyed only by the approved pair has no correct complete-row value for the demonstrated duplicate case. Storing a bucket and matching its members by path, adding an occurrence ordinal derived from path order, or bypassing colliding pairs would respectively add path/equivalent identity or violate the required same-pair-hit rule. Recomputing all dynamic decisions after a hit would avoid stale rows but would no longer cache the complete scan/list result or avoid the intended downstream status work.

The same-mtime/changed-size case is independently unsafe. Current behavior must miss the existing recent-success cache and process the changed source. A Manager+mtime complete-row hit could return the previous fast-skip/complete status instead.

### 2. Incremental leaf delivery is feasible but requires snapshot isolation

The existing 750 ms polling loop is bounded and suitable. The immediate blockers are the service lock and mutable controller state, not a lack of transport. A safe design would publish immutable, generation-tagged snapshots under a short lock and let `get_dashboard_state` return the latest snapshot while the original batched `scan_reports` call remains in progress. React would accept only the active generation and replace deterministic prefixes/snapshots, not merge rows by filename ad hoc.

Global duplicate correctness means leaf rows cannot be finalized merely as files are encountered. Discovery and the cheap global metadata/duplicate phase must first know all candidates. Status work can then be completed by final containing-folder groups in deterministic discovery order, with one publish per completed group. The final aggregate must still be produced by the same canonical classification path. This preserves one traversal, zero PPTX/COM discovery work, and correct duplicates, though time-to-first-leaf starts after the required global discovery/duplicate metadata phase.

The incremental design can be implemented after the cache identity contract is resolved; implementing it alone now would contradict the explicit instruction to stop when cache correctness needs another identity component.

## Proposed Remediation

**STOP — user approval is required before production changes.** The exact Manager+mtime-only complete-row identity cannot satisfy the required semantics. Do not silently implement a broader key, path-based bucket matching, collision bypass, or stale row reuse.

**Preferred after approval**:

1. Use a session-memory source identity containing at least normalized source path, Management Number, size, and mtime. This matches the already-established source-change semantics and distinguishes equal-Manager/equal-mtime duplicates.
2. Keep scan context separate from source identity, but invalidate or version cached row decisions when period/range, force mode, template/output/master workbook state, or recent-success-cache state changes. If the cache is expected to avoid reopening the master, define an explicit safe master snapshot/version contract first.
3. Cache immutable row-reconstruction data, not a boolean. Recompute deterministic aggregate counters/order from current candidates and never reuse manual-selection or job-run state as source classification.
4. Refactor classification into a global metadata/duplicate phase followed by deterministic final-containing-folder status phases. Publish exactly one immutable generation-tagged dashboard snapshot per completed leaf and one final snapshot.
5. Reuse the existing dashboard polling loop. Shorten service lock scope around snapshot publication/read only; do not invoke JavaScript from a worker thread, add per-file bridge calls, or add per-file React polling.
6. Extend PROMPT-031 metrics with bounded `SCAN_CACHE_*`, `SCAN_LEAF_*`, and `SCAN_INCREMENTAL_SUMMARY` aggregates and safe basename/hash labels only.

**Alternatives requiring an explicit contract change**:

- Keep exactly Manager+mtime only for immutable intrinsic facts such as parsed Management Number/date, while recomputing duplicate, period, recent-cache, master, and force decisions every scan. This can be safe but is not a complete list-row cache and is unlikely to provide the requested downstream-work avoidance.
- Mark duplicate-key collisions and same-mtime source changes uncacheable, and flush on every context change. This preserves correctness but violates the unconditional same-pair-hit requirement and reduces warm-cache coverage.
- Accept stale/collision risk. This is not recommended and conflicts with the explicit semantic-preservation requirements.

**Files likely to change after the contract is resolved**:

- `app/prescan.py`
- `app/gui_controller.py`
- `app/application_service.py`
- `app/scan_diagnostics.py`
- `frontend/src/state/store.tsx`
- frontend dashboard DTO/types as needed
- `tests/test_prescan.py`
- `tests/test_scan_performance.py`
- `tests/test_application_service.py`
- frontend store/startup tests

**Tests to add or update after approval**:

- cold scan then warm scan: exact final-row equivalence and operation-count reduction;
- changed Manager, changed mtime, changed size with preserved mtime, moved/renamed path, period change, force change, master-row change, and recent-success-cache change;
- same Manager+same mtime in two leaf folders with deterministic path tie-breaking and no cached-row aliasing;
- cache isolation across roots, configurations, cancelled scans, corrupt/ineligible sources, and generations;
- three or more leaf folders with exactly one publish per completed leaf, no per-file events/calls, and earlier-leaf visibility before a delayed later leaf completes;
- final incremental snapshot equals canonical non-incremental complete result in IDs/order/status/reason/counters;
- cancellation mid-leaf and between leaves, stale-generation rejection, repeated scan, and source mutation during scan;
- 30/medium/large cold/warm work counts and timings, with time-to-first-leaf evidence;
- preserved one traversal, one required metadata read per candidate, zero PPTX parse, zero COM starts, and PROMPT-031 redaction/diagnostic bounds.

## Risks & Considerations

- A stale complete-row hit can suppress required processing or show an incorrect duplicate/month/master status. This is a business-correctness risk, not merely a conservative cache miss.
- Path is bridge-sensitive data. It may be used internally for identity but must remain redacted from logs and DTO strings exactly as PROMPT-031 requires.
- A master workbook can change while source files do not. Source identity alone cannot validate master-derived statuses.
- Incremental publication from mutable controller structures would introduce torn snapshots or cross-generation contamination unless snapshots are immutable and lock scope is explicit.
- Publishing a leaf before the complete duplicate set is known can later reverse its winner. The global duplicate phase must precede final leaf publication.
- Windows pywebview behavior, real filesystem latency, and PowerPoint/UNC/removable-drive acceptance cannot be claimed from Linux.
- The full Linux suite, frontend suite, compile/lint, benchmarks, and Windows acceptance workflow were intentionally not run because no fix was applied after the mandated stop condition was proven.

## Open Questions

- [NEEDS USER DECISION: May the session cache identity be expanded to normalized path + Management Number + size + mtime, with scan-context versioning/invalidation outside that source key?]
- [NEEDS USER DECISION: If the identity must remain exactly Manager+mtime, may colliding/unsafe pairs be explicitly uncacheable and may all context-dependent decisions be recomputed, accepting that this is not a complete-row cache?]
- [NEEDS CONTRACT DEFINITION: Which master-workbook change token is acceptable for warm reuse of master-derived statuses, or must the current workbook remain authoritative and be read every Quét?]
- [NEEDS WINDOWS RETEST AFTER A FIX: Does the existing pywebview host service concurrent `scan_reports` and `get_dashboard_state` calls as expected, and what are the resulting `SCAN_LEAF_*` time-to-first-leaf measurements?]
