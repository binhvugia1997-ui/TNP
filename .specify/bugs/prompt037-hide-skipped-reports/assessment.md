# Bug Assessment: Skipped reports (“Bỏ qua — đã cập nhật”) remain in “Danh sách báo cáo”

- **Slug**: prompt037-hide-skipped-reports
- **Created**: 2026-10-13
- **Source**: pasted text (PROMPT-037)
- **Verdict**: valid (reproduced on Linux, development evidence only)
- **Severity**: medium (UI noise and inflated “Tổng” on large folders; no data loss, no processing impact)
- **Starting SHA**: `2dbf355` (base branch `arena/2507b6dc-tnp`, includes PROMPT-036R + PROMPT-036S, version 1.3.6 Build 019)
- **Working branch**: `arena/9eb765c7-tnp`
- **Spec Kit execution**: Executed from checked-in Spec Kit skill instructions; native Arena skill invocation was not available.
- **Extension hooks**: `.specify/extensions.yml` has `hooks: {}` — no `before_bug_assess` / `after_bug_assess` hooks registered; none executed.

## Report (summarized)

Reports whose final/current application status is “Bỏ qua” (`skipped`) stay visible in
“Danh sách báo cáo”, inflate the displayed total, and can remain selected or active. Requirement:
a skipped report must disappear from the user-visible list and must not come back through polling,
incremental scan deltas, or rescans while it stays skipped. Source files, scan/history state, and
the skip classification rules must not change. No CSS-only hiding.

## Symptom

Expected: a row with status `skipped` is not rendered in the report table, is not counted in the
visible total, cannot be selected or active, and is not reinserted by later reconciliation.
Actual: the row is rendered with status “Bỏ qua — đã cập nhật”, counted in “Tổng số báo cáo” /
footer “Tổng”, and the table header reads `N/M` with M including it.

## Reproduction

1. Start the real store (`frontend/src/state/store.tsx`) against a mocked `pywebview.api`
   whose `get_dashboard_state` returns three rows: `waiting` (id 0), `skipped` (id 1), `completed` (id 2).
2. Let the bridge connect and the first dashboard poll apply.
3. Observe `reports`, `stats`, and `queue`.

Observed (probe run on the base commit, `node` + Vite SSR + JSDOM, probe file not committed):

```
visible ids: 0,1,2 stats.total: 3 skippedBucket: 1 queue: [ '0' ]
```

The skipped row (id 1) is in the visible list and counted in `stats.total`. Reproduced: **yes**.

## Skipped-report flow (traced)

### 1. Where `status = "skipped"` is created

- **Pre-scan classification** — `app/prescan.py:702`: a discovered file whose Excel master row is
  already complete gets `it.action = ACTION_MASTER_COMPLETE` (“Bỏ qua — đã cập nhật (dòng N)”).
  `app/prescan.py:75` maps it to pipeline key `skipped`.
- **Batch processing** — `app/batch_processor.py:330`, `:402`, `:418`: `fr.status = "skipped"` for
  complete-master rows (no Qwen, no extraction), then `self._remember(mgmt, path, "skipped")` at
  `:336` / `:407` (recent-success cache, `app/prescan.py:42` `SAFE_CACHE_STATUSES`).
- **DTO mapping** — `app/application_service.py:530-546` `_status_key()`:
  `final_map["skipped"] → "skipped"` (when `run_stage` is final) and
  `action_map[ACTION_MASTER_COMPLETE] → "skipped"` (when no run stage). This is the only place
  the DTO `status` string is produced for report rows.

### 2. Why skipped rows currently reach the visible list

- `app/application_service.py:393` `_dashboard_state_unlocked` calls `c.scan_rows("Tất cả file đã quét")`
  which is the **unfiltered** list (`app/gui_controller.py:1003-1025`: the “all” filter applies no
  row filter). Every row, including `status: "skipped"`, is serialized by `_report_dto` (`:490`).
- Frontend: `frontend/src/state/store.tsx:529-531` `acceptDashboard` stores `data.reports` as-is
  (full path) or via `mergeIncrementalReports` (incremental path). Nothing removes `skipped`.
- `filteredReports` (`store.tsx:791-803`) only filters by search and the selected FILTERS option;
  the default filter `Tất cả file đã quét` returns everything. `ReportTable.tsx:11` renders `filteredReports`.
- Conclusion: **no filter on status exists for the default list.** The only place `skipped` is
  treated specially is the explicit “File bị bỏ qua” filter (`store.tsx:799`), which is meant to
  show skipped rows and is therefore the opposite of the requirement.

### 3. Can incremental scan deltas reinsert them? — **Yes (reinsertion path proven)**

- The backend emits per-leaf deltas: `app/application_service.py:263-282` `publish_leaf` builds a
  snapshot via `_dashboard_state_unlocked` and extends `_scan_pending_reports` with its rows, which
  include `skipped` rows whose folder completed in that leaf.
- Frontend `store.tsx:164-172` `mergeIncrementalReports(current, incoming)`:
  `byId.set(item.id, item)` for every incoming row. A row that transitions visible → skipped in a
  delta **overwrites** its visible copy with the skipped copy and nothing removes it, so it remains
  in the list. A skipped row that first appears in a delta is added.
- Therefore incremental updates can both introduce and keep skipped rows visible.

### 4. Can dashboard polling reinsert them? — **Yes**

- `store.tsx:522-575` `acceptDashboard` is the single ingestion point for `get_dashboard_state`,
  `scan_reports`, and the connect handshake. `loadDashboard` (`:~580`) is called from the poll
  (`store.tsx:~690` `poll`) and `runOperation` (`:~800`). Each call stores the DTO rows unfiltered,
  so every poll re-presents skipped rows.

### 5. Can `selectedIds` / `activeId` retain a skipped report? — **Yes**

- `store.tsx:565-570`: `selectedIds` is reset only when `nextReports.length === 0`; otherwise the
  previous selection survives unchanged even if a selected id is now skipped.
  `activeId` is kept when `nextReports.some(r => r.id === current)` — a skipped row is still in
  `nextReports`, so it stays active; `activeReport` (`store.tsx:738`) then points at it and
  `ReportDetail` displays it.
- `toggleAllSelected` (`store.tsx:929`) receives `allIds` from `ReportTable.tsx:11` (`filteredReports`),
  so select-all includes skipped rows today.
- `changeSelected` (`store.tsx:919-923`) would send skipped ids to `exclude_reports` /
  `restore_reports` if they remain selected.

### 6. Which visible counters currently include skipped rows?

UI-facing, derived from the store's `reports` state (all currently include skipped rows):

| Counter | Where | Current source | Includes skipped? |
| --- | --- | --- | --- |
| “Tổng số báo cáo: N” | `LogAndStats.tsx:61` (StatsCard) | `stats.total = reports.length` (`store.tsx:775`) | **Yes** |
| footer “Tổng: N” | `FooterBar.tsx:26` | `stats.total` | **Yes** |
| “N/M file” header | `ReportTable.tsx:19` | `rows.length / s.reports.length` | **M yes** |
| “Bỏ qua” stat card and sentence | `LogAndStats.tsx:72,76` | `stats.skipped` (bucket `skipped`, `store.tsx:776`) | **Yes** (and outside_period / source_duplicate / fast_skip) |
| “hàng đợi N”, footer/processing queue count | `ReportTable.tsx:19`, `FooterBar.tsx:29`, `ProcessingCard.tsx:30` | `queue` (`store.tsx:786`) | Only during an active run (`run.queue.slice(doneCount)`): a candidate that becomes `skipped` at batch time stays counted until the worker advances past it |

Backend / internal counters (**must not change**; they describe scan and processing, not the list):
`BatchSummary.skipped` (`app/batch_processor.py:99,318,334,405,421`), the “Bỏ qua: N” log and summary
text (`batch_processor.py:262,721`, `gui_controller.py:1076,1086,1818,1836`), `SCAN_TOTAL status_skipped`
(`application_service.py:360`), `ScanMetrics.status_skipped_files` (`gui_controller.py:982`), and
the prescan count `master_complete` (`prescan.py:491`). `run.queue.length` in `ProcessingCard.tsx:25-28`
is a run-scoped progress denominator, not the visible list.

## Root Cause Hypothesis

**Confidence: high (demonstrated).** The skip classification is correctly emitted as a DTO
`status` string, but **nothing at the presentation boundary removes that status from the
user-visible report list**. The single ingestion point, `acceptDashboard`
(`store.tsx:522`) with `mergeIncrementalReports` (`store.tsx:164`), stores every DTO row unfiltered.
Because all downstream state (`reports`, `filteredReports`, `stats`, `queue`, `activeReport`,
select-all) is derived from that unfiltered `reports` state, the skipped rows propagate into the
table, the totals, selection, and the active row, and every poll or delta reintroduces them.

The backend is **not** the right place to drop them: the incremental delta protocol is
keyed by row id and only overwrites or adds rows. If the backend simply omitted skipped rows
from a delta, the frontend could not tell “became skipped” from “unchanged”, so a previously
visible row would linger. The DTO must keep reporting `status: "skipped"` so the presentation
layer can reconcile transitions.

## Proposed Remediation

**Preferred (frontend presentation boundary, one authoritative rule):**

1. Add one pure rule in `frontend/src/types.ts` (or a dedicated module): `isVisibleReport(report)`
   returns `false` only for `status === 'skipped'`, plus `visibleReports(list)` that filters and
   renumbers `stt` in O(N). This is the only place the hide rule exists.
2. In `acceptDashboard`, apply it after the full-replace or the incremental merge, so
   `reportsRef` and `reports` hold only visible rows. In `mergeIncrementalReports`, merge first
   (so a visible → skipped transition overwrites the stale visible copy) and filter/renumber
   afterwards.
3. Reconcile selection and active row in the same pass: prune `selectedIds` to visible ids
   (keeping state identity when unchanged); if `activeId` is no longer visible, move it to the next
   visible row after its previous position, else the previous visible row, else clear it.
4. Derive `queue` (the running branch) from visible ids only, so a candidate that became skipped
   is not counted as pending. Leave `run.queue.length` (run progress) unchanged.
5. Remove the `skipped` case from the “File bị bỏ qua” filter list for a single consistent rule
   (rows never reach `filteredReports` with `status === 'skipped'` anyway).
6. Keep the backend DTO, `_status_key`, skip classification, batch summary, scan diagnostics,
   and file system untouched.

**Alternatives** (considered, rejected):
- *Backend omits skipped rows from `reports`:* breaks incremental transitions (see root cause);
  also changes `scan_rows`/DTO semantics used by diagnostics and the report-details bridge.
- *Backend emits a `hidden` flag:* adds a second contract field to keep in sync; the frontend
  would still need the same rule for stored rows. Rejected in favour of one rule.
- *CSS `display:none`:* explicitly forbidden; still counts toward totals and selection.

**Files likely to change**:
- `frontend/src/types.ts` (the visibility rule)
- `frontend/src/state/store.tsx` (ingestion, merge, selection/active reconciliation, stats/queue/filter)
- `frontend/src/components/report/ReportTable.tsx` only if needed (expected: no change; it already renders `filteredReports`)
- `frontend/tests/prompt037-hide-skipped.test.mjs` (new regression tests)
- `frontend/package.json` (add `test:prompt037` script and include it in `test:frontend`)

**Tests to add**:
- Rule: waiting / completed / needs_review / error → visible; skipped → absent; mixed list keeps only skipped removed; 1000+ mixed rows give exact visible count within a time bound; `stt` renumbered.
- Store (real `store.tsx` via Vite SSR + JSDOM, mocked bridge): dashboard refresh cannot reinsert skipped; incremental delta (visible → skipped and skipped first-seen) cannot reinsert; rescan with the same skipped classification stays hidden; skipped row cannot remain selected; active row becoming skipped advances safely; select-all (rendered `ReportTable`) excludes skipped; visible totals (`Tổng`, `N/M`, bucket counts) match the visible list.
- Safety: hide flow issues no `pywebview.api` call other than read/scan methods (no delete/exclude); the source PPTX is not deleted (no delete-capable call).
- Classification unchanged: Python test that `ApplicationService._status_key(ACTION_MASTER_COMPLETE, …)` still returns `"skipped"` and that `_report_dto` still carries `status: "skipped"` (DTO preserved for processing / transitions).

## Risks & Considerations

- **Scope of “skipped”:** the requirement names `skipped` / “Bỏ qua”. The status literal `skipped`
  (“Bỏ qua — đã cập nhật”) is hidden. The other skip-family statuses — `outside_period`
  (“Bỏ qua ngoài thời gian xử lý”), `source_duplicate`, and `fast_skip` (“Bỏ qua nhanh — đã xử lý gần đây”)
  — have distinct labels and are **not** hidden by this change (see Open Questions).
- **Stats “Bỏ qua” card** will then count only the non-hidden skip-family statuses; this is consistent
  with the visible list but its meaning shrinks. Documented in the fix report.
- **Performance:** the rule is O(N) per reconciliation (one filter + renumber + Set lookups for
  selection/active). The existing incremental merge already sorts O(N log N) per delta; this change
  does not add a bridge call, a filesystem call, or a per-row backend call. The PROMPT-031 incremental
  path (`_scan_pending_reports`, `publish_leaf`) is untouched.
- **Stale detail:** a row hidden while its details are cached is simply absent from `reports`; the
  detail fetch effect keys on `activeReport`, which is now a visible row.
- **Windows acceptance:** Linux JSDOM and Vite SSR evidence only; real pywebview on Windows not run.
- **PROMPT-036S Learning tab:** not touched.

## Open Questions

- [NEEDS CLARIFICATION: Should `outside_period`, `source_duplicate`, and `fast_skip` (other "Bỏ qua…" statuses)
  also be hidden from “Danh sách báo cáo”? This fix hides only the `skipped` status literal (“Bỏ qua — đã cập nhật”),
  which is the status the requirement names. Changing it is a one-line change to the visibility rule's set.]
