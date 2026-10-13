# Bug Fix: Skipped reports (“Bỏ qua — đã cập nhật”) hidden from “Danh sách báo cáo”

- **Slug**: prompt037-hide-skipped-reports
- **Fixed**: 2026-10-13
- **Assessment**: ./assessment.md
- **Status**: applied
- **Spec Kit execution**: Executed from checked-in Spec Kit skill instructions; native Arena skill invocation was not available.

## Summary

Added one presentation rule (`isVisibleReport` / `presentReports` in `frontend/src/types.ts`) that
hides rows with status `skipped` from the user-visible report list, and applied it at the single
ingestion point of the dashboard DTO in the store (`acceptDashboard`, and the incremental merge it
uses). Selection and the active row are reconciled in the same pass, and the visible queue
excludes skipped candidates. The backend, skip classification, DTO contract, scan diagnostics,
and source files are unchanged.

## Changes

| File | Change | Notes |
| --- | --- | --- |
| `frontend/src/types.ts` | added | `HIDDEN_REPORT_STATUSES = {skipped}`, `isVisibleReport`, `presentReports` (O(N), filters and renumbers `stt` over visible rows only). The only definition of the hide rule. |
| `frontend/src/state/store.tsx` | modified | `mergeIncrementalReports` overwrites by id first, then presents (visible → skipped transitions drop the stale row). `acceptDashboard` uses `presentReports` for full DTOs, prunes `selectedIds` to visible ids (state identity kept when unchanged), and sets `activeId` via `resolveActiveReportId`. New exported pure `resolveActiveReportId` (keep / next visible / previous visible / first / ''). `queue` in the running branch counts only visible candidates. The “File bị bỏ qua” filter list drops `skipped` (rows are never present). |
| `frontend/tests/prompt037-hide-skipped.test.mjs` | added | 16 regression tests (rule, store with mocked bridge, rendered `ReportTable`/`StatsCard`, 1000+ rows). |
| `frontend/package.json` | modified | `test:prompt037` script; file added to `test:frontend`. No dependency change. |
| `tests/test_prompt037_skipped_dto.py` | added | 3 backend guard tests: classification unchanged, DTO keeps `status: "skipped"`, building DTO rows never touches source PPTX files. |
| `.specify/bugs/prompt037-hide-skipped-reports/{assessment,fix,test}.md` | added | Bug workflow reports. |

No Python runtime file, no version/build, no Learning code, no TAT_QPN_PHASE6 file was touched.

## Diff Highlights

```ts
// frontend/src/types.ts — the one rule
export const HIDDEN_REPORT_STATUSES: ReadonlySet<StatusKey> = new Set<StatusKey>(['skipped'])
export function isVisibleReport(report: Pick<Report, 'status'>): boolean {
  return !HIDDEN_REPORT_STATUSES.has(report.status)
}
```

```ts
// frontend/src/state/store.tsx — acceptDashboard (excerpt)
const previousReports = reportsRef.current
const nextReports = data.scan.incremental
  ? mergeIncrementalReports(previousReports, data.reports)   // merge, then presentReports
  : presentReports(data.reports)
reportsRef.current = nextReports
...
const visibleIds = new Set(nextReports.map((report) => report.id))
setSelectedIds((current) => { const kept = current.filter((id) => visibleIds.has(id)); return kept.length === current.length ? current : kept })
setActiveIdValue((current) => resolveActiveReportId(previousReports, nextReports, current, visibleIds))
```

## Tests Added or Updated

- `frontend/tests/prompt037-hide-skipped.test.mjs` — 16 tests, see `test.md` for the list and results.
- `tests/test_prompt037_skipped_dto.py` — 3 tests (`_status_key` precedence, DTO status preserved, no source deletion).

## Local Verification

- `npm run typecheck` (frontend) → exit 0.
- `node --test tests/prompt037-hide-skipped.test.mjs` → 16/16 pass after the fix; 15/16 tests fail on the base `src/` (red check; the only passing test is the classification guard, see `test.md`).
- `npm run test:frontend` → 55 tests, 54 pass, 1 pre-existing unrelated failure (version metadata assertion, see `test.md`).
- `pytest` on scan/incremental/application-service suites → 67 pass, 1 pre-existing unrelated failure (same version assertion).
- `py_compile tests/test_prompt037_skipped_dto.py` → ok. `git diff --check` → clean.

## Deviations from Assessment

1. **STT renumbering on the full-replace path.** The assessment did not call this out. Previously
   the full DTO path kept the backend `stt` (which counted every row, including hidden ones), so
   the list would have shown gaps after hiding. `presentReports` renumbers over visible rows, the
   same way the incremental path already did. The backend `stt` field is still sent and unchanged.
2. **Queue during a run.** The assessment proposed filtering the running queue; this is implemented
   in `queue` only. `run.queue.length` in `ProcessingCard` (run progress denominator) is intentionally
   unchanged.
3. **Filter option “File bị bỏ qua”.** The `skipped` entry was removed from its list. Without this,
   the list would have two different hide rules. `outside_period`, `source_duplicate`, and `fast_skip`
   stay in that filter and remain visible (scope decision, see Open Questions in the assessment).
4. No backend change was needed. The assessment's “files likely to change” did not include Python,
   and none was changed. Hiding only at the frontend boundary is required because an incremental
   delta must still carry the skipped row to express a visible → skipped transition.

## Follow-ups

- Decide whether `outside_period`, `source_duplicate`, and `fast_skip` should also be hidden; if yes,
  change only `HIDDEN_REPORT_STATUSES` and update the affected tests.
- Pre-existing failing test `frontend/tests/prompt027-overlays.test.mjs` (version `1.3.5` expected) and
  `tests/test_application_service.py::test_configuration_ollama_update_state_and_disabled_install`
  (same `1.3.5` expectation) are unrelated to this bug and were not modified (version/build is out of scope).
- Windows / pywebview acceptance of the hide behaviour with a real scan has not been performed.
