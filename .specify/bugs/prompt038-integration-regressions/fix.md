# PROMPT-038 fix report

Executed from checked-in Spec Kit skill instructions; native Arena skill invocation was not available.

## Bug A — report-list viewport not visible at runtime
- Proven cause: stale on-disk `frontend/dist` (gitignored/untracked) served silently in source
  mode. Fresh build + real-Chromium measurement proves source/build correct (see test.md).
- Fix (`app/desktop.py`): source-mode stale-dist guard `frontend_dist_freshness()` — newest
  frontend source input vs `dist/index.html` mtime; `main()` refuses to start with
  "Rebuild first: npm.cmd --prefix frontend run build" when stale, logs
  `WEBVIEW_FRONTEND_FRESHNESS` and `WEBVIEW_FRONTEND_ASSETS` (asset names + sha256 prefixes).
  Packaged mode exempt. No rows hidden/truncated, no backend/count change.

## Bug B — hidden "Bỏ qua" statuses visible
- Proven cause: same stale dist; backend DTO traced to exactly `outside_period` etc. via the
  single `_report_dto`/`_status_key` path (full + incremental).
- Hardening (`frontend/src/types.ts`): the ONE authoritative rule now compares the canonical
  status — exported `normalizeReportStatus()` (trim + lowercase) used by `isVisibleReport()`,
  which `presentReports()` (both ingestion paths, selection/active pruning unchanged) relies on.
  No second rule, no path blacklist, DTO/backend classification untouched, PPTX untouched.

## Bug C — Learning preview tiny (~80px)
- Proven cause: mount-once `useEffect([])` measured a null viewport ref because `ImageReview`
  early-returns with no candidate; async candidate arrival never re-measured (regression test
  measured exactly 80×45 pre-fix).
- Fix (`frontend/src/tabs/LearningTab.tsx`): lifecycle-safe callback ref `setViewportNode` —
  disconnects the previous observer, measures the new node immediately, observes resizes;
  unmount disconnects. No zoom-limit/dimension/renderer/geometry changes.

## Files changed
- `app/desktop.py`, `frontend/src/types.ts`, `frontend/src/tabs/LearningTab.tsx`,
  `frontend/package.json` (test script registration only — no version change)
- Tests: `frontend/tests/prompt038-integration.test.mjs`, `tests/test_prompt038_desktop_freshness.py`
- Spec Kit artifacts: this bug directory (`assessment.md`, `fix.md`, `test.md`)
- `frontend/dist` rebuilt locally for verification only; NOT committed (repo policy: gitignored).
