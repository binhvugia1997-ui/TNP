# PROMPT-038 assessment — integration acceptance regressions (A/B/C)

Slug: `prompt038-integration-regressions` (generated, automated mode).
Base: `integration/prompts-034-035-036-037` @ 53818c2, v1.3.6 Build019.
Executed from checked-in Spec Kit skill instructions; native Arena skill invocation was not available.
Environment: Linux sandbox = development evidence only, NOT Windows acceptance.

## Evidence captured before any fix

### Common A/B evidence — fresh build vs Windows observation
- `frontend/dist` is **untracked and gitignored** (`.gitignore:11 dist/`, `git ls-files frontend/dist` = 0).
  Merging PROMPT-034/035/036/037 therefore updates `frontend/src` only; whatever `dist` happens to be
  on the Windows disk is silently served by `python -m app.desktop` with no freshness/hash check,
  and the startup log did not record which asset hashes were served.
- Fresh `npm --prefix frontend run build` in this sandbox produced `assets/index-ETqkDjvt.js` +
  `assets/index-ySgHXgSJ.css`; the built CSS **does** contain
  `.report-list-viewport{...max-height:calc(6*47px + 43px + 2px);overflow-y:auto}` (hypothesis A2/A3
  disproved for a fresh build).
- Real-Chromium run against the **built dist** (static server, mocked pywebview bridge, 1440×940,
  17 DTO reports incl. all four hidden statuses): viewport `clientHeight=325`, `max-height=327px`,
  `scrollHeight=654`, **exactly 6 visible rows**, 13 rows in DOM, and the page text contains neither
  "Bỏ qua ngoài thời gian xử lý" nor any "Bỏ qua" status. → Source + fresh build are correct for
  both A and B; the Windows runtime must have served a **stale pre-merge dist** (hypothesis 1 for
  both A and B).

### Bug A — hypotheses
1. stale dist: **PROVEN as the only consistent cause** (see above; no other hypothesis reproduces).
2. built CSS missing rule: disproved (rule present in built CSS).
3. Tailwind layer/order override: disproved (computed `max-height: 327px` in real Chromium).
4. runtime DOM lacks class: disproved for fresh build (`.report-list-viewport` present, cap applied).
5. other layout override: disproved (measured 6 visible rows, internal scroll).
6. pywebview cache: **disproved** — pywebview 6.2.1 `webview/http.py` sends
   `Cache-Control: no-cache, no-store, must-revalidate` (+ Pragma/Expires) for every static asset.
7. merge/build regression: disproved (current source + fresh build measure correctly).

Root cause A: unchecked stale `frontend/dist` on disk served in source mode.

### Bug B — hypotheses
- REAL DTO status traced in source: `app/application_service.py::_status_key` maps
  `ACTION_OUTSIDE_PERIOD → "outside_period"` (also via `final_map` for run_stage), and **all** DTO
  paths (full + incremental `_scan_pending_reports`) go through the single `_report_dto`. Prescan
  labels ("Bỏ qua ngoài thời gian xử lý") are display text, never the DTO status; the DTO value is
  exactly `outside_period`. So hypothesis 2 (DTO differs) is disproved for the checked-in backend.
  Hypotheses 4/5/6 (bypass paths) disproved by inspection + prompt037 tests: both ingestion paths
  go through `presentReports`/`mergeIncrementalReports`, selection/active pruned in `acceptDashboard`.
- Hypothesis 1 (stale frontend assets) is the proven common cause: with stale pre-037 assets the
  row renders normally.
- Residual hardening (defence in depth, keeps ONE rule): the visibility predicate compared the raw
  DTO string; any future non-canonical casing/whitespace would silently leak a "Bỏ qua" row.
  Normalising inside the single authoritative predicate closes that without a second rule.

Root cause B: stale dist (primary); exact-match predicate fragility (secondary hardening).

### Bug C — Learning tiny preview
Confirmed by code inspection (regression test written first, fails pre-fix):
`ImageReview` early-returns while no candidate exists, so the viewport `<div ref={viewportRef}>`
is not in the DOM on first render; the `useEffect(..., [])` runs once, sees `viewportRef.current
=== null`, and never re-runs when the async candidate arrives. `viewport` stays `{w:0,h:0}`,
`Math.max(80, …)` yields ~80px fit, 400% ≈ 320px.

## Proposed remediation
- A+B: source-mode stale-dist guard in `app/desktop.py` — compare newest `frontend/src` (+
  `index.html`, vite/tailwind/postcss configs) mtime against `frontend/dist/index.html`; refuse to
  start with a clear rebuild instruction when stale, and log served asset names + sha256 prefixes
  (`WEBVIEW_FRONTEND_ASSETS`) so runtime evidence is capturable. Packaged mode exempt (bundle is
  proven by the portable builder). No version bump, no dist committed.
- B (extra): normalise status (trim/lowercase) inside the ONE authoritative visibility rule
  (`isVisibleReport`/`normalizeReportStatus`); no other list source, no path blacklist, DTO data
  untouched, hidden rows still available to backend/diagnostics.
- C: lifecycle-safe callback ref for the preview viewport — attaches when the node appears,
  disconnects from the old node, measures immediately, observes resizes; unmount disconnects.
- Tests: new `frontend/tests/prompt038-integration.test.mjs` (B normalisation + C lifecycle/zoom/
  resize/switch/aspect/overlay alignment), new `tests/test_prompt038_desktop_freshness.py`
  (fresh/stale/missing/no-source-tree/packaged-exempt + asset report). Existing 034/037/Learning
  suites re-run; typecheck + build; `git diff --check`.

## Risks
- Freshness guard must never block the packaged exe or a source checkout without `frontend/src`
  adjacency — guarded by `is_packaged()` and "no source tree → cannot judge → allow".
- mtime comparison only (no content diffing) — a rebuilt dist is always newer than sources.
