# Bug Verification: “Danh sách báo cáo” grows the whole page with hundreds of reports

- **Slug**: prompt034r-report-list-viewport
- **Tested**: 2026-10-13
- **Assessment**: ./assessment.md
- **Fix**: ./fix.md
- **Result**: verified (Linux/Chromium development evidence) with a Windows acceptance limitation

## Summary

The symptom was reproduced by measurement before the fix (a 1000-report list produced a 47,043px
list box and a 47,855px workspace; the column header and toolbar scrolled away with the page) and
the fixed build was measured in the same real browser: the list is capped at six rows plus the
header, scrolls inside itself, keeps the column header pinned (≤1px from the list top at every
scroll position, hit-testing as `TH`), keeps the page height independent of the dataset, and shows
the 1000th report at the bottom of the scroll. Ten new real-browser tests pass; the same suite fails
7/10 against the pre-fix source with geometry messages. No regression: the only remaining frontend
failure is the pre-existing version-metadata test that also fails on pristine HEAD.

## Checks Performed

| Check | Command / Action | Result | Notes |
|-------|------------------|--------|-------|
| Reproduction (pre-fix) | Chromium 153 (checked-in `@sparticuz/chromium`) rendering the real app via Vite harness, 1440×940 | pass (bug proven) | list box 298px @0 rows, 298px @1 row, 47,043px @1000 rows; workspace 47,855px; nothing scrolled inside the list; header+toolbar scrolled away |
| 0 rows: no oversized empty area | `node --test tests/prompt034r-report-list-layout.test.mjs` #1 | pass | viewport 209px instead of the pre-fix 298px; card 326px; no scroll |
| 1–6 rows: natural compact height | suite #2 | pass | 1 row 90px, 6 rows 325px = header 43 + 6×47 + 2px frame; no scrollbar; strictly increasing |
| >6 rows: bounded viewport + internal scroll | suite #3 | pass | cap 327px = 6×47 + 43 + 2 for 7/30/1000 rows; scrollHeight 372/1453/47043; ~6 rows visible |
| Sticky header | suite #4 | pass | header offset 1px (box border) at scrollTop 0/1/12/250/46718; hit-test `TH`; workspace/document scrollTop stay 0 |
| All rows kept / reachable | suite #3, #4, #5 | pass | 1000 rows rendered and scrollable to STT 1000 (`995…1000`); 1200-row fixture fully present |
| No pagination / truncation | suite #3 (rowCount === dataset size), #5 | pass | row count equals the fixture size in every scenario |
| Sticky/column structure | suite #5 | pass | 8 `<th>`, 8 `<colgroup><col>`, 8 `<td>` per row preserved |
| Search still works | suite #7 | pass | `260900012` → 1 row (STT 12), box collapses to 90px, no scrollbar |
| Filter still works | suite #7 | pass | `File lỗi` → 6 rows, all chips `Lỗi`, exactly fits with no scrollbar; clearing restores 30 |
| Selection / select-all | suite #8 | pass | select-all marks all 30 rows (not just the visible six), toggles off, single selection works, active outline exclusive |
| Incremental updates | suite #9 | pass | scan leaf 8→11 rows renders, viewport height unchanged; `processing`/`error` status updates render and never resize |
| No horizontal scrollbar | suite #6 | pass | widths 1440/1366/1280/1024/900: list/workspace/document horizontal overflow all 0 |
| Large fixture (1000+) usable | suite #3, #4, #5, #10 | pass | bounded box, reachable last row, 1200-row structural validity, traffic independent of row count |
| No per-row bridge traffic | suite #10 | pass | 1000 rows vs 6 rows over 3.2 s: same bounded traffic (`get_dashboard_state` 4–5, `get_report_details` ≤2/poll, 0 `ping`) |
| Pre-fix red check | same suite against stashed (pristine) source | pass (7 fail / 3 pass) | failures are geometric, e.g. `0 rows must not reserve a blank report area (viewport was 298px)`, `7 rows: the list is specified for six visible rows`, `a one-row search result must not keep the full list height (got 298px)` |
| Focused frontend suite | `npm run test:report-list` | pass | 10 passed / 0 failed |
| Existing frontend suite | `npm run test:frontend` | partial | 48 passed / 1 failed — `the frontend package metadata agrees with the authoritative backend version` (stale `1.3.5`/018 constant vs backend `1.3.6`/019), reproduced with the change stashed on pristine HEAD; out of scope (version bumps forbidden) |
| Typecheck | `npm run typecheck` (`tsc --noEmit`) | pass | clean |
| Production build | `npm run build` | pass | `index-S0xAL67Y.css` 28.45 kB, `index-DkKVnJ3Z.js` 282.52 kB |
| Diff hygiene | `git diff --check` | pass | clean |
| Python/backend tests | not-run | n/a | no Python file changed — a pure frontend layout fix must not trigger the long backend suite |

## Output Excerpts

```text
PRE-FIX (1440×940, real Chromium)
rows=  0  list clientHeight=  298 scrollHeight=  298 visible=  0  mainScrollHeight=  1021
rows=  1  list clientHeight=  298 scrollHeight=  298 visible=  1  mainScrollHeight=  1535
rows=  6  list clientHeight=  325 scrollHeight=  325 visible=  6  mainScrollHeight=  1535
rows=  7  list clientHeight=  372 scrollHeight=  372 visible=  7  mainScrollHeight=  1535
rows= 30  list clientHeight= 1453 scrollHeight= 1453 visible= 30  mainScrollHeight=  2265
rows=1000 list clientHeight=47043 scrollHeight=47043 visible=1000  mainScrollHeight= 47855
   main.scrollTop=300 → page scrolled: toolbar/column captions above the fold, rows 1…18 on screen

POST-FIX (same window)
rows=  0  cap=327  clientHeight=  209  (empty state)             mainScrollHeight= 1021
rows=  1  cap=327  clientHeight=   90  scrollHeight=   90        mainScrollHeight= 1535
rows=  6  cap=327  clientHeight=  325  scrollHeight=  325        mainScrollHeight= 1535
rows=  7  cap=327  clientHeight=  325  scrollHeight=  372  6 visible   mainScrollHeight= 1535
rows= 30  cap=327  clientHeight=  325  scrollHeight= 1453 6 visible   mainScrollHeight= 1535
rows=1000 cap=327  clientHeight=  325  scrollHeight=47043 6 visible   mainScrollHeight= 1535
   sticky header offset from list top: 1px at scrollTop 0/1/12/250/46718 (max), hit-test = TH
   bottom of scroll → visible window 995…1000, last row reachable, workspace scrollTop = 0
   horizontal overflow at 900/1024/1280/1366/1440 px: list 0, workspace 0, document 0

node --test tests/prompt034r-report-list-layout.test.mjs
# tests 10  # pass 10  # fail 0        (fixed source)
# tests 10  # pass  3  # fail 7        (pristine pre-fix source, geometry assertions)

npm run test:frontend → # tests 49  # pass 48  # fail 1  (pre-existing version constant, HEAD-identical)
npm run typecheck     → clean
npm run build         → ✓ built in 8.85s
git diff --check      → clean
```

Screenshot evidence captured in this session (kept outside Git, workspace folder
`prompt034r-evidence/`): `before-1-top.png` (single row inside a 300px blank box),
`before-1000-top.png`, `before-1000-scrolled.png` (page-level scroll, controls gone),
`after-0-top.png` (compact empty state), `after-1-top.png`, `after-6-top.png`
(natural six-row height), `after-1000-top.png` (bounded box, header pinned),
`after-1000-scrolled.png` (rows 13–18 under the pinned header, toolbar in place),
`zoom-sticky-header.png` (2× close-up of a row passing under the header).

## Residual Risks

- Linux headless Chromium is development evidence only. Windows acceptance (WebView2 wheel/scrollbar
  rendering, DPI scaling of the 47px/43px rows, pywebview at other window sizes) is **not performed**
  here and must not be reported as passed.
- The 47px/43px row and header heights were measured at the shipped font stack and 13px/19px body
  size; a Windows font fallback with different metrics would change the number of rows shown inside
  the cap, not the page-growth behaviour (the cap is fixed, so the list still scrolls). The regression
  test pins the cap formula and the rendered row height so a metric change surfaces as a test signal
  rather than silently.
- The new suite needs a usable Chromium; where none is available it skips with a printed reason
  (never a silent pass).

## Recommendation

Close the bug — verified end-to-end in a real engine with the original defect reproduced before the
fix and seven geometric assertions failing against the old source. Hold the Windows pywebview
acceptance (visible row count, scroll feel, DPI) for the Windows check, then merge the PR.
