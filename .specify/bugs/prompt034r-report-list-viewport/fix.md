# Bug Fix: “Danh sách báo cáo” grows the whole page with hundreds of reports

- **Slug**: prompt034r-report-list-viewport
- **Fixed**: 2026-10-13
- **Assessment**: ./assessment.md
- **Status**: applied

## Summary

The report list is now a bounded viewport: it shows about six report rows plus the sticky column
header and scrolls inside itself, while a shorter list keeps its natural height. The already-present
sticky `<thead>` is now effective because its scrolling ancestor actually scrolls, and it was made
opaque for `border-collapse: collapse` rendering. The fixed `min-h-[430px]` column placeholder that
reserved blank space was removed. No pagination, no truncation, no dataset or backend change, no
per-row DOM/timer/observer/bridge work: three class-level edits plus one CSS component rule.

## Changes

| File | Change | Notes |
|------|--------|-------|
| `frontend/src/index.css` | modified | New `.report-list-viewport` component rule: documented `--report-rows: 6`, `--report-row-h: 47px`, `--report-head-h: 43px`, `--report-frame: 2px`, `max-height: calc(rows*rowH + headH + frame)`, `overflow-y: auto`. Plus `.report-list-viewport thead th { background:#f4f7fb; box-shadow: inset 0 -1px 0 #d8e1ec }` so the sticky header stays opaque and keeps its bottom edge over scrolling rows (a collapsed table border would stay behind). |
| `frontend/src/components/report/ReportTable.tsx` | modified | List box `min-h-[300px] flex-1 overflow-auto …` → `report-list-viewport overflow-auto …`. `min-height` removed (natural height below the cap, nothing reserved); the existing `sticky top-0` `<thead>`, colgroup, rows, selection, chips and result cell are untouched. |
| `frontend/src/tabs/ReportListTab.tsx` | modified | Workspace column `flex min-h-[430px] flex-col` → `flex flex-col` (the fixed 430px blank minimum is gone). |
| `frontend/package.json` | modified | New `test:report-list` script; the new suite is part of `test:frontend`. Version/build unchanged (1.3.6 / Build019). |
| `frontend/tests/layout/report-list-harness.html`, `frontend/tests/layout/report-list-harness.tsx` | added test | Real-browser harness: the REAL `App` + `StoreProvider` + `ReportTable` rendered through Vite, only `window.pywebview.api` mocked (Proxy answering every method, counting native calls). `?rows=N` builds a 0…5000-row dataset; `window.__tnpLayoutHarness` can grow the dataset incrementally (modelling one completed scan leaf), flip a row's status and read the bridge-call tally. |
| `frontend/tests/prompt034r-report-list-layout.test.mjs` | added test | 10 real-Chromium geometry/behaviour tests (see below). Locates the list box structurally (`.report-list-viewport` **or** the pre-fix `div.overflow-auto`) so the same assertions fail on geometry against the old markup instead of on a missing class. |

## Diff Highlights

```css
.report-list-viewport {
  --report-rows: 6; --report-row-h: 47px; --report-head-h: 43px; --report-frame: 2px;
  max-height: calc(var(--report-rows) * var(--report-row-h) + var(--report-head-h) + var(--report-frame));
  overflow-y: auto;
}
```

- The cap is derived from the measured row (34px result cell + 2×6px cell padding + 1px separator)
  and header (two-line uppercase caption) heights, not from a magic fixed box — verified by the test
  against the rendered row height.
- `max-height` (not `height`) is what gives requirement 5 for free: 0–6 rows render at their natural
  height, >6 rows pin at the cap and scroll.

## Tests Added or Updated

`frontend/tests/prompt034r-report-list-layout.test.mjs` (node:test, real Chromium through the
checked-in `@sparticuz/chromium` + `puppeteer`; skips with a printed reason when no browser is
usable so a measurement is never reported as passing):

1. **0 rows** — no oversized empty report area (`< 260px`, card `< 350px`; pre-fix 298px), no scroll,
   no horizontal overflow.
2. **1/2/3/6 rows** — natural compact height equals header + rows + the 2px frame (≤2px), stays under
   the cap, no internal scrollbar, every row visible, height strictly increasing with row count,
   1 row `< 150px`.
3. **7/30/1000 rows** — the cap equals the documented 6-rows-plus-header rule, the rendered box fills
   it, rows scroll inside the list, ~6 rows visible, `<thead>` computed `position: sticky`, all rows
   still present, no horizontal overflow; workspace/document heights identical for 7 and 1000 rows.
4. **Sticky header + reachability (1000 rows)** — header offset from the list top stays ≤1px at
   scrollTop 0/1/12/250/max, header fully inside the box, hit-test at the header centre returns `TH`
   (header above the rows), workspace and document `scrollTop` stay 0, exactly six rows visible, the
   last row (STT 1000) reachable with visible window `995…1000` and a monotonic forward window.
5. **Large fixture (1200 rows)** — all 1200 rows rendered, 8 columns/8 colgroup entries/8 cells per
   row, scroll range > 100 viewports, 6 visible, sticky, no horizontal overflow.
6. **Desktop widths 1440/1366/1280/1024/900** — zero horizontal overflow on the list, workspace and
   document; rows fit the column width; still ~6 visible rows.
7. **Search + status filter** — one-row search result collapses the box (`< 150px`, no scrollbar); a
   no-match search keeps an empty state (`cardHeight < 350`); the status filter narrows to the six
   error reports with correct status chips and no scrollbar; clearing restores all 30.
8. **Selection / select-all** — select-all marks all 30 reports (not only the six visible), toggles
   off, single-row selection works, active-row highlight stays exclusive.
9. **Incremental scan updates** — appending a completed scan leaf grows the dataset to 11 rows without
   changing the viewport height, new rows render, a `processing` status renders per row while other
   rows keep their status, an `error` status update does not resize the list.
10. **No per-row native bridge traffic** — with 1000 rows, a 3.2 s window costs the same traffic as
    6 rows (detail fetches bounded by the dashboard poll cadence, no reconnects, total `< 40`).

## Local Verification

- `node --test tests/prompt034r-report-list-layout.test.mjs`: **10 passed / 0 failed** (real Chromium).
- Same suite against the pre-fix source (stash): **7 failed / 3 passed** with geometry messages
  (`0 rows must not reserve a blank report area (viewport was 298px)`, `1 rows must keep natural
  height (got 298px, expected 91.5px)`, `7 rows: the list is specified for six visible rows`,
  `about six rows stay visible while scrolling — 1000 !== 6`, `the 1200 rows live behind the bounded
  viewport scroll`, `a one-row search result must not keep the full list height (got 298px)`,
  `eight rows already scroll`). The three that pass pre-fix are the “must not regress” assertions
  (horizontal overflow, selection/select-all, bridge traffic).
- `npm run test:frontend`: 48 passed / 1 failed — the single failure is the **pre-existing**
  `the frontend package metadata agrees with the authoritative backend version` (test constant
  `1.3.5`/Build 018 vs backend `1.3.6`/Build019), reproduced on pristine HEAD with the change
  stashed; it is unrelated to this layout fix and version bumps are forbidden by this prompt.
- `npm run typecheck` (`tsc --noEmit`): pass. `npm run build`: pass
  (`index-S0xAL67Y.css` 28.45 kB, `index-DkKVnJ3Z.js` 282.52 kB).
- `git diff --check`: clean.
- Python tests: **not run — no Python file changed** (pure frontend layout fix).

## Deviations from Assessment

None. The remediation stayed with the chosen bounded-viewport approach; no pagination, no
virtualization, no dataset/back-end change, and no semantics change to selection, filters, search,
status or incremental updates.

## Follow-ups

- Windows/pywebview acceptance of the real window (1440×940 default, WebView2 scroll rendering,
  mouse-wheel over the list, sticky header under DPI scaling) remains for the Windows check —
  Linux Chromium numbers are development evidence only (constitution VII).
