# Bug Assessment: “Danh sách báo cáo” grows the whole page with hundreds of reports

- **Slug**: prompt034r-report-list-viewport
- **Created**: 2026-10-13
- **Source**: pasted text (PROMPT-034R)
- **Verdict**: valid
- **Severity**: medium/high (core working list becomes unusable at the report volumes the app is built for;
  no data loss, no processing impact)
- **Starting SHA**: `2dbf355` (base branch `arena/2507b6dc-tnp`, includes PROMPT-036R + PROMPT-036S)
- **Working branch**: `arena/1597dcdd-tnp`
- **Application version**: unchanged — v1.3.6 Build019
- **Spec Kit execution**: Executed from checked-in Spec Kit skill instructions; native Arena skill invocation was not available.

## Report (summarized)

The “Danh sách báo cáo” table can contain hundreds or 1000+ reports. Today it grows as tall
as the dataset, so the page becomes inconvenient to use. Required UX: the list behaves like a
compact viewport — title/search/filter/action controls stay where they are, the column header
stays visible while the rows scroll, about six rows are visible at normal application size,
the rows scroll inside the list instead of stretching the page, a short list keeps its natural
height (no reserved blank area), no pagination/truncation, no horizontal scrollbar at supported
window sizes, selection/select-all/filters/search/status/active row/incremental updates/order
preserved, 1000+ rows usable, and no change to processing or backend scan semantics.

## Reproduction

Real engine, real application. Chromium 153 (the offline browser shipped in the checked-in dev
dependency `@sparticuz/chromium`) renders the app through Vite with only the pywebview bridge
mocked (`frontend/tests/layout/report-list-harness.tsx`, committed with the fix); 1440×940 is the
default pywebview window size (`app/desktop.py:534`). Measurements below are `getBoundingClientRect`
values from that page.

| Reports | list box `clientHeight` | list box `scrollHeight` | rows fully visible | workspace `main.scrollHeight` | document `scrollHeight` |
|--------:|------------------------:|------------------------:|-------------------:|------------------------------:|------------------------:|
| 0       | 298                     | 298                     | 0                  | 1021                          | 940 (viewport)          |
| 1       | 298                     | 298                     | 1                  | 1535                          | 1209                    |
| 6       | 325                     | 325                     | 6                  | 1535                          | 1209                    |
| 7       | 372                     | 372                     | 7                  | 1535                          | 1209                    |
| 30      | 1453                    | 1453                    | 30                 | 2265                          | 1940                    |
| 1000    | **47043**               | **47043**               | **1000**           | **47855**                     | **47530**               |

So one report row measured 47px and the column header 43px. Scroll behaviour before the fix:
with 1000 reports, scrolling the workspace moved the whole page — the report card header, the
search/filter/action row and the **column header** all scrolled away (rows 1–18 were shown at
`main.scrollTop = 300` with the table header above the fold). Nothing scrolled inside the list:
the list box was never shorter than its content in any scenario.

Also captured: `0 rows` reserved a 298px-tall empty list box inside a fixed `min-h-[430px]`
column, and `1 row` produced a 298px box around a single 47px row.

Evidence screenshots (captured in this session, kept outside Git):
`before-1-top.png` (one row inside a 300px blank box), `before-1000-top.png` (page runs off the
window), `before-1000-scrolled.png` (header, toolbar and column captions scrolled away).

## DOM / CSS inspection (the real structure)

- `frontend/src/tabs/ReportListTab.tsx:15` — the 65% workspace column:
  `<div className="flex min-h-[430px] flex-col">`. This fixed minimum is why even an empty or
  one-row list reserved a large blank area.
- `frontend/src/components/report/ReportTable.tsx:55` — the list box:
  `<div className="min-h-[300px] flex-1 overflow-auto rounded-sm2 border border-line">`.
  `min-h-[300px]` is the direct cause of the 298px render height for 0–1 rows, and
  `flex-1` (no `flex-basis: auto` clamp, no `max-height`) makes the box grow with its content —
  the table has no height constraint of its own.
- `frontend/src/components/report/ReportTable.tsx:159` — `<thead className="sticky top-0 z-10">`
  already existed, **but it was inert**: a sticky element only pins against its nearest scrolling
  ancestor, and the list box had nothing to scroll (`scrollHeight === clientHeight`), so the sticky
  header could never activate. No parent rule prevented it (`main` is `overflow-auto`, not
  `overflow-hidden`; no `transform`/`filter`/`contain` ancestors), so the missing ingredient was a
  bounded, actually-scrolling viewport, plus an opaque header background that survives
  `border-collapse: collapse` (the collapsed bottom border is painted by the table layer and stays
  behind while the header sticks).
- Row markup: 8 columns, `table-fixed` with an explicit `<colgroup>` (34/44/145/92/116/auto/150/26%),
  cells `px-2 py-1.5`, the last cell wrapping a fixed `h-[34px]` result box → 34 + 12 + 1 = 47px.
- Cost of a large dataset: with 1000 rows the workspace grew to ≈47.8k px, so every scroll,
  hover and layout pass ran over a 47k-px-tall table (no virtualization, no windowing, and the
  sticky header could not help). No per-row bridge traffic exists in the list today (only the
  750 ms dashboard poll plus one `get_report_details` for the active report, which the dashboard
  snapshot re-requests on each poll).

## Root cause (demonstrated, not assumed)

1. **The list box has no maximum height.** `min-h-[300px] flex-1` + an unbounded `<table>` means
   the box is always exactly as tall as all rows: 1000 rows → 47,043px inside the card.
2. **Because the box never scrolls, `overflow-auto` never engages and the already-present
   `sticky` `<thead>` can never stick** — the only scrolling ancestor left is `main`, which scrolls
   the whole page (controls, card header and column captions leave the viewport).
3. **The fixed `min-h-[430px]` column reserved blank space regardless of content**, which also
   breaks requirement 5 (≤6 rows must shrink naturally).

## Proposed remediation (chosen)

Bounded viewport on the existing list box, nothing else:

- Replace `min-h-[300px] flex-1` with a documented height cap ≈ six report rows + the column
  header + the box frame (`max-height: calc(var(--report-rows) * var(--report-row-h) +
  var(--report-head-h) + var(--report-frame))`, `overflow-y: auto`). Natural height below the cap
  is automatic — a block box with `max-height` shrinks to its content and scrolls only when the
  content exceeds the cap. No fixed empty box, no per-row markup change.
- Keep the existing `sticky top-0` header and give `thead th` an explicit background + inset
  bottom border so the header stays opaque over scrolling rows under `border-collapse: collapse`.
- Remove the fixed `min-h-[430px]` column minimum so an empty/short list is compact (the empty
  state itself keeps its own comfortable height).

Alternatives considered and rejected: pagination or result limiting (forbidden by the
requirements and by the “all reports remain available” rule); a JS `ResizeObserver`/row-count
height calculation (new observers and per-render work for a purely visual bound); virtualized
rows (a much larger change to selection/scroll semantics than the defect justifies).

## Files expected to change

- `frontend/src/index.css` (viewport component class)
- `frontend/src/components/report/ReportTable.tsx` (list box class)
- `frontend/src/tabs/ReportListTab.tsx` (remove the fixed 430px column minimum)
- `frontend/tests/layout/**`, `frontend/tests/prompt034r-report-list-layout.test.mjs`,
  `frontend/package.json` (real-browser regression coverage; no version change)

## Constraints confirmed before the fix

No backend, scanner, processing, updater, version/build, Portable or unrelated-file change is
needed: the defect is entirely in the report-list layout. Python tests are therefore not touched
(the layout fix changes no Python file).
