# PROMPT-038 verification report (Linux sandbox = development evidence, NOT Windows acceptance)

Executed from checked-in Spec Kit skill instructions; native Arena skill invocation was not available.

## Bug A — before/after (real Chromium, 1440×940, built dist, mocked bridge, 17 DTO reports)
- Windows (stale dist): ~10–12+ rows visible, no internal bound (user report).
- Fresh build served here: `.report-list-viewport` clientHeight 325px, computed max-height 327px
  (= 6×47 + 43 + 2), scrollHeight 654, exactly 6 visible rows, 13 rows in DOM (no truncation).
- Existing real-browser suite `test:report-list` (prompt034): 10/10 pass (0–6 natural height,
  >6 bounded, sticky header, 1000 rows accessible, no page growth).
- Stale-dist guard live check: touching `frontend/src/types.ts` then
  `python -m app.desktop --debug` logs `WEBVIEW_FRONTEND_FRESHNESS stale=True
  reason=source-newer-than-dist` and exits with the rebuild instruction; after
  `npm --prefix frontend run build` freshness reports `stale=False`.

## Bug B — before/after
- Same real-Chromium dist run: page text contains neither "Bỏ qua ngoài thời gian xử lý" nor any
  "Bỏ qua" status; 17 DTO reports (incl. all 4 hidden) → 13 visible rows.
- Existing `test:prompt037`: 20/20 pass (full + incremental reinsertion, selection/active
  pruning, totals, 1200-row bulk).
- New prompt038 B test: non-canonical spellings (`OUTSIDE_PERIOD`, ` fast_skip `, …) hidden by
  the single rule; visible statuses untouched; STT renumbered over visible rows.

## Bug C — before/after
- Pre-fix regression run: fit measured exactly 80×45 (the reported ~80px bug reproduced).
- Post-fix prompt038 tests (3/3 pass): async candidate arrival → 100% = 952×535.5 for a
  1000×700 viewport (16:9 preserved); 400% = 4× fit; resize to 500×400 → 452×254.25;
  candidate switch stays measured with new aspect; overlays remain at exact DTO percentages.

## Suites run
- `node --test tests/prompt038-integration.test.mjs`: 3/3 pass.
- `npm run test:report-list` (PROMPT-034): 10/10 pass.
- `npm run test:prompt037`: 20/20 pass.
- `npm run test:frontend` (all incl. Learning hooks/labels, prompt027/027r, 031r, update-state):
  71/72 pass — the 1 failure is PRE-EXISTING on the unmodified base (prompt027 test hardcodes
  app version 1.3.5 while package.json is 1.3.6; verified by stash + rerun). Not fixed here:
  version/build changes are forbidden by PROMPT-038.
- `pytest tests/test_prompt038_desktop_freshness.py tests/test_prompt028_portable.py`: 79 pass,
  3 failures PRE-EXISTING on the unmodified base (stash-verified; Windows/pywebview-probe and
  stale version-literal tests), 9 skipped.
- `npm run typecheck`: pass. `npm run build`: pass (assets `index-BoW2LrVw.js`,
  `index-ySgHXgSJ.css`; built CSS contains the `.report-list-viewport` cap).
- `git diff --check`: clean.

## Windows verification status
NOT performed — Linux sandbox only. Windows must re-run: `npm.cmd --prefix frontend run build`,
then `python -m app.desktop --debug`, confirm ~6 rows + internal scroll, no "Bỏ qua" rows, and
Learning 100%/400% fit; the new log lines `WEBVIEW_FRONTEND_FRESHNESS` /
`WEBVIEW_FRONTEND_ASSETS` record exactly which bundle was served.
