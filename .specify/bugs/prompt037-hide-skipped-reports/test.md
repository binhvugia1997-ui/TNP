# Bug Verification: Skipped reports hidden from “Danh sách báo cáo”

- **Slug**: prompt037-hide-skipped-reports
- **Tested**: 2026-10-13
- **Assessment**: ./assessment.md
- **Fix**: ./fix.md
- **Result**: partial

`partial` is used because the original symptom is gone and all PROMPT-037 checks pass, but two
pre-existing, unrelated version-metadata assertions fail (they expect `1.3.5`; the app is `1.3.6`),
and no Windows / pywebview acceptance was performed. Environment: Linux sandbox (Node 22, Python 3.11,
JSDOM + Vite SSR). Linux results are development evidence only, not Windows or Portable acceptance.

## Summary

The original defect reproduces on the base commit (skipped row visible, counted in the total, 15 of 16
regression tests red) and is gone after the fix (16/16 PROMPT-037 tests pass, visible count, selection and active-row
reconciliation proven through the real store and rendered `ReportTable`). No regression
in the frontend suite beyond the pre-existing version assertion; no regression in the scan / incremental /
application-service Python suites beyond the same pre-existing version assertion.

## Checks Performed

| Check | Command / Action | Result | Notes |
| --- | --- | --- | --- |
| Reproduction (pre-fix) | Probe: real store + mocked bridge, dashboard `[waiting, skipped, completed]` | fail (reproduced) | `visible ids: 0,1,2 stats.total: 3 skippedBucket: 1` |
| New tests on base `src/` (red check) | `git show HEAD:` for `store.tsx`/`types.ts`, `node --test tests/prompt037-hide-skipped.test.mjs`, then restored | 15 fail / 1 pass | The only pass is “skip classification semantics unchanged” (guard, expected to pass both ways) |
| Original case (post-fix) | Probe, same dashboard | pass | `visible ids: 0,2 stats.total: 2 skippedBucket: 0` |
| New frontend tests | `npm run test:prompt037` (`node --test tests/prompt037-hide-skipped.test.mjs`) | pass 16/16 | Full list below |
| Frontend suite | `npm run test:frontend` | 54 pass / 1 fail | Failure: `prompt027-overlays` “package metadata agrees with the authoritative backend version” (`'1.3.6' !== '1.3.5'`) — pre-existing on base (base run: 38 pass / 1 fail of 39) |
| Typecheck | `npm run typecheck` (= `npm.cmd run typecheck` on Windows; run here with `npm run`) | pass | `tsc --noEmit` exit 0 |
| Scan / incremental / application-service (Python) | `pytest tests/test_incremental_scan_cache.py tests/test_scan_performance.py tests/test_scan_list.py tests/test_scan_list_compact.py tests/test_application_service.py tests/test_prompt037_skipped_dto.py tests/test_scanner_visibility.py` | 67 pass / 1 fail | Fail: `test_application_service.py::test_configuration_ollama_update_state_and_disabled_install` — same `1.3.5` vs `1.3.6` version assertion, pre-existing and unrelated |
| Prompt037 backend guards | `pytest tests/test_prompt037_skipped_dto.py` | pass 3/3 | classification unchanged; DTO keeps `skipped`; no source file touched |
| py_compile | `python -m py_compile tests/test_prompt037_skipped_dto.py` | pass | No application Python file was changed |
| Whitespace / conflicts | `git diff --check` | pass (clean) | |
| Build / Portable / LAN publish / merge | not run | not-run | Out of scope per PROMPT-037 |
| Windows pywebview acceptance | not run | not-run | Linux sandbox only |

### PROMPT-037 test inventory (`frontend/tests/prompt037-hide-skipped.test.mjs`)

| Requirement | Test |
| --- | --- |
| waiting → visible; completed → visible; needs_review → visible; error → visible | visibility rule: waiting, completed, needs_review and error reports stay visible; full dashboard test |
| skipped → absent | visibility rule: a skipped report is absent… ; full dashboard test |
| mixed list → only skipped removed | mixed list test (STT renumbered, identity kept) |
| 1000+ mixed reports → correct visible count, no pathological behaviour | 1000+ mixed reports (pure, 12 000 rows, linear bound); 1000+ mixed reports render exact visible count (store, 1 200 rows, DOM count, ≤1 detail read) |
| skipped row cannot remain selected | a selected report that becomes skipped is removed from selectedIds |
| active row becoming skipped advances/clears | active row … advances to the next visible row; last → previous; empty clears; resolveActiveReportId unit test |
| select-all does not include skipped | select-all in the rendered table never includes skipped reports |
| dashboard refresh cannot reinsert skipped | dashboard refresh cannot reinsert…, several real polls observed |
| incremental scan update cannot reinsert skipped | incremental scan deltas cannot reinsert…; visible→skipped delta removes it; final scan response |
| rescan with same skipped classification stays hidden | a rescan with the same skipped classification stays hidden |
| no source PPTX deletion | hide flows record bridge calls; `mutationCalls` (anything other than ping/get/scan/details) must be empty in every hide test; backend `test_building_skipped_dto_rows_never_deletes_source_files` |
| skip classification semantics unchanged | pure guard (frontend `BUCKET_OF`/`STATUS_LABEL`); backend `test_skip_classification_is_unchanged_in_the_dto_mapping` |
| visible totals remain consistent | visible totals stay consistent with the rendered rows (`Tổng`, `N/M file`, buckets, row count) |
| reappearance after legitimate status change (no blacklist) | a skipped report that later becomes actionable again reappears |

## Output Excerpts

```
# tests 16
# pass 16
# fail 0
```

```
# tests 55
# pass 54
# fail 1
not ok 25 - the frontend package metadata agrees with the authoritative backend version
```

```
1 failed, 67 passed in 4.98s
FAILED tests/test_application_service.py::test_configuration_ollama_update_state_and_disabled_install
E       AssertionError: assert '1.3.6' == '1.3.5'
```

## Residual Risks

- Windows / pywebview runtime was not exercised; the store behaviour was verified with a mocked bridge in JSDOM.
- Real folder scans were not run in this sandbox (no Windows PowerPoint); the incremental path was verified with
  leaf-by-leaf dashboard snapshots shaped like `publish_leaf` output.
- Scope decision: only the `skipped` status literal is hidden. `outside_period`, `source_duplicate`, and `fast_skip`
  remain visible (their labels also start with “Bỏ qua”). The stat card “Bỏ qua” therefore counts only those.
- Two pre-existing version-metadata assertions (`1.3.5` vs `1.3.6`) fail on the base commit and were not modified.
- Renumbering STT over visible rows changes the displayed STT when hidden rows precede visible ones (intended).

## Recommendation

Hold for review — verified in the Linux development environment, partial overall because (a) Windows pywebview
acceptance with a real scan is still needed, and (b) the scope question for the other “Bỏ qua…” statuses is open.
The fix itself is complete for the `skipped` status: the hide rule is single, reconciled at the store ingestion point,
and covered by red-then-green regression tests. Do not merge without the user's explicit authorization.
