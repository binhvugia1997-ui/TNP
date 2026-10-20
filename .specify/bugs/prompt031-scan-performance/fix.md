# Bug Fix: Quét scan response no longer performs quadratic source-path I/O

- **Slug**: prompt031-scan-performance
- **Fixed**: 2026-10-10
- **Assessment**: ./assessment.md
- **Status**: applied
- **Spec Kit execution**: Executed from checked-in Spec Kit skill instructions; native Arena skill invocation was not available.

## Confirmed Plan

- Replace per-string/per-row source-path replacement rebuilding with one response-scoped lexical redactor.
- Preserve recursive discovery, month/status semantics, ordering, DTO fields, opaque IDs, and path-redaction safety.
- Instrument the real sequential production path with bounded aggregate stages, exact slow-file diagnostics, work counts, and the pywebview return boundary.
- Add a deterministic no-COM benchmark/regression fixture that proves list correctness, termination, one traversal, one candidate metadata read, zero PPTX/COM work, and one redaction-context build.
- Do not add persistent caching, speculative concurrency, parser/COM changes, or a global timeout.

## Summary

The response builder now creates one filesystem-free, compiled source-path sanitizer per dashboard response and reuses it for all list rows/logs/messages. This removes the demonstrated `O(N²)` path-resolution/stat storm while keeping all list semantics and source-path redaction intact. The real Quét path now records bounded stage/count timing and emits an exact safe file index/basename/hash only when an operation remains slow enough to look hung.

## Changes

| File | Change | Notes |
|------|--------|-------|
| `app/scan_diagnostics.py` | added | One-call counters/timings, DEBUG per-file start/end, bounded top-5 slow operations, and an INFO watchdog snapshot only for slow/stuck stages/files. No timeout or execution concurrency. |
| `app/scanner.py` | modified | Optional observation-only metrics count entries, candidates, cheap unrelated filters, structural rejects, directories, and traversals without changing decisions/order. |
| `app/prescan.py` | modified | Optional sequential file metadata/status timing; exactly one candidate stat count; corrupt/stat failures keep existing isolation; PPTX and COM counters remain explicitly zero. |
| `app/gui_controller.py` | modified | Passes one scan metric record through enumerate/filter; logs `SCAN_ENUMERATE_*`, `SCAN_FILTER_*`; reuses one normalized path per `ScanRow`. |
| `app/application_service.py` | modified | Logs start/build/aggregate/total stages; response-scoped lexical redactor replaces per-field filesystem resolution; one normalized row key is reused; stores only the latest bounded metric record for diagnostics/tests. |
| `app/bridge.py` | modified | Marks the true pywebview-managed serialization boundary and emits `SCAN_RETURN ... bridge_calls=1`. It does not pretend a second `json.dumps` measures pywebview. |
| `tests/test_scan_performance.py` | added | Deterministic 610-entry realistic scan fixture plus list/work-count, redaction-I/O, error isolation, empty-folder, watchdog, month, parser, COM, and termination regressions. |
| `.specify/bugs/prompt031-scan-performance/assessment.md` | added | Pre-fix trace, reproducible profile, root-cause matrix, and remediation contract. |
| `.specify/bugs/prompt031-scan-performance/fix.md` | added | This remediation record. |

## Diff Highlights

- Old behavior: each report DTO called `_redact_source_paths()` twice; each call rebuilt variants for every source and called `Path.resolve()` twice per source.
- New behavior: `_source_path_redactor()` builds one lexical replacement map and one compiled pattern per response; `_redact_source_paths(value, redactor)` scans each string once and performs no source filesystem access.
- Normal INFO logging contains aggregate scan lines only. Every `SCAN_FILE_START/END` exists at DEBUG; if one operation exceeds the bounded stall threshold, the watchdog emits one INFO `SCAN_FILE_START ... status=still_running` snapshot for that exact stage/index, never an absolute path.
- The scan remains sequential. No production operation is cancelled, timed out, retried, cached, or parallelized.

## Tests Added or Updated

- `tests/test_scan_performance.py::test_large_quiet_scan_is_one_bounded_batch_with_identical_rescan`
  - 610 entries: 133 in-month corrupt synthetic PPTX candidates, 67 outside-month corrupt synthetic PPTX candidates, 400 unrelated files, and 10 structurally ineligible PowerPoint-looking files.
  - Pins 200 returned rows, ordering/status/month values, one traversal, 200 candidate metadata reads, 133 accepted files, 67 status skips, 10 structural rejects, 400 cheap unrelated filters, one redaction build, zero parser opens, and zero COM starts.
  - Repeats the scan and requires an identical list with one traversal again.
- `tests/test_scan_performance.py::test_redaction_context_is_lexical_and_reused_without_source_io`
  - Makes `Path.resolve()` fail if redaction touches it, proves source paths remain hidden, and reuses one context for 500 fields.
- `tests/test_scan_performance.py::test_stat_error_and_corrupt_pptx_are_isolated_and_scan_terminates`
  - A candidate `stat()` failure is isolated; both corrupt candidates remain listed and the scan terminates.
- `tests/test_scan_performance.py::test_empty_folder_returns_an_empty_scanned_list`
  - Pins correct empty-list semantics and zero per-file work.
- `tests/test_scan_performance.py::test_watchdog_identifies_exact_slow_file_without_absolute_path`
  - Pins one bounded slow-file line with exact index/phase/safe basename and no parent path.
- `tests/test_scan_performance.py::test_month_filter_contract_stays_filename_metadata_only`
  - Pins filename/date month decisions and confirms unrelated files are absent from candidate inspection.

## Local Verification

- `.venv/bin/pytest -q tests/test_scan_performance.py -vv` → **6 passed in 0.36s**.
- `.venv/bin/pytest -q tests/test_prescan.py tests/test_scanner_visibility.py tests/test_scan_list.py tests/test_application_service.py` → **82 passed in 10.95s**.
- `.venv/bin/python -m compileall -q app` → pass.
- `.venv/bin/pyflakes app/scan_diagnostics.py app/scanner.py app/prescan.py app/gui_controller.py app/application_service.py app/bridge.py` → pass.
- `git diff --check` → pass at this stage.

### Same-fixture before/after development benchmark

The benchmark runs used the same 610-entry fixture and `cProfile`. Timing is supporting Linux evidence; work counts are the regression gate.

| Metric | Before | After |
|---|---:|---:|
| directory entries | 610 | 610 |
| candidate PowerPoint files | 200 | 200 |
| accepted files | 133 | 133 |
| directory traversals | 1 | 1 |
| PPTX opens/parses | 0 | 0 |
| candidate metadata reads | 200 | 200 |
| PowerPoint COM starts | 0 | 0 |
| bridge-equivalent service calls | 1 | 1 |
| redaction-context builds | 425 full rebuilds (`_redact_source_paths` calls) | 1 |
| `Path.resolve()` calls (profile) | 253,134 | 734 |
| filesystem `lstat` + `stat` calls (profile) | 1,265,875 | 3,875 |
| scan elapsed under `cProfile` | 17.5278 s | 0.2461 s |
| largest identified stage | response redaction, 17.3527 s | response build/redactor compile, 0.1752 s |

Supporting scaling changed from clearly quadratic (50/100/200 candidates: 1.3185/4.7306/17.5278 s) to approximately linear (0.0643/0.1231/0.2461 s). At 200 candidates, profiled Python calls fell from 32,165,471 to 592,547. No wall-clock limit is asserted in CI.

## Deviations from Assessment

None. The assessment anticipated a small dedicated diagnostics module; that kept timing/watchdog behavior isolated from scan decisions. The literal Windows never-return case remains a required manual retest, so the eventual verification verdict must remain partial even though the demonstrated Linux backend cause and regression are fixed.

## Follow-ups

- Run the same real Windows flow and capture the exact `SCAN_*` lines listed in `test.md`/the final report.
- If Windows stalls, use `SCAN_STALLED` or slow `SCAN_FILE_START ... status=still_running` to reassess the exact operation; do not add a global timeout.
- No Portable rebuild/publish is required merely to assess this source change.
