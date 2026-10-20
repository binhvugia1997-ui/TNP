# Bug Verification: Quét scan response performance and hang diagnostics

- **Slug**: prompt031-scan-performance
- **Tested**: 2026-10-10
- **Assessment**: ./assessment.md
- **Fix**: ./fix.md
- **Result**: partial
- **Environment**: Linux agent, Python 3.11.2; no real Windows pywebview/PowerPoint/UNC/removable-drive acceptance
- **Spec Kit execution**: Executed from checked-in Spec Kit skill instructions; native Arena skill invocation was not available.

## Summary

The deterministic Linux reproduction no longer exhibits the quadratic response-building path: the same 610-entry fixture returns the same 200-row ordered/status-preserving list, response redaction is built once without source filesystem I/O, and the profiled runtime falls from 17.5278 s to 0.2461 s. Work-count regressions prove one traversal, one metadata read per candidate, no PPTX open, no PowerPoint COM start, and termination despite corrupt bytes or one metadata error.

The verdict is **partial**, not verified, because the strengthened original symptom is a real Windows “list never appears” report and that exact environment has not been rerun. The new bounded production instrumentation is ready to distinguish a blocked file/stage, backend completion, and the pywebview return boundary on that retest.

## Checks Performed

| Check | Command / Action | Result | Notes |
|-------|------------------|--------|-------|
| Reproduction (pre-fix) | Real `ApplicationService.scan_reports()` + `cProfile`, 610-entry fixture | fail before fix | 17.5278 s; 17.3527 s in redaction; 253,134 resolves; 1,265,875 filesystem metadata syscalls. |
| Reproduction (post-fix) | Same production service path + same fixture + `cProfile` | pass | 0.2461 s; same 200 rows; one redactor context; approximately linear scaling. |
| New focused regressions | `.venv/bin/pytest -q tests/test_scan_performance.py` | pass | 6 passed in 0.42s. |
| Final focused rerun | `.venv/bin/pytest -q tests/test_scan_performance.py tests/test_application_service.py` | pass | 27 passed in 2.09s after final diagnostic/redaction refinements. |
| Scan/backend regression | `.venv/bin/pytest -q tests/test_prescan.py tests/test_scanner_visibility.py tests/test_scan_list.py tests/test_scan_list_compact.py tests/test_application_service.py tests/test_gui_controller.py` | pass | 109 passed in 16.42s. |
| PROMPT-030 regression | `.venv/bin/pytest -q tests/test_prompt030_runtime.py` | pass | 99 passed, 5 skipped in 0.75s. Skips are host/platform-specific. |
| Full Linux Python suite (run once) | `.venv/bin/pytest -q` | environment-incomplete | 1197 passed, 15 skipped, 2 failed in 108.04s. Both failures reported absent pywebview/pythonnet packages in the fresh venv, not code assertions in changed paths. |
| Failed-node rerun after installing declared webview runtime plus Linux import-probe packages | `.venv/bin/pytest -q tests/test_prompt028_portable.py::test_official_hook_webview_misses_the_js_bridge_so_our_manual_js_collection_is_necessary tests/test_prompt028_portable.py::test_live_import_probe_reports_every_module_in_this_venv` | pass | 2 passed in 0.65s after installing `requirements-webview.txt`, `pythonnet==3.0.5`, and `clr-loader==0.2.10`. Per instruction, the full suite was not run a second time. |
| Compile | `.venv/bin/python -m compileall -q app tests tools` | pass | No syntax/bytecode compilation error. |
| Changed-path lint | `.venv/bin/pyflakes app/scan_diagnostics.py app/scanner.py app/prescan.py app/gui_controller.py app/application_service.py app/bridge.py tests/test_scan_performance.py` | pass | No warning in changed Python paths. |
| Repository-wide pyflakes observation | `.venv/bin/pyflakes app tests/test_scan_performance.py` | known baseline warnings | Nine existing intentional optional/import-probe unused-import warnings in unchanged files; no changed-path warning. |
| Whitespace | `git diff --check` | pass | Clean. |
| Frontend tests/typecheck | not run | not applicable | No frontend source changed. |
| Real Windows Quét flow | maintainer action pending | not-run | Required before a verified verdict. |
| Real PowerPoint COM | not-run | not applicable to list discovery / pending Windows | Regression proves Quét does not call COM. |

## Output Excerpts

```text
$ .venv/bin/pytest -q tests/test_scan_performance.py
......                                                                   [100%]
6 passed in 0.42s
```

```text
$ .venv/bin/pytest -q tests/test_prescan.py tests/test_scanner_visibility.py tests/test_scan_list.py tests/test_scan_list_compact.py tests/test_application_service.py tests/test_gui_controller.py
........................................................................ [ 66%]
.....................................                                    [100%]
109 passed in 16.42s
```

```text
$ .venv/bin/pytest -q tests/test_prompt030_runtime.py
.......sss..................................s........................... [ 69%]
...........................s....                                         [100%]
99 passed, 5 skipped in 0.75s
```

```text
$ .venv/bin/pytest -q
2 failed, 1197 passed, 15 skipped in 108.04s (0:01:48)
```

The two full-run failures were:

```text
collect_data_files("webview", subdir="js") == []
ModuleNotFoundError: No module named 'webview'
ModuleNotFoundError: No module named 'pythonnet'
```

After installing the runtime packages, only those exact failed nodes were rerun:

```text
..                                                                       [100%]
2 passed in 0.65s
```

## Before/After Performance Evidence

Same 610-entry fixture; cProfile timing is supporting Linux evidence, while operation counts are the stable gate.

| Metric | Before | After |
|---|---:|---:|
| directory entries | 610 | 610 |
| candidate PowerPoint files | 200 | 200 |
| accepted files | 133 | 133 |
| structural PowerPoint rejects | 10 | 10 |
| unrelated cheap rejects | 400 | 400 |
| outside-month/status skips | 67 | 67 |
| directory traversals | 1 | 1 |
| PPTX opens/parses | 0 | 0 |
| candidate metadata reads | 200 | 200 |
| PowerPoint COM starts | 0 | 0 |
| bridge calls | 1 | 1 |
| full redaction/replacement builds | 425 calls rebuilding all sources | 1 response context |
| `Path.resolve()` calls | 253,134 | 734 |
| profiled `lstat` + `stat` calls | 1,265,875 | 3,875 |
| scan elapsed | 17.5278 s | 0.2461 s |
| largest identified stage | response redaction: 17.3527 s | response build/redactor compile: 0.1752 s |

Additional scaling evidence:

| Candidate files | Before | After |
|---:|---:|---:|
| 50 | 1.3185 s | 0.0643 s |
| 100 | 4.7306 s | 0.1231 s |
| 200 | 17.5278 s | 0.2461 s |

The hard regression does **not** assert these times. It pins the work contract: 610 entries, 200 candidates, one traversal, 200 candidate stats, one redaction context, 0 parser opens, 0 COM starts, and an identical second-scan snapshot.

## Correctness and Termination Findings

- Same folder scanned twice returns the same IDs, order, names, Management Numbers, occurrence dates, statuses, and blank external `path` fields.
- All 133 September rows remain processing candidates; all 67 October rows remain visible with `outside_period` status.
- All synthetic PPTX bytes are corrupt/non-ZIP. They are listed without being opened, proving corrupt contents cannot freeze discovery.
- 400 unrelated files and 10 structurally ineligible PowerPoint-looking files receive no candidate metadata inspection.
- A synthetic `Path.stat()` error is isolated; both candidate rows remain present and the scan terminates.
- Empty folder returns `scanned=true`, an empty report list, one traversal, and zero per-file operations.
- Normal per-file start/end diagnostics are DEBUG-only. A slow/stuck operation emits one bounded INFO snapshot with exact index/phase and safe basename/hash, no absolute parent path.
- The backend remains synchronous within one pywebview call, but the demonstrated work is now bounded/linear. No spinner-only workaround was used.

## Windows Retest Procedure and Required Log Capture

No Portable publish/rebuild is required merely to assess this source change. On integrated Windows source/build, run exactly:

1. Chọn tháng.
2. Chọn thư mục used by the real report set.
3. Bấm **Quét** once.
4. Wait until the list appears, or until a slow/stall diagnostic appears.
5. Copy from `logs/app.log` the contiguous block beginning at `SCAN_START` and ending at `SCAN_RETURN` (or the last available `SCAN_*` line if it still hangs).

Capture these exact lines/markers:

```text
SCAN_START
SCAN_ENUMERATE_START
SCAN_ENUMERATE_END elapsed_ms=... entries=... candidates=... rejected=... unrelated=... traversals=...
SCAN_FILTER_START
SCAN_FILE_START index=... phase=... file=... elapsed_ms=... status=still_running   # only if a file is slow/stuck
SCAN_FILE_END index=... phase=... file=... elapsed_ms=...                       # DEBUG diagnostics if DEBUG capture is enabled
SCAN_FILTER_END elapsed_ms=... candidates=... accepted=... skipped=...
SCAN_BUILD_RESPONSE_START
SCAN_STALLED stage=... elapsed_ms=... status=still_running                       # only if a non-file stage stalls
SCAN_BUILD_RESPONSE_END elapsed_ms=... files=... redaction_contexts=... redaction_sources=...
SCAN_PER_FILE elapsed_ms=... operations=... sequential=true
SCAN_METADATA elapsed_ms=... reads=... errors=... master_open_ms=...
SCAN_PARSE_PPTX elapsed_ms=... opens=...
SCAN_POWERPOINT_COM elapsed_ms=... starts=...
SCAN_SLOW_OP rank=... phase=... index=... file=... elapsed_ms=...
SCAN_TOTAL elapsed_ms=... entries=... candidates=... accepted=... rejected=... unrelated=... status_skipped=... returned=... traversals=... largest_measured_stage=...
SCAN_SERIALIZE elapsed_ms=unavailable boundary=pywebview_after_python_return
SCAN_RETURN elapsed_ms=... ok=True bridge_calls=1
```

Interpretation:

- Last line inside enumeration → filesystem walk/I/O issue.
- Slow `SCAN_FILE_START ... status=still_running` → exact sequential candidate index/phase is blocked.
- `SCAN_BUILD_RESPONSE_START` without END → response construction is blocked.
- `SCAN_TOTAL` present but `SCAN_RETURN` absent → failure is between service completion and bridge return.
- `SCAN_RETURN` present but UI list absent → backend returned; investigate pywebview delivery/frontend rendering separately.
- `opens` and COM `starts` must remain 0 for Quét.

Also record whether the selected folder is local disk, UNC/network, or removable media. Do not send sensitive absolute report paths; the new diagnostics do not log them.

## Residual Risks

- The literal Windows never-return symptom is not reproduced/accepted in the Linux agent.
- Actual pywebview serialization happens after the Python method returns and cannot be timed from this source boundary; it is marked honestly as unavailable. `SCAN_RETURN` identifies the boundary.
- Very large Excel masters may make `master_open_ms` or `SCAN_FILTER` the largest remaining legitimate stage. Instrumentation will demonstrate that rather than masking it.
- Full-suite aggregate remained formally non-green because its single allowed run preceded installation of optional runtime probe packages. Both exact environmental failures pass after installation; no second full run was performed.
- No real Windows COM or Portable acceptance is claimed.

## Recommendation

Proceed to the required Windows Quét retest and capture the complete `SCAN_*` block. The demonstrated quadratic backend bug is fixed and regression-tested, and production diagnostics are sufficient for the stronger hang report. Do not label the issue fully verified until the same real Windows folder returns its file list and the captured counts/stages are reviewed.
