# Bug Assessment: Quét scan appears to hang before the file list is returned

- **Slug**: prompt031-scan-performance
- **Created**: 2026-10-10
- **Source**: pasted text (PROMPT-031 plus strengthened Windows reproduction evidence)
- **Verdict**: valid
- **Severity**: high
- **Starting SHA**: `f0401d04e725ec2d3324ef911103fea1d462fa01`
- **Working branch**: `arena/2e088bee-tnp`
- **Spec Kit execution**: Executed from checked-in Spec Kit skill instructions; native Arena skill invocation was not available.

## Report (verbatim or summarized)

The user-visible flow is:

1. Chọn tháng.
2. Chọn thư mục.
3. Bấm **Quét**.
4. The scan appears to hang indefinitely on Windows and the file list does not appear even after a long wait.

Folder/month selection itself is not reported slow. The request requires the real Quét call path to be profiled before a fix, bounded stage/file diagnostics, a deterministic Linux benchmark that does not require PowerPoint COM, work-count regressions, and a later real-Windows retest.

## Preflight and source authority

- Read `AGENTS.md`, `.specify/memory/constitution.md`, `docs/SPECKIT_ARENA.md`, and all three checked-in bug skill files before assessment.
- Ran `git fetch origin --prune`, `git status`, `git branch --show-current`, and `git log --oneline --decorate -15` before assessment.
- Working tree was clean on `arena/2e088bee-tnp` at exact expected SHA `f0401d0` (`Merge PROMPT-030T into PR #7`).
- `git ls-remote --heads origin arena/28cc318f-tnp` resolves to the same full SHA. The checkout is not stale and did not require synchronization.
- `tests/test_prompt030_runtime.py` and the committed PROMPT-030T Spec Kit reports are present. The source contains the PR #7/PROMPT-030T work.
- `.specify/extensions.yml` is valid and has `hooks: {}`; there were no bug-assessment hooks to execute.
- No production file was edited during assessment.

## Reproduction

### Reported real flow

1. On Windows, select a month/year.
2. Select the report folder.
3. Press **Quét**.
4. Observe that the list does not appear for a very long time / appears indefinitely hung.

The literal never-returning Windows case is not executable in the Linux agent. The production backend delay was nevertheless reproduced deterministically without PowerPoint, COM, PPTX parsing, Ollama, preview rendering, or export.

### Deterministic pre-fix benchmark

An ad-hoc assessment benchmark drove the real `ApplicationService.scan_reports()` production path with a temporary folder containing:

- 133 September `.pptx` candidates accepted by the selected September 2026 period;
- 67 October `.pptx` candidates retained in the list with `OUTSIDE_PERIOD` status;
- 400 unrelated `.txt` files filtered structurally;
- 5 `~$*.pptx` lock/temp files and 5 `*.pptx.tmp` files rejected structurally;
- 610 inspected directory entries total;
- invalid synthetic PPTX bytes, deliberately proving scan discovery does not open their ZIP/PPTX contents;
- no template/output workbook, so Excel workbook loading could not confound the measured response-building path.

The benchmark used one service/bridge-equivalent batch operation and a selected `month=9, year=2026`. `cProfile`, operation counters, and bounded wrappers measured the production function boundaries. It did not modify source.

### Pre-fix measurements

| Candidate PowerPoint files | Directory entries | Returned reports | Total | Discover | Pre-scan/filter | Build dashboard response | Path redaction |
|---:|---:|---:|---:|---:|---:|---:|---:|
| 50 | 160 | 50 | 1.3185 s | 0.0083 s | 0.0026 s | 1.3035 s | 1.2747 s |
| 100 | 310 | 100 | 4.7306 s | 0.0149 s | 0.0053 s | 4.7037 s | 4.6449 s |
| 200 | 610 | 200 | 17.5278 s | 0.0294 s | 0.0103 s | 17.4762 s | 17.3527 s |

The 200-file profile recorded:

- one directory traversal;
- 200 candidate files and 200 returned list DTOs;
- 133 month-accepted processing candidates and 67 month-skipped/outside-period rows;
- 0 PPT/PPTX opens/parses;
- 0 PowerPoint COM starts;
- 0 subprocess waits;
- 0 network/Ollama calls;
- 200 legitimate per-candidate `stat()` metadata inspections in `prescan`;
- 425 calls to `ApplicationService._redact_source_paths()` while building one response;
- 253,134 `Path.resolve()` calls caused chiefly by response redaction;
- 1,012,535 `lstat` and 253,340 `stat` calls in the profile (1,265,875 filesystem metadata syscalls total);
- 32,165,471 Python calls;
- one batched backend call, not one bridge call per file.

The measured scaling (1.32 s → 4.73 s → 17.53 s while candidates double) demonstrates accidental quadratic response work. On a network/removable folder, each redundant `resolve()` can be much slower or block in filesystem I/O, making the finite backend work appear indefinite.

## Exact production Quét call path

1. **React click and state**
   - `frontend/src/components/report/SourcePanel.tsx:25-26`: the **Quét** button calls `s.scan`.
   - `frontend/src/state/store.tsx:505-524`: `scan()` sets `scanBusy`, publishes “Đang quét…”, then makes one `callNative<DashboardDTO>('scan_reports', configDto(), inputFileToken)` call.
2. **JavaScript/pywebview proxy**
   - `frontend/src/state/store.tsx:267-268`: `callNative` delegates to `callBridgeWithOptions`.
   - `frontend/src/lib/bridge.ts:174-224`: waits for the named pywebview method, makes one proxy call, validates the one success/error envelope, and returns its `data`.
3. **Python bridge**
   - `app/bridge.py:52-54`: `BridgeService.scan_reports()` validates/bounds the DTO and token and invokes `ApplicationService.scan_reports()` through `_respond()`.
4. **Application service**
   - `app/application_service.py:172-200`: acquires the service `RLock`, checks idle state, applies folder/month/config values, validates/resolves the selected folder, calls `GuiController.scan()`, saves settings, and calls `dashboard_state()` to build the file-list response.
5. **Filesystem enumeration and structural filtering**
   - `app/gui_controller.py:873-910`: `GuiController.scan()` calls `discover()` exactly once for this Quét invocation.
   - `app/gui_controller.py:824-838`: `discover()` calls `scan_inputs()` and rebuilds internal rows.
   - `app/scanner.py:46-71`: `scan_folder()` recursively walks the selected folder once with `os.walk`, sorts directories/files deterministically, rejects hidden/temp/non-report PowerPoint-looking entries, and accepts `.ppt`, `.pptm`, and `.pptx` candidates. Unrelated files are rejected from candidacy by suffix without opening them.
6. **Period/status/metadata filtering**
   - `app/gui_controller.py:882-909`: resolves the selected processing period, optionally opens the current Excel master once in probe mode, optionally opens the small recent-success cache, and calls `prescan()`.
   - `app/prescan.py:437-564`: files are inspected sequentially. The first pass performs one `Path.stat()` per candidate, filename-only Management Number extraction, date derivation, and month filtering; later in-memory passes handle duplicate selection and optional cache/master status. This code does not open PPT/PPTX bytes.
   - If configured, `ExcelWriter(..., probe=True)` loads the current workbook once and `MasterLookup` supplies status lookups. This is status work required by the current list contract; it was not the demonstrated largest stage in the no-workbook reproduction.
7. **Response construction — demonstrated bottleneck**
   - `app/application_service.py:202-269`: `dashboard_state()` obtains scan rows, creates the list DTOs, collects logs, and returns a dashboard object.
   - `app/application_service.py:295-327`: `_report_dto()` calls `_redact_source_paths()` twice for every row (`shortResult` and `warning`).
   - `app/application_service.py:909-947`: `_collect_logs()` also redacts each new log separately. Every `_redact_source_paths()` invocation reconstructs variants for every report source and invokes `Path.resolve()` twice per source. This produces `O(number of output strings × number of files)` filesystem work, which is `O(N²)` for the per-row strings.
   - The operation is synchronous while the service lock and pywebview call are outstanding. No deadlock, future wait, subprocess, COM call, or parser loop was present in the reproduced profile.
8. **Serialization/return boundary**
   - `app/bridge.py:234-248`: `_respond()` wraps the whole dashboard in one envelope. Actual Python-to-JavaScript serialization is owned by pywebview after the Python method returns and is not directly measurable at the current source boundary.
9. **Frontend list availability/render**
   - `frontend/src/state/store.tsx:509-520`: only after the proxy Promise resolves does the handler clear detail/selection state, call `acceptDashboard`, and clear `scanDirty`/`scanBusy`.
   - `frontend/src/state/store.tsx:284-328`: `acceptDashboard()` normalizes the returned reports and calls `setReports()`.
   - `frontend/src/state/store.tsx:448-460` and `frontend/src/components/report/ReportTable.tsx`: `filteredReports` is an in-memory memoized filter and the table maps the already batched report array. No per-file bridge call is made to populate the initial list. The details effect explicitly skips `waiting`, `new_row`, and `processing` rows, so it does not eagerly parse list entries.

## Proven root cause

**Confidence: high.** The demonstrated bottleneck is unnecessary, synchronous, quadratic filesystem/path redaction during `dashboard_state()` response construction—not PPTX parsing. Each row redacts two ordinary list strings by rebuilding a replacement set across all source paths. Each rebuild resolves every source path twice. For 200 candidates this caused 425 full replacement rebuilds, 253,134 `Path.resolve()` calls, more than 1.26 million filesystem metadata syscalls, and 17.35 of 17.53 seconds.

The implementation therefore recomputes unchanged scan-scoped path information unnecessarily and performs network/removable-drive-hostile metadata access while only trying to sanitize strings for the bridge. The backend has not completed when the UI is waiting: it remains in `SCAN_BUILD_RESPONSE`/`_report_dto` path redaction. A sufficiently slow filesystem can make this appear hung.

No literal infinite loop or deadlock was reproduced on Linux. The strengthened Windows “never appears” symptom still requires Windows log confirmation. The fix must therefore include stage/current-file diagnostics and a slow-stage watchdog so Windows can distinguish a genuinely blocked file operation from this demonstrated unbounded response-building work.

## Root-cause question matrix

| Question | Finding | Evidence / disposition |
|---|---|---|
| A. Parses every PPTX for list | No | Corrupt/non-ZIP `.pptx` fixtures returned correctly; parser count 0; scan path contains no `parse_pptx` call. |
| B. Parses same PPTX more than once | No | Parser count 0. |
| C. Opens PowerPoint COM per file | No | COM count 0 and no renderer/COM call in traced path. |
| D. Starts/stops PowerPoint repeatedly | No | COM count 0. |
| E. Reads ZIP/PPTX contents unnecessarily | No | Invalid synthetic bytes are accepted/listed without reads/parsing. Only filename and stat metadata are used. |
| F. Recurses unexpectedly | Recurses, but this is existing intended contract | `scanner.py` and existing tests explicitly require recursive report folders. One traversal was measured. It is not identified as the bug. |
| G. Walks directory multiple times | No within one Quét | `GuiController.scan()` calls `discover()` once; benchmark counter = 1 traversal. Folder selection in React only sets state and does not scan. |
| H. Performs synchronous expensive work not needed before display | **Yes** | Quadratic filesystem path resolution for output sanitization runs synchronously before response return. |
| I. Calculates preview/evidence eagerly | No | No preview, rendering, image, learning, or evidence call is in the path. |
| J. Recomputes unchanged scan information | **Yes** | Replacement variants for all N source paths are rebuilt 2N+ times in one response. |
| K. Makes bridge call per file | No | One `scan_reports` call returns the batch. |
| L. Serializes oversized objects | Not demonstrated as root cause | DTO contains required list fields plus dashboard state in one response; no images/details. Actual pywebview serialization is not currently measurable. |
| M. Network/removable-hostile access | **Yes** | Redaction performs hundreds of thousands of unnecessary `resolve`/`stat` operations; this is especially harmful on UNC/removable sources. |

## Proposed remediation

**Preferred**:

1. Build a lexical source-path redaction context once per dashboard response and reuse it for logs, scan message, diagnostics, and every report DTO. Do not call filesystem-resolving `Path.resolve()` merely to sanitize a string. This removes the proven `O(N²)` filesystem work without caching scan results or changing list semantics.
2. Reuse the already computed normalized row identity within response construction rather than re-normalizing the same source path for manual-field lookups.
3. Add scan-scoped metrics around the real production path: `SCAN_START`, enumeration start/end, period/status filter start/end, sequential per-file start/end at debug level, a bounded slow-operation top-N, metadata counts/timing, response start/end, bridge return boundary, and total. Include zero-valued PPTX parse and COM counters to make absence explicit.
4. Add a bounded slow-scan watchdog. Normal fast scans emit only aggregate INFO lines; if a scan remains in one stage long enough to appear hung, emit one safe stage/index/basename-or-hash snapshot so Windows can identify the exact last active stage without logging thousands of files or imposing a speculative global timeout.
5. Do not add concurrency, persistent caching, parser changes, COM timeouts, or a global scan timeout. No single corrupt PPTX open occurs at list discovery, so there is no parser/COM operation to time out.

**Files likely to change**:

- `app/application_service.py`
- `app/bridge.py`
- `app/gui_controller.py`
- `app/scanner.py`
- `app/prescan.py`
- a small scan-diagnostics module if needed to keep instrumentation bounded
- `tests/test_scan_performance.py`
- `.specify/bugs/prompt031-scan-performance/fix.md`
- `.specify/bugs/prompt031-scan-performance/test.md`

**Tests to add or update**:

- deterministic large synthetic service/bridge fixture with eligible, outside-month, corrupt synthetic, unrelated, and structurally ineligible entries;
- same returned list/order/status/DTO before and after the optimization;
- one directory traversal and one bridge call;
- unrelated/ineligible files never receive candidate metadata inspection;
- one metadata inspection per discovered candidate;
- zero parser/COM/subprocess calls during Quét;
- bounded linear redaction-context work rather than per-row full rebuild;
- each accepted file inspected no more than required;
- empty folder and isolated file-stat error termination;
- instrumentation includes stage/count summaries and a deterministic watchdog snapshot test;
- existing scanner/list/application-service tests and `tests/test_prompt030_runtime.py`.

## Risks & Considerations

- Redaction must remain fail-safe and must continue removing absolute source paths from every bridge-visible message. The regression must assert no source path leaks.
- Duplicate reasons may contain a selected absolute path; the one-response replacement context must sanitize it exactly as before.
- Existing ordering, month interpretation, statuses, manual-field identity, report isolation, and opaque IDs must remain unchanged.
- Recursive scanning is established behavior and must not be changed merely for speed.
- Excel master lookup can be separately expensive for very large workbooks, but it was not the demonstrated dominant stage. Instrumentation must measure it rather than speculate.
- The pywebview serialization interval is outside the Python method after return. `SCAN_RETURN` will identify that boundary honestly rather than claiming to measure inaccessible transport internals.
- Linux timings are development evidence only. The literal Windows hang, pywebview responsiveness, UNC/removable I/O, and Windows acceptance remain pending.

## Open Questions

- [NEEDS WINDOWS RETEST: Does the real machine reach `SCAN_BUILD_RESPONSE_START`, and does it now reach `SCAN_BUILD_RESPONSE_END`, `SCAN_RETURN`, and `SCAN_TOTAL`?]
- [NEEDS WINDOWS RETEST: If it stalls before response construction, which stage/current index does the bounded slow-scan diagnostic report?]
- [NEEDS WINDOWS RETEST: What are the real folder entry/candidate/accepted/rejected counts and selected-folder storage type (local, UNC, or removable)?]
