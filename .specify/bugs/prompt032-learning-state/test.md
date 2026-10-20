# Bug Test & Verdict: Learning state

- **Slug**: prompt032-learning-state
- **Date**: 2026-10-10
- **Starting SHA**: `2ca7dbfeb12a414933f0efc2b4f3f8fe6a501361`
- **Working branch**: `arena/2e088bee-tnp`
- **Assessment**: `.specify/bugs/prompt032-learning-state/assessment.md`
- **Fix record**: `.specify/bugs/prompt032-learning-state/fix.md`
- **Spec Kit execution**: Executed from checked-in Spec Kit skill instructions; native Arena skill invocation was not available.

## Verdict

**READY FOR WINDOWS LEARNING RETEST.**

The deterministic Linux production fixture and all requested backend/bridge/frontend regressions pass. A Learning read now finalizes a completed worker's queued lifecycle event and returns its retained image and content candidates without a prior dashboard read or restart. Successful empty, no-run, processing, unavailable, malformed-response, and backend-failure outcomes are distinguishable and visibly rendered. Preview failure remains display-only and does not erase candidate metadata.

This is not Windows acceptance. Linux did not execute pywebview/WebView2 on Windows, PowerPoint COM rendering, or the user's real PowerPoint source files.

## 23 PROMPT-032 regression contracts

| # | Contract | Evidence | Result |
|---:|---|---|---|
| 1 | Production terminal batch creates an eligible reviewable image candidate | `test_completed_production_batch_learning_read_reconciles_and_returns_both_candidate_kinds` | PASS |
| 2 | The same production batch creates reviewable content candidates | same | PASS |
| 3 | Both object sets remain in `FileResult` after worker completion | same | PASS |
| 4 | Direct Learning read reconciles queued terminal event without dashboard | same | PASS |
| 5 | Direct read returns current image candidates | same | PASS |
| 6 | Direct read returns current content candidates | same | PASS |
| 7 | A second manual-style bridge read is fresh backend state | same | PASS |
| 8 | Image hard exclusions remain filtered | production raw-versus-DTO count assertion | PASS |
| 9 | Content hard exclusions remain filtered and content remains verbatim | production raw-versus-DTO/text assertions | PASS |
| 10 | Successful zero-image snapshot is explicit valid empty | production second read + availability matrix | PASS |
| 11 | Successful zero-content snapshot is explicit valid empty | production second read + availability matrix | PASS |
| 12 | No completed run is distinct from valid empty | `test_learning_availability_distinguishes_no_run_processing_valid_empty_and_unavailable` | PASS |
| 13 | Processing and learning-unavailable states are independently explicit | same | PASS |
| 14 | Preview exception retains candidate identity/decision/eligibility/geometry/evidence/fallback | `test_image_preview_exception_keeps_truthful_candidate_metadata` | PASS |
| 15 | Bridge returns availability/count contract with no absolute source path | production bridge assertion | PASS |
| 16 | `LEARNING_*` state diagnostics are transition-bounded and payload/path-free | `test_learning_diagnostics_are_transition_bounded_and_do_not_leak_payloads` | PASS |
| 17 | Entering top-level Learning performs a backend call | `frontend/tests/learning-hooks.test.mjs` | PASS |
| 18 | Manual refresh performs an additional backend call and can recover | same | PASS |
| 19 | Populated image response visibly renders image data | same + PROMPT-027 mounted tests | PASS |
| 20 | Populated content response visibly renders verbatim content | same | PASS |
| 21 | Valid-empty image and content each render truthful visible feedback | same | PASS |
| 22 | Missing required candidate field renders contract failure, not empty | same | PASS |
| 23 | Backend rejection is visible in both review paths and retryable | same | PASS |

## Executed verification

### Focused PROMPT-032 and affected rendering/scan/runtime suite

```text
$ source .venv/bin/activate && pytest -q \
    tests/test_prompt032_learning_state.py \
    tests/test_application_service.py \
    tests/test_image_learning.py \
    tests/test_content_learning.py \
    tests/test_prompt024r_learning_workspace.py \
    tests/test_after_evidence.py \
    tests/test_prompt027_rendering.py \
    tests/test_prompt027r_regions.py \
    tests/test_prompt027r_renderer.py \
    tests/test_scan_performance.py \
    tests/test_incremental_scan_cache.py \
    tests/test_prompt030_runtime.py
312 passed, 1 skipped in 24.32s
```

This covers affected image/content evidence and rendering, PROMPT-024R Learning workspace, PROMPT-027/027R, PROMPT-031/031R, and PROMPT-030. The one skip is the established platform-conditional test; it was not converted to a pass.

### Frontend tests, typecheck, and production build

```text
$ cd frontend && npm run test:frontend
32 passed, 0 failed

$ npm run typecheck
PASS (tsc --noEmit)

$ npm run build
PASS (tsc -b && vite build; 1,588 modules transformed)
```

The first combined frontend attempt exposed an existing PROMPT-027R source-contract assertion that requires the literal hook-safe `if (!cand) return` guard. The implementation was adjusted without moving or adding hooks after that guard; the complete frontend suite was rerun and passed as shown above.

### Full Linux Python suite — run once after focused suites were green

```text
$ source .venv/bin/activate && pytest -q
1222 passed, 3 skipped in 120.00s
```

No failed or xfailed test was hidden. The three skips are existing platform/fixture conditions.

### Compile, changed-path lint, version, and diff checks

```text
$ python -m compileall -q app tests/test_prompt032_learning_state.py
PASS

$ python -m pyflakes app/application_service.py app/batch_processor.py tests/test_prompt032_learning_state.py
PASS

$ git diff --check
PASS
```

Version remains `1.3.4`; authoritative build remains `017` (`BUILD_NUMBER = 17`). No publish/build batch file, Portable output, LAN update, stable backup, `H:`, PR merge, or main-branch operation was run.

## Preserved contracts

- Actual candidate generation and eligibility are unchanged; tests use production-generated candidates rather than hardcoded candidates to prove lifecycle retention.
- Ollama remains optional; the production regression runs with `use_ollama=False` and still generates both candidate kinds.
- Report terminal status and Excel semantics are unchanged.
- Content remains verbatim and no new content-learning meaning was introduced.
- PROMPT-027/027R target/evidence overlay separation, item ownership, renderer preference/fallback diagnostics, preview cache identity, unmodified bitmap display, and report isolation remain green.
- PROMPT-031/031R incremental scan generations, leaf publishing, bounded polling, stale-revision rejection, and scan performance remain green.
- Preview failure affects display fields only; candidate metadata remains truthful.

## Exact Windows source retest and log capture

Run this against the assigned source branch only; do not use `BUILD_AND_PUBLISH.bat` and do not create/publish a Portable package.

### 1. Update and identify the exact source

From PowerShell in the source checkout:

```powershell
git fetch origin --prune
git switch arena/2e088bee-tnp
git pull --ff-only origin arena/2e088bee-tnp
git status --short
git rev-parse HEAD
```

Expected: branch `arena/2e088bee-tnp`, clean tracked source before local frontend build, and the pushed PROMPT-032 commit reported in the delivery message.

### 2. Build the local frontend from this same source checkout

```powershell
cd frontend
npm ci
npm run typecheck
npm run build
cd ..
```

This is a local source build only. Do not run the publish script.

### 3. Start the source application and process a known eligible report

```powershell
$env:PYTHONUTF8 = "1"
python -m app.desktop
```

In the application:

1. Confirm the header still shows `v1.3.4 · Build 017`.
2. Select a real report whose improvement slide visibly has at least one eligible **Sau cải tiến** picture and at least one reviewable production-improvement text block.
3. Select the real template/output, choose **Quét**, review the one-report queue, and process with normal settings. Ollama may be unavailable; deterministic candidates must still be produced.
4. Wait until the report is visibly terminal (`Hoàn thành` or `Cần kiểm tra`). Do not restart the app.
5. Immediately enter **Học cải tiến → Kiểm tra ảnh cải tiến**. Confirm at least one current-run candidate is visible, its report/slide/item/decision/eligibility metadata is present, and the full-slide preview or truthful cropped-image fallback appears.
6. Confirm the two PROMPT-027 overlays retain their meanings: dashed amber = the picture being reviewed; solid blue = the authoritative item-scoped region that would be exported to Excel. The PowerPoint bitmap must remain unmodified.
7. Enter **Kiểm tra nội dung cải tiến**. Confirm at least one current-run content candidate and its verbatim source text are visible.
8. Press **Làm mới trạng thái học** in each review path. Confirm it performs a real refresh and the candidates remain current without restart.
9. If a report legitimately has no candidate for one kind, confirm the view says the current run has no candidate; it must not claim that Python omitted data. Do not change eligibility to manufacture a candidate.
10. Keep PowerPoint available for the fidelity check. If the preview falls back, the UI/log must say so; fallback is not PowerPoint-fidelity acceptance.

### 4. Capture bounded Learning diagnostics

After closing the application (or while tailing from another PowerShell), inspect the output's log directory:

```powershell
$log = Join-Path "<EXACT_OUTPUT_FOLDER>" "logs\app.log"
Select-String -Path $log -Pattern "LEARNING_|SLIDE_RENDER purpose=learning_preview" |
  Select-Object LineNumber, Line |
  Format-Table -AutoSize
```

Required evidence:

```text
LEARNING_CANDIDATES_RETAINED report=<opaque-id> image_raw=<N> image_reviewable=<N> content_raw=<N> content_reviewable=<N>
LEARNING_STATE lifecycle=idle worker_alive=False reports=<N> image_raw=<N> image_reviewable=<N> content_raw=<N> content_reviewable=<N> image_state=ready ... content_state=ready ...
SLIDE_RENDER purpose=learning_preview ... selected_backend=powerpoint faithful=true ...
```

For the chosen acceptance report, `image_reviewable` and `content_reviewable` must each be greater than zero. The log must not contain an absolute source path in any `LEARNING_*` line, candidate text/payload, data URI, or image bytes. Repeated idle refreshes should not add duplicate `LEARNING_STATE` lines unless the lifecycle/count/status signature changed.

If the source report does not naturally produce both candidate kinds under current rules, stop and select a known eligible report; do not disable rules or fabricate candidates. If PowerPoint rendering is unavailable or the Windows result differs, preserve the source, exact SHA, screenshot, and the filtered log lines and report the retest as failed/pending rather than changing semantics.

## Remaining platform acceptance

- Windows pywebview/WebView2 direct post-run Learning read and manual refresh.
- Real PowerPoint faithful preview and visual overlay alignment against the authored slide.
- Real customer report image/content eligibility under unchanged production rules.
- Windows log-path/redaction and transition-bounded diagnostic confirmation.
