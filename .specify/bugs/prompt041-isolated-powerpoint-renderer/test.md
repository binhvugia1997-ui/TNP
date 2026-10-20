# Bug Verification: PowerPoint isolation while the user's PowerPoint is open (PROMPT-041)

- **Slug**: prompt041-isolated-powerpoint-renderer
- **Tested**: 2026-10-17
- **Assessment**: ./assessment.md
- **Fix**: ./fix.md
- **Result**: partial — automated/fake-COM verification passes; **real Windows PowerPoint acceptance NOT performed**. The headline request (PowerPoint render while the user's PowerPoint is open) is **not achievable and is not claimed**; what is verified is that the refusal is safe, correctly diagnosed and no longer over-broad.
- **Spec Kit execution mode**: Executed from the checked-in Spec Kit skill instructions; native Arena skill invocation was not available. `.specify/extensions.yml` has `hooks: {}`.
- **Environment**: Linux sandbox, Python 3.11.2, pytest 9.1.1 in a throw-away `.venv` (gitignored, not committed), Node v22.22.3 / npm 10.9.8. No Windows, no PowerPoint, no pywin32, no LibreOffice.

## Summary

The reported behaviour reproduces deterministically against the PROMPT-035 fake single-instance COM machine on the **unmodified** base (`16 failed, 18 passed` for the new test file as a characterization copy) and the target contract holds after the fix (`35 passed`). No existing test was edited or weakened, including the whole PROMPT-035 ownership suite. The full-suite failure set is unchanged from base `6346388` (one pre-existing, Linux-only failure that needs the win32-only `pythonnet`). Because PowerPoint, the Toolhelp process reader and `ProcessIdToSessionId` are Windows-only, the real-PowerPoint portion is unverified.

## Checks Performed

| Check | Command / Action | Result | Notes |
|-------|------------------|--------|-------|
| Reproduction (pre-fix) | new test file as a characterization copy, unmodified `app/` at `6346388` | **reproduced: 16 failed, 18 passed** | `'PowerPointStageError' object has no attribute 'code'`, `'SlideRenderer' object has no attribute 'preview_reason'`, `no capability probe in base`, `no attribute '_enumerate_powerpoint_pids'`, missing single-instance documentation |
| Behaviour trace (pre-fix) | throw-away harness printing the refusal | `stage=ownership`, `reason=a PowerPoint session is already running…`, `dispatches=0`, `availability={'powerpoint': True, …}`, `last_backend='builtin'`, `detail=-` | proves the cause is the PROMPT-035 pre-flight guard and that nothing machine-readable leaves the renderer |
| Reproduction (post-fix) | `pytest tests/test_prompt041_isolated_powerpoint.py` | **35 passed** (33 functions, `test_5` parametrised ×3) | run repeatedly, stable including the 3-thread test |
| Mutation A: remove the session filter | temporary edit (`return pids` before the session loop), restored | fails `test_16_a_powerpoint_in_another_windows_session_does_not_cause_a_false_refusal` | file restored byte-identical (`diff -q` clean) |
| Mutation B: drop `code=REASON_USER_SESSION_RUNNING` | temporary edit, restored | fails `test_2_user_*`, `test_5[user-session]`, `test_17_the_learning_dto_*` | file restored byte-identical |
| PROMPT-035 ownership suite | `pytest tests/test_prompt035_powerpoint_isolation.py` | **pass, file unmodified** | `git status` shows no change to it; `test_11` also reads its key assertions back out of its source |
| Existing renderer / preview suites | `pytest tests/test_prompt027r_renderer.py tests/test_prompt027_rendering.py tests/test_prompt027r_regions.py tests/test_prompt030b_integration.py tests/test_prompt038_desktop_freshness.py tests/test_prompt035_powerpoint_isolation.py` | **153 passed** | `SlideRenderer.availability()` dict unchanged, so cache identity is unchanged |
| Full Python regression | `.venv/bin/python -m pytest -q tests` | **1 failed, 1356 passed, 13 skipped** (earlier runs: 1354/14 and 1355/13 — one test is conditionally skipped, the total is 1370 either way) | the failure is pre-existing, see below |
| Pre-existing failure isolated | same test with `app/` + `frontend/src` stashed | **fails identically on base** | `test_prompt028_portable.py::test_live_import_probe_reports_every_module_in_this_venv` needs win32-only `pythonnet` / `clr` / `clr_loader` |
| Compile | `python -m compileall -q app tools run.py` | pass | |
| Diagnostics panel | `python -c "from app.diagnostics import run_diagnostics; ..."` via `test_24_*` (three capability states) | pass | `PowerPoint renderer` vs the new `PowerPoint isolation` row stay distinct |
| Lint | `python -m pyflakes app/qpn_renderer.py app/application_service.py app/diagnostics.py tests/test_prompt041_isolated_powerpoint.py` | only pre-existing deliberate "imported but unused" availability probes: 3 in `qpn_renderer.py` (`win32com.client`, `fitz`, `PIL.ImageChops`) and 4 in `diagnostics.py` (`pptx`, `openpyxl`, `PIL`, `requests`); new test file clean | |
| Whitespace | `git diff --check` | clean | |
| Frontend tests | `npm run test:frontend` | **86 passed / 0 failed** (83 baseline + 3 new) | includes the unchanged PROMPT-027 notice contract |
| Frontend typecheck | `npm run typecheck` (`tsc --noEmit`) | clean | |
| Frontend build | `npm run build` (`tsc -b && vite build`) | succeeds — `dist/assets/index-Cwa3JmJU.js`, `index-zEMh39gD.css` | `frontend/dist` is gitignored and is NOT staged |
| Static safety contracts | `test_18_*`, `test_19_*`, `test_22_*` + `grep` over `app/` | pass | no `taskkill`/`TerminateProcess`/`psutil`/`pkill`/`Stop-Process`/`wmic process`, no registry write, no `OpenProcess`, no attach, no enumerate-Presentations, exactly one `.Close()` and one `.Quit()` |
| Real Windows / PowerPoint acceptance | — | **not-run** | requires real Windows + PowerPoint — `docs/WINDOWS_ACCEPTANCE_PROMPT041.md` |
| Windows `ctypes` paths (`Toolhelp32`, `ProcessIdToSessionId`) | — | **not-run** | only the non-Windows `None` branches execute here; the filtering *logic* is tested with injected seams |
| Portable / PyInstaller build | — | **not-run (out of scope)** | no build, no publish, no version/build change |

## Output Excerpts

```
$ pytest -q tests/test_prompt041_isolated_powerpoint.py
...................................                                      [100%]
35 passed

$ pytest -q tests/test_prompt035_powerpoint_isolation.py tests/test_prompt027r_renderer.py \
      tests/test_prompt027_rendering.py tests/test_prompt027r_regions.py \
      tests/test_prompt030b_integration.py tests/test_prompt038_desktop_freshness.py
153 passed

$ pytest -q tests
FAILED tests/test_prompt028_portable.py::test_live_import_probe_reports_every_module_in_this_venv
1 failed, 1356 passed, 13 skipped

$ npm run test:frontend
# tests 86 / # pass 86 / # fail 0
```

Pre-fix characterization run (same tests, unmodified `6346388`):

```
16 failed, 18 passed
AttributeError: 'PowerPointStageError' object has no attribute 'code'
AttributeError: 'SlideRenderer' object has no attribute 'preview_reason'
AttributeError: no capability probe in base
AttributeError: module 'app.qpn_renderer' has no attribute '_enumerate_powerpoint_pids'
AssertionError: qpn_renderer must keep documenting WHY an isolated PowerPoint instance cannot be created
```

## Mandatory-test coverage (the report's 15 items)

| # | Required | Where | Status |
|---|---|---|---|
| 1 | no user PowerPoint → owned renderer works | `test_1_*`, `test_24_*` | pass |
| 2 | user PowerPoint exists → isolated owned renderer succeeds | `test_2_*`, `test_2b`, `test_2c` | **cannot pass — see Deviation** |
| 3 | user presentation never closed/modified | `test_2_*`, `test_3_and_4_*` | pass |
| 4 | unsaved user presentation untouched | `test_3_and_4_*` (refusal + export failure + Quit failure) | pass |
| 5 | isolation unprovable → builtin fallback | `test_5_*` ×3, `test_5b`, `test_5c` | pass |
| 6 | owned export failure → only owned cleanup | `test_6_*`, `test_6b` | pass |
| 7 | cleanup failure → diagnostic, never user cleanup | `test_7_*`, `test_7b` | pass |
| 8 | repeated renders → no ownership leak | `test_8_*` (10 renders, no orphan) | pass |
| 9 | concurrent/rapid requests → lifecycle safe | `test_9_*` (3 threads) | pass |
| 10 | cancellation/shutdown → no orphan | `test_10_*`, `test_10b`, `test_21_*` | pass |
| 11 | existing ownership-refusal regression still meaningful | `test_11_*` + the PROMPT-035 suite run unchanged | pass |
| 12 | PROMPT-030B persistence/restart green | `test_12_*` + `tests/test_prompt030b_integration.py` | pass |
| 13 | PROMPT-038 lifecycle/cache/freshness green | `test_13_*` + `tests/test_prompt038_desktop_freshness.py` | pass |
| 14 | PROMPT-039 updater green | `test_14_*` + `tests/test_updater.py`, `test_auto_update_check.py`, `test_update_progress.py`, `test_publish_update.py` | pass |
| 15 | PROMPT-040 Excel persistence green | `test_15_*` + `tests/test_prompt040_excel_persistence.py`, `test_excel_lock_retry.py` | pass |

## Deviation: mandatory test 2

Test 2 asks for "user PowerPoint exists → isolated owned renderer succeeds". That is **not achievable** and is not faked: PowerPoint is a Multiuse (Single Instance) COM server, so no second `POWERPNT.EXE` can exist in the user's Windows session and there is no isolated instance to own (evidence in ./assessment.md). `test_2_*` pins the honest contract instead — no dispatch, no open, no close, no quit, `code=powerpoint-user-session-running` — and `test_2b` is a tripwire that fails if the single-instance rationale is ever removed from `qpn_renderer.py`, so the isolation question gets re-opened deliberately rather than by accident.

## What Linux cannot prove

- That a real `DispatchEx` on the tester's machine really returns the running `POWERPNT.EXE` (the premise of the whole refusal).
- That the Toolhelp snapshot and `ProcessIdToSessionId` behave as modelled, including the session ids of a real second session / RDP session.
- That the user's unsaved edit, active presentation, window and Save-prompt behaviour are genuinely unaffected end to end.
- That no orphan `POWERPNT.EXE` accumulates on real Windows across repeated renders, failures and app shutdown.
- The `SLIDE_RENDER_BACKEND_FAILED … detail=powerpoint-user-session-running` and `POWERPOINT_ISOLATION state=refused …` lines on a real machine, and the Learning notice as rendered by WebView2.

All of the above are Scenarios A–D in `docs/WINDOWS_ACCEPTANCE_PROMPT041.md`.
