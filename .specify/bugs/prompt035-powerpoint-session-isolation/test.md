# Bug Verification: PowerPoint session isolation (PROMPT-035)

- **Slug**: prompt035-powerpoint-session-isolation
- **Tested**: 2026-10-13
- **Assessment**: ./assessment.md
- **Fix**: ./fix.md
- **Result**: partial — automated/fake-COM verification passes; real Windows PowerPoint acceptance NOT performed
- **Spec Kit execution mode**: Executed from checked-in Spec Kit skill instructions; native Arena skill invocation was not available.
- **Environment**: Linux sandbox, Python 3.11.2, pytest 9-series in a throw-away venv (`/home/user/.venv`, not committed), no Windows, no PowerPoint, no pywin32.

## Summary

The original unsafe lifecycle (user's instance returned by `DispatchEx`, opened into, closed, `Quit()`) reproduces deterministically against a single-instance fake COM machine on the unmodified base (12 of 22 initial tests fail) and no longer reproduces after the fix (24/24 pass). No regressions: the full-suite failure set is byte-identical to base `502968c`. Because the single-instance behaviour and the Toolhelp process-table reader are Windows-only, the real-PowerPoint portion remains unverified.

## Checks Performed

| Check | Command / Action | Result | Notes |
|-------|------------------|--------|-------|
| Reproduction (pre-fix) | `pytest tests/test_prompt035_powerpoint_isolation.py` on unmodified `app/qpn_renderer.py` | reproduced: 12 failed, 10 passed | Quit sent to user's instance, user presentation closed, CancellationRequested misclassified |
| Reproduction (post-fix) | same file after fix | pass: 24 passed (run 3×, stable incl. thread test) | |
| Mutation: remove lock | temporary edit, restored | fails `test_concurrent_jobs_*` as intended | file restored; diff unchanged |
| Mutation: Quit without exclusivity guard | temporary edit, restored | fails `test_7_*` as intended | file restored |
| Existing renderer/COM suite | `pytest tests/test_prompt027r_renderer.py tests/test_prompt027r_regions.py` | pass | 027R harness adapted, assertions untouched |
| Full regression suite (base) | `pytest -q` in a detached worktree of `502968c` | 35 failed, 1182 passed, 18 skipped | pre-existing |
| Full regression suite (fixed) | `pytest -q` | 35 failed, 1206 passed, 18 skipped | failure sets identical (`diff` empty); +24 = new tests |
| Static / compile | `python -m py_compile` on changed files; `python -m pyflakes` | pass (compile); pyflakes clean on new test, only 3 pre-existing deliberate "imported but unused" probes in `qpn_renderer.py` | |
| Whitespace | `git diff --check` | pass | |
| Global-kill contract | `test_8*` static tests + `grep` for taskkill/TerminateProcess/os.kill/psutil in `app/ tools/ *.bat` | pass | only `app/updater.py` POSIX-only `os.kill(pid, 0)` liveness probe of the *old app's* PID (pre-existing, unrelated) |
| Real Windows / PowerPoint acceptance | — | **not-run** | requires real Windows + PowerPoint |
| Windows process-table reader (`ctypes` Toolhelp) | — | **not-run** | only the non-Windows `None` branch is exercised |
| Portable/PyInstaller build | — | **not-run (out of scope)** | no build, no publish, no version/build change |

## Output Excerpts

```
tests/test_prompt035_powerpoint_isolation.py ........................   24 passed
base:  35 failed, 1182 passed, 18 skipped
fixed: 35 failed, 1206 passed, 18 skipped      (identical FAILED set)
```
Pre-existing failures at base are unrelated: stale `1.3.5 / Build 018` version assertions (release is 1.3.6 / Build 019), updater/Portable/runtime-path tests that need Windows or a frozen build.

## Windows Acceptance Plan (minimal; NOT yet performed)

1. On Windows with PowerPoint, open PowerPoint manually, open any unrelated presentation, type an **unsaved** edit; leave it open.
2. Run one small Report Extractor PPTX job (single report, with a learning preview or normal extraction that renders slides).
3. Verify: PowerPoint is still running; the user's presentation is still open; the unsaved edit is still there; Report Extractor completes (log shows `SLIDE_RENDER_BACKEND_FAILED ... backend=powerpoint stage=ownership` then LibreOffice/built-in fallback, with the reduced-fidelity notice); no Quit/close dialog appeared.
4. Close the user's PowerPoint completely, run the same job again: verify the render uses `backend=powerpoint faithful=true`, and that **no `POWERPNT.EXE` remains** afterwards (Task Manager / `Get-Process POWERPNT`).
5. (Optional, 1 min) While a job is rendering with PowerPoint, double-click an unrelated .pptx: it must stay open afterwards.

## Residual Risks

- **Not verified on Windows**: that real PowerPoint behaves as the fake models (single instance), that `ctypes` Toolhelp enumeration works in the Portable build, and the end-to-end "unsaved edit survives" outcome.
- **Residual TOCTOU**: a user launching an *empty* PowerPoint in the few milliseconds between the pre-check and `DispatchEx` returning cannot be distinguished by PID/presentation/window checks from a fresh instance. Narrow, mitigated by two pre-checks + post-dispatch proof + pre-Quit exclusivity check, not eliminated.
- Fail-closed means reduced-fidelity rendering whenever the user's PowerPoint is running (by design).
- An app-owned instance whose `Quit()` fails is left running (logged, never killed); later jobs fail closed until it exits.
- 35 pre-existing failures remain at base and are untouched.

## Recommendation

Hold for Windows acceptance. Automated evidence shows the audited lifecycle (the only COM path in the product) can no longer Close/Quit a user-owned PowerPoint session or presentation; run the short Windows plan above before labelling the fix `verified` end-to-end.
