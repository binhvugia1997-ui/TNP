# Bug Verification: worker dies silently mid-batch after a PowerPoint render refusal

- **Slug**: worker-death-on-render-failure
- **Tested**: 2026-11-19
- **Assessment**: ./assessment.md
- **Fix**: ./fix.md
- **Result**: partial

## Summary

The six new regression tests pass on Linux, and the full Linux suite is green
except one pre-existing environment failure (pythonnet not installed in the
Linux dev venv — unrelated to this fix). The original live case (PowerPoint
open on the company machine during a 61-file run) was not re-exercised here;
Windows full-suite + portable-build acceptance on ollama is still pending.

## Checks Performed

| Check | Command / Action | Result | Notes |
|-------|------------------|--------|-------|
| New regression tests | `.venv-build/bin/python -m pytest tests/test_worker_death_retry.py -q` | pass | 6 passed in 12.21s (Linux) |
| Full regression suite (Linux) | `.venv-build/bin/python -m pytest tests -q -p no:cacheprovider` | partial | 1413 passed, 13 skipped, 1 failed in 590.40s — see below |
| Bump-affected files re-run after the Build 025 bump (Linux) | `.venv-build/bin/python -m pytest tests/test_worker_death_retry.py tests/test_auto_update_check.py tests/test_build_and_publish.py tests/test_content_learning.py tests/test_gui_redesign.py tests/test_gui_responsive.py tests/test_image_learning.py tests/test_prompt001.py tests/test_prompt027_rendering.py tests/test_prompt028_portable.py tests/test_publish_update.py tests/test_update_progress.py tests/test_updater.py tests/test_application_service.py -q` | partial | 403 passed, 8 skipped, 1 failed in 248.57s — the same single environment failure (see below); all version assertions now expect Build 025 and pass |
| Windows full suite (system Python 3.14) | ollama `C:\Python314\python.exe -m pytest tests -q` | partial | 1406 passed, 16 skipped, 5 failed in 1028.99s — see analysis below |
| Windows re-run of 3 failures (system Python 3.14) | ollama `pytest tests/test_runtime_paths.py::test_startup_error_log_written_next_to_exe tests/test_runtime_paths.py::test_startup_logging_writes_startup_env_line tests/test_application_service.py::test_learning_candidates_counts_label_review_saved_reapply_retry_and_model_update -q` | pass | 3 passed in 34.76s (after the `test_runtime_paths.py` bump fix) |
| Windows full suite (build venv) | ollama `tools\build_portable.py --python <uv cpython-3.12.16> --no-publish --fresh-venv` (runs `pytest -q` in `.venv-build`) | partial | Run 1: **1424 passed, 2 skipped, 1 failed** in 940.13s. Run 2: **same 1424/1/2 in 939.32s**. Run 3 (clean PowerPoint state): **same 1424/1/2 in 966.11s** — the same single flaky COM test every time (run 3 proximate cause: `powerpoint-presentation-open`). Both pythonnet tests and both `test_runtime_paths` tests PASS in this venv in all runs. |
| Portable build | ollama `run_build2.bat` equivalent (see above; `run_build2.bat` itself no longer exists on the machine — invoked `tools\build_portable.py` directly, log at `build2.log`) | failed | **BLOCKED: all 3 build runs failed at the test gate (step [04])** on the same single pre-existing flaky test. No portable was produced. No PR opened; no deploy. See blocker note below. |
| Original live case (company machine, PowerPoint open) | Gia live test | not-run | needs Gia on the company machine |

## Output Excerpts

New tests: `......  [100%]  6 passed in 12.21s`.

Full Linux suite: `1 failed, 1413 passed, 13 skipped, 1 warning in 590.40s`.
The single failure:

```
FAILED tests/test_prompt028_portable.py::test_live_import_probe_reports_every_module_in_this_venv
   import webview: OK
   import pythonnet: LỖI – ModuleNotFoundError: No module named 'pythonnet'
```

## Windows full-suite failure analysis (system Python 3.14, 19/11)

`5 failed, 1406 passed, 16 skipped in 1028.99s`. None of the five is a
regression from this fix (the fix touches only `app/batch_processor.py` and
`app/logger.py`):

- `test_runtime_paths.py::test_startup_error_log_written_next_to_exe` and
  `test_startup_logging_writes_startup_env_line` — **bump miss, fixed**: both
  asserted a hardcoded `"version=1.4.0 build=024"` / `"STARTUP version=1.4.0
  build=024 …"` while the code correctly reports `build=025`. The strings live
  in `tests/test_runtime_paths.py` (lines 158, 189), which the first repo-wide
  sweep missed (its output was cut at `head -40`). Bumped to `025`; both pass
  on re-run (Linux 15/15, Windows 2/2). A second untruncated sweep confirmed
  no other current-build "024" remains.
- `test_prompt028_portable.py::test_live_import_probe_reports_every_module_in_this_venv`
  and `test_prompt030_runtime.py::test_validate_only_passes_a_healthy_folder` —
  **wrong interpreter, not a code failure**: both require pythonnet / clr_loader
  / cffi in the test interpreter, and the system `C:\Python314` has none of
  them (`pytest.fail: Windows live-runtime integration prerequisites are
  missing from C:\Python314\python.exe`). The Build 024 Windows run passed
  both because it ran in a venv with pythonnet. The build venv
  (`.venv-build`, `requirements-webview.txt` pins pythonnet 3.0.5 +
  clr-loader 0.2.10) is the correct interpreter; the build's internal
  `pytest -q` run is the deciding check for these two.
- `test_application_service.py::test_learning_candidates_counts_label_review_saved_reapply_retry_and_model_update`
  — **flaky PowerPoint COM, not a regression**: failed in the full run with
  `Windows fatal exception: code 0x80010108` (RPC_E_DISCONNECTED) inside
  `qpn_renderer._render_with_powerpoint_locked` during a learning-label
  preview; the same test passed in the Build 024 Windows run, passed on Linux
  (render stub), and **passed on isolated Windows re-run** (3 passed in
  34.76s, with the same COM stderr noise present but the assertions green).
  The test never calls `BatchProcessor.run()` or `BatchResultLog`, so this
  fix cannot cause it. The COM disconnect is timing/state-dependent.
  Mechanism (read from the test + build log): the test's `slow_train_models`
  waits at most 5s for the main thread to call `release.set()`; the main
  thread first calls `learning_state()`, which renders slide previews through
  real PowerPoint COM — when those renders are slow (COM timeouts, leftover
  PowerPoint processes), the 5s window expires and the test fails. In the
  build run 1 the failure was `release.wait(timeout=5)` = False with
  `SLIDE_RENDER_BACKEND_FAILED … code=powerpoint-presentation-open` and
  `app-exit-wait-timeout` for PowerPoint pids created during the run.
  A stale headless PowerPoint (`POWERPNT.EXE /AUTOMATION`, pid 15128, started
  10:58, no window, orphaned) was found on ollama and terminated before the
  build re-run; it had been visible in the renderer's snapshot during run 1.
  Run 2 failed the same way with a different proximate cause: the renderer's
  own PowerPoint (pid 16860, created 12:56:49 by a `learning_preview` render)
  did not quit within the cleanup wait (`app-exit-wait-timeout` 12:57:04, "still
  running; not terminated"), so the next render in the same test was refused
  (`POWERPOINT_ISOLATION state=refused reason=powerpoint-user-session-running`)
  and the 5s train window expired again. The test therefore fails 3/3 recent
  full-suite runs (system Python + build venv run 1 + build venv run 2) while
  passing in isolation and in the Build 024 full run — a pre-existing
  PowerPoint-COM timing fragility in the test/renderer, not this fix. Runs 1–3
  of the build all failed at the test gate on this one test (run 3: same
  1424/1/2 in 966.11s, proximate cause `powerpoint-presentation-open` in the
  learning preview), so the build is BLOCKED: no portable was produced, no PR
  was opened and nothing was deployed. The flaky test needs its own bug batch
  (do NOT weaken the 5s timeout or the test to force green, and do NOT bypass
  the build's test gate with --skip-tests). The Build 025 code, tests, bump
  and this test evidence are committed on the Arena branch
  `arena/06a1c335-tnp` (scoped paths only) so the work is preserved in git;
  the PR waits on the flake fix and a green build.

## Recommendation

Do NOT mark this `verified`. The fix itself is verified on Linux and in the
Windows build venv (1424/1425, the only failure being the pre-existing flaky
PowerPoint COM test, which passes in isolation with the fix present). The
portable build, PR and deploy are blocked on that flaky test; the next step is
a separate bug batch for
`test_learning_candidates_counts_label_review_saved_reapply_retry_and_model_update`
(PowerPoint render slowness/refusal under full-suite load vs the 5s train
timeout), then a green build re-run.

## Environment notes (ollama, 19/11)

- `C:\Users\muse.OLLAMA\TNP` is a **junction to `D:\muse.OLLAMA\TNP`** (created
  17/11); the real repo, venv and build outputs live on D:. All C: paths used
  in this batch resolve to the same D: files.
- `run_build2.bat` no longer exists on the machine; the build was invoked
  directly as `tools\build_portable.py --python <uv cpython-3.12.16>
  --no-publish --fresh-venv`, log at `build2.log`.
- `uv python install cpython-3.12.16` was needed — the managed 3.12.16 was
  missing (only 3.13.18 was installed; the registry pointed at the deleted
  `D:\HanNguPython\Python312`).
- Push blocked (19/11): the `arena/06a1c335-tnp` branch + commit `1d67d6b`
  exist only on ollama — `git push` fails because the HTTPS remote has no
  usable credential in a headless SSH session (`could not read Username for
  'https://github.com'`, wincredman unavailable) and `gh` is not logged in
  on the machine. Needs an interactive `gh auth login` (or equivalent) on
  ollama before the branch can be pushed and the PR opened.

## Residual Risks

- The Linux failure is environmental, not a regression: the test probes
  whether `pythonnet` imports in the current venv, and `.venv-build` on this
  Linux VM does not have `pythonnet` installed (verified: `import pythonnet`
  → `ModuleNotFoundError`). The fix touched only `app/batch_processor.py` and
  `app/logger.py`; it cannot change which packages are installed. The same
  test passed on Windows in the Build 024 run (1419 passed, 0 failed) and in
  the Build 025 build venv (1424 passed), where the venv has pythonnet. The
  test was NOT weakened or skipped.
- The exact exception that killed the 19/11 worker remains unverified (the
  delivered `app-4.log` is a partial slice; see assessment). The fix closes
  every demonstrated structural hole, but a live re-run on the company
  machine with PowerPoint open is the only end-to-end confirmation.
- Retry + `MASTER_NEW_RETRY` interaction (a file that created a row, then
  failed, then retries) is covered by the existing partial-row logic but has
  no dedicated new test in `append` mode; the Windows acceptance run should
  include a 61-file live batch if Gia re-runs it.
