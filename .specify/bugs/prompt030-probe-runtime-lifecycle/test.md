# Bug Verification: PROMPT-030T — live-probe pythonnet lifecycle / clean shutdown

- **Slug**: prompt030-probe-runtime-lifecycle
- **Tested**: 2026-10-10
- **Assessment**: ./assessment.md
- **Fix**: ./fix.md
- **Result**: partial

`partial` and not `verified`: the original case is a **Windows native** subprocess crash, and this
sandbox is Linux with no .NET Framework. Every check that can run here was run and is recorded below,
including direct execution of the production probe string; the Windows rerun of
`test_validate_only_passes_a_healthy_folder` is **not-run** and remains the acceptance gate.

## Summary

The reproduced defect was the probe modelling only the initialization half of pythonnet's lifecycle. At
the base commit the same new regression tests fail (`16 failed, 73 passed, 5 skipped`); after the fix all
`99 passed, 5 skipped`. The full Linux suite gained 29 passing tests with no new failures
(`1162 → 1191 passed`, identical 2 pre-existing environment failures). The critical regression — healthy
JSON plus a nonzero subprocess exit still failing — is pinned and was never green-by-weakening: that gate
was already correct and is unchanged. What is NOT yet proven is that a real Windows probe now exits 0
with no `GIL must always be released`; that requires the Windows rerun.

## Checks Performed

| Check | Command / Action | Result | Notes |
|-------|------------------|--------|-------|
| Reproduction (pre-fix) | `python -m pytest tests/test_prompt030_runtime.py -q` at `0502855` with the new tests | fail (as intended) | `16 failed, 73 passed, 5 skipped` |
| Reproduction (pre-fix, executed probe only) | `… -k "probe_source_runs_the_complete_lifecycle or probe_source_never_shuts_down or probe_source_reports_a_failed_shutdown or probe_source_reports_an_unresolvable_shutdown or probe_source_keeps_the_original_failure or probe_source_sends_pythonnets or fails_closed_when_the_shutdown_half"` | fail (as intended) | `7 failed, 97 deselected` |
| Reproduction (post-fix) | `python -m pytest tests/test_prompt030_runtime.py -q` | pass | `99 passed, 5 skipped` |
| New / updated tests | same command | pass | 29 net new tests; 5 skips are pre-existing Windows/`clr_loader`-absent cases |
| Regression — portable packaging | `python -m pytest tests/test_prompt028_portable.py -q` | pass (unchanged) | `2 failed, 74 passed, 10 skipped` — **identical to base**; both need `pywebview`/`pythonnet` (win32-only per `requirements-webview.txt`) |
| Regression — renderer | `python -m pytest tests/test_prompt027r_renderer.py -q` | pass | `25 passed` |
| Regression — regions | `python -m pytest tests/test_prompt027r_regions.py -q` | pass | `20 passed` |
| Regression — full Linux suite | `python -m pytest -q` | pass (unchanged failures) | `2 failed, 1191 passed, 15 skipped` vs base `2 failed, 1162 passed, 15 skipped` |
| Lint / type-check | `python -m pyflakes tools/build_portable.py tests/test_prompt030_runtime.py` | pass | 0 findings |
| Byte-compile | `python -m compileall -q tools/build_portable.py tests/test_prompt030_runtime.py app/desktop.py` | pass | exit 0 |
| Whitespace | `git diff --check` | pass | clean |
| Frontend | — | not-run | no frontend file changed (`git diff --stat`: 2 files); reported as not-affected, not as a pass |
| **Windows acceptance** | `.\.venv-build\Scripts\python.exe -m pytest tests/test_prompt030_runtime.py::test_validate_only_passes_a_healthy_folder -vv -s` | **not-run** | requires real Windows + .NET Framework; skipped in this sandbox, never inferred |
| Portable build / publish / version | — | not-run | out of scope; release safety per constitution §IX |

## Output Excerpts

Pre-fix (base `0502855`, new tests in place):

```
FAILED tests/test_prompt030_runtime.py::test_probe_fails_closed_when_the_shutdown_half_of_the_lifecycle_is_absent
FAILED tests/test_prompt030_runtime.py::test_probe_source_runs_the_complete_lifecycle_and_exits_zero
FAILED tests/test_prompt030_runtime.py::test_probe_source_never_shuts_down_a_runtime_whose_initialize_failed
FAILED tests/test_prompt030_runtime.py::test_probe_source_reports_a_failed_shutdown_and_still_finalizes
FAILED tests/test_prompt030_runtime.py::test_probe_source_reports_an_unresolvable_shutdown
FAILED tests/test_prompt030_runtime.py::test_probe_source_keeps_the_original_failure_when_finalize_also_fails
FAILED tests/test_prompt030_runtime.py::test_probe_source_sends_pythonnets_own_entry_point_arguments
… 16 failed, 73 passed, 5 skipped in 0.66s
```

Post-fix:

```
tests/test_prompt030_runtime.py:  99 passed, 5 skipped in 0.83s
tests/test_prompt027r_renderer.py: 25 passed
tests/test_prompt027r_regions.py:  20 passed
tests/test_prompt028_portable.py:  2 failed, 74 passed, 10 skipped   (== base commit)
full suite:                        2 failed, 1191 passed, 15 skipped in 114.62s
base full suite:                   2 failed, 1162 passed, 15 skipped in 112.48s
```

Environment evidence for the two pre-existing failures (base and post-fix alike):

```
import webview: LỖI – ModuleNotFoundError: No module named 'webview'
import pythonnet: LỖI – ModuleNotFoundError: No module named 'pythonnet'
```

A test that found a real behaviour detail while being written (kept as
`test_probe_source_does_not_finalize_a_clr_it_never_started`): making `pyclr_initialize` itself raise
produces no `finalized` key at all, because the `finally` guard `if stage != "start":` correctly refuses
to finalize a CLR that never started. The probe reported `error="OSError: pyclr_initialize exploded"`,
`stage="start"`, no `finalized` — cleanup was neither performed nor claimed.

## Residual Risks

- **Windows native outcome unproven.** Linux cannot execute the netfx native chain, so "subprocess
  returncode == 0" and "no `GIL must always be released`" are pending the Windows rerun. The executed-
  probe tests prove the probe's Python-level control flow, not real CLR behaviour.
- **Reconstructed finalizer interleaving.** The assessment demonstrates from source that `Py.cs:41`'s
  `~GILState()` is the unique origin of the observed message and that skipping
  `Loader.Shutdown(b"full_shutdown")` leaves `AppDomain.CurrentDomain.ProcessExit` subscribed with the
  managed Finalizer queue full. Which specific `GILState` was collected rather than disposed on the
  Windows machine was not instrumented here.
- **A different residual cause is possible.** If Windows still exits nonzero after an ordered shutdown,
  the next candidates are `PythonEngine.cs:375-383` (`"Python error indicator is set"` →
  `Loader.Shutdown` returns 1, now visible as `probe shutdown_rc 1`) or another undisposed finalizable
  object. The new evidence fields make that diagnosable from the build log without re-running a
  debugger, but it would be a new defect requiring `$speckit-bug-assess`.
- **MOTW untouched.** `Zone.Identifier` remains an unproven hypothesis for the original `H:\` failure;
  nothing here proves or disproves it.

## Recommendation

Hold for Windows acceptance — do not close. Integrate this branch into the PR #7 Windows branch and run

```powershell
.\.venv-build\Scripts\python.exe -m pytest `
    tests/test_prompt030_runtime.py::test_validate_only_passes_a_healthy_folder `
    -vv -s
```

Expect `PASSED`, with `probe appdomain NULL`, `probe resolved True`, `probe initialize_rc 0`,
`probe shutdown_resolved True`, `probe shutdown_rc 0`, `probe finalized True`, `root/current AppDomain`,
subprocess returncode 0, and no `GIL must always be released`. If it still fails, the printed
`shutdown_rc` / `finalized` / `stage` / `finalize_error` fields now identify which lifecycle stage broke;
re-run `$speckit-bug-assess slug=prompt030-probe-runtime-lifecycle` with that output rather than
adjusting the gate.
