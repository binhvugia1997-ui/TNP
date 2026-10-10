# Bug Fix: PROMPT-030T — the live probe now runs pythonnet's whole lifecycle, not just Initialize

- **Slug**: prompt030-probe-runtime-lifecycle
- **Fixed**: 2026-10-10
- **Assessment**: ./assessment.md
- **Status**: applied

## Summary

`PROBE_PACKAGED_RUNTIME` in `tools/build_portable.py` now models the lifecycle the pinned upstream
sources actually define — `pyclr_initialize` → `pyclr_create_appdomain(NULL, NULL)` →
`Loader.Initialize(b"")` → **`Loader.Shutdown(b"full_shutdown")`** → **`pyclr_finalize()`** — instead of
stopping after `Initialize`. `pyclr_finalize()` runs in a `finally`, gated on the lifecycle stage actually
reached, so a probe never abandons a CLR it started, never shuts down a runtime that never initialized,
and never finalizes a CLR it never started. `probe_packaged_runtime()` now demands positive evidence for
every stage it relies on (`resolved`, `initialize_rc`, `shutdown_resolved`, `shutdown_rc`, `finalized`,
no `finalize_error`) and still fails CLOSED on anything missing, malformed, non-True or nonzero. The
subprocess exit-code gate was **not** weakened. The live probe is still the real one: it still dlopens the
packaged `_internal/clr_loader/ffi/dlls/<arch>/ClrLoader.dll` and resolves both entry points out of the
packaged `_internal/pythonnet/runtime/Python.Runtime.dll`; no mock, static-only or import-only substitute
was introduced.

## Changes

| File | Change | Notes |
|------|--------|-------|
| `tools/build_portable.py` | modified | (1) `PROBE_PACKAGED_RUNTIME`: added a `stage` tracker (`start` → `clr-initialized` → `appdomain` → `pythonnet-initialized` → `pythonnet-shutdown`); resolve `Python.Runtime.Loader.Shutdown` and call it with `b"full_shutdown"` (buffer length 13, matching `clr_loader.types.ClrFunction.__call__`) **only inside the `initialize_rc == 0` branch**; record `shutdown_resolved` / `shutdown_rc` / `stage`; added a `finally` block that calls `fw.pyclr_finalize()` exactly once, only when `stage != "start"`, recording `finalized` (`True`/`False`) and a separate `finalize_error`. `out["error"]` is still assigned in exactly one place (the `except` block), so a cleanup exception cannot bury the original finding; `error_stage` records where it died. `pyclr_close_appdomain` stays uncalled for the root/NULL handle (upstream `NetFx.shutdown()` skips it, `ClrLoader.CloseAppDomain` no-ops on `IntPtr.Zero`). The cdef and the packaged paths are unchanged. The `#:` comment block above the probe now records the verified lifecycle with file/line evidence. (2) New `_lifecycle_rc()` helper validates one stage's `int Func(IntPtr, int)` result (non-bool `int`, equal to 0) for both `Initialize` and `Shutdown`. (3) `probe_packaged_runtime()`: extended the printed evidence keys to `shutdown_resolved`, `shutdown_rc`, `finalized`, `stage`, `finalize_error`, `error_stage` (field width 14 → 17); the `initialize_rc` check now goes through `_lifecycle_rc` and yields an `initialized` flag; the shutdown/finalization checks run **only** after `initialized` is true, so a failed `Initialize` stays the single visible finding. New fail-closed checks: `shutdown_resolved` must be `True`, `shutdown_rc` must be integer 0, `finalized` must be `True`, `finalize_error` must be absent. Docstring updated; the returncode/JSON-object/arch/appdomain/`resolved` checks and the `error` early-return are unchanged. |
| `tests/test_prompt030_runtime.py` | modified | `HEALTHY_ROOT_DOMAIN_EVIDENCE` extended to the complete lifecycle; new `INITIALIZE_ONLY_EVIDENCE` pins the pre-030T payload that must now fail closed; added the PROMPT-030T lifecycle section (validator-level fail-closed matrix + source-level ordering assertions) and a new section that **executes** `bp.PROBE_PACKAGED_RUNTIME` itself against a stub `cffi` so the probe's stage/cleanup control flow is proven on any host; `test_probe_accepts_a_runtime_that_resolves`, the end-to-end gate test and the Windows live test were updated to the lifecycle shape; module docstring documents the lifecycle. No existing test was weakened or deleted. |

`app/desktop.py` was **not** modified: it diagnoses the frozen app's own start-up
(`Failed to resolve Python.Runtime.Loader.Initialize`, `appdomain_report()`), and the frozen app's
pythonnet lifecycle is owned by pythonnet's own `atexit.register(unload)`. It has no relationship to the
build-time probe subprocess. `git diff --stat` confirms only the two files above changed.

## Diff Highlights

Probe — the second half of the lifecycle, nested inside the successful-initialize branch:

```python
    if out["resolved"]:
        buffer = ffi.from_buffer("char[]", b"")
        out["initialize_rc"] = int(func(ffi.cast("void*", buffer), 0))
        if out["initialize_rc"] == 0:
            stage = "pythonnet-initialized"
            shutdown = fw.pyclr_get_function(domain, str(assembly).encode("utf8"),
                                             b"Python.Runtime.Loader", b"Shutdown")
            out["shutdown_resolved"] = shutdown != ffi.NULL
            if out["shutdown_resolved"]:
                command = ffi.from_buffer("char[]", b"full_shutdown")
                out["shutdown_rc"] = int(shutdown(ffi.cast("void*", command), len(command)))
                if out["shutdown_rc"] == 0:
                    stage = "pythonnet-shutdown"
except BaseException as exc:
    out["error"] = f"{type(exc).__name__}: {exc}"[:400]     # the ONLY assignment to out["error"]
    out["error_stage"] = stage
finally:
    if stage != "start":                                    # only if pyclr_initialize() really ran
        try:
            fw.pyclr_finalize()
            out["finalized"] = True
        except BaseException as exc:
            out["finalized"] = False
            out["finalize_error"] = f"{type(exc).__name__}: {exc}"[:400]
    out["stage"] = stage
```

Validator — shutdown/finalization demanded only once initialization is proven:

```python
        initialized = _lifecycle_rc(out, "initialize_rc", "Python.Runtime.Loader.Initialize",
                                    "pythonnet không khởi động được trong gói đã build", problems)
    if not initialized:
        return problems          # never bury a failed Initialize under invented shutdown complaints
    …
    if not shutdown_ok:
        return problems          # never judge finalization of a runtime that could not be shut down
    finalized = out.get("finalized", None)
    if finalized is not True:
        problems.append(…)
    if out.get("finalize_error"):
        problems.append(f"pyclr_finalize() thất bại: {out['finalize_error']}")
```

## Tests Added or Updated

Validator-level (decision contract, any host):

- `test_probe_fails_closed_when_the_shutdown_half_of_the_lifecycle_is_absent` — **D**, the reproduced
  Windows payload (Initialize-only evidence) must fail closed.
- `test_probe_accepts_the_complete_root_domain_lifecycle` — **A**, NULL root domain + Initialize 0 +
  Shutdown resolved rc 0 + finalized, rc 0 → PASS, and every stage stays visible in the build log.
- `test_probe_fails_when_shutdown_cannot_be_resolved_after_successful_initialize` — **B**.
- `test_probe_fails_closed_on_malformed_shutdown_evidence` (5 params) — **E** missing/non-boolean
  `shutdown_resolved`, missing/`false`/`true` `shutdown_rc`.
- `test_probe_rejects_a_nonzero_shutdown_result` — **C**.
- `test_probe_fails_when_clr_loader_finalization_did_not_succeed` (4 params) — **F** raised, reported
  false, evidence missing, JSON `1` instead of `true`.
- `test_probe_fails_closed_on_nonzero_exit_even_with_a_complete_healthy_lifecycle` — **G**, the critical
  regression: perfect lifecycle JSON + `returncode=1` + the real GIL message on stderr still FAILS, and
  the .NET cause stays visible.
- `test_probe_does_not_demand_shutdown_evidence_from_a_runtime_that_never_initialized` — **H**, asserts
  the exact single problem string.
- `test_probe_invents_no_lifecycle_evidence_after_an_exception_before_initialize` — **I**.
- `test_gate_rejects_an_initialize_only_probe_end_to_end` — end-to-end gate stops the build.
- Updated: `test_probe_accepts_the_root_appdomain_null_sentinel` (**J**),
  `test_probe_fails_closed_on_incomplete_or_malformed_evidence` (**K**),
  `test_probe_fails_closed_on_nonzero_exit_with_parseable_json`,
  `test_probe_fails_closed_on_unparseable_output`, `test_probe_reports_a_timeout_as_a_failure`,
  `test_probe_accepts_a_runtime_that_resolves`,
  `test_gate_accepts_the_healthy_root_domain_probe_end_to_end`,
  `test_validate_only_passes_a_healthy_folder` (Windows acceptance now also asserts the shutdown/
  finalization lines and the absence of `GIL must always be released`).

Source-level (**L**):

- `test_probe_source_models_the_verified_shutdown_and_finalize_lifecycle` — `Loader.Shutdown` +
  `b"full_shutdown"` + `fw.pyclr_finalize()` present, in the order Initialize → Shutdown → finalize,
  with a `finally`, and all three evidence keys emitted.
- `test_probe_source_only_shuts_down_a_runtime_that_initialized` — Shutdown nested under
  `if out["initialize_rc"] == 0:`, finalize gated on `if stage != "start":`, root-domain request and the
  `appdomain` sentinel line preserved (030S-R).
- `test_probe_source_never_lets_a_cleanup_error_hide_the_original_failure` — one assignment to
  `out["error"]`, `finalize_error` separate, `finalized = False` on failure.

Executed-probe (runs the production string itself against a stub `cffi`; distinguishes
initialized / shutdown / finalized / process-exit):

- `test_probe_source_runs_the_complete_lifecycle_and_exits_zero` — full evidence dict and rc 0, plus the
  packaged `host`/`assembly` paths.
- `test_probe_source_never_shuts_down_a_runtime_whose_initialize_failed` — no `shutdown_*` keys, rc
  preserved, finalize still ran, stage `appdomain`.
- `test_probe_source_reports_a_failed_shutdown_and_still_finalizes` — stage stops at
  `pythonnet-initialized`, `finalized True`.
- `test_probe_source_reports_an_unresolvable_shutdown` — `shutdown_resolved False`, no `shutdown_rc`.
- `test_probe_source_invents_no_lifecycle_evidence_when_dlopen_fails` — `error` + `stage start`, no
  `finalized`, no `shutdown_resolved`.
- `test_probe_source_keeps_the_original_failure_when_finalize_also_fails` — `error` intact,
  `finalized False`, `finalize_error` set.
- `test_probe_source_does_not_finalize_a_clr_it_never_started` — `pyclr_initialize` raising ⇒ no
  `finalized` key at all (cleanup never claimed).
- `test_probe_source_finalizes_exactly_once` — the stub raises on a second `pyclr_finalize()`, so
  `finalized is True` proves exactly one call.
- `test_probe_source_sends_pythonnets_own_entry_point_arguments` — the stub returns 99 for a wrong
  buffer/size, so `b""`/0 and `b"full_shutdown"`/13 are pinned.

## Local Verification

Environment: Linux (Debian sandbox), Python 3.11.2, pytest 9.1.1, pyflakes 4.0.3, venv
`/home/user/.venv-tnp`. No Windows, no .NET Framework, no PyInstaller, no Portable artifact.

- `python -m pytest tests/test_prompt030_runtime.py -q` → `99 passed, 5 skipped` (5 skips are the
  pre-existing `clr_loader`-not-installed / Windows-only cases).
- Pre-fix reproduction at `0502855` with the same new tests: `16 failed, 73 passed, 5 skipped`, and the
  executed-probe subset alone: `7 failed, 97 deselected`. Post-fix: all pass.
- `python -m pytest tests/test_prompt028_portable.py tests/test_prompt027r_renderer.py
  tests/test_prompt027r_regions.py -q` → 027r suites fully green (25 + 20 passed); 028 reports the same
  `2 failed, 74 passed, 10 skipped` as the base commit — both failures need `pywebview`/`pythonnet`,
  which `requirements-webview.txt` installs only on `sys_platform == "win32"`.
- Full suite: `2 failed, 1191 passed, 15 skipped` vs base `2 failed, 1162 passed, 15 skipped`
  (+29 tests, identical pre-existing failures).
- `python -m compileall -q tools/build_portable.py tests/test_prompt030_runtime.py app/desktop.py` → 0.
- `python -m pyflakes tools/build_portable.py tests/test_prompt030_runtime.py` → 0 findings.
- `git diff --check` → clean.
- Frontend: no frontend file changed (`git diff --stat` lists only the two files above), so no frontend
  test run was performed — reported as not-affected rather than as a pass.

## Deviations from Assessment

One, and it is a refinement rather than a departure: the assessment's stage list treated
`pyclr_close_appdomain` as "n/a (correctly skipped)". The fix keeps it uncalled and adds no evidence
field for it, because upstream `NetFx.shutdown()` guards on a truthy handle (`netfx.py:56`) and
`ClrLoader.CloseAppDomain` no-ops on `IntPtr.Zero` (`ClrLoader.cs:91`), while `pyclr_finalize()` disposes
every registered `DomainData` regardless. Introducing a call or a field for it would have added a code
path the probe can never reach without a named AppDomain, which §6 forbids.

The lifecycle PROMPT-030T proposed in §4 is confirmed correct by the pinned sources; no alternative root
cause had to be reported.

## Follow-ups

- Windows acceptance: rerun
  `.\.venv-build\Scripts\python.exe -m pytest tests/test_prompt030_runtime.py::test_validate_only_passes_a_healthy_folder -vv -s`
  and confirm `probe shutdown_resolved True`, `probe shutdown_rc 0`, `probe finalized True`, subprocess
  returncode 0, and no `GIL must always be released`.
- If Windows still exits nonzero after an ordered shutdown, re-run `$speckit-bug-assess` with the new
  evidence — the next candidates are `PythonEngine.cs:375-383` ("Python error indicator is set") or a
  second undisposed finalizable object, neither of which is addressed here.
- MOTW / `Zone.Identifier` remains an unproven hypothesis for the original `H:\` failure and was
  deliberately not touched.
