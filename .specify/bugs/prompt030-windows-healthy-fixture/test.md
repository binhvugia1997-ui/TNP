# Bug Verification: PROMPT-030 Windows healthy-fixture regression

- **Slug**: prompt030-windows-healthy-fixture
- **Tested**: 2026-10-10
- **Assessment**: ./assessment.md
- **Fix**: ./fix.md
- **Result**: partial

## Summary

The synthetic fixture is now confined to structural and injected-decision tests, and the Windows validate-only success test is configured to stage the genuine installed payloads and run the production live probe. Linux focused and full-suite checks pass, but the original Windows native-loading case was not executed here, so the fix is not verified end-to-end.

## Checks Performed

| Check | Command / Action | Result | Notes |
|-------|------------------|--------|-------|
| Reproduction (post-fix) | `tests/test_prompt030_runtime.py::test_validate_only_passes_a_healthy_folder` | not-run on Windows | Linux skips this intentionally Windows-only test; no claim of native loading is made. The supplied pre-fix Windows failure remains authoritative. |
| New / updated tests | `.venv/bin/python -m pytest -q tests/test_prompt030_runtime.py` | pass | 59 passed, 1 skipped; the skipped case is the Windows-only live integration. |
| Regression suite | `.venv/bin/python -m pytest -q tests/test_prompt028_portable.py tests/test_prompt027r_renderer.py` | pass | 104 passed, 7 skipped. |
| Full Linux suite | `.venv/bin/python -m pytest -q` | pass | 1156 passed, 8 skipped in 112.18s. **LINUX DEVELOPMENT TESTS — not Windows acceptance.** |
| Compile check | `.venv/bin/python -m compileall -q tests/test_prompt030_runtime.py tests/dotnet_image_fixtures.py` | pass | No syntax errors. |
| Lint | `.venv/bin/python -m pyflakes tests/test_prompt030_runtime.py tests/dotnet_image_fixtures.py` | pass | No findings. |

## Output Excerpts

- Focused PROMPT-030: `59 passed, 1 skipped`.
- PROMPT-028/PROMPT-027 regression: `104 passed, 7 skipped`.
- Full Linux suite: `1156 passed, 8 skipped`.
- Local structural parser check found the five expected `pyclr_*` export names in the synthetic image; Linux `probe_packaged_runtime()` explicitly printed `BỎ QUA` because .NET Framework loading is Windows-only.

## Residual Risks

- Real Windows CFFI/CLR loading and `pyclr_*` resolution after this change remain untested. The dedicated Windows integration test must run with installed `pythonnet`, `clr_loader`, CFFI, and the architecture-matched real DLLs.
- The Windows full suite must be rerun after the change is present on the Windows machine. Until then, the reported pre-fix result (`1 failed, 1161 passed`) is not closed.
- Mark-of-the-Web / `Zone.Identifier` remains unproven. The old `H:\ReportExtractor_v1.3.4_Portable` evidence must remain untouched.

## Recommendation

Hold for real Windows acceptance. On the Windows machine, run `python -m pytest tests/test_prompt030_runtime.py::test_validate_only_passes_a_healthy_folder` and the full suite with the installed development runtime, confirm the real packaged probe resolves and calls `pyclr_initialize` through `pyclr_finalize`, then record that evidence. Linux success is not a substitute.
