# Bug Fix: PROMPT-030 Windows healthy-fixture regression

- **Slug**: prompt030-windows-healthy-fixture
- **Fixed**: 2026-10-10
- **Assessment**: ./assessment.md
- **Status**: applied

## Summary

Separated structural PE fixtures from live Windows runtime validation. The Windows validate-only success case now stages the genuine `Python.Runtime.dll` and architecture-matched `ClrLoader.dll` discovered through the exact interpreter validate-only will use; the synthetic fixture is retained only for structural and injected-decision tests. No production/application source was changed.

## Changes

| File | Change | Notes |
|------|--------|-------|
| `tests/test_prompt030_runtime.py` | modified | Added exact-interpreter DLL discovery and deterministic byte staging for the Windows-only live test. Added explicit Windows skip rationale and a structural validate-only decision test with per-test injected probe results. The structural rejection test also injects the probe so no Windows loader sees a synthetic DLL. |
| `tests/dotnet_image_fixtures.py` | modified | Clarified that `clr_loader_image()` models structural PE/export metadata and placeholder RVAs only; it does not implement a loadable/callable native mixed-mode DLL. |

## Diff Highlights (optional)

- `test_validate_only_passes_a_healthy_folder` now requires Windows and uses real files found by the same selected interpreter as the probe. Missing `Python.Runtime.dll`, matching `ClrLoader.dll`, or CFFI is an explicit test failure.
- Linux keeps structural/validator-decision coverage via `test_validate_only_obeys_injected_probe_result_for_structural_fixture`; its injected result is scoped to that test and does not claim native execution.
- The original synthetic fixture remains available to the parser, metadata, export-name, identity, architecture, and malformed-image cases.

## Tests Added or Updated

- `tests/test_prompt030_runtime.py::test_validate_only_passes_a_healthy_folder` — Windows integration using genuine installed payloads and the production live probe.
- `tests/test_prompt030_runtime.py::test_validate_only_obeys_injected_probe_result_for_structural_fixture` — parametrized success/failure decision contract without native-loadability claims.
- `tests/test_prompt030_runtime.py::test_validate_only_reuses_the_same_gate_and_reports_a_real_cause` — explicitly injects the live-probe result so the structural negative fixture is never sent to Windows loading.

## Local Verification

- Commands run: `.venv/bin/python -m pytest -q -vv tests/test_prompt030_runtime.py` → **59 passed, 1 skipped** (the Windows-only live integration is skipped on Linux).
- Commands run: `.venv/bin/python -m pytest -q -vv tests/test_prompt028_portable.py tests/test_prompt027r_renderer.py` → **104 passed, 7 skipped**.
- Environment: Linux, Python 3.11.2. The focused PROMPT-028/PROMPT-027 run initially exposed missing development/runtime packages in the clean Agent environment; after installing the declared webview development requirements plus `pythonnet==3.0.5` and `clr-loader==0.2.10` into the ignored local `.venv`, the rerun passed. No dependency files were changed.
- Windows native DLL loading was not run in this environment; the live integration still requires the real Windows retest.

## Deviations from Assessment

None. The implementation stays in the two assessed test files; production runtime code and the test payload rules are unchanged.

## Follow-ups

- Rerun `test_validate_only_passes_a_healthy_folder` and the full suite on the actual Windows machine after this change is present there. Do not treat the Linux run as Windows acceptance.
- Keep the real `H:\ReportExtractor_v1.3.4_Portable` folder and its ADS evidence untouched. MOTW/Zone.Identifier remains unproven.
