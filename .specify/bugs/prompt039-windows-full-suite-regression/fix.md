# Bug Fix: PROMPT-039 Windows full-suite integration regressions

- **Slug**: `prompt039-windows-full-suite-regression`
- **Fixed**: 2026-10-15 (UTC)
- **Assessment**: ./assessment.md
- **Status**: applied

## Summary

Corrected the shared synthetic Portable fixtures to model the strict Windows runtime layout, made the renderer and freshness regression tests deterministic, and made equal-mtime freshness diagnostics stable. Production Portable validation and PowerPoint ownership behavior were not weakened or changed.

## Changes

| File | Change | Notes |
|------|--------|-------|
| `tests/portable_fixtures.py` | added canonical helper | Supplies one structural-only synthetic Windows runtime layout to directory and ZIP fixture builders; explicitly non-loadable placeholders. |
| `tests/test_updater.py` | modified shared builders and added tests | Both `make_portable()` and `make_release_zip()` now include the required runtime paths. Added forced-Windows acceptance, stage-path pass-through, and missing Python.Runtime/ClrLoader/WebView2 rejection coverage. |
| `tests/test_prompt027_rendering.py` | modified renderer regression | Injects an available PowerPoint probe and a deterministic staged failure; asserts one diagnostic, purpose/backend/stage/reason, and builtin fallback. |
| `app/desktop.py` | modified freshness diagnostics | Source candidates are ordered by normalized relative path; equal mtimes consistently select the lexical-first path. The strict `source_mtime > dist_mtime` stale rule is unchanged. |
| `tests/test_prompt038_desktop_freshness.py` | modified freshness fixtures | Sets source `src/index.css`, source `index.html`, and built `dist/index.html` mtimes explicitly using whole-second values; tests source-file ordering, `index.html` participation, and tie behavior. |
| `.specify/bugs/prompt039-windows-full-suite-regression/assessment.md` | added | Root-cause assessment and remediation contract. |
| `.specify/bugs/prompt039-windows-full-suite-regression/fix.md` | added | This implementation report. |
| `.specify/bugs/prompt039-windows-full-suite-regression/test.md` | added | Verification evidence and Windows retest status. |

## Diff Highlights

- The shared ZIP and directory builders now use the same path map for `Python.Runtime.dll`, both hook-provided `ClrLoader.dll` architecture variants, WebView2 Core/WinForms, and `WebView2Loader.dll`.
- Tests force the production validator's Windows branch with `windows_runtime=True`. Stage-level tests also run the actual updater staging flow against that branch, including ZIPs with a recomputed valid checksum but missing required runtime DLLs.
- Freshness still considers the Vite entry `frontend/index.html`; it was not excluded. If source timestamps tie, the lexically first normalized source-relative path is reported, and freshness still requires a strictly newer source timestamp.

## Tests Added or Updated

- `tests/test_updater.py::test_shared_portable_fixtures_match_the_windows_runtime_contract` — both directory and extracted ZIP fixtures satisfy the strict Windows contract.
- `tests/test_updater.py::test_stage_update_passes_the_windows_runtime_gate_with_shared_fixture` — update staging proceeds through Windows runtime validation.
- `tests/test_updater.py::test_stage_update_rejects_windows_packages_missing_runtime_dlls` — valid-checksum packages missing Python.Runtime or all ClrLoader DLLs fail stage validation.
- `tests/test_updater.py::test_staged_portable_rejects_missing_required_windows_runtime_files` — missing Python.Runtime, ClrLoader, each required WebView2 DLL are rejected.
- `tests/test_prompt027_rendering.py::test_backend_failure_is_logged_not_silently_swallowed` — deterministic injected failure remains logged and falls back.
- `tests/test_prompt038_desktop_freshness.py` — explicit mtime cases for fresh/stale bundles, Vite `index.html`, and equal-time path tie-breaking.

## Local Verification

Linux sandbox only; this is development evidence, not Windows acceptance.

- `.venv/bin/python -m pytest tests/test_updater.py -q` → **62 passed**.
- `.venv/bin/python -m pytest tests/test_update_progress.py -q` → **12 passed**.
- `.venv/bin/python -m pytest tests/test_content_learning.py -q` → **20 passed**.
- `.venv/bin/python -m pytest tests/test_image_learning.py -q` → **28 passed**.
- `.venv/bin/python -m pytest tests/test_prompt027_rendering.py -q` → **54 passed**.
- `.venv/bin/python -m pytest tests/test_prompt038_desktop_freshness.py -q` → **7 passed**.
- `.venv/bin/python -m pytest tests/test_prompt027r_renderer.py tests/test_prompt035_powerpoint_isolation.py -q` → **49 passed** (including deterministic unavailable and ownership-refusal paths).
- `.venv/bin/python -m pytest tests/test_prompt028_portable.py tests/test_updater.py::test_stage_update_rejects_windows_packages_missing_runtime_dlls tests/test_updater.py::test_staged_portable_rejects_missing_required_windows_runtime_files -q` → **90 passed, 3 skipped**.
- `.venv/bin/python -m pytest -q` → **1,287 passed, 4 skipped**.
- `compileall` on every changed Python file → passed.
- `pyflakes` on the changed test/helper files → passed.
- `git diff --cached --check` → passed after all implementation and Spec Kit report files were staged.

The ignored local `.venv` was provisioned with the dev/webview test dependencies so the complete suite could run. No dependency manifest or frontend file was changed.

## Deviations from Assessment

- Added end-to-end updater staging tests with a forced Windows validator in addition to direct structural validation tests. This confirms the shared fixtures pass the same stage gate that caused the reported cascade and that correctly hashed packages missing Python.Runtime/ClrLoader still fail there.
- Deterministic freshness tie-breaking was implemented in production because the existing unsorted `rglob()` walk made the diagnostic source path depend on enumeration order when mtimes tied. The stale decision and source-input set remain unchanged.
- No production Portable validator or PowerPoint renderer change was needed.

## Follow-ups

- Run the Windows acceptance command after this branch is available on the target machine: `.\.venv-build\Scripts\python.exe -m pytest -q`.
- Linux results do not establish the requested Windows full-suite acceptance; the Windows result remains pending.
