# Bug Verification: PROMPT-039 Windows full-suite integration regressions

- **Slug**: `prompt039-windows-full-suite-regression`
- **Tested**: 2026-10-15 (UTC)
- **Assessment**: ./assessment.md
- **Fix**: ./fix.md
- **Result**: partial

## Summary

The shared fixtures now satisfy the updater's Windows-only Portable contract, while forced-Windows tests prove missing runtime DLLs remain rejected. The changed renderer and freshness regressions are deterministic, and the full available Linux suite passes; the original Windows acceptance was not rerun because this Agent environment is Linux, so Windows acceptance remains pending.

## Checks Performed

| Check | Command / Action | Result | Notes |
|-------|------------------|--------|-------|
| Supplied Windows reproduction | User-reported `18 failed, 1262 passed`; forced-Windows validator exercise against the old shared directory and ZIP fixtures | pass (root cause reproduced) | Both pre-fix fixtures returned `thiếu .NET/pythonnet runtime bắt buộc` when `windows_runtime=True`; no real Windows run was available. |
| Updater | `.venv/bin/python -m pytest tests/test_updater.py -q` | pass | 62 passed, including stage pass-through and missing-runtime rejection. |
| Update progress | `.venv/bin/python -m pytest tests/test_update_progress.py -q` | pass | 12 passed. |
| Content learning | `.venv/bin/python -m pytest tests/test_content_learning.py -q` | pass | 20 passed. |
| Image learning | `.venv/bin/python -m pytest tests/test_image_learning.py -q` | pass | 28 passed. |
| Renderer | `.venv/bin/python -m pytest tests/test_prompt027_rendering.py -q` | pass | 54 passed. Failure logging and builtin fallback are asserted with injected state. |
| Freshness | `.venv/bin/python -m pytest tests/test_prompt038_desktop_freshness.py -q` | pass | 7 passed; includes explicit source mtimes, Vite `index.html`, and equal-time tie-break. |
| PowerPoint availability and ownership | `.venv/bin/python -m pytest tests/test_prompt027r_renderer.py tests/test_prompt035_powerpoint_isolation.py -q` | pass | 49 passed; includes deterministic unavailable and ownership-refusal/fallback tests. |
| Strict Portable negatives | `.venv/bin/python -m pytest tests/test_prompt028_portable.py tests/test_updater.py::test_stage_update_rejects_windows_packages_missing_runtime_dlls tests/test_updater.py::test_staged_portable_rejects_missing_required_windows_runtime_files -q` | pass | 90 passed, 3 skipped. Missing Python.Runtime, all ClrLoader variants, and required WebView2 files are rejected. |
| Full regression suite | `.venv/bin/python -m pytest -q` | pass | **1,287 passed, 4 skipped** in 102.15 seconds. Linux development evidence only. |
| Changed-Python compile | `.venv/bin/python -m compileall -q app/desktop.py tests/portable_fixtures.py tests/test_updater.py tests/test_prompt027_rendering.py tests/test_prompt038_desktop_freshness.py` | pass | All changed Python files compile. |
| Changed test/helper lint | `.venv/bin/python -m pyflakes tests/portable_fixtures.py tests/test_updater.py tests/test_prompt027_rendering.py tests/test_prompt038_desktop_freshness.py` | pass | No warnings in changed test/helper modules. |
| Diff whitespace | `git diff --cached --check` | pass | Checked after all implementation and Spec Kit report files were staged. |
| Windows full-suite acceptance | `.\.venv-build\Scripts\python.exe -m pytest -q` | not-run | No Windows host in this environment. Run this exact command on Windows before declaring acceptance. |

## Output Excerpts

```text
pre-fix current fixture forced-Windows validation: (False, 'thiếu .NET/pythonnet runtime bắt buộc')
pre-fix release ZIP forced-Windows validation: (False, 'thiếu .NET/pythonnet runtime bắt buộc')

62 passed (tests/test_updater.py)
12 passed (tests/test_update_progress.py)
20 passed (tests/test_content_learning.py)
28 passed (tests/test_image_learning.py)
54 passed (tests/test_prompt027_rendering.py)
7 passed (tests/test_prompt038_desktop_freshness.py)
49 passed (PowerPoint renderer + ownership suites)
90 passed, 3 skipped (strict Portable validation group)
1287 passed, 4 skipped (full Linux suite)
```

## Residual Risks

- The Windows full suite has not been run on a Windows host. Linux forced-branch tests validate the updater's Windows runtime logic but do not replace the requested Windows integration acceptance.
- Four tests are skipped in the Linux full run; the Windows command must report its own result. Do not treat the user-reported 18 as 18 independent production defects: the updater group was a shared fixture cascade, with separate renderer and freshness test nondeterminism.

## Recommendation

Ready for the Windows full-suite retest. Run `.\.venv-build\Scripts\python.exe -m pytest -q`; expected acceptance is 0 failed. If any test remains failing, capture its exact name and traceback before changing production code.
