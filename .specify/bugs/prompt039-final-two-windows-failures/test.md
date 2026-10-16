# Bug Verification: PROMPT-039 final two Windows updater failures

- **Slug**: `prompt039-final-two-windows-failures`
- **Tested**: 2026-10-15 (UTC)
- **Assessment**: `./assessment.md` (unchanged)
- **Fix**: `./fix.md`
- **Branch**: `arena/b9301dab-tnp`
- **Result**: Linux verification passed; Windows acceptance pending

## Original evidence and reproduction status

The user-provided Windows log (`prompt039-final2-windows.txt`, UTF-16 LE) was decoded and reviewed in full. The two failures were captured separately in that log: the learning-data test's second staging setup fails at `os.open()` for a 260-character WebView2 WinForms DLL path, and the wait test's second staging setup fails at `os.open()` for a 265-character WebView2 loader DLL path. Both are wrapped as `RuntimeError: Cập nhật thất bại ở bước chuẩn bị (FileNotFoundError).` The log reaches `UPDATE_VERIFY_OK`, not `UPDATE_EXTRACT_OK`.

This agent runs on Linux, so it cannot reproduce Windows `MAX_PATH` behavior directly. Before the fixture edit, both exact tests passed separately on Linux, including separate runs that forced `validate_staged_portable(..., windows_runtime=True)`. These results did not establish Windows filesystem acceptance. The cause was established by the supplied Windows tracebacks and the exact over-limit target paths, not inferred from Linux passes.

## Checks after the fixture fix

| Check | Command / action | Result | Notes |
|-------|------------------|--------|-------|
| Learning-data + rollback target | `.venv/bin/python -m pytest -q tests/test_image_learning.py::test_28_apply_update_keeps_learning_data_and_rollback_restores_it` | **1 passed** | Run separately; 0.38s. |
| Wait target | `.venv/bin/python -m pytest -q tests/test_updater.py::test_updater_waits_for_main_process_exit` | **1 passed** | Run separately; 0.66s. |
| Learning-data target with forced Windows Portable validation | In a separate Python process, wrapped `validate_staged_portable` to call the original with `windows_runtime=True`, then ran the exact pytest node | **1 passed** | 0.29s. Exercises the strict Windows runtime validation branch, not Windows path semantics. |
| Wait target with forced Windows Portable validation | Same wrapper in a separate Python process, then ran the exact pytest node | **1 passed** | 0.30s. Wait behavior still uses the test's deterministic `alive` and `sleep` callbacks. |
| Related updater and learning regressions | `.venv/bin/python -m pytest -q tests/test_updater.py tests/test_image_learning.py` | **90 passed** | 4.22s. |
| Full available suite | `.venv/bin/python -m pytest -q` | **1,287 passed, 4 skipped** | 108.03s, Linux host. |
| Diff whitespace | `git diff --check` | **passed** | Run after the source edits. |

## Contract and safety checks

- The learning test still checks that label/model/thumbnails survive the successful update and that labels/model survive rollback while the old executable is restored.
- The wait test still checks three alive polls then exit, then a zero-timeout never-exits case with no replacement and no backup. It supplies a positive non-helper PID, deterministic `alive`/`sleep` callbacks, and a `spawn` adapter returning `None`; it does not depend on real process timing. `wait_for_exit()` uses validated PID input, a bounded timeout, monotonic-deadline polling, and capped intervals. Windows `process_alive()` was inspected (Win32 `OpenProcess` / `GetExitCodeProcess`, `STILL_ACTIVE`, and fail-closed handling of ambiguous query errors) but is bypassed by the injected `alive` callback and cannot be OS-exercised on this Linux host. Production wait and process-probe code remains unchanged.
- The root-level ZIP layout is supported and has dedicated coverage in `test_valid_zip_without_top_folder_passes`; wrapper-folder ZIP tests remain in the updater suite.
- Only test fixture path/layout and this bug's Spec Kit follow-up reports changed in this continuation. `app/updater.py`, the pre-existing `assessment.md`, user data, release artifacts, and production Windows validation were not changed.

## Windows acceptance

**Not run:** this environment is Linux. Forced Windows runtime validation covers the validator branch only; it is not a substitute for Windows filesystem/process behavior. Retest both exact tests on Windows, then the Windows suite, before declaring Windows acceptance.

**Status: READY FOR WINDOWS RETEST.**
