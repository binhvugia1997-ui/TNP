# Bug Fix: PROMPT-039 final two Windows updater failures

- **Slug**: `prompt039-final-two-windows-failures`
- **Fixed**: 2026-10-15 (UTC)
- **Assessment**: `./assessment.md` (left unchanged, as requested)
- **Branch**: `arena/b9301dab-tnp`, continuing the existing PROMPT-039 branch from `28cf7ef89d161920a049074a6f0b26cbbd4c8242`
- **Status**: test-fixture fix applied; Windows acceptance remains pending

## Root cause and classification

The two supplied Windows tracebacks are manifestations of **one shared test-fixture root cause**, not independent updater defects. In both tests, the second `_staged()` call starts an extraction beneath pytest's long temporary root, updater staging directories, and a ZIP wrapper directory. `_extract_archive()` reaches `os.open()` with a target path beyond the classic Windows `MAX_PATH` boundary. Windows reports `FileNotFoundError`, which `stage_update()` wraps as `RuntimeError: Cập nhật thất bại ở bước chuẩn bị (FileNotFoundError).`

| Test | Failing second-stage target reported in the Windows traceback | Length | Intended scenario not reached |
|------|---------------------------------------------------------------|--------|--------------------------------|
| `tests/test_image_learning.py::test_28_apply_update_keeps_learning_data_and_rollback_restores_it` | `_internal\\webview\\lib\\Microsoft.Web.WebView2.WinForms.dll` | 260 characters | The injected `_internal` replacement failure and rollback assertions |
| `tests/test_updater.py::test_updater_waits_for_main_process_exit` | `_internal\\webview\\lib\\runtimes\\win-x64\\native\\WebView2Loader.dll` | 265 characters | The second, never-exits / zero-timeout wait assertions |

The logs reach `UPDATE_VERIFY_OK` but not `UPDATE_EXTRACT_OK`, matching failure during archive extraction. The first update and learning-data assertions in the learning test had already succeeded; its later rollback setup failed before rollback ran. The wait test's first poll/exit scenario had already succeeded; its second stage failed before the never-exits case called `apply_update()`. Thus neither failure demonstrates a problem in learning-data preservation, rollback, Windows process detection, wait timing, or PID handling.

## Changes

| File | Change | Rationale |
|------|--------|-----------|
| `tests/test_image_learning.py` | Use `_staged(tmp_path / "r", monkeypatch, top=False)` for the second, rollback-only setup. | Shortens the fixture root and removes the ZIP wrapper component from every extracted target. Existing preservation and rollback assertions remain intact. |
| `tests/test_updater.py` | Use `_staged(tmp_path / "s", monkeypatch, top=False)` for the second, never-exits setup. | Applies the same path-length correction only to the setup that failed before the wait assertions. Existing wrapper-folder package coverage remains. |

Root-level update ZIPs are supported by `_extract_archive()` and explicitly covered by `test_valid_zip_without_top_folder_passes`; wrapper-folder archives continue to be covered by the other updater package tests. No validation was weakened, and no production code was changed.

## Lifecycle and contract review

- **Stage → verify → extract:** package copy and verification complete; extraction fails at `os.open()` before `UPDATE_EXTRACT_OK`. The shorter fixture root and root-level archive layout reduce the destination path without changing extraction safety checks.
- **Learning data → apply → rollback:** the successful first apply and learning-data assertions remain unchanged. The second setup still creates data only beneath pytest's temporary directory, injects the same `_internal` replace failure, and asserts the old app is restored while labels/model data remain. `app/updater.py` and its preservation allowlist are unchanged.
- **Wait / PID / timeout:** the existing test injects deterministic callbacks rather than relying on a real external process: its first `alive(pid)` mock returns true for three polls and then false, recording three sleeps; the second returns true with `wait_timeout=0.0` and checks that the app is not replaced and no backup is created. `_staged()` supplies a positive `old_pid` different from the helper PID, as required by the staged-manifest contract. `wait_for_exit()` validates the PID, bounds timeout to 0–90 seconds, polls with a monotonic deadline, and caps the interval at 0.5 seconds. The test's injected `alive` callback bypasses the OS probe; production Windows `process_alive()` was inspected and is unchanged: it uses `OpenProcess` / `GetExitCodeProcess`, treats `STILL_ACTIVE` as alive, and only `ERROR_INVALID_PARAMETER` as proof that the PID is gone, failing closed on ambiguous query errors. The mocked `spawn` returns `None`, which the helper treats as a test adapter rather than a real child process. The failure occurred before the second `apply_update()` call. No waiting was removed; Windows API behavior itself still requires Windows retest.

## Files changed for this continuation

- `tests/test_image_learning.py`
- `tests/test_updater.py`
- `.specify/bugs/prompt039-final-two-windows-failures/assessment.md` (existing assessment retained as-is)
- `.specify/bugs/prompt039-final-two-windows-failures/fix.md`
- `.specify/bugs/prompt039-final-two-windows-failures/test.md`

## Safety

This is a test-only path-layout change. It does not touch a user's installation, `learning_data`, release artifacts, runtime validation, process waiting, rollback logic, backups, or cleanup behavior. The targeted tests continue to assert data preservation, rollback restoration, wait timeout handling, and no backup/replacement on timeout.
