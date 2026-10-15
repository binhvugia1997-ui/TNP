# PROMPT-039 continuation assessment — final two Windows updater failures

- **Slug**: `prompt039-final-two-windows-failures` (generated for this continuation; new directory)
- **Assessed**: 2026-10-15 (UTC)
- **Checkout**: branch `arena/b9301dab-tnp`, fast-forwarded to existing PROMPT-039 commit `28cf7ef89d161920a049074a6f0b26cbbd4c8242`; no branch switch or new baseline implementation.
- **Reported Windows result**: 2 failed, 1,289 passed in 697.59s.
- **Environment**: Linux, Python 3.11.2; Windows behavior cannot be accepted here.
- **Verdict**: reported failures are credible, but their exact root cause and relationship are **not yet established**. Do not change production updater or tests until the Windows traceback/assertion evidence is available.

## Reproduction performed before any changes

The two requested tests were run separately with verbose output and long tracebacks enabled:

1. `.venv/bin/python -m pytest -vv --tb=long tests/test_image_learning.py::test_28_apply_update_keeps_learning_data_and_rollback_restores_it` → **1 passed** on Linux.
2. `.venv/bin/python -m pytest -vv --tb=long tests/test_updater.py::test_updater_waits_for_main_process_exit` → **1 passed** on Linux.

Both were also rerun separately with a temporary in-process wrapper that forced `validate_staged_portable(..., windows_runtime=True)` for every validation call → **1 passed each**. This verifies that the current synthetic fixtures satisfy the Windows Portable layout, but it does not emulate Windows filesystem/process semantics and does not reproduce either supplied failure. Therefore there is no local traceback/assertion to report for either test.

The Windows summary supplied only the names of the two failed tests and the lines `UPDATE_STAGE_START build=021`, `UPDATE_COPY_DONE bytes=7756`, and `UPDATE_VERIFY_OK`. It does not establish whether stage validation completed, whether `UPDATE_WAIT_EXIT_OK` or `UPDATE_INSTALL_FAILED` occurred, or which assertion/exception failed.

## Lifecycle/code inspection

- The learning test exercises successful apply, preservation of `learning_data`, then a simulated failure at replacement of the second Portable root's `_internal` entry, followed by rollback. The production backup/replace allowlist excludes `learning_data`, and the test checks label/model data after both paths.
- The wait test supplies explicit `alive` and `sleep` callbacks. It checks three alive polls followed by exit, then a zero-timeout case where the process remains alive and no target replacement/backup is allowed. On this test path, the production `process_alive()` OS probe is intentionally bypassed by the injected `alive` callback.
- Code inspection alone does not prove a production lifecycle defect or a test-fixture defect. In particular, the Windows failure at the rollback injection point could depend on path normalization; the wait test has deterministic callback state in the Linux run. Neither hypothesis is proven by the supplied evidence.

## Required evidence before remediation

Please capture each Windows failure independently, including the complete pytest section from the test name through the traceback/assertion diff. Also include the updater log lines around each `UPDATE_VERIFY_OK` through the terminal stage/result, especially whether `UPDATE_STAGE_VALIDATE_OK`, `UPDATE_WAIT_EXIT_OK`, `UPDATE_BACKUP_OK`, `UPDATE_INSTALL_FAILED`, `UPDATE_ROLLBACK_*`, `UPDATE_INSTALL_OK`, and `UPDATE_RESTART*` appear. This will distinguish an extraction/validation problem, rollback-injection problem, wait/poll failure, restart/cleanup issue, or assertion mismatch.

## Proposed remediation

No remediation is authorized by the current evidence. Once the two tracebacks are available:

- classify whether both failures share a root cause or are independent;
- reproduce each separately on Windows if available, or with the narrowest deterministic Windows-specific simulation if not;
- change only the proven production lifecycle defect or the affected fixture/mock;
- retain strict Portable validation, real process waiting, and learning-data preservation.

## Files likely to change after evidence

- `tests/test_image_learning.py` and/or `tests/test_updater.py` if a deterministic test-only path/mocking defect is proven.
- `app/updater.py` only if the trace proves a production update/rollback/process-wait defect.
- This continuation's Spec Kit `fix.md` and `test.md` after a verified remediation.

## Risks & open question

- The Agent environment is Linux and cannot run the requested Windows suite. A local pass (including forced Windows runtime validation) is not Windows acceptance.
- **[NEEDS CLARIFICATION]** What are the complete Windows traceback/assertion outputs and terminal updater log lines for each failure?
