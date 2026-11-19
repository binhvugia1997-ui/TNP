# Version bump — Build 025 (v1.4.0)

- **Date**: 2026-11-19
- **Previous build**: 024 (v1.4.0)
- **New build**: 025 (v1.4.0)
- **Trigger**: fix batch for `.specify/bugs/worker-death-on-render-failure/` — Gia approved
  ("Sửa thành Build 025 đi", 19/11 11:38), plus the retry-pass requirement he added
  the same morning (retry error files once after the queue finishes; a first-pass
  error is only a real error if it still fails on retry).

## Exception record

Standing rule (Gia, 2026-11-16): every fix batch bumps the build number. This
bump is the authorized version step for that batch — no specification, assessment,
or plan was re-run for the bump itself (the fix batch it belongs to went through
the bug workflow: assessment → fix → test in
`.specify/bugs/worker-death-on-render-failure/`).

## What changed in this batch (for the bump record)

- `app/batch_processor.py`: per-file `BaseException` guard in the run loop
  (`_error_result()`); `_run_guarded` catches `BaseException` with a
  `CancellationRequested` pass-through; `finally` hardens `result_log.finish()`
  and `_emit_done()`; new `_retry_errors()` — one retry pass over `status ==
  "error"` files only, skipped when Stop/Stop-all is set, final summary counts
  each file once by its last outcome.
- `app/logger.py`: `FileResult.retried` flag; `BatchResultLog.update()` replaces
  the entry for a `source_file` instead of appending a duplicate.
- `tests/test_worker_death_retry.py`: 6 new regression tests (all pass on Linux).

## Bump procedure (mirrors commit dc55817 / Build 024)

1. `app/__init__.py`: `BUILD_NUMBER` 24 → 25 (`__version__` stays "1.4.0");
   docstring examples "Build 024" → "Build 025".
2. All test expectations asserting the current build swept 24 → 25 /
   "024" → "025" / "Build 024" → "Build 025", including the shapes the Build 024
   bump missed:
   - `UpdateInfo("1.4.0", 24, …)` fixtures in `tests/test_application_service.py`
     (the "available update" fixture and its `"build": "024"` assertion).
   - `frontend/index.html` title ("Report Extractor — 1.4.0 — Build 024" →
     "… — Build 025") — this missed file failed the first Windows build of 024.
   - `frontend/tests/` mocks/fixtures (`APP_BUILD`, harness `build: '024'`,
     `buildNumber: 24`, bridge `get_app_version` mocks, update-state fixture).
   - **Miss found by the first Windows run (19/11):** `tests/test_runtime_paths.py`
     lines 158 and 189 (`"version=1.4.0 build=024"`,
     `"STARTUP version=1.4.0 build=024 packaged=true"`) were NOT in the swept
     files — the first repo-wide grep's output was cut at `head -40` and the
     tighter follow-up grep did not match the unquoted `build=024` shape.
     Fixed 19/11 (both → `build=025`); a second full sweep with no truncation
     confirmed no other current-build "024" assertion remains anywhere.
3. Repo-wide sweep afterwards: no remaining current-build "024" assertion in
   `app/`, `tests/`, `frontend/`, `tools/`.
4. Historical "Build 024" references deliberately NOT changed (they name the
   batch that introduced a behaviour): `app/batch_processor.py:166`,
   `app/gui_controller.py:815`, `tests/test_excel_batch.py:217`,
   `tests/test_gui_progress_ollama.py:285`, `frontend/src/state/store.tsx:898`,
   `frontend/tests/prompt034r-report-list-layout.test.mjs:295,346,521`.
   Non-version "024" strings untouched (PROMPT-024, management codes, dependency
   versions such as fast-fifo 1.3.2).

## Verification (Linux)

- Bump-affected files + new tests:
  `.venv-build/bin/python -m pytest tests/test_worker_death_retry.py
  tests/test_auto_update_check.py tests/test_build_and_publish.py
  tests/test_content_learning.py tests/test_gui_redesign.py
  tests/test_gui_responsive.py tests/test_image_learning.py tests/test_prompt001.py
  tests/test_prompt027_rendering.py tests/test_prompt028_portable.py
  tests/test_publish_update.py tests/test_update_progress.py tests/test_updater.py
  tests/test_application_service.py -q` →
  **403 passed, 8 skipped, 1 failed** in 248.57s. The single failure is the
  pre-existing Linux-environment failure
  `test_live_import_probe_reports_every_module_in_this_venv` (pythonnet is not
  installed in the Linux dev venv; the same test passed on Windows in the
  Build 024 run). Not caused by the bump; not weakened.
- Full Linux suite (before the bump): 1413 passed, 13 skipped, 1 failed (same
  environment failure), 590.40s. Full-suite re-run after the bump is covered by
  the Windows full suite in the Build 025 step.

## Not authorized

No merge, no publish, no further bump. Windows full suite + portable build +
PR + deploy are the next steps of this batch, not of this bump.
