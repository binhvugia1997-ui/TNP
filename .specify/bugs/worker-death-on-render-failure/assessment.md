# Bug Assessment: worker dies silently mid-batch after a PowerPoint render refusal

- **Slug**: worker-death-on-render-failure
- **Created**: 2026-11-19
- **Source**: pasted text + company-machine evidence files (`~/workspace/user/files/app-4.log`, `~/workspace/user/files/batch_result-2.json`)
- **Verdict**: valid
- **Severity**: high

## Report (verbatim or summarized)

Gia ran the Python ReportExtractor Build 024 on the company machine on 2026-11-19,
11:06, queue of 61 reports, Excel on the Desktop. 6 files finished (3 completed,
3 needs_review), then the run stopped at file 7/61 with `WORKER_RECONCILE` at
11:15:31. Log tail (11:15:21):

- `POWERPOINT_ISOLATION state=refused reason=powerpoint-user-session-running`
- `POWERPOINT_OWNERSHIP decision=refused reason=powerpoint-user-session-running`
- `[ERROR] SLIDE_RENDER_POWERPOINT_ONLY_FAILED purpose=after_evidence report=(CTMS)_261009010_... code=powerpoint-user-session-running`

File 7: `261009010-VOC`, `(CTMS)_261009010_Đối sách cải tiến model A185 rear lỗi
bong lệch đuôi ATN, dị vật tape spk veritical, dị vật sơn ngày 08.10 (IQC SEV).pptx`.

`batch_result.json` of that run: 6 results, no `finished_at` in `meta`.
Build 024's worker guard (`_run_guarded` in `app/batch_processor.py`) was supposed
to log `WORKER_FAILED` + deliver the done event on any worker death; neither
happened — the frontend only learned about it from `reconcile()`.

Gia approved the fix as Build 025 (11:38), and added (11:39): after the queue
finishes, the batch retries error files once more (one retry pass, error files
only, never `needs_review`); a first-pass error is NOT yet a real error — only a
file that still fails on retry appears in the final error list (summary +
`batch_result.json` + results table). No retry when Stop / Stop-all was pressed.

## Symptom

A render failure on one file (here: PowerPoint already open on the machine, so
the PR #19 ownership guard refused to attach) kills the whole batch worker
thread silently: no exception, no `WORKER_FAILED`, no done event, 55 files never
run. Expected: the failing file is recorded as an error and the batch continues
with the next file; a worker death is always logged and always delivers done.

## Reproduction

1. On a Windows machine with PowerPoint running (any PowerPoint window open),
   start a batch whose queue contains a file needing a PowerPoint render
   (`after_evidence` / QPN panel).
2. The render is refused with `powerpoint-user-session-running` (11:15:21).
3. The worker thread dies; the UI shows `WORKER_RECONCILE` and the run ends
   (here: 7/61) instead of continuing to file 8.
4. `batch_result.json` has no `finished_at`; the 6 finished files are in the
   Excel, the remaining 55 are never attempted.

[NEEDS CLARIFICATION: the exact exception that killed the thread — see
"Root Cause Hypothesis". The delivered `app-4.log` is a partial slice (see
below); the 11:15:20–11:15:35 segment of the full `app.log` (or `errors.log`)
from the company machine would settle it. The fix below closes every
structural hole that can produce this symptom, so it does not depend on which
exact exception it was.]

## Suspected Code Paths

- `app/batch_processor.py:168` — `BatchProcessor._run_guarded`: catches only
  `Exception`; any `BaseException` escaping `run()` kills the thread with no
  `WORKER_FAILED` and no done event.
- `app/batch_processor.py:210` — `BatchProcessor.run`: the per-file loop calls
  `self._process_one(...)` with no guard of its own; anything escaping
  `_process_one` (or `result_log.add` / `on_batch` after it) aborts the batch.
- `app/batch_processor.py:262` — `run()`'s `finally`: unguarded
  `result_log.finish(...)` (and `_log()` → `on_log`) run before `_emit_done()`;
  a failure there skips done delivery.
- `app/batch_processor.py:398` — `_process_one`: catches `Exception` into an
  error `FileResult`, re-raises only `CancellationRequested` (a plain
  `Exception` subclass, `app/cancellation.py:11`); `BaseException` passes through.
- `app/image_extractor.py:262` — `export_after_pictures`: catches render
  `Exception`, logs `REGION_RENDER ... result=failed`, falls back to embedded
  pictures; a render `RuntimeError` degrades to a review reason, never an error.
- `app/qpn_renderer.py:1303` — `_raise_powerpoint_only`: raises
  `PowerPointUnavailableError(RuntimeError)` (`app/qpn_renderer.py:177`).
- `app/gui_controller.py:1638` — `reconcile()`: fires only when the worker
  thread is dead, the queue is drained and no done was applied; it appends to
  UI `log_lines` only (no `LOG` call — that is why `WORKER_RECONCILE` is absent
  from `app.log`).
- `app/logger.py:137` — `BatchResultLog`: append-only `add()`; no replace/update
  by `source_file` (needed for the retry pass to keep one entry per file).

## Root Cause Hypothesis

**Demonstrated (code + evidence):**

1. The render refusal itself CANNOT kill the batch. `PowerPointUnavailableError`
   is a `RuntimeError`; `export_after_pictures` catches it, logs
   `REGION_RENDER ... result=failed`, falls back to embedded pictures, and
   `_process_one` step 5 catches any remaining render `Exception` into a review
   reason (file becomes `needs_review`, batch continues). The render error is
   fully handled at three layers.
2. The worker thread DID die without delivering done: `reconcile()` only fires
   when `processor.is_running()` is False and no done was applied, and it did
   fire (UI log 11:15:31). `batch_result.json` has no `finished_at`, so
   `run()` never reached `result_log.finish()` in its `finally`.
3. The delivered `app-4.log` is a partial slice, not the full log: it is
   658,791 bytes / 2,497 lines (longest line 746 B) while the `app.log` on the
   company machine is 517,325 KB. Lines are missing in the 11:15:21–11:27:36
   window — proven by the absent mandatory follow-up line: after the
   `SLIDE_RENDER_POWERPOINT_ONLY_FAILED` ERROR, `export_after_pictures` MUST
   log `REGION_RENDER ... result=failed`, but `app-4.log` contains only the six
   `REGION_RENDER ... result=ok` lines of files 1–6. So "no `WORKER_FAILED` /
   no traceback in the log" is NOT established from this file.

**Hypothesis (medium confidence):** the thread was killed by a `BaseException`
(SystemExit / KeyboardInterrupt / GeneratorExit, or a `BaseException` raised by
a dependency / COM callback) escaping `run()`, OR by a failure in the
unprotected part of `run()`'s `finally` (`result_log.finish()` /
`_log()` → `on_log`) before `_emit_done()`. Both paths bypass the Build 024
guard exactly as observed: CPython's thread bootstrap swallows a
`BaseException` to stderr (`threading.Thread._bootstrap_inner` uses a bare
`except:` → `_invoke_excepthook`), which is invisible in the packaged pywebview
app — the "silent death". The exact exception is unverified pending the full
log segment (see Open Questions).

**Structural holes (demonstrated in code, fixed regardless of which exception
it was):**

- H1 — `_run_guarded` catches `Exception` only (`batch_processor.py:168`).
- H2 — `run()`'s `finally` calls unguarded `result_log.finish()` before
  `_emit_done()` (`batch_processor.py:262`).
- H3 — no per-file guard around `_process_one(...)` in the run loop
  (`batch_processor.py:252`).
- H4 — first-pass errors are final: no retry; a transient cause (e.g. a
  PowerPoint left open, closed mid-run) still reports the file as failed.

## Proposed Remediation

**Preferred:**

1. Per-file guard in the `run()` loop: wrap `self._process_one(...)` (and the
   `result_log.add` / `on_batch` epilogue) in `try/except BaseException` →
   build an error `FileResult` (status `error`, `error="Type: message"`),
   append to `self.results` + `result_log`, continue with the next file.
   `CancellationRequested` still propagates (cancellation contract, VI).
2. Widen `_run_guarded` to `except BaseException`: log `WORKER_FAILED`,
   count unaccounted files as failed, deliver done exactly once. A failing
   `on_done` inside the handler is already guarded.
3. Harden `run()`'s `finally`: wrap `result_log.finish()` and the final
   `_emit_done()` in their own `try/except` so done is always delivered even
   when the final save/finish fails.
4. One retry pass, inside the same run, after the first-pass loop and before
   the `finally` (so the final save + `result_log.finish()` + done cover the
   retry results):
   - Snapshot the files whose first-pass status is `error` only (never
     `needs_review` / `not_written` / `skipped` / `cancelled`).
   - Skip the retry pass entirely when Stop / Stop-all was requested (check
     before the pass and before each retried file, same boundaries as the
     main loop). Exactly one pass — no loop back.
   - Re-process each error file via `_process_one` (fresh `seq`, same
     `idx`/`path`); replace its entry in `self.results` and in
     `batch_result.json` (new `BatchResultLog.update()` keyed by
     `source_file`); undo its first-pass accounting first (`failed -= 1`,
     remove its entry from `summary.errors`) so the final summary counts
     each file exactly once, by its LAST outcome.
   - A file that succeeds on retry is `completed`/`needs_review` and is NOT
     in the final error list; a file that still fails stays `error` with the
     retry's error message. Done is delivered only after the retry pass.
   - Mark retried outcomes (`FileResult.retried = True`) so the retry is
     visible in `batch_result.json` / the review report.

**Alternatives:**

- Widen only `_run_guarded` (fix 2): the batch would deliver done but still
  end early at the failed file — all remaining files unattempted. Rejected:
  does not fix the user-visible symptom.
- Pre-flight PowerPoint check before each render: avoids this one trigger but
  leaves the general silent-death hole open. Rejected as the sole fix.
- Retry the whole batch on failure: wasteful (re-runs completed files) and
  still needs the per-file guard. Rejected.

**Files likely to change**:

- `app/batch_processor.py` (guards + retry pass)
- `app/logger.py` (`BatchResultLog.update`, `FileResult.retried`)
- `tests/test_worker_death_retry.py` (new)
- `app/__init__.py` + test expectations + `.specify/version-bump-build-025/note.md`
  (Build 025 bump, per the standing every-batch rule; same procedure as
  `.specify/version-bump-build-024/note.md`)

**Tests to add or update**:

- Render/base-exception failure in one file → file becomes `error`, batch
  continues, done delivered (new).
- `BaseException` escaping `run()` outside the loop → `WORKER_FAILED` logged
  + done delivered (new).
- Retry: file errors on pass 1, succeeds on retry → absent from final error
  list; `batch_result.json` holds one entry with the success (new).
- Retry: file errors on both passes → in final error list with the retry's
  error (new).
- Retry: `needs_review` file is never retried (new).
- Retry: Stop requested before the retry pass → no retry, first-pass error
  stands (new).
- Retry: done event delivered only after the retry pass completes (new).
- Full suite (Linux + Windows) must stay green; no existing test weakened.

## Risks & Considerations

- Constitution V (Excel transaction safety): retry re-processes a file whose
  first pass may already have created a row — the existing
  `MASTER_NEW_RETRY` / `_row_looks_created_by_us` path reuses the partial row
  instead of creating a second one; the retry tests must cover a failed file
  that created a row. No double-commit.
- Constitution VI (cancellation): retry must never run after Stop / Stop-all;
  `CancellationRequested` still propagates out of `_process_one` and out of
  the per-file guard.
- Constitution I/II (evidence semantics): retry does not change what is
  extracted; it only re-runs the same extraction. No LLM text enters Excel.
- Constitution IV (report isolation): retry touches only the retried file's
  own result entry; `update()` replaces by `source_file`.
- The `finally` hardening must not mask a real save failure: failures are
  still logged (`Final save failed`, `result_log.finish failed`).
- Catching `BaseException` per file must not swallow `KeyboardInterrupt` on
  the main thread: the guard lives only in the daemon worker thread, where
  SIGINT cannot be delivered; `CancellationRequested` (an `Exception`) is
  re-raised before the `BaseException` handler.

## Open Questions

- [NEEDS CLARIFICATION: which exact exception killed the 19/11 worker? The
  full `app.log` segment 11:15:20–11:15:35 (or `errors.log`) from the company
  machine would settle it. The fix closes all demonstrated holes, so it does
  not block on this answer — but the answer would confirm whether any
  additional hole exists.]
- The company machine runs the app from
  `H:\ReportExtractor_1.3.7\ReportExtractor_v1.3.7_Portable\` (folder name
  v1.3.7, binary Build 024); its logs live next to the output Excel
  (`Desktop\logs\`), not next to the exe — worth confirming with Gia when
  asking for the log segment.
