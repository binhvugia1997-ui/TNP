# Bug Fix: worker dies silently mid-batch after a PowerPoint render refusal

- **Slug**: worker-death-on-render-failure
- **Fixed**: 2026-11-19
- **Assessment**: ./assessment.md
- **Status**: applied

## Summary

One file's failure can no longer kill the batch, and the worker thread can no
longer die silently: per-file failures (including `BaseException`) become an
error result for that file only and the batch continues; the worker guard now
catches `BaseException` and always delivers the done event; the run's `finally`
cannot skip done delivery anymore. Error files get exactly one retry pass
before done/final save; a first-pass error is provisional — only a file that
still fails on retry is a real error in the final results.

## Changes

| File | Change | Notes |
|------|--------|-------|
| `app/batch_processor.py` | modified | Per-file `BaseException` guard in the run loop; `_run_guarded` catches `BaseException` (+ explicit `CancellationRequested` pass-through); `finally` hardens `result_log.finish()` / `_emit_done()`; new `_error_result()` helper; new `_retry_errors()` one-pass retry |
| `app/logger.py` | modified | `FileResult.retried` flag; `BatchResultLog.update()` replaces the entry for a `source_file` (retry keeps one entry per file) |
| `tests/test_worker_death_retry.py` | added | 6 regression tests (see below) |
| `.specify/bugs/worker-death-on-render-failure/fix.md` | added | this report |
| `.specify/bugs/worker-death-on-render-failure/assessment.md` | added | assessment (assess phase) |

## Diff Highlights

Per-file guard — the loop no longer lets one file abort the batch:

```python
try:
    fr = self._process_one(idx, Path(path), seq, writer, renderer, client, history, assets)
    self.results.append(fr)
    result_log.add(fr)
    done += 1
    self.on_batch(done, total)
except CancellationRequested:
    raise                       # cooperative cancel is never a file error (VI)
except BaseException as exc:    # one file must never kill the batch (Build 025)
    fr = self._error_result(idx, Path(path), exc)
    ...
```

Retry pass — after the loop, before the `finally`:

```python
seq = self._retry_errors(queue, writer, renderer, client, history, assets, result_log, seq)
```

`_retry_errors` snapshots `status == "error"` files only, skips entirely when
Stop / Stop-all is set (and records `stopped`, like the main loop), undoes each
retried file's first-pass accounting (`failed -= 1`, drop its `summary.errors`
entry) before re-processing, replaces the result in `self.results` and via
`result_log.update()`, marks it `retried = True`. Exactly one pass.

## Tests Added or Updated

- `tests/test_worker_death_retry.py::test_base_exception_in_one_file_becomes_error_and_batch_continues` — a `BaseException` from one file's parse becomes an `error` result; the other files finish; done delivered once.
- `tests/test_worker_death_retry.py::test_base_exception_outside_loop_logs_worker_failed_and_delivers_done` — a `BaseException` escaping `run()` is caught by `_run_guarded`: `WORKER_FAILED` in the log, done delivered, unaccounted files counted failed.
- `tests/test_worker_death_retry.py::test_retry_success_removes_file_from_final_errors` — file errors on pass 1, succeeds on retry: absent from final errors, one `batch_result.json` entry with the success, done is the last event.
- `tests/test_worker_death_retry.py::test_retry_still_failing_stays_in_final_errors_with_retry_error` — file errors on both passes: stays `error` with the RETRY error message.
- `tests/test_worker_death_retry.py::test_needs_review_is_never_retried` — a `needs_review` file is processed once, never retried.
- `tests/test_worker_death_retry.py::test_no_retry_when_stop_requested_before_retry_pass` — Stop during the last file skips the retry pass; the first-pass error stands.

No existing test was modified or weakened.

## Local Verification

- Commands run: `.venv-build/bin/python -m pytest tests/test_worker_death_retry.py -q` → 6 passed (Linux, 12.21s).
- Full Linux suite: running at fix-report time; result recorded in `test.md`.
- Windows full suite + portable build: pending (ollama); result recorded in `test.md`.

## Deviations from Assessment

- None in the fix itself. One addition during implementation: when the retry
  pass is skipped because Stop / Stop-all arrived while the last report was
  still running, `_retry_errors` now sets `summary.stopped = True` (and
  `cancel_requested` for cancel), mirroring the main loop — otherwise the
  summary would claim the run was not stopped even though the retry was
  suppressed by the user's stop. Covered by
  `test_no_retry_when_stop_requested_before_retry_pass`.

## Follow-ups

- The exact exception that killed the 19/11 worker is still unverified (the
  delivered `app-4.log` is a partial slice; see assessment Open Questions).
  Ask Gia for the full `app.log` 11:15:20–11:15:35 segment (or `errors.log`)
  from the company machine if the cause needs confirmation beyond the closed
  structural holes.
- Gia live test on the company machine: close ALL PowerPoint windows first,
  then Start — the 6 finished files are already in the Excel, the run
  continues with the remaining 55.
