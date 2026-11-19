"""Build 025: worker-death-on-render-failure — per-file isolation, worker guard, retry pass.

Regression coverage for `.specify/bugs/worker-death-on-render-failure/`:

- A `BaseException` escaping one file's processing (the 19/11 silent death:
  render refusal `powerpoint-user-session-running` on 261009010-VOC killed the
  worker with no WORKER_FAILED and no done event) becomes an error result for
  that file only; the batch continues and done is delivered.
- A `BaseException` escaping `run()` outside the per-file loop is caught by
  `_run_guarded`: WORKER_FAILED is logged and done is delivered.
- Retry pass (Gia, 19/11): error files are retried exactly once, after the
  queue finishes and before done/final save. A first-pass error is provisional:
  only a file that still fails on retry is a real error in the final summary /
  batch_result.json. `needs_review` is never retried; no retry after Stop.

Headless tests (real BatchProcessor + real workbook); PowerPoint COM is stubbed
on Linux. Windows / PowerPoint-COM acceptance still needs the manual pass.
"""
from __future__ import annotations

import json
import logging
from pathlib import Path

import pytest

import app.batch_processor as bp
from app.batch_processor import BatchOptions, BatchProcessor
from tests.powerpoint_render_stub import powerpoint_render_stubbed


class _ThreadKiller(BaseException):
    """Simulates the exotic death `except Exception` cannot catch (the 19/11
    silent worker death: no WORKER_FAILED, no done event)."""


def _opts(files, template, out, **kw):
    return BatchOptions(files=[Path(f) for f in files], template=template, output_file=out,
                        use_ollama=False, row_mode="append", **kw)


def _run(proc, **kw):
    with powerpoint_render_stubbed():
        return proc.run(**kw)


def _parse_counter(monkeypatch, fail_first_n, fail_type, target_name):
    """Wrap `app.batch_processor.parse_pptx`: fail the target file's first
    `fail_first_n` parses with `fail_type`, then delegate to the real parser.
    Returns the per-path call counts."""
    orig = bp.parse_pptx
    calls: dict[str, int] = {}

    def _parse(path, *args, **kwargs):
        key = Path(path).name
        calls[key] = calls.get(key, 0) + 1
        if key == target_name and calls[key] <= fail_first_n:
            raise fail_type(f"boom-{key}-{calls[key]}")
        return orig(path, *args, **kwargs)

    monkeypatch.setattr(bp, "parse_pptx", _parse)
    return calls


def _result_for(results, name):
    return next(r for r in results if Path(r.source_file).name == name)


def _batch_json(out):
    return json.loads((out.parent / "logs" / "batch_result.json").read_text(encoding="utf-8"))


# ================================================================== per-file isolation
def test_base_exception_in_one_file_becomes_error_and_batch_continues(sample_tree, tmp_path, monkeypatch):
    files = sample_tree["files"][:3]
    target = files[1].name
    calls = _parse_counter(monkeypatch, 99, _ThreadKiller, target)
    out = tmp_path / "o" / "r.xlsx"
    dones = []
    proc = BatchProcessor(_opts(files, sample_tree["template"], out), on_done=dones.append)
    summary = _run(proc)

    assert calls[target] == 2                       # first pass + exactly one retry pass
    assert len(proc.results) == 3
    bad = _result_for(proc.results, target)
    assert bad.status == "error" and bad.retried is True
    assert "_ThreadKiller: boom-" in bad.error
    assert summary.failed == 1
    assert [e["file"] for e in summary.errors] == [str(files[1])]
    assert summary.completed + summary.needs_review == 2   # the other files finished
    assert len(dones) == 1 and dones[0] is summary          # done delivered exactly once
    assert out.exists()


def test_base_exception_outside_loop_logs_worker_failed_and_delivers_done(sample_tree, tmp_path, caplog, monkeypatch):
    files = sample_tree["files"][:3]
    out = tmp_path / "o" / "r.xlsx"
    dones = []
    proc = BatchProcessor(_opts(files, sample_tree["template"], out), on_done=dones.append)

    def _boom():
        raise _ThreadKiller("boom-outside-loop")

    monkeypatch.setattr(proc, "run", _boom)
    with caplog.at_level(logging.DEBUG, logger="report_extractor"):
        proc._run_guarded()

    assert "WORKER_FAILED" in caplog.text
    assert len(dones) == 1
    assert dones[0].failed == 3                     # nothing accounted for -> all counted failed
    assert dones[0].errors and "WORKER_FAILED" in dones[0].errors[0]["error"]


# ================================================================== retry pass
def test_retry_success_removes_file_from_final_errors(sample_tree, tmp_path, monkeypatch):
    files = sample_tree["files"][:3]
    target = files[1].name
    out = tmp_path / "o" / "r.xlsx"
    order: list = []
    calls: dict[str, int] = {}
    orig_parse = bp.parse_pptx

    def _parse(path, *args, **kwargs):
        key = Path(path).name
        calls[key] = calls.get(key, 0) + 1
        order.append(("parse", key))
        if key == target and calls[key] == 1:
            raise RuntimeError(f"boom-{key}-1")
        return orig_parse(path, *args, **kwargs)

    monkeypatch.setattr(bp, "parse_pptx", _parse)
    proc = BatchProcessor(_opts(files, sample_tree["template"], out),
                          on_done=lambda s: order.append(("done",)))
    summary = _run(proc)

    assert calls[target] == 2                       # failed once, then retried and succeeded
    assert order[-1] == ("done",)                   # done delivered only after the retry pass
    assert order.count(("parse", target)) == 2
    assert len(proc.results) == 3                   # one entry per file, not one per attempt
    good = _result_for(proc.results, target)
    assert good.status in ("completed", "needs_review") and good.retried is True
    assert summary.failed == 0 and summary.errors == []
    assert summary.completed + summary.needs_review == 3
    res = _batch_json(out)
    assert len(res["results"]) == 3
    entry = next(r for r in res["results"] if Path(r["source_file"]).name == target)
    assert entry["status"] == good.status and entry["retried"] is True
    assert res["meta"]["summary"]["failed"] == 0 and res["meta"]["summary"]["errors"] == []
    assert "finished_at" in res["meta"]


def test_retry_still_failing_stays_in_final_errors_with_retry_error(sample_tree, tmp_path, monkeypatch):
    files = sample_tree["files"][:3]
    target = files[1].name
    _parse_counter(monkeypatch, 99, RuntimeError, target)
    out = tmp_path / "o" / "r.xlsx"
    dones = []
    proc = BatchProcessor(_opts(files, sample_tree["template"], out), on_done=dones.append)
    summary = _run(proc)

    assert len(proc.results) == 3
    bad = _result_for(proc.results, target)
    assert bad.status == "error" and bad.retried is True
    assert bad.error == "RuntimeError: boom-" + target + "-2"   # the RETRY error, not the first-pass one
    assert summary.failed == 1
    assert len(summary.errors) == 1 and summary.errors[0]["file"] == str(files[1])
    assert "boom-" + target + "-2" in summary.errors[0]["error"]
    assert len(dones) == 1
    res = _batch_json(out)
    entry = next(r for r in res["results"] if Path(r["source_file"]).name == target)
    assert entry["status"] == "error" and entry["error"] == bad.error


def test_needs_review_is_never_retried(sample_tree, tmp_path, monkeypatch):
    files = sample_tree["files"][:3]
    target = files[1].name
    calls = _parse_counter(monkeypatch, 0, RuntimeError, target)
    orig_extract = bp.extract_record

    def _force_review(report, *args, **kwargs):
        rec = orig_extract(report, *args, **kwargs)
        if Path(report.path).name == target:
            rec.review_reasons.append("forced: needs review")
        return rec

    monkeypatch.setattr(bp, "extract_record", _force_review)
    out = tmp_path / "o" / "r.xlsx"
    proc = BatchProcessor(_opts(files, sample_tree["template"], out))
    summary = _run(proc)

    assert calls[target] == 1                       # processed once: no retry for needs_review
    flagged = _result_for(proc.results, target)
    assert flagged.status == "needs_review" and flagged.retried is False
    assert summary.failed == 0 and summary.needs_review >= 1


def test_no_retry_when_stop_requested_before_retry_pass(sample_tree, tmp_path, monkeypatch):
    files = sample_tree["files"][:3]
    target = files[1].name
    calls = _parse_counter(monkeypatch, 99, RuntimeError, target)
    out = tmp_path / "o" / "r.xlsx"
    last = files[2].name
    proc = BatchProcessor(_opts(files, sample_tree["template"], out))

    def _stop_at_last(i, s, d):
        if Path(files[i]).name == last and s in ("completed", "completed_new", "needs_review"):
            proc.request_stop()

    proc.on_file = _stop_at_last
    summary = _run(proc)

    assert calls[target] == 1                       # first-pass error stands: no retry after Stop
    bad = _result_for(proc.results, target)
    assert bad.status == "error" and bad.retried is False
    assert summary.failed == 1 and summary.stopped is True
    assert len(proc.results) == 3
