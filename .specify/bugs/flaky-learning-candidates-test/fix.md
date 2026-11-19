# Fix — flaky-learning-candidates-test

**Skill:** Spec Kit bug workflow (repo `AGENTS.md` + `.specify/memory/constitution.md` +
`docs/SPECKIT_ARENA.md`). Native speckit invocation unavailable → executed from the checked-in
skill instructions in `.agents/skills/speckit-*/SKILL.md`; this fallback is disclosed here and
in the final report.

**Bug:** `tests/test_application_service.py::test_learning_candidates_counts_label_review_saved_reapply_retry_and_model_update`
failed 4/4 recent full-suite runs on Windows (and 2/2 isolated runs in this batch) with
`'Huấn luyện thất bại: assert False ... wait(timeout=5)'` — the 5 s `train_models` handshake
expired while the main thread was still rendering slide previews through real PowerPoint COM.

**Root cause (proven in `assessment.md`):** each successful render's cleanup waits up to
`_PROCESS_EXIT_TIMEOUT_S = 15.0` s for the quit PowerPoint to leave the process table
(`_wait_for_process_exit`, `app/qpn_renderer.py:632`). The wait is observation-only — Quit is
already sent — so it never makes PowerPoint exit faster; it only stalls the caller. Measured
19/11 on the ollama machine: every render's quit took the full 15.036 s (process exited
~0.6 s after the timeout), 4 consecutive renders = ~60 s of pure stall in one test, 6 renders
in the isolated run = ~90 s (test took 141.76 s vs ~35 s normal). Terminating the process is
forbidden by PROMPT-041 (tested: `test_7`, `test_16d`, `test_18`, `test_19`), so the wait can
only be bounded, not escalated.

## Changes (this fix only)

1. **`app/qpn_renderer.py`** — `_PROCESS_EXIT_TIMEOUT_S`: 15.0 → **2.0 s**. The bound caps the
   worst-case caller stall per render at ~2 s (export ~1 s + wait 2 s ≈ 3 s), so one slow render
   inside the 5 s train window can no longer eat it. 2.0 s is the value the PROMPT-035 harness
   already pins for its fake machine. No process is ever terminated; the timeout only produces
   the diagnostic `app-exit-wait-timeout` note. No test threshold touched.
2. **`tests/test_application_service.py`** — one best-effort `service.learning_state()` warm-up
   call immediately before the `train_models()` handshake, so the two `learning_state()` calls
   inside the window are cache hits instead of fresh PowerPoint renders. No assertion
   removed/weakened/added to the existing checks; a degraded warm-up simply leaves the cache
   cold (the bounded wait is what caps a slow render).
3. **`tests/test_prompt035_powerpoint_isolation.py`** — two regression tests:
   - `test_cleanup_exit_wait_is_bounded_below_the_train_window`: pins the production constant
     (`_PROCESS_EXIT_TIMEOUT_S < 5.0`; deliberately takes no `machine` fixture, which
     monkeypatches the value to 2.0).
   - `test_render_returns_promptly_when_quit_process_never_exits`: fake machine whose quit
     process never leaves the table (`quit_lingers_polls = 10**9`) → render returns after the
     bounded wait, export intact, `app-exit-wait-timeout` note logged, `App.Quit` still sent.

## Not changed

- No test skipped, no assertion removed or weakened, no timeout raised to buy green.
- No change to the no-termination contract (PROMPT-041) or to `docs/WINDOWS_ACCEPTANCE_PROMPT041.md`
  (its "15 s" is the degraded-preview TTL, not this wait).
- The Build 025 batch code (`arena/06a1c335-tnp`, commit `4f0482d`) is untouched by this fix.

## Residual risk

- The ollama machine's PowerPoint is currently degraded (every quit takes ~15.6 s, even from a
  fresh process). The fix is robust in that state (in-window calls are cache hits; any single
  render is capped at ~3 s), but the degraded state itself is a machine issue, not fixed here.
- The healthy-state PowerPoint exit distribution was not measurable while the machine is
  degraded. If a healthy machine shows spurious `app-exit-wait-timeout` notes, the 2.0 s bound
  can be revisited (the note is diagnostic-only; renders still succeed).

## Verification

- Linux targeted tests (development evidence only): see `test.md`.
- Windows full suite in the build venv (the gate that failed) + isolated test run: see `test.md`.
- Full build + deploy + PR: per the task brief, after the gate is green.
