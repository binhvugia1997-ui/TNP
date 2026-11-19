# Bug Assessment: flaky learning-candidates test blocks the Build 025 gate

- **Slug**: flaky-learning-candidates-test
- **Created**: 2026-11-19
- **Source**: pasted text (task brief from the parent agent) + measured build log `build2.log` run 3 (19/11 13:13, `.venv-build`, full suite)
- **Verdict**: valid
- **Severity**: high

## Report (verbatim or summarized)

`tests/test_application_service.py::test_learning_candidates_counts_label_review_saved_reapply_retry_and_model_update`
fails 4/4 recent full-suite runs on Windows (1× system Python 3.14 + 3×
build venv `.venv-build`) with three near-different proximate causes
(COM `0x80010108`, `powerpoint-presentation-open`,
`powerpoint-user-session-running` — a self-created PowerPoint not exiting in
time during cleanup). It passes in isolation (with the Build 025 fix present)
and passed in the Build 024 full run. The test does not call any code changed
by Build 025. The test calls `learning_state()`, which renders slide previews
through real PowerPoint COM; the `train_models` 5s timeout races that render.

## Symptom

In a full-suite run the test fails at its final assertion with
`'Huấn luyện thất bại: assert False ... wait(timeout=5)'` instead of
`"Test model update completed."` — the background training thread's 5s
`release.wait()` expires because the main thread is still stuck inside a
`learning_state()` PowerPoint render. In isolation the same test passes.

## Reproduction

1. On ollama, in the build venv (`.venv-build`, Python 3.12.16, pythonnet
   present), run the full suite:
   `.\.venv-build\Scripts\python.exe -m pytest tests -q -p no:cacheprovider`
   (the same command the build gate runs).
2. Observe: `1 failed, 1424 passed, 2 skipped` — the single failure is this
   test (reproduced 3/3 build runs 19/11; run 3 in `build2.log`, 966.11s).
3. Run the test alone: it passes (baseline; isolated re-run in progress at
   assessment time — see Open Questions).

Reproduced by this batch (isolated, build venv, 19/11 ~13:52):
`1 failed in 141.76s` (normal: ~35 s) — same assertion
(`test_application_service.py:726`, `'Huấn luyện thất bại: assert False ...
wait(timeout=5)'`), with **six** `app-exit-wait-timeout` events in one
test (pids 9772, 11872, 12924, 16340, 15964, 20612), each followed by a
`powerpoint-presentation-open` failure — 6 renders × 15 s = ~90 s of pure
stall. A fresh-process timing probe (`measure_render_timing.py`, scratch,
not committed) reproduced the same stall on its very first render
(`app-exit-wait-timeout pid=17700`), and its second render failed with
`presentation-open` — so the degradation is machine state (PowerPoint's
automation server), not state held by the pytest process. No PowerPoint
was running before that probe run.

## Suspected Code Paths

- `tests/test_application_service.py:627` — the test; `slow_train_models`
  (`release.wait(timeout=5)`, line ~705) and the two `learning_state()`
  calls on the main thread inside the 5s window (`train_models()` return,
  line ~711, and the `training` assert, line ~712).
- `app/application_service.py:1119` `learning_state()` — calls
  `_image_candidate_dto()` → `_slide_preview_for()` (line ~1686), which
  calls `SlideRenderer.render(..., purpose="learning_preview")` — a real
  PowerPoint COM render. Degraded/failed previews are deliberately NOT
  cached, so every `learning_state()` re-attempts the render.
- `app/qpn_renderer.py:695` `_render_with_powerpoint_locked()` — the render;
  cleanup (lines ~790–815) closes the presentation, `Quit()`s the owned
  instance, then `_wait_for_process_exit()` waits up to
  `_PROCESS_EXIT_TIMEOUT_S = 15.0` (line 413) and on timeout only logs
  `app-exit-wait-timeout` and returns — **it never terminates the process**.
- `app/qpn_renderer.py:632` `_wait_for_process_exit()` — PID-poll wait,
  docstring: "Observation only: on timeout it reports and returns False —
  it never terminates anything."
- `app/qpn_renderer.py:666` ownership proof — any PowerPoint process in the
  session refuses the next render (`powerpoint-user-session-running`).

## Root Cause Hypothesis

**Demonstrated (high confidence for the mechanism):** the 5s train window is
only the tripwire. The stall is a render whose cleanup burns the whole
15s quit-wait and then leaves its own PowerPoint behind; the next render
then fails or is refused, and the window expires.

Measured from build run 3 (`build2.log`, 19/11 13:13, all timestamps from the
test's captured stderr — the log holds only the failing test's output, not
the whole suite):

- `13:13:35,528` — snapshot: no PowerPoint running (`raw=[]`).
- `13:13:36,046` — render creates its own PowerPoint, pid 19236.
- `13:13:51,082` — `SLIDE_RENDER_CLEANUP ... app-exit-wait-timeout
  detail=pid=19236 still running; not terminated` — **15.036 s after
  creation, i.e. the full `_PROCESS_EXIT_TIMEOUT_S = 15.0` expired**.
  The single render call therefore took ~15.6 s (13:13:35.5 → 13:13:51.08),
  **3× the 5 s train window**.
- `13:13:51,083` — the next render fails at `presentation-open`
  (`com_error 0x80020009`, `powerpoint-presentation-open`).
- `13:13:51,163`–`13:13:51,400` — four further attempts refused
  `powerpoint-user-session-running` (pid 19236 still alive).
- Final assertion (`test_application_service.py:726`):
  `assert 'Huấn luyện thất bại: assert False ... wait(timeout=5)' ==
  'Test model update completed.'`

The stall is not a one-off: the same test produced **four** consecutive
renders, each burning the full 15 s quit-wait (timeouts at 13:13:51,
13:14:06, 13:14:22, 13:14:37; pids 19236, 9344, 12920, 7780 — all
different, so each previous process did exit, ~0.6 s *after* its timeout:
pid 19236 was gone by the next render's snapshot at 13:13:51,728).
**4/4 renders × 15 s = ~60 s of pure stall inside one test.** PowerPoint
takes ~15.6 s to exit after `Quit()` in this degraded state, and the wait
— observation-only, it never causes the exit — blocks the calling thread
for all of it.

The test's own assertions are correct; nothing in the test's contract is
violated by the product. The same lingering-process mechanism also hurts
production batches: every file render spawns PowerPoint, and a lingering
instance makes the next file's render refuse (the 19/11 company-machine
incident had the same shape — a PowerPoint already running blocks the
render). The ~20 `Windows fatal exception: code 0x80010108`
(RPC_E_DISCONNECTED) lines in the same log are a separate COM-instability
symptom of the same churn; they do not by themselves fail this test.

Refinement (19/11, during assessment): the stall is specifically the
**observation-only** quit-wait, not the refusals. Refusals are fast
(~80 ms apart in the log above), so a degraded preview never consumes the
window — only a *successful* render whose PowerPoint is slow to exit does,
because `_wait_for_process_exit()` blocks the calling thread for the full
15 s. A second, isolated run of the same test (19/11 ~13:50, build venv)
also failed — in 142.88 s instead of the usual ~35 s — with
`powerpoint-presentation-open` in the learning preview; the machine had no
lingering PowerPoint before that run, so the stall was created inside the
test itself. Full failure output of that run is pending (captured to
`flaky-isolated.log` on ollama).

## Contract constraint discovered during assessment

Terminating the owned PowerPoint is **not** an available fix. PROMPT-041
deliberately forbids it, and the ban is tested:

- `tests/test_prompt041_isolated_powerpoint.py::test_7` — "cleanup failure
  is reported and never escalates to the user or to a kill"; the harness
  arms `os.kill` / `subprocess` to fail, so any escalation attempt explodes.
- `test_16d` — the session helpers are read-only queries; asserts no
  `OpenProcess` / `TerminateProcess` / `os.kill` / `.terminate(` in
  `qpn_renderer.py`.
- `test_18` — "no process termination or registry writing anywhere in the
  application"; scans all of `app/*.py` for `taskkill`,
  `terminateprocess`, `psutil`, `pkill`, `killall`, `stop-process`, etc.
- `test_19` — exactly one `.Quit()` and one `.Close()` in the renderer.
- `tests/test_prompt035_powerpoint_isolation.py` harness — `os.kill`,
  `subprocess.run`, `subprocess.Popen` are armed to raise
  `AssertionError("process termination attempted")`.

The ban exists to protect the user's unsaved PowerPoint work: a kill
aimed at the wrong instance would destroy it. Lifting it would be an
explicit contract change requiring Gia's approval — out of scope for this
batch. The wait itself is observation-only (Quit is already sent; waiting
does not make PowerPoint exit faster), so the fix must bound the *wait*,
not escalate to a kill.

## Proposed Remediation

**Preferred (two parts, both inside the existing contract):**

1. **Product — bound the cleanup exit-wait** (`app/qpn_renderer.py`):
   lower `_PROCESS_EXIT_TIMEOUT_S` from 15.0 s to 2.0 s — the value the
   PROMPT-035 harness already pins for its tests, i.e. the suite's own
   model of this bound. Justification by measurement: the wait is
   observation-only (Quit is already sent; waiting never makes PowerPoint
   exit faster), and the 15 s bound is 3× the 5 s train window, so one
   stalled render alone kills the test. The wait's diagnostic role is kept
   (a timeout still logs `app-exit-wait-timeout`); only the worst-case
   caller stall shrinks, from 15 s to 2 s. No threshold in any test is
   touched, and no process is ever terminated. (The exact normal-exit
   distribution will be measured before the fix lands — see Open Questions.)

2. **Test — warm the preview cache before training**
   (`tests/test_application_service.py`): call `service.learning_state()`
   once immediately before `train_models()` so the two `learning_state()`
   calls inside the 5 s window are cache hits (milliseconds) instead of
   fresh renders. No assertion is removed or weakened and no timeout is
   changed; the added call only removes the test's inherent race (a 5 s
   thread handshake sharing the main thread with COM renders). The test's
   purpose is the train flow (counts / labels / retry / model-update
   message), not render timing.

Together: (1) removes the product stall that also hurts production
batches; (2) makes the train window deterministic even on a cache miss
(a single render then costs export ~0.5–1 s + wait ≤ 2 s, not 15 s).

**Alternatives** (rejected):
- *Terminate the owned PowerPoint after the wait* — forbidden by the
  PROMPT-041 no-kill contract (see above); would need an explicit
  contract change approved by Gia.
- *Raise the 5 s train timeout*: changes the test contract without fixing
  the cause. The measurement shows the stall is 15 s, so a "justified"
  timeout would have to exceed 15 s — absurd for a train handshake.
- *conftest sweeper killing leftover headless PowerPoint between tests*:
  follow-up only, and it would violate the same no-kill rule in test
  infrastructure. It addresses cross-test leakage from crashed runs, not
  our own render stalling its caller.

**Files likely to change**:
- `app/qpn_renderer.py` (handle-based wait + terminate escalation)
- `tests/test_prompt041_isolated_powerpoint.py` or a new test file
  (regression tests for the cleanup escalation)
- `.specify/bugs/flaky-learning-candidates-test/{assessment,fix,test}.md`

**Tests to add or update**:
- Regression test proving an owned PowerPoint that ignores `Quit()` is
  terminated and the next render is not refused (Windows COM path; a
  portable unit test for the wait/terminate helper on a sacrificial
  subprocess where the platform allows).
- Acceptance: full Windows suite green in `.venv-build` (the exact gate
  that failed, 1424/1/2 → 0 failures) + the isolated test run.

## Risks & Considerations

- Terminating the wrong process on PID reuse — mitigated by handle
  targeting; must be documented in the fix.
- The escalation only fires after the existing 15 s wait, so the success
  path is unchanged; other render callers (batch, QPN) get the same
  determinism benefit.
- The `0x80010108` disconnects may persist as log noise; they are not the
  cause of this test's failure.
- No test assertion is removed, no test is skipped, no timeout is raised
  to force green (constitution VIII).

## Open Questions

- Full-suite reproduction for this batch: the gate run itself (after the
  fix) is the verification; the pre-fix full-suite failures are already
  documented (4/4 runs, run 3 in `build2.log`).
- Healthy-state PowerPoint exit distribution: unmeasured — the machine is
  currently degraded (every render's quit takes ~15.6 s), so no healthy
  sample exists right now. The 2.0 s bound is justified by the harness's
  own pin and by the stall being 7.5× the bound; if a healthy machine
  shows spurious timeout notes, the bound can be revisited (the note is
  diagnostic-only; renders still succeed).
- `gh auth login` was completed by Gia on ollama 19/11 (logged in as
  binhvugia1997-ui, HTTPS) — the earlier push blocker is cleared; push +
  PR proceed after a green build, no further auth step needed.
