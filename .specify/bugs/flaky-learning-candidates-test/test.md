# Test — flaky-learning-candidates-test

**Skill:** Spec Kit bug workflow (fallback: checked-in skill instructions; native invocation
unavailable). Linux results are development evidence only; Windows results are the gate.

## Fix under test

- `app/qpn_renderer.py`: `_PROCESS_EXIT_TIMEOUT_S` 15.0 → 2.0 s (bounded, observation-only
  cleanup wait; no process ever terminated).
- `tests/test_application_service.py`: best-effort `learning_state()` warm-up before the
  `train_models()` 5 s handshake.
- `tests/test_prompt035_powerpoint_isolation.py`: 2 new regression tests.

## Linux (development evidence only)

- `tests/test_prompt035_powerpoint_isolation.py` + `tests/test_prompt041_isolated_powerpoint.py`:
  **61 passed** in 7.86 s (includes the 2 new regression tests).
- Targeted: 2 new regression tests + the flaky test in isolation: **3 passed** in 4.17 s.

## Windows — isolated flaky test (build venv `.venv-build`, ollama)

- **Before fix (19/11):** FAILED 2/2 isolated runs — `1 failed in 141.76 s` (and 142.88 s),
  same assertion (`test_application_service.py:726`), 6× `app-exit-wait-timeout`.
- **After fix (19/11):** **1 passed in 27.96 s, EXIT=0** (`flaky-fixed-isolated.log`).
  0× `app-exit-wait-timeout`, 0× `powerpoint-user-session-running` /
  `powerpoint-presentation-open`. The 11× `Windows fatal exception: code 0x80010108` lines
  are the usual COM RPC-disconnect noise during renders; the test passed through them.

## Windows — full suite (the gate that failed)

- Run as part of the full build (`tools\build_portable.py --python <uv cpython-3.12.16>
  --no-publish --fresh-venv`, log `build2.log` on ollama, 19/11 ~14:11):
  **1427 passed, 2 skipped, 1 warning in 327.48s (0:05:27), 0 failed — GATE GREEN.**
- Before the fix the same gate was **1424 passed / 1 failed / 2 skipped in 966.11s**
  (the flaky test failing). Now: 1427 = 1424 + the flaky test fixed + 2 new regression
  tests. Suite time dropped 966 s → 327 s because the 2 s bound also removed the
  15 s stall from every other stalled render in the suite.
- The `0x80010108` COM RPC-disconnect tracebacks in the log are the usual noise during
  renders; the suite passed through them (0 failed).

## Result

- `test.md` status: **verified** — full-suite gate green on Windows + isolated test green.
