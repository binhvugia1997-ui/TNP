# Bug Fix: PowerPoint session isolation — cleanup operates only on proven app-owned resources (PROMPT-035)

- **Slug**: prompt035-powerpoint-session-isolation
- **Fixed**: 2026-10-13
- **Assessment**: ./assessment.md
- **Status**: applied (automated/fake-COM proven; real Windows PowerPoint acceptance pending)
- **Spec Kit execution mode**: Executed from checked-in Spec Kit skill instructions; native Arena skill invocation was not available.

## Summary

`render_with_powerpoint` no longer assumes that `DispatchEx` returned an app-owned instance. It proves ownership (no PowerPoint process before, exactly one new process after, pristine instance), opens a presentation only into that pristine instance, closes only the presentation it opened, and quits only an Application that is proven app-created **and** still exclusive at cleanup time. If ownership cannot be proven it fails closed (`PowerPointStageError("ownership")`) and the existing renderer chain falls back to LibreOffice/built-in.

## Ownership model

| | Before (`502968c`) | After |
|---|---|---|
| Application | "owned" because `DispatchEx` was called (`created_app = True` unconditionally) | owned only if: process table readable **and** empty before; `GetActiveObject` finds nothing; exactly one new `POWERPNT.EXE` PID after; instance has 0 presentations and 0 windows |
| User Application | could be returned by `DispatchEx`, opened into, and `Quit()` | never attached/opened/closed/quit: refused before activation, or released untouched if it appears during activation |
| Presentation | object from `Open()` inside possibly-shared instance | object from our `Open()` on a pristine instance (`Count == 0` re-checked immediately before Open) |
| Quit gate | none | only if `owned` **and** remaining presentation count is exactly what we expect (0, or 1 if our own Close failed); unreadable count ⇒ no Quit |
| Concurrency | none | module lock serialises the COM lifecycle |
| Process termination | none (but unsafe Quit) | none; only a read-only Toolhelp snapshot + bounded wait for the app-owned PID to exit; timeout is logged, never escalated |

## Cleanup behaviour

- **Success**: close our presentation → (exclusivity re-check) → Quit our instance → release refs → bounded wait for our PID to exit → `CoUninitialize`.
- **Error** (`presentation-open`, `slide-export`, cleanup failures): same ownership-gated `finally`; user-owned objects are never reached. Quit failure is logged, never escalated to a kill.
- **Cancellation**: `CancellationRequested` raised in the export loop now propagates unchanged (it was previously wrapped into `PowerPointStageError("slide-export")`); cleanup closes/quits only owned resources. Pre-backend cancel never touches COM.
- **Timeout** (COM call raising `TimeoutError`/RPC timeout): treated as `slide-export`; owned-only cleanup, no process kill.
- **Refusal / shared / ambiguous**: no Open/Close/Quit; `CoUninitialize` stays balanced; renderer chain continues; failure stage is `ownership` (not `availability`) so a degraded preview is retried rather than pinned.

## Changes

| File | Change | Notes |
|------|--------|-------|
| `app/qpn_renderer.py` | modified | Added `_POWERPOINT_LOCK`, `_powerpoint_process_ids` (read-only Toolhelp snapshot, `None` = unknown), `_running_powerpoint_registered`, `_is_pristine`, `_presentation_count`, `_ownership_refusal`, `_wait_for_process_exit`; `render_with_powerpoint` split into a lock wrapper + `_render_with_powerpoint_locked` with ownership-gated cleanup; `CancellationRequested` re-raised unchanged. Public signature unchanged. Imports `threading`, `time`, `FrozenSet`. |
| `tests/test_prompt035_powerpoint_isolation.py` | added | 24 tests: fake single-instance COM machine recording Close/Quit with owner and order; TESTS 1–8, race, ROT-only visibility, unreadable process table, cancellation/timeout, cleanup-failure-no-kill, concurrency lock, renderer-chain degradation, static no-kill contracts. |
| `tests/test_prompt027r_renderer.py` | modified (harness only) | `_fake_com` now models a machine with no user PowerPoint (process table, `GetActiveObject` "nothing running", `Presentations.Count`, `Windows.Count`). No assertion changed, removed or weakened. |
| `.specify/bugs/prompt035-powerpoint-session-isolation/{assessment,fix,test}.md` | added | workflow evidence |

## Diff Highlights

```python
before = _powerpoint_process_ids()
if before is None:  raise _ownership_refusal("process table unreadable; ownership cannot be established", pptx)
if before:          raise _ownership_refusal("a PowerPoint session is already running; not attaching to it", pptx)
if _running_powerpoint_registered(win32com.client): raise _ownership_refusal(...)
app = win32com.client.DispatchEx(POWERPOINT_PROGID)           # may return a USER's instance: unproven
created = after - before
if created is None or len(created) != 1 or not _is_pristine(app): raise _ownership_refusal(...)  # owned stays False
...
finally:
    ...close only our pres...
    if owned and remaining == (0 if pres_closed else 1): app.Quit()   # exclusive + proven
```

## Tests Added or Updated

- `test_1_*`, `test_1b_*` — existing user session is never Quit; fail closed at stage `ownership`, no activation.
- `test_2_*` — user's open (unsaved) presentation is never closed / opened into.
- `test_3_*` — app-opened presentation is closed.
- `test_4_*`, `test_4b_*` — app-owned Application is Quit, after Close, with bounded wait; reusable afterwards.
- `test_5_*`, `test_5b_*`, `test_5c_*` — exception paths close only owned resources; Quit failure never escalates to kill (`os.kill`/`subprocess` armed to fail).
- `test_6_*`, `test_6b_*`, `test_6c_*` — cancellation (unchanged exception), timeout, pre-start cancel with a user session.
- `test_7_*`, `test_7b_*` — user work arriving in our instance mid-render ⇒ our presentation closed, **no Quit**; user + app presentations coexist ⇒ only ours closed.
- `test_race_*`, `test_running_instance_invisible_*`, `test_unreadable_process_table_*`, `test_ownership_refusal_degrades_*` — fail-closed ownership proofs and graceful fallback.
- `test_concurrent_jobs_*`, `test_real_process_table_reader_*` — lock and unknown-off-Windows.
- `test_8_*`, `test_8b_*`, `test_8c_*` — static contract: no `taskkill`/`TerminateProcess`/`psutil`/`os.kill`/`.kill(`/`.terminate(`/`OpenProcess` in `app/`; no enumerate-all-presentations; no attaching `Dispatch`/`GetObject`.

Pre-fix result of the new suite against unmodified `502968c` code: 12 failed / 10 passed (see assessment). Mutation checks: removing the lock fails the concurrency test; removing the Quit exclusivity guard fails TEST 7.

## Local Verification

- `pytest tests/test_prompt035_powerpoint_isolation.py tests/test_prompt027r_renderer.py` → 49 passed.
- Full suite: 35 failed / 1206 passed / 18 skipped after, vs 35 failed / 1182 passed / 18 skipped at base `502968c`; failure sets identical (pre-existing, see test.md).
- Manual checks: none (Linux sandbox; no PowerPoint).

## Deviations from Assessment

- Added `_presentation_count`-based expected-count logic so a failed `pres.Close()` of our own presentation does not leak a hidden app (expected remaining = 1 in that case). Minor refinement within the planned boundary.
- Added a pristine re-check immediately before `Open` (assessment step 4 described as `Count == 0`; implemented as `_is_pristine`, which also checks windows).

## Follow-ups

- **Windows acceptance (pending)** — see test.md.
- Product decision (not made here): while the user's PowerPoint is running, extraction uses the lower-fidelity LibreOffice/built-in renderer. A future, Windows-validated design could attach read-only to a user instance and close only its own presentation, but that is intentionally out of this fix.
- The pre-existing stale version assertions (1.3.5 vs 1.3.6) are unrelated and untouched.
