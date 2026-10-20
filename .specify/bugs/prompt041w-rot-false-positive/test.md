# Bug Verification: PowerPoint false positive from a stale Running Object Table entry (PROMPT-041W)

- **Slug**: prompt041w-rot-false-positive
- **Tested**: 2026-10-17
- **Assessment**: ./assessment.md
- **Fix**: ./fix.md
- **Result**: **automated verification passes; real Windows PowerPoint acceptance NOT performed and NOT claimed.** The agent runs on Linux with no PowerPoint, no pywin32 and no Toolhelp API, so everything Windows-only (the real process table, `ProcessIdToSessionId`, the real Running Object Table) is exercised through the PROMPT-035 fake COM machine. The Windows checklist is `docs/WINDOWS_ACCEPTANCE_PROMPT041W.md`, and its **first** case is "no PowerPoint running".
- **Spec Kit execution mode**: Executed from the checked-in Spec Kit skill instructions; native Arena skill invocation was not available. `.specify/extensions.yml` has `hooks: {}`.
- **Environment**: Linux sandbox, Python 3.11.2, pytest 9.1.1 in a throw-away `.venv` (gitignored, not committed), Node v22.22.3 / npm 10.9.8.

## Summary

The false positive **reproduces deterministically against the unmodified base** (`08b3de1`) and the target contract
holds after the fix. 33 new Python regressions plus 3 frontend regressions; **no existing test was edited or
weakened** — including the whole PROMPT-035 ownership suite, whose file is asserted unchanged by regression 12.
Six mutations were run to prove the new tests actually bind to the fix rather than to the fakes.

## Checks performed

| # | Check | Command / action | Result |
|---|---|---|---|
| 1 | **Reproduction (pre-fix)** | throw-away harness `tests/test_zz_repro041w.py` on the unmodified `08b3de1` tree, then deleted | **reproduced**: `process_ids()=frozenset()`, `stage=ownership`, `code=powerpoint-registered-instance`, `dispatches=0`, `last_backend=builtin`, `preview_reason=powerpoint-registered-instance`, `degraded=True` |
| 2 | **Reproduction of the second defect (pre-fix)** | same harness | `_preview_expired(age=3 s, TTL=15.0) = False` → a degraded preview *and its reason* are served verbatim inside the window |
| 3 | **Message attribution** | `grep -rn "Một phiên PowerPoint đang chạy" frontend/src app docs .specify` | **exactly one hit**, `LearningTab.tsx` → the slug is `powerpoint-registered-instance`, not `powerpoint-user-session-running` |
| 4 | New regressions (post-fix) | `.venv/bin/python -m pytest -q tests/test_prompt041w_rot_false_positive.py` | **33 passed** |
| 5 | Frontend regressions (post-fix) | `cd frontend && node --test tests/prompt041w-preview-reason.test.mjs` | **3 passed** |
| 6 | Full Python suite | `.venv/bin/python -m pytest -q tests` | **1389 passed, 13 skipped, 1 failed** |
| 7 | The one failure | `tests/test_prompt028_portable.py::test_live_import_probe_reports_every_module_in_this_venv` | **pre-existing and Linux-only** — needs the win32-only `pythonnet` / `clr_loader`. Unchanged from the PROMPT-041 baseline; not caused by this fix |
| 8 | The 13 skips | `pytest -q tests -rs` | all environmental: `pythonnet` / `clr_loader` / `git history unavailable` / `frontend/dist not built here` (that skip disappears once `npm run build` has run, which is why the skip count is 13 rather than 14) / one Windows-only native-load test |
| 9 | Frontend suite | `cd frontend && npm run test:frontend` | **89 passed** (86 before + 3 new) |
| 10 | Frontend typecheck | `cd frontend && npm run typecheck` (`tsc --noEmit`) | clean |
| 11 | Frontend build | `cd frontend && npm run build` | `✓ 1588 modules transformed`, `✓ built in 9.75s` (`dist/` is gitignored and was not staged) |
| 12 | Lint | `.venv/bin/python -m pyflakes app/qpn_renderer.py app/application_service.py tests/test_prompt041w_rot_false_positive.py` | only the **3 pre-existing** `imported but unused` notes in `qpn_renderer.py` (`win32com.client`, `fitz`, `PIL.ImageChops` — deliberate side-effect imports); nothing new |

## The 12 mandatory regressions

| Required | Test | Asserts |
|---|---|---|
| 1 — raw snapshot empty → no user-session-running refusal | `test_01`, `test_03`, `test_03b` | render succeeds, `dispatches == 1`, capability `isolated=True reason=ok`, no `App.Quit` on a user instance |
| 2 — same-session PID → correct refusal | `test_04` | `stage=ownership code=powerpoint-user-session-running`, `dispatches == 0`, `user.quit_calls == 0`, user presentation not closed |
| 3 — different-session PID only → no same-session refusal | `test_05`, `test_05b` | session filter drops it; capability `ok`; render succeeds and the out-of-session PowerPoint is never touched |
| 4 — `ProcessIdToSessionId` failure → distinct code | `test_06`, `test_06b`, `test_06c` | `powerpoint-session-unproven`, **never** `powerpoint-user-session-running`; `test_06c` drives the *real* helper with `sys.platform="win32"` |
| 5 — snapshot API failure → distinct code | `test_07`, `test_07b` | `powerpoint-process-table-unreadable`, fail closed, `dispatches == 0` |
| 6 — empty set not confused with `None`/error | `test_08`, `test_08b`, `test_08c`, `test_08d` | `frozenset() is not None` and `!= None`; different reasons; the self-vouch; three ROT states |
| 7 — cached prior refusal → retry clears it | `test_09`, `test_09b`, `test_09c`, `test_09d` | memory cache discarded and re-rendered the moment PowerPoint disappears; non-PowerPoint fallbacks not hammered; the on-disk copy cannot resurrect the reason |
| 8 — successful PowerPoint render clears `slidePreviewReason` | `test_10` | `backend=powerpoint`, `faithful=True`, `reason == ""`, real base64 bitmap |
| 9 — candidate switching cannot leak a reason | `test_11` | blocked candidate and a *different* report alternate without cross-talk in either direction |
| 10 — PROMPT-035 ownership tests unchanged/green | `test_12` + run + `git status` | the file is read back: `assert user.quit_calls == 0`, `assert machine.dispatches == 0`, `assert machine.touched_user_app is False`, `assert info.value.stage == "ownership"`, `assert machine.user_presentations_closed == []`, ≥23 `def test_`, and the literal string `PROMPT-041W` must be **absent**. Suite: **24 passed**; `git status --short` on the file is empty |
| 11 — PROMPT-041 suite green | run + `git status` | `tests/test_prompt041_isolated_powerpoint.py` → **35 passed**, file **unmodified** (`git status --short` empty). Its `test_20` asserts its own explicit 8-slug set is stable and mapped in the UI, so the *added* slug does not disturb it; the 9-slug set including `powerpoint-session-unproven` is pinned by `test_13` here |
| 12 — PROMPT-030B / 038 / 039 / 040 regressions green | combined run | **254 passed** across `prompt035`, `prompt041`, `prompt041w`, `prompt027r_renderer`, `prompt027_rendering`, `prompt030b_integration`, `prompt038_desktop_freshness`, `prompt040_excel_persistence`, `prompt020_report_scoping`, `prompt025_multi_item`, `improvement_visual_regions` |

Plus the diagnostics surface: `test_18b` drives the real `run_diagnostics()` and asserts that
`đóng PowerPoint của bạn` appears **only** for the proven `powerpoint-user-session-running` case and never for
`powerpoint-session-unproven`, `powerpoint-process-table-unreadable` or `powerpoint-unavailable`, while the
PROMPT-041 `test_24` separation of "no PowerPoint here" from "your PowerPoint is open" still holds.

Plus the observability and safety contracts: `test_14` / `test_14b` / `test_14c` assert every required log field
(`pid`, `session`, `status`, `raw`, `same_session`, `vouched`, `session_provable`, `rot`, `decision`, `reason`,
`cached_backend`, `cached_reason`, `action`); `test_15` / `test_16` / `test_17` scan the whole `app/` package for
process termination, registry writes, `Dispatch(`, `GetObject(`, `ActivePresentation`, `OpenProcess`,
`TerminateProcess`, `os.kill`, `.terminate(`, and assert `.Quit()`/`.Close()` each appear exactly once;
`test_18` pins the capability memo; `test_19` pins that the reported sentence is gone while the legacy slug keeps
wording for already-persisted previews.

## Mutation testing — do the new tests bind to the fix, or only to the fakes?

Each mutation was applied, the suite run, and the file restored byte-identical from a `/tmp` copy
(`git diff --numstat` re-checked after every restore).

| Mutation | Result | Killed by |
|---|---|---|
| **M1** — ROT gate reverted to the pre-fix unconditional refusal (`if rot == ROT_REGISTERED:`) | **2 failed, 29 passed** | `test_01`, `test_02` — this *is* the reported bug |
| **M2** — `_degraded_reason_is_stale()` hard-wired to `False` (contract E off) | **4 failed, 27 passed** | `test_09`, `test_09c`, `test_09d`, `test_10` |
| **M3** — `_session_filter_is_provable()` hard-wired `True` (contract D collapses into user-session-running) | first run **survived** (my tests stubbed the helper) → added `test_06c`, which drives the *real* helper with `sys.platform="win32"` → **1 failed, 31 passed** | `test_06c` |
| **M4** — `_process_table_vouched()` hard-wired `True` off Windows (the PROMPT-035 blind-spot case stops failing closed) | caught immediately by `tests/test_prompt035_powerpoint_isolation.py` (**1 failed, 54 passed**); first run of my own file survived → added an off-Windows assertion to `test_08b` → **1 failed, 31 passed** | PROMPT-035's own suite, and `test_08b` |
| **M6** — the diagnostics row collapses every PowerPoint-present reason back into "đóng PowerPoint của bạn" | **1 failed, 67 passed** | `test_18b` — the diagnostics panel had the same over-claim as the preview notice and is fixed too |
| **M5** — the memory-path `pop` replaced with `pass` | **survived — behaviourally equivalent**, not an escape | n/a: the branch does not `return`, so the entry falls through and is re-rendered either way, and `_remember_preview` overwrites the same key. Observable behaviour is identical. |

M3 and M4 surviving the *first* run is the useful result: it showed two assertions were testing my stubs rather than
the code, and both were rewritten to drive the real functions.

## What is NOT verified

- **Real Windows behaviour.** No Toolhelp32 snapshot, no `ProcessIdToSessionId`, no real Running Object Table and no
  real `POWERPNT.EXE` were exercised. The vouch, the session filter and the three ROT states are all faked here.
- **Whether the tester's machine has a stale ROT entry.** That is the proven mechanism, but confirming it *on that
  machine* is a Windows step — the acceptance doc includes the exact command that reads it.
- **Real PowerPoint fidelity.** Unchanged by this fix, and still a Windows acceptance step.
- **No claim of Windows acceptance is made anywhere in this change.**
