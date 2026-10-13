# Bug Assessment: Report Extractor can Quit()/close the user's already-open PowerPoint (PROMPT-035)

- **Slug**: prompt035-powerpoint-session-isolation
- **Created**: 2026-10-13
- **Source**: pasted text (PROMPT-035 report); public documentation of PowerPoint's single-instance automation behaviour (see Evidence)
- **Verdict**: valid — proven by automated fake-COM reproduction of the lifecycle; real-PowerPoint confirmation is Windows acceptance (pending)
- **Severity**: critical — data-loss risk (a user's unrelated, possibly unsaved, presentation is closed by a background batch)
- **Spec Kit execution mode**: Executed from checked-in Spec Kit skill instructions; native Arena skill invocation was not available.
- **Base**: `502968c` (v1.3.6 — Build 019). Session branch is `arena/0997da21-tnp` (fixed by Arena, branched from the requested `bug/prompt035-powerpoint-session-isolation` @ `502968c`).

## Report (summarized)

While Report Extractor processes PPTX files, presentations the user already has open in PowerPoint appear to be closed automatically. Safety contract (A–I): an existing PowerPoint instance and the user's presentations must stay open; unsaved work must never be closed/saved/modified; the app may Close only presentations it opened and Quit only an Application it created and exclusively owns; attach/reuse must never Quit; error/timeout/cancel paths obey the same rules; no broad `POWERPNT.EXE` termination; if ownership cannot be established, fail closed.

## Symptom

Starting a PPTX job while the user's PowerPoint is open closes the user's presentation(s). Expected: nothing in the user's session is touched.

## Reproduction

Automated (Linux, fake COM with documented single-instance semantics): `tests/test_prompt035_powerpoint_isolation.py`, run against the **unmodified** `app/qpn_renderer.py` at `502968c`:

```
12 failed, 10 passed
Application.Quit() was sent to the user's PowerPoint                (test_1)
DID NOT RAISE PowerPointStageError  (no fail-closed on a running session)  (test_1b)
the user's open presentation was closed                              (test_2)
assert (['user_unsaved.pptx'] == [])  user work lost                 (test_7b)
a PowerPoint that now holds user work is not exclusively ours        (test_7: Quit sent to shared process)
app.qpn_renderer.PowerPointStageError: slide-export: cancel          (test_6: CancellationRequested misclassified as failure)
DID NOT RAISE PowerPointStageError  (race: user launches PPT during DispatchEx)
DID NOT RAISE PowerPointStageError  (ROT-visible instance not noticed; unreadable process table not fail-closed)
```

Order of the failing sequence in the fake (identical to the real code path): `DispatchEx` → *returns the user's instance* → `Presentations.Open(job.pptx)` inside the user's session → `pres.Close()` → **`app.Quit()` on the user's instance** → every user presentation in that process is closed.

PROVEN BY AUTOMATED REPRODUCTION: the code unconditionally treats whatever `DispatchEx` returns as app-owned and quits it.
REQUIRES WINDOWS/REAL POWERPOINT ACCEPTANCE: that the real PowerPoint build actually hands `DispatchEx` the running instance on the user's machine, and the end-to-end "user's unsaved edit survives" check. Not performed here.

## Evidence for the single-instance premise (external, not run here)

- "In PowerPoint 2000, there can be one and only one instance of a PowerPoint Application object; `CreateObject()` references an already running instance of PowerPoint" — https://flylib.com/books/en/1.19.1.86/1/
- Independent bug report on the same hazard (`Quit()` on an attached PowerPoint closes the user's document): https://github.com/agentscope-ai/QwenPaw/issues/8002
These support, but do not replace, Windows acceptance.

## Lifecycle audit (full repository; evidence per question)

Searched `app/ tools/ tests/ run.py *.bat *.spec requirements*` (excluding the immutable `backup/`) for: `PowerPoint.Application`, `win32com`, `pythoncom`, `Dispatch`, `DispatchEx`, `GetActiveObject`, `CreateObject`, `EnsureDispatch`, `Presentations`, `Open`, `Close`, `Quit`, `POWERPNT`, `taskkill`, `TerminateProcess`, `os.kill`, `.terminate(`, `.kill(`, `psutil`, `subprocess`, `atexit`, `signal`.

The **only** COM lifecycle in the product is `app/qpn_renderer.py::render_with_powerpoint` (called only from `SlideRenderer.render`, which is used by `batch_processor.py`, `image_extractor.py`, `application_service.py` learning preview). `pptx_parser.py` is python-pptx (no COM). `scanner.py` etc. only mention PowerPoint in text. `powerpoint_probe()` is side-effect free (`CLSIDFromProgID`/registry; never starts PowerPoint).

| # | Question | Answer (evidence) |
|---|----------|-------------------|
| 1 | Creates a private instance? | It *requests* one with `DispatchEx` (`qpn_renderer.py:278`, comment "always a NEW, application-owned instance"). PowerPoint cannot provide one: it is a single-instance automation server. |
| 2 | Attaches to a user instance? | Not deliberately (no `Dispatch`/`GetActiveObject`/`GetObject` anywhere). **Implicitly yes**: `DispatchEx` returns the running instance. |
| 3 | Can COM return/reuse an existing instance unexpectedly? | Yes — see evidence above; modelled by `FakeMachine.dispatch_ex`. |
| 4 | What owns each Presentation? | `pres` from `app.Presentations.Open(...)` (`:284`); closed in `finally` (`:296-300`). Safe in isolation, but opened *inside the shared instance*; also `Open()` on a path the user already has open returns the user's own presentation object, so `pres.Close()` would then close user work. |
| 5 | What owns the Application? | `created_app = True` is set **immediately after `DispatchEx` returns** (`:279`) with no proof — the ownership flag is an assumption, not evidence. |
| 6 | Where are Close()/Quit()? | `pres.Close()` `:298`; `app.Quit()` `:304`; both in the single `finally`. No other Close/Quit/kill in `app/`. |
| 7 | Ownership-aware? | Only by the unproven `created_app` flag (Application) and "opened by me" (Presentation). Neither checks that the process is actually ours or exclusive. Quit has no guard for presentations that the user opened later. |
| 8 | Broad POWERPNT.EXE termination? | **No.** `taskkill`/`TerminateProcess`/`psutil`/`.kill(`/`.terminate(` are absent from `app/`, `tools/`, `*.bat`. `app/updater.py:1131 os.kill(pid, 0)` is a POSIX-only liveness probe for the *old app's* PID (Windows branch uses `OpenProcess`/`GetExitCodeProcess`) — unrelated to PowerPoint, out of scope, untouched. LibreOffice `subprocess.run(..., timeout=300)` kills only its own child. |
| 9 | Can timeout/error recovery kill PowerPoint globally? | No kill path exists. But the error path runs the same unsafe `Quit()`. There is no timeout around COM calls at all; a hung COM call is not interrupted. |
| 10 | Can concurrent jobs confuse ownership? | Yes. `render_with_powerpoint` is reachable from the batch worker *and* the learning-preview path (`application_service`) concurrently. Job B's `DispatchEx` returns job A's instance (single instance) and sets `created_app=True`; whichever finishes first `Quit()`s the other's instance and closes its presentation. No in-process lock exists. |

Additional defect found on the cancel path: `except Exception` around the export loop wraps `CancellationRequested` (an `Exception`) into `PowerPointStageError("slide-export")`, misclassifying cooperative cancel as a renderer failure (violates constitution §VI and contract G). Proven by `test_6`.

## Root Cause Hypothesis (confidence: high for the code defect; medium-high for real-PowerPoint behaviour pending Windows acceptance)

`render_with_powerpoint` assumes `DispatchEx("PowerPoint.Application")` always yields a fresh, exclusively owned process and sets `created_app = True` unconditionally. Because PowerPoint is a single-instance automation server, that call returns the user's running instance whenever one exists; the function then opens a document into it, closes it, and `Quit()`s it, which closes every presentation the user has open (saved or unsaved, depending on `DisplayAlerts`). The same unproven assumption makes concurrent in-process renders able to quit each other, and a presentation the user opens in our hidden instance mid-render is also destroyed by our `Quit()`.

## Proposed Remediation

**Preferred — explicit ownership proof at the narrow COM boundary (`render_with_powerpoint` only):**

1. Serialize the whole COM lifecycle with a module lock (in-process jobs no longer collide).
2. **Pre-flight, fail closed:** snapshot running `POWERPNT.EXE` PIDs read-only (Toolhelp snapshot via `ctypes`; no handle to any process is opened, nothing can be signalled). Refuse (`PowerPointStageError("ownership")`, no activation at all) if the snapshot is unreadable or non-empty, or if `GetActiveObject` finds a registered instance (second, independent signal).
3. **Post-dispatch proof:** take ownership only if exactly one *new* PID appeared and the returned app holds zero presentations and zero windows. Otherwise release the reference untouched (no Open/Close/Quit).
4. Before `Open`, require `Presentations.Count == 0`; close only the object returned by our own `Open`.
5. **Before `Quit`, re-prove exclusivity:** quit only if `Presentations.Count == 0` after closing ours; if anything else is there (user opened a file in our instance), leave the process running. Count unreadable → do not quit.
6. After an authorised `Quit`, wait (bounded) for the owned PID to disappear; never escalate to process termination. Cleanup failures are logged only.
7. Re-raise `CancellationRequested` unchanged.
8. On refusal the existing renderer chain falls through to LibreOffice/built-in (already reported as reduced fidelity); the failure stage is `ownership` (not `availability`) so a degraded preview is retried, not pinned.

**Alternatives rejected:** (a) removing `Quit()` — leaks hidden PowerPoint processes; (b) `taskkill`/PID kill of PowerPoint — forbidden, can hit user processes; (c) attaching to the user's instance with a read-only windowless `Open` and `Close` only ours — permitted by the contract but cannot be made safe without Windows evidence (same-path `Open` returns the user's presentation; focus/alerts/add-ins in a shared session), so rejected for this fix; (d) custom out-of-process COM server per job — larger redesign, not needed.

**Files likely to change:** `app/qpn_renderer.py`; `tests/test_prompt027r_renderer.py` (adapt the existing fake-COM harness to the new ownership seams — assertions are not weakened); new `tests/test_prompt035_powerpoint_isolation.py`.

**Tests to add:** TEST 1–8 of the report (see the new test file) plus race, ROT-visible, unreadable-snapshot, concurrency-lock, cleanup-failure-no-kill and renderer-chain degradation tests.

## Risks & Considerations

- **Behavioural trade-off:** while the user's PowerPoint is running, extraction uses the LibreOffice/built-in renderer instead of PowerPoint (reduced fidelity, already surfaced in the UI) — correct under "fail closed". With PowerPoint closed, behaviour is unchanged.
- **Residual TOCTOU:** a user launching an *empty* PowerPoint within the milliseconds between our snapshot and `DispatchEx` returning is not distinguishable from a fresh instance by PID/presentations/windows alone. Mitigated (two pre-checks, post-dispatch checks, re-check before Quit) but not provably eliminated; **requires Windows acceptance** and is recorded as a residual risk.
- The Toolhelp snapshot code path only runs on Windows and is **unverified on Linux**; tests use an injected process table.
- Pre-existing unrelated failures at `502968c`: 3 version-identity tests in `tests/test_prompt027_rendering.py` still expect 1.3.5/Build 018 (release identity is 1.3.6/Build 019). Not modified (release identity out of scope).
- An owned instance whose `Quit()` fails is left running and logged (no kill); subsequent renders fail closed until it exits.
