# Bug Assessment: app reports a user PowerPoint running when no `POWERPNT.EXE` exists (PROMPT-041W)

- **Slug**: prompt041w-rot-false-positive
- **Created**: 2026-10-17
- **Source**: Windows acceptance run of PR #17 (`arena/d2fec0b6-tnp` @ `08b3de1`, PROMPT-041) — targeted automated tests 138 passed / 0 failed, **manual acceptance FAILED**
- **Verdict**: **valid, and a genuine correctness defect, not a wording defect.** The renderer asserted a machine fact it had never established. Root cause **proven by execution**, not inferred (see *Proof*).
- **Severity**: high — the message instructs the reviewer to close a program that is not running, so it is unactionable, and it blocks a PowerPoint-faithful render that was safe to perform. No data risk: the failure direction was over-refusal.
- **Spec Kit execution mode**: Executed from the checked-in Spec Kit skill instructions (`.agents/skills/speckit-bug-*`); native Arena skill invocation was not available. `.specify/extensions.yml` registers `hooks: {}`, so there are no pre/post hooks to run.
- **Base**: `arena/d2fec0b6-tnp` @ `08b3de1` (PROMPT-041). PR #17 open against `integration/prompts-034-035-036-037`; this bug **blocks that merge**.

## Report (summarized)

Windows, no PowerPoint running at all:

| Check | Result |
|---|---|
| `Get-Process POWERPNT -ErrorAction SilentlyContinue \| Select-Object Id, StartTime, SessionId, MainWindowTitle` | **zero rows** |
| PowerPoint open by hand? | no |
| `.pptx` open? | no |
| **Actual** | `Render: Trình vẽ tích hợp` + warning *"Một phiên PowerPoint đang chạy trên máy này nên app không tạo được phiên riêng. Đóng PowerPoint rồi chọn lại ảnh."* |
| **Expected** | PowerPoint backend attempts activation → `Render = PowerPoint`, `faithful = true`; if activation fails for another reason, the diagnostic must state *that* reason. Never `powerpoint-user-session-running` when the real process snapshot is empty. |

The tester's `Get-Process` output is **correct and is taken as ground truth.** The bug is in the app.

## Symptom

The Learning preview says a PowerPoint session is running on the machine. The message is uniquely attributable: `grep -rn "Một phiên PowerPoint đang chạy"` returns **exactly one hit repo-wide**, in `frontend/src/tabs/LearningTab.tsx` → `fidelityHint()`, and only the `powerpoint-registered-instance` case produces it. So the reason code that reached the UI was **`powerpoint-registered-instance`**, *not* `powerpoint-user-session-running` — the first thing the evidence establishes, and it rules out the same-session process check as the direct cause.

## Trace (as required, in order)

`_powerpoint_process_ids()` → `_current_session_id()` / `_process_session_id()` session filter → pre-activation snapshot in `_render_with_powerpoint_locked` → ownership refusal → `PowerPointStageError.code` → `SlideRenderer.render` fallback → `SlideRenderer.preview_reason` → DTO `slidePreviewReason` → `fidelityHint()` in the Learning UI.

Pre-fix pre-flight, in execution order:

```python
before = _powerpoint_process_ids()                    # frozenset()  → correct, agrees with Get-Process
if before is None:      raise … PROCESS_TABLE_UNREADABLE     # not taken (set, not None)
if before:              raise … USER_SESSION_RUNNING         # NOT taken (an empty set is falsy)
if _running_powerpoint_registered(win32com.client):
                        raise … REGISTERED_INSTANCE          # ← THE ONLY PATH THAT CAN FIRE
```

With a correct empty snapshot, only the **Running Object Table** lookup can refuse — and it did.

## Root cause — PROVEN, not speculated

`_running_powerpoint_registered()` called `win32com.client.GetActiveObject("PowerPoint.Application")` and read
`found is not None` as "a PowerPoint is running". Two independent facts make that reading false:

1. **The ROT outlives the process.** Microsoft's own guidance is that Office programs register themselves in the Running Object Table *when they lose focus*. Registration is torn down on a **graceful** exit. A crashed, killed, or abandoned instance — a PowerPoint closed by Task Manager, a previous Report Extractor run whose `Quit()` failed, a machine reboot with a hibernation image — leaves the entry behind with no process behind it.
2. **`GetActiveObject` can never report "nothing".** From `pywin32/com/win32com/client/__init__.py`:

   ```python
   def GetActiveObject(Class, clsctx=pythoncom.CLSCTX_ALL):
       resultCLSID = pywintypes.IID(Class)
       dispatch = pythoncom.GetActiveObject(resultCLSID)
       dispatch = dispatch.QueryInterface(pythoncom.IID_IDispatch)
       return __WrapDispatch(dispatch, Class, resultCLSID=resultCLSID, clsctx=clsctx)
   ```

   `__WrapDispatch` always returns a `CDispatch` wrapper — **never `None`**. The only "nothing registered" signal is a `pythoncom.com_error` carrying `MK_E_UNAVAILABLE` (`0x800401E3`). A lookup can also fail for unrelated reasons (COM not initialised on this thread, a bad marshalled proxy), and the pre-fix code collapsed *every* one of those into `False`, so it could not even distinguish "nothing registered" from "the lookup is broken".

So: a stale ROT entry, no process, and the renderer converted that into a user-facing **process claim**. Contract A violated.

### Proof by execution (not inference)

A throw-away harness (`tests/test_zz_repro041w.py`, run against the unmodified `08b3de1` tree, then deleted) drove the real renderer with the PROMPT-035 fake COM machine, an empty process table and a registered ROT object:

```
process_ids()   = frozenset()            ← matches the user's empty Get-Process POWERPNT
stage           = ownership
code            = powerpoint-registered-instance
dispatches      = 0                      ← activation was never attempted
last_backend    = builtin
preview_reason  = powerpoint-registered-instance
degraded        = True
```

Same harness, second defect:

```
_preview_expired(age=3 s, TTL=15.0) = False
```

The empty snapshot and the empty `Get-Process` agree; the only difference between the two machines is the ROT. That is the whole of the root cause.

## Second defect found while tracing — stale preview reason (contract E)

`_slide_preview_for` served a cached preview for the whole `DEGRADED_PREVIEW_TTL_S = 15.0` s window with **no expiry check other than age**, and the cache entry carries `reason`. Consequences:

- A preview rendered while PowerPoint *was* open keeps telling the reviewer "PowerPoint is running" for up to 15 s after they close it, on every poll.
- `LearningAssets.slide()` re-stamps `created` from `saved_at` (`time.monotonic() - (time.time() - saved_at)`), so a degraded preview persisted **to disk** is *not* expired after a restart either — it is served with its stale reason intact.
- The offline branch (`assets.offline_slide(...)`) applied no expiry check at all.

So even a correct pre-flight fix would have left a window in which the UI asserted a machine fact that had stopped being true.

## Third weakness — the lookup could not report its own failure

`_rot_lookup` collapsed three distinct states into one boolean: *a server is registered*, *nothing is registered*, and *the lookup itself could not run*. A diagnostic that cannot say which of those happened cannot be used to make a safety decision, and cannot be handed to a tester on Windows.

## Contracts, restated

| | Condition | Required behaviour |
|---|---|---|
| **A** | no POWERPNT (raw **and** same-session empty) | must **not** refuse as user-session-running; attempt activation, let the post-dispatch PID proof decide |
| **B** | POWERPNT in **this** session | refuse `powerpoint-user-session-running`, fall back to built-in, never attach (PROMPT-035, unchanged) |
| **C** | POWERPNT only in **another** session | not user-session-running here; then still perform the safe ownership proof |
| **D** | snapshot unreadable / session unprovable | fail closed with a **distinct** code, never dressed up as user-session-running |
| **E** | cached/persisted fallback recorded while PowerPoint was open | must not keep claiming "PowerPoint is running" once the machine has moved on |

## Explicit audit of everything the report asked to check

| Suspect | Finding |
|---|---|
| Toolhelp32 snapshot filtering | **correct.** `_enumerate_processes` matched `POWERPNT.EXE` and returned the true set; `None` on any failure. |
| Executable-name comparison | **correct.** case-insensitive compare against `_POWERPOINT_IMAGE_NAME = "powerpnt.exe"`. |
| `ProcessIdToSessionId` | **correct but its failure was invisible.** On failure it returns `None` and `_powerpoint_process_ids` conservatively keeps the PID (treats it as ours) — safe direction, but the caller then reported it as `powerpoint-user-session-running`, i.e. an *asserted* same-session fact. Contract D gap. |
| Current-process SessionId | read via `ProcessIdToSessionId(os.getpid())`; `None` → "cannot tell sessions apart" → all PIDs kept. Same D gap. |
| PID / session conversion | no truncation or sign issues (`ctypes` `DWORD` out-param). |
| Error vs `None` semantics | `_powerpoint_process_ids` → `None` = unknown (fail closed), `frozenset()` = none running. **Not confused**; verified. |
| **Empty set vs `None`** | **not confused** in the pre-flight — which is precisely why `if before:` was correctly skipped and the ROT check was reached. |
| Cached state | **defect** (contract E, above). |
| Stale preview reason / cache | **defect** (contract E, above). |
| Is the warning from an OLD cached preview? | **Both.** A fresh refusal produced it *and* the cache replayed it for 15 s. The repro proves the fresh-refusal path independently. |
| Is a previous refusal reason reused across candidates? | **No cross-talk.** `reason` is read per render from `renderer.preview_reason` on a fresh `SlideRenderer` built inside `_slide_preview_for`; pinned by regression 9 anyway. |
| 15-second retry / cache behaviour | `DEGRADED_PREVIEW_TTL_S = 15.0` s with an age-only test — see contract E. |
| Frontend store / state retaining `slidePreviewReason` | **No leak.** `const previewReason = cand.slidePreviewReason || ''` is read per render from the current candidate DTO; nothing is memoised in the store. The frontend is *not* the cause. |

## Explicitly NOT the cause

Not a process-detection bug. Not a frontend bug. Not a wrong `Get-Process` by the tester. Not `DispatchEx` (PROMPT-041 already proved it cannot isolate). The refusal was created in `_render_with_powerpoint_locked`, from the Running Object Table, before any activation was attempted.

## Fix direction

Make the **process table the only source of process facts** and demote the ROT to a three-state diagnostic; add a self-vouch so "empty" can be told apart from "could not see"; give the unprovable-session case its own code; and re-check a cached PowerPoint-blocked fallback against the live machine before serving it again. Full design in ./fix.md.
