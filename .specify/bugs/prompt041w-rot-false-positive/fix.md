# Bug Fix: the process table is the only source of process facts — the Running Object Table is a diagnostic (PROMPT-041W)

- **Slug**: prompt041w-rot-false-positive
- **Fixed**: 2026-10-17
- **Assessment**: ./assessment.md
- **Verification**: ./test.md
- **Status**: applied and proven against fakes (Python `1386 passed, 14 skipped, 1 pre-existing Linux-only failure`; frontend `89 passed`, typecheck clean, build clean). **Real Windows acceptance pending** — the checklist is `docs/WINDOWS_ACCEPTANCE_PROMPT041W.md`.
- **Spec Kit execution mode**: Executed from the checked-in Spec Kit skill instructions; native Arena skill invocation was not available. `.specify/extensions.yml` has `hooks: {}`.

## Summary

The app told the reviewer *"a PowerPoint session is running on this machine"* on a machine whose process table was
empty, because a **Running Object Table** entry was read as a process fact. The fix removes that inference entirely
rather than re-wording it:

1. **The process table is the only source of process facts.** The ROT is demoted to a three-state diagnostic
   (`registered` / `none` / `unavailable`) that can never, on its own, produce a user-facing claim about the machine.
2. **A self-vouch makes "empty" trustworthy.** A successful Toolhelp32 snapshot always contains *our own* PID. If it
   does not, the enumeration did not see the whole table and an empty answer is worthless. `_process_table_vouched()`
   is exactly that test, and it is what lets contract A and the PROMPT-035 fail-closed case coexist.
3. **An unprovable session gets its own code.** `powerpoint-session-unproven` — never `powerpoint-user-session-running`.
4. **A cached PowerPoint-blocked fallback is re-checked against the live machine** before it is served again, so a
   reason recorded while PowerPoint was open cannot outlive the condition that produced it.

Nothing in the ownership model was weakened. No attach, no `GetActiveObject` *use*, no process kill, no registry write.

## Files changed

| File | Change |
|---|---|
| `app/qpn_renderer.py` | the fix (new reason code, self-vouch, 3-state ROT, rewritten pre-flight, observability) |
| `app/diagnostics.py` | the `PowerPoint isolation` row said "đóng PowerPoint của bạn" for **every** PowerPoint-present reason — the same over-claim as the preview notice, on the panel support reads. It now says that only for the proven same-session case |
| `app/application_service.py` | contract E: `_degraded_reason_is_stale()` + the cache/persisted discard path + the `SLIDE_PREVIEW_CACHE` log line |
| `frontend/src/tabs/LearningTab.tsx` | wording for the new slug; the legacy slug's wording no longer asserts a process fact |
| `tests/test_prompt041w_rot_false_positive.py` | **new** — 33 deterministic regressions |
| `frontend/tests/prompt041w-preview-reason.test.mjs` | **new** — 3 frontend regressions, incl. a Python↔TS drift check |
| `frontend/package.json` | `test:prompt041w` script; the new file appended to `test:frontend` |
| `docs/WINDOWS_ACCEPTANCE_PROMPT041W.md` | **new** — the Windows retest checklist |

## 1. The refusal order, before and after

**Before** (`08b3de1`) — one ROT hit was enough to refuse, and it was reported as a process fact:

```python
before = _powerpoint_process_ids()
if before is None:  raise … PROCESS_TABLE_UNREADABLE
if before:          raise … USER_SESSION_RUNNING
if _running_powerpoint_registered(win32com.client):
                    raise … REGISTERED_INSTANCE          # ← fired with an EMPTY process table
```

**After** — the ROT can only refuse when the process table cannot vouch for itself:

```python
before = _powerpoint_process_ids()
facts = powerpoint_snapshot_facts(include_rot=True, client=win32com.client)
rot = facts.get("rot", ROT_NONE)
_log_snapshot(facts)                                     # ← the Windows evidence line

if before is None:                                       # contract D: table unreadable
    raise … PROCESS_TABLE_UNREADABLE
if before:
    if not facts["session_provable"]:                    # contract D: PowerPoint seen, our session unreadable
        raise … SESSION_UNPROVEN                         #   distinct code, never "user session"
    raise … USER_SESSION_RUNNING                         # contract B: proven same-session PowerPoint
if rot == ROT_REGISTERED and not facts["vouched"]:       # PROMPT-035: table could not see the whole picture
    raise … REGISTERED_INSTANCE                          #   fail closed — unchanged behaviour
# otherwise: the table vouched AND saw nothing → attempt activation (contract A)
app = win32com.client.DispatchEx(POWERPOINT_PROGID)
after = _powerpoint_process_ids()
created = None if after is None else after - before
if created is None or len(created) != 1 or not _is_pristine(app):
    raise … OWNERSHIP_UNPROVEN                           # the PID proof decides, not the ROT
```

Two things make this safe rather than merely permissive:

- The **post-dispatch PID proof is unchanged and still decisive.** Reaching the `DispatchEx` line only means "the
  machine we can see has no PowerPoint in our session". Ownership is then proven by *exactly one new PID* plus
  *pristine instance*, and the short-circuit means `_is_pristine(app)` is never evaluated when the PID proof fails,
  so no method is ever called on an instance we do not own.
- `_process_table_vouched()` is **False off Windows**, so on the PROMPT-035 fake machine a ROT hit still refuses with
  `dispatches == 0`. On a real Windows box it is True because our own PID is always in the table. That single
  difference is what satisfies contract A *and* keeps PROMPT-035 byte-identical.

## 2. The self-vouch

```python
def _process_table_vouched() -> bool:
    if sys.platform != "win32":
        return False
    every = _enumerate_processes(None)        # UNFILTERED: every process, not just POWERPNT
    return every is not None and os.getpid() in every
```

`_enumerate_processes(image_name=None)` was extracted from `_enumerate_powerpoint_pids()` so one Toolhelp walk can
answer both questions; `_enumerate_powerpoint_pids()` is now the thin wrapper `_enumerate_processes(_POWERPOINT_IMAGE_NAME)`.
**The wrapper is unchanged in name and signature, so every existing test that patches it keeps working.**

## 3. The Running Object Table, three states instead of one boolean

```python
ROT_REGISTERED, ROT_NONE, ROT_UNAVAILABLE = "registered", "none", "unavailable"

def _rot_lookup(client) -> str:
    try:
        found = client.GetActiveObject(POWERPOINT_PROGID)
    except Exception as exc:
        return ROT_NONE if _is_mk_e_unavailable(exc) else ROT_UNAVAILABLE
    return ROT_REGISTERED if found is not None else ROT_NONE
```

- `MK_E_UNAVAILABLE` (`0x800401E3`) is the ROT's **normal** "nothing registered" answer → `ROT_NONE`.
- Anything else (COM not initialised on this thread, a bad proxy, …) → `ROT_UNAVAILABLE`: *the lookup could not run*,
  which is a different fact from *nothing is registered* and must not be reported as if it were the same.
- `_running_powerpoint_registered(client)` is kept as a one-line compatibility wrapper (`_rot_lookup(client) == ROT_REGISTERED`).
  No test anywhere referenced it directly (grep-verified), so the return-type change is inert.
- The looked-up object is **never returned, stored or called.** It is a boolean-shaped fact and nothing more.

## 4. Contract E — a cached fallback cannot outlive its cause

`app/application_service.py`:

```python
@staticmethod
def _degraded_reason_is_stale(entry) -> bool:
    if not entry.get("degraded"):
        return False
    from .qpn_renderer import POWERPOINT_PRESENT_REASONS, powerpoint_render_capability
    if (entry.get("reason") or "") not in POWERPOINT_PRESENT_REASONS:
        return False
    return bool(powerpoint_render_capability(max_age_s=PREVIEW_CAPABILITY_MAX_AGE_S).isolated)
```

Applied on **both** reuse paths — the in-memory cache hit and the `LearningAssets` restore (whose re-stamped `created`
means `_preview_expired` alone would serve a stale reason after a restart). A stale entry is **popped and re-rendered
immediately** instead of waiting out `DEGRADED_PREVIEW_TTL_S`.

Two deliberate limits, so this does not become a COM hammer:

- Only reasons in `POWERPOINT_PRESENT_REASONS` are re-checked. A `powerpoint-unavailable` fallback (no PowerPoint on
  this PC) is *not* a machine fact that just changed, so the TTL still governs it — verified by regression 09b.
- `powerpoint_render_capability(max_age_s=PREVIEW_CAPABILITY_MAX_AGE_S)` (1.0 s) lets one Learning poll with many
  candidates take at most one extra process-table snapshot. `_CAPABILITY_MEMO` is only consulted when `max_age_s > 0`;
  **no-argument calls still re-read live**, so diagnostics and every existing test see current state.

## 5. Observability — what the Windows log must show

Two new greppable lines per render, PIDs and session ids only (no paths, no document names, no user content):

```
POWERPOINT_SNAPSHOT   pid=<ours> session=<ours|None> status=readable|unreadable raw=[…] same_session=[…]
                      vouched=true|false session_provable=true|false rot=registered|none|unavailable
POWERPOINT_OWNERSHIP  decision=created|refused reason=<slug> new_pids=[…] rot=…
```

plus the cache provenance line in the application log:

```
SLIDE_PREVIEW_CACHE hit=1 source=memory stale_reason=0|1 cached_backend=… cached_reason=… action=reuse|discard-and-rerender
```

and the pre-existing `POWERPOINT_ISOLATION state=created|refused reason=… detail=…`, which now also fires on the
refusal path together with `POWERPOINT_OWNERSHIP decision=refused`.

## 6. UI wording

`frontend/src/tabs/LearningTab.tsx` — `fidelityHint()` is the single producer of the notice (grep-verified: one hit
repo-wide).

| Slug | Wording now |
|---|---|
| `powerpoint-user-session-running` | **unchanged** — "Máy này CÓ PowerPoint, nhưng bạn đang mở PowerPoint … Đóng PowerPoint rồi chọn lại ảnh …" This is still correct and still actionable: here the process table proved it. |
| `powerpoint-session-unproven` | **new** — "Có tiến trình PowerPoint nhưng app không đọc được phiên Windows của chính mình nên từ chối dùng (an toàn cho tài liệu của bạn). Hãy thử render lại." |
| `powerpoint-registered-instance` | **re-worded** — no longer claims a running process. Kept at all only because previews *persisted before the fix* still carry the slug. "COM vẫn còn ghi nhận một phiên PowerPoint cũ trong khi không thấy tiến trình POWERPNT nào, nên app từ chối dùng để an toàn. Hãy thử render lại." |

The sentence the tester saw, `Một phiên PowerPoint đang chạy trên máy này…`, **no longer exists in the codebase**
(asserted by a test, not by inspection).

The **diagnostics panel** had the same defect in a second place: its `PowerPoint isolation` row rendered *every*
reason in `POWERPOINT_PRESENT_REASONS` as "…đóng PowerPoint của bạn để có bản xem trước trung thực cao". Only the
proven same-session case earns that instruction now; the others say what the app could not prove and tell the
reviewer to retry. `tests/test_prompt041_isolated_powerpoint.py::test_24` still passes untouched, because it
exercises the same-session case.

## 7. What was deliberately NOT done

| Rejected | Why |
|---|---|
| Re-wording the message without fixing the verdict | the report explicitly forbids patching only the Vietnamese text; the *verdict* was wrong |
| Deleting the ROT pre-check outright | the PROMPT-035 "instance invisible to the process table" case must keep failing closed |
| Calling a property on the ROT object to probe liveness | that **is** attaching — forbidden, and it can destroy an unsaved document |
| Adding a session/isolation key to `SlideRenderer.availability()` | feeds the preview-cache identity and breaks the exact-dict assertion at `tests/test_prompt027r_renderer.py:111` |
| Relaxing any PROMPT-035 assertion | explicitly forbidden; none was touched (regression 12 reads the file back and asserts that) |
| Treating the tester's `Get-Process` output as wrong | it was right; the app was wrong |

## 8. Safety invariants re-asserted

- Never attach / reuse / close / kill / modify the user's PowerPoint or documents. No `Dispatch`, no `GetObject`, no
  enumeration of `Presentations`, no `ActivePresentation`.
- No `OpenProcess` / `TerminateProcess` / `os.kill` / `subprocess`. Only `CreateToolhelp32Snapshot` + `Process32Next`
  + `ProcessIdToSessionId` + `CloseHandle`, all read-only queries.
- No registry writes.
- Fail closed on unprovable ownership: every doubt resolves towards "there is one" / "we cannot prove it".
- Learning preview and production evidence share one `SlideRenderer`, so the change is automatically safe for both.
- `pythoncom.CoInitialize` / `CoUninitialize` stay balanced on success, error, timeout and cancellation.
