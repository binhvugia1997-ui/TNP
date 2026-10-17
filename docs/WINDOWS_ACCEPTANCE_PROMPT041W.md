# Windows Acceptance — PROMPT-041W: PowerPoint false positive from a stale Running Object Table entry

**Branch** `arena/d2fec0b6-tnp` · **PR** #17 · **base** `integration/prompts-034-035-036-037` · version **1.3.7 / Build 020**

**Run the cases IN THIS ORDER. Do not skip to case B.** Case A is the reported bug and it must pass before anything
else is worth testing — if case A fails, stop and send the output back.

The fix, in one line: **the process table is now the only source of process facts.** A Running Object Table entry —
which Office leaves behind after a crash, a kill or an abandoned instance, and which `GetActiveObject` can never
report as "nothing" — is no longer allowed to become the claim *"a PowerPoint session is running on this machine"*.

Nothing here weakens the PROMPT-035 ownership guard: the app still never attaches to, opens into, closes, quits or
kills your PowerPoint, and it still fails closed whenever it cannot prove an instance is its own.

---

## 0. Prepare (once)

```powershell
cd C:\path\to\ReportExtractor        # the folder that contains ReportExtractor.exe
.\.venv-build\Scripts\python.exe -m pytest -q
```

Expected: the same result the agent got on Linux plus whatever your Windows box adds — **no new failures**. The
agent's Linux run was `1387 passed, 14 skipped, 1 failed`, the one failure being `test_live_import_probe_...`
which is expected to *pass* on Windows (it needs `pythonnet`).

Then, before every case, snapshot the machine:

```powershell
# (1) GROUND TRUTH — the process table, exactly as the tester saw it
Get-Process POWERPNT -ErrorAction SilentlyContinue | Select-Object Id, StartTime, SessionId, MainWindowTitle
# and our own session, so "same session" is comparable:
(Get-Process -Id $PID).SessionId
```

The five things to compare for every case are produced by one command after the render:

```powershell
# (2) raw snapshot  (3) same-session snapshot  (4) ownership decision  (5) resulting backend/reason
Get-Content .\logs\app.log -Tail 400 | Select-String "POWERPOINT_SNAPSHOT|POWERPOINT_OWNERSHIP|POWERPOINT_ISOLATION|SLIDE_RENDER_BACKEND_FAILED|SLIDE_PREVIEW_CACHE" | Select-Object -Last 12
```

Read the lines like this:

| Field | Meaning |
|---|---|
| `pid=` | Report Extractor's own PID |
| `session=` | Report Extractor's own Windows session id (`(Get-Process -Id $PID).SessionId`) |
| `status=` | `readable` = the Toolhelp32 snapshot worked; `unreadable` = the table could not be read (fail closed) |
| `raw=` | **every** `POWERPNT.EXE` PID on the machine, all sessions → compare with command **(1)** |
| `same_session=` | the subset in **our** session → this is what the ownership decision uses |
| `vouched=` | `true` = the snapshot contained our own PID, so it really did see the whole table |
| `session_provable=` | `true` = our own session id was readable, so "same session" is a fact we may assert |
| `rot=` | Running Object Table: `registered` / `none` / `unavailable` — **diagnostic only, never a process fact** |
| `decision=` | `created` = an app-owned instance was proven; `refused` = the app declined, `reason=` says why |
| `new_pids=` | the PID(s) that appeared after activation and were proven app-created |

`POWERPOINT_ISOLATION state=created|refused reason=<slug> detail=…` is the same decision in a form that is safe to
read over the phone. `SLIDE_PREVIEW_CACHE … action=reuse|discard-and-rerender` tells you whether the preview you are
looking at came from the cache and with which cached reason.

---

## CASE A — **NO PowerPoint running** (the reported bug — test this FIRST)

This is the exact scenario that failed. It must now render with PowerPoint.

### A1. Prove the machine is clean

```powershell
Get-Process POWERPNT -ErrorAction SilentlyContinue | Select-Object Id, StartTime, SessionId, MainWindowTitle
```
**Expected: zero rows.** If any row appears, close that PowerPoint properly first (File → Close) and re-run.

### A2. Read the Running Object Table by hand

This is what produced the false positive. A stale entry is *expected and harmless* now — it must not block the render.

```powershell
.\.venv-build\Scripts\python.exe -c "import pythoncom, win32com.client;`ntry:`n    o = win32com.client.GetActiveObject('PowerPoint.Application'); print('ROT: registered (STALE if Get-Process is empty)')`nexcept Exception as e:`n    print('ROT: none/unavailable ->', hex(getattr(e, 'hresult', 0) or 0))"
```

Record what it prints. **Either answer is fine for case A** — that is the whole point of the fix.

### A3. Diagnostics

```powershell
.\.venv-build\Scripts\python.exe -m app.main --diag --no-ollama-check
```

**Expected** `PowerPoint isolation` row:

```
OK – app sẽ tạo và sở hữu một phiên PowerPoint riêng khi render
```

**FAIL** if it says `Không dùng được (powerpoint-user-session-running)` or
`Không dùng được (powerpoint-registered-instance)` while `Get-Process POWERPNT` is empty.

### A4. Render in the UI

Report Extractor → **Học cải tiến** → **Kiểm tra ảnh cải tiến** → select any candidate. Wait for the preview.

**Expected (all four):**

| # | Check | Expected |
|---|---|---|
| 1 | `Render:` chip | **`PowerPoint`** |
| 2 | `faithful` | **`true`** (the reduced-fidelity banner does **not** appear) |
| 3 | warning text | **absent** — in particular *not* "Một phiên PowerPoint đang chạy trên máy này…" |
| 4 | log | `POWERPOINT_SNAPSHOT … status=readable raw=[] same_session=[] vouched=true rot=…` and `POWERPOINT_OWNERSHIP decision=created reason=ok new_pids=[<pid>]` |

**The `rot=` value does not matter here.** If `rot=registered` and the render still succeeds, that is precisely the
stale-ROT case that used to fail — please say so explicitly in your report, it is the strongest possible evidence.

### A5. Compare all five

```powershell
Get-Process POWERPNT -ErrorAction SilentlyContinue | Select-Object Id, StartTime, SessionId, MainWindowTitle   # (1) must still be zero rows
Get-Content .\logs\app.log -Tail 400 | Select-String "POWERPOINT_" | Select-Object -Last 6                    # (2)(3)(4)
```

| | Expected |
|---|---|
| (1) `Get-Process POWERPNT` | zero rows **before and after** — the instance the app created was quit again |
| (2) `raw=` | `[]` (or, if a *different-session* PowerPoint exists, that PID) |
| (3) `same_session=` | `[]` |
| (4) `decision=` | `created`, `reason=ok` |
| (5) backend / reason | `powerpoint` / (no reason) |

**CASE A PASSES only if all five agree.** Any `decision=refused` with `same_session=[]` is a fail — send the log.

---

## CASE B — the user's PowerPoint IS open (only after A passes)

```powershell
Start-Process powerpnt            # then open any .pptx and LEAVE IT OPEN, edited or not
Get-Process POWERPNT -ErrorAction SilentlyContinue | Select-Object Id, StartTime, SessionId, MainWindowTitle
(Get-Process -Id $PID).SessionId   # must equal the POWERPNT SessionId for this case to be meaningful
```

Render in the UI as in A4.

**Expected:**

| # | Check | Expected |
|---|---|---|
| 1 | `Render:` chip | `Trình vẽ tích hợp` (built-in fallback) |
| 2 | warning | "…Máy này **CÓ** PowerPoint, nhưng bạn đang mở PowerPoint nên app không dùng chung phiên đó… **Đóng PowerPoint** rồi chọn lại ảnh…" |
| 3 | log | `same_session=[<your POWERPNT pid>]` · `decision=refused reason=powerpoint-user-session-running` · `new_pids=[]` |
| 4 | **your document** | still open, **unsaved edits intact**, no save prompt, PowerPoint still running afterwards |
| 5 | `Get-Process POWERPNT` | **the same PID as before** — the app must not have quit, closed or restarted it |

Then close PowerPoint properly and re-render **within 15 seconds** (contract E):

**Expected:** `SLIDE_PREVIEW_CACHE … stale_reason=1 cached_reason=powerpoint-user-session-running action=discard-and-rerender`, and the very next preview is `Render: PowerPoint`, `faithful=true`, **no warning**. The stale notice must not linger for the rest of the degraded-TTL window.

---

## CASE C — PowerPoint in a DIFFERENT Windows session (only if you can)

Open a second interactive/RDP session as another user, start PowerPoint **there**, leave it running, and render in
the first session.

**Expected:** the render is **not** blocked as `powerpoint-user-session-running`. The log must show
`raw=[<other session pid>]` but `same_session=[]`, then `decision=created reason=ok`.

---

## CASE D — fail-closed paths (optional, needs a tool)

These must be reported with their **own** code, never as "your PowerPoint is running":

| Situation | Expected `reason=` |
|---|---|
| process table cannot be read | `powerpoint-process-table-unreadable` |
| a POWERPNT is seen but our own session id is unreadable | `powerpoint-session-unproven` |
| a ROT entry exists **and** `vouched=false` | `powerpoint-registered-instance` (fail closed; legacy wording, no "close PowerPoint" instruction) |
| the instance returned by activation is not proven app-created | `powerpoint-ownership-unproven` |

In every case: no PowerPoint is quit, no document is touched, and the UI must **not** claim a user PowerPoint is
running unless `same_session=` actually contains a PID.

---

## Report back

Paste, verbatim:

1. the `Get-Process POWERPNT` output for case A (must be empty),
2. the `GetActiveObject` line from A2,
3. the `PowerPoint isolation` diagnostics row from A3,
4. the last 6 `POWERPOINT_*` / `SLIDE_PREVIEW_CACHE` log lines from A5,
5. the `Render:` chip and whether the warning banner appeared,
6. for case B, the PID of your PowerPoint before and after the render.

**Do not merge PR #17 until case A passes on Windows.**
