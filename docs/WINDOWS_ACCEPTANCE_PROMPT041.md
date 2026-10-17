# PROMPT-041 — Windows isolation acceptance (NOT executed in the Linux agent)

Accepted base: `integration/prompts-034-035-036-037` @ **`6346388`**.
Review branch: **`arena/d2fec0b6-tnp`**. Check `git log -1` against the review commit/PR before testing — do not
infer acceptance from an older base or from Linux tests.

**Read this first — what PROMPT-041 does and does not claim.**

PROMPT-041 asked for "user PowerPoint open **and** Report Extractor renders with PowerPoint" via an app-owned,
isolated PowerPoint automation instance. **That is not delivered, on purpose.** PowerPoint is a **Multiuse
(Single Instance)** COM server: inside one Windows session `DispatchEx("PowerPoint.Application")` returns the
user's running `POWERPNT.EXE`, and a second one is never started (Microsoft:
<https://learn.microsoft.com/en-us/previous-versions/office/troubleshoot/office-developer/use-visual-c-automate-run-program-instance>
— "Only one instance of PowerPoint (Powerpnt.exe) can run at any given time"). The only documented route to a
second instance is launching it as a **different Windows user**, which this app must never do. So there is no
isolated instance to create and nothing to prove ownership of; the report's own stop clause applies and the
PROMPT-035 fallback stays.

What **must** be true on Windows, and what this checklist verifies:

1. The user's PowerPoint instance, documents, unsaved edits, active presentation and window are untouched.
2. The refusal is reported as **`powerpoint-user-session-running`**, never as "PowerPoint unavailable".
3. The Learning notice says the cause and the remedy ("đóng PowerPoint…").
4. No orphan `POWERPNT.EXE` accumulates across repeated renders.
5. With no user PowerPoint, the PowerPoint renderer still works exactly as in the accepted PROMPT-030B run.

Rules for the tester: never `Stop-Process`, `taskkill` or otherwise terminate PowerPoint as a cleanup shortcut;
never modify Office registration or the registry; never force `sys.platform` or mock availability during a real
acceptance; do not build/publish a Portable, do not bump version/build. Use copies of real reports and a
disposable output workbook. All observation commands below are **read-only**.

## 1. Confirm revision and prerequisites

```powershell
Set-Location 'D:\arena\TNP\TNP'          # your real checkout
if ((git branch --show-current).Trim() -ne 'arena/d2fec0b6-tnp') { throw 'Wrong acceptance branch; stop.' }
git status --short
git log -1 --format='%H %s'
git merge-base HEAD 6346388786ea39a72b19b471d793484cc9c27838
# Must print 6346388786ea39a72b19b471d793484cc9c27838. Do not overwrite a dirty checkout.

$py = '.\.venv-webview\Scripts\python.exe'   # or '.\.venv\Scripts\python.exe'
& $py --version
& $py -m pip install -r requirements-dev.txt -r requirements-webview.txt
& $py -m pip check
npm.cmd --prefix frontend ci --no-audit --no-fund

& $py -c 'import sys; from app import __version__, BUILD_ID; from app.qpn_renderer import powerpoint_probe, renderer_status, powerpoint_render_capability as cap; print(sys.platform, __version__, BUILD_ID); print(powerpoint_probe()); print(renderer_status()); print(cap())'
```

Expected: `win32`, **1.3.7 / 020** (unchanged), `powerpoint_probe()` still the tuple API `(available, stage, detail)`,
and `powerpoint_render_capability()` printing `PowerPointCapability(available=True, isolated=…, reason=…, detail=…)`.
With **no** PowerPoint running it must print `isolated=True reason=ok`.

### Diagnostics panel (the support-facing view of the same answer)

```powershell
& $py -m app.main --diag --no-ollama-check
```

Two rows must be present and must mean different things:

- `PowerPoint renderer : OK` — PowerPoint is **installed** (this row's meaning is unchanged).
- `PowerPoint isolation : OK – app sẽ tạo và sở hữu một phiên PowerPoint riêng khi render` when nothing is running,
  or `PowerPoint isolation : Không dùng được (powerpoint-user-session-running) – PowerPoint CÓ sẵn trên máy; đóng
  PowerPoint của bạn để có bản xem trước trung thực cao` while your PowerPoint is open.

Run it twice — once with your PowerPoint closed and once with it open — and record both outputs. A machine with no
PowerPoint must instead read `PowerPoint renderer : Unavailable` and
`PowerPoint isolation : Không dùng được (powerpoint-unavailable)`.

### Read-only process observation (the only process commands allowed)

```powershell
# Snapshot BEFORE / AFTER every scenario. Read-only: lists, never signals, never terminates.
Get-Process POWERPNT -ErrorAction SilentlyContinue |
  Select-Object Id, ProcessName, StartTime, @{n='SessionId';e={$_.SessionId}}, MainWindowTitle |
  Format-Table -AutoSize | Out-String -Width 200 | Tee-Object -Append .\logs\prompt041-pids.txt

# Per-process detail incl. session id and owner (read-only CIM query).
Get-CimInstance Win32_Process -Filter "Name='POWERPNT.EXE'" |
  Select-Object ProcessId, SessionId, CreationDate, CommandLine | Format-List
```

Record: your PowerPoint PID(s), session id, main window title, the open document name, and whether it has an
unsaved edit. Keep the file out of any public PR (it contains your machine's paths).

## 2. Targeted tests (exact commands)

```powershell
& $py -m pytest -q tests/test_prompt041_isolated_powerpoint.py
if ($LASTEXITCODE -ne 0) { throw 'PROMPT-041 targeted tests failed.' }

& $py -m pytest -q tests/test_prompt035_powerpoint_isolation.py tests/test_prompt027_rendering.py tests/test_prompt027r_renderer.py tests/test_prompt027r_regions.py tests/test_prompt030b_integration.py tests/test_prompt032_learning_state.py tests/test_prompt036r_bridge_hang.py tests/test_prompt036s_learning_latency.py tests/test_prompt038_desktop_freshness.py tests/test_prompt040_excel_persistence.py tests/test_application_service.py
if ($LASTEXITCODE -ne 0) { throw 'PROMPT-035/027/030B/038/040 regression failed.' }

# PROMPT-039 updater regression, exact cases:
& $py -m pytest -q tests/test_image_learning.py::test_28_apply_update_keeps_learning_data_and_rollback_restores_it tests/test_updater.py::test_updater_waits_for_main_process_exit
if ($LASTEXITCODE -ne 0) { throw 'PROMPT-039 retest failed.' }

npm.cmd --prefix frontend run test:prompt041
npm.cmd --prefix frontend run test:prompt027
npm.cmd --prefix frontend run test:prompt027r
npm.cmd --prefix frontend run test:prompt030b
```

`tests/test_prompt041_isolated_powerpoint.py` must report **34 passed**, and
`tests/test_prompt035_powerpoint_isolation.py` must be **unmodified** (`git diff --stat -- tests/test_prompt035_powerpoint_isolation.py`
prints nothing). Those COM tests use controlled fakes even on Windows: they establish regression behaviour, **not**
real Office ownership/fidelity acceptance. Scenarios A–D below are the real acceptance.

## 3. Full suite and fresh frontend

```powershell
& $py -m pytest -q -ra
if ($LASTEXITCODE -ne 0) { throw 'Full Python suite failed.' }
npm.cmd --prefix frontend run test:frontend
if ($LASTEXITCODE -ne 0) { throw 'Frontend suite failed.' }
npm.cmd --prefix frontend run typecheck
if ($LASTEXITCODE -ne 0) { throw 'Typecheck failed.' }
npm.cmd --prefix frontend run build
if ($LASTEXITCODE -ne 0) { throw 'Frontend build failed.' }
& $py -m compileall -q app tools run.py
& $py -m pyflakes app/qpn_renderer.py app/application_service.py tests/test_prompt041_isolated_powerpoint.py
git diff --check
& $py -c 'from app.desktop import frontend_dist_freshness; print(frontend_dist_freshness())'
```

Record counts, every skip reason and the exact commit. `frontend/dist` is local/ignored and must be rebuilt after
source changes (the PROMPT-038 stale-dist guard is intentionally unchanged); do not commit it.

Launch the real desktop app (so a missing WebView runtime cannot silently fall back to the legacy Tk UI):

```powershell
& $py -m app.desktop --debug
```

## 4. SCENARIO A — USER POWERPOINT OPEN (the reported case)

1. Start PowerPoint by hand. Open a `.pptx` of your own that is **not** used by Report Extractor.
2. Type something into a slide and **do not save**. Leave the window visible and in focus.
3. Record the baseline: PID, session id, main window title, document name, "unsaved edit present = yes".

```powershell
Get-Process POWERPNT -ErrorAction SilentlyContinue | Select-Object Id, SessionId, StartTime, MainWindowTitle
```

4. Open Report Extractor ( Học cải tiến → Kiểm tra ảnh cải tiến ). Select **several** candidates, on more than one
   report if available. Wait for each preview.
5. **Expect**: the backend chip reads `Render: Trình vẽ tích hợp`, and the notice reads
   "Bản xem trước đơn giản — bố cục có thể khác PowerPoint (backend: Trình vẽ tích hợp). Kết quả Excel vẫn dùng
   đúng toạ độ PPTX. Máy này CÓ PowerPoint, nhưng bạn đang mở PowerPoint nên app không dùng chung phiên đó …
   Đóng PowerPoint rồi chọn lại ảnh để có bản xem trước trung thực cao."
   It must **not** say PowerPoint is missing. The notice must stay small and non-blocking; the slide image and
   both overlays still render; Learning does not crash.
6. **Expect in the log**:

```powershell
Get-ChildItem .\logs -Filter '*.log' -ErrorAction SilentlyContinue |
  Select-String -Pattern 'POWERPOINT_ISOLATION|SLIDE_RENDER_BACKEND_FAILED|SLIDE_RENDER_CLEANUP|SLIDE_RENDER '
```

   - `POWERPOINT_ISOLATION state=refused reason=powerpoint-user-session-running detail=a PowerPoint session is already running; not attaching to it`
   - `SLIDE_RENDER_BACKEND_FAILED purpose=learning_preview backend=powerpoint stage=ownership … detail=powerpoint-user-session-running`
   - **no** `SLIDE_RENDER_CLEANUP … step=app-quit`, and **no** `POWERPOINT_ISOLATION state=created`.
7. **Verify the user instance is untouched**:
   - the same PID is still running, with the same `StartTime`;
   - the document is still open, still the **active** presentation, still with the unsaved edit (the modified
     marker is still there);
   - the window was not closed, minimized, hidden or re-ordered;
   - **no** Save / "Do you want to save changes" dialog appeared at any point;
   - no second `POWERPNT.EXE` appeared.
8. Repeat steps 4–7 after moving to another candidate and after switching reports.

**FAIL** if: PowerPoint is closed or quit; the unsaved edit is gone; a Save prompt appears; the active
presentation changes; a second `POWERPNT.EXE` appears; the message says PowerPoint is unavailable; or the app
reports `backend=powerpoint` while your instance is the one that rendered it.

## 5. SCENARIO B — NO USER POWERPOINT

1. Close your PowerPoint normally through its own UI (save or discard as you wish).
2. Confirm the process table is empty:

```powershell
Get-Process POWERPNT -ErrorAction SilentlyContinue
```

3. In Report Extractor, re-select the same candidates. The refused preview is `degraded` and expires after
   **15 s**, so either wait ~15 s or restart the app.
4. **Expect**: `Render: PowerPoint`, no "Bản xem trước đơn giản" notice at all, and in the log
   `POWERPOINT_ISOLATION state=created reason=ok detail=exactly one new PowerPoint process, pristine` followed by
   `SLIDE_RENDER purpose=learning_preview backend=powerpoint … selected_backend=powerpoint faithful=true`.
5. Compare the preview against the original slide (authored Before/After captions, arrows, shapes, fonts, the
   previously reported black-background case). This must match the accepted PROMPT-030B fidelity result.
6. **Owned cleanup**: after the render, the app-created instance must be gone:

```powershell
Get-Process POWERPNT -ErrorAction SilentlyContinue | Select-Object Id, StartTime
# and
Get-ChildItem .\logs -Filter '*.log' | Select-String -Pattern 'SLIDE_RENDER_CLEANUP'
```

   Expect no lingering `POWERPNT.EXE` created by the app, and no `app-exit-wait-timeout` note. If a
   `step=app-quit` failure is logged, the process may legitimately stay running — it must **not** be killed by
   the app, and you must record it rather than terminate it as a shortcut.

## 6. SCENARIO C — FORCED ISOLATION FAILURE

Prove the fail-closed path and the diagnostic without ever touching your instance.

1. With your PowerPoint **closed**, start PowerPoint again with a document open (so a session exists), then run a
   Learning preview: this is the natural "isolation unavailable" case — expect
   `reason=powerpoint-user-session-running`, the built-in renderer, the correct notice, and your instance intact.
2. Read-only capability check while your PowerPoint is open (must report installed **and** not isolatable):

```powershell
& $py -c 'from app.qpn_renderer import powerpoint_render_capability as cap; c = cap(); print(c)'
# Expect: PowerPointCapability(available=True, isolated=False, reason='powerpoint-user-session-running', detail=...)
```

3. Optional, only if you can do it without touching Office registration: make `Presentations.Open` fail (for
   example point the run at a `.pptx` that is locked by another program) and confirm
   `SLIDE_RENDER_BACKEND_FAILED … stage=presentation-open … detail=powerpoint-presentation-open`, an owned-only
   cleanup, and a built-in fallback.
4. **Never** simulate this by killing PowerPoint, editing the registry, or editing `sys.platform`.

**FAIL** if any of these paths attaches to, opens into, closes, quits or kills your instance, or reports
"PowerPoint unavailable" while PowerPoint is installed.

## 7. SCENARIO D — REPEATED RENDER (orphan watch)

1. Close your own PowerPoint (Scenario B conditions).
2. Record the baseline process list, then render **at least 10** previews: several candidates, several slides,
   switching candidates quickly, and at least one cancel (switch report / press cancel mid-run).
3. After each batch and after a normal app exit, re-run the read-only snapshot and compare PID sets.

```powershell
Get-Process POWERPNT -ErrorAction SilentlyContinue | Select-Object Id, StartTime, SessionId
Get-ChildItem .\logs -Filter '*.log' | Select-String -Pattern 'POWERPOINT_ISOLATION|SLIDE_RENDER_CLEANUP'
```

4. **Expect**: at most one app-created `POWERPNT.EXE` alive at a time (renders are serialised), the count back to
   zero after the app exits, one `state=created` per successful render, and balanced `CoInitialize`/`CoUninitialize`
   (no COM error strings in the log).
5. Repeat the whole batch with your own PowerPoint **open**: expect zero `state=created`, zero new
   `POWERPNT.EXE`, and your single PID unchanged throughout.

**FAIL** if `POWERPNT.EXE` processes accumulate, if a second instance appears next to yours, or if your PID
disappears.

## 8. Second-session check (optional, only with a second Windows account you control)

PowerPoint is single-instance **per login session**, and PROMPT-041 scopes the process snapshot to the app's own
session. If you can sign in a second account (or open a second RDP session) and start PowerPoint there:

1. Sign in as user B, start PowerPoint there, leave it running.
2. As user A (no PowerPoint of your own), run a Learning preview.
3. **Expect**: `isolated=True`, `POWERPOINT_ISOLATION state=created`, `backend=powerpoint`, `faithful=true` — i.e.
   user B's PowerPoint no longer causes a false refusal.
4. Confirm user B's PowerPoint PID and document are untouched.

If you cannot set this up, record it as **not tested** — do not mark it passed.

## 9. Record for the report

| Item | Value |
|---|---|
| Commit / branch | |
| Windows + Office version, 32/64-bit | |
| `powerpoint_probe()` / `powerpoint_render_capability()` output | |
| `--diag` rows, PowerPoint closed vs open (`PowerPoint renderer` + `PowerPoint isolation`) | |
| Python / frontend test counts (+ every skip reason) | |
| Scenario A: user PID before/after, unsaved edit survived, no Save prompt, notice text | |
| Scenario A: log lines (`POWERPOINT_ISOLATION`, `SLIDE_RENDER_BACKEND_FAILED … detail=`) | |
| Scenario B: `backend=powerpoint`, `faithful=true`, no notice, fidelity vs original | |
| Scenario C: forced-failure reason code, user instance untouched | |
| Scenario D: PID counts before/during/after, orphan count | |
| Second-session result (or "not tested") | |
| Anything not verified | |

Do not attach private report paths or content to a public PR.
