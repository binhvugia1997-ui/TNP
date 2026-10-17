# Bug Fix: PowerPoint isolation is impossible in-session — refuse safely, diagnose precisely, stop over-refusing (PROMPT-041)

- **Slug**: prompt041-isolated-powerpoint-renderer
- **Fixed**: 2026-10-17
- **Assessment**: ./assessment.md
- **Status**: applied (automated/fake-COM proven). **The headline request is NOT delivered, deliberately**: an isolated PowerPoint automation instance cannot be created inside the user's Windows session, so the report's own stop clause applies and the fallback stays. Real Windows acceptance pending.
- **Spec Kit execution mode**: Executed from the checked-in Spec Kit skill instructions; native Arena skill invocation was not available. `.specify/extensions.yml` has `hooks: {}`.

## Summary

The requested "user PowerPoint open **and** Report Extractor renders with PowerPoint" is not implementable safely. PowerPoint is a **Multiuse (Single Instance)** COM server — Microsoft's own words — so `DispatchEx("PowerPoint.Application")` inside the user's session returns *their* `POWERPNT.EXE`; a second one is never started. The only documented route to a second instance is launching it as a *different Windows user*, which a desktop app must not do. Lifting the PROMPT-035 guard would therefore attach to the user's instance — the exact data-loss bug that guard exists to prevent — and would buy nothing.

What shipped instead is everything that *is* safe and was actually broken:

1. **The refusal now names its cause.** A stable slug vocabulary (`REASON_*`) rides on `PowerPointStageError.code`, on `SlideRenderer.preview_reason` and in the log, so "PowerPoint is installed but your own session is in the way" is never reported as "PowerPoint unavailable".
2. **The reviewer can act on it.** `slidePreviewReason` reaches the Learning DTO and `fidelityHint()` appends a cause + remedy to the existing non-blocking notice — "Máy này CÓ PowerPoint, nhưng bạn đang mở PowerPoint … Đóng PowerPoint rồi chọn lại ảnh để có bản xem trước trung thực cao."
3. **False refusals removed.** The PowerPoint process snapshot is now scoped to *this Windows session*, matching PowerPoint's real single-instance scope and COM's session-local activation. A PowerPoint in another session/RDP session no longer blocks a perfectly safe render.
4. **One greppable line per decision**: `POWERPOINT_ISOLATION state=created|refused reason=… detail=…`.

Nothing in the ownership model changed. No process is ever terminated, no registry is written, no instance is attached.

## Isolation mechanism — researched, and it does not exist

| Question | Answer | Evidence |
|---|---|---|
| Does `DispatchEx` create a second `POWERPNT.EXE`? | **No.** "Only one instance of PowerPoint (Powerpnt.exe) can run at any given time. Therefore, PowerPoint is a Multiuse (Single Instance) server." | learn.microsoft.com — *How to use Visual C# to automate a running instance of an Office program* |
| Any COM flag / context that changes this? | No. `Dispatch`/`DispatchEx`/`CoCreateInstance`/`GetActiveObject`/`BindToMoniker` all resolve to the one running server. | same source: "For a Multiuse (Single Instance) server like PowerPoint, it does not matter, because the automation reference points to the same running instance." |
| Two `POWERPNT.EXE` in Task Manager? | A zombie from a crash, not a supported second instance. | stackoverflow.com/questions/10955496 |
| Real isolation? | Only a **different Windows logon session** (`runas /user:…`), i.e. another account. | stackoverflow.com/questions/8389415, /questions/10955496, superuser.com/questions/1570107 |
| Can this app do that? | No — it would need another account's credentials, and the report forbids registry/Office-registration tricks. `test_2c` asserts `runas` / `CreateProcessWithLogon` / `CreateProcessAsUser` / `LogonUser` appear nowhere in `app/`. |

So there is no isolated instance to create and therefore nothing to prove ownership of. Per the report: **STOP, report the limitation, keep the fallback.**

## How ownership is decided (unchanged from PROMPT-035, plus session scope)

| | Rule |
|---|---|
| **App-owned Application** | process table readable **and** empty (this session) before activation; nothing in the Running Object Table; exactly ONE new `POWERPNT.EXE` PID after `DispatchEx`; the returned object is pristine (0 presentations, 0 windows). All four, or `owned` stays `False`. |
| **App-owned Presentation** | only the object returned by *our* `Presentations.Open`, on an instance re-proven pristine immediately before the call. |
| **User instance** | never dispatched to, never opened into, never closed, never quit, never killed, never enumerated. Refused *before* COM activation. |
| **Process termination** | none anywhere. The process table is read with a Toolhelp32 snapshot (no handle opened); the session lookup uses `ProcessIdToSessionId` (a query, not a handle). |
| **Concurrency** | `_POWERPOINT_LOCK` serialises the whole COM lifecycle. |
| **Doubt** | always resolves to "there is a PowerPoint" ⇒ refuse. |

## Lifecycle: CREATE → VERIFY → OPEN → EXPORT → CLOSE → QUIT → RELEASE

```
with _POWERPOINT_LOCK:                                    # serialised
  CoInitialize()
  before = _powerpoint_process_ids()                      # READ-ONLY, THIS SESSION
    None            -> refuse  powerpoint-process-table-unreadable
    non-empty       -> refuse  powerpoint-user-session-running     <-- the reported case
  GetActiveObject   -> found   -> refuse  powerpoint-registered-instance
  app = DispatchEx("PowerPoint.Application")              # CREATE (only reached when nothing was running)
  after = _powerpoint_process_ids(); created = after - before
    len(created) != 1 or not pristine -> refuse  powerpoint-ownership-unproven   (object released, never used)
  owned, owned_pid = True, the one new PID                # VERIFY OWNERSHIP (explicit token = owned_pid)
  log POWERPOINT_ISOLATION state=created reason=ok
  not pristine anymore -> owned = False; refuse  powerpoint-not-exclusive
  pres = app.Presentations.Open(path, ReadOnly=True, Untitled=False, WithWindow=False)   # OPEN read-only, hidden
  Slides(n).Export(png, "PNG", w, h)                      # EXPORT
finally:                                                  # success / error / timeout / cancellation alike
  pres.Close()                                            # CLOSE only what we opened
  if owned and Presentations.Count == (0 if closed else 1):
      app.Quit()                                          # QUIT only a proven-owned, still-exclusive instance
      wait <= 15 s for owned_pid to leave the table        # observe only; never terminate
  release refs; CoUninitialize()
  log every cleanup failure / skipped quit
```

The user's PowerPoint is outside this lifecycle in every branch: it is never the `app`, never the `pres`, and `owned` can only be `True` for a PID that did not exist before we activated.

## Fail-closed behaviour

| Situation | Outcome | Reason code |
|---|---|---|
| PowerPoint not installed / not Windows / no pywin32 | LibreOffice → built-in | `powerpoint-unavailable`, `powerpoint-import` |
| User PowerPoint open (the reported case) | LibreOffice → built-in | `powerpoint-user-session-running` |
| Process table unreadable | LibreOffice → built-in | `powerpoint-process-table-unreadable` |
| Instance registered in the ROT but invisible to the snapshot | LibreOffice → built-in | `powerpoint-registered-instance` |
| `DispatchEx` returned something unprovable | released untouched, LibreOffice → built-in | `powerpoint-ownership-unproven` |
| Someone else arrived in our instance | left running, LibreOffice → built-in | `powerpoint-not-exclusive` |
| `Open` / `Export` failed | owned-only cleanup, next backend | `powerpoint-presentation-open`, `powerpoint-slide-export` |
| `Quit` failed | logged, process left running, **never killed** | `SLIDE_RENDER_CLEANUP step=app-quit` |

A refusal keeps `stage="ownership"` (not `availability`), so `SlideRenderer.degraded` stays `True` and the preview is retried after `DEGRADED_PREVIEW_TTL_S = 15 s`: closing PowerPoint restores full fidelity without a restart. Safety still beats fidelity in every row.

## Diagnostics

- `POWERPOINT_ISOLATION state=created|refused reason=<slug> detail=<path-free text>` — one line per decision.
- `SLIDE_RENDER_BACKEND_FAILED … stage=ownership … detail=<slug>` — the slug now fills the previously empty `detail=-`.
- `SLIDE_RENDER_CLEANUP backend=powerpoint stage=cleanup step=<step> …` — unchanged, covers close/quit failures and skipped quits.
- `SLIDE_RENDER purpose=… backend=… selected_backend=… faithful=…` — unchanged.
- `powerpoint_render_capability() -> (available, isolated, reason, detail)` — read-only probe for support/diagnostics; it never initialises COM, never starts PowerPoint and never touches the ROT (so `isolated` may be optimistic, never pessimistic about `available`).
- The **diagnostics panel** (`--diag`, GUI "Chẩn đoán") gains a `PowerPoint isolation` row next to `PowerPoint renderer`. `PowerPoint renderer` keeps meaning "installed"; the new row says whether a faithful render can happen *right now* and, when it cannot but PowerPoint exists, tells the user to close their own PowerPoint. The two opposite causes produce two different rows.

## Learning UI

| Case | Backend chip | Notice |
|---|---|---|
| Isolated PowerPoint render | `Render: PowerPoint`, `faithful=true` | **none** — the "Bản xem trước đơn giản" warning is not shown |
| Refused because the user's PowerPoint is open | `Render: Trình vẽ tích hợp` | original sentence + "Máy này CÓ PowerPoint, nhưng bạn đang mở PowerPoint nên app không dùng chung phiên đó (để không đụng tới tài liệu của bạn). Đóng PowerPoint rồi chọn lại ảnh để có bản xem trước trung thực cao." |
| No PowerPoint on this machine | `Render: Trình vẽ tích hợp` | original sentence + "Máy này không có PowerPoint (hoặc thiếu thành phần COM của PowerPoint)." |
| Any other refusal | `Render: Trình vẽ tích hợp` | original sentence + the matching cause |
| Unknown/absent slug | `Render: Trình vẽ tích hợp` | original sentence only (never a crash) |

The notice stays a single, small, non-blocking banner; the slide bitmap and both overlays still render. Learning never crashes on an unknown slug.

## Production evidence impact

**None.** The change is confined to the PowerPoint backend's diagnostics, the process-snapshot scope and one new DTO field.

- `render_qpn_panel` / `render_qpn` (item-scoped After regions, QPN panel crop) — untouched; they call `SlideRenderer.render` exactly as before.
- Excel crop semantics (`improvement_visual.py`, `improvement_pictures.py`, `excel_writer.py`) — untouched; `test_15` asserts the new symbols never appear there and that no Excel module owns a PowerPoint lifecycle.
- PROMPT-020/021/025 report scoping and output eligibility — untouched.
- Learning preview and production evidence share the same `SlideRenderer` and therefore the same ownership rules, so the fix is safe for both by construction. `test_15` pins that separation.
- `SlideRenderer.availability()` is **unchanged** (`test_prompt027r_renderer.py:111` still asserts the exact 3-key dict), so preview cache identity and PROMPT-027R §30 freshness are unchanged. A faithful preview cached while PowerPoint was closed keeps being reused after the user opens PowerPoint — the new field is metadata, not identity.

## Changes

| File | Change | Notes |
|------|--------|-------|
| `app/qpn_renderer.py` | modified | `PowerPointStageError.code`; `REASON_*` vocabulary + `POWERPOINT_PRESENT_REASONS`; `PowerPointCapability` + `powerpoint_render_capability()`; `_powerpoint_process_ids()` split into a session-scoped filter over `_enumerate_powerpoint_pids()` with `_current_session_id()` / `_process_session_id()`; `_ownership_refusal(..., code=)`; `_log_isolation()`; codes on every refusal/`import`; `SlideRenderer.last_codes` + `preview_reason` + `_fail(code=)`. Public signatures unchanged. |
| `app/application_service.py` | modified | `_slide_preview_for` records `renderer.preview_reason` as `"reason"` (also on the no-bitmap path); DTO gains `slidePreviewReason`; cache/asset dicts documented. |
| `frontend/src/types.ts` | modified | `slidePreviewReason?: string` with the contract in the doc comment. |
| `frontend/src/tabs/LearningTab.tsx` | modified | exported `fidelityHint(slug)`; notice appends the cause/remedy. Original wording untouched. |
| `app/diagnostics.py` | modified | new `PowerPoint isolation` row from `powerpoint_render_capability()`; `renderer_status()` untouched. |
| `frontend/package.json` | modified | `test:prompt041` script; the new file added to `test:frontend`. |
| `tests/test_prompt041_isolated_powerpoint.py` | added | 33 test functions / 35 test items. Imports the PROMPT-035 fake machine so both prompts are tested against ONE model of Windows. |
| `frontend/tests/prompt041-preview-reason.test.mjs` | added | 3 tests: slug→text mapping, the banner in a real `LearningTab` render, and the no-notice PowerPoint case. |
| `docs/WINDOWS_ACCEPTANCE_PROMPT041.md` | added | Scenarios A–D with exact PowerShell commands. The PID/process script is **read-only** (`Get-Process`, `Get-CimInstance`) — it never kills anything. |
| `.specify/bugs/prompt041-isolated-powerpoint-renderer/{assessment,fix,test}.md` | added | workflow evidence |

**Not changed:** `tests/test_prompt035_powerpoint_isolation.py` and every other existing test; version/build (still 1.3.7 / Build 020); `ReportExtractor.spec`; `backup/…STABLE/`; the updater; the Excel pipeline.

## Diff Highlights

```python
# qpn_renderer.py — session-scoped, conservative in the safe direction
def _powerpoint_process_ids():
    pids = _enumerate_powerpoint_pids()
    if pids is None or not pids:
        return pids
    mine = _current_session_id()
    if mine is None:
        return pids                                    # cannot tell sessions apart: assume they are all ours
    return frozenset(pid for pid in pids if (_process_session_id(pid) in (None, mine)))

# refusal keeps stage="ownership" (PROMPT-035) and gains a stable cause
raise _ownership_refusal("a PowerPoint session is already running; not attaching to it", pptx,
                         code=REASON_USER_SESSION_RUNNING)
```

```tsx
// LearningTab.tsx
Bản xem trước đơn giản — bố cục có thể khác PowerPoint (backend: {backendLabel(previewBackend)}).
Kết quả Excel vẫn dùng đúng toạ độ PPTX.{fidelityHint(previewReason)}
```

## Tests Added or Updated

- `test_1_*` (required 1) — no user PowerPoint ⇒ owned instance created, used, closed, quit, gone; `POWERPOINT_ISOLATION state=created`.
- `test_2_*`, `test_2b`, `test_2c` (required 2, honestly) — user PowerPoint open ⇒ no dispatch, no open, no close, no quit, `code=powerpoint-user-session-running`; the single-instance fact is documented in the module (tripwire); no `runas`/logon-API anywhere in `app/`.
- `test_3_and_4_*` (required 3 + 4) — user presentation and unsaved edits survive refusal, export failure and Quit failure.
- `test_5_*`, `test_5b`, `test_5c` (required 5) — every unprovable-isolation path falls back to built-in with the right code; a machine without PowerPoint reports `powerpoint-unavailable` and is *not* degraded; the capability probe separates installed from isolatable.
- `test_6_*`, `test_6b` (required 6) — owned-instance export failure cleans only owned resources; post-dispatch ownership failure never opens/closes/quits.
- `test_7_*`, `test_7b` (required 7) — cleanup failure is logged, never escalated to a kill, never reaches the user.
- `test_8_*` (required 8) — 5 refused + 5 owned renders, no orphan process, `CoInitialize`/`CoUninitialize` balanced 10/10.
- `test_9_*` (required 9) — 3 concurrent renders, lock-serialised, 3 owned quits, no orphan.
- `test_10_*`, `test_10b` (required 10) — cancellation propagates unchanged, owned-only cleanup; pre-backend cancel never touches COM.
- `test_11_*` (required 11) — PROMPT-035's key assertions are read back out of its own source; the suite itself runs unchanged.
- `test_12_*` … `test_15_*` (required 12–15) — PROMPT-030B persistence round-trips the new field (including a pre-change asset with no field), PROMPT-038 freshness semantics unchanged, PROMPT-039/040 modules untouched.
- `test_16_*` … `test_16d` — session scoping: another session's PowerPoint is ignored; every doubt keeps the PID; the helpers are read-only queries.
- `test_17_*` — the reason reaching the DTO is a slug: no `/`, `\` or `:`.
- `test_18_*`, `test_19_*` — static contracts: no `taskkill`/`TerminateProcess`/`psutil`/`Stop-Process`/logon API anywhere in `app/`, no registry writes, no attach, no enumerate-Presentations, exactly one `.Close()` and one `.Quit()`.
- `test_20_*` — the slug vocabulary is stable and every value has wording in `LearningTab.tsx`.
- `test_21_*`, `test_22_*`, `test_23_*` — shutdown mid-export leaves no orphan and balanced COM; the Toolhelp handle is closed and `OpenProcess` never appears; the fake machine is shared with PROMPT-035.
- `test_24_*` — the diagnostics panel reports "PowerPoint installed but your session is in the way" and "no PowerPoint here" as two different rows.
- Frontend: `fidelityHint` mapping, one non-blocking banner with cause + remedy, and no notice for a genuine PowerPoint render.

## Local Verification

- `pytest tests/test_prompt041_isolated_powerpoint.py` → **35 passed**.
- `pytest tests/test_prompt035_powerpoint_isolation.py tests/test_prompt027r_renderer.py tests/test_prompt027_rendering.py tests/test_prompt027r_regions.py tests/test_prompt030b_integration.py tests/test_prompt038_desktop_freshness.py` → **153 passed**, none of those files edited.
- Full suite → **1355 passed, 13 skipped, 1 failed**; the one failure is `test_prompt028_portable.py::test_live_import_probe_reports_every_module_in_this_venv`, which needs the win32-only `pythonnet`/`clr_loader` and fails identically on the unmodified base (verified by re-running it with the changes stashed).
- `compileall app tools run.py` clean; `pyflakes` on the changed files reports only the 3 pre-existing deliberate "imported but unused" probes in `qpn_renderer.py`; `git diff --check` clean.
- Frontend: `test:frontend` **86 passed** (83 baseline + 3 new), `tsc --noEmit` clean, `vite build` succeeds.
- Pre-fix reproduction: the same test file as a characterization copy against unmodified `6346388` → **16 failed, 18 passed**.
- Mutations: removing the session filter fails `test_16`; dropping `code=REASON_USER_SESSION_RUNNING` fails `test_2`, `test_5[user-session]` and `test_17`. Both files restored byte-identical afterwards.

## Deviations from Assessment

- None. The assessment's preferred remediation is what shipped; the rejected alternatives stay rejected.

## Follow-ups

- **Windows acceptance (pending)** — see ./test.md and `docs/WINDOWS_ACCEPTANCE_PROMPT041.md`. This is the only thing that can confirm the real `POWERPNT.EXE`, the session-scoped `ctypes` paths and the end-to-end "the user's unsaved edit survives" claim.
- **Product decision (not made here):** a *consented*, read-only attach to the user's own PowerPoint is the only route to the requested fidelity while their session is open. This report excludes it; it would need its own safety design and its own Windows acceptance.
- **Product decision (not made here):** offering/bundling LibreOffice so the fallback is better than "simple preview".
