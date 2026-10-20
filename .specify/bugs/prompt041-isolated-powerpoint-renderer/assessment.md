# Bug Assessment: Learning preview loses PowerPoint fidelity while the user's own PowerPoint is open (PROMPT-041)

- **Slug**: prompt041-isolated-powerpoint-renderer
- **Created**: 2026-10-17
- **Source**: pasted text (PROMPT-041 report) + public Microsoft documentation of PowerPoint's COM server type (see Evidence)
- **Verdict**: valid as a **diagnostic/UX defect**, and **not implementable as asked** — an isolated PowerPoint automation instance cannot be created inside the user's Windows session, so the requested "user PowerPoint open + PowerPoint render at the same time" is refused by design and the honest fix is precise diagnostics + the safe fidelity wins, not a workaround
- **Severity**: medium — no data risk (the PROMPT-035 guard works exactly as intended); the cost is fidelity plus a misleading message that gives the reviewer no cause and no remedy
- **Spec Kit execution mode**: Executed from the checked-in Spec Kit skill instructions (`.agents/skills/speckit-bug-*`); native Arena skill invocation was not available. `.specify/extensions.yml` registers `hooks: {}`, so there are no pre/post hooks to run.
- **Base**: `integration/prompts-034-035-036-037` @ `6346388`. Session branch `arena/d2fec0b6-tnp` (fixed by Arena, branched from the requested base).

## Report (summarized)

Confirmed on Windows by the user:

- **No user PowerPoint running** → Learning preview uses the PowerPoint renderer, fidelity correct.
- **User PowerPoint running** → the PROMPT-035 ownership protection refuses PowerPoint and the chain falls back to the built-in renderer.

That protects the user's data but the UX is not acceptable in practice: the reviewer has to close PowerPoint to get a high-fidelity preview.

Requested: let `USER POWERPOINT OPEN` and `REPORT EXTRACTOR RENDERS WITH POWERPOINT` coexist, by giving Report Extractor its **own** PowerPoint automation instance/process. Absolutely no attach / reuse / close / kill / modification of the user's PowerPoint instance or documents. Mandatory pre-work: reproduce, trace the whole path, find the exact cause, write a regression test for the current behaviour, **research whether a truly isolated PowerPoint automation instance is possible on Windows/COM**, and implement only if ownership can be proven safe. Explicit stop clause: *if Windows COM offers no reliable way to create and prove an isolated instance, STOP, report the limitation and keep the current fallback; do not implement a dangerous workaround.*

## Symptom

With the user's PowerPoint open, the Learning preview is rendered by the built-in renderer and shows only "Bản xem trước đơn giản — bố cục có thể khác PowerPoint (backend: Trình vẽ tích hợp)". Expected by the report: a PowerPoint render. Expected by the safety contract: a fallback **that says why and what to do about it**.

## Reproduction

1. Windows machine, PowerPoint installed, Report Extractor built from `6346388`.
2. Start PowerPoint by hand, open any `.pptx`, leave it open (edited or not).
3. Open Report Extractor → Học cải tiến → Kiểm tra ảnh cải tiến → select any candidate.
4. The slide preview is the built-in render; the notice names no cause. The log shows
   `SLIDE_RENDER_BACKEND_FAILED purpose=learning_preview backend=powerpoint stage=ownership … reason=a PowerPoint session is already running; not attaching to it detail=-`.
5. Close PowerPoint, re-select the candidate (the degraded preview expires after `DEGRADED_PREVIEW_TTL_S = 15 s`) → PowerPoint render, `faithful=true`, no notice.

Automated (Linux, the PROMPT-035 fake single-instance COM machine), against the **unmodified** base:

```
tests/test_prompt041_isolated_powerpoint.py (as a characterization copy)
16 failed, 18 passed
AttributeError: 'PowerPointStageError' object has no attribute 'code'
AttributeError: 'SlideRenderer' object has no attribute 'preview_reason'
AttributeError: no capability probe in base
AttributeError: module 'app.qpn_renderer' has no attribute '_enumerate_powerpoint_pids'
AssertionError: qpn_renderer must keep documenting WHY an isolated PowerPoint instance cannot be created
```

and the behaviour itself, printed by a throw-away harness on the unmodified code:

```
stage         = ownership
reason        = a PowerPoint session is already running; not attaching to it
dispatches    = 0                 <- COM activation is never attempted
user quit     = 0 ; user docs = ['user_unsaved.pptx']   <- the guard works
availability  = {'powerpoint': True, 'builtin': True}   <- PowerPoint IS installed
last_backend  = 'builtin' ; degraded = True
SLIDE_RENDER_BACKEND_FAILED … stage=ownership … detail=-
```

PROVEN BY AUTOMATED REPRODUCTION: the refusal is the PROMPT-035 **pre-flight** ownership guard (`app/qpn_renderer.py`, `if before: raise _ownership_refusal(...)`), it fires before `DispatchEx` is called, and nothing machine-readable about its cause leaves the renderer — `detail=-`, no code, no DTO field.
REQUIRES WINDOWS/REAL POWERPOINT ACCEPTANCE: everything about a real `POWERPNT.EXE` (see `docs/WINDOWS_ACCEPTANCE_PROMPT041.md`).

## Full trace (Learning preview → cleanup)

| Step | Where | What it does |
|---|---|---|
| Learning preview | `app/application_service.py::_slide_preview_for` | cache identity = (prefer chain, width, `RENDER_SCHEMA_VERSION`, `SlideRenderer.availability()`, learning identity); on miss → `SlideRenderer.render(..., purpose="learning_preview")` |
| Backend selection | `app/qpn_renderer.py::SlideRenderer.render` | chain `("powerpoint", "libreoffice", "builtin")`; `powerpoint_available()` first |
| PowerPoint probe | `powerpoint_probe()` | `pythoncom.CLSIDFromProgID` then HKCR in default/64/32-bit views. **Side-effect free** — never starts PowerPoint. Reports `powerpoint: True` even while the user's session is open |
| COM activation | `render_with_powerpoint` → `_POWERPOINT_LOCK` → `_render_with_powerpoint_locked` | `CoInitialize` → process snapshot → **pre-flight refusal here** → `GetActiveObject` → `DispatchEx` |
| Ownership detection | `_powerpoint_process_ids()` (Toolhelp32, read-only) + `_running_powerpoint_registered` + `_is_pristine` | owned only if: table readable, **empty before**, nothing in the ROT, exactly one new PID after, 0 presentations / 0 windows |
| Open / export | `Presentations.Open(path, ReadOnly=True, Untitled=False, WithWindow=False)` → `Slides(n).Export(..., "PNG", w, h)` | only into a proven-owned pristine instance |
| Cleanup / fallback | the single `finally` | close only our presentation; `Quit()` only if owned **and** the remaining count is exactly what we expect; bounded wait for our PID; `CoUninitialize`; `PowerPointStageError` → next backend |
| UI | `slidePreviewBackend` / `slidePreviewFaithful` → `frontend/src/tabs/LearningTab.tsx` | generic notice; **no cause, no remedy** |

## Exact cause of "user PowerPoint present ⇒ ownership refusal"

`_render_with_powerpoint_locked` snapshots the PowerPoint process table **before** touching COM. If that set is non-empty it raises `PowerPointStageError("ownership")` immediately — before `DispatchEx`, before `GetActiveObject`, before `CoInitialize`'s work is used. So the mere *presence* of any `POWERPNT.EXE` is the trigger; it is not an availability failure and it is not the post-dispatch proof failing.

The guard is not accidental. It exists because of the platform fact below, and PROMPT-035's `test_1b` pins `machine.dispatches == 0` ("must not even touch COM activation while a PowerPoint session exists").

Two secondary false-refusal sources were found while tracing:

1. **The snapshot is machine-wide, but PowerPoint's single-instance scope is the login session.** `CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS)` lists every session, so a `POWERPNT.EXE` belonging to another logged-in user / RDP session refuses our render even though the interactive user has no PowerPoint. COM activation and the ROT are session-local, so that other-session process can never be handed to us.
2. **The refusal carries no machine-readable cause.** `stage="ownership"` is the only signal; `detail` is empty and the DTO has no field for it, so the UI cannot distinguish "no PowerPoint here" from "your PowerPoint is in the way".

## Evidence: an isolated PowerPoint automation instance does not exist inside one Windows session

Microsoft, *How to use Visual C# to automate a running instance of an Office program*
(learn.microsoft.com/en-us/previous-versions/office/troubleshoot/office-developer/use-visual-c-automate-run-program-instance):

> "When a request for a new COM object comes to a Multiuse (Single Instance) COM server, the server uses only one
> instance of the .exe file to create that object. No matter how many clients request a new COM object, there will
> be only one server .exe process. … Multiple instances of Word (Winword.exe), Excel (Excel.exe), and Microsoft
> Access (MSAccess.exe) can run simultaneously. Therefore, these servers are defined as Single Use (Multiple
> Instances) servers. **Only one instance of PowerPoint (Powerpnt.exe) can run at any given time. Therefore,
> PowerPoint is a Multiuse (Single Instance) server.** … For a Multiuse (Single Instance) server like PowerPoint,
> it does not matter, because the automation reference points to the same running instance."

Corroborating, independent:

- "PowerPoint only allows one instance of itself, so multiple simultaneous instances isn't possible, **at least on a per user login basis** … you might be able to make it fly by running multiple instances inside multiple login sessions." — stackoverflow.com/questions/8389415 (this is also the evidence for session-scoping the process snapshot).
- "PowerPoint only allows one instance of itself… If there's already an instance running, you'll get a reference to it." — stackoverflow.com/questions/10433507
- Two `POWERPNT.EXE` in Task Manager is a zombie, not a feature: "It only allows one instance of itself. If you see two instances of Powerpnt.exe in the task list, as sometimes happens, it means that something's gone wrong and left a zombie in memory." — stackoverflow.com/questions/10955496
- The one documented route to a second instance is **another Windows account**: `runas /user:username "…\POWERPNT.EXE"` — stackoverflow.com/questions/10955496, superuser.com/questions/1570107, ask.metafilter.com/47089.
- The same hazard, independently reported: attaching to `PowerPoint.Application` binds to the user's app and `Quit()` destroys the user's document — github.com/agentscope-ai/QwenPaw/issues/8002.

Consequences for the request:

| Option | Verdict |
|---|---|
| `DispatchEx("PowerPoint.Application")` while the user's PowerPoint is open | Returns **the user's instance**. No second process is created. Using it is exactly what PROMPT-035 forbids. |
| `Dispatch` / `GetActiveObject` / `BindToMoniker` | Explicit attach — forbidden by the report and by PROMPT-035 (`test_8c`). |
| A pre-created long-lived "app-owned" instance | Does not help: the user opening PowerPoint afterwards lands **inside our process**, so our `Quit()` gate would have to protect their document forever, and the render would still be refused from then on (our own leftover PID makes the pre-check non-empty). Strictly more dangerous than today. |
| Another logon session / `runas` / a second Windows account | Technically the only real isolation, and out of bounds: needs credentials the app must never handle, plus a per-account Office licence/first-run. Forbidden by the report's safety list in effect. |
| Registry / AppID / `RunAs` / COM surrogate tricks to force a second server | Forbidden ("không thay đổi registry/Office registration để ép COM behavior"). |
| Kill the user's PowerPoint, or kill "all POWERPNT" then start ours | Forbidden, and the exact data-loss bug PROMPT-035 was written for. |
| Windows Sandbox / a VM / App-V virtualized Office | Not something this desktop app can require of a target PC. |

**Conclusion (confidence: high, from the vendor's own description of the server type):** inside the user's Windows
session there is no isolated PowerPoint automation instance to create, so there is nothing to prove ownership of.
The report's stop clause applies. The refusal stays; what changes is that it stops being silent and stops
over-refusing.

## Root Cause Hypothesis

Two causes, both confirmed by the reproduction above.

1. **By design, correctly:** PowerPoint is a Multiuse (Single Instance) COM server, so the PROMPT-035 pre-flight
   guard must refuse whenever a PowerPoint session exists. The requested feature is not implementable safely.
2. **By defect:** that refusal is (a) reported without a machine-readable cause, so the UI shows the same
   "simple preview" text as a machine with no PowerPoint at all — the specific diagnostic error the report
   forbids — and gives no remedy; and (b) evaluated against a machine-wide process snapshot, so a PowerPoint in
   another Windows session causes a refusal that protects nothing.

## Proposed Remediation

**Preferred (implemented):** keep fail-closed, fix the diagnosis, remove the false refusals.

1. **Stable reason codes** on `PowerPointStageError.code` and a `REASON_*` vocabulary:
   `powerpoint-unavailable`, `powerpoint-import`, `powerpoint-process-table-unreadable`,
   `powerpoint-user-session-running`, `powerpoint-registered-instance`, `powerpoint-ownership-unproven`,
   `powerpoint-not-exclusive`, `ok`. `stage` is untouched, so every PROMPT-035 assertion still holds.
2. **`SlideRenderer.preview_reason`** + `last_codes` so the chain reports the faithful backend's own answer
   (the first code in chain order), and `detail=<code>` in `SLIDE_RENDER_BACKEND_FAILED`.
3. **`powerpoint_render_capability()`** — read-only `(available, isolated, reason, detail)` for diagnostics; no
   COM init, no PowerPoint start, no ROT lookup.
4. **Session-scoped snapshot** — `_enumerate_powerpoint_pids()` (unchanged Toolhelp reader) +
   `_current_session_id()` / `_process_session_id()` (`ProcessIdToSessionId`, a query, not a handle). Any doubt
   (unreadable own session, unreadable PID session, unreadable table) keeps the PID / returns `None` ⇒ refuse.
5. **`POWERPOINT_ISOLATION state=created|refused reason=… detail=…`** — one greppable line per decision,
   path-free and PID-free.
6. **`slidePreviewReason`** in the Learning DTO (a slug: no path, no PID), persisted with the preview and
   surviving a restart, mapped by `fidelityHint()` in `LearningTab.tsx` into an actionable sentence appended to
   the existing non-blocking notice.

**Alternatives rejected:** (a) lifting the pre-flight guard so `DispatchEx` is attempted anyway — on real Windows
it binds to the user's instance (a COM call into their process, with a hang risk if it is busy) and buys nothing;
it would also break PROMPT-035's `dispatches == 0` guarantee. (b) Attaching read-only to the user's instance
with the user's consent — explicitly excluded by this report. (c) Bundling LibreOffice to raise fallback
fidelity — a packaging decision, out of scope here (and `BUILD_AND_PUBLISH.bat` is off limits for this task).
(d) A pre-created long-lived owned instance — see the table above.

**Files likely to change:** `app/qpn_renderer.py`, `app/application_service.py`, `frontend/src/types.ts`,
`frontend/src/tabs/LearningTab.tsx`, `frontend/package.json`, new `tests/test_prompt041_isolated_powerpoint.py`,
new `frontend/tests/prompt041-preview-reason.test.mjs`, new `docs/WINDOWS_ACCEPTANCE_PROMPT041.md`.

**Tests to add or update:** the report's 15 mandatory items (item 2 becomes the honest, pinned contract plus the
platform-fact tripwire), plus session-scoping, capability probe, reason-code and DTO/frontend mapping tests, plus
the static no-kill / no-registry-write / no-attach contracts. **No existing test is weakened or edited.**

## Risks & Considerations

- **The headline request is not delivered.** It is refused with evidence rather than approximated. The reviewer's
  remedy is unchanged (close PowerPoint) but is now stated in the UI, and recovery is automatic: the refused
  preview is `degraded` and expires after 15 s, so closing PowerPoint restores fidelity without a restart.
- **Session scoping touches a safety-critical reader.** It can only ever *keep* a PID it is unsure about, so the
  failure mode is over-refusal (fidelity), never under-refusal (attaching to the user). The `ctypes` calls
  themselves are Windows-only and cannot be executed on Linux; the surrounding logic is tested with injected
  seams.
- **Residual TOCTOU (unchanged from PROMPT-035):** a user launching an *empty* PowerPoint between our snapshot
  and `DispatchEx` returning is not distinguishable by PID/presentations/windows alone. Still mitigated by the
  two pre-checks, the post-dispatch proof and the pre-`Quit` re-check; still requires Windows acceptance.
- **The reason slug is a cross-language contract.** `test_20` fails if a `REASON_*` value loses its wording in
  `LearningTab.tsx`.
- A leftover app-owned PowerPoint whose `Quit()` failed still blocks later renders until it exits (pre-existing,
  logged, never killed). Its reason is now visible instead of looking like "PowerPoint unavailable".

## Open Questions

- [NEEDS CLARIFICATION: product decision, not engineering] Is a *consented*, read-only attach to the user's own
  PowerPoint acceptable in a future prompt? It is the only way to the requested fidelity while their session is
  open, and this report explicitly excludes it. It would need its own safety design (no `Quit`, no `Close` of
  anything but our own presentation, `DisplayAlerts` handling, focus/active-presentation guarantees) and its own
  Windows acceptance.
- [NEEDS CLARIFICATION: product decision] Should LibreOffice be offered/bundled so the fallback is better than
  "simple preview"? Out of scope for this fix.
