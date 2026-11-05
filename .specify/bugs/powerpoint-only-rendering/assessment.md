# Bug Assessment: PowerPoint-only rendering — silent fallback to the built-in renderer

- **Slug**: powerpoint-only-rendering
- **Created**: 2026-11-04
- **Source**: pasted text (user report in chat, 2026-11-04; user log + screenshot evidence)
- **Verdict**: valid
- **Severity**: high (user's standing requirement from the start of the build: rendering MUST use PowerPoint; the preview and the Excel evidence regions currently degrade silently to the built-in renderer, which the user rejects)

## Report (verbatim or summarized)

User (translated): "There is no renderer setting anywhere. From the start of the build I already required PowerPoint to be used." Screenshot showed the learning preview banner: "Bản xem trước đơn giản — bố cục có thể khác PowerPoint (backend: Trình vẽ tích hợp). Kết quả Excel vẫn dùng đúng tọa độ PPTX." ("Simplified preview — layout may differ from PowerPoint (backend: built-in renderer).") The user's log showed many `~$` PowerPoint lock files at the time, i.e. PowerPoint was open with files. PowerPoint IS installed on the machine.

## Symptom

When PowerPoint cannot be used (most commonly: a PowerPoint user session is already running), slide rendering silently falls through to LibreOffice / the built-in renderer for the learning preview and the Excel After-evidence regions, instead of stopping with a clear error. There is no user-facing renderer setting, and the only signal is a reduced-fidelity banner.

## Reproduction

1. Open PowerPoint with any report file open (creates `~$` lock files) — or otherwise make the PowerPoint backend unavailable (ProgID unregistered / pywin32 missing).
2. Open the app, open a report, view a slide in the learning preview (or export After evidence).
3. Observed: preview banner reports backend "Trình vẽ tích hợp" (built-in renderer); rendering proceeds instead of failing closed.

[NEEDS CLARIFICATION: live reproduction on Windows was NOT performed — this Linux sandbox has no PowerPoint/COM. Reproduction evidence is code trace + the user's log/screenshot. Windows acceptance (close-PowerPoint-then-retry, refusal reason codes) remains pending on a real Windows machine.]

## Suspected Code Paths

- `app/qpn_renderer.py:1111-1113` — `SlideRenderer.__init__` hardcodes `prefer=("powerpoint", "libreoffice", "builtin")`; no user setting exists anywhere.
- `app/qpn_renderer.py:1161-1210` — `SlideRenderer.render()` tries backends in order; every failure only records `_fail()` and `continue`s to the next backend; only total failure logs `SLIDE_RENDER_FAILED`.
- `app/qpn_renderer.py:622-648` — `render_with_powerpoint` REFUSES when a PowerPoint user session is already running (ownership guard, PROMPT-035/041: PowerPoint is a Single Instance COM server; `DispatchEx` would return the user's instance). Refusal raises `PowerPointStageError("ownership")` and the chain falls through.
- `app/qpn_renderer.py:236` — only `powerpoint` is in `FAITHFUL_BACKENDS`; LibreOffice/builtin are documented "honest fallbacks".
- `app/main.py:258` — docstring states "PowerPoint is NOT required here (rendering degrades through the renderer chain)", contradicting the user's requirement.
- Render entry points feeding preview/evidence (all use the same chain): `app/application_service.py:1606,1647` (`learning_preview`), `app/image_extractor.py:97` (`improvement_sheet`), `app/image_extractor.py:264-265` (`after_evidence`), `app/qpn_renderer.py:1302` (`qpn_panel`).
- Diagnostics already exist and must be kept: `POWERPOINT_ISOLATION` log lines, `SLIDE_RENDER` / `SLIDE_RENDER_BACKEND_FAILED` with stable `REASON_*` codes (`app/qpn_renderer.py:123-133`, `last_codes` at `:1220`).

## Root Cause Hypothesis

Confidence: high. The renderer chain is designed to degrade gracefully (first working backend wins), and the PowerPoint ownership guard deliberately refuses whenever the user already has PowerPoint open — the common case on the user's machine (log shows `~$` lock files). Nothing in the chain consults the user's PowerPoint-only requirement, so the refusal degrades silently to the built-in renderer instead of failing closed with the reason and remedy. The 16/10 log in the same file shows `SLIDE_RENDER ... selected_backend=powerpoint faithful=true`, proving COM works on that machine when PowerPoint is not already running — the guard, not PowerPoint availability, is the trigger.

## Proposed Remediation

**Preferred**: fail closed at the render entry points used by preview/evidence (`learning_preview`, `after_evidence`, `qpn_panel`, `improvement_sheet`): when the PowerPoint backend cannot be used, STOP with a clear, specific error carrying the stable reason code and the remedy (e.g. `powerpoint-user-session-running` → "close all PowerPoint windows and retry"; ProgID/pywin32 missing → install/check guidance) instead of trying LibreOffice/builtin. Keep the ownership guard itself (never attach to or kill the user's PowerPoint) and keep the diagnostic logging and reason codes. Do NOT delete the LibreOffice/builtin implementations; gate them behind the fail-closed decision at those entry points. If some non-render path legitimately needs a fallback, document exactly which path and why.

**Alternatives**:
- Add a user-facing renderer setting (PowerPoint-only vs. allow fallback). Rejected for this fix: the user requirement is unconditional PowerPoint-only, and a setting adds UI + persistence surface for a decision already made.
- Surface the refusal reason in the banner only, still rendering with builtin. Rejected: the user explicitly rejects builtin output.

**Files likely to change**:
- `app/qpn_renderer.py` (fail-closed gating in `SlideRenderer.render` / a PowerPoint-only entry path)
- `app/main.py` (docstring correction at :258)
- `app/application_service.py`, `app/image_extractor.py`, `app/batch_processor.py` (only if call sites must handle/present the new error — keep minimal)
- `tests/test_prompt027_rendering.py`, `tests/test_prompt027r_renderer.py`, `tests/test_prompt028_portable.py`, `tests/test_prompt035_powerpoint_isolation.py`, `tests/test_prompt041_isolated_powerpoint.py` (new regression tests; NO weakening of existing tests)

**Tests to add or update**:
- PowerPoint unavailable (ownership refusal simulated) → render raises/reports instead of returning builtin output; error carries the stable code + remedy.
- Refusal reason surfaces end-to-end (renderer `last_codes` / diagnostic) for the user-session-running case.
- Existing fallback tests must keep passing only where they test the (retained) fallback implementations directly, not the gated entry points — any test that asserts silent degradation through the gated entry points must be re-examined, never silently weakened.

## Risks & Considerations

- The ownership guard exists to protect the user's running PowerPoint (PROMPT-035/041). The fix must not weaken it.
- Failing closed changes behavior for machines without PowerPoint: preview/evidence will error instead of rendering degraded output. That is the user's explicit requirement; the error message must say exactly why and what to do.
- The LibreOffice path is used on machines without PowerPoint in other contexts; gating must be scoped to the render entry points above, not a global removal.
- Windows/COM acceptance cannot be verified on Linux; the cloud Windows build + a real PowerPoint run must verify: (1) PowerPoint closed → PowerPoint backend used; (2) PowerPoint open → clear ownership error, no builtin output.

## Open Questions

- [NEEDS CLARIFICATION: none blocking the fix. The exact remedy wording for each reason code should reuse the existing stable codes; UI presentation of the error (banner vs dialog) follows the existing error-display path used by the call sites.]
