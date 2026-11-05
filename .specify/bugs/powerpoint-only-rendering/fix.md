# Fix: PowerPoint-only rendering, fail-closed (powerpoint-only-rendering)

- **Slug**: powerpoint-only-rendering
- **Assessment**: `.specify/bugs/powerpoint-only-rendering/assessment.md`
- **Status**: fixed (Linux dev evidence only; Windows/PowerPoint-COM acceptance still required)
- **Mode B disclosure**: Executed from checked-in Spec Kit skill instructions; native Arena skill invocation was not available.

## Root cause (confirmed by code trace + user log/screenshot)

`app/qpn_renderer.py::SlideRenderer` tried the hardcoded chain `("powerpoint", "libreoffice", "builtin")`, first working backend wins. `render_with_powerpoint` REFUSES when a PowerPoint user session is already running (ownership guard, PROMPT-035/041: PowerPoint is Single Instance COM, `DispatchEx` returns the user's instance), and the chain then fell through silently to the built-in renderer. `app/main.py:258` even documented "PowerPoint is NOT required here (rendering degrades through the renderer chain)", contradicting the user's standing requirement (rendering MUST use PowerPoint; silent fallback is NOT acceptable; there is no user-facing renderer setting). The user's log showed many `~$` PowerPoint lock files at the time the preview fell back to builtin. Stable reason codes, `POWERPOINT_ISOLATION` / `SLIDE_RENDER` logging, and the frontend `fidelityHint(reason)` mapping already existed — the chain just never surfaced them.

## Changes

- `app/qpn_renderer.py`:
  - New `REASON_RENDER_FAILED = "powerpoint-render-failed"`; new `powerpoint_remedy(code)` (Vietnamese remedies, e.g. `powerpoint-user-session-running` → close all PowerPoint windows and retry; `powerpoint-unavailable` → install PowerPoint; `powerpoint-import` → reinstall the full portable / pywin32).
  - New `PowerPointUnavailableError(RuntimeError)` with `.code` and `.remedy`.
  - `SlideRenderer.__init__` gained `powerpoint_only: bool = False`. `render()` uses the chain `["powerpoint"]` when set and raises `PowerPointUnavailableError` (via `_raise_powerpoint_only`, logging `SLIDE_RENDER_POWERPOINT_ONLY_FAILED`) on availability failure (code from the `powerpoint_probe()` stage), on `PowerPointStageError` (stable `e.code`), and on generic render exceptions (`REASON_RENDER_FAILED`). A first `except PowerPointUnavailableError: raise` keeps the fail-closed raise from being re-caught by the generic handler.
  - The ownership guard is untouched. The LibreOffice/builtin implementations are NOT deleted; the default chain (`powerpoint_only=False`) is preserved for direct/test use.
- `app/application_service.py`: `_slide_preview_for` constructs `SlideRenderer(width_px=LEARNING_PREVIEW_WIDTH_PX, powerpoint_only=True)`; the preview-cache signature now includes `renderer.powerpoint_only` (busts old cached builtin previews); a new `except PowerPointUnavailableError` handler returns an empty preview with `"reason": exc.code, "degraded": True`, uncached; `_image_candidate_dto` preserves `exc.code` as the preview reason (the DTO already mapped `preview["reason"]` → `slidePreviewReason`).
- `app/batch_processor.py:216`: the batch renderer is now `SlideRenderer(width_px=opts.render_width, powerpoint_only=True)` (feeds `render_qpn_panel` and `export_after_pictures`).
- `app/image_review.py`: `reapply_labels` now passes `renderer=SlideRenderer(powerpoint_only=True)` to `export_after_pictures` (save-time re-apply path).
- `app/main.py:258`: docstring corrected — PowerPoint IS required for rendering (fail-closed).
- `frontend/src/tabs/LearningTab.tsx`: the empty-preview branch now shows, when `previewReason` is set, "Không render được bằng PowerPoint ({reason})." + `fidelityHint(previewReason)` instead of a bare "Preview không khả dụng."
- `tests/test_powerpoint_only_rendering.py` (new, 7 tests): unavailable → `powerpoint-unavailable` + no PNG + builtin never called; pywin32 import failure → `powerpoint-import`; ownership refusal → `powerpoint-user-session-running` + "Đóng tất cả cửa sổ PowerPoint" remedy + no PNG; generic COM failure → `powerpoint-render-failed`; PowerPoint success still renders; default chain still falls back to builtin (control); remedy covers every stable code and the session-running remedy differs from the unavailable remedy.

## Deliberate exception (decision, report it)

`export_after_pictures` (`app/image_extractor.py:215-330`) keeps its embedded-picture fallback (original PPTX picture blobs, NOT a re-render) when rendering fails, with the render error recorded visibly in `problems` / `REGION_FALLBACK` review reasons. The renderer chain itself is gated; this source-data fallback was judged out of scope to remove (it would change Excel output semantics broadly). If the user wants even that removed, it is a separate decision.

## Test updates for the new contract (intent preserved, each documented)

The fail-closed gate changed what Linux batch/learning tests must do: production paths now need a PowerPoint render, which does not exist on Linux. New shared helper `tests/powerpoint_render_stub.py` stubs the COM probe + export (white slide with every picture shape drawn as its average-colour rectangle at authored geometry — enough for geometry and colour-presence assertions, never a faithful render). Applied at the batch harness / direct-call level in: `test_prompt004`, `test_content_region`, `test_after_evidence` (via `test_prompt004._run`), `test_image_learning` (direct `proc.run()` / `reapply_labels` / controller calls), `test_fixed_master_rules`, `test_master_not_found`, `test_new_master_row`, `test_new_row`, `test_prescan` (incl. a direct `GuiController` run), `test_production_rows`, `test_excel_batch` (incl. the worker-thread test), `test_prompt020_report_scoping`, `test_prompt025_multi_item` (incl. a direct construction), `test_prompt024r_cancel_all`, `test_prompt024r_learning_workspace` (autouse fixture; no test in that file asserts degraded behavior), `test_validation_phase`, `test_gui_controller`, `test_scan_list`. No assertion was weakened: counts, colours, geometry, and status assertions are unchanged; only the render backend behind them is stubbed.

Contract tests pinning the old degraded behavior were updated to the new fail-closed contract (see also the Bug B record for the shared `test_prompt030b_integration.py` updates):

- `tests/test_prompt027_rendering.py` (5): module helper `_fake_powerpoint(monkeypatch)` stubbing the COM call with white PNGs so the real render loop runs; cache-identity, cache-isolation, failed-preview-recovery, DTO backend/faithful, and preview-purity tests updated to the PowerPoint backend.
- `tests/test_prompt027r_renderer.py` (2): failed-preview identity test (failed preview empty + reason, never cached); the degraded-fallback TTL test renamed to `test_powerpoint_only_preview_is_retried_on_every_access_and_recovers` (failed previews never cached, retried every access, reason `powerpoint-slide-export`).
- `tests/test_prompt041w_rot_false_positive.py` (3): ownership-blocked preview empty + reason + uncached + recovers on PowerPoint close; the `09b` test renamed to `test_09b_an_unavailable_powerpoint_preview_never_claims_a_fallback_render`; the `14c` test drives the unblocked environment and asserts a cache line with `cached_backend=powerpoint` + `action=reuse`. The static-safety tests (`09c`, `09d`, `10`, `15`–`18`: no taskkill/terminate/attach/registry writes) are untouched.
- `tests/test_prompt030b_integration.py` (4 tests, shared with Bug B): production-flow backend assertion `builtin`/unfaithful → `powerpoint`/faithful; persistent-cache `availability` param now flips availability True→False (the COM stub starts True) and asserts an empty preview with reason `powerpoint-unavailable`; ownership-refusal test rewritten to the retry-every-access contract (empty preview + `powerpoint-ownership` reason, `calls` 1→2→3, diagnostic log lines preserved).

## Verification (Linux)

- New Bug A tests: 7/7 pass.
- Renderer suites (`test_prompt027_rendering`, `test_prompt027r_renderer`, `test_prompt028_portable`, `test_prompt035_powerpoint_isolation`, `test_prompt041_isolated_powerpoint`, `test_prompt041w_rot_false_positive`): 248 passed, 8 skipped, 1 failed — the failure is the pre-existing environment probe `test_live_import_probe_reports_every_module_in_this_venv` (`ModuleNotFoundError: pythonnet`, Windows-only; same class as the pre-change baseline).
- Full Python suite: 1404 passed, 1 failed (the pythonnet probe above), 13 skipped — vs the pre-change baseline of 1385 passed / 4 failed (3 LibreOffice-selection environment failures now pass; the pythonnet failure remains).
- Pyflakes on changed Python files: clean except pre-existing unused imports.

## Still needs Windows verification (via the cloud workflow build)

1. PowerPoint closed → the PowerPoint backend is used for previews, batch renders, and re-apply.
2. PowerPoint open with files → a clear ownership error with the stable code + remedy, NO builtin output, and the preview shows the reason message.
3. PowerPoint / pywin32 missing → the distinct `powerpoint-unavailable` / `powerpoint-import` codes and remedies.
4. Batch + re-apply evidence renders on Windows produce the same crops the stubbed Linux tests assert.
