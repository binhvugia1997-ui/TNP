# Bug Fix: the picked Excel files are persisted at pick time and reused without the picker

- **Slug**: `prompt040-excel-selection-persistence`
- **Fixed**: 2026-10-16 (UTC)
- **Assessment**: `./assessment.md` (unchanged)
- **Status**: applied
- **Branch**: `arena/b092bbeb-tnp` (Arena session branch, starts at `034d339de7f1a9c3a4176d3ab5d49bdc7a9d95f0` = `integration/prompts-034-035-036-037`)
- **Spec Kit execution**: Executed from checked-in Spec Kit skill instructions (`speckit-bug-assess` → `speckit-bug-fix`); native Arena skill invocation was not available. Slug came from conversation context.
- **Extension hooks**: `.specify/extensions.yml` has `hooks: {}` — no `before_bug_fix` / `after_bug_fix` hooks are registered, so none executed.

## Summary

A successful native Excel pick is now written into the existing user config in the same action, through the existing
`GuiController.save_settings()` → `AppConfig` path (`last_template` / `last_output_file`; no new field, no migration,
no second config system). The remembered files are published on every dashboard response and on demand through a new
`get_excel_selection` bridge method with an honest state
(`unconfigured | checking | ok | missing | inaccessible | invalid`), and the UI shows the current default with
`Đổi file…` / `Xóa lựa chọn`. Nothing re-opens the picker on its own and a stored-but-broken path is reported, never
silently replaced. Excel lock/retry/atomic-commit/workbook-validation and output transaction semantics are untouched.

## Changes

| File | Change | Notes |
|------|--------|-------|
| `app/bridge.py` | modified | `choose_template_file` / `choose_output_file` delegate to a new `_choose_excel(kind)` (`:286`): the same native dialog runs first, then `service.persist_excel_selection(kind, selected_path)`. Response shapes are byte-identical (`{cancelled, path, kind}` / `{cancelled, path}`); a cancelled dialog persists nothing. New `get_excel_selection` (`:125`) and `clear_excel_selection(kind)` (`:129`). No new method accepts a path from JavaScript, so the fixed-capability bridge contract is intact. |
| `app/application_service.py` | modified | `EXCEL_KINDS`/`EXCEL_STATES` + Vietnamese messages (`:50`); `excel_selection()` (`:220`), `persist_excel_selection(kind, path)` (`:225`), `clear_excel_selection(kind)` (`:250`); persistence helper `_persist_excel_settings_unlocked` (`:271`) re-reads `AppConfig.load(...)` to report `persisted` honestly and refuses to write while a run is in progress; `_excel_kind_dto` (`:305`) builds `{path, name, folder, state, message, persisted}`; `_template_excel_state` (`:320`) does `stat`-only checks plus one cached deep check, `_output_excel_state` (`:339`) checks the write contract only (the output may not exist yet); `_start_excel_deep_check`/`_run_excel_deep_check` (`:357`/`:366`) run `excel_writer.validate_template` **once per selection** in one daemon thread, cached by `(path, st_mtime_ns, st_size)`; `dashboard_state()` now also returns `"excel"` (`:611`). |
| `app/gui_controller.py` | modified | `config_storage_path()` (`:2167`) exposes the config file the controller persists to, so the service can verify the write instead of trusting memory. No behaviour change. |
| `frontend/src/types.ts` | modified | `ExcelKind`, `ExcelState`, `ExcelKindState`, `ExcelSelectionState`, `EXCEL_STATE_LABEL` (view models for the DTO above). |
| `frontend/src/state/store.tsx` | modified | `normalizeExcelSelection()` tolerates an absent/partial `excel` block (older or mocked backend ⇒ “Chưa chọn”, never a crash); `acceptDashboard` applies the block on every poll (`:616`); `chooseTemplateFile` / `chooseOutputFile` re-read the state after a pick so the panel shows what Python actually persisted; new `clearExcel(kind)` (`:1013`) and `excel` in the store value (`:1438`). Python/config remains the only authority — no `localStorage`. |
| `frontend/src/components/report/SourcePanel.tsx` | modified | New `ExcelDefaultField` row (`:11`): label + editable path (unchanged manual entry) + `Đổi file…` / `Chọn…` + `Xóa lựa chọn` + “Excel mặc định:” line with file name, full folder, state chip and the Vietnamese explanation; also warns when the selection could not be written to `config.json`. |
| `frontend/package.json` | modified | New `test:prompt040` script; the new test file added to `test:frontend`. |
| `tests/test_prompt040_excel_persistence.py` | added | 21 regression/contract tests (see below). |
| `frontend/tests/prompt040-excel-default.test.mjs` | added | 4 real-browser tests (real `App` + `StoreProvider`, mocked bridge). |
| `frontend/tests/layout/excel-default-harness.tsx`, `.html` | added | Harness that boots the real application in Chromium for the 4 tests above. |
| `.specify/bugs/prompt040-excel-selection-persistence/fix.md`, `test.md` | added | This report and the verification report. |

Nothing else was touched: no version/build bump, no change to `ReportExtractor.spec`, `backup/…STABLE/`, `app/updater.py`,
`app/excel_writer.py`, `app/config.py` (schema unchanged) or the extraction pipeline.

## Diff Highlights

```python
# app/bridge.py — the pick is persisted where it is validated; shapes unchanged
def _choose_excel(self, kind: str) -> dict:
    if kind == "template":
        selected = self._choose_file(XLSX_FILTER, ".xlsx", "template")
    else:
        selected = self._choose_save_file("Kiem_chung_Ket_qua.xlsx", XLSX_FILTER, ".xlsx")
    if selected.get("cancelled"):
        return selected
    self.service.persist_excel_selection(kind, selected["path"])
    return selected
```

```python
# app/application_service.py — existing config mechanism, verified by re-reading the file
def _persist_excel_settings_unlocked(self, kind: str) -> bool:
    if self.controller.is_running():
        self._excel_persisted_key = None
        return False
    self.controller.save_settings()
    self._excel_persisted_key = None                     # re-read the file, never trust memory
    return bool(self._excel_persisted_unlocked().get(kind))
```

```tsx
// frontend/src/components/report/SourcePanel.tsx — what the operator sees
<Button disabled={disabled} aria-label={`Chọn ${label}`} onClick={onPick}>
  {state.path ? 'Đổi file…' : 'Chọn…'}
</Button>
<Button variant="ghost" disabled={disabled || !state.path} aria-label={`Xóa lựa chọn ${label}`} onClick={onClear}>
  <Trash2 className="h-3.5 w-3.5" /> Xóa lựa chọn
</Button>
```

## Tests Added or Updated

`tests/test_prompt040_excel_persistence.py` (production objects only: `BridgeService` → `ApplicationService` →
`GuiController` → `AppConfig`, scripted native-dialog stand-in):

- `test_excel_pick_is_written_into_the_existing_config_file` — pick ⇒ `config.json` holds both paths (the reproduced bug).
- `test_pick_response_shape_is_unchanged` — response shapes and `kind` field byte-compatible.
- `test_restart_restores_both_excel_paths_without_any_dialog`, `test_restarted_default_is_exposed_in_the_dashboard` — restart ⇒ restored, `excel` states, zero dialog calls.
- `test_processing_uses_the_stored_excel_files_without_a_picker` — `BatchOptions.template`/`output_file` come from the stored default; no dialog.
- `test_second_pick_replaces_the_persisted_default` — “Đổi file” ⇒ new path persisted/displayed.
- `test_deleted_stored_template_is_reported_and_never_silently_replaced`, `test_corrupt_template_is_reported_as_invalid`, `test_unconfigured_state_before_the_first_pick` — state matrix; stored path never substituted; `start_processing` keeps raising the existing `VALIDATION_FAILED`.
- `test_unicode_space_and_relative_pick_paths_round_trip` — Vietnamese diacritics + spaces + `..` segments.
- `test_legacy_config_without_excel_selection_loads_normally`, `test_corrupt_config_file_keeps_the_existing_load_contract`, `test_unknown_legacy_keys_survive_a_pick` — old/corrupt config compatibility, no destructive migration.
- `test_updater_preserves_the_saved_excel_selection` — update + rollback keep `config/config.json` (PROMPT-039 contract) and `learning_data`.
- `test_persisting_a_pick_does_not_open_or_modify_the_workbooks` — bytes/mtime unchanged, no `~$` lock, no `.saving.xlsx`, no `.tmp`.
- `test_pick_reports_when_the_selection_could_not_be_persisted` — read-only config dir ⇒ pick still works, `persisted: false` (honest, not silent).
- `test_folder_pptx_and_update_pickers_do_not_write_an_excel_default`, `test_cancelled_excel_pick_is_not_persisted_and_changes_nothing` — folder/PPTX/update pickers unchanged; cancel writes nothing.
- `test_clear_excel_selection_is_persisted_and_idempotent`, `test_excel_states_cover_every_documented_state`, `test_excel_selection_keeps_learning_storage_untouched`.

`frontend/tests/prompt040-excel-default.test.mjs` (real Chromium from the checked-in `@sparticuz/chromium`; SKIPs with a
printed reason when no browser is available — never false-pass):

- `a restored Excel default is shown and no picker runs on startup` — name, folder, “Dùng được” chip, `Đổi file…` label, zero picker calls across several dashboard polls.
- `a stored file that disappeared is reported and not replaced` — “Không tìm thấy” + message, no picker call.
- `“Đổi file…” persists a new pick and displays it again` — exactly one native dialog per click, new path re-read from Python.
- `“Xóa lựa chọn” clears only that Excel default` — backend called once, row back to “Chưa chọn”, the other row untouched.

## Local Verification

Linux development evidence only (no Windows dialog/WebView2/COM/Portable acceptance claimed).

- Pre-fix baseline (fresh worktree at `034d339` + the new test file): `pytest -q tests/test_prompt040_excel_persistence.py` → **18 failed, 3 passed**.
- Post-fix: same command → **21 passed**.
- Reproduction script (`/home/user/TNP_scratch/repro_prompt040.py`, not committed): `config after pick` now holds both paths; `restart (no scan)` restores both; the raw script previously printed `last_template='' last_output_file=''`.
- Targeted Python suites: `pytest -q tests/test_prompt040_excel_persistence.py tests/test_application_service.py tests/test_gui_controller.py tests/test_updater.py tests/test_image_learning.py tests/test_prompt030b_integration.py tests/test_excel_lock_retry.py tests/test_excel_batch.py tests/test_prompt036r_bridge_hang.py tests/test_prompt038_desktop_freshness.py` → **212 passed**.
- Full available suite: `.venv/bin/python -m pytest -q` → **1,332 passed, 3 skipped** (2 skipped = shallow Arena checkout, “git history unavailable”; 1 = Windows-only CLR loading test).
- Frontend: `npm run typecheck` clean; `npm run test:prompt040` → **4 passed**; `npm run test:frontend` → **83 passed, 0 failed**; `npm run build` → `dist/` rebuilt; `desktop.frontend_dist_freshness()` → `{"stale": false, "reason": "fresh"}`.
- Pre-fix frontend baseline (same worktree at `034d339`): the new test file → **0 passed, 4 failed**.
- `pyflakes app/application_service.py app/bridge.py app/gui_controller.py tests/test_prompt040_excel_persistence.py` clean; `git diff --check` clean.
- Known pre-existing flake (not caused by this fix): `PROMPT-034R · >6 rows` fails intermittently (workspace-height equality) — reproduced **1 of 4 runs on the untouched base commit**.

## Deviations from Assessment

- **Working branch**: the assessment (and the user instruction) named a new task branch derived from
  `integration/prompts-034-035-036-037`. This session is bound to the Arena branch `arena/b092bbeb-tnp`, which starts at
  exactly `034d339de7f1a9c3a4176d3ab5d49bdc7a9d95f0` (the same commit), and the session rules plus `AGENTS.md` §6 forbid switching to
  or creating another branch. All work was therefore committed on `arena/b092bbeb-tnp`; nothing else changed about the base.
- The rest of the remediation was applied as proposed (persist at pick time through the existing config, cached one-shot
  deep check in a daemon thread, dashboard + `get_excel_selection` exposition, UI in `SourcePanel`). No new findings
  invalidated the assessment; `assessment.md` was not edited.
- Two cosmetic test-file fixes were needed after the first post-fix run (an unused import and an unused local variable
  in the new test module) plus one scripted-dialog ordering fix in the folder/PPTX test; the production code was not
  adjusted to make a test pass.

## Follow-ups

- Windows acceptance: run the retest steps in `test.md` on a Windows host with PowerPoint COM and the Portable build.
- Optional later cleanup: the `checking` state is visible for a moment on very large check workbooks (the deep check is
  deliberately off the request thread); if that ever proves distracting, the deep check could also be surfaced in the
  Cài đặt → Chẩn đoán card.
