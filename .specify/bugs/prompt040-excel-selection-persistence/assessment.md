# Bug Assessment: Repeated Excel picker — a successful “File kiểm chứng / File kết quả” selection is not persisted

- **Slug**: prompt040-excel-selection-persistence
- **Created**: 2026-10-16 (UTC)
- **Source**: pasted text (PROMPT-040 requirement in the Arena session)
- **Verdict**: valid (reproduced on Linux against the production bridge/service/controller/config objects; development evidence only)
- **Severity**: medium (repeated manual work and lost selections between restarts; no data loss, no Excel transaction change, no silent file substitution)
- **Starting SHA**: `034d339de7f1a9c3a4176d3ab5d49bdc7a9d95f0` = `refs/heads/integration/prompts-034-035-036-037` (verified against `git ls-remote origin`); PROMPT-039 and PROMPT-030B already merged there
- **Working branch**: `arena/b092bbeb-tnp` (the Arena session branch that starts at that exact commit; the session rules forbid switching to another branch)
- **Spec Kit execution**: Executed from checked-in Spec Kit skill instructions; native Arena skill invocation was not available.
- **Extension hooks**: `.specify/extensions.yml` contains `hooks: {}` — no `before_bug_assess` / `after_bug_assess` hooks are registered, so none executed.

## Report (summarized)

The user must pick the Excel files (“file Excel kiểm/đầu ra” = `File kiểm chứng Excel` and `File kết quả`)
again and again. Requirement: pick the Excel file **once**, have the app remember it and reuse it on later
runs, without opening the picker again; Python/config stays the source of truth (no frontend
`localStorage` authority); the picker may only be needed when nothing was configured yet, or when the stored
file was deleted/moved/becomes unreachable/invalid. The UI must show the current default Excel with
name/path, a “Đổi file” button, a clear status when the file is missing/invalid, and optionally
“Xóa lựa chọn”. Excel lock/retry/atomic-commit/workbook-validation semantics must not change, the
PROMPT-039 updater preservation contract must stay intact, and folder/PPTX pickers must not regress.

## Symptom

Expected: `choose_template_file()` / `choose_output_file()` (native picker) persist the chosen path into the
current user config, so a later app start restores it and processing reuses it without any picker.

Actual: the picker result is returned to the browser only. Nothing writes the path to `config.json`, so a
fresh process starts with `paths.template` / `paths.output` empty and the user has to open the picker again.
The path is persisted only **incidentally**, when some unrelated later action happens to call
`GuiController.save_settings()` (a completed scan, “Lưu cấu hình”, “Kiểm tra kết nối”, “Làm mới model”,
“Kiểm tra cập nhật”, applying a discovered Ollama server, toggling auto-update, or starting a processing run).
A selection that is not followed by one of those actions is lost on the next start.

## Reproduction

Raw reproduction script (not committed; run with the repository venv on the base commit):

```bash
python3 -m venv .venv && .venv/bin/python -m pip install -r requirements-dev.txt
PYTEST_CURRENT_TEST=repro .venv/bin/python /home/user/TNP_scratch/repro_prompt040.py
```

It drives the real production objects — `BridgeService` → `ApplicationService` → `GuiController` →
`AppConfig` — with one fake native dialog window (`create_file_dialog`) whose choices are a template and an
output path, exactly like `tests/test_application_service.py` does:

```
runtime root      : /tmp/prompt040-repro-…/runtime
config file       : /tmp/prompt040-repro-…/runtime/config/config.json
picked template   : {'ok': True, 'data': {'cancelled': False, 'path': '…/picked/Bieu mau kiem chung 2026.xlsx', 'kind': 'template'}}
picked output     : {'ok': True, 'data': {'cancelled': False, 'path': '…/picked out/Kiem chung ket qua 2026.xlsx'}}
config after pick : last_template='' last_output_file=''
restart (no scan) : template='' output=''
config after scan : last_template='…/picked/Bieu mau kiem chung 2026.xlsx' last_output_file='…/picked out/Kiem chung ket qua 2026.xlsx'
restart (after scan): template='…/picked/Bieu mau kiem chung 2026.xlsx' output='…/picked out/Kiem chung ket qua 2026.xlsx'
```

Reproduced: **yes** (bug). Both picks returned `ok: true` with a fully resolved path, yet `config.json` still
held `last_template: ""` / `last_output_file: ""`, and a new `ApplicationService` (a restart) started with
`template=''`, `output=''` → the user must open the picker again. The same paths survive a restart only after
a scan is run, which confirms the missing write is exactly “at pick time”.

Manual reproduction on Windows (one command sequence, no Linux-only step):

1. Start the app with an empty `config/config.json` (delete the two Excel fields or use a fresh Portable copy).
2. Pick `File kiểm chứng Excel`, then `File kết quả` with the native pickers.
3. Close the app immediately (do not press `Quét` and do not open Cài đặt → `Lưu cấu hình`).
4. Inspect `<Portable>\config\config.json`: `last_template` and `last_output_file` are still `""`.
5. Start the app again: both Excel fields are empty and the fields’ `Chọn…` buttons are the only way back.

## Traced flow (frontend → bridge → controller/service → config → picker)

1. **Frontend (picker button)** — `frontend/src/components/report/SourcePanel.tsx:53` renders
   `Chọn…` → `s.chooseOutputFile`; `:43` renders `Chọn…` → `s.chooseTemplateFile`.
2. **Frontend (store)** — `frontend/src/state/store.tsx:921-936`:
   `chooseTemplateFile` calls the bridge `choose_template_file`, then **only** `setTemplate(result.path)`;
   `chooseOutputFile` calls `choose_output_file`, then **only** `setOutput(result.path)`. No
   `save_configuration` call, no other persistence — the React state is the only result.
3. **Bridge** — `app/bridge.py:118-122`:
   `choose_template_file` → `self._choose_file(XLSX_FILTER, ".xlsx", "template")`;
   `choose_output_file` → `self._choose_save_file("Kiem_chung_Ket_qua.xlsx", XLSX_FILTER, ".xlsx")`.
   Both validate suffix/existence (`app/bridge.py:269-300`) and return `{"cancelled": False, "path": …}`.
   Neither touches `ApplicationService.controller` or `AppConfig`.
4. **Service/controller** — the only writer of the Excel fields is
   `GuiController.save_settings()` (`app/gui_controller.py:2167-2186`), which copies
   `self.template`/`self.output` into `cfg.last_template`/`cfg.last_output_file` and calls `cfg.save(...)`.
   Call sites: `app/application_service.py:195` (`save_configuration` — reachable only from
   `Cài đặt → Lưu cấu hình` / `Kiểm tra kết nối` / `Làm mới model` / `Kiểm tra cập nhật`),
   `app/application_service.py:289` (after a **successful scan**, `save_reports` → `scan_reports`),
   `app/application_service.py:831` (after applying a discovered Ollama server),
   `app/gui_controller.py:1443` (`save_ollama_settings`), `:1553` (start of a processing run),
   `:1932` (`set_auto_update_check`); plus the classic Tk GUI `app/gui.py` (1165/1172/1182/1667/1671/1718/2342).
   ⇒ **No picker path calls any of them.** `scan_reports` (`:217+`) and `start_processing` (`:549+`) apply the
   DTO through `_apply_configuration_for_scan` (`:466-488`), which mutates controller fields only.
5. **Config** — `AppConfig` (`app/config.py:119-160`) already has the persistence mechanism and the fields:
   `last_template`, `last_output_file`, written atomically to `<portable>/config/config.json`
   (`runtime_paths.config_file()`, preserved by the updater through `PRESERVED_ENTRIES`, `app/updater.py:69`).
   => no new config system is needed; the existing one is simply never invoked by the picker.
6. **Restore path (already correct)** — `GuiController.__init__` (`app/gui_controller.py:389-391`) loads
   `cfg.last_template` / `cfg.last_output_file`; `ApplicationService.current_config`
   (`app/application_service.py:140-156`) publishes them; the frontend collects them in the one startup
   handshake (`frontend/src/lib/startup.ts:32`) and applies them in `applyConfig`
   (`frontend/src/state/store.tsx:525-528`, called from `:736`). So whatever is in `config.json` *is*
   restored without a picker — the defect is purely the missing write at pick time.
7. **Picker re-entry** — nothing in the backend ever opens a dialog on its own; the user is forced back to the
   picker because the restored value is empty. When the stored value exists but the file is gone, processing
   fails validation with a message (`GuiController.validate`, `app/gui_controller.py:1177-1211`,
   “File Kiểm chứng (.xlsx) không tồn tại.”) and there is **no UI status** that explains it before the run.

## Root Cause Hypothesis

**Confidence: high.** The pick action is a read-only native-dialog capability: `choose_template_file` /
`choose_output_file` return the validated path to JavaScript and stop there, while Python's own Excel state
(`GuiController.template` / `.output`) and `config.json` are written only by unrelated later actions. The
frontend keeps the pick in React state only and never pushes it back through `save_configuration`
(`store.tsx:921-936`), so a restart — or any exit before the incidental `save_settings()` — re-empties the
Excel fields and forces another pick. There is no second, missing storage layer: the existing
`AppConfig.last_template` / `last_output_file` fields, the startup handshake and `applyConfig` already
implement restore; only the write is missing.

## Proposed Remediation

**Preferred**: persist the pick where the pick is validated, and surface the stored selection.

1. `app/bridge.py` — after a successful Excel dialog selection, register the path with the service
   (`ApplicationService.persist_excel_selection(kind, path)`), keeping the existing response shapes byte-for-byte
   (`{"cancelled": False, "path": …}` / `…, "kind": "template"`). The path still comes only from the native
   dialog, so the “no arbitrary path API” bridge contract is untouched.
2. `app/application_service.py` — new `persist_excel_selection(kind, path)`, `excel_selection()` and
   `clear_excel_selection(kind)`:
   - set the controller field through the existing `set_template` / `set_output` (scan-invalidation semantics kept),
   - persist with the existing `GuiController.save_settings()` (existing `AppConfig`/`config.json`,
     no schema change, no migration),
   - verify the persisted value by re-reading `AppConfig.load(...)` and report `persisted` honestly,
   - publish a per-kind DTO (`path`, `name`, `folder`, `state`, `message`) with states
     `unconfigured | checking | ok | missing | inaccessible | invalid`, embedded in `dashboard_state()` and
     exposed through a dedicated `get_excel_selection` bridge method. The template’s deep Excel contract
     (`excel_writer.validate_template`, the same call `validate()` uses) is evaluated once per selection in a
     daemon thread and cached by `(path, mtime_ns, size)`, so the UI can show “Hợp lệ”/“Không hợp lệ” without
     blocking a dashboard poll and without a second validation implementation.
3. `frontend/` — show the stored default Excel in `SourcePanel` (name + full path + status chip + Vietnamese
   message), rename the Excel picker buttons to `Đổi file…`, add `Xóa lựa chọn`, refresh the state after a pick
   / after clearing, and apply the `excel` block from the running dashboard poll. Nothing re-opens the picker
   automatically; a stored-but-broken path is reported, never silently replaced.

**Alternatives**:
- *Frontend-only*: call `save_configuration(configDto())` right after a pick. Rejected as the primary fix: it
  keeps two writes (pick + save) that can race or be forgotten, and it makes React state the effective
  authority for a persisted value. It is the natural fallback only if the bridge must stay read-only.
- *New config field(s)* (e.g. `default_excel_template`): rejected — would duplicate the existing persisted
  selection and force a migration for no behavioural gain.

**Files likely to change**: `app/bridge.py`, `app/application_service.py`, `app/config.py` (only if a helper is
needed; no schema change expected), `frontend/src/state/store.tsx`, `frontend/src/types.ts`,
`frontend/src/components/report/SourcePanel.tsx`, `frontend/package.json` (register the new frontend test),
plus tests and this bug directory.

**Tests to add or update** (new `tests/test_prompt040_excel_persistence.py`, plus one frontend test file):
1. pick → `config.json` holds both Excel paths (regression for the demonstrated bug);
2. restart (new service over the same config) → both paths restored, no `create_file_dialog` call;
3. stored valid Excel → `scan_reports` + `start_processing` use it and never open a dialog;
4. “Đổi file” (second pick) → new path persisted and displayed;
5. stored file deleted → state `missing`, stored path not silently replaced, processing raises the existing
   `VALIDATION_FAILED` and no dialog opens;
6. Unicode + spaces (+ `..`) paths round-trip through `config.json`;
7. legacy `config.json` without the new state loads normally (`excel_selection` → unconfigured);
8. corrupt/unreadable `config.json` keeps the existing fallback contract (defaults, no crash, no data loss);
9. updater preservation: the saved Excel selection survives an update/rollback (existing preservation rules);
10. Excel safety: pick/persist neither opens nor modifies the template workbook (no lock file, unchanged
    bytes/mtime) — lock/retry/atomic-commit behaviour is untouched;
11. folder picker / PPTX picker unchanged (existing parametrized picker test kept; folder selection is not
    persisted as an Excel default);
12. `learning_data` persistence (PROMPT-030B) untouched;
13. PROMPT-039 updater tests unchanged;
14. frontend: default-Excel display, `Đổi file`, `Xóa lựa chọn`, missing-file status, and “no picker call on
    startup” in the real-browser harness.

## Risks & Considerations

- **Restore change is additive**: existing fields/schema are reused; old `config.json` files must keep loading
  (`AppConfig.load` ignores unknown keys into `extra`, and missing keys keep their dataclass defaults).
- **Bridge surface**: only fixed backend-owned actions are added (`get_excel_selection`, `clear_excel_selection`);
  no method accepts an arbitrary path from JavaScript, so validation stays on the Python side.
- **Validation cost**: `validate_template` opens the workbook; it must never run per dashboard poll. Hence the
  once-per-selection cached background check; the cheap structural state stays `stat`-only.
- **Portable read-only locations**: if `config/config.json` cannot be written, the existing
  `save_settings()` behaviour (log a line, no crash) is kept, and the new DTO reports `persisted: false`
  instead of pretending success.
- **Busy state**: a pick cannot be persisted while a run is in progress (the UI is locked anyway); the DTO
  reports it rather than failing silently, and the next scan/processing start persists it.
- **Not changed**: lock detection, retry, atomic commit, workbook validation rules, output transaction
  semantics, updater preservation, version/build, `backup/…STABLE/`.

## Open Questions

- [NEEDS CLARIFICATION: none blocking. UX wording (“Đổi file”, “Xóa lựa chọn”, status labels) follows the
  prompt; Windows/PowerPoint COM acceptance and Portable acceptance remain unperformed on this Linux host and
  must be re-run on Windows.]
