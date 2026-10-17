# Bug Fix: the PROMPT-040 persistence-failure test is now cross-platform

- **Slug**: `prompt040-windows-test-portability`
- **Fixed**: 2026-10-17 (UTC)
- **Assessment**: `./assessment.md`
- **Status**: applied (test-only; no production change)
- **Branch**: `arena/b092bbeb-tnp` (same PROMPT-040 branch / PR #16), base `integration/prompts-034-035-036-037`
- **Starting SHA**: `8c00d9bf76500ed52a5b1bf518f29be7c906cf03`
- **Spec Kit execution**: Executed from checked-in Spec Kit skill instructions (`speckit-bug-fix`); native Arena skill invocation was not available. Slug came from conversation context.
- **Extension hooks**: `.specify/extensions.yml` has `hooks: {}` — no `before_bug_fix` / `after_bug_fix` hooks are registered, so none executed.

## Summary

Windows acceptance of `tests/test_prompt040_excel_persistence.py` failed one test before any assertion ran:
`AttributeError: module 'os' has no attribute 'geteuid'`. The test manufactured the “config cannot be written”
condition with Unix account semantics (`os.geteuid`) and POSIX permission bits (`chmod 0o500`); neither works on
Windows, and `chmod` does not make a directory unwritable there. The failure is now injected deterministically at
the real disk-write boundary — a monkeypatched `AppConfig.save` raising `PermissionError` (an `OSError`) — so the
same test proves the same contract on Windows, macOS and Linux, root or not, with no dependency on real protected
directories. **Production code was not changed**; the assessed root cause was purely in the test.

## Changes

| File | Change | Notes |
|------|--------|-------|
| `tests/test_prompt040_excel_persistence.py` | modified | Rewrote `test_pick_reports_when_the_selection_could_not_be_persisted`: removed `os.geteuid()` + `chmod 0o500`; injects an `OSError` from `AppConfig.save` for the duration of the pick and restores it afterwards. Assertions strengthened (see below). Removed the now-unused `import pytest`; `os` is still used for `os.getpid()` in the updater-preservation test. |
| `.specify/bugs/prompt040-windows-test-portability/assessment.md` | added | Reproduction (Windows attribute surface simulated on Linux), trace of the real error contract, root cause, plan. |
| `.specify/bugs/prompt040-windows-test-portability/fix.md` | added | This report. |
| `.specify/bugs/prompt040-windows-test-portability/test.md` | added | Verification report. |

No production file changed: `git diff --stat` for this fix is `tests/test_prompt040_excel_persistence.py | 32 +, 15 -`.
`app/application_service.py`, `app/bridge.py`, `app/config.py`, `app/gui_controller.py` and the frontend are byte-identical
to PR #16's previous head.

## Diff Highlights

```python
def test_pick_reports_when_the_selection_could_not_be_persisted(tmp_path, monkeypatch):
    ...
    real_save = AppConfig.save

    def unwritable_save(_self, _path=None):
        raise PermissionError(13, "config.json không ghi được (mô phỏng thư mục chỉ đọc)")

    monkeypatch.setattr(AppConfig, "save", unwritable_save)

    picked = bridge.choose_template_file()
    assert picked["ok"] is True and picked["data"]["path"] == str(template.resolve())
    assert not config_path.exists()                                # the write really did not happen
    assert not config_path.with_suffix(".json.tmp").exists()       # …not even a leftover temp file
    assert any("Không lưu được config.json" in line for line in service.controller.log_lines)
    state = bridge.get_excel_selection()["data"]
    assert state["template"]["path"] == str(template)              # used for this session
    assert state["template"]["state"] in EXCEL_STATES
    assert state["template"]["persisted"] is False                 # and reported honestly

    # The failure is not sticky: once writes succeed again the same selection is persisted for real.
    monkeypatch.setattr(AppConfig, "save", real_save)
    repaired = service.persist_excel_selection("template", str(template))
    assert repaired["template"]["persisted"] is True
    assert AppConfig.load(config_path).last_template == str(template)
```

Why this boundary is faithful rather than a mock-only shortcut: the real `GuiController.save_settings()` body runs,
sets `cfg.last_template`, calls the (failing) `AppConfig.save`, and its **production** `except OSError` handler writes
`Không lưu được config.json: …` — which the test asserts. `ApplicationService` then re-reads `AppConfig.load(...)` and
reports `persisted: false`. Nothing in the production path is bypassed.

## Tests Added or Updated

- `tests/test_prompt040_excel_persistence.py::test_pick_reports_when_the_selection_could_not_be_persisted` —
  proves the contract under a **genuine** write failure, now identically on every OS:
  - the pick response stays `ok: true` with the resolved path;
  - the config file (and its `.json.tmp`) really is absent;
  - the production “cannot save config.json” handler ran;
  - the DTO keeps the path/state but reports `persisted: false`;
  - the failure is recoverable — the same selection persists once saving works again.
- Every other test in the file is unchanged.

## Local Verification

Linux development evidence only — real Windows execution is not claimed.

- Pre-fix reproduction (Windows attribute surface, `del os.geteuid`): exact test → **1 failed**, `AttributeError: module 'os' has no attribute 'geteuid'` at `tests/test_prompt040_excel_persistence.py:427`.
- Post-fix, exact test: Linux → **1 passed**; with `os.geteuid` deleted → **1 passed**.
- Post-fix, full file: Linux → **21 passed**; with `os.geteuid` deleted → **21 passed**.
- Mutation checks (prove the test is not vacuous):
  - DTO forced to report `persisted: true` → test **fails** (`assert True is False`);
  - `save_settings` with its `except OSError` removed → test **fails** with the propagated `PermissionError`.
- Bridge/controller/config regressions: `pytest -q tests/test_application_service.py tests/test_gui_controller.py tests/test_prompt036r_bridge_hang.py tests/test_core.py tests/test_gui_lifecycle.py tests/test_auto_update_check.py` → **97 passed**.
- Full available suite: `.venv/bin/python -m pytest -q` → **1,332 passed, 3 skipped** (2 × “git history unavailable” in the shallow Arena checkout, 1 × Windows-only CLR test).
- `pyflakes tests/test_prompt040_excel_persistence.py` clean; `git diff --check` clean.
- `npm run build` (includes `tsc -b`) OK; `desktop.frontend_dist_freshness()` → `{'stale': False, 'reason': 'fresh'}` (frontend sources unchanged).

## Deviations from Assessment

None. The preferred remediation was applied exactly: failure injected at the disk-write boundary, no uid/chmod/
administrator/protected-directory dependence, assertions strengthened instead of weakened, production untouched.

## Follow-ups

- Re-run the Windows targeted acceptance on this branch head; exact commands are in `test.md`.
- No production follow-up: `AppConfig.save` / `GuiController.save_settings` behaviour under a real read-only or
  locked config folder is unchanged and already covered by their own handlers.
