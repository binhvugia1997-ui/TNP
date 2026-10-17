# Bug Verification: the PROMPT-040 persistence-failure test no longer depends on POSIX semantics

- **Slug**: `prompt040-windows-test-portability`
- **Tested**: 2026-10-17 (UTC)
- **Assessment**: `./assessment.md`
- **Fix**: `./fix.md`
- **Branch**: `arena/b092bbeb-tnp` (PR #16), starting from `8c00d9bf76500ed52a5b1bf518f29be7c906cf03`
- **Result**: partial — verified on Linux including a simulated Windows attribute surface; **the real Windows run has not been performed on this host** and is the next step
- **Spec Kit execution**: Executed from checked-in Spec Kit skill instructions (`speckit-bug-test`); native Arena skill invocation was not available. Slug came from conversation context.
- **Extension hooks**: `.specify/extensions.yml` has `hooks: {}` — no `before_bug_test` / `after_bug_test` hooks are registered, so none executed.

## Summary

The reported Windows failure reproduces exactly on Linux once the Windows attribute surface is simulated
(`del os.geteuid`): `AttributeError: module 'os' has no attribute 'geteuid'` at line 427. After the test-only fix the
same exact test and the whole file pass both normally and with `os.geteuid` absent, and the test still detects a lying
DTO or a production `save_settings` that stops swallowing `OSError` (two mutations, both killed). Production code and
every other test are unchanged; the full suite is green at the same totals as the previously accepted branch state.

## Checks Performed

| Check | Command / action | Result | Notes |
|-------|------------------|--------|-------|
| Reproduction (pre-fix) | `python -c "del os.geteuid; pytest.main([...::test_pick_reports_when_the_selection_could_not_be_persisted"])"` | reproduced | Exact Windows error: `AttributeError: module 'os' has no attribute 'geteuid'` at `tests/test_prompt040_excel_persistence.py:427`. |
| Exact test after fix (Linux) | `.venv/bin/python -m pytest -q "tests/test_prompt040_excel_persistence.py::test_pick_reports_when_the_selection_could_not_be_persisted"` | pass | **1 passed**. |
| Exact test after fix (Windows attribute surface) | same node, run in a process that deleted `os.geteuid` | pass | **1 passed** — the Windows-only failure mode is gone. |
| Full PROMPT-040 file (Linux) | `.venv/bin/python -m pytest -q tests/test_prompt040_excel_persistence.py` | pass | **21 passed**. |
| Full PROMPT-040 file (Windows attribute surface) | same file with `os.geteuid` deleted | pass | **21 passed**. |
| Assertion strength (mutation A) | `ApplicationService._persisted_excel_flags = lambda self: {"template": True, "output": True}` then run the exact test | killed (test fails) | `assert True is False` on `persisted is False` — a lying DTO is detected. |
| Assertion strength (mutation B) | replace `GuiController.save_settings` with a variant whose `except OSError` is removed, then run the exact test | killed (test fails) | `PermissionError` propagates and the pick is no longer `ok` — the production catch is genuinely exercised. |
| Bridge / controller / config regressions | `pytest -q tests/test_application_service.py tests/test_gui_controller.py tests/test_prompt036r_bridge_hang.py tests/test_core.py tests/test_gui_lifecycle.py tests/test_auto_update_check.py` | pass | **97 passed**. |
| Full available suite | `.venv/bin/python -m pytest -q` | pass | **1,332 passed, 3 skipped** (2 × “git history unavailable” shallow checkout; 1 × Windows-only CLR test). |
| Static checks | `pyflakes tests/test_prompt040_excel_persistence.py`; `git diff --check` | pass | Clean. |
| Frontend build / freshness (unchanged sources) | `npm run build` (`tsc -b` + vite); `desktop.frontend_dist_freshness()` | pass | Build OK; `{'stale': False, 'reason': 'fresh'}`. |
| Real Windows rerun | `RUN_TESTS.bat` on Windows with this branch head | not-run | This host is Linux; the simulation reproduces the same attribute surface but is **not** Windows acceptance. |

## Output Excerpts

```
# pre-fix reproduction (Windows attribute surface simulated)
tests/test_prompt040_excel_persistence.py:427: AttributeError
E       AttributeError: module 'os' has no attribute 'geteuid'
1 failed in 0.40s

# post-fix
1 passed in 0.20s                                  # exact test, Linux
1 passed in 0.17s                                  # exact test, os.geteuid absent
21 passed in 1.56s                                 # full file, os.geteuid absent
21 passed in 1.16s                                 # full file, Linux

# mutations
assert True is False                               # lying DTO is caught
PermissionError: [Errno 13] …                      # missing production OSError handler is caught

# regressions / full suite
97 passed in 10.98s
1332 passed, 3 skipped in 134.66s
```

## Residual Risks

- **Real Windows execution not performed here**: the fix removes the only platform-specific construct in the test
  (`os.geteuid` / `chmod`), and the same attribute surface was simulated, but the genuine Windows run must confirm it.
- The injected `PermissionError` models the same `OSError` family that Windows raises for an unwritable or locked
  `config/` folder (`PermissionError` / `FileNotFoundError` / `OSError` on replace). The production handler catches
  `OSError`, so all of those take the identical path.
- Windows file locking semantics on `config.json` itself (e.g. an AV scanner holding the file during `os.replace`) are
  not reproduced; they are outside the scope of this test and unchanged by the fix.

## Recommendation

Re-run the Windows targeted acceptance on the new branch head. Expected: **21 passed, 0 failed**:

```bat
python -m pytest -q tests\test_prompt040_excel_persistence.py
python -m pytest -q tests\test_prompt040_excel_persistence.py::test_pick_reports_when_the_selection_could_not_be_persisted
python -m pytest -q tests\test_application_service.py tests\test_gui_controller.py
RUN_TESTS.bat
```
