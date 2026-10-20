# Bug Assessment: PROMPT-040 persistence-failure test is not portable (uses `os.geteuid` / POSIX `chmod`)

- **Slug**: prompt040-windows-test-portability
- **Created**: 2026-10-17 (UTC)
- **Source**: Windows targeted acceptance of `tests/test_prompt040_excel_persistence.py` → 20 passed, 1 failed
- **Verdict**: valid (reproduced on Linux by removing the Windows-absent `os.geteuid` attribute)
- **Severity**: low for the product, **high for acceptance** — the production behaviour under test was never exercised on Windows; the suite cannot go green there and the reported contract (`persisted: false`) is unproven on the target OS
- **Starting SHA**: `8c00d9bf76500ed52a5b1bf518f29be7c906cf03` = `refs/heads/arena/b092bbeb-tnp` (PROMPT-040 fix + verification reports, PR #16)
- **Spec Kit execution**: Executed from checked-in Spec Kit skill instructions (`speckit-bug-assess` → `speckit-bug-fix` → `speckit-bug-test`); native Arena skill invocation was not available.
- **Extension hooks**: `.specify/extensions.yml` contains `hooks: {}` — no hooks are registered, so none executed.

## Report (summarized)

Windows ran the PROMPT-040 regression file and got **20 passed, 1 failed**. The single failure is
`test_pick_reports_when_the_selection_could_not_be_persisted`
(`tests/test_prompt040_excel_persistence.py:427`):

```
AttributeError: module 'os' has no attribute 'geteuid'
```

The test opens with `if os.geteuid() == 0: pytest.skip("root ignores directory permissions")` and then attempts to
make the config folder unwritable with `os.chmod(config_path.parent, 0o500)`. Neither construct means the same thing
on Windows: `os.geteuid` does not exist at all, and POSIX mode bits do not govern directory write access on Windows
(`os.chmod` only touches the read-only attribute, which Windows ignores when creating files inside a directory).

## Symptom

Expected on any OS: the test proves the production contract “a pick whose config write genuinely fails is still usable
for the session and is reported with `persisted: false`, never silently claimed as saved”.

Actual on Windows: the test cannot even start — `AttributeError` before any assertion, so the contract is **untested**
on the target platform and the suite is red for a reason unrelated to the product.

## Reproduction

```bash
# Simulate the Windows attribute surface on Linux: os.geteuid simply does not exist there.
.venv/bin/python - <<'PY'
import os, sys, pytest
del os.geteuid
sys.exit(pytest.main(["-q", "tests/test_prompt040_excel_persistence.py::"
                          "test_pick_reports_when_the_selection_could_not_be_persisted"]))
PY
```

Observed (exact, matches the Windows report):

```
tests/test_prompt040_excel_persistence.py:427: AttributeError
E       AttributeError: module 'os' has no attribute 'geteuid'
1 failed in 0.40s
```

## Traced flow (test → production)

1. **Test setup** — `tests/test_prompt040_excel_persistence.py:429-437` builds a real `ApplicationService` over a
   `tmp_path` config, binds a scripted dialog, stages the pick, then `os.chmod(config_path.parent, 0o500)`.
2. **Production write path** — `BridgeService._choose_excel` → `ApplicationService.persist_excel_selection`
   (`app/application_service.py:225`) → `_persist_excel_settings_unlocked` (`:271`) → `GuiController.save_settings()`
   (`app/gui_controller.py:2171`). The controller copies the fields and calls `c.save(self._config_path)`.
3. **Disk boundary** — `AppConfig.save` (`app/config.py:184-200`) serialises to `config.json.tmp` and
   `os.replace(tmp, target)`. An unwritable folder raises `OSError` here.
4. **Production error contract** — `GuiController.save_settings` already wraps that call:
   `except OSError as e: self.log_lines.append(f"Không lưu được config.json: {e}")`. No crash, no partial file.
5. **Honest reporting** — `_persisted_excel_flags` (`app/application_service.py:281`) re-reads
   `AppConfig.load(c.config_storage_path())` instead of trusting memory, so an absent/old file yields
   `persisted: false`, while the picked path stays active in `GuiController.template`.
6. **Where the test breaks** — the test *creates* that failure only through POSIX directory permissions. It never
   injects the failure at the disk boundary, so on Windows (a) it crashes on `os.geteuid`, and (b) even with that
   guarded, `chmod` would not stop the write and the assertion `persisted is False` would fail for the wrong reason.

## Root Cause Hypothesis

**Confidence: high, and it is test-only.** The test encodes the failure using Unix account semantics
(`os.geteuid`) and POSIX permission bits (`chmod 0o500`). Neither is part of the production contract, and neither
exists in the same form on Windows. Production never calls `os.geteuid`, `os.getuid`, `os.chmod` or `os.access`
(verified: `grep -rn "geteuid\|getuid\|os\.chmod\|access(" app/ tools/ run.py` → no matches), and the production
failure contract (`save_settings` swallowing `OSError` + `persisted` re-read) is already platform-independent.
There is **no production defect to fix**: the only defect is the test's failure injection mechanism.

## Proposed Remediation

**Preferred**: inject the write failure at the real disk boundary, deterministically and identically on every OS.

1. Replace the `os.geteuid` guard and `chmod` calls in
   `tests/test_prompt040_excel_persistence.py::test_pick_reports_when_the_selection_could_not_be_persisted` with a
   monkeypatched `AppConfig.save` that raises `PermissionError` (an `OSError`) — exactly what an unwritable or locked
   `config/` folder produces on Windows. The real `GuiController.save_settings()` body still runs, including its
   production `except OSError` handler, so the test proves the product contract rather than a bypass.
2. Keep and strengthen the intended assertions, so the test still proves behaviour under a *genuine* persistence
   failure:
   - the pick response is still `ok: true` with the resolved path;
   - the config file really was **not** written (`config_path.exists()` is false);
   - the production handler ran (the `"Không lưu được config.json"` log line is present);
   - the state DTO still carries the picked path and state, with `persisted: false`;
   - the failure is **not sticky**: after restoring `AppConfig.save`, persisting the same selection reports
     `persisted: true` and the file then contains the path.
3. Do not depend on uid/root, `chmod`, administrator state or real protected directories anywhere.

**Files likely to change**: `tests/test_prompt040_excel_persistence.py` only, plus this bug directory (assessment,
fix, test reports). No production file is expected to change.

**Tests to add or update**:
1. the rewritten `test_pick_reports_when_the_selection_could_not_be_persisted` (works on Windows, macOS and Linux,
   root or not);
2. the rest of `tests/test_prompt040_excel_persistence.py` must stay green (unchanged behaviour);
3. bridge/controller/config regression suites must stay green;
4. the full available Python suite must stay green.

## Risks & Considerations

- **Weakening risk**: replacing a real read-only folder with an injected `OSError` could drift into “testing the mock”.
  Mitigated by asserting the production catch path executed (log line), that the file is genuinely absent, and that
  the same selection persists once writes succeed again.
- **Linux coverage**: the injected failure is identical on Linux; the previous POSIX test could only run as non-root
  anyway and is already skipped for root, so no meaningful Linux coverage is lost. Real read-only paths remain
  covered indirectly by `AppConfig.save`'s own error handling (unchanged production code).
- **Not changed**: production code, the DTO contract, Excel lock/retry/atomic-commit behaviour, version/build,
  `backup/…STABLE/`, unrelated tests.

## Open Questions

- [NEEDS CLARIFICATION: none blocking. Windows acceptance of the *product* (native picker, WebView2, PowerPoint COM,
  Portable) still has to be run on Windows after this test fix — this assessment can only make the suite portable.]
