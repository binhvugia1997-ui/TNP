# PROMPT-039 assessment — Windows full-suite integration regressions

- **Slug**: `prompt039-windows-full-suite-regression` (generated in automated mode; new unique directory)
- **Assessed**: 2026-10-15 (UTC)
- **Reported result**: Windows full suite, 18 failed / 1,262 passed
- **Checkout baseline**: `bb1fd4664b00c5a8ae019f6330992320d020e353` (`release: v1.3.7 build 020`), exactly the requested commit. The Arena session branch is `arena/b9301dab-tnp`; the fetched `origin/integration/prompts-034-035-036-037` ref is not present in this clone. Initial worktree was clean; the prompt's two named local files were absent and were not touched.
- **Environment**: Linux, Python 3.11.2. This is development evidence only, not Windows acceptance.
- **Verdict**: valid test regressions with three shared/independent causes; production Portable and PowerPoint ownership behavior must remain strict.
- **Severity**: high for release confidence (the Windows suite is blocked), low for runtime behavior based on the supplied evidence (the failures are predominantly test/fixture nondeterminism, not 18 independent production defects).

## Evidence captured before fixes

### Cluster A — Windows Portable/updater fixture failures (cascading)

- `app.updater.validate_staged_portable()` deliberately activates Windows-only validation when `sys.platform == "win32"`. That contract requires `_internal/pythonnet/runtime/Python.Runtime.dll`, a `_internal/clr_loader/ffi/dlls/*/ClrLoader.dll`, the WebView2 Core and WinForms assemblies, and the win-x64 `WebView2Loader.dll`. The `MZ` check on the synthetic EXE and frontend/pywebview bridge checks are separate required parts of the layout.
- The test-only `make_release_zip()` and `make_portable()` builders in `tests/test_updater.py` model the older minimal layout and omit those Windows runtime files. `tests/test_update_progress.py` imports both builders; content/image-learning tests share the staged helper from `tests/test_updater.py`.
- Forced-Windows validation of both existing fixtures returned `(False, 'thiếu .NET/pythonnet runtime bắt buộc')`. The same fixtures pass the Windows-runtime portion on Linux only because the production function correctly selects its host-dependent branch. This explains why the failures cluster on Windows and why they cascade before each test reaches its intended staging/progress/rollback behavior.
- The production contract matches the build/package hooks and the existing PROMPT-028/029 packaging validation. Weakening `validate_staged_portable()` is not a valid remediation.

### Cluster B — renderer diagnostic test assumes local machine state

- `tests/test_prompt027_rendering.py::test_backend_failure_is_logged_not_silently_swallowed` expects `reason=unavailable` while using the real `powerpoint_available()` probe and real PowerPoint backend.
- On Windows where PowerPoint is installed/running, the probe reports available and the renderer attempts the ownership-protected backend. Production then correctly refuses to attach to the user's registered PowerPoint (`stage=ownership`, reason identifies that a running instance is not safe to attach to), and falls through to builtin. That is not an `unavailable` result.
- Existing deterministic coverage already exists for true unavailability (`tests/test_prompt027r_renderer.py`) and ownership refusal/fallback (`tests/test_prompt035_powerpoint_isolation.py`). The failing test needs a fake backend/probe state and assertions on its intended injected failure, not a production behavior change.

### Cluster C — freshness test fixture timestamps an included source after its reference time

- The production freshness input list intentionally includes `src`, `index.html`, Vite/Tailwind/PostCSS configuration, and `package.json`. `frontend/index.html` is the Vite entry document; excluding it would hide a real source change. The freshness function examines files, not directory mtimes, and does not recurse into `dist`, `node_modules`, test, or cache output.
- In `tests/test_prompt038_desktop_freshness.py`, `now = time.time()` is captured before `_make_frontend()` writes `frontend/index.html`. The fixture then assigns explicit times only to `src/index.css` and `dist/index.html`, leaving the legitimate `frontend/index.html` at its later write time. Thus the supposedly fresh dist can appear stale, and the `newest_source` assertion can report `index.html` instead of `index.css`.
- The source walk uses `rglob()` without sorting. Equal mtimes do not make the stale/fresh result ambiguous (`source_mtime > dist_mtime` is the strict stale rule), but they can make the diagnostic `newest_source` depend on filesystem enumeration order. A stable path tie-break plus explicit fixture mtimes will make the contract portable across Windows/POSIX and timestamp resolutions.

## Proposed remediation

- Add one canonical, test-only Windows runtime contract helper and use it from the shared `make_release_zip()` and `make_portable()` fixture builders. Use clearly synthetic placeholder bytes; these files model the paths required by the structural updater validator and are not executable runtime payloads. Do not duplicate construction across consumer tests.
- Add deterministic Windows-branch updater validation tests that accept the complete fixture and reject missing Python.Runtime, ClrLoader, and WebView2 runtime files. Keep the production validator unchanged.
- Make the renderer backend-failure test deterministic by monkeypatching the probe and raising a staged backend error; assert one failure record, purpose/backend/stage/reason and builtin fallback. Retain separate deterministic unavailable and ownership-refusal coverage.
- Keep `index.html` in freshness inputs. Sort candidate paths for deterministic tie reporting and revise the fixture to set timestamps for every relevant source file explicitly, with large separations where freshness is under test. Add coverage that `index.html` itself is a valid source input and that equal-time diagnostics are deterministic.
- Run the requested targeted groups, full available suite, `compileall`, `git diff --check`, and the required negative Portable tests. State explicitly that Linux results do not constitute Windows acceptance.

## Files likely to change

- `app/desktop.py` — deterministic freshness diagnostic tie-breaking only, if confirmed needed.
- `tests/portable_fixtures.py` — canonical synthetic Windows runtime file layout.
- `tests/test_updater.py` — use the helper and add updater-level Windows-runtime acceptance/rejection tests.
- `tests/test_prompt027_rendering.py` — deterministic backend failure/fallback test.
- `tests/test_prompt038_desktop_freshness.py` — timestamp-controlled source fixtures, index.html and equal-time coverage.
- `.specify/bugs/prompt039-windows-full-suite-regression/{assessment.md,fix.md,test.md}` — workflow records.

## Risks & considerations

- Never let structural test placeholders be mistaken for a real packaged runtime. Builder/runtime authenticity remains covered by the existing PROMPT-028/029 tests and production packaging hook validation.
- Do not alter PowerPoint ownership isolation or claim PowerPoint unavailable when it is actually available.
- Do not omit `index.html` or relax the source-newer freshness protection to satisfy an incorrectly timed test.
- Linux cannot prove Windows full-suite acceptance; the final recommendation must include the exact PowerShell retest command.

## Open questions

- None. The supplied failure patterns and the code/test inspection identify the remediation without changing production Portable or renderer behavior.
