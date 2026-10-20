# Bug Fix: Safe production Portable self-update

- **Slug**: `prompt033-production-self-update`
- **Date**: 2026-10-12
- **Starting SHA**: `c92f7e423a04a3ca02d9ee1a25324f9294d6a8b2`
- **Working branch**: `arena/c27c3203-tnp`
- **Assessment**: `.specify/bugs/prompt033-production-self-update/assessment.md`
- **Spec Kit execution**: Executed from checked-in Spec Kit skill instructions; native Arena skill invocation was not available.
- **Release identity**: v1.3.5 / Build 018; unchanged. No Portable release was built or published.

## Implemented fix

The existing configured-folder discovery flow remains read-only and is reused. Installation is now reachable from the default React + pywebview UI only when the backend proves it is running from a valid, local, frozen Portable root. Every install boundary repeats that policy; the frontend's DTO/button state is not trusted as authorization.

### Settings → configured source → build/manifest → bridge/UI

1. Settings persists the existing configured update folder and `autoCheck` setting. `ApplicationService` performs the existing policy-controlled, one-shot asynchronous startup check when enabled; it does not add an update polling loop.
2. `GuiController.check_update_async()` calls `updater.check_for_update()` with that configured path. The checker reads a bounded `version.json`, validates version/build/package/SHA-256/optional size, requires a safe package basename and direct ordinary package file, checks package size, and compares numeric build numbers. Same/older builds remain ineligible. It does not copy or extract from the configured folder.
3. `ApplicationService._update_dto()` publishes runtime capability, available package metadata, progress, status/error, install permission, and a specific disabled reason. It fails closed for non-Portable/development mode, busy work, lock contention, or lock access failure. Failed/rollback states never advertise an enabled install.
4. `BridgeService` preserves the standard response envelope for check/install. Installation is backend-enforced. A separate `close_for_update` bridge call requires a bound window and a confirmed helper handoff; the window is destroyed only after the bridge response can return.
5. React validates the complete update DTO and rejects missing, malformed, or contradictory permission/status combinations. Settings labels development versus packaged Portable, presents same/older/no-update/error/progress/handoff states, enables installation only for an available newer build and explicit backend permission, and offers a close retry if a confirmed helper handoff cannot close the window.

### Production policy and process lifecycle

- `production_install_capability()` requires a frozen executable named `ReportExtractor.exe`, the exact executable-parent Portable root, a verified current VERSION/build and application layout, and a Windows local drive. Network/UNC/mapped-remote roots fail closed; source/development installations can still discover updates but cannot stage, invoke the helper, or install.
- The service, controller, staging function, updater command, and helper each enforce the relevant production/newer-build/root contract. Calling a bridge/UI method directly cannot bypass it.
- A shared inter-process lock serializes packaged scans/batches and update installation. The installing application holds it through handoff; the helper acquires it before replacement. Existing service idle/close guards continue to protect active report processing, scans, diagnostics, training, and server-apply work. Lock contention and inability to open/create the lock are reported distinctly and fail closed.
- Process-exit and restart checks are bounded. The helper rechecks the target identity/layout after it gets the lock, so a stale second helper cannot replace a target changed while waiting.

### Staging, package ownership, backup, replacement, rollback

- The ZIP is copied from the configured local/LAN source to a unique UUID-named stage below the Portable root. The source is never used as an extraction or installation workspace. The copied bytes are checked against required SHA-256 and declared size before extraction. Update, helper, and rollback-backup manifests are size-bounded.
- A custom ZIP plan/extractor rejects traversal, absolute/drive/UNC names, ADS and reserved Windows components, noncanonical names, case collisions, duplicate entries, symlinks/reparse/special files, encrypted members, oversized expansion, excessive compression ratios, unexpected package roots, and packaged runtime/user data. The package must contain the expected Portable executable, frontend/bridge layout, and matching bounded `VERSION.txt`; the executable is not the running executable.
- Replacement is restricted to the exact package-owned root allowlist (`ReportExtractor.exe`, `_internal`, `frontend`, `VERSION.txt`, and the documented release support files). Unknown root entries are not swept. `config/`, `logs/`, `Output/`, `backup/`, `learning_data/`, update staging/backup/lock state, settings/history/manual data, and root-level user spreadsheet/presentation/log/JSON files remain outside replacement.
- Before destructive work, the helper makes an application-owned backup and verifies tree signatures against the source. It replaces only allowlisted package entries and validates the installed tree. A replacement or new-process startup failure triggers rollback from the verified backup; if safe rollback/relaunch cannot be proven, the helper records failure rather than claiming success. Backup pruning retains bounded history.
- Stage progress, helper handoff, result state, and secret-free update log events are connected to the service/frontend. Failure before handoff keeps the old application open and cleans only updater-owned staging. Terminal results stay visible until the helper exits; updater-owned stage cleanup is coordinated under the same lock.

## Tests and code deltas

The implementation reuses the existing updater and publisher contract rather than introducing a second updater. New and expanded tests cover strict manifest/package inputs, bounded manifests, secure ZIP layout, frozen/source gates, numeric ordering, lock failures/contention, active cross-process scan/batch exclusion, target identity recheck after lock wait, actual DTO permission/progress, bridge close guards, frontend fail-closed normalization, and simulated installation/rollback with persistent user-data preservation. The final executed test results are in `test.md`.

A final DTO regression also verifies that a failed or rolled-back install with a still-available remote package cannot accidentally serialize `installEnabled: true`.

## Changed paths

- `app/updater.py` — strict discovery, bounded manifest/version reads, package verification/extraction, production gate, shared lock, staging, ownership allowlist, verified backup, helper, replacement, restart and rollback.
- `app/gui_controller.py` — production gate, process-wide activity/install lock lifecycle, bounded worker/handoff/error/progress handling.
- `app/application_service.py` — startup check policy, production update DTO/install boundary, idle and handoff lifecycle, helper result consumption.
- `app/bridge.py`, `app/gui.py` — guarded update-window close and legacy UI disable reason.
- `frontend/src/state/store.tsx`, `frontend/src/tabs/SettingsTab.tsx`, `frontend/package.json`, `frontend/tests/update-state.test.mjs` — typed/fail-closed update state, install/handoff UI, and regression coverage.
- `tests/test_updater.py`, `tests/test_application_service.py`, `tests/test_update_progress.py`, `tests/test_auto_update_check.py` — backend/service lifecycle, lock, update, package and rollback regressions.
- `tests/test_content_learning.py`, `tests/test_image_learning.py` — fixtures explicitly emulate packaged runtime where existing learning-data-preservation tests exercise the production updater gate.
- `.specify/bugs/prompt033-production-self-update/{assessment,fix,test}.md` — checked-in bug workflow evidence.

## Scope preserved

- Version/build remain **1.3.5 / Build 018**. No Build 019 was created.
- No hardcoded LAN folder, address, share credential, or password was added. Access uses the current Windows session and the configured update location.
- No release build/publish, LAN Build 018 overwrite, `BUILD_AND_PUBLISH.bat`, `H:`, stable backup, `TAT_QPN_PHASE6/`, PR #7/main merge, reset, rebase, or force-push was performed.
- The code-level Linux tests are not Windows packaged-runtime acceptance.

## Bootstrap decision

**B — a one-time manual installation of a later updater-capable Portable release is required.** The deployed Build 018 executable predates this implementation and has an unconditional install lock; it cannot bootstrap itself from the existing Build 018 LAN package. This change deliberately does not build or publish a replacement. Windows production retest must wait for an authorized later updater-capable Portable package and a separate, strictly higher-build test package. See `test.md` for the disposable-copy procedure and stop conditions.
