# Bug Test & Verdict: Production Portable self-update

- **Slug**: `prompt033-production-self-update`
- **Date**: 2026-10-12
- **Starting SHA**: `c92f7e423a04a3ca02d9ee1a25324f9294d6a8b2`
- **Working branch**: `arena/c27c3203-tnp`
- **Assessment**: `.specify/bugs/prompt033-production-self-update/assessment.md`
- **Fix record**: `.specify/bugs/prompt033-production-self-update/fix.md`
- **Spec Kit execution**: Executed from checked-in Spec Kit skill instructions; native Arena skill invocation was not available.

## Verdict

**PARTIAL — the implementation and Linux contract/regression tests pass; Windows packaged-runtime acceptance remains pending.**

The full Python suite, affected frontend suite, typecheck, production frontend build, compile, changed-path lint, and diff validation pass. Tests include simulated package replacement, failure/rollback, preserved user data, a real second-process activity-lock case, the helper's stale-target recheck, service/bridge permission state, and frontend fail-closed DTO handling.

No Windows Portable executable, WebView2 session, configured LAN access, Windows helper handoff, real replacement/restart, or Windows rollback was exercised here. Linux results are **not** Windows acceptance. Release identity remains **v1.3.5 / Build 018**; no release was built or published.

## Executed verification

### Full Python suite — after the final updater/service/test changes

```text
$ .venv/bin/python -m pytest -q
1226 passed, 9 skipped in 118.31s (0:01:58)
```

The skips are the suite's existing platform/fixture-conditional tests; no failure was hidden or converted to a pass.

### Focused final regressions

```text
$ .venv/bin/python -m pytest \
    tests/test_updater.py::test_oversized_manifest_is_rejected_before_unbounded_read \
    tests/test_updater.py::test_helper_rejects_oversized_stage_manifest_before_parsing \
    tests/test_updater.py::test_helper_rechecks_target_identity_after_acquiring_shared_lock \
    tests/test_updater.py::test_update_lock_probe_distinguishes_access_failure_from_contention \
    tests/test_application_service.py::test_failed_update_state_never_advertises_install_enabled -q
5 passed in 1.45s

$ .venv/bin/python -m pytest tests/test_updater.py::test_oversized_backup_manifest_is_rejected_before_parsing -q
1 passed in 0.41s
```

The broader affected updater/service/progress/startup/data-preservation group also passed during validation; the final full suite above was run after all final code edits.

### Frontend tests, typecheck, and production build

```text
$ cd frontend && npm run test:frontend && npm run typecheck && npm run build
```

- Frontend tests: **36 passed, 0 failed**.
- Typecheck: **PASS** (`tsc --noEmit`).
- Production frontend build: **PASS** (`tsc -b && vite build`; 1,588 modules transformed).

This run includes the final fixture-only cleanup and `tests/update-state.test.mjs` coverage for packaged install permission, development-only discovery, malformed/contradictory DTO fail-closed behavior, and staging progress.

### Compile, lint, and whitespace validation

```text
$ .venv/bin/python -m py_compile \
    app/updater.py app/gui_controller.py app/application_service.py \
    tests/test_updater.py tests/test_application_service.py
PASS

$ .venv/bin/python -m pyflakes \
    app/updater.py app/gui_controller.py app/application_service.py app/bridge.py \
    tests/test_updater.py tests/test_application_service.py
PASS

$ git diff --check
PASS
```

### High-value transaction and lifecycle regressions included in the full suite

| Contract | Test evidence | Result |
|---|---|---|
| Source/development mode may discover but cannot stage/install | `test_production_install_gate_and_shared_activity_lock`; service-disabled install contract; frontend development DTO test | PASS |
| Normal update eligibility is numeric and rejects same/older builds | updater ordering/anti-downgrade tests and service boundary | PASS |
| Update, helper, and rollback-backup manifests are size-bounded | `test_oversized_manifest_is_rejected_before_unbounded_read`; `test_helper_rejects_oversized_stage_manifest_before_parsing`; `test_oversized_backup_manifest_is_rejected_before_parsing` | PASS |
| ZIP traversal/package layout checks do not reach outside fresh staging | package safety and `test_zip_slip_rejected` | PASS |
| Configured package is copied into local staging before installation | `test_package_copied_to_local_staging_before_installation` | PASS |
| SHA mismatch cannot advance to apply | `test_sha_mismatch_never_proceeds_to_apply` | PASS |
| Another OS process holding the activity lock prevents packaged scan and batch start | `test_another_process_lock_blocks_packaged_scan_and_batch_until_done` | PASS |
| Stale helper cannot replace a target changed while it waited for the lock | `test_helper_rechecks_target_identity_after_acquiring_shared_lock` | PASS |
| Simulated successful replacement preserves runtime/user data | `test_application_files_replaced_and_user_data_preserved`; learning-data-preservation regressions | PASS |
| Simulated replacement failure restores prior app files | `test_failed_install_triggers_rollback` | PASS |
| Failure/rollback DTO cannot claim a new install is enabled | `test_failed_update_state_never_advertises_install_enabled` | PASS |
| Bridge close cannot be requested before a bound window/handoff | `test_bridge_close_without_bound_window_does_not_mutate_close_state` | PASS |
| Progress snapshots reflect real staging state and never authorize a second install | service/update-progress/frontend DTO tests | PASS |

## Exact disposable-copy Windows bootstrap and retest guide

### Bootstrap limitation — read before testing

The deployed **Build 018** executable predates this change and has the unconditional install lock identified in the assessment. It cannot bootstrap itself into this updater-capable code from the current Build 018 LAN package. **Bootstrap verdict B:** a one-time manual installation of a later updater-capable Portable release is required.

No Build 019 or other release was created here. Do not run `BUILD_AND_PUBLISH.bat`, the release builder, or publisher for this retest; do not overwrite the LAN Build 018 package. Wait for an independently authorized updater-capable Portable bootstrap package with a build greater than 018. To exercise an actual self-update after bootstrapping, also obtain a separate disposable test ZIP/manifest with a build greater than the bootstrap build. If those authorized test artifacts are unavailable, stop: source-mode checks and Linux simulation are not production Windows acceptance.

### A. Create isolated local copies

1. Close all TNP instances. Do not point the test at the live Portable folder, the production update share, `H:`, or a stable backup.
2. Use a local disk folder writable by the same standard Windows account that will run the app. The example below creates a disposable path with spaces; choose another local path if appropriate.

```powershell
$stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$testRoot = Join-Path $env:TEMP "TNP PROMPT-033 $stamp"
$bootstrap = Join-Path $testRoot 'Portable updater-capable copy'
$updateSource = Join-Path $testRoot 'Test update source'
New-Item -ItemType Directory -Force -Path $bootstrap, $updateSource | Out-Null
```

3. Keep the original Build 018 untouched. Manually extract/copy the **authorized later updater-capable Portable bootstrap** into `$bootstrap` (or copy that exact already-installed test Portable into `$bootstrap`). Ensure the folder itself contains `ReportExtractor.exe`, `_internal`, `frontend`, and `VERSION.txt`; do not copy only the EXE. Record the bootstrap's version/build and its source archive SHA-256. Do not install it over the user's working app.
4. Copy only the authorized, higher-build **test** ZIP and its matching `version.json` into `$updateSource`. Do not edit a production manifest, reuse a same/older release as a fake newer build, or place a test package on the live LAN release folder.
5. Check package identity/size/hash before starting the app:

```powershell
$manifest = Get-Content -Raw (Join-Path $updateSource 'version.json') | ConvertFrom-Json
$zipPath = Join-Path $updateSource $manifest.package
$zip = Get-Item -LiteralPath $zipPath
$hash = (Get-FileHash -Algorithm SHA256 -LiteralPath $zipPath).Hash.ToLowerInvariant()
[pscustomobject]@{
  Version = $manifest.version
  Build = $manifest.build
  Package = $manifest.package
  ActualBytes = $zip.Length
  DeclaredBytes = $manifest.size
  ActualSha256 = $hash
  DeclaredSha256 = $manifest.sha256.ToLowerInvariant()
}
```

Proceed only if the test package came from the authorized test source and its declared size/hash match the computed values. The application must read this folder from its configured update location; do not bypass the configured-path flow.

### B. Seed disposable user data and take a baseline

Create harmless sentinel files only inside `$bootstrap`; do not copy customer data into this test. This makes preservation observable across the transaction.

```powershell
$sentinels = @(
  (Join-Path $bootstrap 'config\prompt033-sentinel.json'),
  (Join-Path $bootstrap 'logs\prompt033-sentinel.log'),
  (Join-Path $bootstrap 'Output\prompt033-sentinel.xlsx'),
  (Join-Path $bootstrap 'backup\prompt033-sentinel.txt'),
  (Join-Path $bootstrap 'learning_data\prompt033-sentinel.txt'),
  (Join-Path $bootstrap 'prompt033-user-sentinel.xlsx'),
  (Join-Path $bootstrap 'history.json')
)
foreach ($path in $sentinels) {
  New-Item -ItemType Directory -Force -Path (Split-Path -Parent $path) | Out-Null
  Set-Content -LiteralPath $path -Value "PROMPT033-$stamp" -NoNewline
}
$before = $sentinels | ForEach-Object {
  [pscustomobject]@{ Path = $_; Hash = (Get-FileHash -Algorithm SHA256 -LiteralPath $_).Hash }
}
```

### C. Exercise discovery, gates, install, and success

1. Start the copied bootstrap as the standard user, from its own folder:

```powershell
Start-Process -FilePath (Join-Path $bootstrap 'ReportExtractor.exe') -WorkingDirectory $bootstrap
```

2. In **Settings → Cập nhật**, confirm the app identifies itself as a packaged Portable runtime. Set the update path using the folder picker/text field to `$updateSource`; keep the normal configured-path mechanism and current Windows account rights. No credentials should be requested or stored.
3. Select **Kiểm tra cập nhật**. Confirm the newer test build is reported. The button must be enabled only if the backend says the package is newer, the current installation passes the production/root/layout gate, and the shared lock is free. Confirm the reported current/remote build ordering; same-build and older-build manifests must never permit installation.
4. Click **Cập nhật ngay** once. Observe stage/progress and the explicit helper-handoff state. Allow the application to close and the helper to finish; do not kill the helper or manually replace files. The staged copy must be under the local Portable folder, even though the configured update source may later be tested from an authorized LAN test share.
5. Confirm the new process restarts from `$bootstrap` and reports the authorized test package's version/build. Inspect `logs\update.log`, the consumed/result state if still present, and `update_backup` to confirm helper milestones and a verified old application backup were recorded. Confirm the user/runtime sentinels still exist with their original content/hash:

```powershell
$after = $sentinels | ForEach-Object {
  [pscustomobject]@{ Path = $_; Hash = (Get-FileHash -Algorithm SHA256 -LiteralPath $_).Hash }
}
Compare-Object $before $after -Property Path, Hash
```

Expected: no output from `Compare-Object`. Also verify the configured update location/settings remain intact. No elevation should be needed for a user-writable disposable root. Record the Windows version, exact bootstrap/test package builds and SHA-256 values, app result, and secret-free update log milestones.

### D. Exercise a safe Windows rollback in a second disposable copy

Do not reuse `$bootstrap` after the success case. Make a second fresh local copy of the same updater-capable bootstrap, seed the same kind of harmless sentinels, and use a separately prepared **test-only** higher-build package whose structure/hash/layout passes validation but whose new executable is deliberately unable to start. This test artifact must never be published, copied to the production update share, or used against any non-disposable root.

Run the same check/stage/helper flow. Expected: replacement starts, the failed new process is detected within the bounded startup grace, only application-owned entries are restored from the verified backup, the prior application is relaunched if Windows permits it, the result reports rollback/failure (not success), and all sentinels are byte-for-byte unchanged. If no authorized rollback-test package exists, leave Windows rollback **untested**; do not improvise against the live app.

### E. Source-mode and failure-stop checks

- Optional source-mode Windows check: launch the source app without a frozen executable, configure the disposable update source, and verify discovery still works while installation is explicitly disabled. This is a development-gate check, **not** packaged acceptance.
- Verify package/manifest/hash/layout rejection leaves the copied live application tree unchanged. Check denied access, lock contention, active processing, and a second app instance only within a disposable root. Do not elevate to work around a failure.
- Stop and preserve logs if root identity/layout changes, a backup is incomplete/mismatched, the helper cannot obtain the shared lock, old-process wait times out, user data changes, rollback cannot be verified, or any unexpected file outside the allowlist changes. Do not retry against the working Portable installation.
- For an authorized LAN test source, select the test share through Settings using the current account. Never embed that path in source, copy the ZIP directly over the application, request a password, or alter the production release share.

## Uncompleted platform acceptance

- Windows packaged `ReportExtractor.exe` capability/root gate and actual WebView2 startup.
- Windows standard-user ACL/drive detection and configured LAN share access.
- Real inter-process Windows lock semantics and old-process file-handle release.
- Separate helper process launch, replacement, restart, and rollback/old-app relaunch on Windows.
- Customer-data preservation and Windows update logs on a disposable Portable copy.

These items are intentionally pending until the required later updater-capable bootstrap and higher-build disposable test artifacts exist. The Linux suite is evidence for deterministic logic only.
