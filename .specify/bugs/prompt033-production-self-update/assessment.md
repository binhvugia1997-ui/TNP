# Bug Assessment: Production Portable self-update is locked and the existing apply path is not yet safe enough

- **Slug**: prompt033-production-self-update
- **Created**: 2026-10-12
- **Source**: pasted text (PROMPT-033; real Windows observation supplied by the user)
- **Verdict**: valid
- **Severity**: high — production users can discover a release but cannot install it through the default Portable UI; enabling the present legacy apply path without hardening it would create material installation/data-integrity risk.
- **Spec Kit execution mode**: Executed from checked-in Spec Kit skill instructions; native Arena skill invocation was not available.

## Report (verbatim or summarized)

The user reports on Windows that a Portable Build 017 (1.3.4) reaches its configured LAN update directory, successfully checks `version.json`, detects 1.3.5 / Build 018, and displays “Có bản cập nhật”. “Cập nhật ngay” remains disabled and the UI says applying/installing is locked in the development integration. The LAN path is not reported as faulty. Required behavior is an explicit, safe, production-Portable-only update flow with a separate replacement process, validation, backup, rollback, preservation of persistent data, visible state, and regression coverage.

The current checkout was fetched before inspection. The assigned Arena session branch is `arena/c27c3203-tnp` (the session branch is fixed by Arena); its HEAD is exactly `c92f7e423a04a3ca02d9ee1a25324f9294d6a8b2`, the requested minimum commit, and the worktree was clean. The requested named base branch `arena/2e088bee-tnp` is not the branch assigned to this session; no branch switch was made. The required source baseline is present at the exact SHA.

Release identity is synchronized in source: `app/__init__.py` declares `__version__ = "1.3.5"`, `BUILD_NUMBER = 18`, and derives `BUILD_ID = "018"`, `BUILD_LABEL = "Build 018"`, `APP_TITLE = "Report Extractor v1.3.5"`, `VERSION_LABEL = "1.3.5 — Build 018"`, and `VERSION_LINE = "version=1.3.5 build=018"`. Frontend package metadata and `frontend/index.html` also identify 1.3.5 / Build 018. `tools/build_portable.py` writes those canonical values to packaged `VERSION.txt`; no generated Portable artifact is present in this checkout to inspect directly.

## Symptom

The reported production user can discover a newer Build but cannot apply it from the default React + pywebview Settings UI. Direct source inspection proves that installation is unconditionally disabled at both the UI and service boundary, rather than disabled because update discovery or the configured LAN location failed.

## Reproduction

1. **User-provided Windows reproduction:** launch the older Portable (1.3.4 / Build 017), configure its existing update folder, and select Settings → Update.
2. Select “Kiểm tra cập nhật”; the user reports a successful check and the expected 1.3.5 / Build 018 offer.
3. Observe “Cập nhật ngay” disabled and the development-integration explanation.
4. **Source reproduction:** in `frontend/src/tabs/SettingsTab.tsx`, the React button is rendered with `disabled={true}`. If invoked through the bridge anyway, `ApplicationService.install_update()` returns `status="unavailable"` and the development-lock message without reaching the updater. `ApplicationService._update_dto()` also serializes `installEnabled: false` unconditionally.
5. The source-side path is established; the real Windows UI was not launched in this Linux sandbox. An attempted focused pytest run could not execute because neither `pytest` nor the `pytest` Python module is installed in the initial environment.

## Suspected Code Paths

- `run.py` → `app.main.main()` → `launch_ui()` → `app.desktop.main()` — default packaged entry path. `ReportExtractor.spec` points PyInstaller at `run.py`; `app.main.main()` defaults to the React + pywebview UI, not `--legacy-gui`.
- `frontend/src/tabs/SettingsTab.tsx::UpdateSection` — calls `s.checkUpdate(false)` for discovery, hard-codes `disabled={true}` for “Cập nhật ngay”, and always displays the development-lock explanation; it also labels every runtime “development”.
- `frontend/src/state/store.tsx::checkUpdate` / `installUpdate` / `acceptDashboard` — invokes the native bridge and consumes dashboard update state. The `UpdateState` has an optional `installEnabled`, but the UI ignores it and the backend always sends false.
- `app/bridge.py::BridgeService.check_update` → `ApplicationService.save_configuration` → `ApplicationService.check_update` → `GuiController.check_update_async(startup=False)` → `app.updater.check_for_update` — the existing read-only check path. It uses the configured `update_path`, reads `version.json`, verifies that the named package exists, and orders releases by numeric build. The bridge wraps the DTO in the normal `{"ok": true, "data": ...}` envelope.
- `app/application_service.py::ApplicationService.install_update` — exact production integration lock: always returns `status="unavailable"`, with “Cài đặt bản cập nhật bị tắt trong tích hợp phát triển; không thay thế hay khởi động lại ứng dụng.”
- `app/application_service.py::ApplicationService._update_dto` — exact DTO lock: `installEnabled` is always false, regardless of frozen/Portable status or update availability.
- `app/gui_controller.py::GuiController.install_update[_async]` and `app/gui.py` — a separate legacy Tk path already calls the old updater, enables its button after a positive newer-build check, and closes its root 300 ms after helper handoff. It is not the default UI of the packaged `run.py` entry point and it has no source/development installation gate; it is not a safe workaround for the reported production lock.
- `app/updater.py::parse_manifest`, `check_for_update`, `stage_update`, `updater_command`, `launch_updater`, `apply_update`, `release_entries`, and rollback helpers — existing discovery, copy/stage, helper, wait, replacement, backup, and rollback implementation; several safety invariants required by this report are incomplete (see Root Cause Hypothesis).
- `app/runtime_paths.py::is_packaged` / `portable_root` — proven runtime identity/root primitives: `is_packaged()` uses PyInstaller’s `sys.frozen`; packaged writable state/root is `Path(sys.executable).resolve().parent`, independent of the current working directory. Source mode resolves to the checkout root.
- `app/main.py::main` `--apply-update` dispatch plus `ReportExtractor.spec` — the packaged EXE can run the existing helper mode; `app.updater` is in the PyInstaller hidden imports. The package entry point therefore contains some updater machinery, but that fact does not make the default React installation contract reachable or the present apply implementation safe enough for the required production use.
- `tools/build_portable.py::write_version_manifest`, `validate_artifact`, `validate_frontend`, and ZIP creation; `tools/publish_update.py::publish_release` — publisher contract. The client manifest is version/build/package/SHA256/size. Publishing copies/verifies the ZIP first and atomically publishes `version.json` last. The runtime client does not need a hardcoded LAN location: the configured setting is passed through `update_path`.
- `DEVELOPMENT_WEBVIEW.md` — explicitly documents the historical decision to disable update installation in the React development service/UI; the packaged application uses that same default React service.
- `tests/test_application_service.py::test_configuration_ollama_update_state_and_disabled_install`, `tests/test_updater.py`, and `tests/test_auto_update_check.py` — prior contracts cover the intentionally disabled service, the legacy updater, numeric ordering, staging, backup/rollback, and Tk startup checks. These tests require replacement/expansion for PROMPT-033, not weakening.

## Existing Updater Trace and Evidence

**Default discovery path (as found):** Settings → update path from the current configuration → `check_update` bridge method → saved configuration → asynchronous `GuiController.check_update_async` → `updater.check_for_update(configured_path)` → `version.json` parsed by `parse_manifest` → package existence → numeric `remote build > current build` comparison → controller `UpdateCheck` → service update DTO/dashboard → bridge success envelope → React store’s dashboard state → “Có bản cập nhật” status. The observed successful discovery is consistent with this source path. No observed LAN address is embedded in the runtime updater; `tools/publish_update.py` has publishing-machine defaults only.

**Existing dormant/legacy apply path (as found):** the Tk controller’s install worker invokes `stage_update`; the latter copies the ZIP into `<portable>/update_staging`, verifies it, checks ZIP members, extracts a new tree, then writes `update.json`. `updater_command` launches the staged release’s `ReportExtractor.exe --apply-update <manifest>` in packaged mode (or `python run.py --apply-update ...` in source mode) via structured argument-list `Popen`. `app.main.main()` dispatches the helper mode to `apply_update`; the helper waits up to 90 seconds for the recorded old PID, backs up/replaces top-level release entries, attempts rollback if replacement copying fails, writes a result/log, and starts the installed EXE. The Tk UI then destroys its root so the helper can continue. The `ReportExtractor.spec` bundles `app.updater` and the CLI entry path.

This is reusable implementation evidence, not proof of production readiness. Specific gaps found against PROMPT-033 include:

- The React application-service install method is a hard lock, and the frontend button is hard-disabled. The package’s normal default UI cannot reach the otherwise present updater.
- The controller/legacy source-mode install path has no backend production-only gate; `updater_command(packaged=False)` attempts to apply to the source runtime root. A UI-only gate would therefore be insufficient.
- SHA256 is optional in `parse_manifest`/`verify_package`; the client does not enforce publisher `size`; check-time package validation follows ordinary `Path.is_file()` semantics. The filename check is not cross-platform-safe for Windows separators/drive-qualified forms when evaluated on POSIX.
- The fixed `<portable>/update_staging` is cleared with `rmtree(ignore_errors=True)` rather than a unique, bounded, ownership-checked stage. Copy/extract failure cleanup is incomplete.
- ZIP validation checks for an executable and `VERSION.txt` but does not reject all path forms, drive-qualified/normalized collisions, symlink/reparse-like members, or resource-exhaustion cases before extraction. The staged release is not checked against the actual required Portable layout (`ReportExtractor.exe`, `_internal` runtime, frontend bundle, and matching VERSION metadata) before replacement.
- The helper trusts mutable manifest paths and metadata too broadly, does not revalidate a complete staged package/hash at handoff, does not require a valid original PID, and does not enforce a monotonic current-build guard at the helper API. A direct helper invocation can bypass the normal `UpdateCheck.available` check.
- Per-controller click locks exist, but there is no cross-process install lock protecting fixed staging/backup paths if multiple application instances are opened.
- `release_entries` broadly treats “not in the preserved-name/suffix set” as replaceable rather than validating a precise package-owned root layout. The build artifact includes `Output/`, `logs/`, and `config/` scaffolding, which must stay excluded from replacement. Repository-proven persistent user data includes `config/` (including update settings), `logs/`, `Output/`, `backup/`, `learning_data/`, root `config.json`, `history.json`, root-level spreadsheet/presentation/log/JSON files, and update staging/backup state. The replacement policy must explicitly preserve these and avoid sweeping unknown user files.
- Backup setup clears the previous backup with ignored deletion errors and does not verify a complete durable backup before starting destructive work. On replacement failure, rollback is attempted, but `apply_update` logs a new-EXE launch failure without restoring the prior release; the helper can therefore report success while restart fails.
- The existing runtime `update_result.json` is consumed by the legacy Tk GUI only. The default React service/frontend does not expose stage/error/handoff/restart states or consume the result/cleanup staging on startup.
- The legacy Tk path has an idle-processing guard in `GuiController`; the React service’s locked install method does not exercise it. Existing shared service shutdown guards cover processing/scan/training/diagnostics/server-apply work and should be reused before handoff.
- The Tk UI schedules automatic check-on-start after a delay. The default React store currently has no startup call to `check_update`; the Settings copy says automatic checks happen on startup, but this path must be wired/pinned for the packaged React app without adding polling loops.

## Root Cause Hypothesis

**Confidence: high.** The update-discovery system is separate and already works. Installation is intentionally disabled in the React + pywebview integration: the current `SettingsTab` hard-codes the button disabled, the service rejects every install as `unavailable`, and the DTO advertises `installEnabled: false`. That old development safety lock is unconditional, so it applies to a frozen Portable runtime too. The root cause is not a broken LAN location. The old updater module contains reusable parts, but they do not currently constitute the required safe and production-gated installer end-to-end.

## Proposed Remediation

**Preferred:** keep and harden `app/updater.py` instead of creating a second updater. Keep `check_for_update` and the configured-path contract read-only; make install availability explicit in the update DTO from a robust packaged/frozen capability check, and enforce the same gate in the backend/service/controller/helper so frontend state cannot authorize installation. Source execution must still be able to check/display updates but never stage or install.

Before enabling the production button, make the manifest strict (required version/build/package/SHA256, optional bounded size validation; safe basename across POSIX/Windows; package confined to the configured directory and regular-file checks), copy into a unique local stage, hash the exact copied bytes, securely validate/extract every ZIP member, validate the actual release layout and matching VERSION before replacement, then launch the packaged helper in a separate process with structured arguments. Use the proven PyInstaller root (`sys.executable` parent) rather than CWD. Add an exclusive update lock, busy-task guard, bounded old-PID wait, exact package-owned root allowlist, verified application-only backup, and bounded rollback that restores only application-owned files while preserving persistent data. A failed restart must attempt rollback/relaunch where safe. Report every stage/failure to a bounded update log and bridge DTO.

Wire the React UI to the explicit production capability and state: development explanation and disabled control in source mode; no-update/same/older disabled; newer Build enabled only in packaged mode; progress and explicit validation errors; single-click handoff state; no success claim until the helper starts. Reuse `ApplicationService` idle/close lifecycle and the existing app shutdown path. Preserve automatic checks in the default app (add the currently missing React startup call if confirmed by tests) using one bounded check, not polling. Retain the configured LAN path and existing publisher contract.

Add isolated temporary-directory backend, bridge, frontend, package/rollback, and simulated Portable integration tests; include paths with spaces, build ordering, source/frozen gating, duplicate launches, failure at each pre-destructive stage, backup failure, rollback/restart failures, and persistent-data preservation. Update only bounded PROMPT-033 docs/spec artifacts and provide a disposable-copy Windows acceptance procedure; do not build/publish/bump Build 019 in this work.

**Alternatives** (not preferred):
- Reimplement the updater as an unrelated helper architecture. This would duplicate the existing staging/replacement code and create two authorities for package validation/preservation rules.
- Simply remove the service lock and enable the button. Rejected: the source/frozen gate, archive/manifest hardening, helper manifest trust, backup/restart rollback, and frontend state are not sufficient for that change.

**Files likely to change**:
- `app/updater.py`, `app/gui_controller.py`, `app/application_service.py`, `app/bridge.py`, `app/desktop.py` or the established shutdown adapter as needed.
- `frontend/src/state/store.tsx`, `frontend/src/tabs/SettingsTab.tsx`, and focused frontend tests.
- `tests/test_updater.py`, `tests/test_update_progress.py`, `tests/test_application_service.py`, and focused bridge/integration tests.
- `ReportExtractor.spec` and `tools/build_portable.py` only if a shared package-layout validator or helper packaging adjustment is needed; never run the release builder/publisher.
- `DEVELOPMENT_WEBVIEW.md` / user-facing updater documentation as necessary.
- `.specify/bugs/prompt033-production-self-update/{assessment,fix,test}.md`.

**Tests to add or update**:
- Strict manifest/package/path/SHA/size and numeric-build check tests.
- ZIP traversal, absolute path, drive-qualified path, normalization collision, symlink/reparse rejection; no extraction outside a fresh stage.
- Failure-injection tests proving install root unchanged for copy/hash/extract/layout-validation/backup failures.
- Helper process wait/timeout, argument-list paths-with-spaces, single/cross-process lock, source/frozen capability, and backend anti-downgrade gates.
- Simulated Portable success + replacement failure/rollback proving package files change/restore and config/logs/Output/update settings/learning data survive.
- Bridge DTO roundtrip and exactly-once call/error contract; UI status/button/failure/handoff tests; no-update/same/older/source-mode disabled and packaged-newer enabled.
- Regression for React automatic startup check and legacy updater behavior.

## Risks & Considerations

- Linux can prove the mock/process/file transaction logic, not Windows file-lock behavior, WebView2/PyInstaller startup, permissions, UNC credentials/access, or real Portable replacement/restart. Windows disposable-copy acceptance remains mandatory.
- The existing frozen onedir package layout is clear in `ReportExtractor.spec`/`tools/build_portable.py`, and `runtime_paths.portable_root()` is independent of CWD. A real packaged EXE has not been built or run in this sandbox; the production capability must fail closed if these markers/layout checks do not hold.
- The default app shares the backend with a separately testable legacy Tk GUI; safety gates must cover all public install entry points, not just the React button/service.
- Do not modify version/build metadata, build a Windows release, publish, or change the existing LAN Build 018 artifact.
- No build/package artifact is available locally to inspect the exact historical Build 017/018 byte set. The current checked-in Build 018 source/spec proves the default packaged app launches React and includes the updater module/helper CLI, but its install endpoint is compiled as unavailable. Therefore the effective deployed Build 017/018 default runtime **cannot bootstrap PROMPT-033 by detecting the existing Build 018 LAN package**. Bootstrap verdict: **B — a one-time manual install of a later updater-capable Portable release is required**. The legacy Tk entry point is not a safe substitute: it is not the default path and lacks the required source-mode backend gate and full package safety checks.

## Open Questions

- [NEEDS CLARIFICATION: none blocking source remediation; real Windows testing must confirm that the packaged staged EXE can run its `--apply-update` mode, wait for the original EXE to release file handles, replace the onedir tree without elevation, restart, and roll back if the new EXE cannot launch.]
- [NEEDS CLARIFICATION: Build 017/018 production ZIP bytes are not present in this checkout; the bootstrap conclusion is conservative and is based on the checked-in build/runtime contract and explicit unconditional lock, not forensic inspection of historical binaries.]
