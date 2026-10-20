# Bug Verification: PROMPT-030B — safe integration on accepted 9faba26

- **Slug**: `prompt-030b` (continued from assessment/fix context)
- **Tested**: 2026-10-16 (UTC)
- **Assessment**: ./assessment.md
- **Fix**: ./fix.md
- **Result**: **partial — all available Linux checks pass; real Windows acceptance NOT RUN**
- **Branch**: `arena/1ec5e900-tnp`, based exactly on `9faba26486a77d2637c5fd4216537df758e1eb01`
- **Review revision**: the containing implementation/report commit; obtain its exact hash with `git log -1` on this branch (also supplied in the PR/final report).

Executed from checked-in Spec Kit skill instructions; native Arena skill invocation was not available.
Assessment and fix reports reviewed before verification. `.specify/extensions.yml` still has `hooks: {}`; no mandatory/optional pre/post hooks are registered. Verification did not edit runtime code or existing tests.

## Summary / audit verdict

**Do not merge PR #14 directly.** Its source `arena/f4a4d15e-tnp` / `aee4580` has merge-base **`588466e`** with accepted integration **`9faba26`**. Only `aee4580` is exclusive vs accepted; the eight earlier commits GitHub lists against old main are already ancestors. Exclusive patch: 11 files, +2585/-137. PR-tree vs accepted-tree: 133 files, +3902/-24448 (mostly missing newer accepted work, not an import plan). Four Git conflicts and eight semantic conflict categories are documented in assessment.md.

The missing behaviors were **reimplemented selectively**, not blindly merged/cherry-picked: durable candidate/image/full-slide review, source-scoped evidence identity/cache, top-of-right-panel labels, scoped multi-select/bulk and guarded quick confirmation. New tests demonstrate saved and unlabeled candidates survive restart/offline; wrong cross-report cached pixels no longer reproduce; real Python-rendered production geometry remains matched to its bitmap. Full current tests remain intact and green.

The accepted renderer, ownership proof, source fidelity/region builder, 036S fast writes/stale-poll protection, 038 viewport/freshness and 039 updater/Portable contracts remain authoritative. No `app/qpn_renderer.py`, `app/image_learning.py`, `app/updater.py`, `app/desktop.py`, existing tests, Portable validation/build tools or stable-backup changes.

## Environment / real-host observations

Linux, Python **3.11.2**, Node **22.22.3**, npm **10.9.8**. Development requirements installed into ignored `.venv`; frontend `npm ci` used `PUPPETEER_SKIP_DOWNLOAD=1`. The existing runtime dependencies pywebview **6.2.1**, pythonnet **3.0.5**, clr-loader **0.2.10** were installed for complete import/packaging-probe coverage; PyMuPDF **1.28.2** is available. No Windows desktop, Office installation, .NET Framework host or Windows Portable runtime is available here.

Actual unmodified host probe:

```text
sys.platform = linux
powerpoint_probe() = (False, 'availability', 'not-windows')
renderer_status() = {'powerpoint': 'Unavailable', 'libreoffice': 'Unavailable', 'builtin': 'OK'}
```

COM simulations in unit tests are explicitly test seams, never a claim of real PowerPoint availability/ownership success. Real PPTX end-to-end tests use the honest built-in fallback on this host. Frontend 034 report-list measurements used real Linux Chromium; Learning interaction tests use jsdom/StrictMode and stubbed geometry. Neither is Windows/WebView2 acceptance.

## Checks performed

| Check | Command / action | Result / evidence |
| --- | --- | --- |
| Baseline reproduction | Temporary production service + save label + new service, two sources with same canonical ID/different pixels | Before fix: 1 label retained, 0 restored rows, no previews directory; cross-source cached pixels identical. Exact evidence in assessment. |
| Original-case rerun + new regression coverage | `.venv/bin/python -m pytest -q tests/test_prompt030b_integration.py` | **23 passed**. Real production PPTX → candidates → full Python slide → snapshot → fresh service (including unlabeled) → confirm → updated render → offline restart → relabel; mandatory pixels, dimensions, regions, item index 0, note/ID and path-free payload assertions. |
| Broad targeted Python | Command below | **289 passed in 31.46s**, exit 0. Covers renderer/ownership, Learning preview/state/persistence/cache, scoped identity, fast writes/bridge lifecycle, 038 freshness and updater preservation. |
| Final focused rerun after fix report | Command below | **56 passed in 7.63s**, exit 0. Includes exact two PROMPT-039 failure nodes unchanged. |
| Targeted frontend | `npm --prefix frontend run test:prompt030b` + `test:learning-labels` + `test:learning-hooks` + `test:prompt038` | **7 + 1 + 1 + 3 passed**, zero failures/skips. Native click/click/dblclick sequence, no guessed label, one save in StrictMode, cross-report same-basename/same-slide boundaries, optimistic bulk, deferred polling, visible unsaved failure/retry, note preservation, identity pruning, hooks/viewport. |
| Full available Python | `.venv/bin/python -m pytest -q -ra` | **1313 passed, 1 skipped in 116.08s**, exit 0. No failed tests. |
| Only final Python skip | `tests/test_prompt030_runtime.py:1239` | `Windows-only: exercises real CFFI/CLR DLL loading; Linux cannot claim native loadability`. No skip/assertion modification. |
| Full frontend | `npm --prefix frontend run test:frontend` | **79 passed, 0 failed/cancelled/skipped**, exit 0; 62.10s final run. Existing 034/037/038 tests all included. |
| Typecheck | `npm --prefix frontend run typecheck` | PASS (`tsc --noEmit`), exit 0. |
| Frontend production asset build | `npm --prefix frontend run build` | PASS, 1588 modules; local ignored dist only, no Portable/EXE/release build. |
| Desktop freshness | `from app.desktop import frontend_dist_freshness; print(frontend_dist_freshness())` | `stale: False`, `reason: fresh`; existing 038 guard unchanged. |
| Scoped static analysis | `.venv/bin/python -m pyflakes app/learning_assets.py app/application_service.py app/gui_controller.py app/bridge.py tests/test_prompt030b_integration.py` | PASS, exit 0. |
| Environment integrity | `.venv/bin/python -m pip check` | `No broken requirements found.`, exit 0. |
| Whitespace / safety diff | `git diff --check`; protected-path diff against `9faba26` | PASS; protected-path output empty, explicit changed-file inventory reviewed. |
| Real Windows COM / desktop / machine restart / installed Portable acceptance | Not executable on Linux | **NOT RUN**, not simulated as acceptance. Exact retest/checklist in docs/WINDOWS_ACCEPTANCE_PROMPT030B.md. |

### Reproducible broad targeted command

```sh
.venv/bin/python -m pytest -q tests/test_prompt030b_integration.py tests/test_prompt035_powerpoint_isolation.py tests/test_prompt027_rendering.py tests/test_prompt027r_renderer.py tests/test_prompt027r_regions.py tests/test_image_learning.py tests/test_prompt020_report_scoping.py tests/test_prompt032_learning_state.py tests/test_prompt036r_bridge_hang.py tests/test_prompt036s_learning_latency.py tests/test_prompt038_desktop_freshness.py tests/test_application_service.py tests/test_updater.py
```

### Final focused rerun

```sh
.venv/bin/python -m pytest -q tests/test_prompt030b_integration.py tests/test_prompt035_powerpoint_isolation.py tests/test_prompt038_desktop_freshness.py tests/test_image_learning.py::test_28_apply_update_keeps_learning_data_and_rollback_restores_it tests/test_updater.py::test_updater_waits_for_main_process_exit
```

## Important test details / earlier failures

- First full run: **1299 passed, 13 skipped, 2 failed**. The failures were existing live import/collection probes with webview/pythonnet missing in the new venv. Installed existing dependencies in the environment and reran the unchanged tests; final suite is green. No dependency/lock metadata changed to obtain this result.
- Fix-phase compatibility regression: an initial optional-storage wiring incorrectly assumed preview-only callers had a controller, caught by eight existing 027/027R tests. Storage became optional; tests stayed unchanged and passed. Initial new updater test incorrectly called strict `release_entries` on an installed app (which correctly rejects non-package roots); corrected the test to inspect the validated staged source. No production validation was weakened.
- Nonfatal Vite teardown `The build was canceled` appeared on some runs, including baseline; node reported zero failed/cancelled tests and exited 0. Final focused frontend reruns were clean. This warning was not hidden by a pipe or exit-code masking.
- New updater parameterized tests execute **real `apply_update()` and rollback** on the existing complete validated Portable fixtures, with accepted short/root-level extraction layouts. They compare **all learning_data file bytes** before/after, including candidate snapshots, full-slide metadata, rendered image bytes and JSONL labels. Only process liveness/spawn callbacks are deterministic adapters; no actual installed app is replaced/restarted.
- Bulk validates the entire request before mutation. A simulated second JSONL append failure leaves the first append committed, returns an error and retains pending choices for retry; this is **reported partial persistence**, not fake all-or-none atomicity. Existing Excel transaction/lock tests remain green.

## Residual risks / Windows acceptance

1. Real Windows full suite still needs rerunning on the review commit, especially native CLR loading and both exact PROMPT-039 nodes. Prior base acceptance does not automatically accept this delta.
2. Real Office single-instance behavior, user-unsaved presentation isolation, faithful export/prior black-background case, and owned-only cleanup need the manual checklist. No process may be killed to manufacture a PASS.
3. Real WebView2 click latency, Ctrl/Shift selection, actual double-click, DPI/resize/400% zoom, source-dist freshness, machine restart and offline source need Windows desktop observation.
4. Shared classification labels intentionally remain shared; source-scoped review tokens separate visual assets, not the training-label schema. Unknown legacy provenance is not permission to guess a region or path. Existing unaccepted PR #14 sidecars are preserved but not treated as authoritative render truth.
5. A changed source requires new processing for current candidate geometry. Offline label/model changes retain saved pixels but remove outdated production-region overlays. Missing/corrupt display assets do not remove labels. Content-addressed image/history files are not garbage-collected automatically; no user assets are deleted by this change.
6. Actual installed Portable update/process restart acceptance requires a separately authorized disposable Windows run; this task does not authorize publishing or a new Portable release. Automated strict updater fixture coverage is not that acceptance.

## Final safety / recommendation

PR #14 remains **OPEN**, head `aee4580`, target main unchanged; it was not merged/closed/edited. The new task is based on accepted integration and its review PR targets that integration, **not main**. No integration/main merge, LAN publish, BUILD_AND_PUBLISH, release/Portable creation, version/build bump (still **1.3.7 / 020**), stable-backup modification, updater weakening, existing test weakening, `git add .`, unrelated/user-file staging or branch switching. Only reviewed explicit paths are committed; venv/node_modules/dist/learning_data/test outputs remain ignored.

**Recommendation:** ready for code review of the selective integration; retain a **partial** cross-platform verdict until the documented real Windows checks are performed. Review the eleven-file keep/drop matrix in assessment.md and the fourteen-file actual delta in fix.md. Do not blind-merge PR #14 or treat Linux evidence as Windows acceptance.

**PROMPT-030B READY FOR REVIEW**
