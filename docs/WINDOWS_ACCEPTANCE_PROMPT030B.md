# PROMPT-030B — Windows retest (NOT executed in the Linux agent)

Accepted base: `integration/prompts-034-035-036-037` / `9faba26` (including the user's PROMPT-039 Windows acceptance).
Review branch: **`arena/1ec5e900-tnp`**, not the old PR #14 branch. Check `git log -1` against the review commit/PR before testing. Do not infer a new Windows acceptance from the old accepted base or Linux tests.

Use PowerShell in the real checkout (example path below). Use an existing Python 3.12 Windows dev environment with PowerPoint, pywin32 and WebView2 installed. The commands below do **not** build/publish a Portable or modify version/build. Never modify Office registration, force `sys.platform`, mock availability in a manual acceptance, or terminate PowerPoint processes.

## 1. Confirm revision and development prerequisites

```powershell
Set-Location 'D:\arena\TNP\TNP'
if ((git branch --show-current).Trim() -ne 'arena/1ec5e900-tnp') { throw 'Wrong acceptance branch; stop.' }
git status --short
git log -1 --format='%H %s'
git merge-base HEAD 9faba26486a77d2637c5fd4216537df758e1eb01
# Must print 9faba26486a77d2637c5fd4216537df758e1eb01. Do not overwrite a dirty checkout.

$py = '.\.venv-webview\Scripts\python.exe'
# If using the existing .venv instead, set $py = '.\.venv\Scripts\python.exe'.
& $py --version
& $py -m pip install -r requirements-dev.txt -r requirements-webview.txt
& $py -m pip check
npm.cmd --prefix frontend ci --no-audit --no-fund
& $py -c 'import sys; from app import __version__, BUILD_ID; from app.qpn_renderer import powerpoint_probe, renderer_status; print(sys.platform, __version__, BUILD_ID); print(powerpoint_probe()); print(renderer_status())'
```

Expected version/build is unchanged: **1.3.7 / 020**. Real `sys.platform` is `win32`. `powerpoint_probe()` retains the accepted **tuple** API `(available, stage, detail)`; it does not have PR #14's `.available` property or `renderer_status_detailed()`. A positive probe permits an attempt; it does NOT prove successful export or safe ownership. A user-owned session must still cause a truthful `ownership` refusal.

## 2. Targeted tests (exact commands)

```powershell
& $py -m pytest -q tests/test_prompt030b_integration.py tests/test_prompt035_powerpoint_isolation.py tests/test_prompt027_rendering.py tests/test_prompt027r_renderer.py tests/test_prompt027r_regions.py tests/test_image_learning.py tests/test_prompt020_report_scoping.py tests/test_prompt032_learning_state.py tests/test_prompt036r_bridge_hang.py tests/test_prompt036s_learning_latency.py tests/test_prompt038_desktop_freshness.py tests/test_application_service.py tests/test_updater.py
if ($LASTEXITCODE -ne 0) { throw 'Targeted Python retest failed.' }

# Exact PROMPT-039 regression cases must still reach apply/rollback and wait/timeout assertions:
& $py -m pytest -q tests/test_image_learning.py::test_28_apply_update_keeps_learning_data_and_rollback_restores_it tests/test_updater.py::test_updater_waits_for_main_process_exit
if ($LASTEXITCODE -ne 0) { throw 'PROMPT-039 retest failed.' }

npm.cmd --prefix frontend run test:prompt030b
npm.cmd --prefix frontend run test:learning-labels
npm.cmd --prefix frontend run test:learning-hooks
npm.cmd --prefix frontend run test:prompt038
```

These COM tests use controlled fakes even on Windows. They establish regression behavior, **not** real Office ownership/fidelity acceptance. New updater tests use the existing strictly validated Portable fixtures and a non-spawning helper adapter; they really execute apply and rollback on disposable directories, not the user's installation.

## 3. Full available suite and fresh frontend

```powershell
& $py -m pytest -q -ra
if ($LASTEXITCODE -ne 0) { throw 'Full Python suite failed.' }
npm.cmd --prefix frontend run test:frontend
if ($LASTEXITCODE -ne 0) { throw 'Frontend suite failed.' }
npm.cmd --prefix frontend run typecheck
if ($LASTEXITCODE -ne 0) { throw 'Typecheck failed.' }
npm.cmd --prefix frontend run build
if ($LASTEXITCODE -ne 0) { throw 'Frontend build failed.' }
& $py -m pyflakes app/learning_assets.py app/application_service.py app/gui_controller.py app/bridge.py tests/test_prompt030b_integration.py
git diff --check
& $py -c 'from app.desktop import frontend_dist_freshness; print(frontend_dist_freshness())'
```

Record counts, every skip reason and exact commit. The existing 034 browser harness uses `@sparticuz/chromium`; if unavailable on Windows, it reports real-browser measurement skips. Do **not** turn those into a PASS: manually measure the six-row viewport, sticky header and horizontal overflow in the real desktop app, or separately arrange an appropriate browser test environment. No existing tests were weakened or skipped in this integration.

`frontend/dist` is ignored/local and must be rebuilt after source changes (038 stale-dist guard is intentionally unchanged). Do not commit it. Launch the React desktop directly so a missing WebView runtime cannot silently substitute the legacy Tk UI:

```powershell
& $py -m app.desktop --debug
```

## 4. Real PowerPoint ownership / export / cleanup

Use copies of source reports and a disposable output workbook. Do not use production output for an acceptance experiment.

1. **User instance present:** Open PowerPoint yourself with a user presentation (include an unsaved edit). Record its PID and window. Run report processing / Learning preview. It must refuse activation with `stage=ownership`, use an honestly marked fallback, and leave all user slides, edit state, windows and processes untouched. A positive availability probe must not be mislabeled as missing Office.
2. **No user instance:** Save/close your own PowerPoint normally through its UI. Run a new preview (allow the 15-second degraded TTL to expire, or restart after expiry). With working Office registration, the actual selected backend must be `powerpoint`, faithful badge true, no fallback warning. Compare original and preview, including authored Before/After captions, arrows, shapes, fonts and the prior black-background case.
3. **Owned cleanup:** On success and a real error/cancellation path, only the app-created presentation may close; only a proven exclusively app-owned instance may quit. Never use `Stop-Process`, `taskkill` or a process terminator as a test cleanup shortcut. Any ownership ambiguity must fail closed; a refused/unproven process is not permission to kill it.
4. **Observe, don't alter:**

```powershell
Get-Process POWERPNT -ErrorAction SilentlyContinue | Select-Object ProcessName, Id, StartTime
# After each scenario / normal app exit:
Get-Process POWERPNT -ErrorAction SilentlyContinue | Select-Object ProcessName, Id, StartTime
Get-ChildItem .\logs -Filter '*.log' -ErrorAction SilentlyContinue |
  Select-String -Pattern 'SLIDE_RENDER_BACKEND_FAILED|SLIDE_RENDER_CLEANUP|SLIDE_RENDER .*backend=powerpoint'
```

Record before/after PIDs, selected backend, failure stage/detail and whether the user's presentation remains unmodified. Do not attach private report paths/content to a public PR.

## 5. Real Learning UI, restart and offline assets

- On a multi-item slide, labels must be first in the right panel. Picture overlay and production evidence-region overlay remain distinct and match their own item; Python full-slide preview retains its authored geometry (no new masking/paint changes).
- Single click focuses only. Ctrl/Meta + click and Shift range can select multiple candidates of the **same source report and slide**, including separate ImprovementItems. Same basename and same slide number from another directory must not join the selection.
- Labeling multiple selected images updates immediately and commits in one batch save. Each note stays with its own candidate. A failure must remain visible/unsaved and retryable, not claim all-or-none or silent success. Excel lock/retry must preserve saved labels and must not create partial workbook output.
- Double-click without an explicit/current confirmed label asks for one, never guesses. With a label, a real double-click sequence confirms once; with a multi-selection it confirms the selected same-scope candidates once. No duplicate save, hook error or unexpected auto-navigation.
- Fresh app/no candidate → process report → candidates arrive asynchronously: 100% preview must fit the real viewport, not ~80px. Test candidate switches, 400% zoom, resize and panel drag. Report list 034/037 behavior must remain correct.
- Enter review for both unlabeled and confirmed candidates; close the app normally, reopen without processing. Candidate image, full slide, IDs, item index **0**, label and note must return from `learning_data/previews` / existing label JSONL. Repeat across machine restart.
- For an **acceptance copy only**, make its source PPTX unavailable, reopen and relabel. Stored Python-rendered images remain available. When saved labels/model identity changes offline, the old production-region overlay must clear rather than pretend it is current. Restoring/changing source contents requires a new processing run for authoritative candidate geometry; changed-source snapshots must not overlay old boxes on new pixels.
- Missing/corrupt preview must keep label rows and counts. Never delete the real learning dataset to test corruption; use a disposable copy. No absolute path may appear in Learning DTOs/asset URLs; source paths in private snapshot files are local provenance, not browser payloads.

```powershell
Get-ChildItem .\learning_data\previews -File | Select-Object Name, Length, LastWriteTime
Get-ChildItem .\learning_data\previews -Filter '*.tmp'  # empty after completed writes
```

Opaque `reviewId`/`reportScope` separate report evidence. Canonical `candidate_id`/JSONL label dedup remains intentionally shared, as in the accepted classifier; do not reinterpret evidence-row count as a change to training-record identity.

## 6. Portable updater acceptance boundary

Automated apply/rollback and exact PROMPT-039 cases above must pass on Windows, including long-path-sensitive fixture setup. Production updater code, Portable validation and stable backup are unchanged.

A real packaged-runtime/installed Portable update is **not executed or newly authorized by this task**. If a reviewer separately authorizes that acceptance using an already prepared disposable Portable, compare hashes of all `learning_data` files (labels, models, old thumbnails, new snapshot JSON and rendered image files) before/after update and rollback, plus real old-PID wait/restart behavior. Do not publish to LAN, run `BUILD_AND_PUBLISH.bat`, create/bump a release, or modify the immutable stable backup as part of these retest commands.

**Required evidence before declaring new Windows acceptance:** exact revision, full/targeted logs and skip reasons, real COM ownership/fidelity/cleanup observations, real desktop selection/bulk/zoom/freshness results, app/machine restart/offline results, and any separately authorized packaged-runtime acceptance results. All remain NOT RUN in the Linux agent.
