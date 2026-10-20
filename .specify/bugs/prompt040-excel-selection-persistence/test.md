# Bug Verification: the Excel picker is no longer needed on every run

- **Slug**: `prompt040-excel-selection-persistence`
- **Tested**: 2026-10-16 (UTC)
- **Assessment**: `./assessment.md` (unchanged)
- **Fix**: `./fix.md` (commit `6bf3a5207406c943e3f202202d9071b88e7001fd`, branch `arena/b092bbeb-tnp`)
- **Result**: partial — Linux verification passed, including the scripted reproduction; **Windows/manual acceptance not performed on this host** (see below)
- **Spec Kit execution**: Executed from checked-in Spec Kit skill instructions (`speckit-bug-test`); native Arena skill invocation was not available. Slug came from conversation context.
- **Extension hooks**: `.specify/extensions.yml` has `hooks: {}` — no `before_bug_test` / `after_bug_test` hooks are registered, so none executed.

## Summary

The reproduced defect (a successful Excel pick was returned to the browser but never written to the user config, so the
next start had empty Excel fields and forced another pick) no longer reproduces: the raw scripted reproduction now shows
both paths in `config.json` immediately after the picks, and a fresh process restores them without any dialog. All new and
existing suites pass. No regression was found in the folder/PPTX/update pickers, the PROMPT-039 updater preservation, or
PROMPT-030B learning storage. Windows acceptance is **not** claimed — this host is Linux, so the native dialog,
WebView2, PowerPoint COM and Portable update paths were not exercised.

## Checks Performed

| Check | Command / action | Result | Notes |
|-------|------------------|--------|-------|
| Reproduction (post-fix, scripted) | `.venv/bin/python /home/user/TNP_scratch/repro_prompt040.py` (scratch, not committed) | pass | `config after pick` holds both paths; `restart (no scan)` restores both. Pre-fix the same script printed `last_template='' last_output_file=''`. |
| Reproduction baseline (pre-fix, same host) | fresh worktree at `034d339` + the new test file → `.venv/bin/python -m pytest -q tests/test_prompt040_excel_persistence.py` | fail (as expected) | **18 failed, 3 passed** — the bug is proven by the tests, not assumed. |
| New tests (post-fix) | `.venv/bin/python -m pytest -q tests/test_prompt040_excel_persistence.py` | pass | **21 passed**. |
| Targeted regression (Python) | `pytest -q tests/test_prompt040_excel_persistence.py tests/test_application_service.py tests/test_gui_controller.py tests/test_updater.py tests/test_image_learning.py tests/test_prompt030b_integration.py tests/test_excel_lock_retry.py tests/test_excel_batch.py tests/test_prompt036r_bridge_hang.py tests/test_prompt038_desktop_freshness.py` | pass | **212 passed**. Includes PROMPT-039 updater, PROMPT-030B learning storage and the Excel lock/retry/atomic-commit suites. |
| Full available suite (Python) | `.venv/bin/python -m pytest -q` | pass | **1,332 passed, 3 skipped** (120.3 s). Skips: 2 × “git history unavailable” (shallow Arena checkout), 1 × Windows-only CLR loading test. |
| Frontend targeted (pre-fix baseline) | same Chromium run from the worktree at `034d339` → `node --test tests/prompt040-excel-default.test.mjs` | fail (as expected) | **0 passed, 4 failed**. |
| Frontend targeted (post-fix) | `npm run test:prompt040` | pass | **4 passed** (real Chromium via the checked-in `@sparticuz/chromium`, real `App` + mocked bridge only). |
| Frontend full suite | `npm run test:frontend` | pass | **83 passed, 0 failed**. |
| Frontend type-check | `npm run typecheck` | pass | `tsc --noEmit`, clean. |
| Frontend build + PROMPT-038 freshness | `npm run build`, then `app.desktop.frontend_dist_freshness()` | pass | `dist/index.html` + assets rebuilt; `{"stale": false, "reason": "fresh"}`. `frontend/dist` stays untracked (`.gitignore`). |
| Static checks | `.venv/bin/python -m pyflakes app/application_service.py app/bridge.py app/gui_controller.py tests/test_prompt040_excel_persistence.py`; `git diff --check` | pass | Clean. |
| Manual Windows steps | assessment’s 5-step manual reproduction (native pickers, close app, inspect `config.json`, restart) | not-run | Requires Windows + WebView2; cannot be performed on this Linux host. |

## Output Excerpts

```
# reproduction (post-fix)
picked template   : {'ok': True, 'data': {'cancelled': False, 'path': '…/picked/Bieu mau kiem chung 2026.xlsx', 'kind': 'template'}}
picked output     : {'ok': True, 'data': {'cancelled': False, 'path': '…/picked out/Kiem chung ket qua 2026.xlsx'}}
config after pick : last_template='…/picked/Bieu mau kiem chung 2026.xlsx' last_output_file='…/picked out/Kiem chung ket qua 2026.xlsx'
restart (no scan) : template='…/picked/Bieu mau kiem chung 2026.xlsx' output='…/picked out/Kiem chung ket qua 2026.xlsx'

# pre-fix baseline
18 failed, 3 passed in 1.81s

# pre-fix frontend baseline
1..4  # pass 0  # fail 4

# post-fix frontend targeted
1..4  # pass 4  # fail 0

# full suites
1332 passed, 3 skipped in 120.29s (0:02:00)
# tests 83  # pass 83  # fail 0

# freshness
{"stale": false, "reason": "fresh", "dist": "/home/user/TNP/frontend/dist", "newest_source": "/home/user/TNP/frontend/package.json"}
```

## Residual Risks

- **Windows not exercised here**: native dialog behaviour, WebView2, PowerPoint COM and the Portable update path are
  unverified on this host; only the Python/JS contract and the real-Chromium UI behaviour were verified.
- **Deep workbook check is asynchronous**: a very large check workbook briefly shows “Đang kiểm tra” before
  “Dùng được”. It runs once per `(path, mtime, size)` in a daemon thread and never per dashboard poll.
- **Read-only config location**: if `config/config.json` cannot be written (e.g. a Portable copy in a protected folder),
  the pick is still applied for the session and the UI states that the selection could not be saved (`persisted: false`)
  instead of pretending success. That is the pre-existing `save_settings()` contract, now reported honestly.
- **Pre-existing flake (unrelated)**: `PROMPT-034R · >6 rows: bounded viewport…` failed **1 of 4 runs on the untouched
  base commit** with the same 30 px workspace-height difference. Not introduced by this change.

## Windows retest commands

```bat
:: 1. Fresh copy of the Portable build with this branch, empty Excel defaults
del /q config\config.json

:: 2. Start the app, pick "File kiểm chứng Excel" and "File kết quả" with the native pickers
run.bat            (or the built ReportExtractor.exe)

:: 3. Close the app WITHOUT scanning and without Cài đặt → Lưu cấu hình, then inspect the config:
type config\config.json
::    last_template and last_output_file must hold the two picked paths

:: 4. Start again: both Excel rows must show the remembered name + path + "Dùng được",
::    with no picker opening on startup; "Quét" then "Bắt đầu xử lý" must reuse them
::    (press "Đổi file…" only when a different workbook is wanted).

:: 5. Delete/move the stored check workbook and start again: the row must show "Không tìm thấy"
::    with the Vietnamese explanation, processing must refuse with the existing validation message,
::    and only "Đổi file…" may choose a new file.

:: 6. Press "Xóa lựa chọn" on one row, restart: only that row must be unconfigured.

:: 7. Run the automated suites on Windows:
RUN_TESTS.bat
cd frontend && npm ci && npm run test:frontend && npm run typecheck && npm run build
```

## Recommendation

Hold the bug open until Windows acceptance is done: the defect is proven fixed by the scripted reproduction and by the
new/regression suites on Linux, but the native-dialog, WebView2 and Portable paths must be re-run on Windows using the
steps above. No further code change is expected.
