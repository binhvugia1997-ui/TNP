# React + pywebview development integration

This is a local development launcher, not a Portable/EXE build. `app.desktop` opens the built React files from
`frontend/dist/index.html` in pywebview. pywebview's loopback HTTP server is rooted at `frontend/dist` and serves
only the built UI assets; it does not expose report, configuration, learning, or log directories. The relative local URL
also lets Edge/WebView2 load Vite's ES-module assets without `file://` CORS restrictions. Rebuild after UI changes, then
relaunch the desktop window.

The integration keeps Python authoritative for scanning, extraction, Excel, learning, Ollama, configuration, logs, and
update state. The development updater's **install** action is deliberately disabled; the launcher must not update or
replace this checkout.

## Exact Windows PowerShell commands

Run from the repository root (replace the path with the local checkout path):

```powershell
Set-Location 'C:\path\to\TNP'

py -3.14 -m venv .venv-webview
& .\.venv-webview\Scripts\python.exe -m pip install --upgrade pip
& .\.venv-webview\Scripts\python.exe -m pip install -r requirements-webview.txt
& .\.venv-webview\Scripts\python.exe -m pip install pytest pyflakes

Push-Location frontend
npm ci --ignore-scripts
npm run build
Pop-Location

& .\.venv-webview\Scripts\python.exe -m app.desktop --debug
```

The target desktop has already verified Python 3.14.7, Node.js 24.21.0, pywebview 6.2.1, pythonnet 3.0.5, native
window startup, HTML/JavaScript rendering, JavaScript-to-Python `ping()`, and the returned `Python bridge OK` value.
The launcher selects pywebview's Edge Chromium backend on Windows. A WebView2 Runtime is required.

After changes, repeat the frontend build and relaunch. To run automated checks from the repository root:

```powershell
& .\.venv-webview\Scripts\python.exe -m pytest -q
& .\.venv-webview\Scripts\python.exe -m compileall -q app tests tools
& .\.venv-webview\Scripts\python.exe -m pyflakes `
  app/application_service.py app/bridge.py app/desktop.py `
  app/batch_processor.py app/excel_writer.py app/gui_controller.py app/logger.py `
  tests/test_application_service.py
Push-Location frontend
npm run typecheck
npm run build
Pop-Location
git diff --check
```

Do **not** run `BUILD_AND_PUBLISH.bat`, build Portable/EXE, publish an update/release ZIP, or edit production
`version.json` for development acceptance.

## Acceptance checklist and status vocabulary

| Status | Scope |
|---|---|
| **REAL CONNECTED — AUTOMATED TESTED** | The React bridge contract and JSON-safe DTOs; real Python report scanning/list/details; configuration and status mapping; manual Vendor/date persistence and Excel lock/retry; a production `GuiController` + `BatchProcessor` fixture run; worker progress/logs and stop-after-current lifecycle; image/content candidate review, label persistence, Excel reapply/retry, model-update dispatch, export, and fixed output actions. The synthetic worker test covers lifecycle only; it is not the processing integration test. |
| **VERIFIED ON WINDOWS** | Target-machine checks listed above: native window startup, HTML/JavaScript rendering, JS-to-Python bridge invocation, and Python-to-JS `Python bridge OK` response. |
| **IMPLEMENTED — WINDOWS ACCEPTANCE REQUIRED** | Actual user-selected native dialogs; the complete React page inside Edge WebView2 (not just the bridge ping); local real-world PPTX processing and rendered image evidence; Excel lock behavior under Microsoft Excel; output/open-folder actions; and 100%, 125%, and 150% DPI/window-size behavior. Linux automated tests do not substitute for these checks. |
| **DISABLED WITH REASON** | Applying/installing updates is disabled in the development UI/service so no updater replaces or restarts the checkout. Update checking remains a read-only backend operation. |
| **NOT IMPLEMENTED** | Hot-module-reload Vite server embedded directly in the pywebview window. The supported development flow is `npm run build`, then `python -m app.desktop`; opening Vite in a normal browser has no mock/demo bridge fallback. |

### Windows manual acceptance

1. Launch with the commands above and confirm version **1.3.2 — Build 015** and the connected Python state.
2. Exercise folder selection, one-PPTX selection, template/output selection, cancellation, and the report scan. Confirm the
   report table shows filenames and no absolute source paths.
3. Open a report, edit Vendor and Occurrence date, reject an invalid date, save valid values, and confirm the other Excel
   cells remain unchanged. With the result workbook open in Excel, verify the locked message and retry after closing it.
4. Process a disposable copy of a representative PPTX with Ollama disabled. Observe real progress and logs; request stop
   during a file and confirm that file completes while the next file is not started. Check extracted details, rendered
   QPN/After evidence, Excel output, and the output-file/folder actions.
5. Exercise Ollama connection/model refresh and read-only update check. Confirm the update-install button stays disabled.
6. Review image and content candidates, save labels/notes, verify local persistence, test saved-label Excel retry while the
   workbook is locked, and exercise model update/export.
7. Repeat layout and native-dialog checks at Windows display scaling 100%, 125%, and 150%; verify the minimum window size.

Do not use production reports or the production output workbook for the manual run. Keep the run local; no LAN update
publication or release package is part of this acceptance.
