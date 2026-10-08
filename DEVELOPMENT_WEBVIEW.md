# React + pywebview development integration

This is a local development launcher, not a Portable/EXE build. `app.desktop` opens the built React files from
`frontend/dist/index.html` in pywebview. pywebview's loopback HTTP server is rooted at `frontend/dist` and serves
only the built UI assets; it does not expose report, configuration, learning, or log directories. The relative local URL
also lets Edge/WebView2 load Vite's ES-module assets without `file://` CORS restrictions. Rebuild after UI changes, then
relaunch the desktop window.

The integration keeps Python authoritative for scanning, extraction, Excel, learning, Ollama, configuration, logs, and
update state. The development updater's **install** action is deliberately disabled; the launcher must not update or
replace this checkout.

## Python compatibility

Use **Python 3.12** for the Windows pywebview development environment. `requirements-webview.txt` pins pywebview 6.2.1
and, on Windows, `pythonnet==3.0.5`; pip's `Requires-Python` metadata rejects the previously attempted Python 3.14
combination. Install the declared requirements normally—do not use pip flags that bypass package metadata. Python 3.12
has launched the integrated app on Windows.

This is a narrower compatibility choice for the **Windows pywebview development frontend**, not a change to the backend
or the ordinary source/Portable workflow. The backend/setup range remains Python 3.10 through 3.14 (`>=3.10,<3.15`),
so Python 3.11 CI remains supported. No extra Python-version preflight is implemented; normal pip metadata is the
compatibility gate for the pywebview environment.

## Exact Windows PowerShell commands

Run from the repository root (replace the path with the local checkout path). If `.venv-webview` was created with a
different interpreter, first confirm it contains no needed local data, then recreate it with Python 3.12. Do not try to
repair an incompatible environment by bypassing pip's metadata checks.

```powershell
Set-Location 'C:\path\to\TNP'

py -3.12 -m venv .venv-webview
& .\.venv-webview\Scripts\python.exe -m pip install --upgrade pip
& .\.venv-webview\Scripts\python.exe -m pip install -r requirements.txt
& .\.venv-webview\Scripts\python.exe -m pip install -r requirements-webview.txt
& .\.venv-webview\Scripts\python.exe -m pip install pytest pyflakes

Push-Location frontend
npm.cmd ci --ignore-scripts
npm.cmd run typecheck
npm.cmd run build
Pop-Location

& .\.venv-webview\Scripts\python.exe -m app.desktop --debug
```

The Windows launcher selects pywebview's Edge Chromium backend; a WebView2 Runtime is required. The prior, pre-PROMPT-023
Windows run verified startup, Edge WebView2 and production React rendering, the JavaScript-to-Python bridge and returned
Python response, native folder/PPTX pickers, scanning a real folder of six PPTX reports and one selected real PPTX,
report DTO-to-React table rendering, Management Number and occurrence-date display, and selected-report details. Those
checks were performed before the PROMPT-023 fixes; they are not post-patch verification.

## Automated checks

Run from the repository root after installing the development requirements:

```powershell
& .\.venv-webview\Scripts\python.exe -m pytest -q tests/test_application_service.py tests/test_runtime_paths.py tests/test_updater.py
& .\.venv-webview\Scripts\python.exe -m pytest -q
& .\.venv-webview\Scripts\python.exe -m compileall -q app tests tools
& .\.venv-webview\Scripts\python.exe -m pyflakes `
  app/application_service.py app/bridge.py app/desktop.py app/runtime_paths.py `
  app/batch_processor.py app/excel_writer.py app/gui_controller.py app/logger.py `
  tests/conftest.py tests/test_application_service.py tests/test_runtime_paths.py tests/test_updater.py
Push-Location frontend
npm.cmd run typecheck
npm.cmd run build
Pop-Location
git diff --check
```

Do **not** run `BUILD_AND_PUBLISH.bat`, build Portable/EXE, publish an update/release ZIP, or edit production
`version.json` for development acceptance.

## Acceptance status

| Status | Scope |
|---|---|
| **AUTOMATED TESTED — PROMPT-023** | 875 Python tests passed (82 targeted picker/runtime-isolation/updater tests); compileall and configured pyflakes passed; frontend typecheck and production build passed. The build emits a hashed local favicon referenced through `./assets/...`; `git diff --check` passed. The suite covers the bridge filter/cancel/path contracts, isolated persistent state, and prior extraction, evidence, learning, Excel safety/lock-retry, force-reprocessing, Ollama, and updater regressions. |
| **AUTOMATED TESTED — PROMPT-025** | 894 Python tests passed (875 baseline + 19 new in `tests/test_prompt025_multi_item.py`); compileall, pyflakes (new/changed files), and `git diff --check` passed; frontend typecheck and production build passed. Multi-improvement segmentation + item-scoped After regions are covered by the §30 two-item fixture (3-picture After block = one region, item #1 crop excludes item #2 text, whitespace policy, caption/arrow association, Before exclusion), heading-style variants (defect-style "Lỗi … sau ép nhỰA", mid-block non-bold colon heading), 3-item and 1-item slides, report isolation, force-reprocess, learning identity per report+slide+item, render-once-per-slide with two item crops, the cooperative cancellation hook, file-atomic stop-after-current with a multi-item slide, Excel mapping (one row, both items in order, two image groups), temporary-action exclusion, and the `cải進` caption fold. All prior PROMPT-014/015/020/021/023/024 regressions still pass. |
| **AUTOMATED TESTED — PROMPT-027** | 965 Python tests passed (911 baseline + 54 new in `tests/test_prompt027_rendering.py`); compileall, configured pyflakes and `git diff --check` passed; frontend `test:bridge` (14), `test:learning-hooks` (1) and the new `test:prompt027` (7) passed; typecheck and production build passed. Covered headlessly: `No Fill` / `No Line` parsed as *paints nothing* (`<a:noFill/>` → `MSO_FILL.BACKGROUND`, which is **not** `None`) and no black frame in the built-in renderer, including grouped children; full-slide render keeps the authored aspect ratio, whitespace and object pixel position with no white-trim; the built-in renderer honours the authored vertical anchor / text insets / line spacing / `normAutofit` scale (the cause of the vertically shifted preview); EMU→pixel crop conversion against the ACTUAL rendered dimensions; item-scoped region ownership with no cross-item contamination (pixel-level); authored whitespace preserved between two After pictures; bounded deterministic region padding; `SLIDE_RENDER` / `SLIDE_RENDER_BACKEND_FAILED` diagnostics; preview-cache identity now includes renderer chain + width + render schema version with report isolation and mtime invalidation; the learning DTO exposes the preview backend, faithfulness and the authoritative evidence-region bbox; overlay alignment for BOTH geometries under fit/resize/letterbox/zoom/DPI; frontend/backend version agreement at 1.3.3 / Build 016. |
| **WINDOWS RETEST REQUIRED AFTER PROMPT-027** | First confirm the header shows **v1.3.3 · Build 016** and Settings shows **1.3.3 — Build 016**; if Windows still shows 1.3.2 / Build 015, STOP — that is stale code or a stale `frontend/dist`. Then, on the real PPTX slide with 2 improvements, authored whitespace, an invisible `No Fill` / `No Line` content frame and multiple After pictures, verify: the full-slide preview matches PowerPoint with content no longer shifted upward; authored whitespace preserved; the invisible frame stays invisible with no fake black rectangle in either the preview or the final Excel evidence; the preview reports its renderer backend and logs `SLIDE_RENDER purpose=learning_preview backend=powerpoint` and `purpose=after_evidence backend=powerpoint`; the picture-candidate overlay (dashed) and the final evidence-region overlay (solid, "Mục #n · Vùng xuất Excel") are both aligned and visually distinct; Mục #1 and Mục #2 keep independent regions when switching candidates; the exported Excel crop keeps the intended whitespace and contains no neighbouring item text. No Linux check substitutes for Windows acceptance of the real file, and PowerPoint-faithful rendering cannot be claimed until it is seen there. |
| **WINDOWS RETEST REQUIRED AFTER PROMPT-025** | The real two-item PPTX (item #1 "Lỗi xước rear …" with a three-picture After group + "Sau cải進" caption, item #2 "Cải tiến lỗi lệch ATN …"): confirm exactly 2 improvement items, 2 item-scoped After crops in the Excel image cell (one horizontal image row per item), item #1's crop free of item #2 heading/body, the transition arrow excluded, learning preview showing the "Mục #n" chip per candidate, and stop-after-current/lock-retry behaviour unchanged. No Linux check substitutes for Windows acceptance of the real file. |
| **WINDOWS VERIFIED BEFORE PROMPT-023** | The pre-patch checks listed above. They confirm the earlier integrated runtime, not the picker-filter or persistent-state fixes in this change. |
| **WINDOWS RETEST REQUIRED AFTER PROMPT-023** | Native picker open/cancel behavior and all post-patch real-report processing, Excel edit/lock/retry, stop-after-current, rendered evidence, and batch checks in the checklist below. No Linux GUI feasibility check substitutes for Windows acceptance. |
| **DISABLED WITH REASON** | Applying/installing updates is disabled in the development UI/service so no updater replaces or restarts the checkout. Update checking remains a read-only backend operation. |
| **NOT IMPLEMENTED** | Hot-module-reload Vite server embedded directly in the pywebview window. The supported development flow is `npm.cmd run build`, then `python -m app.desktop`; opening Vite in a normal browser has no mock/demo bridge fallback. |

## Post-patch Windows manual acceptance

Use disposable copies of reports, templates, and output workbooks. Keep the test local; do not publish an update or release.

1. Run the Python 3.12 setup sequence above, launch the app, and confirm version **1.3.3 — Build 016** and the connected
   Python state.
2. Exercise native folder, single-PPTX, XLSX template/open, XLSX output/save, learning-export, and update-folder pickers.
   Test both choosing a valid path and cancelling each applicable dialog. Confirm XLSX/PPTX/ZIP extensions and verify
   that cancelling leaves the current setting unchanged.
3. Scan a real report folder and select one representative real PPTX. Process that one report first; confirm real progress,
   QPN, structural cause, improvement text, the whole rendered “Sau cải tiến” crop, and the resulting Excel output.
4. Edit Vendor and Occurrence date for that report, save valid values, reject an invalid date, and verify the intended Excel
   cells changed without altering neighboring cells.
5. With two disposable reports, test **Dừng sau báo cáo hiện tại** and verify the active report finishes while the next
   report does not start. Open the result workbook in Excel, verify the lock/retry message, close Excel, retry, and confirm
   the saved values are applied.
6. Only after the single-report, manual-edit, stop, and lock/retry checks pass, run a small batch and confirm its report
   order, progress, QPN/cause/improvement details, rendered whole “Sau cải tiến” crop, and consolidated Excel output.
7. Exercise Ollama connection/model refresh and the read-only update check. Confirm update installation stays disabled.
8. Review image/content learning candidates, save labels/notes, test the saved-label Excel retry while the workbook is
   locked, and exercise the local learning-data export. Repeat layout and native-dialog checks at Windows display scaling
   100%, 125%, and 150%; verify the minimum window size.
