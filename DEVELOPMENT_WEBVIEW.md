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

Do **not** run `BUILD_AND_PUBLISH.bat`, publish an update/release ZIP, or edit production `version.json` for
development acceptance.

## Windows Portable acceptance build (no publishing)

`ReportExtractor.exe` launches the same React application as `python -m app.desktop`:
`run.py` → `app.main.main()` → `launch_ui()` → `app.desktop.main()` (pywebview + EdgeChromium/WebView2).
Build an acceptance package **without** publishing anything:

```powershell
cd frontend
npm.cmd install ; npm.cmd run build      # the exe serves frontend/dist; a stale bundle = a blank window
cd ..
build_portable.bat --no-publish
```

`tools/build_portable.py` rebuilds the frontend itself (step 5), runs pyflakes + pytest, packages with
`ReportExtractor.spec`, copies `frontend/dist` next to `_internal`, writes `README.txt` / `FIRST_RUN.txt` /
`VERSION.txt`, then **validates** the artifact: the React bundle in both locations (byte-identical),
root-relative `/assets/` URLs with no developer path, `webview/js/*` (the JS↔Python bridge) and the WebView2
interop assemblies. `--no-publish` stops before the LAN update folder; the ZIP and `SHA256SUMS.txt` stay in
`release/`. Use `--skip-frontend` only when `frontend/dist` was just built.

PyInstaller cannot cross-compile, so this must run on Windows. The target machine needs neither Python,
Node.js, npm, Git nor the source checkout; it does need the Microsoft Edge WebView2 Runtime. PowerPoint stays
**optional** — without it the renderer chain falls back to LibreOffice and then the built-in renderer, and the
UI reports which backend produced the preview.

Confirm the running build in the header (`vX.Y.Z · Build NNN`), in Settings → *Phiên bản hiện tại*, and in
`VERSION.txt` before trusting any acceptance result.

## PROMPT-028R — packaged .NET start-up failure (`Failed to resolve Python.Runtime.Loader.Initialize`)

Real failure observed on Windows with the 1.3.3 / Build 016 Portable package: double-clicking
`ReportExtractor.exe` died before the React UI appeared. Chain:
`webview.platforms.winforms` → `import clr` → `pythonnet.load()` → `clr_loader.get_netfx()` →
`clr_loader/netfx.py:50` → `RuntimeError: Failed to resolve Python.Runtime.Loader.Initialize from
_internal\pythonnet\runtime\Python.Runtime.dll`.

That message is raised whenever the native `pyclr_get_function()` returns `NULL`. It reports the path
**pythonnet asked for**, not a path that exists, so it is equally consistent with the assembly being absent,
being a different build, or being present but unloadable. Every candidate cause was therefore measured
against installed source rather than assumed:

| Candidate | Verdict | Evidence |
|---|---|---|
| Duplicate `Python.Runtime.dll` in the package | **Eliminated as a cause of wrong content** | `Analysis` ends with `normalize_toc(self.datas + self.binaries)`; `_TOC_TYPE_PRIORITIES` gives `BINARY`/`EXTENSION` priority 1 over `DATA` 0, so a same-destination `datas` entry is discarded in favour of hook-clr's `binaries` entry. Pinned by `test_pyinstaller_normalize_toc_gives_binaries_priority_over_same_dest_datas`. |
| Wrong PyInstaller category | **Eliminated** | Same mechanism — the hook's `BINARY` entry always wins at `pythonnet/runtime`. |
| `collect_data_files("pythonnet")` interfering with hook-clr | **Eliminated (redundant, not fatal)** | Measured: it returns `pythonnet/runtime/Python.Runtime.dll` as DATA; normalization drops it. Removed anyway so the hook is the single owner. |
| `collect_dynamic_libs(...)` interference | **Eliminated** | Returns the *same* file at the *same* destination; also dropped by normalization. `collect_dynamic_libs` is no longer even imported by the spec. |
| Version mismatch / stale `.venv-build` | **Open — now gated** | `requirements-webview.txt` pinned `clr-loader==0.2.10` (pythonnet 3.0.5 only declares `clr_loader<0.3.0,>=0.2.7`). `verify_runtime_imports()` imports `webview`/`pythonnet`/`clr`/`clr_loader` in the build venv **before** freezing and checks the installed clr-loader satisfies pythonnet's own range. `build_portable.py --fresh-venv` recreates `.venv-build` deterministically. |
| `clr_loader` netfx module missing | **Open — now hardened** | `clr_loader/__init__.py` has no module-level netfx import; `get_netfx()` starts with `from .netfx import NetFx`. The spec now declares `clr_loader.ffi/netfx/types/util` explicitly. |
| DLL collected to the wrong directory | **Concrete, now caught** | hook-clr falls back to `ctypes.util.find_library('Python.Runtime')` with destination `'.'` when `importlib.metadata.files('pythonnet')` does not yield exactly one match — landing the DLL at `_internal/Python.Runtime.dll`, one directory away from where pythonnet looks. Executed in `test_hook_clr_legacy_fallback_would_collect_the_dll_one_directory_too_high`; rejected by `validate_pythonnet_runtime()`. |
| Wrong/modified DLL bytes | **Open — now gated** | `validate_pythonnet_runtime()` requires exactly one `Python.Runtime.dll`, at `_internal/pythonnet/runtime/`, byte-identical (SHA256) to the build venv's copy, and prints both hashes into the build log. |
| Missing .NET on the target | **Real prerequisite, now detected** | `Python.Runtime.dll` targets `.NETStandard,Version=v2.0` (its own `deps.json`), so it needs the netstandard facade of **.NET Framework 4.7.2+** (registry `Release >= 461808`). `desktop.dotnet_framework_report()` reads `HKLM\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full` and logs it; `explain_clr_failure()` names the requirement. Documented in `release_docs/README.txt` §1 and §5. |
| Unsupported combination / wrong clr_loader runtime | **Not reproducible here** | pythonnet's default on win32 is `netfx`; `runtime_diagnostics()` logs the selected runtime as `dotnet_runtime`. |
| Stale PyInstaller cache | **Addressed** | `build_portable.py` already cleans `build/`+`dist/`; `--fresh-venv` covers the venv. |

Runtime facts are written to `logs/app.log` as a single `WEBVIEW_RUNTIME key=value …` line **before** any
window is created, so a build that dies instantly still explains itself: `frozen`, `python`, `platform`,
`pywebview`, `pythonnet`, `clr_loader`, `backend`, `python_runtime_dll`, `python_runtime_dll_exists`,
`dotnet_runtime`, `dotnet_framework`, and — when the probe fails — `clr_import` / `clr_import_error`.
A `WEBVIEW_DOTNET_PROBLEM … | HINT:` line follows when `import clr` fails, and `webview_start()` re-raises
with the hint attached. Nothing user-sensitive is added to the React UI.

The one manual pywebview collection kept is `collect_data_files("webview", subdir="js")`: hook-webview uses
`subdir='lib'` and therefore never collects the JS↔Python bridge injection. It is deliberately restricted to
`js` so it cannot overlap hook-webview's handling of `webview/lib`. React + pywebview/EdgeChromium remains the
default UI; the spec never falls back to Tkinter to mask a start-up failure.

**Windows acceptance still required.** This analysis was performed on Linux, where PyInstaller cannot
cross-compile and `import clr` legitimately fails for want of a .NET host. Run on the Windows build machine:

```
build_portable.bat --no-publish
```

Step 3 now prints the dependency versions and the pre-freeze import results; step 10 fails the build unless
the packaged `Python.Runtime.dll` is unique, correctly placed and byte-identical to the source. If the exe
still does not open, read the `WEBVIEW_RUNTIME` line in `logs/app.log` — it states which of the open causes
above applies instead of leaving it to guesswork.

## Acceptance status

| Status | Scope |
|---|---|
| **INTEGRATED ON LINUX — PROMPT-029 (v1.3.4 / Build 017)** | `git merge --no-ff` of PROMPT-027R (`8f979eb`, tip of `arena/6e89ee9d-tnp`) onto PROMPT-028R (`78805fa`); merge commit `3dae422`, merge-base `b5f8a29` on both sides, so both commits remain ancestors of HEAD and neither line was reimplemented or dropped. Canonical version is now **1.3.4 / Build 017** (`VERSION_LABEL` "1.3.4 — Build 017"), matching `frontend/package.json`; the packaging commits never touched `app/__init__.py`, so nothing regressed to 1.3.3/016. Only 3 of 45 changed paths conflicted: `PROGRESS_LOG.md` (both sides appended at the same anchor — resolved by keeping BOTH blocks verbatim in authoring order), `DEVELOPMENT_WEBVIEW.md` and `tests/test_publish_update.py` (both auto-merged, disjoint regions). Integration verified rather than assumed: **32/32** paths only PROMPT-027R touched are byte-identical to `8f979eb`, and **9/10** paths only PROMPT-028R touched are byte-identical to `78805fa` — the single intentional deviation is `tests/test_prompt028_portable.py`, whose version assertion and baseline-diff scope guard had premises invalidated by the integration and were rewritten (still 80 tests there; none deleted). `frontend/dist` rebuilt for the changed `index.html` (`index-B5WzBAyA.js`, `index-BTW1TSfn.css`). **1087 Python tests passed** (1045 at the PROMPT-028R head + PROMPT-027R's 42: 20 in `tests/test_prompt027r_regions.py`, 22 in `tests/test_prompt027r_renderer.py`), frontend **29/29** (22 + the new `test:prompt027r` 7), `tsc -b && vite build` and `tsc --noEmit` clean, `compileall` clean, `pyflakes app tools run.py` (the build gate's exact scope) clean, `git diff --check` clean, no conflict markers left anywhere. Expected artifacts on Windows: `dist\ReportExtractor_v1.3.4_Portable\ReportExtractor.exe` (from `PORTABLE_NAME = f"ReportExtractor_v{VERSION}_Portable"`) and `release\ReportExtractor_1.3.4.zip`. **STILL NOT DONE:** no Windows build, no EXE, no launch. This sandbox is Linux x86_64 and PyInstaller cannot cross-compile, so BOTH outstanding Windows acceptances — PROMPT-028R's (`build_portable.bat --no-publish`, launch from `dist\`, relaunch after copying the folder to another drive, JS↔Python bridge / Settings / Học cải tiến) and PROMPT-027R's (real multi-item PPTX with PowerPoint, `SLIDE_RENDER … backend=powerpoint faithful=true`, `REGION_EVIDENCE` for Mục #1 ending above Mục #2's heading, Mục #2 independent, final Excel crop) — remain unperformed and must not be reported as passed. |
| **WINDOWS ACCEPTANCE REQUIRED — PROMPT-028R** | On Linux: **1044 Python tests passed** (1012 baseline + 33 new/extended in `tests/test_prompt028_portable.py`, now 80 in that file), `compileall` and configured `pyflakes` passed, `git diff --check` passed, frontend `test:frontend` 22/22 and `tsc --noEmit` passed. The shipped `hook-clr.py`, `hook-clr_loader.py`, `hook-webview.py` were executed directly and PyInstaller's `normalize_toc` priority behaviour was measured, so the duplicate-DLL / wrong-category / `collect_*`-interference hypotheses are **eliminated with evidence** rather than assumed. Version stays **1.3.3 / Build 016** and no PROMPT-027R evidence-region file was touched. **No Windows EXE was built or launched here** — PyInstaller cannot cross-compile from Linux, so the mandated acceptance (`build_portable.bat --no-publish`, launch from `dist\`, launch again after copying the folder to another drive, bridge/Settings/Học cải tiến, then the real-PPTX `powerpoint` backend check) is **NOT DONE** and must not be reported as passed until it is performed on Windows. |
| **AUTOMATED TESTED — PROMPT-023** | 875 Python tests passed (82 targeted picker/runtime-isolation/updater tests); compileall and configured pyflakes passed; frontend typecheck and production build passed. The build emits a hashed local favicon referenced through `./assets/...`; `git diff --check` passed. The suite covers the bridge filter/cancel/path contracts, isolated persistent state, and prior extraction, evidence, learning, Excel safety/lock-retry, force-reprocessing, Ollama, and updater regressions. |
| **AUTOMATED TESTED — PROMPT-025** | 894 Python tests passed (875 baseline + 19 new in `tests/test_prompt025_multi_item.py`); compileall, pyflakes (new/changed files), and `git diff --check` passed; frontend typecheck and production build passed. Multi-improvement segmentation + item-scoped After regions are covered by the §30 two-item fixture (3-picture After block = one region, item #1 crop excludes item #2 text, whitespace policy, caption/arrow association, Before exclusion), heading-style variants (defect-style "Lỗi … sau ép nhỰA", mid-block non-bold colon heading), 3-item and 1-item slides, report isolation, force-reprocess, learning identity per report+slide+item, render-once-per-slide with two item crops, the cooperative cancellation hook, file-atomic stop-after-current with a multi-item slide, Excel mapping (one row, both items in order, two image groups), temporary-action exclusion, and the `cải進` caption fold. All prior PROMPT-014/015/020/021/023/024 regressions still pass. |
| **AUTOMATED TESTED — PROMPT-027** | 965 Python tests passed (911 baseline + 54 new in `tests/test_prompt027_rendering.py`); compileall, configured pyflakes and `git diff --check` passed; frontend `test:bridge` (14), `test:learning-hooks` (1) and the new `test:prompt027` (7) passed; typecheck and production build passed. Covered headlessly: `No Fill` / `No Line` parsed as *paints nothing* (`<a:noFill/>` → `MSO_FILL.BACKGROUND`, which is **not** `None`) and no black frame in the built-in renderer, including grouped children; full-slide render keeps the authored aspect ratio, whitespace and object pixel position with no white-trim; the built-in renderer honours the authored vertical anchor / text insets / line spacing / `normAutofit` scale (the cause of the vertically shifted preview); EMU→pixel crop conversion against the ACTUAL rendered dimensions; item-scoped region ownership with no cross-item contamination (pixel-level); authored whitespace preserved between two After pictures; bounded deterministic region padding; `SLIDE_RENDER` / `SLIDE_RENDER_BACKEND_FAILED` diagnostics; preview-cache identity now includes renderer chain + width + render schema version with report isolation and mtime invalidation; the learning DTO exposes the preview backend, faithfulness and the authoritative evidence-region bbox; overlay alignment for BOTH geometries under fit/resize/letterbox/zoom/DPI; frontend/backend version agreement at 1.3.3 / Build 016. |
| **WINDOWS RETEST REQUIRED AFTER PROMPT-027** | First confirm the header shows **v1.3.3 · Build 016** and Settings shows **1.3.3 — Build 016**; if Windows still shows 1.3.2 / Build 015, STOP — that is stale code or a stale `frontend/dist`. Then, on the real PPTX slide with 2 improvements, authored whitespace, an invisible `No Fill` / `No Line` content frame and multiple After pictures, verify: the full-slide preview matches PowerPoint with content no longer shifted upward; authored whitespace preserved; the invisible frame stays invisible with no fake black rectangle in either the preview or the final Excel evidence; the preview reports its renderer backend and logs `SLIDE_RENDER purpose=learning_preview backend=powerpoint` and `purpose=after_evidence backend=powerpoint`; the picture-candidate overlay (dashed) and the final evidence-region overlay (solid, "Mục #n · Vùng xuất Excel") are both aligned and visually distinct; Mục #1 and Mục #2 keep independent regions when switching candidates; the exported Excel crop keeps the intended whitespace and contains no neighbouring item text. No Linux check substitutes for Windows acceptance of the real file, and PowerPoint-faithful rendering cannot be claimed until it is seen there. |
| **AUTOMATED TESTED — PROMPT-027R** | 1007 Python tests passed (965 baseline + 42 new: 20 in `tests/test_prompt027r_regions.py`, 22 in `tests/test_prompt027r_renderer.py`; no test deleted); compileall, pyflakes (same message set as baseline) and `git diff --check` passed; frontend `test:bridge` (14), `test:learning-hooks` (1), `test:prompt027` (7) and the new `test:prompt027r` (7) passed; typecheck and production build passed. Covered headlessly on a synthetic slide with the reported layout: After regions are structural (same-row/column blocks with clear gaps, picture-relative spacing, caption-side separation, group ancestry as association only); a caption never labels a picture across another item's structure; inline `+ Sau:` lines cannot override an arrow vote; callouts and short labels on or under a picture are not item headings; item spans end at the next item heading and the final crop is clamped against every heading it horizontally reaches (members always kept, conflicts logged); invisible `No Fill`/`No Line` containers are neither drawn nor boundaries; body text is excluded unless it lies in the footprint; the Learning DTO exposes the production region, its member picture ids and the boundary that limited it, computed with the same learning decisions Excel uses; the PowerPoint probe is side-effect free, checks COM then the 64/32-bit registry views, COM work uses an application-owned instance with balanced CoInitialize/CoUninitialize, and every failure names its stage and HRESULT; preview cache identity includes file size, mtime_ns and backend availability, and a degraded fallback expires after 15 s; the Learning preview shows ONE unmodified slide bitmap with outline overlays only. |
| **WINDOWS RETEST REQUIRED AFTER PROMPT-027R** | Confirm **v1.3.4 · Build 017** / **1.3.4 — Build 017** first. On the real multi-item PPTX (slide 3): (1) with PowerPoint installed, `SLIDE_RENDER purpose=learning_preview backend=powerpoint … faithful=true` and no reduced-fidelity notice; if a `SLIDE_RENDER_BACKEND_FAILED … stage=` line appears instead, record the stage, exception_type and hresult; (2) compare the same slide in PowerPoint and the Learning preview (title, body, weight, wrapping, gray text, whitespace); (3) Mục #1: `REGION_EVIDENCE` lists every intended After picture id, `region_bottom_y` < `next_heading_y` of Mục #2, `boundary_source=next_item_heading`, and the Excel crop contains no Mục #2 heading or body text; (4) Mục #2: its own After pictures only, its Before picture absent; (5) without PowerPoint or LibreOffice the app still starts, reports stage `availability`, and the built-in preview shows its reduced-fidelity notice. No Linux check substitutes for this acceptance. |
| **WINDOWS RETEST REQUIRED AFTER PROMPT-025** | The real two-item PPTX (item #1 "Lỗi xước rear …" with a three-picture After group + "Sau cải進" caption, item #2 "Cải tiến lỗi lệch ATN …"): confirm exactly 2 improvement items, 2 item-scoped After crops in the Excel image cell (one horizontal image row per item), item #1's crop free of item #2 heading/body, the transition arrow excluded, learning preview showing the "Mục #n" chip per candidate, and stop-after-current/lock-retry behaviour unchanged. No Linux check substitutes for Windows acceptance of the real file. |
| **WINDOWS VERIFIED BEFORE PROMPT-023** | The pre-patch checks listed above. They confirm the earlier integrated runtime, not the picker-filter or persistent-state fixes in this change. |
| **WINDOWS RETEST REQUIRED AFTER PROMPT-023** | Native picker open/cancel behavior and all post-patch real-report processing, Excel edit/lock/retry, stop-after-current, rendered evidence, and batch checks in the checklist below. No Linux GUI feasibility check substitutes for Windows acceptance. |
| **DISABLED WITH REASON** | Applying/installing updates is disabled in the development UI/service so no updater replaces or restarts the checkout. Update checking remains a read-only backend operation. |
| **NOT IMPLEMENTED** | Hot-module-reload Vite server embedded directly in the pywebview window. The supported development flow is `npm.cmd run build`, then `python -m app.desktop`; opening Vite in a normal browser has no mock/demo bridge fallback. |

## Post-patch Windows manual acceptance

Use disposable copies of reports, templates, and output workbooks. Keep the test local; do not publish an update or release.

1. Run the Python 3.12 setup sequence above, launch the app, and confirm version **1.3.4 — Build 017** and the connected
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
