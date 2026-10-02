# Progress log

- Phase 1  GUI skeleton, config, recursive scanner, DnD path parsing ............ done (tests)
- Phase 2  Ollama client (/api/tags, /api/generate json, temp 0), URL normalisation,
           Qwen-4B preference, fake-HTTP-server tests ............................ done
- Phase 3  python-pptx parser (text/tables/pictures/positions/reading order),
           heuristic + LLM classifier with QPN exact-text guard, section walker,
           temporary-handling exclusion, verbatim-only LLM values ................ done
- Phase 4  openpyxl writer: template copy, alias header detection, style copy,
           blank vendor/date/weeks, wrap+top, row height, mapping sheet .......... done
- Phase 5  QPN full-slide render chain (PowerPoint COM → LibreOffice → built-in) . done
           (COM / LibreOffice paths can only execute on a machine that has them)
- Phase 6  Improvement contact sheet + original picture dump ..................... done
- Phase 7  Batch processor, stop-after-current, errors.log, batch_result.json,
           history.json duplicate protection, CLI mode, diagnostics .............. done
- Phase 8  PyInstaller spec + build_portable.bat / RUN_DEV.bat ................... written
           Sandbox is Linux without libpython/tkinter and GitHub assets blocked →
           the .exe could not be produced here; run build_portable.bat on Windows.
- Fixes    openpyxl closed-image-buffer on 2nd save; header mis-detection on reopen;
           cover title mis-classified as improvement; defect cell extraction.
- Tests    40 passed (pytest).

- PROMPT-004  Final production extraction / fixed-master rules (v1.0.3, no new release):
           no auto-created production rows (MASTER_NOT_FOUND status, prescan counter), duplicate rows →
           topmost used + others red, diagnostics "Dòng sử dụng / bị trùng tại dòng / Đã đánh dấu đỏ",
           one safety backup before first workbook modification per batch (none for skip-only, failure →
           master untouched), content-region / After-only picture rules re-verified.
           tests/test_prompt004.py (25 tests) ................................... done, 488 passed
           Real-data validation (September five-report dataset, Windows) ........ NOT run in sandbox

- PROMPT-004A Ollama local-first connection: 127.0.0.1 → saved server → manual/LAN (never automatic
           LAN scan); probe_timeout (3 s) separate from request_timeout (180 s); GUI controls updated
           when local is selected. tests/test_ollama_local_first.py (14 tests) ..... done, 502 passed

- PROMPT-004B GUI contrast hotfix: declarative theme-aware style spec with explicit fg/bg for every state,
           clam preferred, WCAG contrast audit; palette tuned (secondary/warning/error/muted);
           tests/test_gui_contrast.py (39 tests) ............................... done, 541 passed

- PROMPT-004C Scanner visibility: discovery separated from eligibility, SCAN_REJECT diagnostics, default list
           shows every discovered report, attention rows never hidden, bucketed counters;
           tests/test_scanner_visibility.py (7 tests) ........................... done, 548 passed

- PROMPT-004C-B After-picture evidence without captions: parser exposes per-paragraph colours (RGB + theme)
           and directional arrows (auto-shapes/connectors, rotation/flip); selector adds blue-After-text and
           arrow-destination evidence (block-scoped, contiguous groups, conflicts → ambiguous);
           tests/test_after_evidence.py (18 tests) ............................ done, 566 passed

- PROMPT-004D Automatic new Management Number row restored (reverses the PROMPT-004 "never create" rule, keeps 004C):
           valid key absent from the master → prescan action PROCESS_NEW_ROW ("Sẽ xử lý — Management Number mới"),
           BatchProcessor creates exactly ONE production row (format-only clone, no business values / WEEK / manual
           fields), writes the key, saves, then processes normally → stage "completed_new"
           ("Hoàn thành — đã thêm Management Number mới"); date from the key only; backup before the row,
           backup failure → no row; no valid key in the filename → no row ("Cần kiểm tra — Không xác định được
           Management Number từ tên file"); counters "Management Number mới: N"; diagnostics
           "Đã tạo dòng mới: <row> – Management Number mới: <key>".  tests/test_master_not_found.py rewritten to the
           004D rule (8 tests), tests/test_new_master_row.py restored, tests/test_new_row.py (19) . done, 590 passed
           Real-data validation (September dataset, Windows) ......................... NOT run in sandbox
