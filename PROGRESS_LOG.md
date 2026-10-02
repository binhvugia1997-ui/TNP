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

- PROMPT-004E Portable build must not require Git: tools/build_portable.py `get_git_revision()` (never raises;
           FileNotFoundError / CalledProcessError / TimeoutExpired / OSError → "unavailable"),
           `write_release_metadata()` (VERSION.txt `git=unavailable` + `Git revision: unavailable`, README
           `Git: unavailable`); also fixed latent `render_doc(name=…)` keyword collision in the same step.
           tests/test_build_git_optional.py (12 tests) ............................ done, 602 passed
           Windows portable build rerun ........................................... NOT possible in Linux sandbox

- PROMPT-004F Portable validator false positive: python-pptx runtime resource _internal/pptx/templates/default.pptx
           was flagged as sample data.  Narrow allowlist ALLOWED_DEPENDENCY_RESOURCES (exact component tuple,
           case/separator independent) via is_allowed_dependency_resource(); every other .pptx/.ppt/.xlsx,
           real_data/September reports, Kiem_chung.xlsx, config.json, ollama/gguf, tests/__pycache__ stay
           rejected.  8 regression tests in tests/test_build_git_optional.py .......... done, 610 passed
           Windows portable build rerun ........................................... NOT possible in Linux sandbox

- PROMPT-005 Offline/LAN self-update (v1.0.4, Build 004): numeric BUILD_NUMBER as the only build identity
           (Git = internal metadata), app/updater.py (manifest parsing, numeric comparison, package safety incl.
           SHA256, local staging, external updater via the staged copy, wait-for-exit, update_backup rollback,
           restart, update.log), config.update_path, GUI card "Cập nhật phần mềm" (path + Chọn... + Kiểm tra cập
           nhật + Cập nhật ngay, Cập nhật/Để sau confirmation, async startup check), --apply-update entry point,
           build script writes release/ReportExtractor_<version>.zip + version.json (exact SHA256).
           tests/test_updater.py (46 tests) ...................................... done, 656 passed
           Windows portable build / real update round-trip ......................... NOT possible in Linux sandbox


- PROMPT-006 Improvement-image learning & correction mode (v1.1.0-beta, Build 006 – experimental, not for
           production release).  §0 stable backup backup/ReportExtractor_v1.0.4_Build004_STABLE (83 tracked source
           files, verified, never modified) + tools/make_stable_backup.py.  app/image_learning.py: ImageCandidate
           records (normalised PPT geometry, caption/blue/arrow/inspection/logo evidence, IMAGE_FEATURE_SCHEMA=1,
           path-independent candidate_id), calibrated deterministic confidence + Vietnamese evidence lines,
           LabelStore (learning_data/image_labels.jsonl, append-only, last label wins, relabel audit, duplicate
           protection), pure-Python logistic regression AFTER vs NON_AFTER (MIN 10 examples / 3 per class,
           learning_data/model/image_model.json with metadata; missing/corrupt/incompatible → rules only),
           decision = hard exclusions > user labels > deterministic thresholds (0.75 / 0.35) > model on the
           uncertain band only.  Pipeline: extract_record(learning=…), BatchOptions.learning_dir, FileResult
           .image_candidates; app/image_review.py re-applies confirmed labels to the improvement_image cell only
           (ExcelWriter.replace_improvement_images, backup rule unchanged).  GUI tab 2 card "Dữ liệu học ảnh cải
           tiến" (counts, Kiểm tra ảnh cải tiến review window, Cập nhật mô hình ảnh, Mở thư mục dữ liệu học, Xuất
           dữ liệu học).  Updater preserves learning_data; build validator forbids it in the artefact.
           tests/test_image_learning.py (27 tests) ................................ done, 683 passed
           Real-data labelling / learned-model validation on Windows ............... NOT performed (no labels yet)

- PROMPT-006B Learning / correction for improvement CONTENT regions (v1.1.0-beta, Build 007 – beta).
           app/content_learning.py: ContentCandidate per text block of the improvement slides
           (CONTENT_FEATURE_SCHEMA=1, 28 normalised geometry/typography/keyword/section features, identity
           MN|S<slide>|SH<shape>|<order>), deterministic kind + confidence + VI evidence, hard exclusions (title,
           sidebar, caption buttons, footer/logo/decoration, XỬ LÝ TẠM THỜI), labels IMPROVEMENT_CONTENT /
           EXCLUDE_CONTENT → learning_data/content_labels.jsonl, shared logistic regression →
           learning_data/model/content_model.json (MIN 20 / 5 per class), decision = hard exclusions > user label >
           strong rule evidence > model (uncertain band) > extractor fallback; verbatim rebuild in reading order
           (Section.sources traces every line to its shape).  image_review.reapply_content_labels +
           ExcelWriter.replace_improvement_text (improvement cell only).  GUI card: "Kiểm tra nội dung cải tiến"
           review window, "Cập nhật mô hình học" trains both models independently with separate counts/results.
           tests/test_content_learning.py (20 tests) .............................. done, 703 passed
           Real-data content labelling / learned-model validation on Windows ....... NOT performed (no labels yet)
