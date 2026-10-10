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

## PROMPT-006B addendum — Danh sách báo cáo gọn + file loại thủ công tự xuống cuối (Build 007, không đổi build)
- Cột hiển thị: STT · Management Number · Ngày phát sinh · Vendor · Tên file (giãn) · Trạng thái. Không còn cột "Đường dẫn";
  đường dẫn đầy đủ / lý do quét / ghi chú vẫn nằm trong `ScanRow` và hiện qua nhấp đúp (`ScanRow.details()`).
- Nhóm hiển thị cố định (`GuiController.scan_rows`, sort ổn định theo thứ tự quét trong từng nhóm):
  đang xử lý → chờ / sẽ xử lý (kể cả Management Number mới) → cần kiểm tra / lỗi → hoàn thành → bỏ qua (đã đủ dữ liệu,
  ngoài thời gian, trùng nguồn) → **Đã loại thủ công (luôn ở cuối)**.
- Loại thủ công: dòng giữ nguyên, trạng thái `Đã loại thủ công`, ra khỏi hàng đợi (`final_queue`), xuống cuối ngay, không quét lại.
  `Khôi phục` trả về đúng nhóm + hàng đợi. STT = thứ tự hiển thị hiện tại (tính lại 1..n), `index`/`key` không đổi.
- Quét lại cùng thư mục giữ nguyên danh sách đã loại (khoá = đường dẫn chuẩn hoá + Management Number); đổi thư mục báo cáo
  mới xoá danh sách này. Bộ lọc mới "File đã loại thủ công"; "File cần xử lý" không hiện dòng đã loại.
- Danh sách chỉ sắp xếp lại khi một báo cáo đạt trạng thái cuối hoặc khi kết thúc batch (không nhảy dòng khi đang xử lý).
- Kiểm thử: `tests/test_scan_list_compact.py` (13) — tổng 716 passed. Chưa build lại Portable.

## PROMPT-007A — Giao diện co giãn + cuộn dọc (v1.1.0-beta / Build 008)
- Nguyên nhân Build 007 "card chỉ còn tiêu đề": tab Cấu hình là một grid phẳng, không có lớp cuộn, và
  `page.rowconfigure(2, weight=1)` đặt trọng số vào hàng card *Cập nhật phần mềm* (cộng với hàng log). Khi tổng chiều cao
  yêu cầu vượt cửa sổ (màn hình nhỏ / DPI 125–150 %), grid thu các hàng có trọng số về 0 → thân card bị cắt, card sau
  đè lên card trước.
- Sửa: helper `_scroll_page()` (Canvas + Scrollbar dọc + Frame trong) dùng chung cho cả hai tab; frame trong luôn rộng bằng
  viewport và cao = max(viewport, chiều cao yêu cầu) → cửa sổ thấp thì cuộn, cửa sổ cao thì bảng/log nở ra; chỉ hàng log
  có trọng số ở tab 2; không card nào có chiều cao cố định; không tạo lại widget khi resize.
- Nhãn giải thích dài bám theo bề rộng trang (`_wrap_label` / `_refit_wrap_labels`, tối thiểu 320 px) thay vì 900 px cố định;
  ô Server / Đường dẫn cập nhật / Thư mục báo cáo / File kiểm chứng / Kết quả dãn ngang (`sticky="ew"`); nút giữ bề rộng tự nhiên.
- Con lăn chuột: một handler, định tuyến theo widget dưới con trỏ tới đúng trang cuộn; Treeview/Text/Combobox/Spinbox tự cuộn;
  cửa sổ khác không bị chiếm.
- Kích thước tối thiểu cửa sổ 880×540 (cuộn là phương án dự phòng); cửa sổ kiểm tra ảnh / nội dung cải tiến: resizable,
  minsize 640×420, vùng chữ có thanh cuộn riêng, hàng nút nhãn / điều hướng không có trọng số nên luôn còn trong tầm với.
- Danh sách báo cáo giữ 6 cột gọn; Treeview có thanh cuộn dọc + ngang riêng.
- Kiểm thử: `tests/test_gui_responsive.py` (17) với fake-tk ghi lại Canvas/Toplevel; tổng 733 passed. Chưa build Portable;
  cần kiểm tra thủ công trên Windows (A–F, 100/125/150 %).

## Build 008 blocker — autoconnect không được ghi đè server LAN đã xác nhận (v1.1.0-beta / Build 008, không đổi build)
- Nguyên nhân: `auto_connect()` chạy trong thread nền từ lúc khởi động và ghi thẳng `host/port` (= 127.0.0.1) vào
  controller khi probe xong; trên Windows probe local chậm hơn nên kết thúc SAU khi người dùng xác nhận server tìm được
  trong LAN → ghi đè lựa chọn mới (`DISCOVERY_APPLY` rồi `OLLAMA_AUTOCONNECT source=local`).
- Sửa (deterministic, không sleep): `endpoint_epoch` tăng mỗi khi người dùng chọn endpoint (gõ host/port, xác nhận
  discovery); autoconnect chụp epoch khi bắt đầu và chỉ được ghi kết quả / phát sự kiện GUI nếu epoch không đổi; kết quả cũ
  bị bỏ qua (`OLLAMA_AUTOCONNECT ignored=stale`, đếm `autoconnect_stale`). Ưu tiên: lựa chọn người dùng > autoconnect.
- Kiểm thử: +4 trong `tests/test_gui_discovery.py` (autoconnect bắt đầu trước → discovery xác nhận → callback cũ đến sau;
  server gõ tay; khởi động bình thường vẫn chọn local; kịch bản GUI). Tổng 737 passed.

## PROMPT-009 — Tự động kiểm tra cập nhật LAN + thông báo (v1.1.0-beta / Build 009)
- Controller: `auto_update_check` (cfg.extra, mặc định bật), `check_update(startup)`, `check_update_async(startup)` tôn trọng
  tuỳ chọn, `update_status_text()` trả thông báo nhẹ khi kiểm tra tự động thất bại, `pending_update_offer()` (một lần/build/phiên).
- GUI: checkbox trong thẻ Cập nhật phần mềm; `_startup_update_check` sau `STARTUP_UPDATE_DELAY_MS` = 3000 ms; worker chỉ đặt
  cờ `update_dirty`, `_poll` (Tk thread) render + gọi `_offer_update` → hộp thoại [Cập nhật ngay]/[Để sau];
  `install_update(confirmed=True)` đi vào luồng cập nhật hiện có. Thẻ không đổi chiều cao cố định, vẫn nằm trong trang cuộn 008.
- Kiểm tra phiên bản chỉ đọc version.json (+ `is_file()` gói); ZIP chỉ được copy khi người dùng bấm Cập nhật ngay.
- Kiểm thử: `tests/test_auto_update_check.py` (15), cập nhật `test_updater.py` (startup thất bại mềm, độ trễ 2–5 s). Tổng 752 passed.
- Chưa build Portable.
- Sửa blocker Windows (test `test_unavailable_path_does_not_crash_and_shows_soft_status`): phân giải tên UNC thật trên
  Windows có thể mất hàng chục giây nên `join(10)` hết hạn trước khi worker công bố kết quả (race trong test, không phải lỗi
  production). Thêm `update_check_done` (threading.Event) vào controller — chỉ set SAU khi `update_check` được gán; test chờ
  event này và phân biệt "đang chạy" với "đã xong = inaccessible"; hai trường hợp UNC dùng stand-in offline tất định,
  thư mục cục bộ thiếu vẫn đi qua updater thật. +1 test hợp đồng trạng thái. Tổng 753 passed.

## Build 010 — Tự động xuất bản vào thư mục cập nhật LAN (v1.1.0-beta / Build 010)
- `tools/publish_update.py` (mới, tập trung đích xuất bản): ZIP → .tmp → kiểm tra size + SHA256 → đổi tên; version.json
  qua .tmp + os.replace SAU CÙNG; xoá ZIP cũ sau khi bản mới đã sống; thất bại giữ nguyên bản cũ, không để manifest trỏ tới
  ZIP thiếu; có CLI để xuất bản lại.
- `tools/build_portable.py`: bước 12 chỉ chạy sau toàn bộ kiểm tra; `--no-publish`, `--publish-dir`; mã thoát 3 =
  build OK nhưng PUBLISH FAILED (bat hiển thị riêng). Không đổi updater runtime.
- Kiểm thử: `tests/test_publish_update.py` (15). Tổng 768 passed. Chưa build Portable.

## PROMPT-011 — Cửa sổ tiến trình cập nhật (v1.1.0-beta / Build 011)
- `app/updater.py`: `copy_with_progress()` copy theo khối 1 MB, callback `(bytes_copied, total_bytes, percent)` tính từ
  byte thật (không timer); `stage_update(progress=)` phát PREPARING → COPYING → VERIFYING → READY từ thread gọi.
- `app/gui_controller.py`: `UpdateProgress` (snapshot bất biến, tốc độ = bytes/elapsed monotonic), `UPDATE_STAGES_VI`,
  `install_update_async()` (worker `update-install`, một lần duy nhất, `update_installing` khoá nút), sự kiện
  `update_progress` / `update_done`; `install_update()` đồng bộ vẫn giữ cho CLI/test; lỗi → stage ERROR, dọn staging.
- `app/gui.py`: cửa sổ "Đang cập nhật Report Extractor / Build X → Build Y" (resizable), Progressbar determinate khi copy,
  "58.4 MB / 87.1 MB  67%", "Tốc độ: 11.2 MB/s"; mọi cập nhật widget qua `_poll` (Tk thread); đóng cửa sổ khi đang copy
  chỉ ẩn; sau khi gói đã kiểm tra 100% hiển thị "Gói cập nhật đã sẵn sàng. Đang khởi động trình cài đặt..." rồi mới
  bàn giao cho trình cập nhật ngoài (EXE đang chạy không bao giờ tự ghi đè; backup/rollback/khởi động lại giữ nguyên).
- Kiểm thử: `tests/test_update_progress.py` (10); test_15 discovery được làm tất định (gate thay vì sleep). Tổng 778 passed.

## Phiên bản 1.2.0 / Build 011 — rời chuỗi 1.1.0-beta
- `app/__init__.py`: `__version__ = "1.2.0"`, `BUILD_NUMBER = 11` (hiển thị v1.2.0 / Build 011). Gói: `ReportExtractor_1.2.0.zip`;
  `version.json`: `"version": "1.2.0", "build": 11`. Số Build vẫn là tiêu chí sắp xếp cập nhật → máy đang chạy
  1.1.0-beta / Build 010 nhận 1.2.0 / Build 011 là bản cập nhật bình thường (test `test_legacy_beta_client_sees_1_2_0_build_011_as_normal_update`).
- Test nhận dạng hiện tại cập nhật sang 1.2.0; fixture "bản mới hơn" dùng 1.2.1 / Build 012+; tham chiếu lịch sử giữ nguyên.
- Dọn các file do một lần chạy nhầm trình cập nhật ngoài để lại ở gốc repo (ReportExtractor.exe, VERSION.txt, README.txt,
  _internal/) — đã bỏ khỏi Git, thêm vào .gitignore; conftest chặn mọi test spawn trình cập nhật thật. Tổng 779 passed.

## BUGFIX — khối chữ trong GROUP: cửa sổ "Kiểm tra nội dung cải tiến" chỉ hiện "Massage"
- Nguyên nhân gốc (`app/pptx_parser.py`): shape con của `p:grpSp` được lưu theo toạ độ con (`a:chOff/a:chExt`); parser
  dùng thẳng giá trị thô → khi group đã bị kéo/thu phóng, mọi textbox trong group bị đặt sai vị trí (rơi vào dải tiêu đề /
  cột sidebar) nên tiêu đề/nội dung mục cải tiến bị loại trừ cứng hoặc đổi thứ tự, còn nhãn nhỏ "Massage" giữ bằng chứng
  vùng. Kèm theo: shape copy-paste trong group có thể trùng `cNvPr id` → sections/candidates/nhãn bị gán chéo theo
  `(slide, shape_id)`.
- Sửa: `GroupXform` (off/ext ↔ chOff/chExt, lồng nhau, có scale) cho textbox/bảng/ảnh/mũi tên; `_dedupe_shape_ids`
  mỗi slide; `lookup_override` giữ nhãn cũ theo `report|S|SH` khi thứ tự đọc đổi (chỉ khi không mơ hồ).
- Không đổi quy tắc nghiệp vụ trích xuất; không đổi version/build (1.2.0 / Build 011).
- Test: `tests/test_content_group_candidates.py` (12). Tổng 791 passed.

## PROMPT-011B — Content learning: Excel bị khoá (WinError 32) sau khi lưu nhãn (v1.2.0 / Build 011)
- Nguyên nhân gốc: `reapply_content_labels` thêm dòng vào `updated_rows` ngay khi sửa ô trong bộ nhớ, rồi mới
  `writer.save()`; `save()` ghi `master.saving.xlsx` và `shutil.move` đè lên master → Excel đang mở master
  (khoá ngoài, không phải handle của app) → WinError 32 tại bước đổi tên; ngoại lệ bị gộp vào `errors` nên GUI báo
  "đã cập nhật 7 dòng" + "Lỗi WinError 32" cùng lúc; file `.saving.xlsx` bị bỏ lại.
- Sửa: `ExcelWriter.save()` → `os.replace` nguyên tử, dọn temp, `ExcelLockedError(path)` chỉ khi là sharing/lock
  violation (`is_sharing_violation`: WinError 32/33), PermissionError khác giữ nguyên; `ReapplyResult` tách
  `prepared_rows` / `updated_rows` (= committed) / `locked_path` / `technical_error`, commit chung `_commit()` cho cả
  ảnh và nội dung; controller: nhãn là giao dịch riêng (`CONTENT_REVIEW_LABELS_SAVED`), Excel lỗi → giữ danh sách
  chờ (`pending_content_reapply`), `retry_content_reapply()` chỉ làm lại Excel; GUI: hộp [Thử lại] [Để sau] + nút
  "Cập nhật Excel từ nhãn đã lưu". Log: CONTENT_REAPPLY_START / PREPARED / COMMIT_OK / COMMIT_FAILED.
- Test: `tests/test_excel_lock_retry.py` (13). Tổng 804 passed.

## PROMPT-012 — BUILD_AND_PUBLISH.bat (v1.2.1 / Build 012)
- `app/__init__.py`: 1.2.1 / BUILD_NUMBER 12 (gói `ReportExtractor_1.2.1.zip`). Fixture "bản mới hơn" → 1.2.2 / Build 013+.
- `tools/build_and_publish.py`: `git_sync` (fetch + pull --ff-only; dừng khi dirty/ahead/diverged/detached/sai nhánh;
  không reset/stash/force), `check_not_already_published` (so với version.json đang phát hành), chạy
  `tools/build_portable.py` (stream + log `logs/build_and_publish_*.log`), mã thoát 0/1/2/3/4.
- `BUILD_AND_PUBLISH.bat` ở gốc repo: tìm Python 3.10–3.13 + git, gọi orchestrator, báo kết quả tiếng Việt.
- Test: `tests/test_build_and_publish.py` (19, dùng git thật trên repo tạm). Tổng 823 passed. Chưa build Portable.

## PROMPT-025 — Một slide nhiều mục cải tiến: phân đoạn ImprovementItem + crop "Sau" chính xác theo mục (v1.3.2 / Build 015, không đổi build)
- Giả định cũ "một slide = một mục cải tiến" sai với báo cáo thật: một slide "3. CẢI TIẾN TRONG SẢN XUẤT" có thể chứa
  nhiều mục (ví dụ: mục 1 "Lỗi xước rear (Áp dụng cải tiến 17/9 – Công đoạn Assy Daoltech):" với Before ở trái,
  mũi tên xanh giữa, After là BA ảnh nằm ngang bên phải và caption "Sau cải進" bên dưới; mục 2 "Cải tiến lỗi lệch
  ATN ..." ngay dưới). Hệ quả cũ: một item duy nhất, một region gộp cả hai mục, ảnh "Before" bị nhận nhầm thành
  "Sau", và crop của mục 1 chứa cả heading/body của mục 2.
- Nguyên nhân gốc (đã nhân bản bằng fixture tổng hợp, không cần Windows): (1) `slide_items()` không nhận được
  heading của mục 2 khi heading đó là dòng kiểu tên lỗi "Lỗi lệch ATN sau ép nhựa" (chứa "sau" → đúng tên bị loại
  bởi exclusion regex của branch defect) hoặc là dòng không in đậm, có dấu hai chấm, nằm giữa một khung chữ chung;
  (2) các anchor "+ Trước:/+ Sau:" và caption được áp dụng theo toàn slide, không theo mục; (3) không có mô hình
  ImprovementItem nên không có gì neo crop theo từng mục; (4) đường biên của mục bị cắt bởi nhãn sidebar hẹp ở
  mép trái (oval "Cải tiến trong kiểm tra") nên span của mục 1 mất cả hàng ảnh của chính nó.
- Sửa (không đổi version/build 1.3.2 / Build 015):
  - `app/improvement_items.py` (mới): mô hình `ImprovementItem` (Report → Slide[] → ImprovementItem[] → Before/After
    regions): phân đoạn theo heading, marker Trước/Sau, mũi tên, hình học và khoảng trắng dọc (không OCR, không
    detect khung đỏ/chụp màn hình); span của mục = từ anchor tới đỉnh đường biên kế tiếp có liên quan theo phương
    ngang (hoặc full-width), không bao giờ thu nhỏ bên dưới khung anchor; text của mục theo span (dòng ở khoảng
    trắng giữa các mục thuộc về không mục nào nhưng vẫn vào ô Excel qua join section — không mất nội dung).
  - `app/improvement_pictures.py`: nhận diện anchor của mục theo 2 họ đúng thật (heading có dấu ":" + ngữ nghĩa sản
    xuất, hoặc mở đầu bằng tên lỗi "Lỗi lệch ATN sau ép nhỰA"), luôn kèm tín hiệu thị giác (có caption/ảnh
    trong hàng bên dưới) để không phân đoạn nhầm dòng body; span chỉ bị cắt bởi đường biên có overlap theo trục
    ngang hoặc rộng ≥ 55% slide; anchor (caption / inline / chữ xanh / mũi tên / claims) được scope theo span của
    từng mục với fallback toàn slide khi mục không có anchor (tương thích ngược); log `ITEM_SEGMENTATION`.
  - `app/improvement_visual.py`: `ImprovementVisualRegion` thêm `item_index` + `after_block_index`; gom cụm các ảnh
    "Sau" của một mục thành các After block (cùng hàng/cùng cột hoặc chung caption → 3 ảnh một block là MỘT
    region, không phải 3); mỗi block là một region riêng; bbox bị clamp theo span dọc của mục (khi đường biên có
    overlap theo trục ngang) → crop của mục nọ không bao giờ chứa heading/body của mục sau; caption gán theo
    block gần nhất (hòa thì loại); sort theo (slide, item_index, after_block_index, source_order).
  - `app/image_extractor.py`: render mỗi slide đúng một lần rồi crop theo từng region của từng mục; thêm hook
    huỷ hợp tác `should_cancel` giữa các region (batch không truyền → vẫn "Dừng sau file hiện tại" nguyên khối file,
    không bao giờ ngắt commit Excel); `owners` dùng `region.block_id` (ổn định, duy nhất theo block).
  - `app/extractor.py`: giữ nguyên ô Excel improvement = join_sections (đủ các mục, đúng thứ tự, không gộp chéo);
    thêm `ExtractedRecord.improvement_items` (summary mỗi mục: slide/index/itemId/scopedItemId/heading/text/bounds/
    before/after/captions) tính sau khi chọn ảnh cuối cùng.
  - `app/logger.py` + `app/batch_processor.py`: `FileResult.improvement_items` (batch_result.json).
  - `app/image_learning.py`: `ImageCandidate.item_index` (điền từ segmentation); `candidate_id_for` giữ nguyên
    (report|slide|shape|order → tương thích ngược, migration an toàn).
  - `app/application_service.py`: DTO learning thêm `itemId`/`itemIndex`/`itemHeading`.
  - `frontend/src/types.ts` + `frontend/src/tabs/LearningTab.tsx`: thêm field optional + chip "Mục #n" nhỏ cạnh
    nhãn Slide ở preview học (không redesign; khung đích đã vẽ theo từng candidate).
  - `app/pptx_parser.py`: `norm_key` fold ký tự Hán-Việt `進` → "tien" (caption "Sau cải進" trên báo cáo thật trước
    đây không được nhận là caption "Sau").
- Business logic giữ nguyên: PROMPT-014 (nguyên nhân cấu trúc), 015 (gate ngữ nghĩa), 020 (isolation theo báo
  cáo), 021 (crop đúng slide đã render), 023 (picker/config), 024 (huỷ an toàn + learning + commit Excel nguyên tử);
  Management Number, Vendor, ngày phát sinh, QPN, cause, improvement text, evidence, lock/retry Excel, force
  reprocess, Ollama, learning, updater.
- Test: `tests/test_prompt025_multi_item.py` (19): fixture đúng §30 (2 mục, 3 ảnh After của mục 1 + caption "Sau cải進" dưới, khoảng trắng, không khung chữ bao ngoài) + biến thể heading kiểu "Lỗi ... sau ép nhựa" (ca báo cáo
  Windows), heading giữa khung không in đậm, slide 3 mục, slide 1 mục, 2 báo cáo giống hệt nhau, force reprocess,
  identity học theo report+slide+item, render một lần cho 2 region, hook huỷ hợp tác, mapping Excel (1 dòng, đủ 2
  mục đúng thứ tự, 2 nhóm ảnh), dừng sau file hiện tại vẫn nguyên khối, loại xử lý tạm thời. Tổng **894 passed**
  (baseline 875 + 19 mới); compileall, pyflakes (file mới), `git diff --check`, frontend `tsc --noEmit` + `vite build`
  đều pass.
- Chưa chạy xác minh Windows: cần chạy lại file PPTX 2 mục thật trên Windows (xem checklist trong báo cáo).

## PROMPT-027 — Render trung thực PowerPoint + xem trước đúng vùng After thật + sửa No Fill/No Line (v1.3.3 / Build 016)

Bối cảnh: chạy thử PPTX thật trên Windows lộ ra nhiều lỗi KHÁC nhau (không cùng một nguyên nhân), nên mỗi vấn đề
được truy vết độc lập trước khi sửa (§58): xem trước toàn slide bị dồn lên trên; khung chữ tác giả để **No Fill /
No Line** bị vẽ thành khung đen; khung học chỉ tô đúng MỘT ảnh trong khi Excel xuất cả vùng Sau của Mục; và cần
kiểm chứng crop thật.

- Nguyên nhân gốc (đã chứng minh bằng fixture, không suy đoán):
  - **Khung đen (Problem B)** — `app/pptx_parser.py` đọc `line_visible = line.fill.type is not None`. Với
    `<a:ln><a:noFill/></a:ln>`, python-pptx trả về `MSO_FILL.BACKGROUND` (giá trị 5, **khác None**) nên "No Line"
    bị hiểu nhầm là "có nét"; vì không có màu tường minh, renderer fallback `line_color or "#000000"` → vẽ khung
    đen. Nhánh fill đã có phép thử `!= MSO_FILL.BACKGROUND`, nhánh line thì thiếu → bất đối xứng. Đây đúng là điều
    §7 cấm: suy diễn "thiếu màu" thành "nét đen nhìn thấy được".
  - **Nội dung bị dồn lên trên (Problem A)** — `Block` hoàn toàn KHÔNG lưu `anchor` / text inset / line spacing,
    còn `_draw_text` luôn bắt đầu vẽ ở `y0 + 3` (đỉnh khung). Slide thật có chữ căn giữa (`anchor="ctr"`) và nhiều
    khoảng trắng tác giả chủ ý nên PowerPoint vẽ THẤP HƠN; renderer tích hợp vẽ sát đỉnh → lệch lên. Đo trên
    fixture: chữ bị vẽ cao hơn tâm thật 82px. Không dùng bất kỳ offset cố định nào (§12).
- `app/pptx_parser.py`:
  - Thêm `ShapeStyle` + `_shape_style()`: phân biệt rành mạch **No Fill / No Line tường minh** (`fill_explicit_none`,
    `line_explicit_none`) với **kế thừa / không khai báo / theme**; `BACKGROUND` = "không vẽ gì" cho CẢ hai nhánh
    fill và line; chỉ khi tác giả không khai báo gì mới fallback `fillRef` / `lnRef`.
  - Thêm `_text_frame_layout()` và các trường mới trên `Block`: `vertical_anchor`, `wrap`, 4 inset, `autofit_scale`
    (`<a:normAutofit fontScale>`), `line_spacing` / `space_before` / `space_after` theo từng đoạn. Đây là gợi ý
    render, KHÔNG phải viết lại font engine của PowerPoint (§11).
- `app/qpn_renderer.py`:
  - `_draw_visual_shape()` chỉ vẽ đúng slot được phép: `fill_visible=False` → không tô; `line_visible=False` →
    không viền; cả hai false → không vẽ gì. Màu mặc định chỉ là hằng số đặt tên (`DEFAULT_SHAPE_FILL` /
    `DEFAULT_SHAPE_LINE`) áp dụng CHO SLOT THẬT SỰ HIỂN THỊ nhưng không phân giải được màu. Connector/line: chỉ
    `line_explicit_none` mới ẩn (nét là bản thân hình đó).
  - `_draw_text()` tính vị trí chữ theo inset thật + `vertical_anchor` (t/ctr/b) + line/paragraph spacing +
    `normAutofit`; shrink-to-fit có chặn dưới, suy ra từ chính khung đã trừ inset. Màu chữ lấy theo CHỈ SỐ ĐOẠN
    (trước đây lấy theo chỉ số dòng đã wrap → lệch màu khi một đoạn xuống dòng).
  - Chẩn đoán có cấu trúc (§3): `SLIDE_RENDER purpose=… backend=… faithful=… slide=… rendered=WxH
    slide_emu=WxH schema=v…`, `SLIDE_RENDER_BACKEND_FAILED purpose=… backend=… reason=…` (kể cả
    `reason=unavailable` — không rơi xuống âm thầm), `SLIDE_RENDER_FAILED` khi cả chuỗi thất bại. `_safe_reason()`
    loại đường dẫn tuyệt đối khỏi lý do lỗi.
  - `FAITHFUL_BACKENDS = {"powerpoint"}` + `is_faithful_backend()`; `RENDER_SCHEMA_VERSION = 2` để invalidate cache
    preview cũ sau khi sửa hình học render (§29). `render(..., purpose=…)` là tham số bổ sung, không đổi hành vi.
  - PowerPoint COM vẫn là backend ưu tiên và export NGUYÊN slide (`width_px` × `height` theo đúng tỉ lệ slide):
    không trim khoảng trắng, không crop theo pixel khác trắng, không đổi tỉ lệ (§5). LibreOffice / built-in vẫn là
    fallback hợp lệ nhưng được báo rõ là giảm độ trung thực (§6).
- `app/application_service.py`:
  - Cache preview giờ định danh theo (report scope, đường dẫn, slide, mtime, **chuỗi backend ưu tiên**, **width_px**,
    **RENDER_SCHEMA_VERSION**) → renderer đổi là bỏ cache, không tái dùng ảnh sai cũ; vẫn cách ly tuyệt đối giữa các
    báo cáo (không tái diễn PROMPT-020).
  - `_slide_evidence_regions()` chạy ĐÚNG bộ dựng production (`select_after_pictures` +
    `build_improvement_visual_regions`) và `_evidence_region_for()` chọn region sở hữu ảnh ứng viên theo chính danh
    sách ảnh của region → React KHÔNG BAO GIỜ tự suy ra hình học từ `candidate.bounds` (§15). Region được tính
    TRƯỚC khi render nên preview lỗi vẫn báo đúng sự thật vùng bằng chứng.
  - DTO thêm (chỉ thêm, không bỏ trường cũ): `slidePreviewBackend`, `slidePreviewFaithful`, `evidenceRegionBbox`,
    `evidenceRegionBboxPct`, `evidenceRegionKind`, `evidenceRegionItemId/ItemIndex/ItemHeading/PictureCount`.
    Vẫn không lộ đường dẫn tuyệt đối ra browser.
- `frontend/src/types.ts`, `frontend/src/components/AppHeader.tsx`, `frontend/src/tabs/LearningTab.tsx`:
  - Header hiện **`v1.3.3 · Build 016`** (lấy từ Python qua bridge, không hardcode); Settings vẫn hiện
    **`1.3.3 — Build 016`**. Tester Windows nhận diện ngay bản đang chạy (§50/§52).
  - Preview học vẽ HAI khung phân biệt rõ (§13/§16): **Vùng xuất Excel** = viền liền đậm + nhãn
    `Mục #n · Vùng xuất Excel` (chính), **Ảnh đang đánh giá** = viền nét đứt mảnh + nhãn `Ảnh #id` (phụ). Vùng làm
    nét (clip) đi theo vùng xuất Excel. Toggle gọn 3 nút `[Vùng xuất Excel][Ảnh đang đánh giá][Cả hai]`, mặc định
    "Cả hai" (§17) — KHÔNG redesign lại layout Học cải tiến (§43).
  - Backend hiển thị dạng tên an toàn (`Render: PowerPoint` / `Trình vẽ tích hợp`); khi không phải PowerPoint thì có
    dòng chú ý nhỏ KHÔNG chặn: "Bản xem trước đơn giản — bố cục có thể khác PowerPoint" (§6/§41). Panel phải liệt kê
    toạ độ EMU + % của cả hai hình học để đối chiếu log Windows (§42).
  - Hook `overlayMode` khai báo CÙNG các hook khác và TRƯỚC early-return theo candidate → giữ nguyên fix React #310
    (§44). Làm mờ/làm nét chỉ là hiển thị React, không đổi ảnh nguồn, region, crop hay bytes Excel (§18).
- Phiên bản: `app/__init__.py` → `__version__ = "1.3.3"`, `BUILD_NUMBER = 16`; `frontend/package.json` +
  `package-lock.json` (chỉ 2 trường version gốc — KHÔNG đụng `fast-fifo@1.3.2`), `frontend/index.html`,
  `frontend/public/mock/qpn-slide.svg`, fixtures test frontend. Tài liệu lịch sử (`PROGRESS_LOG.md` các mục cũ,
  `DEVELOPMENT_WEBVIEW.md` dòng kết quả PROMPT-025) và `tests/test_console_encoding.py` (chuỗi mẫu test encoding)
  giữ nguyên đúng §49. `backup/ReportExtractor_v1.0.4_Build004_STABLE/` KHÔNG đổi (§63). KHÔNG xuất bản LAN /
  release / không đụng `main` (§62).
- Giữ nguyên nghiệp vụ: PROMPT-015 gate ngữ nghĩa (AFTER / PRODUCTION_IMPROVEMENT / confident owner /
  excel_output_eligible — nhãn học không vượt gate), PROMPT-020 isolation theo báo cáo, PROMPT-021 crop từ slide đã
  render, PROMPT-025 nhiều mục trên một slide (region theo item, 3 ảnh After = MỘT vùng), PROMPT-024R huỷ hợp tác
  (`should_cancel` vẫn được quan sát ở các điểm an toàn, `purpose` không đổi hành vi huỷ), PROMPT-026R bridge
  (ping/polling gate/reconnect), an toàn transaction Excel, PROMPT-011 hình học group.
- Test: `tests/test_prompt027_rendering.py` (54) + `frontend/tests/prompt027-overlays.test.mjs` (7). Tổng
  **965 passed** (baseline 911 + 54 mới, không giảm, không xoá test cũ); compileall, pyflakes (theo đúng lời gọi và
  bộ lọc của `tools/build_portable.py`), `git diff --check`, `npm run test:bridge` (14), `npm run
  test:learning-hooks` (1), `npm run typecheck`, `npm run build` đều pass.
- CHƯA xác minh trên Windows: độ trung thực PowerPoint thật, hành vi No Fill / No Line trên PowerPoint thật, và crop
  Excel cuối cùng trên báo cáo thật — chỉ được kết luận sau khi chạy PPTX thật trên Windows với v1.3.3 / Build 016.

## PROMPT-028 — Đóng gói Windows Portable cho ứng dụng React + pywebview (v1.3.3 / Build 016, không đổi version)

Yêu cầu: `ReportExtractor.exe` phải mở ĐÚNG ứng dụng mà `python -m app.desktop` đang mở, build frontend trước,
bundle `frontend/dist`, sửa cách tìm đường dẫn frontend khi chạy frozen, đóng gói đủ runtime pywebview, giữ
WebView2, PowerPoint vẫn là TÙY CHỌN, máy đích không cần Python/Node/npm/Git, và CHỈ build gói nghiệm thu
(`--no-publish`).

- Hiện trạng trước khi sửa (đã đọc code, không phỏng đoán):
  - `ReportExtractor.spec` trỏ entry `run.py` → `app.main.main()` → `launch_gui()` → **Tk cổ điển**; `datas`
    KHÔNG có `frontend/dist`; `hiddenimports` KHÔNG có `webview`/`pythonnet`/`clr`. Tức là exe build ra sẽ KHÔNG
    phải ứng dụng React.
  - `app/desktop.py` tìm bundle bằng `Path(__file__).parent.parent / "frontend/dist"` và hardcode
    `FRONTEND_URL = "../frontend/dist/index.html"` — đúng khi chạy source, sai khi frozen.
  - `requirements-build.txt` không kéo `requirements-webview.txt` → venv build KHÔNG có pywebview.
- Nguyên nhân gốc của "cửa sổ trắng" khi đóng gói (đã chứng minh bằng cách đọc source pywebview 6.2.1):
  `webview.util.get_app_root()` trả về `sys._MEIPASS` khi frozen, còn ở chế độ source trả về
  `dirname(realpath(sys.argv[0]))`; `abspath()` nối URL tương đối vào root đó, và `http.start_server()` đặt
  `server.root_path = abspath(dirname(commonpath(urls)))`. Vậy CÙNG một bundle phải được gọi bằng
  `frontend/dist/index.html` khi nằm trong `_internal`, và `../frontend/dist/index.html` khi nằm cạnh
  `_internal` — hardcode một giá trị là sai một trong hai trường hợp.
- `app/desktop.py`: thêm `assumed_app_root()` (xác định: `_MEIPASS` khi frozen, `<repo>/app` khi source — KHÔNG
  đọc `sys.argv[0]` nên không phụ thuộc pytest/IDE/`python -c`), `runtime_app_root()` (ưu tiên
  `webview.util.get_app_root()` thật), `frontend_dist_candidates()`, `frontend_dist_dir()`, `frontend_url()`,
  `frontend_index()`. URL được SUY RA từ thư mục dist đã phân giải (`relpath(dist, root)`), không hardcode theo
  từng chế độ; luôn là đường dẫn tương đối để pywebview phục vụ qua HTTP loopback (giữ nguyên lý do không dùng
  `file://`: ES module của Chromium). Khi frozen CHỈ tìm trong thư mục portable — không fallback về checkout
  nguồn, vì fallback sẽ che mất một gói bị hỏng. Log `WEBVIEW_FRONTEND packaged/app_root/dist/url` và cảnh báo
  `WEBVIEW_FRONTEND_URL_MISMATCH` nếu round-trip sai.
- `app/main.py`: thêm `launch_desktop()` (chạy `app.desktop.main()`), `launch_ui()` (mặc định React, tự lùi về
  Tk khi thiếu pywebview/bundle) và cờ `--legacy-gui`. `launch_gui()` giữ NGUYÊN (test cũ vẫn pass). Mọi lỗi
  khởi động vẫn ghi `logs/startup_error.log`, không bao giờ im lặng.
- `ReportExtractor.spec`: bundle `frontend/dist` → `_internal/frontend/dist`; THU THẬP `webview`
  (`collect_data_files` + `collect_dynamic_libs` + `collect_submodules`) vì `webview/js/*` chính là phần inject
  cầu nối JS↔Python và `webview/lib/*.dll` + `runtimes/win-x64/native/WebView2Loader.dll` là bộ interop
  WebView2; thêm `pythonnet`/`clr`/`clr_loader`/`bottle`/`proxy_tools`; loại trừ PyQt/PySide/gi/cefpython3 để
  luôn đi qua EdgeChromium. PowerPoint/pywin32 chỉ thu thập NẾU có trên máy build (`_module_available`) — máy
  đích không có PowerPoint vẫn chạy vì `qpn_renderer` import COM bên trong hàm và fallback. Build DỪNG NGAY nếu
  thiếu `frontend/dist/index.html` hoặc `assets/` (không để tester nhận một exe cửa sổ trắng).
- `requirements-build.txt`: thêm `-r requirements-webview.txt` (nếu không thì venv build thiếu pywebview và
  spec không thể bundle WebView2).
- `tools/build_portable.py`: thêm `find_npm()`, `build_frontend()` (bước 5, chạy `npm run build` TRƯỚC
  PyInstaller, tự `npm install` khi thiếu `node_modules`, kiểm tra `assets/*.js`), `copy_frontend_into_portable()`
  (bước 8, đặt bundle cạnh `_internal` theo đúng layout Portable), `_same_tree()` và `validate_frontend()`
  (bước 10): yêu cầu bundle ở CẢ HAI vị trí và GIỐNG NHAU từng byte, `index.html` tham chiếu `/assets/`, không
  rò đường dẫn máy dev, có `webview/js/api.js` + `js/lib/dom_json.js` + `js/state.js`, và (chỉ trên Windows) có
  3 assembly WebView2. Thêm `--skip-frontend`. Đánh số bước 1→13, publish vẫn là bước CUỐI.
- `build_portable.bat`: cảnh báo sớm nếu máy build thiếu npm (kèm gợi ý `--skip-frontend`), vẫn chuyển `%*`
  nên `--no-publish` hoạt động. Không đụng `BUILD_AND_PUBLISH.bat`.
- `release_docs/README.txt` + `FIRST_RUN.txt`: viết lại theo UI React thật (tab "Danh sách báo cáo" / "Cài đặt"
  / "Học cải tiến", không còn tên tab Tk cũ), thêm mục "MÁY ĐÍCH CẦN GÌ" (không cần Python/Node/npm/Git; cần
  WebView2 Runtime; PowerPoint tùy chọn), mô tả thư mục `frontend\dist\`, cách chữa cửa sổ trắng
  (`WEBVIEW_FRONTEND` trong `logs\app.log`), `--legacy-gui` / `--version` / `--diag`, và nhắc kiểm tra phiên bản
  trên thanh tiêu đề TRƯỚC khi báo lỗi.
- `DEVELOPMENT_WEBVIEW.md`: thêm mục "Windows Portable acceptance build (no publishing)" với lệnh
  `build_portable.bat --no-publish` và danh sách những gì bước validate kiểm tra.
- Không đổi: `crop_box_px`, `improvement_visual.py`, toàn bộ sửa đổi renderer/parser của PROMPT-027, cầu nối
  PROMPT-026R, huỷ PROMPT-024R, an toàn Excel, `backup/…STABLE/`. KHÔNG xuất bản LAN, KHÔNG sửa `version.json`
  triển khai, KHÔNG chạy `BUILD_AND_PUBLISH.bat`, KHÔNG merge `main`.
- Test: `tests/test_prompt028_portable.py` (47) — phân giải bundle khi frozen ở cả hai layout, ưu tiên bản cạnh
  `_internal`, chế độ source không đổi, không phụ thuộc `sys.argv[0]`/thư mục hiện hành, gói frozen KHÔNG BAO
  GIỜ lấy bundle từ checkout nguồn, báo đủ đường dẫn đã tìm khi thiếu bundle, `launch_ui` ưu tiên React và chỉ
  lùi về Tk khi thật sự không chạy được, `main([])` → React / `--legacy-gui` → Tk / `--cli` không mở cửa sổ,
  chạy thật `desktop.main()` với pywebview giả (đúng title kèm version, URL tương đối round-trip, `js_api` là
  `BridgeService`, `gui="edgechromium"` trên win32, handler đóng cửa sổ hoãn lại khi service đang bận), kiểm tra
  nội dung spec/requirements/docs, và `validate_frontend` với gói thật. Tổng **1012 passed**
  (965 của PROMPT-027 + 47 mới). Đã chạy THẬT `build_frontend()` (npm run build) + lắp ráp + validate một thư
  mục portable đầy đủ trên Linux: `problems: NONE`, hai bản `frontend/dist` giống hệt nhau.
- GIỚI HẠN (quan trọng): sandbox là **Linux x86_64, không có PyInstaller, không có Wine**. PyInstaller KHÔNG
  cross-compile nên `ReportExtractor.exe`, ZIP thật và SHA256 của exe **không thể tạo ở đây** — phải chạy
  `build_portable.bat --no-publish` trên Windows. Ghi nhận thêm: `tests/test_ollama_discovery.py::
  test_15_cancel_stops_early_and_reports_partial` trượt 1 lần trong 4 lần chạy full-suite (chạy riêng luôn pass;
  4/4 lần chạy ở commit gốc cũng pass) — đây là race có sẵn của test đó (phụ thuộc lịch thread), KHÔNG thuộc
  phạm vi đã sửa và không bị thay đổi.

## PROMPT-028R — Gói Portable không khởi động được: `Failed to resolve Python.Runtime.Loader.Initialize` (v1.3.3 / Build 016, không đổi version)

Lỗi thật do nghiệm thu Windows báo về (bản 1.3.3 / Build 016, python 3.12.10, gói ở
`H:\ReportExtractor_v1.3.3_Portable\`): bấm đúp `ReportExtractor.exe` thì tiến trình chết TRƯỚC khi UI React
hiện ra. Chuỗi lỗi: `webview.platforms.winforms` → `import clr` → `pythonnet.load()` →
`clr_loader.get_netfx()` → `clr_loader/netfx.py:50` → `RuntimeError: Failed to resolve
Python.Runtime.Loader.Initialize from _internal\pythonnet\runtime\Python.Runtime.dll`.

- **Điểm mấu chốt về thông báo lỗi**: `netfx.py:50` ném lỗi này khi hàm native `pyclr_get_function()` trả về
  `NULL`. Đường dẫn trong thông báo là đường dẫn **pythonnet HỎI**, không phải đường dẫn **có tồn tại** — nên
  lỗi này tương thích với cả ba khả năng: file thiếu, file sai bản, hoặc file có đó nhưng CLR không nạp được.
  Vì vậy không được suy đoán; từng nguyên nhân phải đo trên mã nguồn đã cài.
- **NGUYÊN NHÂN ĐÃ LOẠI TRỪ (có bằng chứng thực thi, không phải giả thuyết)**: "DLL bị trùng", "sai category
  PyInstaller", "`collect_data_files` phá hook-clr", "`collect_dynamic_libs` gây nhiễu". `Analysis` kết thúc
  bằng `normalize_toc(self.datas + self.binaries)` (`building/build_main.py`), và `_TOC_TYPE_PRIORITIES` của
  `normalize_toc` cho `BINARY`/`EXTENSION` mức 1 còn `DATA` mức 0 → entry `datas` trùng đích bị LOẠI, entry
  `binaries` của hook-clr luôn thắng ở `pythonnet/runtime`. Đã chạy thử `normalize_toc` thật với cả hai thứ tự
  để chốt (`test_pyinstaller_normalize_toc_gives_binaries_priority_over_same_dest_datas`). Hai dòng
  `collect_data_files(...)`/`collect_dynamic_libs(...)` cũ vì thế là THỪA chứ không phải nguyên nhân; vẫn xoá
  để hook chính thức là chủ sở hữu duy nhất, nhờ đó cổng kiểm tra "đúng 1 bản, đúng chỗ, đúng SHA256" có nghĩa.
- **NGUYÊN NHÂN CÒN MỞ → đã dựng cổng kiểm soát để Windows tự kết luận**:
  1. `.venv-build` cũ / lệch phiên bản: `requirements-webview.txt` giờ ghim `clr-loader==0.2.10` (pythonnet
     3.0.5 chỉ khai báo `clr_loader<0.3.0,>=0.2.7`, không ai ghim nên venv tái sử dụng sẽ trôi phiên bản).
     `verify_runtime_imports()` import THẬT `webview`/`pythonnet`/`clr`/`clr_loader` trong venv build TRƯỚC khi
     đóng gói và kiểm tra clr-loader cài đặt có thoả yêu cầu của pythonnet; thêm cờ `--fresh-venv` để tạo lại
     `.venv-build` tất định. Trên Windows thiếu gói nào là DỪNG BUILD, không để tới lúc bấm đúp mới biết.
  2. hook-clr gom DLL SAI THƯ MỤC: khi `importlib.metadata.files('pythonnet')` không trả về đúng 1 kết quả
     (dist-info cũ/sửa tay trong venv tái sử dụng), hook rơi xuống
     `ctypes.util.find_library('Python.Runtime')` và gom với đích `'.'` → DLL nằm ở
     `_internal/Python.Runtime.dll`, LỆCH MỘT THƯ MỤC so với chỗ pythonnet tìm. Đã chạy thật hook với metadata
     rỗng để chứng minh (`test_hook_clr_legacy_fallback_would_collect_the_dll_one_directory_too_high`);
     `validate_pythonnet_runtime()` bắt đúng layout này ("sai vị trí").
  3. DLL sai bản/sửa byte: `validate_pythonnet_runtime()` yêu cầu ĐÚNG 1 `Python.Runtime.dll` trong cả gói, ở
     đúng `_internal/pythonnet/runtime/`, GIỐNG HỆT (SHA256) bản trong venv build, và in cả hai đường dẫn +
     size + SHA256 ra log build; trên Windows còn kiểm `clr_loader/ffi/dlls/<arch>/ClrLoader.dll`.
  4. Thiếu .NET Framework trên máy đích — **yêu cầu thật, nay được phát hiện rõ**: `Python.Runtime.dll` của
     pythonnet 3.0.5 biên dịch cho `.NETStandard,Version=v2.0` (theo chính `deps.json` của nó), nên cần facade
     netstandard của **.NET Framework 4.7.2+** (registry `Release >= 461808`). `dotnet_framework_report()` đọc
     `HKLM\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full` và ghi vào log; `explain_clr_failure()` nêu đích
     danh yêu cầu này. Đã bổ sung vào `release_docs/README.txt` mục 1 và mục 5.
  5. `clr_loader` nạp netfx kiểu lười: `clr_loader/__init__.py` KHÔNG import netfx ở cấp module — thân
     `get_netfx()` mở đầu bằng `from .netfx import NetFx`. spec nay khai báo tường minh
     `clr_loader.ffi/netfx/types/util` + `cffi` thay vì trông vào phân tích bytecode.
- **Chẩn đoán lúc chạy (`app/desktop.py`)**: `runtime_diagnostics()` ghi MỘT dòng `WEBVIEW_RUNTIME key=value …`
  vào `logs/app.log` TRƯỚC khi tạo cửa sổ, gồm `frozen`, `python`, `platform`, `pywebview`, `pythonnet`,
  `clr_loader`, `backend`, `python_runtime_dll`, `python_runtime_dll_exists`, `dotnet_runtime`,
  `dotnet_framework` — nên exe chết ngay vẫn tự giải thích được. Khi probe `import clr` thất bại thì ghi thêm
  `WEBVIEW_DOTNET_PROBLEM … | HINT:`; `webview_start()` ném lại kèm hint. `pythonnet_runtime_dll()` tính đúng
  công thức của pythonnet (`Path(__file__).parent/"runtime"/"Python.Runtime.dll"`). Không đưa đường dẫn nhạy
  cảm của người dùng lên UI React; `startup_error.log` vẫn giữ nguyên vẹn exception.
- **pywebview**: giữ nguyên backend Microsoft Edge WebView2 (React + pywebview vẫn là mặc định, KHÔNG lùi về
  Tkinter để che lỗi). Chỉ giữ đúng MỘT thao tác gom tay: `collect_data_files("webview", subdir="js")` —
  hook-webview dùng `subdir='lib'` nên không bao giờ gom cầu nối JS↔Python (`api.js`/`dom_json.js`/`state.js`,
  6 file). Cố ý giới hạn ở `js` để không chồng lấn `webview/lib` của hook. Thêm `webview.platforms.winforms`
  vào hiddenimports và loại các backend `qt/gtk/cocoa/android/mshtml`.
- **Phiên bản phụ thuộc**: TRƯỚC — `pywebview==6.2.1`, `pythonnet==3.0.5 ; win32`, clr-loader KHÔNG ghim,
  `pyinstaller>=6.6,<7`. SAU — thêm `clr-loader==0.2.10 ; sys_platform == "win32"` (không nâng cấp gì khác;
  thay đổi duy nhất này có test hồi quy `test_requirements_pin_clr_loader_inside_pythonnets_declared_range`).
  Môi trường đã đo ở đây: python 3.11.2, pywebview 6.2.1, pythonnet 3.0.5, clr-loader 0.2.10,
  pyinstaller 6.22.3, pyinstaller-hooks-contrib 2026.8.
- **`Python.Runtime.dll` nguồn (đo thật trong venv)**: `…/site-packages/pythonnet/runtime/Python.Runtime.dll`,
  size **450048**, SHA256 **`d204ad74dc18cd07320c8e665bd32ec6549b555ce97e61e4d3cf88437a64988e`**, target
  `.NETStandard,Version=v2.0`. SHA256 của bản ĐÓNG GÓI phải in ra từ `validate_pythonnet_runtime()` khi build
  trên Windows và phải TRÙNG giá trị này — chưa có vì chưa build được ở đây.
- File đã sửa: `ReportExtractor.spec`, `app/desktop.py`, `tools/build_portable.py`, `requirements-webview.txt`,
  `tests/test_prompt028_portable.py`, `tests/test_publish_update.py`, `release_docs/README.txt`,
  `release_docs/FIRST_RUN.txt`, `DEVELOPMENT_WEBVIEW.md`.
- KHÔNG đổi: version (vẫn 1.3.3 / Build 016 — PROMPT-027R sở hữu 1.3.4), toàn bộ logic vùng bằng chứng của
  PROMPT-027R (`improvement_visual.py`, clustering, `item_span`), `backup/…STABLE/`. KHÔNG merge PR #7, KHÔNG
  xuất bản LAN, KHÔNG sửa `version.json` triển khai, KHÔNG chạy `BUILD_AND_PUBLISH.bat`.
- Test: `tests/test_prompt028_portable.py` 47 → **80**; toàn bộ **1044 passed** (1012 + 32 mới/mở rộng),
  `compileall` + `pyflakes` sạch, `git diff --check` sạch, frontend `test:frontend` **22/22** và
  `tsc --noEmit` sạch. `tests/test_publish_update.py` được bổ sung stub `subprocess.run` cho bước 3 mới (trả
  JSON khi được gọi với `-c`) — giữ nguyên ý nghĩa kiểm tra "không publish sau khi cổng test thất bại".
- **GIỚI HẠN (quan trọng, phải nói rõ)**: sandbox là **Linux x86_64**. PyInstaller KHÔNG cross-compile nên ở
  đây **không tạo và không chạy được `ReportExtractor.exe`**. Đã cài PyInstaller 6.22.3 trên Linux chỉ để ĐỌC
  và CHẠY THẬT các hook chính thức; đã chạy spec tới hết pha gom (log: `.NET runtime module registered …
  clr/clr_loader/pythonnet`, `webview JS bridge files collected: 6`) nhưng pha Analysis dừng vì thiếu
  `libpython3.11.so.1.0` của Python hệ thống. Vì vậy nghiệm thu bắt buộc trên Windows —
  `build_portable.bat --no-publish` (không dùng `--skip-tests`), chạy
  `dist\ReportExtractor_v1.3.3_Portable\ReportExtractor.exe`, chép cả thư mục sang ổ/đường dẫn khác rồi chạy
  lại, kiểm cầu nối JS↔Python, Settings, "Học cải tiến", và cuối cùng là backend `powerpoint` với PPTX thật —
  **CHƯA THỰC HIỆN** và không được báo là đã đạt.

## PROMPT-027R — Vùng bằng chứng Sau thật, ranh giới mục tiếp theo và độ trung thực renderer (v1.3.4 / Build 017, nhánh phát triển)

Bối cảnh: slide 3 thật (hai mục) trên PR #7 (1.3.3 / 016) còn ba lỗi: (A) vùng xuất Excel của Mục #2 nhỏ hơn khối Sau
thật; (B) crop của Mục #1 chứa phần cuối của tiêu đề Mục #2 (thấy như một nhãn "Daoltech" lẻ — không có đối tượng nào
tên như vậy; không lọc chữ); (C) xem trước Learning khác PowerPoint (chữ đậm/tối hơn, xuống dòng và khoảng cách khác,
hiện "Bản xem trước đơn giản"). Không có file PPTX thật trong môi trường này: mọi nguyên nhân dưới đây được chứng minh
trên fixture tái tạo đúng bố cục đã báo cáo (log `REGION_*` + ảnh crop + điểm ảnh), và cần xác nhận lại trên Windows.

- Nguyên nhân gốc (mỗi cơ chế có test/log riêng; không tăng ngưỡng toàn cục, không lọc chữ, không dùng pixel):
  - **M1 — phân loại thời gian sai**: dòng `+ Sau:` sở hữu mọi ảnh BÊN DƯỚI nó (không giới hạn) và được xét trước
    mũi tên, nên ảnh Trước thứ hai (không caption) thành Sau. Sửa: dòng inline phải khớp phiếu mũi tên có cấu trúc;
    xung đột → ambiguous, chỉ được giải bằng hàng lân cận.
  - **M2 — caption gán vượt ranh giới mục**: caption "Sau cải tiến" của Mục #1 nhận ảnh của Mục #2 qua nhánh
    "column header" (cách 1,4") dù tiêu đề Mục #2 nằm giữa → Mục #2 mất ảnh (vùng nhỏ) và crop Mục #1 chứa ảnh Mục #2.
    Sửa: caption không gán qua khoảng có ảnh/caption/tiêu đề khác; owner của caption phải khớp span của mục chứa ảnh,
    nếu không → ambiguous (fail closed).
  - **M3 — phân cụm bằng ngưỡng toàn slide**: ngưỡng 8% chiều rộng slide tách một khối Sau liền hàng thành 3 vùng.
    Sửa: khoảng cách còn so với kích thước ảnh (1,0×) và chỉ nối khi không có vật cản cấu trúc trong khoảng hở;
    ảnh Sau có caption khác phía (trên/dưới cùng phía) là khối khác.
  - **M4 — đưa cả khối chữ bằng giao hộp**: khối thân bài của Mục #2 chạm vùng caption nên được đưa nguyên hộp vào crop
    (kéo theo ảnh Trước và chữ thân bài). Sửa: chữ phải nằm ≥ 60% trong vùng bằng chứng; hình ≥ 50% hoặc chồng ảnh thành viên.
  - **M5 — kẹp đáy chỉ theo ảnh**: điều kiện "ảnh chồng ngang tiêu đề" bỏ qua tiêu đề khi chỉ crop cuối (chú thích,
    dấu khoanh) chồng ngang tiêu đề → phần cuối tiêu đề lọt vào crop. Sửa: kẹp đáy theo mọi tiêu đề bên dưới mà crop
    cuối chạm ngang; thành viên luôn được giữ (xung đột được báo cáo `REGION_BOUNDARY_CONFLICT`).
  - **Thêm — nhãn ngắn trên/dưới ảnh bị coi là tiêu đề mục** (`CHECK`, `120 mm`): tạo mục giả, cắt span mục thật, làm
    sai chủ sở hữu caption. Sửa: nhánh tiêu đề heuristic (không dấu hai chấm, không mở đầu bằng tên lỗi) bỏ qua dòng nằm
    trên ảnh hoặc trong cột ảnh ngay dưới nó.
  - **C/D — xem trước**: `blur(2px) brightness(0.82) saturate(0.85)` áp lên TOÀN ảnh slide + lớp cắt sắc thứ hai → chữ
    xám bị đậm/tối và mờ. Hiện "Trình vẽ tích hợp" nghĩa là PowerPoint COM không được chọn trên máy tester; nguyên nhân
    cụ thể trên máy đó chưa xác định (log mới `SLIDE_RENDER_BACKEND_FAILED ... stage=` sẽ cho biết).

- Sửa:
  - `app/improvement_pictures.py`: `ItemRegion.heading_bounds` (dòng tiêu đề, không phải span); `structural_obstacles()`,
    `gap_is_clear()`; `_caption_distance`/`_claims`/`classify_picture` nhận obstacles; xung đột inline↔mũi tên → ambiguous;
    `_span_owner()` + kiểm tra owner caption ↔ span; `_line_on_visual()`/`_label_beneath_picture()` cho nhánh heuristic;
    log `PICTURE_DECISION` cho mọi ảnh (đủ scope/slide/item/shape/source_order/bbox/temporal/semantic/owner/confident/eligible/group).
  - `app/improvement_visual.py`: phân cụm có cấu trúc (`_cluster_with_reasons`: hàng/cột theo kích thước ảnh, khoảng hở
    trống, caption khác phía, caption mở khối mới, group ancestry chỉ nới khoảng cách); span theo đúng tiêu chí
    `slide_items`; kẹp đáy theo mọi tiêu đề/footer bên dưới theo x của crop cuối (`_clamp_bottom`); đối tượng: containment
    (chữ ≥ 60%, hình ≥ 50% hoặc đè ảnh), loại đối tượng thuộc span/tiêu đề mục khác; `ImprovementVisualRegion` thêm
    `span_top/span_bottom/boundary_source/boundary_item_id/boundary_top/raw_bbox/padding`.
  - `app/pptx_parser.py`: `Block.group_path` (thứ tự nhóm, theo tài liệu) + `parent_group_id`/`group_root_id`; hình học không đổi.
  - `app/qpn_renderer.py`: `powerpoint_probe()` (không khởi động PowerPoint): pywin32 → `CLSIDFromProgID` → HKCR ở view
    mặc định/64/32-bit; `render_with_powerpoint()` dùng `DispatchEx` (instance riêng), mở read-only không cửa sổ, chỉ đóng
    presentation đã mở và chỉ quit instance đã tạo, cân bằng CoInitialize/CoUninitialize; lỗi mang `stage`, `exception_type`,
    `hresult`, lý do không có đường dẫn; `SlideRenderer.availability()`, `degraded`; `RENDER_SCHEMA_VERSION = 3`.
  - `app/application_service.py`: khoá cache preview gồm kích thước + mtime_ns + khả dụng từng backend + định danh Learning
    (nhãn đã lưu + mô hình); preview xuống cấp (faithful backend lỗi) hết hạn sau `DEGRADED_PREVIEW_TTL_S = 15`;
    vùng bằng chứng tính bằng `select_with_learning` (đúng quyết định Excel); DTO thêm `evidenceRegionPictureIds`,
    `evidenceRegionId`, `evidenceRegionBlockIndex`, `evidenceRegionBoundarySource`, `evidenceRegionNextHeadingTop`,
    và `regionId/boundarySource/nextHeadingTop` trong `regions` (chỉ hiển thị/gỡ lỗi).
  - `frontend/src/tabs/LearningTab.tsx`: bỏ hoàn toàn `filter`, lớp cắt thứ hai và overlay làm mờ; còn đúng MỘT bitmap
    nguyên bản + hai khung viền từ DTO (`evidenceRegionBboxPct`, `targetBboxPct`); "Ảnh trong vùng" chỉ hiển thị.
  - `app/image_learning.py`: obstacles dùng chung để đặc trưng học khớp quyết định production.
  - Phiên bản: `app/__init__.py` → `1.3.4` / `BUILD_NUMBER = 17`; `frontend/package.json`, `package-lock.json` (chỉ hai
    trường gốc), `index.html`, `public/mock/qpn-slide.svg`, fixture test. Test cũ đổi chuỗi phiên bản; fixture "bản mới hơn"
    trong `test_updater.py` chuyển 1.3.5 / 18; nhãn bản từ xa `CUR+1/CUR+2` cập nhật theo build mới.
- Không đổi: packaging (`ReportExtractor.spec`, `build_portable.bat`, `tools/build_portable.py`), updater, PROMPT-011,
  015, 020, 021, 024R, 025, 026R, 027 (No Fill/No Line, text layout).

- Chẩn đoán mới (grep được): `ITEM_SPAN` (start_y/end_y/next_item_id/next_heading_y/boundary_source),
  `REGION_CLUSTERS` (joins), `REGION_MEMBER` (scope/slide/item/shape/source_order/bbox/temporal/semantic/owner/confident/
  eligible/group), `REGION_CLAMP`, `REGION_EVIDENCE` (picture_ids, region_bottom_y, next_heading_y), `REGION_BOUNDARY_CONFLICT`,
  `PICTURE_DECISION`, `SLIDE_RENDER purpose=… backend=… slide=… attempted_backend=… selected_backend=… faithful=…`,
  `SLIDE_RENDER_BACKEND_FAILED … stage=… exception_type=… hresult=… reason=…`, `SLIDE_RENDER_CLEANUP`.

- Test: `tests/test_prompt027r_regions.py` (20: CASE 1–9 + 10 quyết định người dùng, M1–M5, nhãn/callout, group, chẩn
  đoán, ràng buộc biên), `tests/test_prompt027r_renderer.py` (22: chọn backend/faithful, probe theo stage, COM giả với
  DispatchEx/cleanup/CoInitialize, cache theo khả dụng, TTL xuống cấp), `frontend/tests/prompt027r-regression.test.mjs` (7,
  nguồn: không filter/clip, overlay từ DTO, pictureIds chỉ hiển thị, hook trước early return). `prompt027-overlays` và
  `learning-hooks` đổi kỳ vọng từ hai bitmap sang một bitmap nguyên bản.
- Kết quả: Python **1007 passed** (baseline 965, +42, không xoá test); `test:bridge` 14, `test:learning-hooks` 1,
  `test:prompt027` 7, `test:prompt027r` 7, `typecheck`, `build` đều pass; `compileall app` pass; pyflakes: cùng tập
  thông báo với baseline (chỉ các dòng "imported but unused" của probe, đã được bộ lọc build chấp nhận); `git diff --check` pass.
- CHƯA xác minh trên Windows (bắt buộc trước khi kết luận): PowerPoint COM thật (`DispatchEx`, từng stage), so sánh
  A/B/C (PowerPoint UI / PNG COM / Learning), PPTX thật (vùng Mục #1 chứa đủ ảnh Sau và kết thúc trên tiêu đề Mục #2;
  Mục #2 độc lập; crop Excel), và máy không có PowerPoint/LibreOffice với bản đóng gói (probe `import`/`availability`).

## PROMPT-029 — Tích hợp PROMPT-027R vào PROMPT-028R (v1.3.4 / Build 017)

Yêu cầu: tích hợp ĐÚNG commit `8f979eb4a520363b4967d9b424e5456d78f5c9f6` (nhánh `arena/6e89ee9d-tnp`) lên
HEAD `78805faad9bad464abff0f2ed3e46c18c7573bfc`. KHÔNG reset về `db92966`, KHÔNG bỏ PROMPT-028R, KHÔNG viết
lại PROMPT-027R từ đầu.

- **Cách tích hợp**: `git merge --no-ff origin/arena/6e89ee9d-tnp`. `merge-base` của hai phía đúng bằng
  `b5f8a29` (PROMPT-027) và PROMPT-027R là MỘT commit duy nhất trên đó, nên đây là tích hợp hai phía sạch.
  Merge giữ `8f979eb` làm ancestor thật (không tạo SHA mới như cherry-pick), nên có thể kiểm chứng bằng
  `git merge-base --is-ancestor`. Kết quả: commit merge `3dae422`; cả `8f979eb` và `78805fa` đều là ancestor
  của HEAD.
- **Phiên bản chuẩn sau tích hợp: v1.3.4 / Build 017 / VERSION_LABEL "1.3.4 — Build 017"**, và
  `frontend/package.json` cũng là 1.3.4. PROMPT-028R vẫn báo 1.3.3 / Build 016 vì nó được viết TRƯỚC khi tích
  hợp — đúng như đề bài nói. Các commit đóng gói KHÔNG hề sửa `app/__init__.py`, nên bản nâng version chỉ có
  thể đến từ PROMPT-027R; KHÔNG có chỗ nào bị kéo lùi về 1.3.3/016 trong lúc giải quyết xung đột.
- **Xung đột: 3 trên 45 đường dẫn** (hai phía cùng sửa):
  - `PROGRESS_LOG.md` — cả hai append cùng một điểm neo (sau entry PROMPT-027). Giải quyết: giữ NGUYÊN VĂN CẢ
    HAI khối, theo thứ tự thời điểm viết (PROMPT-028 → PROMPT-028R → PROMPT-027R), ngăn bằng một dòng trống.
    Không viết lại entry lịch sử nào, nên mỗi entry vẫn ghi đúng phiên bản chuẩn tại thời điểm nó được viết.
  - `DEVELOPMENT_WEBVIEW.md` — git tự merge. Giữ mục điều tra PROMPT-028R + dòng nghiệm thu của nó, ĐỒNG THỜI
    giữ hai dòng nghiệm thu PROMPT-027R và việc 027R đổi "1.3.3 — Build 016" thành "1.3.4 — Build 017" trong
    danh sách kiểm tra sau vá.
  - `tests/test_publish_update.py` — git tự merge vì hai phía sửa hai vùng rời nhau: 027R đổi
    `test_build_010_identity` sang 1.3.4/17/017; 028R làm các assertion thứ tự bước chịu được việc đánh số lại
    và bổ sung stub `subprocess.run` cho truy vấn môi trường build mới ở bước 3.
- **Kiểm chứng tích hợp (đo, không phỏng đoán)**: 32/32 đường dẫn chỉ PROMPT-027R sửa đều GIỐNG HỆT `8f979eb`;
  9/10 đường dẫn chỉ PROMPT-028R sửa đều GIỐNG HỆT `78805fa`. Lệch duy nhất là
  `tests/test_prompt028_portable.py` — do chủ động sửa ở commit này (xem dưới).
- **`tests/test_prompt028_portable.py` phải sửa 3 chỗ vì tích hợp làm tiền đề của chúng hết hiệu lực** (vẫn 80
  test, không xoá test nào):
  - `test_version_was_not_bumped_by_this_packaging_fix` (khẳng định 1.3.3/16/016) → thay bằng
    `test_canonical_version_after_prompt029_integration`: khẳng định 1.3.4/17/017, `BUILD_LABEL == "Build 017"`,
    `VERSION_LABEL == "1.3.4 — Build 017"` và `frontend/package.json` đồng ý với backend.
  - Scope guard cũ diff working-tree với `b5f8a29`, giờ dĩ nhiên liệt kê các file vùng bằng chứng. Đổi thành
    diff DẢI CỐ ĐỊNH `b5f8a29..78805fa` (`test_packaging_commits_never_touched_the_evidence_region_logic`),
    nên vẫn chứng minh được các commit đóng gói không đụng logic đó, và nay còn khẳng định chúng không đụng
    `app/__init__.py`.
  - `test_version_and_evidence_region_scope_…` → thay bằng `test_prompt029_kept_both_lines`: khẳng định cả hai
    commit là ancestor của HEAD, và CHẤT của mỗi phía còn nguyên — phía 028R: spec chỉ để hook chính thức sở
    hữu (không còn `collect_dynamic_libs`, chỉ `collect_data_files("webview", subdir="js")`, có `clr`/
    `clr_loader`), ghim `clr-loader`, `validate_pythonnet_runtime` / `verify_runtime_imports` /
    `report_build_dependencies` / `source_pythonnet_dll`, chẩn đoán trong `app/desktop.py`, `--fresh-venv`;
    phía 027R: log `REGION_EVIDENCE` / `REGION_BOUNDARY_CONFLICT` / `REGION_MEMBER` / `boundary_source`, các hàm
    `_item_span` / `_clamp_bottom` / `_cluster_with_reasons` / `_assign_captions_to_clusters`, và
    `evidenceRegionPictureIds` trong DTO của `app/application_service.py` lẫn `frontend/src/types.ts`.
- **`frontend/dist` được build lại** (027R có sửa `frontend/index.html`): nay là `index-B5WzBAyA.js` và
  `index-BTW1TSfn.css`.
- **Artifact sẽ tạo ra khi build trên Windows** (đã đối chiếu với code, không phải chép lại đề bài):
  `ReportExtractor.spec` đặt `PORTABLE_NAME = f"ReportExtractor_v{VERSION}_Portable"` → `dist\ReportExtractor_v1.3.4_Portable\`
  với `ReportExtractor.exe` bên trong; `tools/build_portable.py` đặt tên ZIP `ReportExtractor_<version>.zip` →
  `release\ReportExtractor_1.3.4.zip`. Cả hai khớp đúng artifact PROMPT-029 mong đợi. Chỉ build bằng
  `build_portable.bat --no-publish`.
- **Test**: Python **1087 passed** (1045 ở HEAD PROMPT-028R + 42 của PROMPT-027R: 20 trong
  `tests/test_prompt027r_regions.py`, 22 trong `tests/test_prompt027r_renderer.py`; KHÔNG xoá test nào).
  `tests/test_prompt028_portable.py` vẫn 80. Frontend **29/29** (22 + 7 test mới
  `tests/prompt027r-regression.test.mjs`, và `test:frontend` đã gồm cả 4 file). `tsc -b && vite build` pass,
  `tsc --noEmit` sạch, `compileall` sạch, `pyflakes app tools run.py` (đúng phạm vi cổng build) sạch,
  `git diff --check` sạch, không còn marker xung đột ở bất kỳ đâu. Ghi chú: `pyflakes tests` có vài cảnh báo
  "local variable assigned but never used" / "redefinition of unused" ở `tests/test_image_learning.py`,
  `tests/test_prompt004.py`, `tests/test_scan_list_compact.py` — có sẵn từ trước, KHÔNG thuộc phạm vi cổng
  build và không phải do tích hợp này sinh ra, nên không đụng tới.
- **GIỚI HẠN (quan trọng, phải nói rõ)**: sandbox vẫn là **Linux x86_64**. PyInstaller KHÔNG cross-compile nên
  ở đây **không tạo được `dist\ReportExtractor_v1.3.4_Portable\`, không tạo được
  `release\ReportExtractor_1.3.4.zip`, và không chạy được `ReportExtractor.exe`**. Toàn bộ nghiệm thu Windows
  bắt buộc của PROMPT-028R (chạy `build_portable.bat --no-publish` không dùng `--skip-tests`, khởi động exe từ
  `dist\`, chép cả thư mục sang ổ/đường dẫn khác rồi khởi động lại, kiểm cầu nối JS↔Python / Settings /
  "Học cải tiến") VÀ của PROMPT-027R (PPTX thật với PowerPoint, `SLIDE_RENDER purpose=… backend=powerpoint
  faithful=true`, `REGION_EVIDENCE` của Mục #1 kết thúc trên tiêu đề Mục #2, Mục #2 độc lập, crop Excel) đều
  **CHƯA THỰC HIỆN** và không được báo là đã đạt.
- KHÔNG đổi: `backup/…STABLE/`. KHÔNG merge PR #7 vào `main`, KHÔNG xuất bản LAN, KHÔNG sửa `version.json`
  triển khai, KHÔNG chạy `BUILD_AND_PUBLISH.bat`.

## PROMPT-030 — Gói Portable chết khi khởi động: `Failed to resolve Python.Runtime.Loader.Initialize` (v1.3.4 / Build 017, không đổi version)

Bối cảnh: trên Windows thật, pytest xanh hết, `build_portable.bat --no-publish` build xong, nhưng
`H:\ReportExtractor_v1.3.4_Portable\ReportExtractor.exe` chết ngay khi khởi động với
`RuntimeError: Failed to resolve Python.Runtime.Loader.Initialize from H:\...\_internal\pythonnet\runtime\Python.Runtime.dll`
trước khi UI hiện ra.

- **Thông báo lỗi đó KHÔNG chứa nguyên nhân.** `clr_loader/ffi/netfx.py` chỉ khai báo đúng năm hàm native
  (`pyclr_initialize`, `pyclr_create_appdomain`, `pyclr_get_function`, `pyclr_close_appdomain`,
  `pyclr_finalize`) và **không có hàm lấy lỗi** (khác coreclr/hostfxr có `util/coreclr_errors.py`), nên
  `clr_loader/netfx.py::_get_callable` gộp MỌI thất bại của `pyclr_get_function` — file thiếu, file bị
  Windows chặn, .NET Framework quá cũ, sai kiến trúc, payload bị đổi tên, chữ ký không bind được delegate —
  vào đúng một chuỗi đó. Nó in ra đường dẫn pythonnet **đã hỏi**, không phải đường dẫn tồn tại.
  `NetFx.__init__` cũng không kiểm tra `pyclr_create_appdomain` có trả về NULL hay không, còn
  `NetFx.info()` hardcode `initialized=True` và `version="<undefined>"`, nên `pythonnet.get_runtime_info()`
  không thể phân biệt "đã tạo AppDomain" với "AppDomain NULL". Câu hint của PROMPT-028R khẳng định
  "clr_loader created a .NET Framework app domain" — điều không thể chứng minh được từ chuỗi lỗi đó, và chính
  nó đã hướng việc điều tra sang lỗi đóng gói mà cổng build đã loại trừ rồi.
- **Nhóm nguyên nhân đóng gói (A) bị LOẠI BẰNG BẰNG CHỨNG, không phải bằng suy đoán.** Cổng build
  PROMPT-028R (`validate_pythonnet_runtime`, bước 10) có `fail()` và build này ĐÃ HOÀN TẤT, nghĩa là gói đã
  được chứng minh: đúng MỘT bản `Python.Runtime.dll`, nằm đúng `_internal/pythonnet/runtime/`, SHA256
  **byte-identical** với bản trong `.venv-build`, và có `ClrLoader.dll` đúng kiến trúc. Bước 3
  (`verify_runtime_imports`) cũng đã `import clr` **thành công trên chính máy đó** trước khi freeze — tức là
  cùng bytes DLL, cùng máy, cùng runtime netfx thì phân giải được. Vậy assembly và máy đều tốt; phần sai
  nằm ở môi trường tiến trình đã freeze.
- **`tools/dotnet_pe.py` (MỚI)** — bộ đọc PE/ECMA-335 thuần stdlib, chạy trên mọi nền tảng, đọc thẳng
  metadata quản lý ra để chứng minh thay vì giả định. Kết quả ĐO trên đúng hai DLL mà bản build này đóng gói:
  `Python.Runtime.dll` = `Python.Runtime, Version=3.0.5.0, Culture=neutral, PublicKeyToken=5000fea6cba702dd`,
  size 450048, SHA256 `d204ad74dc18cd07320c8e665bd32ec6549b555ce97e61e4d3cf88437a64988e`, **IL-only /
  any-cpu** (nên không thể là xung đột kiến trúc), targets `.NETStandard,Version=v2.0`, tham chiếu
  `netstandard 2.0.0.0` + `System.Reflection.Emit(.ILGeneration) 4.0.0.0`, 241 type, và
  **`Python.Runtime.Loader.Initialize` là `static int32 (native int, int32)`** — đúng typedef `entry_point`
  mà clr_loader bind, nên lệch phiên bản pythonnet↔clr_loader cũng bị loại.
  `clr_loader/ffi/dlls/{amd64,x86}/ClrLoader.dll` = `ClrLoader 1.0.0.0`, **mixed-mode C++/CLI**
  (không IL-only), targets `.NETFramework,Version=v4.7.2`, export đúng năm symbol `pyclr_*`.
- **Cổng build nay chứng minh LOADABILITY chứ không chỉ presence/hash** (`tools/build_portable.py`):
  `prove_runtime_identity` (danh tính assembly, IL-only/any-cpu, target framework → suy ra .NET Framework
  `Release` tối thiểu và so với registry máy build, type + method + **chữ ký** entry point),
  `prove_clr_loader` (đúng kiến trúc, mixed-mode, đủ năm export), `mark_of_the_web` (stream
  `Zone.Identifier`), và `probe_packaged_runtime` — chạy **thật** chuỗi đã hỏng
  (`ffi.dlopen` ClrLoader.dll **trong gói** → `pyclr_initialize` → `pyclr_create_appdomain` → kiểm tra NULL →
  `pyclr_get_function` trên `Python.Runtime.dll` **trong gói**) trong một tiến trình con, có timeout, và fail
  build nếu không phân giải được. Thêm `--validate-only <FOLDER>` để chạy đúng cổng đó trên một thư mục
  Portable **đã build xong** (ví dụ `H:\ReportExtractor_v1.3.4_Portable`) mà không phải build lại.
- **Điều kiện duy nhất khớp MỌY dữ kiện và xuất hiện SAU build**: assembly bị Windows chặn.
  `Zone.Identifier` được gắn khi giải nén `release\ReportExtractor_1.3.4.zip` hoặc chép từ máy/USB khác;
  `LoadLibrary` **bỏ qua** cờ đó nên `ClrLoader.dll` native vẫn nạp và AppDomain vẫn được tạo, còn
  `Assembly.LoadFrom` của .NET Framework thì **từ chối** (`COR_E_FILELOAD`, HRESULT 0x80131515 "Operation is
  not supported") — đúng nghĩa một file tồn tại, đúng đường dẫn, đúng SHA256 mà vẫn không initialize được,
  và đúng lý do cùng bytes đó chạy được từ `.venv-build` (pip không bao giờ gắn stream). Đây là **giả thuyết
  có cơ chế khớp bằng chứng**, CHƯA được chứng minh trên Windows thật; cổng `--validate-only` và log khởi động
  mới là thứ xác nhận nó.
- **`app/desktop.py`**: thêm `mark_of_the_web`/`_read_zone_identifier`/`_remove_zone_identifier`/
  `unblock_packaged_runtime` (chỉ chạy khi `sys.platform == "win32"` VÀ `is_packaged()`, chỉ đụng hai file
  runtime dưới `sys._MEIPASS`, gỡ stream đúng như tick "Unblock" của Windows, mọi lỗi đều suy giảm thành
  báo cáo chứ không làm app không khởi động được), `packaged_runtime_files`, `appdomain_report` (đọc thẳng
  handle thay vì tin `info()`), các field log `blocked_runtime` / `unblocked_runtime` / `unblock_failed` /
  `runtime_files_checked` / `dotnet_appdomain`, cảnh báo `WEBVIEW_DOTNET_BLOCKED`, và viết lại hint
  `_CLR_HINTS` để chỉ vào đúng các field phân biệt được nguyên nhân. KHÔNG đổi `pythonnet.load()`, KHÔNG ép
  runtime khác: `clr_loader.get_coreclr()` cần `find_dotnet_root()` + một runtime `Microsoft.NETCore.App` đã
  cài, mà máy đích Portable thì không được phép cần .NET — nên nhánh `except → PYTHONNET_RUNTIME='coreclr'`
  trong `webview/platforms/winforms.py` không bao giờ cứu được app (và vì `pythonnet.load()` để lại
  `_RUNTIME` đã set sau lần thất bại, nhánh đó còn dùng lại đúng runtime netfx cũ và ném lại đúng lỗi cũ).
- **Test**: `tests/dotnet_image_fixtures.py` (MỚI) **tự sinh ảnh PE managed thật** theo ECMA-335 — không cần
  Windows, không cần .NET, không cần PyInstaller, không phụ thuộc gói đã cài — để mutate đúng MỘT thuộc tính
  mỗi lần. `tests/test_prompt030_runtime.py` (MỚI, 58 test). Fixture PROMPT-028R cũ dùng chuỗi 16 byte
  `b"MANAGED-ASSEMBLY"` nên bị cổng mới từ chối đúng: **sửa fixture, KHÔNG nới cổng** (nay ghi ảnh managed
  thật; ca "build cũ" dùng `version=(3,0,4,0)` để cô lập đúng việc so SHA256). Python
  **1104 → 1162 passed**, 0 skip, 0 xfail, không xoá test nào; `tests/test_prompt028_portable.py` +
  `tests/test_prompt027r_renderer.py` **110 passed, 0 skip**. `compileall` sạch, `pyflakes app tools run.py`
  sạch, `git diff --check` sạch. Frontend KHÔNG đụng tới và xanh: `test:bridge` 14/14, `test:frontend`
  29/29, `tsc --noEmit` sạch, bundle vẫn `index-B5WzBAyA.js` / `index-BTW1TSfn.css`.
- **GIỚI HẠN (phải nói rõ)**: sandbox là Linux x86_64, KHÔNG build được artifact PyInstaller nào, KHÔNG
  chạy được `ReportExtractor.exe`. Tiêu chí nghiệm thu cuối — exe mở lên và vào được UI React/pywebview,
  log không còn `Failed to resolve Python.Runtime.Loader.Initialize`, cầu nối Python hoạt động — **CHƯA THỰC
  HIỆN** và không được báo là đã đạt.
- KHÔNG đổi: version/build (vẫn 1.3.4 / Build 017), logic trích xuất, `ImprovementItem` / `AfterVisualRegion`,
  renderer, React UI, `ReportExtractor.spec`, `backup/…STABLE/`. KHÔNG merge PR #7, KHÔNG xuất bản LAN,
  KHÔNG tạo ZIP phát hành, KHÔNG chạy `BUILD_AND_PUBLISH.bat`, KHÔNG force push.
