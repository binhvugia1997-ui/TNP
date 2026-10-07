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
