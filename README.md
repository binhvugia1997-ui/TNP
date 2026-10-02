# Report Extractor – Windows Portable

Batch-process PowerPoint quality reports (`.pptx`) and copy their content into an
existing Excel verification template (sheet **Kiểm chứng**).

```
Report folder ─► recursive .pptx discovery ─► read PPTX structure
   ─► remote Ollama + Qwen 4B classifies WHERE things are (slide numbers only)
   ─► program copies the ORIGINAL text/images from the PPTX
   ─► validation ─► one row per report in the Excel template ─► one consolidated .xlsx
```

**Data-integrity rule:** *AI decides WHERE the information is – the program copies
WHAT the source actually says.* No LLM text is ever written to Excel (model / item /
management number returned by the LLM are only used when found verbatim in the PPTX).

---

## 1. Portable build (Windows) – v1.0.3 / PROMPT-003

On the **development PC** (Python 3.10 – 3.13, 64-bit; the build refuses other versions):

```
build_portable.bat
```

`tools/build_portable.py` creates `.venv-build` from `requirements-build.txt`, runs pyflakes + pytest, runs
PyInstaller with the committed `ReportExtractor.spec` (**onedir, windowed, no console, no UPX, no onefile**),
assembles the portable folder, validates it (no tests / .venv / .git / sample PPTX-XLSX / developer
`config.json` inside, no Ollama or model files), zips it and writes SHA256 sums:

```
dist\ReportExtractor_v1.0.3_Portable\
├── ReportExtractor.exe
├── _internal\                 (read-only bundle – never written to)
├── Output\  logs\  config\    (writable state, created next to the exe)
├── README.txt  FIRST_RUN.txt  VERSION.txt  SHA256SUMS.txt
└── Install_Ollama_Optional.bat
release\ReportExtractor_v1.0.3_Portable.zip + release\SHA256SUMS.txt
```

Runtime paths come from `app/runtime_paths.py`: *resource root* = `sys._MEIPASS` (`_internal`) when packaged,
*portable root* = folder of `ReportExtractor.exe` (never the current working directory). Config is
`config\config.json` (a pre-1.0.3 `config.json` next to the exe is still read once and migrated), start-up log
`logs\app.log` (`STARTUP version=1.0.3 prompt=PROMPT-003 packaged=… executable=… portable_root=…`), fatal start-up
errors → `logs\startup_error.log` + dialog. Defaults shipped: `http://127.0.0.1:11434`, `qwen3:4b`; Ollama and
models are **not** bundled (the app starts and works in heuristic mode without Ollama).

Copy the whole folder to any Windows 10/11 PC (paths with spaces / Vietnamese are fine; avoid `Program Files`)
and double-click `ReportExtractor.exe`. The build must run on Windows – PyInstaller does not cross-compile.
PowerPoint / LibreOffice on the target PC are optional accelerators for QPN rendering; the built-in renderer is used otherwise.

**Ollama in the LAN:** tab *Cấu hình & Ollama* → **Tìm Ollama trong mạng LAN** (user-initiated only; scans port
11434 on the PC's private subnet with bounded concurrency; *Dừng tìm* cancels). Found servers are verified via
`/api/tags`; choose one and confirm **Sử dụng server này** – the app never switches server or model
automatically. The Ollama PC needs `OLLAMA_HOST=0.0.0.0` and port 11434 open.

Development run: `run.bat` / `RUN_DEV.bat` – tests: `RUN_TESTS.bat` – CLI: `run_cli.bat --cli …`.

Headless mode (same pipeline, also works from the exe):

```
ReportExtractor.exe --diag
ReportExtractor.exe --cli D:\Reports --template D:\Templates\Verification.xlsx --output D:\Output\Kiem_chung_doi_sach_TONG_HOP.xlsx [--server 192.168.1.50:11434 --model qwen3:4b] [--no-ai] [--force]
```

## 2. Hướng dẫn sử dụng (GUI)

> **v1.0.2 / PROMPT-002 – giao diện mới.** Tab **Xử lý báo cáo** đi theo đúng thứ tự làm việc: *Nguồn dữ liệu* →
> *Thời gian xử lý* → *Danh sách báo cáo* (ô tìm kiếm `Ctrl+F`, bộ lọc hiển thị, Quét lại `F5`, Khôi phục, Xóa khỏi danh
> sách – chỉ xóa khỏi danh sách xử lý, không xóa file gốc) → **▶ BẮT ĐẦU XỬ LÝ / ■ DỪNG SAU FILE HIỆN TẠI** → tiến trình
> (phần trăm, giai đoạn, file hiện tại, đã chạy / còn khoảng, thẻ Hoàn thành / Cần kiểm tra / Lỗi / Bỏ qua) → *Kết quả xử lý*
> (nháy đúp xem chi tiết). Trạng thái Ollama hiển thị gọn ở đầu danh sách; mọi cấu hình Ollama, tùy chọn xử lý và
> **Nhật ký xử lý** (Mở file log / Mở thư mục log / Xóa phần hiển thị – không xóa `app.log`) nằm ở tab **Cấu hình & Ollama**.
> Kích thước cửa sổ / trạng thái phóng to được ghi nhớ trong `config.json`. Logic trích xuất không thay đổi so với v1.0.1.

1. **Máy AI (Ollama)**: nhập `192.168.1.50:11434` hoặc `http://192.168.1.50:11434`
   → bấm **Kiểm tra kết nối**. Danh sách model được đọc từ `GET /api/tags`;
   model chứa `qwen` + `4b` được chọn sẵn (có thể đổi trong combobox).
2. **Form Excel**: chọn file `.xlsx` mẫu (phải có sheet `Kiểm chứng`; sheet
   `Phân loại` nếu có sẽ dùng để map Model / Item).
3. **Thư mục báo cáo**: chọn thư mục hoặc kéo-thả thư mục / file `.pptx` vào danh
   sách. Toàn bộ thư mục con được quét tự động.
4. **File kết quả**: file `.xlsx` tổng hợp (không bao giờ ghi đè form gốc). Chạy
   lại với cùng file kết quả sẽ ghi tiếp theo quy tắc master cố định:
   - Management Number hợp lệ nhưng chưa có trong file → tự động **thêm đúng một
     dòng mới** (PROMPT-004D) rồi xử lý bình thường → *Hoàn thành — đã thêm
     Management Number mới*; tên file không có Management Number → *Cần kiểm tra —
     Không xác định được Management Number từ tên file* (không ghi gì);
   - dòng đã đủ mọi trường tự động → *Bỏ qua — đã cập nhật* (không gọi Qwen,
     không trích xuất lại, không ghi lại QPN/ảnh);
   - dòng còn thiếu trường → tự động **chỉ điền các ô còn trống**, ô đã có dữ liệu
     (kể cả WEEK +1..+8, cột thủ công) giữ nguyên;
   - lần trước Lỗi / bị dừng / Cần kiểm tra → được xử lý lại.
   Tick “Xử lý lại…” (`--force`) mới ghi đè các trường tự động.
5. **BẮT ĐẦU TRÍCH XUẤT** – theo dõi trạng thái từng file; **Dừng sau file hiện tại**
   dừng an toàn sau khi ghi xong file đang xử lý.
6. Kết thúc: **Mở file kết quả**, **Mở thư mục kết quả**, **Xem chi tiết lỗi**, **Chẩn đoán**.

**Kết nối Ollama** (`IP / Server`, `Port`, `Model`): chấp nhận `127.0.0.1`, `192.168.1.50`,
`192.168.1.50:11434`, `http://192.168.1.50:11434` (tự chuẩn hoá về một endpoint, hỗ trợ port tuỳ ý).
`Kiểm tra kết nối` dùng đúng client sản xuất (máy chủ → API → model có tồn tại); `Làm mới model` lấy
danh sách model thật đã cài trên máy chủ; `Lưu cấu hình` ghi IP/port/model vào `config.json` và được khôi phục
khi mở lại. Đổi IP/model có hiệu lực ngay cho lần kiểm tra/xử lý tiếp theo, không cần khởi động lại.
Ollama không bắt buộc: không kết nối được → chương trình hỏi và tiếp tục với heuristic fallback; mất kết nối
giữa chừng → báo cáo đó dùng heuristic, batch vẫn tiếp tục.

**Tiến độ**: `Đang xử lý: 2 / 5 — 36%` + thanh tiến độ cùng giá trị, tính từ số báo cáo đã kết thúc và mốc giai
đoạn THẬT của báo cáo hiện tại (không ước lượng tiến trình bên trong Qwen); `Đã chạy: MM:SS` (đồng hồ monotonic
bắt đầu khi batch thật sự chạy), `Còn khoảng: …` từ thời gian đo được của các báo cáo đã xong
(`Đang tính...` khi chưa có dữ liệu). Dừng sớm → không hiển thị 100%, không hiển thị ETA.

Cấu hình (máy AI, model, các đường dẫn cuối) được lưu vào `config.json` cạnh exe.

### Trạng thái
`Đang chờ · Đang đọc PPTX · Đang phân tích (AI) · Đang trích xuất QPN · Đang trích xuất hình ảnh · Đang ghi Excel · Hoàn thành · Hoàn thành — đã thêm Management Number mới · Cần kiểm tra · Cần kiểm tra — Không xác định được Management Number từ tên file · Lỗi · Bỏ qua — đã cập nhật`

## 3. Business rules (fixed)

### Vùng nội dung (content-region) – Nguyên nhân / Nội dung đối sách cải tiến
Chỉ **vùng nội dung chính** của slide được sao chép (nguyên văn, đúng thứ tự). Loại bỏ theo cấu trúc/hình học
(tỷ lệ theo kích thước slide, không dùng toạ độ tuyệt đối): tiêu đề slide (`1. NGUYÊN NHÂN`, `3. CẢI TIẾN TRONG
SẢN XUẤT`…), nhãn tròn bên trái (`Nguyên nhân`, `Cải tiến trong kiểm tra`), nút chú thích ảnh (`Trước cải tiến`,
`Sau cải tiến`), logo, footer, mũi tên. Tiêu đề phụ trong nội dung (`Nguyên nhân trong kiểm tra:`) và các dòng
`+ Trước:` / `+ Sau:` là nội dung nghiệp vụ và được giữ nguyên. `XỬ LÝ TẠM THỜI` vẫn bị loại; `ĐỐI SÁCH LÂU DÀI`
vẫn được lấy (phần nội dung).

### Hình ảnh cải tiến = CHỈ ảnh "Sau cải tiến" của cải tiến sản xuất
Thứ tự bằng chứng (PROMPT-004C, chỉ dùng cấu trúc PPTX – không OCR, không pixel, không Qwen):
1. nút chú thích `Trước/Sau cải tiến` gán cho ảnh gần nhất (ảnh cùng hàng dùng chung chú thích);
2. dòng `+ Trước:` / `+ Sau:` sở hữu vùng dọc bên dưới tới dòng kế tiếp;
3. **chữ màu xanh** (màu run/paragraph gốc, kể cả màu theme đã quy đổi; HSV hue 190–260°, bão hoà ≥ 0.35) của khối
   nội dung sản xuất, nằm ngay trên/bên cạnh ảnh → ảnh Sau; chữ xanh ở tiêu đề, logo, sidebar, footer, mục kiểm tra/
   kiểm soát không được tính;
4. **mũi tên định hướng** (auto-shape rightArrow/leftArrow/upArrow/downArrow… có xoay/lật, hoặc connector có đầu mũi
   tên) nằm trong khối nghiệp vụ, có ảnh ở cả hai phía trong hành lang của trục mũi tên: ảnh phía **đích** = Sau, phía
   nguồn = Trước; mỗi mũi tên chỉ gom nhóm ảnh liền kề (khoảng trống > 12 % slide kết thúc nhóm) nên mũi tên của mục
   này không phân loại ảnh của mục khác; mũi tên trong dải tiêu đề/footer hoặc không có ảnh hai phía bị bỏ qua;
   chữ xanh + đích mũi tên trùng nhau = bằng chứng rất mạnh; mâu thuẫn → không chèn.
Ảnh Trước, mũi tên, logo, ảnh của slide/mục `Cải tiến trong kiểm tra` / `kiểm soát` không được chèn. Ảnh duy nhất
không có bằng chứng **không** được coi là Sau. Ảnh không xác định chắc chắn → **không chèn** và ghi
`Cần kiểm tra: Không xác định chắc chắn ảnh Sau cải tiến tại slide X` (không bao giờ chèn toàn bộ ảnh của slide).
Qwen chỉ chọn slide/mục, không chọn từng ảnh. Chi tiết quyết định từng ảnh: `python run.py --inspect <pptx>`.

### Management Number trùng nhiều dòng
Dòng trên cùng là đích; các dòng trùng giữ nguyên và được tô đỏ; batch tiếp tục (không ép `Cần kiểm tra`; xem PROMPT-004).


| Field | Rule |
|---|---|
| Management number | From file name only (`260918080-VOC`, `260922134`); existing row → updated, valid key absent → exactly one new production row (PROMPT-004D) |
| Tên vendor | **always blank** (user fills manually) |
| Ngày phát sinh | derived only from Management Number YYMMDD (`260923045` → 23/09/2026) |
| Model | `SM-A185` → `A185`; sheet *Phân loại* preferred |
| Item | sheet *Phân loại* mapping preferred (keyword → Item) |
| Nội dung lỗi | file name after `LỖI` (date/SEV/parenthetical stripped), e.g. `BONG ATN, MỤN`; no invented quantities |
| QPN | **complete “Quality Problem Notice” slide rendered as PNG**, embedded |
| Nguyên nhân | full original text of cause sections (sub-sections kept) |
| Nội dung đối sách cải tiến | **full original text** of all improvement slides in source order incl. *ĐỐI SÁCH LÂU DÀI*; wrap + top aligned |
| Xử lý tạm thời | **excluded** by section structure (not by phrase) |
| Hình ảnh cải tiến | ONLY `Sau cải tiến` production pictures, source order, separate image objects |
| WEEK +1 … +8 | blank (never invented) |
| Source `.pptx` / template | read-only; output is always a copy |

Rows missing QPN / model / cause / improvement are still exported with blank
fields and marked **Cần kiểm tra** (also written to a `Ghi chú` column if the
template has one).

## 4. Output structure

```
Output\
├── Kiem_chung_doi_sach_TONG_HOP.xlsx
├── assets\  0001_QPN.png, 0001_IMPROVEMENT.jpg, 0001_IMPROVEMENT_pics\…   (audit copies)
└── logs\    app.log, errors.log, batch_result.json, history.json (duplicate protection)
```

The workbook is saved after **every** record.

## 5. Architecture

```
app/
├── main.py            entry point (GUI / --cli / --diag)
├── gui.py             Tkinter GUI, worker-thread + queue, drag & drop (tkinterdnd2)
├── config.py          config.json, Ollama URL normalisation
├── scanner.py         recursive .pptx discovery
├── pptx_parser.py     python-pptx → slides/blocks with positions, reading order
├── ollama_client.py   /api/tags, /api/generate (format=json, temperature 0)
├── classifier.py      heuristic + LLM slide classification (WHERE)
├── extractor.py       copies original text by section (WHAT), temp-handling exclusion
├── qpn_renderer.py    full-slide render: PowerPoint COM → LibreOffice → built-in Pillow
├── image_extractor.py improvement slide contact sheet + original pictures
├── excel_writer.py    openpyxl: template copy, header detection, row append, images
├── history.py         SHA-256 fingerprints / duplicate protection
├── batch_processor.py batch loop, stop-after-current, logging
├── diagnostics.py     runtime / renderer / Ollama / template / output checks
└── logger.py          errors.log + batch_result.json
tools/make_samples.py  generates sample reports + template (used by tests and demos)
tests/                 40 pytest tests
```

## 6. Notes / limitations

* Qwen is only a classifier. If the Ollama server is unreachable the app falls
  back to keyword/heading detection and logs it (`classifier` field in
  `batch_result.json`).
* Header detection is alias based (accent-insensitive), so column order in the
  template does not matter; merged header rows (WEEK +1…+8 under “Kiểm chứng”)
  are supported.
* Excel limits a row to 409 pt; embedded images are scaled to the column width
  and that height – the original resolution is kept inside the file, so the
  image can be enlarged in Excel.
* openpyxl preserves formatting, formulas, merges, widths and existing images;
  charts/VBA in the template are not preserved (use a plain `.xlsx` template).

## Kích thước ảnh trong Excel (QPN & Hình ảnh cải tiến)

* Mỗi ảnh **Sau cải tiến** hợp lệ là **một đối tượng ảnh riêng** trong Excel (không ghép), theo thứ tự xuất hiện trong báo cáo;
  file nguồn lưu tại `Output/assets/<prefix>_IMPROVEMENT_pics/afterNN_slideSS.png`.
* Bề rộng hiển thị = bề rộng vùng đích thực tế (độ rộng cột / vùng merge – cộng **mọi** cột trong vùng merge, tôn trọng
  `<col min max>`, cột ẩn, độ rộng mặc định của sheet) trừ lề 4 px mỗi bên; quy đổi theo ECMA-376
  `px = trunc(((256*W + trunc(128/7))/256)*7)` (openpyxl trả về độ rộng *đã gồm* 5 px đệm – công thức cũ `W*7+5` làm ảnh
  rộng hơn cột 5 px và tràn sang cột bên). Chiều cao theo đúng tỉ lệ gốc – không méo, không cắt. Ảnh nguồn nhỏ/hẹp vẫn
  được phóng vừa bề rộng.
* **Quy tắc chứa cứng**: mọi ảnh (QPN và từng ảnh Sau cải tiến) nằm hoàn toàn trong vùng đích
  (`x >= lề, y >= lề, x+w <= rộng−lề, y+h <= cao−lề`, dung sai 1 px làm tròn EMU), không chồng nhau, không tràn sang
  cột/dòng kế. Trình tự ghi: đọc bề rộng thật → tính bố cục → chốt chiều cao dòng (≤ 409 pt) → **vừa lại ảnh theo cả
  bề rộng lẫn chiều cao thật của dòng** → kiểm tra `assert_image_inside_area` → mới tạo anchor. `save()` kiểm tra lại
  `image_bounds_report()` và từ chối ghi nếu còn ảnh vượt ô (lỗi rõ ràng thay vì file sai). Đơn vị quy đổi tập trung:
  `col_width_to_px`, `pt_to_px`, `px_to_pt`, `px_to_emu`, `emu_to_px`.
* Giới hạn chiều cao tập trung: `MAX_IMAGE_HEIGHT_PT = 300` (app/excel_writer.py). Ảnh quá cao được thu nhỏ đồng tỉ lệ.
* Nhiều ảnh: xếp **một cột** (ảnh sau nằm dưới ảnh trước, cách 6 px). Vì Excel giới hạn chiều cao dòng 409 pt, nếu xếp một cột
  không đủ, hệ thống thu nhỏ đồng tỉ lệ hoặc chuyển sang lưới 2–3 cột – chọn phương án cho **bề rộng ảnh lớn nhất**
  (ưu tiên dễ đọc hơn dòng thấp). Không bao giờ có 5 ảnh thumbnail nằm ngang.
* Chiều cao dòng được nâng theo nhu cầu, **không bao giờ giảm** chiều cao đã có, không đụng dòng khác.
* QPN: ảnh slide được cắt bỏ viền trắng đồng nhất (`trim_white_margins`) rồi vừa bề rộng cột QPN; hình học độc lập với cột Hình ảnh cải tiến.
* Chạy lại: dòng đã đủ ảnh → `Bỏ qua — đã cập nhật`, ảnh giữ nguyên (không nhân đôi, không phóng to dần, không đổi chiều cao dòng).
* Thay đổi độ rộng cột trong template → kích thước ảnh tự thích ứng ở lần chạy sau (không có pixel cứng).

## Quét nhanh thư mục lớn (pre-scan) & Thời gian xử lý

Trước khi mở bất kỳ PPTX nào, mỗi file được quyết định chỉ bằng **tên file + kích thước + mtime + file Excel đang mở + cache JSON**:

```
Tên file → Management Number → Ngày phát sinh (YYMMDD, parser hiện có) → Thời gian xử lý
→ Trùng Management Number trong folder (chọn file mới nhất, hoà → thứ tự đường dẫn)
→ Cache 7 ngày (đã xử lý thành công gần đây, cùng path/size/mtime)
→ Tra dòng Excel (không có → PROCESS_NEW_ROW: tự tạo MỘT dòng báo cáo mới, ghi Management Number trước)
→ Dòng đã đầy đủ → "Bỏ qua — đã cập nhật"; dòng thiếu → PROCESS (bổ sung trường trống)
→ chỉ các báo cáo mới / còn thiếu dữ liệu mới mở PPTX / Qwen / trích xuất
```

* **Thời gian xử lý** (GUI, nhóm `Thời gian xử lý`): `Tự động theo file Excel` (mặc định – nhận diện tháng từ tên file kết quả,
  rồi tên file Kiểm chứng: `09_2026`, `2026-09`, `Tháng 9-2026`, `T09_2026`…; tên mơ hồ như `Kiem_chung_v2_09.xlsx` → báo
  `Không xác định được tháng từ tên file Excel…` và không cho bắt đầu), `Chọn tháng`, `Khoảng thời gian` (dd/mm/yyyy, bao gồm hai
  đầu, không tự đảo), `Tất cả`. Chọn tay luôn thắng tên file (có cảnh báo không chặn). CLI: `--month MM/YYYY`, `--from/--to`, `--all`.
* Ngày để lọc **chỉ** là ngày phát sinh suy ra từ Management Number (không dùng ngày trong tên file/PPT/thư mục/mtime).
* Cache: `Output/logs/fast_scan_cache.json` `{version, entries{mgmt: {path,size,mtime,occurrence_date,processed_at,status}}}`,
  ghi atomic (tmp + replace). Chỉ `Hoàn thành` / `Bỏ qua — đã cập nhật` được cache; `Cần kiểm tra`/`Lỗi`/không tìm thấy/bị dừng
  thì không. Hết hạn khi `ngày phát sinh <= hôm nay − 7 ngày` (hôm nay 14/10 → cutoff 07/10: 07/10 hết hạn, 08/10 còn), dọn
  lúc bắt đầu mỗi batch; cache hỏng/không ghi được → vẫn xử lý bình thường. Dòng Excel luôn là nguồn sự thật: cache hit nhưng
  dòng thiếu trường → `CACHE_MISS reason=master_row_incomplete` và xử lý lại. `--force / Xử lý lại` bỏ qua cache + dòng đầy đủ
  nhưng **vẫn** lọc theo thời gian.
* Tiến độ `Đang xử lý: x / N` chỉ đếm N = báo cáo thực sự vào pipeline; file bị loại ở pre-scan không ảnh hưởng % và ETA.
* Log máy đọc được: `PERIOD_AUTO/PERIOD_MANUAL`, `PERIOD_SKIP`, `SOURCE_DUPLICATE`, `FAST_SKIP`, `CACHE_MISS reason=…`,
  `CACHE_EXPIRE`, `MASTER_SKIP`, `MASTER_MISS`.

### Management Number chưa có trong Excel → tự động thêm MỘT dòng mới (PROMPT-004D, thay thế quy tắc PROMPT-004)

* Management Number chỉ lấy từ tên file (`260918080-VOC`, `260922134`), so khớp chính xác (chuẩn hoá) với cột trong
  master. Bảng quyết định:
  | Tên file | Hành động |
  |---|---|
  | không có Management Number hợp lệ | không tạo dòng, không sửa Excel → `Cần kiểm tra — Không xác định được Management Number từ tên file` |
  | có đúng 1 dòng trong Excel | dòng đó: đủ → `Bỏ qua — đã cập nhật`; thiếu → chỉ điền ô trống |
  | trùng nhiều dòng | dòng trên cùng là đích, các dòng còn lại tô đỏ, batch tiếp tục |
  | hợp lệ nhưng chưa có | **tạo đúng một dòng mới** (không hỏi xác nhận), ghi Management Number, xử lý bình thường → `Hoàn thành — đã thêm Management Number mới` |
* Dòng mới = dòng dữ liệu hợp lệ tiếp theo của bảng (không bao giờ là header/WEEK header/dòng cuối bảng), xác định
  tất định. Chỉ **sao chép định dạng** từ dòng dữ liệu gần nhất (style, viền, căn lề, chiều cao, định dạng số/ngày,
  merge một dòng); **không** sao chép giá trị nghiệp vụ, ảnh, WEEK, ghi chú. WEEK+1..+8 và các ô thủ công để trống.
* Bản sao lưu (một lần mỗi batch) được tạo **trước** khi thêm dòng; sao lưu lỗi → không tạo dòng, không sửa gì, báo lỗi
  rõ (`Không tạo được dòng mới cho Management Number … / Không tạo được bản sao lưu Excel …`). Dòng mới + mã được lưu
  xuống đĩa ngay, trước khi đọc PPTX; lưu lại sau mỗi báo cáo.
* Cùng báo cáo hai lần trong một batch → chỉ một dòng; chạy lần hai → tìm thấy dòng, đủ dữ liệu → `Bỏ qua — đã cập nhật`
  (không gọi Qwen). Chẩn đoán: `Dòng Excel: Tạo mới – Đã tạo dòng mới: <dòng> – Management Number mới: <mã>`.
* Pre-scan: trạng thái `Sẽ xử lý — Management Number mới` (lý do `Sẽ thêm dòng mới cho Management Number …`), dòng tổng
  kết `Management Number mới (sẽ thêm dòng): N`; log `MASTER_MISS … action=PROCESS_NEW_ROW`, `MASTER_NEW management_number=… row=…`,
  `MASTER_NEW_FAILED`. Cache hit nhưng mã không còn trong Excel → `CACHE_MISS reason=master_row_missing` → cũng tạo dòng mới.
* Ngày phát sinh chỉ suy ra từ YYMMDD của Management Number (`260923045` → 23/09/2026); bỏ qua mọi ngày trong tên
  file/slide. Nội dung lỗi ưu tiên tên file sau `LỖI` (bỏ ngày, `SEV`, ngoặc cuối, đuôi file): `… LỖI BONG ATN, MỤN
  22.9.2026 SEV.pptx` → `BONG ATN, MỤN`.

### Cập nhật tăng dần, trùng dòng, sao lưu (PROMPT-004)

* Một file `Kiem_chung.xlsx` cố định. Dòng đã đủ 9 trường quản lý (text + ảnh QPN + ảnh cải tiến) → `Bỏ qua — đã cập nhật`
  kiểm tra rẻ trước khi parse/Qwen. Dòng thiếu → chỉ điền trường còn trống (kể cả trường người dùng đã xoá tay), không
  hỏi, không ghi đè giá trị đã có; ảnh QPN/ảnh cải tiến có sẵn không bị thay/chèn lại; ảnh cải tiến thiếu được điền theo
  rule chỉ-ảnh-Sau. WEEK +1..+8, công thức, cột/sheet khác giữ nguyên; sheet `Data` không dùng.
* Trùng Management Number: dòng trên cùng là dòng dùng; các dòng khác giữ nguyên giá trị, tô đỏ (`FFC7CE`); batch tiếp
  tục và file vẫn có thể `Hoàn thành`. Chẩn đoán: `Dòng sử dụng: 18`, `Management Number bị trùng tại dòng: 24, 31 – Đã
  đánh dấu đỏ các dòng trùng.` Qwen không bao giờ chọn dòng.
* Sao lưu: trước lần sửa workbook **đầu tiên** của batch (điền, QPN, ảnh, tô đỏ) tạo một bản
  `<thư mục output>\backup\<tên>_backup_YYYYMMDD_HHMMSS.xlsx` (log `MASTER_BACKUP file=…`). Batch chỉ bỏ qua → không
  sao lưu, không ghi. Sao lưu thất bại → không sửa master (`Không tạo được bản sao lưu Excel trước khi ghi …`).

### Kết nối Ollama: ưu tiên local (PROMPT-004A)

Khi mở GUI (và khi bấm `Kiểm tra kết nối` với một địa chỉ không trả lời) thứ tự là `127.0.0.1:11434` → server đã lưu →
thủ công / `Tìm Ollama trong mạng LAN`. Ollama local có model → tự chọn, hiển thị `Server: 127.0.0.1`, `Port: 11434`,
trạng thái `Ollama local: Sẵn sàng — qwen3:4b`, **không** quét LAN. Local chạy nhưng thiếu model →
`Ollama local đang chạy nhưng không tìm thấy model qwen3:4b` + danh sách model đã cài (không tự đổi model). Local không
chạy → thử server đã lưu (vd `192.168.1.60:11434`); cả hai không được → `Không kết nối được Ollama local hoặc server đã
lưu.` và người dùng sửa Server/Port, `Kiểm tra kết nối` hoặc `Tìm Ollama trong mạng LAN`. Timeout thăm dò kết nối
(`probe_timeout`, mặc định 3 s) tách riêng khỏi timeout suy luận Qwen (`request_timeout`, 180 s). IP LAN cũ đã lưu
không bao giờ ngăn phát hiện Ollama local; `Lưu cấu hình` vẫn lưu đúng giá trị đang hiển thị.

### Màu chữ / độ tương phản GUI (PROMPT-004B)

Mọi style ttk có chữ (nút, nhãn, tab, Treeview, Entry/Combobox) khai báo tường minh cả foreground lẫn background cho các
trạng thái normal / active / pressed / disabled / selected / readonly trong `build_style_spec(theme)`; theme `clam` được
ưu tiên trên mọi hệ điều hành vì tôn trọng màu nền nút (theme Windows `vista` bỏ qua background của TButton nên chữ trắng
của nút `BẮT ĐẦU XỬ LÝ` từng bị "tàng hình"). Với theme native, nút chính dùng chữ tối trên nền mặc định. Hàm
`audit_style_contrast()` kiểm tra tỉ lệ tương phản WCAG (≥ 4.5:1, disabled ≥ 3:1) và được chạy trong test.

### Quét thư mục ≠ điều kiện xử lý (PROMPT-004C)

`scanner.py` chỉ loại bỏ artefact hệ thống (file khoá `~$…`, file/thư mục ẩn, phần mở rộng không phải PowerPoint) và ghi
log `SCAN_REJECT file=… reason=…` cho mọi mục giống PowerPoint bị bỏ; **không bao giờ** loại file theo nội dung tên
(Management Number, model, FRONT/REAR, ngoặc đơn, thiếu `SEV`/ngày, vendor…). Mọi báo cáo tìm thấy đều hiển thị trong
danh sách (bộ lọc mặc định `Tất cả file đã quét`; bộ lọc `File cần xử lý` vẫn giữ các dòng cần chú ý
`Không xác định được Management Number`; dòng `Sẽ xử lý — Management Number mới` là ứng viên bình thường). Dòng tổng kết:
`Tổng file phát hiện · Sẽ xử lý · Bỏ qua/đã cập nhật · Management Number mới · Lỗi/không hợp lệ · Đã loại thủ công`.

* GUI gồm 2 tab: `Xử lý báo cáo` (thư mục, file, thời gian xử lý, kết quả, tiến độ) và `Cấu hình & Ollama`.

### Danh sách file đã quét (kiểm tra trước khi xử lý)

* Tab `Xử lý báo cáo`: sau `Quét lại` (chạy đúng hàm `prescan` của batch, trong luồng nền, không mở PPTX, không ghi gì)
  hiện bảng: STT · Management Number · Ngày phát sinh · Tên file · Vendor · Trạng thái quét · Đường dẫn (nháy đúp xem chi tiết).
  Bộ lọc `Hiển thị: File cần xử lý / Tất cả file đã quét / File bị bỏ qua` chỉ ảnh hưởng hiển thị.
* `Xóa khỏi danh sách` (nút, phím Delete, chuột phải – cùng một `GuiController.exclude`) chỉ loại file **khỏi hàng đợi hiện tại**
  (`USER_EXCLUDED` – "Đã loại thủ công"); **không bao giờ** xoá/di chuyển/sửa file gốc, không ghi Excel, không tạo dòng mới,
  không mở PPTX, không gọi Qwen, không ghi cache. `Khôi phục` chỉ đảo ngược loại thủ công (không vượt qua ngoài kỳ / trùng /
  mã sai / Excel đầy đủ / cache).
* Hàng đợi thực tế = danh sách đã duyệt (`final_queue`); tiến độ/ETA dùng số này (20 ứng viên − 3 loại = `0 / 17`).
  Dòng Excel mới chỉ được tạo khi ứng viên thực sự được xử lý. Nhãn: `Cần xử lý sau khi quét / Đã loại thủ công / Sẽ xử lý`.
* Đổi thư mục / file Excel / tháng / khoảng / Tất cả / Xử lý lại → danh sách thành cũ: `Danh sách file đã thay đổi điều kiện.
  Vui lòng quét lại.` và không thể Bắt đầu. `Quét lại` xoá các loại thủ công và tính lại (kể cả file trùng chuẩn).
  Khi đang xử lý, các nút loại/khôi phục bị khoá (dùng `Dừng sau báo cáo hiện tại`).
* **Bố cục tab `Xử lý báo cáo`** (từ trên xuống): 1. Nguồn dữ liệu → 2. Thời gian xử lý → 3. AI (trạng thái + nút
  `Cấu hình Ollama…` mở tab 2) → 4. Các nút chức năng (`Danh sách: Quét file · Quét lại · Xóa khỏi danh sách · Khôi phục · Hiển thị`
  / `Xử lý: [ ] Xử lý lại · Bắt đầu xử lý · Dừng sau báo cáo hiện tại`) → 5. Thống kê quét → 6. Danh sách file →
  7. Tiến trình → 8. Kết quả xử lý. Mọi nút thao tác nằm **phía trên** danh sách file.
  Ba lớp cuộn riêng: thanh cuộn dọc của cả tab (cuộn giữa các mục); bảng `Danh sách file` và bảng `Kết quả xử lý`
  mỗi bảng có thanh cuộn dọc + ngang riêng, số dòng hiển thị cố định (12 / 10) nên hàng nghìn file không làm cửa sổ cao lên;
  cột `Tên file`, `Đường dẫn` giữ độ rộng đọc được và xem bằng cuộn ngang. Con lăn chuột trên bảng cuộn bảng, ngoài bảng cuộn trang.

## Cài đặt từ source trên Windows

```text
1. Download/clone source từ GitHub (thư mục bất kỳ, có dấu cách cũng được).
2. Nháy đúp setup.bat.
3. Chọn Ollama/model nếu cần (có thể bỏ qua).
4. Nháy đúp run.bat.
```

* **Python hỗ trợ: >= 3.10 và < 3.15, tức 3.10 – 3.14 (khuyến nghị 3.12)** – định nghĩa duy nhất trong `tools/setup_support.py`.
  `setup.bat` thử `py -3.12 / -3.13 / -3.14 / -3.11 / -3.10 / -3` rồi `python`, kiểm tra phiên bản thật; nếu không có sẽ đề nghị
  cài `Python.Python.3.12` qua **winget** (gói chính thức), không tải từ nguồn lạ. Không cần quyền Admin cho `.venv`/pip.
* Môi trường ảo cục bộ `.venv` (đã có trong `.gitignore`): tạo nếu chưa có, dùng lại nếu lành, tạo lại nếu hỏng/sai phiên bản.
  Thư viện runtime từ `requirements.txt`; công cụ test/dev (`pytest`, …) trong `requirements-dev.txt` – hỏi `Cài thêm công cụ
  phát triển/test? [y/N]`, mặc định không.
* **Ollama local là tuỳ chọn.** Ứng dụng hỗ trợ Ollama trên máy khác qua tab `Cấu hình & Ollama` (IP/port). Nếu chưa cài:
  `[1] Cài Ollama (winget Ollama.Ollama) / [2] Bỏ qua` (mặc định bỏ qua). Có Ollama → menu model `qwen3:1.7b`,
  `qwen3:4b (mặc định/khuyến nghị)`, `Cả hai`, `Không tải model` (mặc định); model đã có không tải lại; tiến trình tải là
  output thật của `ollama pull`; không yêu cầu server Ollama đang chạy khi cài.
* Kiểm tra cuối: import thư viện, module ứng dụng, tkinter, thư mục config – không xử lý báo cáo nào.
* Chạy lại `setup.bat` an toàn (idempotent). `run.bat` = `.venv\Scripts\python.exe run.py` (GUI); `run_cli.bat` chuyển
  tiếp tham số cho CLI (`--cli …`). Cửa sổ dừng lại khi lỗi (`Nhấn phím bất kỳ để thoát...`).
* Đây là cài đặt từ source; bản Portable xem mục 1 (`build_portable.bat`).

Checklist smoke test Windows: (A) máy sạch không Python/Ollama → winget cài Python → .venv → bỏ qua Ollama → run.bat;
(B) có Python, không Ollama; (C) Python + Ollama, chưa model → chọn [4]; (D) đã có qwen3:4b → "đã tồn tại — bỏ qua tải xuống";
(E) bỏ qua Ollama local, cấu hình IP máy khác trong tab Cấu hình; (F) đường dẫn có dấu cách; (G) chạy setup.bat lần 2 →
".venv: dùng lại", không tải lại model.
