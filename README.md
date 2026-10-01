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

## 1. Portable build (Windows)

On the **development PC** (Python 3.9+ installed):

```
BUILD_PORTABLE.bat
```

This creates a venv, installs the build dependencies, runs the tests, runs
PyInstaller (onedir) and produces:

```
dist\ReportExtractor_Portable\
├── ReportExtractor.exe
├── _internal\
├── config.json
└── README.md
```

Copy the whole `ReportExtractor_Portable` folder to any Windows 10/11 PC and
double-click `ReportExtractor.exe`. No Python, pip, or tools required there.

> The build must be run on Windows – PyInstaller does not cross-compile.
> Recommended: Microsoft PowerPoint installed on the *target* PC gives the best
> QPN rendering; otherwise LibreOffice, otherwise the built-in renderer is used.

Development run: `RUN_DEV.bat` – tests: `RUN_TESTS.bat`.

Headless mode (same pipeline, also works from the exe):

```
ReportExtractor.exe --diag
ReportExtractor.exe --cli D:\Reports --template D:\Templates\Verification.xlsx --output D:\Output\Kiem_chung_doi_sach_TONG_HOP.xlsx [--server 192.168.1.50:11434 --model qwen3:4b] [--no-ai] [--force]
```

## 2. Hướng dẫn sử dụng (GUI)

1. **Máy AI (Ollama)**: nhập `192.168.1.50:11434` hoặc `http://192.168.1.50:11434`
   → bấm **Kiểm tra kết nối**. Danh sách model được đọc từ `GET /api/tags`;
   model chứa `qwen` + `4b` được chọn sẵn (có thể đổi trong combobox).
2. **Form Excel**: chọn file `.xlsx` mẫu (phải có sheet `Kiểm chứng`; sheet
   `Phân loại` nếu có sẽ dùng để map Model / Item).
3. **Thư mục báo cáo**: chọn thư mục hoặc kéo-thả thư mục / file `.pptx` vào danh
   sách. Toàn bộ thư mục con được quét tự động.
4. **File kết quả**: file `.xlsx` tổng hợp (không bao giờ ghi đè form gốc). Chạy
   lại với cùng file kết quả sẽ ghi tiếp theo quy tắc master cố định:
   - Management Number không có trong file → *Không tìm thấy Management Number*
     (không tạo dòng mới, bổ sung số rồi chạy lại);
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
`Đang chờ · Đang đọc PPTX · Đang phân tích (AI) · Đang trích xuất QPN · Đang trích xuất hình ảnh · Đang ghi Excel · Hoàn thành · Cần kiểm tra · Không tìm thấy Management Number · Lỗi · Bỏ qua — đã cập nhật`

## 3. Business rules (fixed)

### Vùng nội dung (content-region) – Nguyên nhân / Nội dung đối sách cải tiến
Chỉ **vùng nội dung chính** của slide được sao chép (nguyên văn, đúng thứ tự). Loại bỏ theo cấu trúc/hình học
(tỷ lệ theo kích thước slide, không dùng toạ độ tuyệt đối): tiêu đề slide (`1. NGUYÊN NHÂN`, `3. CẢI TIẾN TRONG
SẢN XUẤT`…), nhãn tròn bên trái (`Nguyên nhân`, `Cải tiến trong kiểm tra`), nút chú thích ảnh (`Trước cải tiến`,
`Sau cải tiến`), logo, footer, mũi tên. Tiêu đề phụ trong nội dung (`Nguyên nhân trong kiểm tra:`) và các dòng
`+ Trước:` / `+ Sau:` là nội dung nghiệp vụ và được giữ nguyên. `XỬ LÝ TẠM THỜI` vẫn bị loại; `ĐỐI SÁCH LÂU DÀI`
vẫn được lấy (phần nội dung).

### Hình ảnh cải tiến = CHỈ ảnh "Sau cải tiến" của cải tiến sản xuất
Neo = nút chú thích `Trước/Sau cải tiến` hoặc dòng `+ Trước:` / `+ Sau:`; mỗi ảnh được gán theo hình học tương đối
với neo (ảnh cùng hàng cạnh ảnh đã có chú thích dùng chung chú thích). Ảnh Trước, mũi tên, logo, ảnh của slide/mục
`Cải tiến trong kiểm tra` / `kiểm soát` không được chèn. Ảnh không xác định chắc chắn → **không chèn** và ghi
`Cần kiểm tra: Không xác định chắc chắn ảnh Sau cải tiến tại slide X` (không bao giờ chèn toàn bộ ảnh của slide).
Qwen chỉ chọn slide/mục, không chọn từng ảnh. Chi tiết quyết định từng ảnh: `python run.py --inspect <pptx>`.

### Management Number trùng nhiều dòng
Dòng trên cùng là đích; các dòng trùng giữ nguyên và được tô đỏ; batch tiếp tục (Cần kiểm tra).


| Field | Rule |
|---|---|
| Management number | From file name / report text (e.g. `260918080-VOC`); blank if not found |
| Tên vendor | **always blank** (user fills manually) |
| Ngày phát sinh | **always blank** (user fills manually) |
| Model | `SM-A185` → `A185`; sheet *Phân loại* preferred |
| Item | sheet *Phân loại* mapping preferred (keyword → Item) |
| Nội dung lỗi | original defect lines with quantities (`Xước: 15ea …`) |
| QPN | **complete “Quality Problem Notice” slide rendered as PNG**, embedded |
| Nguyên nhân | full original text of cause sections (sub-sections kept) |
| Nội dung đối sách cải tiến | **full original text** of all improvement slides in source order incl. *ĐỐI SÁCH LÂU DÀI*; wrap + top aligned |
| Xử lý tạm thời | **excluded** by section structure (not by phrase) |
| Hình ảnh cải tiến | improvement slides rendered and stacked vertically (JPEG), embedded |
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
* Bề rộng hiển thị = bề rộng vùng đích thực tế (độ rộng cột / vùng merge, quy đổi `px = width*7 + 5`) trừ lề 4 px mỗi bên;
  chiều cao theo đúng tỉ lệ gốc – không méo, không cắt. Ảnh nguồn nhỏ/hẹp vẫn được phóng vừa bề rộng.
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

### Management Number chưa có trong Excel → tự thêm dòng mới

* Điều kiện: mã hợp lệ theo parser hiện có (có ngày YYMMDD thật), trong thời gian xử lý, không phải bản trùng bị bỏ.
  Tên file không có mã / ngày sai / ngoài kỳ / bản trùng → **không** tạo dòng.
* Vị trí: `ExcelWriter.create_row` dùng `next_row()` – dòng báo cáo trống đầu tiên của bảng (tận dụng dòng form trống
  có sẵn), nếu không thì ngay sau dòng báo cáo cuối; không bao giờ là header/tiêu đề/sheet khác.
* Chỉ sao chép **định dạng** từ dòng báo cáo gần nhất phía trên (`_copy_row_style`: style/border/fill/font/number format/
  alignment/wrap, merge một dòng, chiều cao dòng). Không sao chép giá trị, ảnh QPN/ảnh cải tiến, WEEK đã nhập tay.
* Ghi Management Number (+ STT) rồi `save()` ngay, sau đó xử lý PPT bằng pipeline chuẩn vào đúng dòng đó; lần chạy sau tìm
  lại được dòng này (kể cả khi trích xuất lỗi) và điền tiếp theo cơ chế incremental – không bao giờ tạo dòng thứ hai.
* Dòng Excel trùng mã đã có sẵn → giữ nguyên rule cũ (dòng trên cùng là đích, các dòng kia tô đỏ), không thêm dòng.
* Cache: cache hit nhưng mã không có trong file Excel hiện tại → `CACHE_MISS reason=master_row_missing` → tạo dòng mới
  (mỗi tháng một file Excel khác nhau).
* Log: `MASTER_NEW management_number=… row=…`, `MASTER_NEW_RETRY …existing_partial_row=…`, `MASTER_NEW_FAILED …`.
  Chẩn đoán GUI: `Dòng Excel: Tạo mới (dòng N)`.
* GUI gồm 2 tab: `Xử lý báo cáo` (thư mục, file, thời gian xử lý, kết quả, tiến độ) và `Cấu hình & Ollama`.
