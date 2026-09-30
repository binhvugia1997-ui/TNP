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

Cấu hình (máy AI, model, các đường dẫn cuối) được lưu vào `config.json` cạnh exe.

### Trạng thái
`Đang chờ · Đang đọc PPTX · Đang phân tích (AI) · Đang trích xuất QPN · Đang trích xuất hình ảnh · Đang ghi Excel · Hoàn thành · Cần kiểm tra · Không tìm thấy Management Number · Lỗi · Bỏ qua — đã cập nhật`

## 3. Business rules (fixed)

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
