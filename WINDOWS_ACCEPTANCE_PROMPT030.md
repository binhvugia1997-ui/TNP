# PROMPT-030 — CHECKLIST NGHIỆM THU WINDOWS

> **Trạng thái tích hợp (PROMPT-030B):** checklist này được giữ nguyên từ PROMPT-030A
> (`aee4580`) và đã được tích hợp vào Build 20 (`bb1fd46`) trên nhánh làm việc
> `arena/fdd1d673-tnp` — thay `arena/f4a4d15e-tnp` bằng nhánh đó khi chạy.
> **Lưu ý hợp nhất:** render pipeline là hợp đồng Build 20/PROMPT-035 (chỉ `DispatchEx`,
> chứng minh ownership, không attach/kết thúc PowerPoint của user); probe string là hợp đồng
> Build 20; diagnostics PROMPT-030 được giữ. **PowerPoint COM + UI thủ công: CHƯA kiểm thử
> trên Windows thật trong phiên tích hợp này (môi trường Linux) — chạy checklist dưới trên
> Windows để nghiệm thu.**

> **Vị trí thực thi trên máy Windows:** `D:\arena\TNP\TNP`  
> **Nhánh:** `arena/fdd1d673-tnp` (không merge main, không chạy `BUILD_AND_PUBLISH.bat`, không tăng version)  
> **Yêu cầu hệ thống:** Windows 10/11 64-bit, PowerPoint đã cài (để kiểm COM), Python 3.10–3.14, `.venv` đã `setup.bat`

Toàn bộ lệnh dưới là PowerShell, copy-paste được. Mở **PowerShell** (không phải CMD) và chạy từng khối.

---

## 0. Chuẩn bị

```powershell
cd "D:\arena\TNP\TNP"
git status
git branch --show-current   # phải là arena/f4a4d15e-tnp
git log --oneline -3
# Kiểm tra không có thay đổi version
Select-String -Path app\__init__.py -Pattern "__version__|BUILD"
Get-Content VERSION.txt 2>$null | Select-Object -First 5
```

## 1. PowerPoint COM — Renderer

### 1.1 Kiểm tra probe và diagnostics (không cần mở UI)

```powershell
cd "D:\arena\TNP\TNP"
.\.venv\Scripts\python -m pytest tests/test_prompt030_powerpoint_renderer.py -v
.\.venv\Scripts\python -c "from app.qpn_renderer import powerpoint_probe, renderer_status_detailed; import json; print(powerpoint_probe()); print(json.dumps(renderer_status_detailed(), ensure_ascii=False, indent=2))"
```

**Tiêu chí đạt:**
- Trên máy có PowerPoint: `powerpoint_probe().available == True`, `stage == ok`, `reason == "PowerPoint có sẵn (ProgID + win32com)"`, `renderer_status_detailed()["powerpoint"]["status"] == "OK"`.
- Trên máy không có PowerPoint: `available == False`, `stage` là `registry_missing`/`not_windows`, UI báo fallback rõ ràng, không crash.

### 1.2 Kiểm tra UI với PPTX lỗi thực tế (PPTX từng bị nền đen)

```powershell
cd "D:\arena\TNP\TNP"
.\.venv\Scripts\python -m app.main
# hoặc: .\run.bat  (không phải BUILD_AND_PUBLISH.bat)
```

**Thao tác trong app:**
1. Mở PPTX lỗi thực tế (ví dụ: báo cáo có slide "3. CẢI TIẾN TRONG SẢN XUẤT" với caption "Trước cải tiến"/"Sau cải tiến").
2. Chạy trích xuất (Quét → Xử lý). Mở tab **Learning / Kiểm tra ảnh cải tiến**.
3. Kiểm tra badge renderer: khi COM thành công phải hiển thị **"PowerPoint"**, KHÔNG phải "Trình vẽ tích hợp".
4. Cảnh báo vàng fallback phải **biến mất** khi PowerPoint render thành công.
5. So sánh ảnh render với PowerPoint gốc: chữ, font, bold phải khớp; caption "Trước cải tiến"/"Sau cải tiến" **không còn bị nền đen sai**.
6. Nếu fallback xảy ra, mở `logs\app.log` và tìm:
   ```
   POWERPOINT_RENDER availability=... stage=...
   POWERPOINT_RENDER com_init=... open=... export=... fallback_stage=... fallback_reason=...
   ```
   Phải có `fallback_stage` là `probe`/`com_init`/`open`/`export` và `fallback_reason` tiếng Việt rõ ràng.

### 1.3 Kiểm tra đóng ứng dụng không để lại tiến trình PowerPoint rác

```powershell
Get-Process POWERPNT -ErrorAction SilentlyContinue | Select-Object ProcessName,Id,StartTime
# Đóng app hoàn toàn (File -> Thoát), đợi 5s
Start-Sleep 5
Get-Process POWERPNT -ErrorAction SilentlyContinue | Select-Object ProcessName,Id
# Phải không còn POWERPNT.EXE do app tạo ra (hoặc chỉ còn instance người dùng mở trước đó)
Get-Content logs\app.log -Tail 60 | Select-String "POWERPOINT_RENDER.*cleanup"
# Phải thấy "cleanup close=ok" và "cleanup quit=ok/skipped"
```

## 2. Fallback renderer (khi không có PowerPoint)

```powershell
cd "D:\arena\TNP\TNP"
# Giả lập không có PowerPoint bằng cách chặn registry (chỉ test tự động đã làm), trên Windows thực:
# - Đổi tên tạm registry PowerPoint.Application (khuyến nghị chỉ test tự động)
# - Thay vào đó, kiểm tra builtin luôn sẵn sàng:
.\.venv\Scripts\python -c "from app.qpn_renderer import SlideRenderer; from pathlib import Path; from tests.conftest import *; print('builtin OK')"
.\.venv\Scripts\python -m pytest tests/test_prompt030_powerpoint_renderer.py::test_non_windows_deterministic_builtin -v
```

**Tiêu chí:** Khi PowerPoint không khả dụng, `SlideRenderer` tự fallback sang `libreoffice` (nếu có) rồi `builtin` Pillow, không exception, logs ghi rõ `fallback_stage=probe`.

## 3. Lưu dữ liệu bền vững (Learning persistence)

### 3.1 Tạo nhãn và kiểm tra file bền vững

```powershell
cd "D:\arena\TNP\TNP"
.\.venv\Scripts\python -m app.main
# Trong Learning:
# - Chọn 1 candidate (ví dụ slide 5, shape #...), bấm nhãn "Sau cải tiến", ghi chú "kiểm thử PROMPT-030", Lưu
# Kiểm tra file hệ thống:
dir learning_data\previews | Format-Table Name,Length,LastWriteTime
dir learning_data\image_labels.jsonl | Format-List Name,Length
dir learning_data\previews\*.tmp    # phải rỗng (ghi nguyên tử .tmp -> replace)
# Phải thấy:
# - previews\<24hex>.jpg        (candidate crop, JPEG q82, ~5-30KB)
# - previews\<24hex>.json       (side-car provenance: candidate_id, source_file, bounds, slide_size, item_id...)
# - previews\<24hex>_slide.jpg  (full slide render, ~30-150KB, shared per slide)
# - image_labels.jsonl          (1 dòng JSON với candidate_id, label, note)
Get-Content learning_data\image_labels.jsonl -Tail 1 | ConvertFrom-Json | Format-List candidate_id,label,note,source_name,slide
Get-Content learning_data\previews\*.json | Select-String "candidate_id|source_file" | Select-Object -First 20
# Kiểm tra không rò rỉ đường dẫn tuyệt đối trong export:
.\.venv\Scripts\python -c "from pathlib import Path; print(Path('learning_data/image_labels.jsonl').read_text(encoding='utf-8')[:800])"
# Phải thấy source_name là tên file, không phải "D:\arena\..." trong JSONL export? (source_file tuyệt đối chỉ trong side-car .json)
```

### 3.2 Kiểm tra va chạm (collision) đã được xử lý

```powershell
# Chuẩn bị 2 thư mục có cùng tên file
New-Item -ItemType Directory -Force -Path "C:\temp\va_cham\A", "C:\temp\va_cham\B" | Out-Null
Copy-Item "D:\arena\TNP\TNP\tests\fixtures\*.pptx" -Destination "C:\temp\va_cham\A\report.pptx" -ErrorAction SilentlyContinue
Copy-Item "D:\arena\TNP\TNP\tests\fixtures\*.pptx" -Destination "C:\temp\va_cham\B\report.pptx" -ErrorAction SilentlyContinue
# Mở lần lượt từng file trong app, tạo nhãn khác nhau, kiểm tra learning_data\previews có 2 file .jpg khác nhau (key khác nhau)
.\.venv\Scripts\python -c "from app.image_learning import learning_asset_key; print(learning_asset_key('ID|S5|#1|1','C:/temp/va_cham/A/report.pptx')); print(learning_asset_key('ID|S5|#1|1','C:/temp/va_cham/B/report.pptx'))"
# Hai key phải khác nhau
dir learning_data\previews\*.jpg | Measure-Object | Select-Object Count
```

### 3.3 Khởi động lại ứng dụng (restart)

```powershell
cd "D:\arena\TNP\TNP"
# Đóng app hoàn toàn (File -> Thoát), đảm bảo không còn process
Get-Process python -ErrorAction SilentlyContinue | Where-Object {$_.Path -like "*TNP*"} | Select-Object Id,ProcessName
# Mở lại
.\.venv\Scripts\python -m app.main
# Quay lại Learning -> kiểm tra:
# - Tổng số mẫu, số "Sau cải tiến" giữ nguyên
# - Candidate đã gán nhãn vẫn hiển thị, src bắt đầu bằng "data:image/" (không phải rỗng)
# - Ảnh preview và slide preview vẫn hiển thị (không phải ô trắng), đường dẫn file trong previews vẫn tồn tại
dir learning_data\previews\*.jpg | Sort-Object LastWriteTime -Descending | Select-Object -First 5 Name,Length
# Thử gán nhãn lại sau restart (đổi từ Sau -> Trước), Lưu -> kiểm tra JSONL có dòng mới với previous_label
Get-Content learning_data\image_labels.jsonl -Tail 2 | ConvertFrom-Json | Format-Table candidate_id,label,previous_label,timestamp
```

### 3.4 Xử lý preview thiếu/hỏng (graceful)

```powershell
# Giả lập preview bị xóa/hỏng
$jpg = Get-ChildItem learning_data\previews\*.jpg | Select-Object -First 1
if ($jpg) { Remove-Item $jpg.FullName -Force; Write-Host "Đã xóa $($jpg.Name) để test thiếu preview" }
.\.venv\Scripts\python -m app.main
# Mở Learning -> candidate đó vẫn hiển thị dòng nhãn, nhưng preview rỗng (không crash), logs có dòng "persisted dto failed ... tolerating" là bình thường
# Khôi phục bằng cách mở lại PPTX gốc và vào Learning (DTO sẽ tự regenerate từ PPTX và persist lại)
```

## 4. Cập nhật ứng dụng (updater) — không mất learning_data

```powershell
cd "D:\arena\TNP\TNP"
# Kiểm tra danh sách bảo toàn
.\.venv\Scripts\python -c "from app import updater; print(updater.PRESERVED_ENTRIES); print(updater.is_preserved('learning_data')); print(updater.release_entries(Path('D:/tmp/fake_staged')))"
# Phải thấy "learning_data" trong PRESERVED_ENTRIES, is_preserved True, release_entries không chứa learning_data
# Giả lập portable có learning_data và staged mới không có learning_data (như test test_updater_preserves_learning_data)
.\.venv\Scripts\python -m pytest tests/test_prompt030_learning_persistence.py::test_updater_preserves_learning_data -v
# Trên Windows thực: chuẩn bị thư mục portable test, chạy updater staging rồi kiểm tra learning_data vẫn còn sau khi copy
```

## 5. Checklist UI Learning (thủ công, bắt buộc)

Mở app (`py -m app.main`), vào **Kiểm tra ảnh cải tiến**, kiểm tra từng mục:

- [ ] Label buttons ở **đầu** panel phải, phía trên "Report nguồn" (không phải cuối)
- [ ] Single click chỉ chọn (không lưu), double-click **không có active label** thì hiện hint "Vui lòng chọn nhãn" và không đoán
- [ ] Double-click **có active label** thì xác nhận đúng **1 lần** (không trùng, kiểm tra JSONL chỉ thêm 1 dòng)
- [ ] Ctrl+click multi-select trong cùng slide hoạt động (vòng highlight + badge ✓, đoạn "Đã chọn N đối tượng")
- [ ] Shift+click chọn dải (range) trong cùng slide hoạt động
- [ ] Bulk label: chọn 2-3 candidate cùng slide → chọn 1 nhãn → Lưu → mỗi candidate giữ identity riêng (JSONL có 2-3 dòng candidate_id khác nhau)
- [ ] Cross-slide: chọn candidate slide 5, Ctrl+click candidate slide 6 → chỉ giữ 1 (không cho bulk khác slide)
- [ ] Sau khi xác nhận, điều hướng tới candidate chưa xác nhận tiếp theo hoạt động
- [ ] Sau restart, DTO khôi phục nhãn + preview + slidePreview (kiểm tra như mục 3.3)

## 6. Lệnh tổng hợp tự động (dùng để báo cáo trước khi nghiệm thu Windows)

```powershell
cd "D:\arena\TNP\TNP"
.\.venv\Scripts\python -m pytest tests/test_prompt030_powerpoint_renderer.py tests/test_prompt030_learning_persistence.py tests/test_prompt030_e2e_real_flow.py -v
.\.venv\Scripts\python -m pytest tests/test_prompt030_powerpoint_renderer.py tests/test_prompt030_learning_persistence.py tests/test_prompt030_e2e_real_flow.py tests/test_image_learning.py tests/test_application_service.py tests/test_updater.py -q
npm --prefix frontend run test:learning-hooks
node --test frontend/tests/learning-030.test.mjs
npm --prefix frontend run build
npx --prefix frontend tsc --noEmit
.\.venv\Scripts\python -m compileall app/application_service.py app/image_learning.py app/qpn_renderer.py app/gui_controller.py
.\.venv\Scripts\python -m pytest tests/ -q
```

## 7. Tiêu chí nghiệm thu cuối

- [ ] PowerPoint COM được dùng khi cài đặt (badge PowerPoint, không fallback, không nền đen sai)
- [ ] Fallback có diagnostics staged rõ ràng khi không dùng PowerPoint
- [ ] Learning preview/slide preview được lưu bền vững dưới `learning_data\previews` (không phải Temp), sống sót sau restart và update
- [ ] Va chạm tên file/thư mục khác nhau đã được xử lý (learning_asset_key)
- [ ] Không duplicate JPEG mỗi lần refresh (atomic .tmp, dedup check)
- [ ] Toàn bộ tests tự động pass (trừ lỗi môi trường PyMuPDF đã được cài lại thì pass)
- [ ] Frontend typecheck + build pass
- [ ] Checklist UI thủ công ở mục 5 đã tick đủ
- [ ] Không merge main, không chạy BUILD_AND_PUBLISH.bat, không tăng version

---

**Ghi chú:** Nếu `.\.venv\Scripts\python -m pytest tests/test_setup_support.py::test_runtime_imports_and_requirements_cover_each_other` vẫn báo `['PyMuPDF']`, chạy:
```powershell
.\.venv\Scripts\pip install "PyMuPDF>=1.24"
```
rồi chạy lại. Trên Linux CI lỗi này từng là do môi trường thiếu PyMuPDF, đã được xác minh bằng `git stash` chạy trên baseline vẫn lỗi, không phải do PROMPT-030.

**Kết thúc:** Sau khi tất cả mục trên đạt, ghi báo cáo với commit hash, số lượng tests, và dòng cuối:

```
PROMPT-030 IMPLEMENTATION COMPLETE — READY FOR WINDOWS ACCEPTANCE
```
