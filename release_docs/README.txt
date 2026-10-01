REPORT EXTRACTOR {version} ({build_id}) – BẢN PORTABLE WINDOWS
=====================================================================

Chuyển báo cáo PowerPoint (PPTX) thành Bảng kiểm chứng Excel, hoàn toàn chạy
trên máy nội bộ. Không cần cài đặt, không cần quyền Administrator.

1. CÁCH DÙNG NHANH
------------------
  1) Giải nén toàn bộ thư mục {name} vào một nơi bạn có quyền ghi
     (ví dụ D:\ReportExtractor hoặc Desktop). KHÔNG đặt trong C:\Program Files.
  2) Chạy ReportExtractor.exe (xem FIRST_RUN.txt nếu Windows hiện cảnh báo SmartScreen).
  3) Tab "Xử lý báo cáo": chọn thư mục báo cáo PPTX, file mẫu Excel, file kết quả → BẮT ĐẦU XỬ LÝ.
  4) Tab "Cấu hình & Ollama": địa chỉ Ollama (tùy chọn), model, nút "Tìm Ollama trong mạng LAN".

2. CẤU TRÚC THƯ MỤC
-------------------
  ReportExtractor.exe   chương trình chính (chỉ chạy từ trong thư mục này)
  _internal\            thư viện – KHÔNG sửa, KHÔNG xóa, KHÔNG lưu dữ liệu vào đây
  config\config.json    cấu hình (tự tạo khi lưu; có thể xóa để về mặc định)
  logs\                 nhật ký khởi động (app.log, errors.log, startup_error.log)
  Output\               nơi gợi ý lưu file Excel kết quả (bạn chọn vị trí khác cũng được)
  README.txt / FIRST_RUN.txt / VERSION.txt / SHA256SUMS.txt
  Install_Ollama_Optional.bat   cài Ollama + model (tùy chọn, tải từ trang chính thức)

Có thể sao chép cả thư mục sang USB hoặc máy khác; đường dẫn có dấu cách hoặc
tiếng Việt đều dùng được.

3. OLLAMA (TÙY CHỌN)
--------------------
Chương trình KHÔNG kèm Ollama hay model. Không có Ollama vẫn chạy bình thường
(chế độ heuristic – AI chỉ dùng để phân loại slide, không bao giờ tạo nội dung Excel).
  - Mặc định: http://127.0.0.1:11434, model qwen3:4b.
  - Dùng Ollama trên máy khác trong mạng LAN:
      a) Trên máy chạy Ollama: đặt biến môi trường OLLAMA_HOST=0.0.0.0 rồi khởi động lại Ollama,
         mở cổng 11434 trên firewall của máy đó.
      b) Trên máy này: tab "Cấu hình & Ollama" → "Tìm Ollama trong mạng LAN" → chọn server →
         "Sử dụng server này" (chương trình không bao giờ tự đổi server), hoặc nhập IP thủ công.
  - Chương trình chỉ quét cổng 11434 trong dải mạng LAN nội bộ của máy, chỉ khi bạn bấm nút.

4. KHẮC PHỤC SỰ CỐ
-------------------
  - Không khởi động được: xem logs\startup_error.log.
  - "Model qwen3:4b chưa có trên server này": chọn model có sẵn trong danh sách hoặc chạy
    `ollama pull qwen3:4b` trên máy server.
  - Không tìm thấy Ollama trong LAN: kiểm tra máy Ollama đã chạy, OLLAMA_HOST=0.0.0.0, firewall,
    hai máy cùng dải mạng; hoặc nhập địa chỉ thủ công.
  - Thư mục chỉ đọc (Program Files, ổ mạng không có quyền ghi): chuyển thư mục sang nơi khác.
  - Kiểm tra tính toàn vẹn: so sánh SHA256 trong SHA256SUMS.txt
      certutil -hashfile ReportExtractor.exe SHA256

5. THÔNG TIN BẢN DỰNG
---------------------
  Phiên bản: {version}   Prompt: {build_id}   Build: {built}   Git: {git}
  Đóng gói: PyInstaller onedir (windowed, không UPX)   Python: {python}
