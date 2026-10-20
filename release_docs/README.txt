REPORT EXTRACTOR {version} – Build {build_id} – BẢN PORTABLE WINDOWS
=====================================================================

Chuyển báo cáo PowerPoint (PPTX) thành Bảng kiểm chứng Excel, hoàn toàn chạy
trên máy nội bộ. Không cần cài đặt, không cần quyền Administrator.

1. MÁY ĐÍCH CẦN GÌ
------------------
  - Windows 10/11 64-bit. KHÔNG cần Python, Node.js, npm, Git hay mã nguồn.
  - Microsoft Edge WebView2 Runtime (đã có sẵn trên Windows 11 và hầu hết Windows 10).
    Nếu cửa sổ hiện ra trắng/rỗng: cài "Microsoft Edge WebView2 Runtime" từ trang chính
    thức của Microsoft rồi chạy lại.
  - .NET Framework 4.7.2 trở lên (khuyến nghị 4.8). Đây là yêu cầu THẬT SỰ của
    pywebview trên Windows: cửa sổ WebView2 được tạo qua pythonnet, và
    Python.Runtime.dll của pythonnet biên dịch cho .NETStandard 2.0 nên cần .NET
    Framework 4.7.2+. Windows 10 1803+ và Windows 11 đã có sẵn. Nếu thiếu, chương
    trình KHÔNG im lặng: logs\app.log có dòng WEBVIEW_RUNTIME ... dotnet_framework=...
    và logs\startup_error.log ghi rõ "cần .NET Framework 4.7.2+".
  - Microsoft PowerPoint là TÙY CHỌN. Có PowerPoint → dùng PowerPoint để render slide
    đúng nhất. Không có PowerPoint → chương trình vẫn chạy, tự chuyển sang LibreOffice
    hoặc trình vẽ tích hợp, và báo rõ backend đã dùng (không bao giờ im lặng).

2. CÁCH DÙNG NHANH
------------------
  1) Giải nén toàn bộ thư mục {name} vào một nơi bạn có quyền ghi
     (ví dụ D:\ReportExtractor hoặc Desktop). KHÔNG đặt trong C:\Program Files.
  2) Chạy ReportExtractor.exe (xem FIRST_RUN.txt nếu Windows hiện cảnh báo SmartScreen).
  3) Tab "Danh sách báo cáo": chọn thư mục báo cáo PPTX, file mẫu Excel, file kết quả → BẮT ĐẦU XỬ LÝ.
  4) Tab "Cài đặt": địa chỉ Ollama (tùy chọn), model, nút "Tìm Ollama trong mạng LAN".
  5) Tab "Học cải tiến": xem trước slide, kiểm tra nhãn ảnh/nội dung cải tiến.
  6) Phiên bản đang chạy hiện ngay trên thanh tiêu đề (vd. "vX.Y.Z · Build NNN") và trong
     Tab "Cài đặt" → "Phiên bản hiện tại". Hãy kiểm tra SỐ NÀY TRƯỚC khi báo cáo lỗi.

3. CẤU TRÚC THƯ MỤC
-------------------
  ReportExtractor.exe   chương trình chính (chỉ chạy từ trong thư mục này)
  _internal\            thư viện Python + pywebview/WebView2 – KHÔNG sửa, KHÔNG xóa,
                        KHÔNG lưu dữ liệu vào đây
  frontend\dist\        giao diện React đã build (index.html + assets\). Chương trình phục vụ
                        thư mục này qua HTTP nội bộ (loopback). Thiếu nó → cửa sổ trắng;
                        xem logs\app.log, dòng WEBVIEW_FRONTEND liệt kê packaged/app_root/dist/url
  config\config.json    cấu hình (tự tạo khi lưu; có thể xóa để về mặc định)
  logs\                 nhật ký khởi động (app.log, errors.log, startup_error.log)
  Output\               nơi gợi ý lưu file Excel kết quả (bạn chọn vị trí khác cũng được)
  README.txt / FIRST_RUN.txt / VERSION.txt / SHA256SUMS.txt
  Install_Ollama_Optional.bat   cài Ollama + model (tùy chọn, tải từ trang chính thức)

Có thể sao chép cả thư mục sang USB hoặc máy khác; đường dẫn có dấu cách hoặc
tiếng Việt đều dùng được.

4. OLLAMA (TÙY CHỌN)
--------------------
Chương trình KHÔNG kèm Ollama hay model. Không có Ollama vẫn chạy bình thường
(chế độ heuristic – AI chỉ dùng để phân loại slide, không bao giờ tạo nội dung Excel).
  - Mặc định: http://127.0.0.1:11434, model qwen3:4b.
  - Dùng Ollama trên máy khác trong mạng LAN:
      a) Trên máy chạy Ollama: đặt biến môi trường OLLAMA_HOST=0.0.0.0 rồi khởi động lại Ollama,
         mở cổng 11434 trên firewall của máy đó.
      b) Trên máy này: tab "Cài đặt" → "Tìm Ollama trong mạng LAN" → chọn server →
         "Sử dụng server này" (chương trình không bao giờ tự đổi server), hoặc nhập IP thủ công.
  - Chương trình chỉ quét cổng 11434 trong dải mạng LAN nội bộ của máy, chỉ khi bạn bấm nút.

5. KHẮC PHỤC SỰ CỐ
-------------------
  - Không khởi động được: xem logs\startup_error.log, và dòng WEBVIEW_RUNTIME trong
    logs\app.log (ghi TRƯỚC khi tạo cửa sổ: phiên bản python/pywebview/pythonnet/
    clr-loader, backend, đường dẫn Python.Runtime.dll đã tìm thấy hay chưa, và
    dotnet_framework=Release=... (ok/TOO-OLD)).
  - Lỗi "Failed to resolve Python.Runtime.Loader.Initialize": Python.Runtime.dll không
    nạp được. Ba nguyên nhân theo thứ tự kiểm tra: (1) thiếu file
    _internal\pythonnet\runtime\Python.Runtime.dll; (2) file đó khác bản trong môi
    trường build (.venv-build cũ); (3) .NET Framework < 4.7.2. Xem đúng dòng
    WEBVIEW_RUNTIME để biết là nguyên nhân nào, không phải đoán.
  - Cửa sổ mở ra nhưng TRẮNG/RỖNG: thiếu Microsoft Edge WebView2 Runtime, hoặc thiếu
    frontend\dist\ (xem logs\app.log, dòng WEBVIEW_FRONTEND).
  - Cần giao diện Tk cổ điển để so sánh: chạy `ReportExtractor.exe --legacy-gui`.
  - In phiên bản: `ReportExtractor.exe --version`.  Chẩn đoán hệ thống: `--diag`.
  - "Model qwen3:4b chưa có trên server này": chọn model có sẵn trong danh sách hoặc chạy
    `ollama pull qwen3:4b` trên máy server.
  - Không tìm thấy Ollama trong LAN: kiểm tra máy Ollama đã chạy, OLLAMA_HOST=0.0.0.0, firewall,
    hai máy cùng dải mạng; hoặc nhập địa chỉ thủ công.
  - Thư mục chỉ đọc (Program Files, ổ mạng không có quyền ghi): chuyển thư mục sang nơi khác.
  - Kiểm tra tính toàn vẹn: so sánh SHA256 trong SHA256SUMS.txt
      certutil -hashfile ReportExtractor.exe SHA256

6. CẬP NHẬT PHẦN MỀM (OFFLINE / MẠNG NỘI BỘ)
  - Tab "Cài đặt" → "Cập nhật phần mềm": nhập "Đường dẫn cập nhật" (vd. D:\ReportExtractor_Update
    hoặc \\SERVER\ReportExtractor\Update) → "Kiểm tra cập nhật" → nếu có bản mới bấm "Cập nhật ngay".
  - Thư mục cập nhật chứa version.json + ReportExtractor_<phiên bản>.zip. Không cần Git/Internet/Python.
  - Gói được sao chép về update_staging\, kiểm tra SHA256, rồi trình cập nhật thay thế file chương trình,
    bản cũ giữ trong update_backup\ để khôi phục nếu lỗi. config\, logs\, Output\, file Excel KHÔNG bị đụng tới.
  - Dùng quyền truy cập Windows hiện có; chương trình không hỏi / không lưu mật khẩu. Nhật ký: logs\update.log.

7. THÔNG TIN BẢN DỰNG
---------------------
  Phiên bản: {version}   Build: {build_id}   Ngày build: {built}   Git (nội bộ): {git}
  Đóng gói: PyInstaller onedir (windowed, không UPX)   Python: {python}
  Giao diện: React + pywebview (Microsoft Edge WebView2)   PowerPoint: tùy chọn
  Build frontend: npm run build (chỉ trên máy BUILD; máy đích không cần Node.js)
