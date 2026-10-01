@echo off
REM ============================================================
REM  Report Extractor - cai dat tu source (Windows)
REM  Idempotent: chay nhieu lan an toan. Khong can quyen Admin
REM  (tru khi winget/Ollama installer tu yeu cau).
REM ============================================================
setlocal EnableExtensions EnableDelayedExpansion
chcp 65001 >nul
cd /d "%~dp0"
set "ROOT=%~dp0"
if "%ROOT:~-1%"=="\" set "ROOT=%ROOT:~0,-1%"
set "VENV_PY=%ROOT%\.venv\Scripts\python.exe"
set "SUPPORT=%ROOT%\tools\setup_support.py"
set "APP_OK=0"
set "OLLAMA_STATE=Không cài / dùng máy khác"

echo ================================================
echo        REPORT EXTRACTOR - CÀI ĐẶT TỪ SOURCE
echo ================================================
echo Thư mục: "%ROOT%"
echo.

REM ---------- 1. Tim Python phu hop ----------------------------
set "SYS_PY="
call :find_python
if not defined SYS_PY (
    echo Không tìm thấy Python phù hợp trên máy.
    where winget >nul 2>nul
    if errorlevel 1 (
        echo Máy không có winget. Hãy cài Python 3.12 từ https://www.python.org/downloads/windows/
        echo ^(tích "Add python.exe to PATH" và "tcl/tk"^) rồi chạy lại setup.bat.
        goto :fail
    )
    set /p "ANS=Cài Python 3.12 bằng winget ^(official Python.org package^)? [Y/n]: "
    if /i "!ANS!"=="n" (
        echo Hãy cài Python 3.12 rồi chạy lại setup.bat.
        goto :fail
    )
    winget install --id Python.Python.3.12 -e --source winget --accept-package-agreements --accept-source-agreements
    if errorlevel 1 (
        echo Cài Python bằng winget thất bại. Hãy cài thủ công rồi chạy lại.
        goto :fail
    )
    REM lam moi PATH cua phien hien tai de tim python vua cai
    call :refresh_path
    call :find_python
    if not defined SYS_PY (
        echo Đã cài Python nhưng chưa nhận diện được. Hãy mở lại cửa sổ và chạy setup.bat lần nữa.
        goto :fail
    )
)
echo Python hệ thống: !SYS_PY!
echo.

REM ---------- 2. Tao / kiem tra .venv ---------------------------
if exist "%VENV_PY%" (
    "%VENV_PY%" -c "import sys; sys.exit(0)" >nul 2>nul
    if errorlevel 1 (
        echo .venv bị hỏng hoặc trỏ tới Python không còn tồn tại.
        set /p "ANS=Tạo lại .venv? [Y/n]: "
        if /i "!ANS!"=="n" goto :fail
        rmdir /s /q "%ROOT%\.venv"
    ) else (
        "%VENV_PY%" "%SUPPORT%" check-python >nul 2>nul
        if errorlevel 1 (
            echo .venv dùng phiên bản Python không phù hợp – tạo lại.
            rmdir /s /q "%ROOT%\.venv"
        )
    )
)
if not exist "%VENV_PY%" (
    echo Đang tạo môi trường ảo .venv ...
    !SYS_PY! -m venv "%ROOT%\.venv"
    if errorlevel 1 (
        echo Tạo .venv thất bại.
        goto :fail
    )
) else (
    echo Môi trường .venv: dùng lại.
)
"%VENV_PY%" "%SUPPORT%" check-python
if errorlevel 1 goto :fail
echo.

REM ---------- 3. Thu vien runtime -------------------------------
echo Đang cài / kiểm tra thư viện ứng dụng ^(requirements.txt^) ...
"%VENV_PY%" -m pip install --upgrade pip --quiet --disable-pip-version-check
"%VENV_PY%" -m pip install -r "%ROOT%\requirements.txt" --disable-pip-version-check
if errorlevel 1 (
    echo Cài thư viện runtime thất bại ^(kiểm tra kết nối Internet / proxy^).
    goto :fail
)
set /p "ANS=Cài thêm công cụ phát triển/test? [y/N]: "
if /i "!ANS!"=="y" (
    "%VENV_PY%" -m pip install -r "%ROOT%\requirements-dev.txt" --disable-pip-version-check
    if errorlevel 1 echo Cảnh báo: cài công cụ phát triển thất bại ^(ứng dụng vẫn chạy được^).
)
echo.

REM ---------- 4. Kiem tra ung dung -----------------------------
echo Kiểm tra ứng dụng ...
"%VENV_PY%" "%SUPPORT%" verify
if errorlevel 1 (
    echo Kiểm tra ứng dụng thất bại – xem thông báo phía trên.
    goto :fail
)
set "APP_OK=1"
echo.

REM ---------- 5. Ollama (tuy chon) -----------------------------
where ollama >nul 2>nul
if not errorlevel 1 (
    echo Ollama: Đã cài đặt
    set "OLLAMA_STATE=Đã cài"
    goto :models
)
echo Ollama chưa được cài đặt.
echo.
echo [1] Cài Ollama ^(winget, gói chính thức Ollama.Ollama^)
echo [2] Bỏ qua
set /p "ANS=Lựa chọn [2]: "
if "!ANS!"=="1" (
    where winget >nul 2>nul
    if errorlevel 1 (
        echo Máy không có winget. Tải Ollama tại https://ollama.com/download ^(trang chính thức^).
        goto :no_ollama
    )
    winget install --id Ollama.Ollama -e --source winget --accept-package-agreements --accept-source-agreements
    if errorlevel 1 (
        echo Cài Ollama thất bại. Ứng dụng vẫn dùng được với Ollama trên máy khác.
        goto :no_ollama
    )
    call :refresh_path
    where ollama >nul 2>nul
    if errorlevel 1 (
        echo Đã cài Ollama nhưng chưa nhận diện được lệnh ollama. Mở lại cửa sổ và chạy setup.bat để tải model.
        goto :no_ollama
    )
    set "OLLAMA_STATE=Đã cài"
    goto :models
)
:no_ollama
echo Bạn có thể cấu hình Ollama trên máy khác trong tab "Cấu hình & Ollama".
goto :summary

REM ---------- 6. Chon model -------------------------------------
:models
echo.
echo ================================================
echo        CHỌN MODEL OLLAMA
echo ================================================
echo.
echo [1] qwen3:1.7b
echo [2] qwen3:4b ^(mặc định/khuyến nghị^)
echo [3] Cả hai
echo [4] Không tải model
echo.
set /p "ANS=Lựa chọn [4]: "
set "SEL="
if "!ANS!"=="1" set "SEL=qwen3:1.7b"
if "!ANS!"=="2" set "SEL=qwen3:4b"
if "!ANS!"=="3" set "SEL=qwen3:1.7b qwen3:4b"
if not defined SEL (
    echo Không tải model. Có thể tải sau bằng: ollama pull qwen3:4b
    goto :summary
)
echo.
echo Model AI có thể cần tải dữ liệu dung lượng lớn. Hãy đảm bảo máy có đủ dung lượng và kết nối Internet.
echo.
set "PULL_FAILED="
for /f "usebackq delims=" %%M in (`"%VENV_PY%" "%SUPPORT%" models-missing !SEL!`) do (
    echo Đang tải %%M ...
    ollama pull %%M
    if errorlevel 1 (
        echo Tải model %%M thất bại.
        set "PULL_FAILED=!PULL_FAILED! %%M"
    )
)
if "!ANS!"=="1" "%VENV_PY%" "%SUPPORT%" set-model qwen3:1.7b
if "!ANS!"=="2" "%VENV_PY%" "%SUPPORT%" set-model qwen3:4b
if defined PULL_FAILED echo Ứng dụng đã cài xong, nhưng tải model!PULL_FAILED! thất bại.

REM ---------- 7. Tong ket ---------------------------------------
:summary
echo.
echo ================================================
echo           CÀI ĐẶT HOÀN TẤT
echo ================================================
"%VENV_PY%" "%SUPPORT%" summary
echo.
echo Cài đặt ứng dụng: HOÀN TẤT
echo AI ^(Ollama/model^) có thể cấu hình sau trong tab "Cấu hình & Ollama".
echo.
echo Khởi chạy:
echo   run.bat
echo ================================================
echo.
pause
endlocal
exit /b 0

:fail
echo.
echo ================================================
echo           CÀI ĐẶT THẤT BẠI
echo ================================================
echo Xem thông báo lỗi phía trên rồi chạy lại setup.bat.
echo Nhấn phím bất kỳ để thoát...
pause >nul
endlocal
exit /b 1

REM ---------- helpers -------------------------------------------
:find_python
REM Uu tien: py -3.12, py -3.13, py -3.14, py -3.11, py -3.10, py -3, python ; phai qua check-python
for %%V in (-3.12 -3.13 -3.14 -3.11 -3.10 -3) do (
    if not defined SYS_PY (
        py %%V "%SUPPORT%" check-python >nul 2>nul
        if not errorlevel 1 set "SYS_PY=py %%V"
    )
)
if not defined SYS_PY (
    python "%SUPPORT%" check-python >nul 2>nul
    if not errorlevel 1 set "SYS_PY=python"
)
exit /b 0

:refresh_path
for /f "usebackq tokens=2,*" %%A in (`reg query "HKCU\Environment" /v Path 2^>nul`) do set "UPATH=%%B"
for /f "usebackq tokens=2,*" %%A in (`reg query "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment" /v Path 2^>nul`) do set "SPATH=%%B"
set "PATH=%SPATH%;%UPATH%;%PATH%"
exit /b 0
