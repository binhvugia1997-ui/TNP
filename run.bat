@echo off
REM Report Extractor - khoi chay giao dien tu source (.venv\Scripts\python.exe -> run.py)
setlocal
chcp 65001 >nul
cd /d "%~dp0"
set "VENV_PY=%~dp0.venv\Scripts\python.exe"
if not exist "%VENV_PY%" (
    echo Chưa cài đặt môi trường. Hãy chạy setup.bat trước ^(tạo .venv và cài thư viện^).
    echo Nhấn phím bất kỳ để thoát...
    pause >nul
    exit /b 1
)
"%VENV_PY%" "%~dp0run.py" %*
if errorlevel 1 (
    echo.
    echo Ứng dụng kết thúc với lỗi. Xem Output\logs\errors.log hoặc chạy lại setup.bat.
    echo Nhấn phím bất kỳ để thoát...
    pause >nul
    exit /b 1
)
endlocal
