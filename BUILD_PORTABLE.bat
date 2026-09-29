@echo off
setlocal
chcp 65001 >nul
title Report Extractor - Build portable
cd /d "%~dp0"

echo ============================================================
echo  Report Extractor - BUILD WINDOWS PORTABLE
echo ============================================================

where py >nul 2>nul && (set "PY=py -3") || (set "PY=python")
%PY% --version >nul 2>nul || (
    echo [ERROR] Python 3.9+ khong duoc tim thay tren may build. Cai Python truoc, tick "Add to PATH".
    pause & exit /b 1
)

if not exist ".venv\Scripts\python.exe" (
    echo [1/5] Tao virtual environment .venv ...
    %PY% -m venv .venv || (echo [ERROR] Khong tao duoc venv & pause & exit /b 1)
) else (
    echo [1/5] Dung lai .venv co san
)
set "VPY=.venv\Scripts\python.exe"

echo [2/5] Cai dat thu vien build ...
"%VPY%" -m pip install --upgrade pip >nul
"%VPY%" -m pip install -r requirements-dev.txt || (echo [ERROR] pip install that bai & pause & exit /b 1)

echo [3/5] Chay test nhanh ...
"%VPY%" -m pytest -q tests || (
    echo [WARN] Co test that bai. Nhan phim bat ky de van tiep tuc build, hoac Ctrl+C de dung.
    pause
)

echo [4/5] Build PyInstaller (onedir) ...
if exist "build" rmdir /s /q "build"
if exist "dist\ReportExtractor_Portable" rmdir /s /q "dist\ReportExtractor_Portable"
"%VPY%" -m PyInstaller --noconfirm --clean ReportExtractor.spec || (echo [ERROR] PyInstaller that bai & pause & exit /b 1)

echo [5/5] Hoan thien thu muc portable ...
if not exist "dist\ReportExtractor_Portable\ReportExtractor.exe" (
    echo [ERROR] Khong thay dist\ReportExtractor_Portable\ReportExtractor.exe
    pause & exit /b 1
)
if not exist "dist\ReportExtractor_Portable\config.json" (
    > "dist\ReportExtractor_Portable\config.json" echo {"ollama_server": "http://192.168.1.50:11434", "model": "", "last_report_folder": "", "last_template": "", "last_output_folder": ""}
)
copy /y README.md "dist\ReportExtractor_Portable\README.md" >nul

echo.
echo Kiem tra khoi dong (diagnostics) ...
"dist\ReportExtractor_Portable\ReportExtractor.exe" --diag --no-ollama-check
if errorlevel 1 (
    echo [WARN] Exe chay --diag tra ve loi. Xem thong bao phia tren.
) else (
    echo [OK] Exe khoi dong duoc.
)

echo.
echo ============================================================
echo  DONE:  dist\ReportExtractor_Portable\ReportExtractor.exe
echo  Copy ca thu muc  dist\ReportExtractor_Portable  sang may khac.
echo ============================================================
pause
endlocal
