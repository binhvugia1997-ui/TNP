@echo off
setlocal
chcp 65001 >nul
title Report Extractor - Build Windows Portable
cd /d "%~dp0"

echo ============================================================
echo  Report Extractor - Build Windows Portable (PyInstaller onedir)
echo ============================================================
echo.

rem --- 1. find a supported Python (same policy as setup.bat: 3.10 - 3.13, 64-bit) ---
set "PY="
for %%V in (3.13 3.12 3.11 3.10) do (
    if not defined PY (
        py -%%V -c "import sys" >nul 2>nul && set "PY=py -%%V"
    )
)
if not defined PY (
    python -c "import sys; v=sys.version_info; sys.exit(0 if (v.major,v.minor) in ((3,10),(3,11),(3,12),(3,13)) else 1)" >nul 2>nul && set "PY=python"
)
if not defined PY (
    echo LOI: Khong tim thay Python 3.10 - 3.13 ^(64-bit^). Cai Python tu https://www.python.org/downloads/windows/
    echo       ^(tick "Add python.exe to PATH" va "py launcher"^).
    pause
    exit /b 1
)
echo Dung Python: %PY%

rem --- 2..11: the Python builder does the rest (venv, pyflakes, pytest, PyInstaller, docs, validate, zip, sha256) ---
%PY% -c "import sys; print(sys.executable)" > "%TEMP%\re_build_py.txt"
set /p PYEXE=<"%TEMP%\re_build_py.txt"
del "%TEMP%\re_build_py.txt" >nul 2>nul
%PY% tools\build_portable.py --python "%PYEXE%" %*
set "RC=%ERRORLEVEL%"
echo.
if "%RC%"=="3" (
    echo BUILD THANH CONG nhung PUBLISH FAILED ^(xem [PUBLISH] phia tren^). Goi nam trong release\.
    echo   Xuat ban lai: %PY% tools\publish_update.py
    pause
    exit /b %RC%
)
if not "%RC%"=="0" (
    echo BUILD THAT BAI ^(ma %RC%^). Xem thong bao phia tren.
    pause
    exit /b %RC%
)
echo BUILD THANH CONG va DA XUAT BAN vao thu muc cap nhat LAN.
echo   Thu muc: dist\ReportExtractor_v^<version^>_Portable
echo   ZIP + SHA256SUMS.txt: release\
pause
exit /b 0
