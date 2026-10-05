@echo off
setlocal
chcp 65001 >nul
title Report Extractor - BUILD + PUBLISH (one click)
cd /d "%~dp0"

echo ============================================================
echo  Report Extractor - BUILD + PUBLISH tu GitHub (mot cu nhap)
echo    1. git fetch / pull --ff-only (an toan, khong ghi de sua doi)
echo    2. pyflakes + pytest + PyInstaller + kiem tra goi
echo    3. Portable + ZIP + SHA256SUMS.txt + version.json
echo    4. Xuat ban vao thu muc cap nhat LAN (ZIP truoc, version.json sau cung)
echo ============================================================
echo.

rem --- find a supported Python (3.10 - 3.13, 64-bit) --- same policy as setup.bat / build_portable.bat
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
where git >nul 2>nul
if errorlevel 1 (
    echo LOI: Khong tim thay git. Cai Git for Windows tu https://git-scm.com/download/win roi chay lai.
    pause
    exit /b 2
)
echo Dung Python: %PY%

%PY% -c "import sys; print(sys.executable)" > "%TEMP%\re_bap_py.txt"
set /p PYEXE=<"%TEMP%\re_bap_py.txt"
del "%TEMP%\re_bap_py.txt" >nul 2>nul
%PY% tools\build_and_publish.py --python "%PYEXE%" %*
set "RC=%ERRORLEVEL%"
echo.
if "%RC%"=="0" (
    echo HOAN TAT: da build va xuat ban vao thu muc cap nhat LAN. Cac may lam viec se tu phat hien ban moi.
) else if "%RC%"=="2" (
    echo DUNG O BUOC GIT: xem huong dan phia tren ^(commit/push, dung nhanh, hoac xu ly phan nhanh^) roi chay lai.
) else if "%RC%"=="4" (
    echo DUNG: so Build nay da duoc xuat ban. Tang BUILD_NUMBER va push, hoac chay voi --force.
) else if "%RC%"=="3" (
    echo BUILD THANH CONG nhung XUAT BAN THAT BAI. Goi nam trong release\. Xuat ban lai: %PY% tools\publish_update.py
) else (
    echo BUILD THAT BAI ^(ma %RC%^). Xem thong bao phia tren va file logs\build_and_publish_*.log
)
pause
exit /b %RC%
