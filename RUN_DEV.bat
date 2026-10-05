@echo off
setlocal
chcp 65001 >nul
cd /d "%~dp0"
where py >nul 2>nul && (set "PY=py -3") || (set "PY=python")
if not exist ".venv\Scripts\python.exe" (
    echo Tao .venv va cai thu vien ...
    %PY% -m venv .venv || (pause & exit /b 1)
    ".venv\Scripts\python.exe" -m pip install --upgrade pip >nul
    ".venv\Scripts\python.exe" -m pip install -r requirements-dev.txt || (pause & exit /b 1)
)
".venv\Scripts\python.exe" run.py %*
if errorlevel 1 pause
endlocal
