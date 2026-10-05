@echo off
REM Report Extractor - CLI (chuyen tiep tham so cho run.py --cli ...). Vi du:
REM   run_cli.bat --cli "D:\Bao_cao" --template "D:\Kiem_chung_09_2026.xlsx" --output "D:\Output\Kiem_chung_09_2026.xlsx" --server 192.168.1.50:11434 --model qwen3:4b
setlocal
chcp 65001 >nul
cd /d "%~dp0"
set "VENV_PY=%~dp0.venv\Scripts\python.exe"
if not exist "%VENV_PY%" (
    echo Chưa cài đặt môi trường. Hãy chạy setup.bat trước.
    exit /b 1
)
"%VENV_PY%" "%~dp0run.py" %*
endlocal
