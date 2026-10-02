@echo off
cd /d "%~dp0"
if exist ".venv\Scripts\python.exe" (set "VPY=.venv\Scripts\python.exe") else (set "VPY=python")
"%VPY%" -m pytest -q tests
pause
