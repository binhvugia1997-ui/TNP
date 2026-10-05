@echo off
setlocal
chcp 65001 >nul
title Report Extractor - Cai Ollama (tuy chon)
echo ============================================================
echo  Cai Ollama + model qwen3:4b (TUY CHON)
echo ============================================================
echo.
echo Report Extractor chay duoc khong can Ollama (che do heuristic).
echo Ollama chi giup phan loai slide chinh xac hon. Ban co the:
echo   - cai Ollama tren may nay (script nay), hoac
echo   - dung Ollama tren may khac trong LAN (nut "Tim Ollama trong mang LAN" trong tab Cau hinh).
echo.
echo Script nay KHONG di kem Ollama hay model; no tai tu trang chinh thuc https://ollama.com
echo.
set /p OK=Tiep tuc cai Ollama bang winget? (Y/N):
if /i not "%OK%"=="Y" goto :end

where ollama >nul 2>nul
if %ERRORLEVEL%==0 (
    echo Ollama da co san tren may.
    goto :model
)
where winget >nul 2>nul
if not %ERRORLEVEL%==0 (
    echo May khong co winget. Tai Ollama thu cong tai https://ollama.com/download roi chay lai script nay.
    goto :end
)
winget install --id Ollama.Ollama -e --accept-source-agreements --accept-package-agreements
where ollama >nul 2>nul
if not %ERRORLEVEL%==0 (
    echo Da cai nhung chua nhan lenh ollama. Mo lai cua so va chay lai script nay de tai model.
    goto :end
)

:model
set /p OK2=Tai model qwen3:4b (~2.5 GB)? (Y/N):
if /i not "%OK2%"=="Y" goto :end
ollama pull qwen3:4b
echo.
echo Xong. Trong Report Extractor: tab "Cau hinh ^& Ollama" -^> Server 127.0.0.1, Port 11434, Model qwen3:4b -^> "Kiem tra ket noi".

:end
echo.
pause
exit /b 0
