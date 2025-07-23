@echo off
setlocal

echo 🔧 Setting up project structure...

:: Diretório raiz (ajuste se necessário)
set ROOT=%~dp0

:: Cria pastas necessárias
mkdir "%ROOT%assets" 2>nul
mkdir "%ROOT%macros" 2>nul
mkdir "%ROOT%archive" 2>nul

:: Move arquivos se existirem
if exist "%ROOT%python.ico" move "%ROOT%python.ico" "%ROOT%assets\" >nul
if exist "%ROOT%Dashboard of Operations.xlsm" move "%ROOT%Dashboard of Operations.xlsm" "%ROOT%macros\" >nul
if exist "%ROOT%Dashboard of Operations 2019-12-19.xlsx" move "%ROOT%Dashboard of Operations 2019-12-19.xlsx" "%ROOT%archive\" >nul
if exist "%ROOT%OnePage.py.bak" move "%ROOT%OnePage.py.bak" "%ROOT%archive\" >nul

echo ✅ Project folders and files organized.
pause
endlocal
