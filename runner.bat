@echo off
cd /d "%~dp0"

set "PY=python"
py --version >nul 2>&1 && set "PY=py"

%PY% main.py %*
set "KODE=%ERRORLEVEL%"

echo.
if "%KODE%"=="0" (echo Selesai dengan sukses.) else (echo Error, kode %KODE%.)

if "%~1"=="" pause
exit /b %KODE%