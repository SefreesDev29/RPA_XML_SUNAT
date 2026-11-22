@echo off
title RPA SUNAT - TEST
cd /d "%~dp0"

set APP_ENV=TEST

echo =======================================================
echo  EJECUTANDO AUTOMATIZACION EN ENTORNO: %APP_ENV%
echo =======================================================

call .venv\Scripts\python.exe Main.py

if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] El proceso finalizo con codigo de error: %ERRORLEVEL%
    pause
    exit /b %ERRORLEVEL%
)

echo.
echo Proceso finalizado en TEST.
pause