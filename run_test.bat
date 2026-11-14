@echo off
title RPA SUNAT - TEST
cd /d "%~dp0"

set APP_ENV=TEST

echo =======================================================
echo  EJECUTANDO AUTOMATIZACION EN ENTORNO: %APP_ENV%
echo =======================================================

call .venv\Scripts\python.exe Main.py

echo.
echo Proceso finalizado en TEST.
pause