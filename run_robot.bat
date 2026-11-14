@echo off
title RPA SUNAT - PROD
cd /d "%~dp0"

set APP_ENV=PROD

echo =======================================================
echo  EJECUTANDO AUTOMATIZACION EN ENTORNO: %APP_ENV%
echo =======================================================

call .venv\Scripts\python.exe Main.py