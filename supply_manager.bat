@echo off
title GESTOR DE APROVISIONAMIENTO - RPA XML SUNAT
cd /d "%~dp0"

:MENU
cls
echo ====================================================================
echo        PANEL DE APROVISIONAMIENTO Y DESPLIEGUE - RPA SUNAT
echo ====================================================================
echo.
echo   [0] Aprovisionar Credenciales para DEV   (Namespace: RPA_SUNAT_DEV)
echo   [1] Aprovisionar Credenciales para TEST  (Namespace: RPA_SUNAT_TEST)
echo   [2] Aprovisionar Credenciales para PROD  (Namespace: RPA_SUNAT_PROD)
echo   [3] Salir
echo.
echo ====================================================================
echo.

choice /c 0123 /n /m "Seleccione una opcion [0, 1, 2 o 3]: "

if errorlevel 4 goto SALIR
if errorlevel 3 goto ENV_PROD
if errorlevel 2 goto ENV_TEST
if errorlevel 1 goto ENV_DEV

:: ----------------------------------------------------------------------
:: OPCION 0: APROVISIONAR DEV
:: ----------------------------------------------------------------------
:ENV_DEV
cls
set APP_ENV=DEV
echo ====================================================================
echo  MODO SELECCIONADO: DESARROLLO (DEV)
echo ====================================================================
call .venv\Scripts\python.exe setup_credentials.py
goto PAUSA_MENU

:: ----------------------------------------------------------------------
:: OPCION 1: APROVISIONAR TEST
:: ----------------------------------------------------------------------
:ENV_TEST
cls
set APP_ENV=TEST
echo ====================================================================
echo  MODO SELECCIONADO: TEST
echo ====================================================================
call .venv\Scripts\python.exe setup_credentials.py
goto PAUSA_MENU

:: ----------------------------------------------------------------------
:: OPCION 2: APROVISIONAR PROD
:: ----------------------------------------------------------------------
:ENV_PROD
cls
set APP_ENV=PROD
echo ====================================================================
echo  MODO SELECCIONADO: PRODUCCION (PROD)
echo ====================================================================
call .venv\Scripts\python.exe setup_credentials.py
goto PAUSA_MENU

:: ----------------------------------------------------------------------
:: CONTROL DE RETORNO Y SALIDA
:: ----------------------------------------------------------------------
:PAUSA_MENU
echo.
echo --------------------------------------------------------------------
echo Operacion finalizada. Presione cualquier tecla para volver al menu...
pause >nul
goto MENU

:SALIR
cls
exit