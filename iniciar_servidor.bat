@echo off
setlocal

:: ============================================================
::  ARCA - Facturación Electrónica
::  Abre el sistema en Chrome
::  El servidor corre como servicio de Windows (siempre activo)
:: ============================================================

set PUERTO=5000
set APP_DIR=%~dp0
set APP_DIR=%APP_DIR:~0,-1%
set SERVICE_NAME=ArcaFacturacion
set URL=http://localhost:%PUERTO%

:: --- Verificar si el servidor está respondiendo ---
netstat -an | findstr ":%PUERTO% " | findstr "LISTENING" >nul 2>&1
if %errorlevel% == 0 goto abrir_chrome

:: --- El servidor no está corriendo, intentar iniciar el servicio ---
echo Iniciando servicio ARCA...
sc start %SERVICE_NAME% >nul 2>&1

:: Esperar hasta 15 segundos
set INTENTOS=0
:esperar
set /a INTENTOS+=1
if %INTENTOS% gtr 15 (
    echo El servicio no respondio. Intentando abrir igual...
    goto abrir_chrome
)
timeout /t 1 /nobreak >nul
netstat -an | findstr ":%PUERTO% " | findstr "LISTENING" >nul 2>&1
if %errorlevel% neq 0 goto esperar

:abrir_chrome
:: Buscar Chrome en rutas conocidas
set CHROME=
if exist "%ProgramFiles%\Google\Chrome\Application\chrome.exe" (
    set CHROME="%ProgramFiles%\Google\Chrome\Application\chrome.exe"
) else if exist "%ProgramFiles(x86)%\Google\Chrome\Application\chrome.exe" (
    set CHROME="%ProgramFiles(x86)%\Google\Chrome\Application\chrome.exe"
) else if exist "%LocalAppData%\Google\Chrome\Application\chrome.exe" (
    set CHROME="%LocalAppData%\Google\Chrome\Application\chrome.exe"
)

if defined CHROME (
    start "" %CHROME% --new-window "%URL%"
) else (
    start "" "%URL%"
)

endlocal
