@echo off
setlocal

:: ============================================================
::  ARCA - Facturación Electrónica
::  Inicia el servidor si no está corriendo y abre Chrome
::  (versión usuario — sin servicio de Windows)
:: ============================================================

set PUERTO=5000
set APP_DIR=%~dp0
set APP_DIR=%APP_DIR:~0,-1%
set URL=http://localhost:%PUERTO%

:: --- Verificar si el servidor ya está corriendo ---
netstat -an | findstr ":%PUERTO% " | findstr "LISTENING" >nul 2>&1
if %errorlevel% == 0 (
    echo Servidor ya en ejecucion. Abriendo Chrome...
    goto abrir_chrome
)

:: --- Verificar Python ---
where python >nul 2>&1
if %errorlevel% neq 0 (
    echo ERROR: Python no encontrado.
    echo Ejecuta instalar_usuario.bat primero.
    pause
    exit /b 1
)

:: --- Iniciar servidor en segundo plano sin ventana ---
echo Iniciando servidor ARCA...
start "" /B pythonw "%APP_DIR%\app.py"

:: --- Esperar hasta 15 segundos a que el servidor responda ---
set INTENTOS=0
:esperar
set /a INTENTOS+=1
if %INTENTOS% gtr 15 (
    echo El servidor tardo demasiado. Abriendo Chrome igual...
    goto abrir_chrome
)
timeout /t 1 /nobreak >nul
netstat -an | findstr ":%PUERTO% " | findstr "LISTENING" >nul 2>&1
if %errorlevel% neq 0 goto esperar
echo Servidor listo.

:: --- Abrir Chrome ---
:abrir_chrome
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
