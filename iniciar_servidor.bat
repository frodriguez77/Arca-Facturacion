@echo off
setlocal

:: ============================================================
::  ARCA - Facturación Electrónica
::  Inicia el servidor si no está corriendo y abre Chrome
:: ============================================================

set PUERTO=5000
set APP_DIR=%~dp0
set APP_DIR=%APP_DIR:~0,-1%

:: --- Verificar si el servidor ya está corriendo en el puerto ---
netstat -an | findstr ":%PUERTO% " | findstr "LISTENING" >nul 2>&1
if %errorlevel% == 0 (
    echo El servidor ya esta corriendo en el puerto %PUERTO%.
    goto abrir_chrome
)

:: --- Buscar Python ---
where python >nul 2>&1
if %errorlevel% neq 0 (
    echo ERROR: Python no encontrado. Instala Python desde https://www.python.org
    pause
    exit /b 1
)

:: --- Iniciar servidor en segundo plano (sin ventana visible) ---
echo Iniciando servidor ARCA...
start "" /B /MIN pythonw "%APP_DIR%\app.py"

:: --- Esperar a que el servidor esté listo (máx 15 segundos) ---
set INTENTOS=0
:esperar
set /a INTENTOS+=1
if %INTENTOS% gtr 15 (
    echo ERROR: El servidor no respondio a tiempo.
    echo Intentando abrir Chrome de todas formas...
    goto abrir_chrome
)
timeout /t 1 /nobreak >nul
netstat -an | findstr ":%PUERTO% " | findstr "LISTENING" >nul 2>&1
if %errorlevel% neq 0 goto esperar
echo Servidor listo.

:: --- Abrir Chrome ---
:abrir_chrome
set URL=http://localhost:%PUERTO%

:: Buscar Chrome en rutas conocidas
set CHROME=""
if exist "%ProgramFiles%\Google\Chrome\Application\chrome.exe" (
    set CHROME="%ProgramFiles%\Google\Chrome\Application\chrome.exe"
) else if exist "%ProgramFiles(x86)%\Google\Chrome\Application\chrome.exe" (
    set CHROME="%ProgramFiles(x86)%\Google\Chrome\Application\chrome.exe"
) else if exist "%LocalAppData%\Google\Chrome\Application\chrome.exe" (
    set CHROME="%LocalAppData%\Google\Chrome\Application\chrome.exe"
)

if not %CHROME% == "" (
    start "" %CHROME% --new-window "%URL%"
) else (
    :: Si no hay Chrome, abrir con el navegador predeterminado
    start "" "%URL%"
)

endlocal
