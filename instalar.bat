@echo off
setlocal EnableDelayedExpansion
chcp 65001 >nul

:: ============================================================
::  ARCA - Facturación Electrónica
::  Script de instalación — ejecutar UNA SOLA VEZ por PC
::  Requiere ejecutar como ADMINISTRADOR
:: ============================================================

:: Verificar que se ejecuta como Administrador
net session >nul 2>&1
if %errorlevel% neq 0 (
    echo.
    echo  ERROR: Este script requiere permisos de Administrador.
    echo  Hacé clic derecho en instalar.bat y elegí "Ejecutar como administrador".
    echo.
    pause
    exit /b 1
)

set APP_DIR=%~dp0
set APP_DIR=%APP_DIR:~0,-1%
set SERVICE_NAME=ArcaFacturacion
set LOG_DIR=%APP_DIR%\logs
set NSSM=%APP_DIR%\tools\nssm.exe
set PUBLIC_DESKTOP=C:\Users\Public\Desktop

echo.
echo  =====================================================
echo   ARCA - Instalacion de Facturacion Electronica
echo  =====================================================
echo.

:: ---- 1. Verificar Python ----------------------------------------
echo [1/5] Verificando Python...
where python >nul 2>&1
if %errorlevel% neq 0 (
    echo.
    echo  ERROR: Python no esta instalado.
    echo  Descargalo desde: https://www.python.org/downloads/
    echo  Asegurate de tildar "Add Python to PATH" al instalar.
    echo.
    pause
    exit /b 1
)
for /f "tokens=*" %%p in ('where python') do set PYTHON_EXE=%%p
for /f "tokens=*" %%v in ('python --version 2^>^&1') do set PYVER=%%v
echo  OK - %PYVER% en %PYTHON_EXE%
echo.

:: ---- 2. Instalar dependencias -----------------------------------
echo [2/5] Instalando dependencias Python...
python -m pip install --upgrade pip --quiet
python -m pip install flask pandas openpyxl zeep requests reportlab qrcode[pil] urllib3 werkzeug --quiet
if %errorlevel% neq 0 (
    echo  ERROR al instalar dependencias.
    pause
    exit /b 1
)
echo  OK - Dependencias instaladas.
echo.

:: ---- 3. Crear carpetas necesarias -------------------------------
echo [3/5] Creando estructura de carpetas...
if not exist "%APP_DIR%\uploads"      mkdir "%APP_DIR%\uploads"
if not exist "%APP_DIR%\certificados" mkdir "%APP_DIR%\certificados"
if not exist "%LOG_DIR%"              mkdir "%LOG_DIR%"
if not exist "%APP_DIR%\tools"        mkdir "%APP_DIR%\tools"
echo  OK - Carpetas listas.
echo.

:: ---- 4. Descargar e instalar NSSM -------------------------------
echo [4/5] Configurando servicio Windows con NSSM...

if not exist "%NSSM%" (
    echo  Descargando NSSM...
    powershell -NoProfile -ExecutionPolicy Bypass -Command ^
        "try { Invoke-WebRequest -Uri 'https://nssm.cc/release/nssm-2.24.zip' -OutFile '%TEMP%\nssm.zip' -UseBasicParsing; Expand-Archive '%TEMP%\nssm.zip' -DestinationPath '%TEMP%\nssm' -Force; Copy-Item '%TEMP%\nssm\nssm-2.24\win64\nssm.exe' '%APP_DIR%\tools\nssm.exe' -Force; Write-Host 'NSSM descargado.' } catch { Write-Host ('ERROR: ' + $_.Exception.Message) }"
    if not exist "%NSSM%" (
        echo.
        echo  ERROR: No se pudo descargar NSSM. Verificá la conexion a internet.
        echo  Descargalo manualmente desde: https://nssm.cc/release/nssm-2.24.zip
        echo  y copiá nssm.exe en: %APP_DIR%\tools\
        echo.
        pause
        exit /b 1
    )
)

:: Detener y eliminar servicio anterior si existe
"%NSSM%" status %SERVICE_NAME% >nul 2>&1
if %errorlevel% == 0 (
    echo  Actualizando servicio existente...
    "%NSSM%" stop %SERVICE_NAME% confirm >nul 2>&1
    "%NSSM%" remove %SERVICE_NAME% confirm >nul 2>&1
)

:: Instalar el servicio
"%NSSM%" install %SERVICE_NAME% "%PYTHON_EXE%" "%APP_DIR%\app.py"
"%NSSM%" set %SERVICE_NAME% AppDirectory        "%APP_DIR%"
"%NSSM%" set %SERVICE_NAME% AppStdout           "%LOG_DIR%\server.log"
"%NSSM%" set %SERVICE_NAME% AppStderr           "%LOG_DIR%\error.log"
"%NSSM%" set %SERVICE_NAME% AppRotateFiles      1
"%NSSM%" set %SERVICE_NAME% AppRotateBytes      1048576
"%NSSM%" set %SERVICE_NAME% AppRestartDelay     3000
"%NSSM%" set %SERVICE_NAME% Description         "ARCA Facturacion Electronica"
"%NSSM%" set %SERVICE_NAME% DisplayName         "ARCA Facturacion"
"%NSSM%" set %SERVICE_NAME% Start               SERVICE_AUTO_START

:: Iniciar el servicio
"%NSSM%" start %SERVICE_NAME% >nul 2>&1
echo  OK - Servicio instalado y en ejecucion.
echo.

:: ---- 5. Crear acceso directo en escritorio publico ---------------
echo [5/5] Creando acceso directo para todos los usuarios...

powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$ws = New-Object -ComObject WScript.Shell;" ^
  "$sc = $ws.CreateShortcut('%PUBLIC_DESKTOP%\ARCA Facturacion.lnk');" ^
  "$sc.TargetPath = '%APP_DIR%\iniciar_servidor.bat';" ^
  "$sc.WorkingDirectory = '%APP_DIR%';" ^
  "$sc.WindowStyle = 7;" ^
  "$sc.Description = 'ARCA - Facturacion Electronica';" ^
  "$sc.Save()"

if %errorlevel% == 0 (
    echo  OK - Acceso directo creado en el escritorio de todos los usuarios.
) else (
    echo  ADVERTENCIA: No se pudo crear el acceso directo en el escritorio publico.
)
echo.

:: ---- Resumen final -----------------------------------------------
echo  =====================================================
echo   Instalacion completada exitosamente.
echo.
echo   El servidor ARCA corre como servicio de Windows:
echo   - Arranca automaticamente con Windows
echo   - Funciona para todos los usuarios
echo   - Se reinicia solo si falla
echo   - Logs en: %LOG_DIR%\
echo.
echo   Uso: doble clic en "ARCA Facturacion" del escritorio
echo   o abrir Chrome en http://localhost:5000
echo  =====================================================
echo.
pause
endlocal
