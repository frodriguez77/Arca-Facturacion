@echo off
setlocal
chcp 65001 >nul

:: ============================================================
::  ARCA - Facturación Electrónica
::  Script de instalación — ejecutar UNA SOLA VEZ por PC
:: ============================================================

set APP_DIR=%~dp0
set APP_DIR=%APP_DIR:~0,-1%

echo.
echo  =====================================================
echo   ARCA - Instalacion de Facturacion Electronica
echo  =====================================================
echo.

:: ---- 1. Verificar Python ------------------------------------------
echo [1/4] Verificando Python...
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
for /f "tokens=*" %%v in ('python --version 2^>^&1') do set PYVER=%%v
echo  OK - %PYVER%
echo.

:: ---- 2. Instalar dependencias -------------------------------------
echo [2/4] Instalando dependencias Python...
echo.
python -m pip install --upgrade pip --quiet
python -m pip install flask pandas openpyxl zeep requests reportlab qrcode[pil] urllib3 werkzeug --quiet

if %errorlevel% neq 0 (
    echo.
    echo  ERROR al instalar dependencias.
    echo  Intenta ejecutar manualmente:
    echo  python -m pip install flask pandas openpyxl zeep requests reportlab qrcode[pil]
    echo.
    pause
    exit /b 1
)
echo  OK - Dependencias instaladas.
echo.

:: ---- 3. Crear carpetas necesarias ---------------------------------
echo [3/4] Creando estructura de carpetas...
if not exist "%APP_DIR%\uploads"       mkdir "%APP_DIR%\uploads"
if not exist "%APP_DIR%\certificados"  mkdir "%APP_DIR%\certificados"
echo  OK - Carpetas listas.
echo.

:: ---- 4. Crear acceso directo en el escritorio --------------------
echo [4/4] Creando acceso directo en el escritorio...

set SHORTCUT_NAME=ARCA Facturacion.lnk
set DESKTOP=%USERPROFILE%\Desktop

powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$ws = New-Object -ComObject WScript.Shell;" ^
  "$sc = $ws.CreateShortcut('%DESKTOP%\%SHORTCUT_NAME%');" ^
  "$sc.TargetPath = '%APP_DIR%\iniciar_servidor.bat';" ^
  "$sc.WorkingDirectory = '%APP_DIR%';" ^
  "$sc.WindowStyle = 7;" ^
  "$sc.Description = 'ARCA - Facturacion Electronica';" ^
  "$sc.Save()"

if %errorlevel% neq 0 (
    echo  ADVERTENCIA: No se pudo crear el acceso directo automaticamente.
    echo  Crea uno manualmente apuntando a: %APP_DIR%\iniciar_servidor.bat
) else (
    echo  OK - Acceso directo creado en el escritorio.
)
echo.

:: ---- Listo --------------------------------------------------------
echo  =====================================================
echo   Instalacion completada exitosamente.
echo.
echo   Para usar el sistema:
echo   - Hace doble clic en "ARCA Facturacion" del escritorio
echo   - El servidor se inicia solo y se abre Chrome
echo   - Si el servidor ya esta corriendo, abre Chrome directo
echo.
echo   Acceso manual: http://localhost:5000
echo  =====================================================
echo.

set /p INICIAR="Deseas iniciar el sistema ahora? (S/N): "
if /i "%INICIAR%"=="S" (
    start "" "%APP_DIR%\iniciar_servidor.bat"
)

endlocal
