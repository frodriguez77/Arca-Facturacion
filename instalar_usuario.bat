@echo off
setlocal EnableDelayedExpansion
chcp 65001 >nul

:: ============================================================
::  ARCA - Facturación Electrónica
::  Instalación por usuario — NO requiere Administrador
::  El servidor arranca cuando el usuario abre el acceso directo
:: ============================================================

set APP_DIR=%~dp0
set APP_DIR=%APP_DIR:~0,-1%
set DESKTOP=%USERPROFILE%\Desktop

echo.
echo  =====================================================
echo   ARCA - Instalacion para este usuario
echo  =====================================================
echo.

:: ---- 1. Verificar Python ----------------------------------------
echo [1/3] Verificando Python...
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

:: ---- 2. Instalar dependencias -----------------------------------
echo [2/3] Instalando dependencias Python...
python -m pip install --upgrade pip --quiet
python -m pip install flask pandas openpyxl zeep requests reportlab qrcode[pil] urllib3 werkzeug --quiet
if %errorlevel% neq 0 (
    echo  ERROR al instalar dependencias.
    pause
    exit /b 1
)
echo  OK - Dependencias instaladas.
echo.

:: ---- 3. Crear carpetas y acceso directo -------------------------
echo [3/3] Configurando acceso directo...
if not exist "%APP_DIR%\uploads"       mkdir "%APP_DIR%\uploads"
if not exist "%APP_DIR%\certificados"  mkdir "%APP_DIR%\certificados"
if not exist "%APP_DIR%\logs"          mkdir "%APP_DIR%\logs"

powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$ws = New-Object -ComObject WScript.Shell;" ^
  "$sc = $ws.CreateShortcut('%DESKTOP%\ARCA Facturacion.lnk');" ^
  "$sc.TargetPath = '%APP_DIR%\iniciar_usuario.bat';" ^
  "$sc.WorkingDirectory = '%APP_DIR%';" ^
  "$sc.WindowStyle = 7;" ^
  "$sc.Description = 'ARCA - Facturacion Electronica';" ^
  "$sc.Save()"

echo  OK - Acceso directo creado en tu escritorio.
echo.

:: ---- Resumen ---------------------------------------------------
echo  =====================================================
echo   Instalacion completada.
echo.
echo   Como usar:
echo   - Doble clic en "ARCA Facturacion" del escritorio
echo   - El servidor inicia solo y se abre Chrome
echo   - Si ya esta corriendo, abre Chrome directo
echo.
echo   Acceso manual: http://localhost:5000
echo  =====================================================
echo.

set /p INICIAR="Deseas iniciar el sistema ahora? (S/N): "
if /i "%INICIAR%"=="S" (
    start "" "%APP_DIR%\iniciar_usuario.bat"
)

endlocal
