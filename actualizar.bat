@echo off
setlocal EnableDelayedExpansion
chcp 65001 >nul 2>&1

set BASE=D:\Arca-Facturacion
set REPO=https://github.com/frodriguez77/Arca-Facturacion.git
set BRANCH=claude/new-pc-download-setup-5AF1K

echo.
echo  ============================================
echo    ARCA Facturacion - Actualizador
echo  ============================================
echo.

REM --- Detener servidor si está corriendo ---
echo  Deteniendo el servidor...
taskkill /f /im arca.exe >nul 2>&1
taskkill /f /im pythonw.exe >nul 2>&1
timeout /t 1 /nobreak >nul

REM --- Verificar que git esté disponible ---
git --version >nul 2>&1
if errorlevel 1 (
    echo  ERROR: No se encontró Git en el sistema.
    echo  Instalá Git desde https://git-scm.com/download/win
    echo.
    pause
    exit /b 1
)

REM --- Inicializar git en la carpeta si no está inicializado ---
cd /d "%BASE%"
if not exist ".git" (
    echo  Inicializando repositorio git...
    git init -q
    git remote add origin %REPO%
)

REM --- Verificar que el remote apunte al repo correcto ---
git remote set-url origin %REPO% >nul 2>&1

REM --- Descargar los cambios del servidor ---
echo  Descargando actualizaciones desde GitHub...
echo.
git fetch origin %BRANCH% --quiet
if errorlevel 1 (
    echo  ERROR: No se pudo conectar con GitHub.
    echo  Verificá la conexión a internet.
    echo.
    pause
    exit /b 1
)

REM --- Actualizar solo los archivos del sistema (sin tocar datos) ---
echo  Aplicando archivos actualizados...
git checkout origin/%BRANCH% -- app.py
if errorlevel 1 goto error_checkout
echo  [OK] app.py

git checkout origin/%BRANCH% -- factura_pdf.py
if errorlevel 1 goto error_checkout
echo  [OK] factura_pdf.py

git checkout origin/%BRANCH% -- wsfe.py
if errorlevel 1 goto error_checkout
echo  [OK] wsfe.py

git checkout origin/%BRANCH% -- wsaa.py
if errorlevel 1 goto error_checkout
echo  [OK] wsaa.py

git checkout origin/%BRANCH% -- templates/admin.html
if errorlevel 1 goto error_checkout
echo  [OK] templates/admin.html

git checkout origin/%BRANCH% -- templates/index.html
if errorlevel 1 goto error_checkout
echo  [OK] templates/index.html

git checkout origin/%BRANCH% -- templates/reportes.html
if errorlevel 1 goto error_checkout
echo  [OK] templates/reportes.html

echo.
echo  ============================================
echo    Actualizacion completada correctamente!
echo  ============================================
echo.

REM --- Reiniciar el servidor ---
set /p REINICIAR= Reiniciar el sistema ahora? (S/N):
if /i "%REINICIAR%"=="S" (
    echo  Iniciando ARCA Facturacion...
    start "" "%BASE%\iniciar_arca.vbs"
)

goto fin

:error_checkout
echo.
echo  ERROR: No se pudieron aplicar los cambios.
echo  Es posible que el repositorio sea privado y necesites
echo  ingresar tu usuario y contraseña (token) de GitHub.
echo.

:fin
echo.
pause
