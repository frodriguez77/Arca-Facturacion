@echo off
setlocal EnableDelayedExpansion

set BASE=D:\Arca-Facturacion
set REPO=https://github.com/frodriguez77/Arca-Facturacion.git
set BRANCH=claude/new-pc-download-setup-5AF1K

echo.
echo  ============================================
echo    ARCA Facturacion - Actualizador
echo  ============================================
echo.

echo  Deteniendo el servidor...
taskkill /f /im arca.exe >/dev/null 2>&1
taskkill /f /im py.exe >/dev/null 2>&1
timeout /t 1 /nobreak >nul

git --version >/dev/null 2>&1
if errorlevel 1 (
    echo  ERROR: No se encontro Git en el sistema.
    pause
    exit /b 1
)

cd /d "%BASE%"

if not exist ".git" (
    echo  Inicializando repositorio git...
    git init -q
    git remote add origin %REPO%
)

git remote set-url origin %REPO% >/dev/null 2>&1

echo  Descargando actualizaciones desde GitHub...
echo.
git fetch origin %BRANCH% --quiet
if errorlevel 1 (
    echo  ERROR: No se pudo conectar con GitHub.
    pause
    exit /b 1
)

echo  Aplicando archivos actualizados...
git checkout origin/%BRANCH% -- app.py
echo  [OK] app.py
git checkout origin/%BRANCH% -- factura_pdf.py
echo  [OK] factura_pdf.py
git checkout origin/%BRANCH% -- wsfe.py
echo  [OK] wsfe.py
git checkout origin/%BRANCH% -- wsaa.py
echo  [OK] wsaa.py
git checkout origin/%BRANCH% -- templates/admin.html
echo  [OK] templates/admin.html
git checkout origin/%BRANCH% -- templates/index.html
echo  [OK] templates/index.html
git checkout origin/%BRANCH% -- templates/reportes.html
echo  [OK] templates/reportes.html

echo.
echo  ============================================
echo    Actualizacion completada correctamente!
echo  ============================================
echo.

set /p REINICIAR= Reiniciar el sistema ahora? (S/N): 
if /i "%REINICIAR%"=="S" (
    echo  Iniciando ARCA Facturacion...
    start "" "%BASE%\iniciar_arca.vbs"
)

echo.
pause
