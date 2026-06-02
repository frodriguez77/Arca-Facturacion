@echo off
setlocal EnableDelayedExpansion
chcp 65001 >nul

set APP_DIR=%~dp0
set APP_DIR=%APP_DIR:~0,-1%
set VBS=%APP_DIR%\iniciar_arca.vbs
set PUBLIC_DESKTOP=C:\Users\Public\Desktop

net session >nul 2>&1
if %errorlevel% neq 0 (
    echo.
    echo  ERROR: Requiere permisos de Administrador.
    echo  Click derecho ^> Ejecutar como administrador
    echo.
    pause
    exit /b 1
)

powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$ws = New-Object -ComObject WScript.Shell;" ^
  "$sc = $ws.CreateShortcut('%PUBLIC_DESKTOP%\ARCA Facturacion.lnk');" ^
  "$sc.TargetPath = '%VBS%';" ^
  "$sc.WorkingDirectory = '%APP_DIR%';" ^
  "$sc.WindowStyle = 1;" ^
  "$sc.Description = 'ARCA - Facturacion Electronica';" ^
  "$sc.Save()"

if %errorlevel% == 0 (
    echo.
    echo  OK - Acceso directo creado para todos los usuarios.
    echo  Los usuarios pueden abrir ARCA desde su escritorio.
) else (
    echo  ERROR al crear el acceso directo.
)
echo.
pause
endlocal
