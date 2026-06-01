@echo off
setlocal
chcp 65001 >nul

:: ============================================================
::  ARCA - Facturación Electrónica
::  Script de desinstalación
::  Requiere ejecutar como ADMINISTRADOR
:: ============================================================

net session >nul 2>&1
if %errorlevel% neq 0 (
    echo ERROR: Requiere permisos de Administrador.
    pause
    exit /b 1
)

set APP_DIR=%~dp0
set APP_DIR=%APP_DIR:~0,-1%
set SERVICE_NAME=ArcaFacturacion
set NSSM=%APP_DIR%\tools\nssm.exe
set PUBLIC_DESKTOP=C:\Users\Public\Desktop

echo.
echo  =====================================================
echo   ARCA - Desinstalacion
echo  =====================================================
echo.
echo  ADVERTENCIA: Esto detiene y elimina el servicio ARCA.
echo  Los datos (facturas, empresas, usuarios) NO se borran.
echo.
set /p CONFIRMAR="Confirmar desinstalacion? (S/N): "
if /i not "%CONFIRMAR%"=="S" (
    echo Cancelado.
    pause
    exit /b 0
)
echo.

:: Detener y eliminar servicio
if exist "%NSSM%" (
    echo Deteniendo servicio...
    "%NSSM%" stop %SERVICE_NAME% confirm >nul 2>&1
    "%NSSM%" remove %SERVICE_NAME% confirm >nul 2>&1
    echo OK - Servicio eliminado.
) else (
    sc stop %SERVICE_NAME% >nul 2>&1
    sc delete %SERVICE_NAME% >nul 2>&1
    echo OK - Servicio eliminado.
)

:: Eliminar acceso directo del escritorio público
if exist "%PUBLIC_DESKTOP%\ARCA Facturacion.lnk" (
    del "%PUBLIC_DESKTOP%\ARCA Facturacion.lnk"
    echo OK - Acceso directo eliminado.
)

echo.
echo  Desinstalacion completada.
echo  Los datos en %APP_DIR% no fueron eliminados.
echo.
pause
endlocal
