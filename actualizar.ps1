# ARCA Facturacion - Actualizador
$BASE    = $PSScriptRoot
$REPO    = 'https://github.com/frodriguez77/Arca-Facturacion.git'
$BRANCH  = 'claude/new-pc-download-setup-5AF1K'

Write-Host ''
Write-Host '  ==========================================' -ForegroundColor Cyan
Write-Host '    ARCA Facturacion - Actualizador' -ForegroundColor Cyan
Write-Host '  ==========================================' -ForegroundColor Cyan
Write-Host ''

# Detener servidor
Write-Host '  Deteniendo el servidor...'
Stop-Process -Name 'arca' -Force -ErrorAction SilentlyContinue
Stop-Process -Name 'py'   -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 1

# Verificar git
if (-not (Get-Command git -ErrorAction SilentlyContinue)) {
    Write-Host '  ERROR: Git no esta instalado.' -ForegroundColor Red
    Read-Host 'Presiona Enter para salir'
    exit 1
}

Set-Location -Path $BASE

# Inicializar repo si no existe
if (-not (Test-Path '.git')) {
    Write-Host '  Inicializando repositorio git...'
    git init -q
    git remote add origin $REPO
}

git remote set-url origin $REPO 2>$null

# Descargar cambios
Write-Host '  Descargando actualizaciones desde GitHub...'
Write-Host ''
git fetch origin $BRANCH --quiet
if ($LASTEXITCODE -ne 0) {
    Write-Host '  ERROR: No se pudo conectar con GitHub.' -ForegroundColor Red
    Read-Host 'Presiona Enter para salir'
    exit 1
}

# Aplicar archivos
Write-Host '  Aplicando archivos actualizados...'
$archivos = @(
    'VERSION',
    'app.py',
    'factura_pdf.py',
    'wsfe.py',
    'wsaa.py',
    'config.py',
    'repository.py',
    'openssl_util.py',
    'wspadron.py',
    'iniciar_arca.vbs',
    'crear_acceso_directo.bat',
    'actualizar.bat',
    'templates/admin.html',
    'templates/index.html',
    'templates/reportes.html',
    'templates/login.html',
    'templates/error.html'
)
foreach ($f in $archivos) {
    git checkout origin/$BRANCH -- $f
    Write-Host "  [OK] $f" -ForegroundColor Green
}

$version = if (Test-Path 'VERSION') { Get-Content 'VERSION' -Raw | ForEach-Object { $_.Trim() } } else { '?' }

Write-Host ''
Write-Host '  ==========================================' -ForegroundColor Green
Write-Host "    Actualizacion completada!  v$version" -ForegroundColor Green
Write-Host '  ==========================================' -ForegroundColor Green
Write-Host ''

$r = Read-Host '  Reiniciar el sistema ahora? (S/N)'
if ($r -match '^[sS]') {
    Write-Host '  Iniciando ARCA Facturacion...'
    Start-Process wscript.exe -ArgumentList "`"$BASE\iniciar_arca.vbs`""
}

Write-Host ''
Read-Host 'Presiona Enter para cerrar'
