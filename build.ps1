#Requires -Version 5.1
<#
.SYNOPSIS
    Baut das OneNoteExporter-Projekt und kopiert das fertige Release in den Ordner "build".
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$ProjectRoot = $PSScriptRoot
$BuildDir    = Join-Path $ProjectRoot 'build'
$PublishDir  = Join-Path $ProjectRoot 'bin\Release\net10.0-windows\win-x64\publish'

# ── 1. build-Ordner leeren / anlegen ─────────────────────────────────────────
Write-Host ">> Bereite build-Ordner vor..." -ForegroundColor Cyan
if (Test-Path $BuildDir) {
    Remove-Item -Path $BuildDir -Recurse -Force
}
New-Item -ItemType Directory -Path $BuildDir | Out-Null

# ── 2. dotnet publish (Release, self-contained, x64) ─────────────────────────
Write-Host ">> Starte dotnet publish..." -ForegroundColor Cyan
dotnet publish "$ProjectRoot\OneNoteExporter.csproj" `
    -c Release `
    -r win-x64 `
    --self-contained true `
    -p:PublishSingleFile=false

if ($LASTEXITCODE -ne 0) {
    Write-Host "FEHLER: dotnet publish fehlgeschlagen (Exit-Code $LASTEXITCODE)." -ForegroundColor Red
    exit $LASTEXITCODE
}

# ── 3. Dateien in build-Ordner kopieren ──────────────────────────────────────
Write-Host ">> Kopiere Dateien nach: $BuildDir" -ForegroundColor Cyan
Copy-Item -Path "$PublishDir\*" -Destination $BuildDir -Recurse -Force

# ── 4. Fertig ─────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "Fertig! Build liegt in: $BuildDir" -ForegroundColor Green
$exe = Join-Path $BuildDir 'OneNoteExporter.exe'
if (Test-Path $exe) {
    $size = [math]::Round((Get-Item $exe).Length / 1MB, 1)
    Write-Host "  OneNoteExporter.exe  ($size MB)" -ForegroundColor Green
}
