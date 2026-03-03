#Requires -Version 5.1
<#
.SYNOPSIS
    Builds the Geolocation Addin and packages it with Inno Setup.

.DESCRIPTION
    1. Builds the solution in Release configuration
    2. Compiles the Inno Setup script into a standalone installer .exe

    Prerequisites:
      - .NET SDK (for dotnet build)
      - Inno Setup 6 (iscc.exe must be on PATH or installed in the default location)

    Output: installer\Output\GeolocationAddin-Setup-<version>.exe

.EXAMPLE
    cd geolocation_addin\installer
    .\Build-Installer.ps1
#>

$ErrorActionPreference = "Stop"

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$RepoRoot = Split-Path -Parent $ScriptDir

if (-not (Test-Path "$RepoRoot\GeolocationAddin.sln")) {
    Write-Host "ERROR: Cannot find GeolocationAddin.sln in $RepoRoot" -ForegroundColor Red
    exit 1
}

# --- Step 1: Build ---
Write-Host "=== Building Solution (Release) ===" -ForegroundColor Cyan
Push-Location $RepoRoot
dotnet restore GeolocationAddin.sln
dotnet build GeolocationAddin.sln -c Release --no-restore
Pop-Location

$OutputDir = "$RepoRoot\src\GeolocationAddin\bin\Release"

# Verify build output exists
$requiredFiles = @("GeolocationAddin.dll", "GeolocationAddin.addin", "Newtonsoft.Json.dll")
foreach ($file in $requiredFiles) {
    if (-not (Test-Path "$OutputDir\$file")) {
        Write-Host "ERROR: Build output missing: $OutputDir\$file" -ForegroundColor Red
        exit 1
    }
}
Write-Host "Build output verified." -ForegroundColor Green

# --- Step 2: Find Inno Setup compiler ---
$iscc = $null

# Check PATH first
$iscc = Get-Command "iscc" -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Source

# Fall back to default install locations
if (-not $iscc) {
    $candidates = @(
        "${env:ProgramFiles(x86)}\Inno Setup 6\ISCC.exe",
        "$env:ProgramFiles\Inno Setup 6\ISCC.exe"
    )
    foreach ($path in $candidates) {
        if (Test-Path $path) {
            $iscc = $path
            break
        }
    }
}

if (-not $iscc) {
    Write-Host "ERROR: Inno Setup compiler (ISCC.exe) not found." -ForegroundColor Red
    Write-Host "Install Inno Setup 6 from: https://jrsoftware.org/isdl.php" -ForegroundColor Yellow
    exit 1
}

Write-Host "Using Inno Setup: $iscc" -ForegroundColor Gray

# --- Step 3: Compile installer ---
Write-Host ""
Write-Host "=== Compiling Installer ===" -ForegroundColor Cyan

$issFile = "$ScriptDir\GeolocationAddin.iss"
& $iscc $issFile

if ($LASTEXITCODE -ne 0) {
    Write-Host "ERROR: Inno Setup compilation failed." -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "=== Installer Created ===" -ForegroundColor Cyan
$outputExe = Get-ChildItem "$ScriptDir\Output\*.exe" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
if ($outputExe) {
    Write-Host "  $($outputExe.FullName)" -ForegroundColor Green
    Write-Host "  Size: $([math]::Round($outputExe.Length / 1KB)) KB" -ForegroundColor Gray
}
