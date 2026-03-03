#Requires -Version 5.1
<#
.SYNOPSIS
    Deploys the Geolocation Addin for Revit 2024.

.DESCRIPTION
    Builds the solution from the local repo and deploys the addin DLL,
    manifest, and sample config to the correct locations.

    Run from within a cloned copy of the repo:
      git clone https://github.com/lwilks/geolocation_addin.git
      cd geolocation_addin\deploy
      .\Deploy-GeolocationAddin.ps1
#>

$ErrorActionPreference = "Stop"

# --- Resolve repo root from script location ---
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$RepoRoot = Split-Path -Parent $ScriptDir

if (-not (Test-Path "$RepoRoot\GeolocationAddin.sln")) {
    Write-Host "ERROR: Cannot find GeolocationAddin.sln in $RepoRoot" -ForegroundColor Red
    Write-Host "Make sure you are running this script from within the cloned repo." -ForegroundColor Red
    exit 1
}

# --- Paths ---
$AddinDir = "$env:APPDATA\Autodesk\Revit\Addins\2024"
$ConfigDir = "C:\ProgramData\GeolocationAddin"

Write-Host "=== Geolocation Addin Deployment ===" -ForegroundColor Cyan
Write-Host "  Repo: $RepoRoot" -ForegroundColor Gray
Write-Host ""

# --- Step 1: Restore NuGet packages ---
Write-Host "[1/4] Restoring NuGet packages..." -ForegroundColor Yellow
Push-Location $RepoRoot
dotnet restore GeolocationAddin.sln
Pop-Location

# --- Step 2: Build ---
Write-Host "[2/4] Building solution (Release)..." -ForegroundColor Yellow
Push-Location $RepoRoot
dotnet build GeolocationAddin.sln -c Release --no-restore
Pop-Location

$OutputDir = "$RepoRoot\src\GeolocationAddin\bin\Release"

# --- Step 3: Deploy to Revit addins folder ---
Write-Host "[3/4] Deploying to Revit 2024..." -ForegroundColor Yellow

if (-not (Test-Path $AddinDir)) {
    New-Item -ItemType Directory -Path $AddinDir -Force | Out-Null
}

$filesToCopy = @(
    "$OutputDir\GeolocationAddin.dll",
    "$OutputDir\GeolocationAddin.addin",
    "$OutputDir\Newtonsoft.Json.dll"
)

foreach ($file in $filesToCopy) {
    if (Test-Path $file) {
        Copy-Item $file -Destination $AddinDir -Force
        Write-Host "  Copied: $(Split-Path $file -Leaf)" -ForegroundColor Green
    } else {
        Write-Host "  WARNING: Not found: $file" -ForegroundColor Red
    }
}

# --- Step 4: Create config directory with samples ---
Write-Host "[4/4] Setting up config directory..." -ForegroundColor Yellow

if (-not (Test-Path $ConfigDir)) {
    New-Item -ItemType Directory -Path $ConfigDir -Force | Out-Null
}

$configFile = "$ConfigDir\config.json"
$mappingFile = "$ConfigDir\mapping.csv"

if (-not (Test-Path $configFile)) {
    Copy-Item "$RepoRoot\config\config.sample.json" -Destination $configFile
    Write-Host "  Created sample config.json" -ForegroundColor Green
} else {
    Write-Host "  config.json already exists, skipping" -ForegroundColor DarkGray
}

if (-not (Test-Path $mappingFile)) {
    Copy-Item "$RepoRoot\config\mapping.sample.csv" -Destination $mappingFile
    Write-Host "  Created sample mapping.csv" -ForegroundColor Green
} else {
    Write-Host "  mapping.csv already exists, skipping" -ForegroundColor DarkGray
}

# --- Done ---
Write-Host ""
Write-Host "=== Deployment Complete ===" -ForegroundColor Cyan
Write-Host ""
Write-Host "Deployed files:" -ForegroundColor White
Write-Host "  DLL + manifest: $AddinDir" -ForegroundColor Gray
Write-Host "  Config:         $ConfigDir" -ForegroundColor Gray
Write-Host ""
Write-Host "NEXT STEPS:" -ForegroundColor Yellow
Write-Host "  1. Edit $configFile" -ForegroundColor White
Write-Host "     - Set siteModelPath to your Desktop Connector site model path" -ForegroundColor Gray
Write-Host "     - Set output folder paths" -ForegroundColor Gray
Write-Host "  2. Edit $mappingFile" -ForegroundColor White
Write-Host "     - Add rows mapping link instance names to target file names" -ForegroundColor Gray
Write-Host "  3. Open Revit 2024 - look for the 'Geolocation' tab in the ribbon" -ForegroundColor White
Write-Host ""
