#Requires -Version 5.1
<#
.SYNOPSIS
    Deploys the Geolocation Addin for Revit 2024.

.DESCRIPTION
    Clones/pulls the private GitHub repo, builds the solution, and deploys
    the addin DLL, manifest, and sample config to the correct locations.

.PARAMETER RepoUrl
    URL of the private GitHub repository. Defaults to the configured repo.

.PARAMETER Branch
    Branch to build from. Defaults to 'main'.
#>

param(
    [string]$RepoUrl = "https://github.com/YOUR_ORG/geolocation_addin.git",
    [string]$Branch = "main"
)

$ErrorActionPreference = "Stop"

# --- Paths ---
$BuildDir = "$env:TEMP\GeolocationAddin_Build"
$AddinDir = "$env:APPDATA\Autodesk\Revit\Addins\2024"
$ConfigDir = "C:\ProgramData\GeolocationAddin"

Write-Host "=== Geolocation Addin Deployment ===" -ForegroundColor Cyan
Write-Host ""

# --- Step 1: Clone or pull ---
if (Test-Path "$BuildDir\.git") {
    Write-Host "[1/5] Updating existing repo..." -ForegroundColor Yellow
    Push-Location $BuildDir
    git checkout $Branch 2>$null
    git pull origin $Branch
    Pop-Location
} else {
    Write-Host "[1/5] Cloning repository..." -ForegroundColor Yellow
    if (Test-Path $BuildDir) { Remove-Item $BuildDir -Recurse -Force }
    git clone --branch $Branch $RepoUrl $BuildDir
}

# --- Step 2: Restore NuGet packages ---
Write-Host "[2/5] Restoring NuGet packages..." -ForegroundColor Yellow
Push-Location $BuildDir
dotnet restore GeolocationAddin.sln
Pop-Location

# --- Step 3: Build ---
Write-Host "[3/5] Building solution (Release)..." -ForegroundColor Yellow
Push-Location $BuildDir
dotnet build GeolocationAddin.sln -c Release --no-restore
Pop-Location

$OutputDir = "$BuildDir\src\GeolocationAddin\bin\Release"

# --- Step 4: Deploy to Revit addins folder ---
Write-Host "[4/5] Deploying to Revit 2024..." -ForegroundColor Yellow

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

# --- Step 5: Create config directory with samples ---
Write-Host "[5/5] Setting up config directory..." -ForegroundColor Yellow

if (-not (Test-Path $ConfigDir)) {
    New-Item -ItemType Directory -Path $ConfigDir -Force | Out-Null
}

$configFile = "$ConfigDir\config.json"
$mappingFile = "$ConfigDir\mapping.csv"

if (-not (Test-Path $configFile)) {
    Copy-Item "$BuildDir\config\config.sample.json" -Destination $configFile
    Write-Host "  Created sample config.json" -ForegroundColor Green
} else {
    Write-Host "  config.json already exists, skipping" -ForegroundColor DarkGray
}

if (-not (Test-Path $mappingFile)) {
    Copy-Item "$BuildDir\config\mapping.sample.csv" -Destination $mappingFile
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
