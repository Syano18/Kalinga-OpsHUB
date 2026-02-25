#!/usr/bin/env pwsh
# Kalinga OpsHUB Build Script (PowerShell version)

Write-Host "======================================" -ForegroundColor Cyan
Write-Host "Kalinga OpsHUB - Build Script" -ForegroundColor Cyan
Write-Host "======================================" -ForegroundColor Cyan

# Check if Python is installed
try {
    $pythonVersion = python --version 2>&1
    Write-Host "[✓] Python found: $pythonVersion" -ForegroundColor Green
} catch {
    Write-Host "[✗] ERROR: Python not found. Please install Python first." -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit 1
}

Write-Host ""
Write-Host "[1/4] Installing dependencies..." -ForegroundColor Yellow
pip install -r requirements.txt
pip install pyinstaller

Write-Host ""
Write-Host "[2/4] Cleaning old builds..." -ForegroundColor Yellow
if (Test-Path "dist") {
    Remove-Item "dist" -Recurse -Force
}
if (Test-Path "build") {
    Remove-Item "build" -Recurse -Force
}

Write-Host ""
Write-Host "[3/4] Building with PyInstaller..." -ForegroundColor Yellow
pyinstaller "Kalinga OpHUB.spec"

if ($LASTEXITCODE -ne 0) {
    Write-Host "[✗] ERROR: PyInstaller build failed!" -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit 1
}

Write-Host "[✓] PyInstaller build completed" -ForegroundColor Green

Write-Host ""
Write-Host "[4/4] Building installer with InnoSetup..." -ForegroundColor Yellow

$isccPath = "C:\Program Files (x86)\Inno Setup 6\ISCC.exe"
if (-Not (Test-Path $isccPath)) {
    Write-Host "[!] WARNING: InnoSetup not found at: $isccPath" -ForegroundColor Yellow
    Write-Host "    Download from: https://jrsoftware.org/isinfo.php" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "[✓] .EXE is ready in: dist\Kalinga OpHUB\" -ForegroundColor Green
    Read-Host "Press Enter to exit"
    exit 0
}

& $isccPath "OpsHUB.iss"

if ($LASTEXITCODE -ne 0) {
    Write-Host "[✗] ERROR: InnoSetup build failed!" -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit 1
}

Write-Host ""
Write-Host "======================================" -ForegroundColor Green
Write-Host "BUILD COMPLETE!" -ForegroundColor Green
Write-Host "======================================" -ForegroundColor Green
Write-Host ""
Write-Host "Output files:" -ForegroundColor Cyan
Write-Host "  - Executable: dist\Kalinga OpHUB\Kalinga OpHUB.exe" -ForegroundColor White
Write-Host "  - Installer: Output\KalingaOpsHUB_Setup_v*.exe" -ForegroundColor White
Write-Host ""
Write-Host "Next steps:" -ForegroundColor Cyan
Write-Host "  1. Test the executable" -ForegroundColor White
Write-Host "  2. Create a git tag: git tag v2.X.X" -ForegroundColor White
Write-Host "  3. Push to GitHub: git push origin v2.X.X" -ForegroundColor White
Write-Host "  4. GitHub Actions will create a draft release" -ForegroundColor White
Write-Host ""
Read-Host "Press Enter to exit"
