@echo off
REM Kalinga OpsHUB Build Script
REM This script builds the exe and installer locally

echo ======================================
echo Kalinga OpsHUB - Build Script
echo ======================================

REM Check if Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python not found. Please install Python first.
    pause
    exit /b 1
)

echo.
echo [1/4] Installing dependencies...
pip install -r requirements.txt
pip install pyinstaller

echo.
echo [2/4] Cleaning old builds...
if exist "dist" rmdir /s /q dist
if exist "build" rmdir /s /q build

echo.
echo [3/4] Building with PyInstaller...
pyinstaller "Kalinga OpHUB.spec"

if errorlevel 1 (
    echo ERROR: PyInstaller build failed!
    pause
    exit /b 1
)

echo.
echo [4/4] Building installer with InnoSetup...

REM Check if InnoSetup is installed
if not exist "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" (
    echo WARNING: InnoSetup not found at C:\Program Files (x86)\Inno Setup 6\
    echo Download from: https://jrsoftware.org/isinfo.php
    echo.
    echo Skipping installer build. .exe is ready in: dist\Kalinga OpHUB\
    pause
    exit /b 0
)

"C:\Program Files (x86)\Inno Setup 6\ISCC.exe" "OpsHUB.iss"

if errorlevel 1 (
    echo ERROR: InnoSetup build failed!
    pause
    exit /b 1
)

echo.
echo ======================================
echo BUILD COMPLETE!
echo ======================================
echo.
echo Output files:
echo   - Executable: dist\Kalinga OpHUB\Kalinga OpHUB.exe
echo   - Installer: Output\KalingaOpsHUB_Setup_v*.exe
echo.
echo Next steps:
echo   1. Test the executable
echo   2. Create a git tag: git tag v2.X.X
echo   3. Push to GitHub: git push origin v2.X.X
echo   4. GitHub Actions will create a draft release
echo.
pause
