#!/bin/bash
# Kalinga OpsHUB Build Script (Bash version for Linux/Mac)

echo "======================================"
echo "Kalinga OpsHUB - Build Script"
echo "======================================"

# Check if Python is installed
if ! command -v python &> /dev/null; then
    echo "[✗] ERROR: Python not found. Please install Python first."
    exit 1
fi

python_version=$(python --version 2>&1)
echo "[✓] Python found: $python_version"

echo ""
echo "[1/4] Installing dependencies..."
pip install -r requirements.txt
pip install pyinstaller

echo ""
echo "[2/4] Cleaning old builds..."
rm -rf dist
rm -rf build

echo ""
echo "[3/4] Building with PyInstaller..."
pyinstaller "Kalinga OpHUB.spec"

if [ $? -ne 0 ]; then
    echo "[✗] ERROR: PyInstaller build failed!"
    exit 1
fi

echo "[✓] PyInstaller build completed"

echo ""
echo "[4/4] Note: InnoSetup installer is Windows-only"
echo "      .EXE is ready in: dist/Kalinga OpHUB/"

echo ""
echo "======================================"
echo "BUILD COMPLETE!"
echo "======================================"
echo ""
echo "Output files:"
echo "  - Executable: dist/Kalinga OpHUB/Kalinga OpHUB.exe"
echo ""
echo "Next steps:"
echo "  1. Transfer .exe to Windows machine"
echo "  2. Build installer with: ./build.ps1 (on Windows)"
echo "  3. Create a git tag: git tag v2.X.X"
echo "  4. Push to GitHub: git push origin v2.X.X"
echo "  5. GitHub Actions will create a draft release"
echo ""
