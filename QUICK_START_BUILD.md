# Quick Start: Build & Release

## TL;DR - Do This:

### 1. Build Locally (Test First)
```powershell
.\build.ps1
```

Or if you prefer batch:
```cmd
build.bat
```

### 2. When Ready to Release
```bash
git tag v2.1.0
git push origin v2.1.0
```

### 3. GitHub Does the Rest ✨
- GitHub Actions automatically builds your app
- Creates a draft release
- Uploads exe and installer as release assets
- You just need to publish the draft!

---

## The Three Build Methods

### 📍 Method 1: GitHub Actions (Recommended)
**When**: You commit code and create a tag
**What happens**: GitHub builds everything automatically

```bash
git tag v2.1.0
git push origin v2.1.0
# GitHub Actions builds and creates draft release
```

**Pros**: 
- ✅ Clean builds every time
- ✅ No dependencies on your computer
- ✅ Consistent builds
- ✅ Multiple contributors can release

**Cons**: 
- ❌ Takes ~5 minutes per build
- ❌ Hard to debug locally

---

### 📍 Method 2: Local Build (Testing)
**When**: You want to test before pushing
**What happens**: Builds on your Windows machine

```powershell
# PowerShell (Recommended)
.\build.ps1

# Or: Batch file
build.bat
```

**Pros**: 
- ✅ Fast iteration
- ✅ Easy to debug
- ✅ No GitHub needed
- ✅ Immediate results

**Cons**: 
- ❌ Need Python + PyInstaller + InnoSetup installed
- ❌ Different computers might build differently

---

### 📍 Method 3: Manual (For Experts)
**When**: You want full control
**What happens**: You run each command yourself

```bash
# Install dependencies
pip install -r requirements.txt
pip install pyinstaller

# Build exe
pyinstaller "Kalinga OpHUB.spec"

# Build installer (Windows only)
"C:\Program Files (x86)\Inno Setup 6\ISCC.exe" "OpsHUB.iss"

# Create release and upload
gh release create v2.1.0 --draft dist/... Output/...
```

---

## File Guide

| File | Purpose |
|------|---------|
| `.github/workflows/build-and-release.yml` | GitHub Actions automation |
| `build.ps1` | PowerShell build script (Windows) |
| `build.bat` | Batch build script (Windows) |
| `build.sh` | Bash build script (Linux/Mac) |
| `Kalinga OpHUB.spec` | PyInstaller configuration |
| `OpsHUB.iss` | InnoSetup installer configuration |
| `BUILD_RELEASE_GUIDE.md` | Detailed documentation |

---

## Common Tasks

### 🔄 Release a New Version

```bash
# 1. Update version in code (if needed)
# Edit: Kalinga OpHUB.py → CURRENT_VERSION = "2.1.0"

# 2. Commit
git add .
git commit -m "Version 2.1.0 release"

# 3. Tag it
git tag v2.1.0

# 4. Push to GitHub
git push origin main
git push origin v2.1.0

# 5. Wait for GitHub Actions (5-10 minutes)
# Then go to GitHub Releases tab and PUBLISH the draft
```

### 🧪 Test Build Locally

```batch
.\build.ps1
# Check: dist\Kalinga OpHUB\Kalinga OpHUB.exe
# Check: Output\KalingaOpsHUB_Setup_v*.exe
```

### 🚀 View Build Progress

1. Go to GitHub
2. Click **Actions** tab
3. Watch "Build and Create Release" workflow
4. View logs if anything fails

### 📦 Manual Release Upload

If GitHub Actions fails, manually upload:

```bash
gh release create v2.1.0 \
  --draft \
  dist/Kalinga\ OpHUB/Kalinga\ OpHUB.exe \
  Output/KalingaOpsHUB_Setup_v*.exe
```

---

## Requirements

### For GitHub Actions (Automatic)
- ✅ Just push a tag
- ✅ No additional setup needed

### For Local Builds
- Python 3.7+
- PyInstaller: `pip install pyinstaller`
- InnoSetup 6: [Download](https://jrsoftware.org/isinfo.php)
- All requirements from `requirements.txt`

---

## Troubleshooting

| Issue | Solution |
|-------|----------|
| Build fails on GitHub Actions | Check Actions tab for logs |
| Release not creating | Make sure tag format is `v*.*.*` (e.g., v2.0.0) |
| .exe file missing | PyInstaller failed - check build logs |
| Installer missing | InnoSetup failed - need to install on Windows |
| Can't run .ps1 script | Run: `Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser` |

---

## Next Steps

1. ✅ Commit the workflow files to GitHub
2. ✅ Test with: `.\build.ps1`
3. ✅ Create a test tag: `git tag v2.0.0b1`
4. ✅ Push: `git push origin v2.0.0b1`
5. ✅ Watch GitHub Actions build
6. ✅ Go to Releases and publish the draft!

---

**Need help?** Check `BUILD_RELEASE_GUIDE.md` for detailed instructions.
