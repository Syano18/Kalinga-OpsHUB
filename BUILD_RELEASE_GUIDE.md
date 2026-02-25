# How to Build & Auto Release on GitHub

## Setup Instructions

### Step 1: Push Your Workflow to GitHub

The workflow file is now at `.github/workflows/build-and-release.yml`

Commit and push it to your repository:

```bash
git add .github/workflows/build-and-release.yml
git commit -m "Add GitHub Actions CI/CD workflow"
git push origin main
```

---

## How to Trigger an Automatic Build & Release

### Option A: Via Git Commands (Local)

1. **Update your version** in `Kalinga OpHUB.py` (update `CURRENT_VERSION`):
   ```python
   CURRENT_VERSION = "2.1"
   ```

2. **Create a version tag** and push it:
   ```bash
   git tag v2.1.0
   git push origin v2.1.0
   ```

3. **GitHub Actions will automatically**:
   - ✅ Build the .exe with PyInstaller
   - ✅ Create InnoSetup installer
   - ✅ Create a **DRAFT release** on GitHub
   - ✅ Upload both files as release assets

4. **Review the release** on GitHub:
   - Go to your repo → **Releases** tab
   - You'll see a **Draft** release with your version
   - Edit the release notes if needed
   - Click **"Publish release"** to make it public

---

### Option B: Via GitHub Web Interface

1. **Go to GitHub** → Your repository
2. **Click "Releases"** on the right sidebar
3. Click **"Create release"** or **"Draft a new release"**
4. **Fill in**:
   - Tag: `v2.1.0`
   - Release title: `Release v2.1.0`
   - Description: Your release notes
5. **Check "Set as a draft"**
6. Click **"Publish release"**

The workflow will automatically trigger when you create the tag!

---

## What Gets Built

When you push a tag (e.g., `v2.0.0`), the workflow:

| File | Build Method | Location |
|------|--------------|----------|
| **Kalinga-OpHUB.exe** | PyInstaller | `dist/Kalinga OpHUB/Kalinga OpHUB.exe` |
| **KalingaOpsHUB-Installer.exe** | InnoSetup | `Output/KalingaOpsHUB_Setup_v*.exe` |

Both get uploaded to the GitHub release automatically! ✨

---

## Troubleshooting

### Release not triggering?

1. **Check the tag format**:
   ```bash
   git tag v2.0.0  # Must be v + semantic versioning
   git push origin v2.0.0
   ```

2. **View workflow status**:
   - Go to your repo → **Actions** tab
   - You should see the "Build and Create Release" workflow running

3. **If it fails**:
   - Click on the failed workflow
   - View the build logs to see the error
   - Common issues:
     - Missing `requirements.txt` dependencies
     - Wrong file paths in the .iss file
     - Python 3.13 compatibility issues

### Modify the workflow

Edit `.github/workflows/build-and-release.yml` if you need to:
- Change the Python version
- Add/remove build steps
- Modify output file names
- Change release naming

---

## Next Steps

1. ✅ Commit the workflow file
2. ✅ Create a test tag: `git tag v2.0.0 && git push origin v2.0.0`
3. ✅ Watch GitHub Actions build your app
4. ✅ Review the draft release
5. ✅ Publish when ready!

---

## Example Workflow

```bash
# 1. Update version in code
# CURRENT_VERSION = "2.1.0"

# 2. Commit changes
git add .
git commit -m "Version 2.1.0 - Add past record functionality"

# 3. Create tag
git tag v2.1.0

# 4. Push commits and tag
git push origin main
git push origin v2.1.0

# 5. Go to GitHub → Releases tab
# The build will run automatically and create a draft release
# Review and publish when ready!
```

---

## Environment Variables Needed

The workflow uses:
- `GITHUB_TOKEN` - Automatically provided by GitHub Actions
- Your `requirements.txt` - Must be up to date

No additional secrets needed! ✅
