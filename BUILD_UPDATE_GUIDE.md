# LogLife Operations - Build & Update Guide

## 📦 Building the Executable

### Quick Build

1. **Run the build script:**

   ```batch
   build_main_ctk.bat
   ```

2. **Find your executable:**
   - Location: `dist\LogLife_Operacao.exe`
   - Also copied to: `release\LogLife_Operacao.exe`

### Manual Build

If you prefer to build manually:

```batch
# Install dependencies
pip install --upgrade pyinstaller pyinstaller-hooks-contrib

# Clean previous builds
rmdir /s /q dist build

# Build
pyinstaller main_ctk.spec --clean --noconfirm
```

## 🔄 Update System

### How It Works

The application includes an **automatic update checker** that:

1. **Checks on startup** - Silently checks for updates when the app starts
2. **Manual check** - Users can check anytime via the "Sobre" tab
3. **Version comparison** - Compares local version with GitHub releases
4. **Download prompt** - Opens browser to download if update is available

### Setting Up Updates

#### 1. Update version.json

Before building a new version, update `version.json`:

```json
{
  "version": "1.1.0",
  "build_date": "2025-11-08",
  "download_url": "https://github.com/ArturSenna/loglife-automations/releases/latest",
  "notes": "Nova versão com melhorias...",
  "min_required_version": "1.0.0",
  "changelog": ["Feature 1", "Bug fix 2", "Improvement 3"]
}
```

#### 2. Update Version in Code

Edit `src\utils\version_manager.py`:

```python
CURRENT_VERSION = "1.1.0"  # Change this
BUILD_DATE = "2025-11-08"   # Change this
```

#### 3. Commit and Push to GitHub

```bash
git add version.json src/utils/version_manager.py
git commit -m "Version 1.1.0 release"
git push origin main
```

#### 4. Create GitHub Release

1. Go to your repository on GitHub
2. Click "Releases" → "Create a new release"
3. Tag: `v1.1.0`
4. Title: `LogLife Operations v1.1.0`
5. Upload `LogLife_Operacao.exe` from the `release` folder
6. Add release notes
7. Publish release

### Force Update Feature

Users can force an update check by:

1. Opening the application
2. Going to the **"Sobre"** tab
3. Clicking **"Verificar Atualizações"** button

The system will:

- Check for the latest version on GitHub
- Show a dialog if an update is available
- Open the browser to download the new version
- Provide instructions for updating

## 🎯 Version Number Convention

Use [Semantic Versioning](https://semver.org/):

- **MAJOR.MINOR.PATCH** (e.g., 1.2.3)
- **MAJOR**: Breaking changes
- **MINOR**: New features (backwards compatible)
- **PATCH**: Bug fixes

Examples:

- `1.0.0` - Initial release
- `1.1.0` - Added new feature
- `1.1.1` - Fixed bug
- `2.0.0` - Major rewrite

## 📋 Release Checklist

When releasing a new version:

- [ ] Update `CURRENT_VERSION` in `src\utils\version_manager.py`
- [ ] Update `BUILD_DATE` in `src\utils\version_manager.py`
- [ ] Update `version.json` with new version and changelog
- [ ] Test the application locally
- [ ] Run `build_main_ctk.bat`
- [ ] Test the built executable
- [ ] Commit changes to Git
- [ ] Push to GitHub
- [ ] Create GitHub Release with tag
- [ ] Upload executable to GitHub Release
- [ ] Test update mechanism from old version
- [ ] Notify users (if applicable)

## 🛠️ Build Configuration

### What Gets Included

The `main_ctk.spec` file configures what's bundled:

- ✅ All Python dependencies
- ✅ CustomTkinter themes and assets
- ✅ Application icon (`src\assets\my_icon.ico`)
- ✅ CSV files (e.g., `Tabela_Malha_ID.csv`)
- ✅ Selenium WebDriver support

### What Gets Excluded

To keep the file size manageable:

- ❌ Testing modules (pytest, unittest)
- ❌ Documentation (sphinx)
- ❌ Unused scientific libraries (matplotlib, scipy)
- ❌ Development tools

### Optimizations

The build uses:

- **UPX compression** - Reduces executable size
- **Strip symbols** - Removes debugging symbols
- **One-file mode** - Single .exe file (no dependencies folder)
- **No console** - GUI-only application

## 🐛 Troubleshooting

### Build Issues

**Problem**: `ModuleNotFoundError` when running executable

**Solution**: Add missing module to `hiddenimports` in `main_ctk.spec`:

```python
hiddenimports=[
    'your_missing_module',
    # ... other imports
]
```

**Problem**: Icon not showing

**Solution**: Verify `src\assets\my_icon.ico` exists and is a valid .ico file

**Problem**: Large file size

**Solution**: Add more modules to `excluded_modules` in `main_ctk.spec`

### Update Issues

**Problem**: Update check always fails

**Solution**:

1. Check internet connection
2. Verify `UPDATE_CHECK_URL` in `version_manager.py` is correct
3. Ensure `version.json` is accessible on GitHub

**Problem**: Version comparison not working

**Solution**: Use only numeric versions (e.g., `1.2.3`, not `v1.2.3`)

## 📁 File Structure

```
loglife-automations/
├── src/
│   ├── automations/
│   │   └── main_ctk.py          # Main application
│   ├── utils/
│   │   └── version_manager.py   # Update checker
│   └── assets/
│       └── my_icon.ico           # App icon
├── main_ctk.spec                 # PyInstaller config
├── build_main_ctk.bat            # Build script
├── version.json                  # Version info (for updates)
├── dist/                         # Built executable
└── release/                      # Release copy
```

## 🚀 Distribution

### For End Users

Simply provide them with:

1. `LogLife_Operacao.exe` from the `release` folder
2. Instructions to run it

**No Python installation required!**

### System Requirements

- Windows 7/8/10/11 (64-bit)
- Minimum 4GB RAM
- 500MB free disk space

## 🔐 Security Notes

- The executable includes all dependencies (no internet required for core functions)
- Update checks require internet connection
- GitHub releases should be the only source for downloads
- Antivirus may flag the executable (false positive) - you may need to whitelist it

## 📞 Support

For issues:

1. Check this documentation
2. Review error logs in the `logs/` folder
3. Contact the development team

---

**Last Updated**: 2025-11-07  
**Document Version**: 1.0
