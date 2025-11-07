# LogLife Operations - Installer Build Guide

## Overview

This project now supports building a **Windows Installer** using Inno Setup, which provides a much better user experience than distributing a single .exe file.

## Benefits of Installer Approach

### ✅ Advantages

1. **Faster Startup** - Application starts quicker (no unpacking needed)
2. **Smaller File Size** - Individual files compress better than one large file
3. **Professional Installation** - Standard Windows install/uninstall experience
4. **Automatic Shortcuts** - Start menu and desktop icons
5. **Easy Updates** - Users can simply run new installer
6. **Better Debugging** - Can see individual DLL and dependency files
7. **Lower Antivirus False Positives** - Folder structure is less suspicious

### ⚠️ One-File vs Folder Structure

| Feature         | One-File (Old)         | Folder + Installer (New) |
| --------------- | ---------------------- | ------------------------ |
| File Count      | 1 exe file             | Multiple files in folder |
| Startup Speed   | Slower (unpacks first) | Fast (direct access)     |
| File Size       | ~70 MB                 | ~50 MB installed         |
| Distribution    | Manual copy            | Professional installer   |
| Uninstall       | Manual delete          | Windows uninstaller      |
| Updates         | Replace file           | Run new installer        |
| User Experience | Basic                  | Professional             |

## Prerequisites

### For Building

1. **Python 3.7+** with virtual environment
2. **PyInstaller** (installed via script)
3. **Inno Setup 6** - Download from https://jrsoftware.org/isdl.php

### Installation of Inno Setup

1. Download from: https://jrsoftware.org/isdl.php
2. Run the installer
3. Install to default location: `C:\Program Files (x86)\Inno Setup 6\`
4. That's it! The build script will find it automatically

## Building the Installer

### Quick Build (Recommended)

```batch
build_with_installer.bat
```

This script will:

1. ✅ Check Python environment
2. ✅ Install PyInstaller
3. ✅ Build application folder structure
4. ✅ Create Windows installer (if Inno Setup is installed)

### Manual Build

If you prefer step-by-step:

```batch
# 1. Clean previous builds
rmdir /s /q dist build installer_output

# 2. Build application folder
pyinstaller main_ctk.spec --clean --noconfirm

# 3. Create installer (requires Inno Setup)
"C:\Program Files (x86)\Inno Setup 6\ISCC.exe" installer.iss
```

## Output Files

After successful build:

### Application Folder

```
dist/
└── LogLife_Operacao/
    ├── LogLife_Operacao.exe      # Main executable
    ├── python312.dll              # Python runtime
    ├── _internal/                 # Dependencies
    │   ├── *.pyd                  # Python modules
    │   ├── *.dll                  # System libraries
    │   └── ...
    └── assets/                    # Application assets
```

### Installer Package

```
installer_output/
└── LogLife_Operations_Setup_v1.0.0.exe   # Windows installer (~50 MB)
```

## Distribution

### For End Users

**Distribute only:** `LogLife_Operations_Setup_v1.0.0.exe`

Users will:

1. Double-click the installer
2. Follow installation wizard
3. Choose installation location
4. Get shortcuts automatically
5. Use Windows to uninstall later

### For GitHub Releases

1. Build the installer
2. Go to GitHub → Releases → Create Release
3. Upload: `installer_output/LogLife_Operations_Setup_v1.0.0.exe`
4. Update `version.json` to point to the installer

## Customizing the Installer

### Edit `installer.iss`

Key sections you can customize:

```ini
; Change version
#define MyAppVersion "1.0.0"

; Change company name
#define MyAppPublisher "LogLife"

; Add/remove shortcuts
[Icons]
Name: "{autodesktop}\{#MyAppName}"; ...

; Include additional files
[Files]
Source: "your_file.txt"; DestDir: "{app}"; ...

; Add pre-installation checks
[Code]
function InitializeSetup(): Boolean;
begin
  // Your custom code here
  Result := True;
end;
```

### Installer Options

The installer includes:

- ✅ Installation location selection
- ✅ Start menu group creation
- ✅ Desktop shortcut (optional)
- ✅ Quick launch icon (optional)
- ✅ License agreement display
- ✅ README display before install
- ✅ Launch application after install
- ✅ Uninstaller creation
- ✅ Portuguese language support

## Testing

### Test the Application Folder

```batch
# Run directly from dist folder
dist\LogLife_Operacao\LogLife_Operacao.exe
```

### Test the Installer

1. Run the installer: `installer_output\LogLife_Operations_Setup_v1.0.0.exe`
2. Complete installation
3. Launch from Start menu
4. Test all features
5. Uninstall via Windows Settings → Apps

## Updating to a New Version

### Update Version Numbers

1. **In `installer.iss`:**

   ```ini
   #define MyAppVersion "1.1.0"
   ```

2. **In `src/utils/version_manager.py`:**

   ```python
   CURRENT_VERSION = "1.1.0"
   BUILD_DATE = "2025-11-08"
   ```

3. **In `version.json`:**
   ```json
   {
     "version": "1.1.0",
     "download_url": "https://github.com/..."
   }
   ```

### Build New Installer

```batch
build_with_installer.bat
```

### Distribute Update

1. Upload new installer to GitHub Releases
2. Users download and run new installer
3. Old version is automatically replaced
4. Settings and data are preserved

## Troubleshooting

### Inno Setup Not Found

**Error:** `[WARNING] Inno Setup not found`

**Solution:**

1. Install from https://jrsoftware.org/isdl.php
2. Or edit `build_with_installer.bat` line 88 to point to your installation

### Build Fails

**Error:** `BUILD FAILED!`

**Solution:**

1. Check PyInstaller output for errors
2. Ensure all dependencies are installed
3. Try: `pip install -r requirements.txt`

### Installer Won't Run

**Error:** Windows blocks installer

**Solution:**

1. Right-click installer → Properties
2. Check "Unblock" if present
3. Or: "Run as administrator"

### Application Won't Start After Install

**Solution:**

1. Check Windows Event Viewer for errors
2. Test `dist\LogLife_Operacao\LogLife_Operacao.exe` directly
3. Rebuild with `--debug` flag for more info

## File Size Comparison

| Method           | Size   | Pros                     | Cons                |
| ---------------- | ------ | ------------------------ | ------------------- |
| One .exe file    | ~70 MB | Simple                   | Slow startup, large |
| Folder structure | ~50 MB | Fast, debuggable         | Multiple files      |
| Installer        | ~30 MB | Compressed, professional | Requires install    |

## Advanced Options

### Silent Installation

```batch
LogLife_Operations_Setup_v1.0.0.exe /SILENT
```

### Custom Install Location

```batch
LogLife_Operations_Setup_v1.0.0.exe /DIR="C:\Custom\Path"
```

### Unattended Installation

```batch
LogLife_Operations_Setup_v1.0.0.exe /VERYSILENT /SUPPRESSMSGBOXES /NORESTART
```

## Best Practices

1. ✅ Always test both folder and installer before distribution
2. ✅ Update version numbers in all locations
3. ✅ Test uninstall to ensure clean removal
4. ✅ Include release notes in installer
5. ✅ Sign installer for production (optional)
6. ✅ Test on clean Windows VM

## For Developers

### Converting from One-File to Folder

The change was made in `main_ctk.spec`:

**Before (One-File):**

```python
exe = EXE(
    pyz,
    a.scripts,
    a.binaries,  # Everything bundled in exe
    a.zipfiles,
    a.datas,
    ...
)
```

**After (Folder):**

```python
exe = EXE(
    pyz,
    a.scripts,
    [],  # Empty
    exclude_binaries=True,  # Put in folder
    ...
)

coll = COLLECT(  # New section
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    ...
)
```

---

**Ready to build?** Run `build_with_installer.bat` and create your professional installer! 🚀
