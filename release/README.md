# SLA LogLife - Executable Build Instructions

This document explains how to create a standalone executable from the SLA.py script.

## Prerequisites

1. **Python Environment**: Ensure you have Python 3.7+ installed
2. **Dependencies**: All required packages should be installed from requirements.txt
3. **Virtual Environment**: Recommended to use a virtual environment

## Quick Start

### Option 1: Automated Build (Recommended)

1. Open Command Prompt or PowerShell as Administrator
2. Navigate to the project directory
3. Run the build script:
   ```batch
   build_executable.bat
   ```

### Option 2: Manual Build

1. Install PyInstaller if not already installed:

   ```bash
   pip install pyinstaller pyinstaller-hooks-contrib
   ```

2. Build the executable:
   ```bash
   pyinstaller SLA.spec --clean --noconfirm
   ```

## Build Configuration

The build is configured using `SLA.spec` file which includes:

- **Main Script**: `src\automations\SLA.py`
- **Icon**: `src\assets\my_icon.ico`
- **Hidden Imports**: All required modules for proper functionality
- **Data Files**: Icon and configuration files are bundled
- **Output**: Single executable file (no console window)

## Output

After successful build:

- Executable location: `dist\SLA_LogLife.exe`
- Size: Approximately 200-300 MB (includes all dependencies)
- No installation required - can be run on any Windows machine

## Troubleshooting

### Common Issues:

1. **Import Errors**:

   - Ensure all dependencies are installed
   - Check that you're using the correct Python environment

2. **Missing Icon**:

   - Verify `src\assets\my_icon.ico` exists
   - Icon will be skipped if missing (application still works)

3. **Build Fails**:

   - Try cleaning previous builds: delete `dist` and `build` folders
   - Ensure no antivirus is blocking the process
   - Run as Administrator if needed

4. **Large File Size**:
   - This is normal for bundled Python applications
   - Consider using `--exclude-module` for unused large modules

### Advanced Optimization:

1. **Reduce Size**:

   ```bash
   pyinstaller SLA.spec --clean --noconfirm --strip
   ```

2. **Debug Mode** (if executable doesn't start):

   ```bash
   pyinstaller SLA.spec --clean --noconfirm --debug=all
   ```

3. **One Directory** (instead of one file):
   - Modify `SLA.spec`: Change `onefile=True` to `onefile=False`
   - Results in faster startup but multiple files

## Distribution

The generated `SLA_LogLife.exe` can be distributed as-is:

- No Python installation required on target machines
- No additional dependencies needed
- Works on Windows 7/10/11 (64-bit)

## File Structure After Build

```
loglife-automations/
├── dist/
│   └── SLA_LogLife.exe          # Main executable
├── build/                       # Build artifacts (can be deleted)
├── SLA.spec                     # PyInstaller specification
├── build_executable.bat         # Build script
└── src/
    ├── automations/
    │   └── SLA.py              # Modified source (with icon support)
    └── assets/
        └── my_icon.ico         # Application icon
```

## Notes

- The executable includes the Python interpreter and all dependencies
- First run may be slower due to unpacking
- Antivirus software might flag the executable (false positive)
- Test thoroughly before distribution
