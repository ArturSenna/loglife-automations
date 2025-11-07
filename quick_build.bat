@echo off
title LogLife Operations - Quick Build
color 0A

echo.
echo  ╔════════════════════════════════════════════════════════╗
echo  ║                                                        ║
echo  ║        LogLife Operations - Quick Build Tool          ║
echo  ║                                                        ║
echo  ╚════════════════════════════════════════════════════════╝
echo.
echo  This script will build your executable in 3 easy steps!
echo.
pause

cls
echo.
echo  ╔════════════════════════════════════════════════════════╗
echo  ║  Step 1/3: Environment Check                           ║
echo  ╚════════════════════════════════════════════════════════╝
echo.

:: Check Python
python --version >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    color 0C
    echo  [ERROR] Python is not installed or not in PATH!
    echo.
    echo  Please install Python 3.7+ from python.org
    echo.
    pause
    exit /b 1
)

echo  [OK] Python is installed
python --version

:: Activate virtual environment if exists
if exist "venv\Scripts\activate.bat" (
    echo  [OK] Activating virtual environment...
    call venv\Scripts\activate.bat
) else if exist ".venv\Scripts\activate.bat" (
    echo  [OK] Activating virtual environment...
    call .venv\Scripts\activate.bat
) else (
    echo  [INFO] No virtual environment found, using system Python
)

echo.
pause

cls
echo.
echo  ╔════════════════════════════════════════════════════════╗
echo  ║  Step 2/3: Installing Build Tools                      ║
echo  ╚════════════════════════════════════════════════════════╝
echo.

pip install --upgrade pyinstaller pyinstaller-hooks-contrib pillow --quiet
if %ERRORLEVEL% NEQ 0 (
    color 0C
    echo  [ERROR] Failed to install build tools!
    pause
    exit /b 1
)

echo  [OK] Build tools installed successfully
echo.
pause

cls
echo.
echo  ╔════════════════════════════════════════════════════════╗
echo  ║  Step 3/3: Building Executable                         ║
echo  ╚════════════════════════════════════════════════════════╝
echo.

:: Clean old builds
echo  [INFO] Cleaning old builds...
if exist "dist" rmdir /s /q "dist" 2>nul
if exist "build" rmdir /s /q "build" 2>nul

:: Build
echo  [INFO] Building executable (this may take 2-5 minutes)...
echo.
pyinstaller main_ctk.spec --clean --noconfirm

if %ERRORLEVEL% NEQ 0 (
    color 0C
    echo.
    echo  ╔════════════════════════════════════════════════════════╗
    echo  ║  BUILD FAILED!                                         ║
    echo  ╚════════════════════════════════════════════════════════╝
    echo.
    echo  Please check the errors above and try again.
    echo.
    pause
    exit /b 1
)

:: Copy to release folder
if not exist "release" mkdir release
copy dist\LogLife_Operacao.exe release\LogLife_Operacao.exe >nul 2>&1

color 0A
cls
echo.
echo  ╔════════════════════════════════════════════════════════╗
echo  ║                                                        ║
echo  ║               BUILD SUCCESSFUL!                        ║
echo  ║                                                        ║
echo  ╚════════════════════════════════════════════════════════╝
echo.
echo  Your executable is ready:
echo.
echo  📁 Main location:    dist\LogLife_Operacao.exe
echo  📁 Release copy:     release\LogLife_Operacao.exe
echo.
echo  ═══════════════════════════════════════════════════════
echo  Next Steps:
echo  ═══════════════════════════════════════════════════════
echo.
echo  1. Test the executable by running it
echo  2. Check functionality of all features
echo  3. Distribute to users from the release folder
echo.
echo  For update system setup, see BUILD_UPDATE_GUIDE.md
echo.
pause

:: Option to open folder
choice /C YN /M "Do you want to open the release folder"
if errorlevel 2 goto :end
if errorlevel 1 explorer release

:end
echo.
echo  Thank you for using LogLife Operations Builder!
echo.
timeout /t 3 >nul
