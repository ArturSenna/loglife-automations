@echo off
title LogLife Operations - Full Build with Installer
color 0A

echo.
echo  ╔════════════════════════════════════════════════════════╗
echo  ║                                                        ║
echo  ║   LogLife Operations - Build + Installer Creator      ║
echo  ║                                                        ║
echo  ╚════════════════════════════════════════════════════════╝
echo.
echo  This script will:
echo  1. Build the application as a folder structure
echo  2. Create a Windows installer (.exe)
echo.
pause

cls
echo.
echo  ╔════════════════════════════════════════════════════════╗
echo  ║  Step 1/4: Environment Check                           ║
echo  ╚════════════════════════════════════════════════════════╝
echo.

:: Check Python
python --version >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    color 0C
    echo  [ERROR] Python is not installed or not in PATH!
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
echo  ║  Step 2/4: Installing Build Tools                      ║
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
echo  ║  Step 3/4: Building Application                        ║
echo  ╚════════════════════════════════════════════════════════╝
echo.

:: Clean old builds
echo  [INFO] Cleaning old builds...
if exist "dist" rmdir /s /q "dist" 2>nul
if exist "build" rmdir /s /q "build" 2>nul
if exist "installer_output" rmdir /s /q "installer_output" 2>nul

:: Build
echo  [INFO] Building application folder (this may take 3-5 minutes)...
echo.
pyinstaller main_ctk.spec --clean --noconfirm

if %ERRORLEVEL% NEQ 0 (
    color 0C
    echo.
    echo  ╔════════════════════════════════════════════════════════╗
    echo  ║  BUILD FAILED!                                         ║
    echo  ╚════════════════════════════════════════════════════════╝
    echo.
    pause
    exit /b 1
)

echo.
echo  [OK] Application built successfully
pause

cls
echo.
echo  ╔════════════════════════════════════════════════════════╗
echo  ║  Step 4/4: Creating Installer                          ║
echo  ╚════════════════════════════════════════════════════════╝
echo.

:: Check for Inno Setup
set "INNO_PATH=C:\Program Files (x86)\Inno Setup 6\ISCC.exe"
if not exist "%INNO_PATH%" (
    echo  [WARNING] Inno Setup not found at default location
    echo.
    echo  Please install Inno Setup 6 from:
    echo  https://jrsoftware.org/isdl.php
    echo.
    echo  Or specify the correct path in this script.
    echo.
    echo  Skipping installer creation...
    goto :skip_installer
)

echo  [INFO] Creating Windows installer...
echo.
"%INNO_PATH%" "installer.iss"

if %ERRORLEVEL% NEQ 0 (
    color 0C
    echo.
    echo  [ERROR] Installer creation failed!
    pause
    exit /b 1
)

echo.
echo  [OK] Installer created successfully

:skip_installer

color 0A
cls
echo.
echo  ╔════════════════════════════════════════════════════════╗
echo  ║                                                        ║
echo  ║               BUILD SUCCESSFUL!                        ║
echo  ║                                                        ║
echo  ╚════════════════════════════════════════════════════════╝
echo.
echo  Your application is ready:
echo.
echo  📁 Application folder:  dist\LogLife_Operacao\
echo  📁 Main executable:     dist\LogLife_Operacao\LogLife_Operacao.exe

if exist "installer_output\*.exe" (
    echo  📦 Installer:          installer_output\LogLife_Operations_Setup_v1.0.0.exe
)

echo.
echo  ═══════════════════════════════════════════════════════
echo  Next Steps:
echo  ═══════════════════════════════════════════════════════
echo.
echo  1. Test the application in dist\LogLife_Operacao\
echo  2. If installer was created, test it
echo  3. Distribute the installer to users
echo.
echo  Benefits of folder structure + installer:
echo  • Faster startup time
echo  • Easier debugging
echo  • Professional installation experience
echo  • Automatic uninstaller
echo  • Start menu shortcuts
echo.
pause

:: Option to open folder
choice /C YN /M "Do you want to open the dist folder"
if errorlevel 2 goto :check_installer
if errorlevel 1 explorer dist\LogLife_Operacao

:check_installer
if exist "installer_output\*.exe" (
    choice /C YN /M "Do you want to open the installer folder"
    if errorlevel 1 explorer installer_output
)

:end
echo.
echo  Thank you for using LogLife Operations Builder!
echo.
timeout /t 3 >nul
