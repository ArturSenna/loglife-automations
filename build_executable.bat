@echo off
echo =========================================
echo    SLA LogLife Executable Builder v2.0
echo =========================================

set "BUILD_TYPE=%1"
if "%BUILD_TYPE%"=="" set "BUILD_TYPE=standard"

echo Build Type: %BUILD_TYPE%
echo.

:: Check if virtual environment is activated
if not defined VIRTUAL_ENV (
    echo Warning: No virtual environment detected.
    echo It's recommended to use a virtual environment.
    echo.
)

:: Ensure we're in the correct directory
cd /d "%~dp0"

echo Step 1: Installing/Updating PyInstaller...
pip install --upgrade pyinstaller pyinstaller-hooks-contrib

echo.
echo Step 2: Cleaning previous builds...
if exist "dist" rmdir /s /q "dist"
if exist "build" rmdir /s /q "build"
if exist "__pycache__" rmdir /s /q "__pycache__"

echo.
echo Step 3: Building executable...

if "%BUILD_TYPE%"=="optimized" (
    echo Building optimized version...
    pyinstaller SLA_optimized.spec --clean --noconfirm
    set "EXE_NAME=SLA_LogLife_Optimized.exe"
) else (
    echo Building standard version...
    pyinstaller SLA.spec --clean --noconfirm
    set "EXE_NAME=SLA_LogLife.exe"
)

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo ERROR: Build failed!
    echo Check the output above for error details.
    pause
    exit /b 1
)

echo.
echo Step 4: Checking build results...
if exist "dist\%EXE_NAME%" (
    echo SUCCESS: Executable created successfully!
    echo Location: %CD%\dist\%EXE_NAME%
    echo.
    
    :: Get file size in MB
    for %%I in ("dist\%EXE_NAME%") do (
        set /a "size_mb=%%~zI / 1048576"
        echo File size: %%~zI bytes ^(!size_mb! MB^)
    )
    
    echo.
    echo Step 5: Creating distribution package...
    if not exist "release" mkdir "release"
    copy "dist\%EXE_NAME%" "release\"
    copy "BUILD_README.md" "release\README.md"
    echo Distribution package created in 'release' folder.
    
    echo.
    choice /C YN /M "Do you want to run the executable now"
    if errorlevel 2 goto :end
    if errorlevel 1 (
        echo Starting executable...
        start "" "dist\%EXE_NAME%"
    )
) else (
    echo ERROR: Executable was not created!
    echo Check the build output for errors.
)

:end
echo.
echo Build process completed.
echo.
echo Available commands:
echo   build_executable.bat          - Standard build
echo   build_executable.bat optimized - Optimized build (smaller size)
echo.
pause