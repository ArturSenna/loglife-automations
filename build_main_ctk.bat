@echo off
echo =========================================================
echo    LogLife Operations - Executable Builder
echo =========================================================
echo.

:: Check if virtual environment is activated
if not defined VIRTUAL_ENV (
    echo Warning: No virtual environment detected.
    echo It's recommended to use a virtual environment.
    echo.
    echo Attempting to activate virtual environment...
    if exist "venv\Scripts\activate.bat" (
        call venv\Scripts\activate.bat
        echo Virtual environment activated.
    ) else if exist ".venv\Scripts\activate.bat" (
        call .venv\Scripts\activate.bat
        echo Virtual environment activated.
    ) else (
        echo No virtual environment found. Continuing with system Python...
    )
    echo.
)

:: Ensure we're in the correct directory
cd /d "%~dp0"

echo Step 1: Installing/Updating Build Dependencies...
pip install --upgrade pyinstaller pyinstaller-hooks-contrib pillow
if %ERRORLEVEL% NEQ 0 (
    echo ERROR: Failed to install dependencies!
    pause
    exit /b 1
)

echo.
echo Step 2: Cleaning Previous Builds...
if exist "dist" rmdir /s /q "dist"
if exist "build" rmdir /s /q "build"
for /d /r "src" %%d in (__pycache__) do @if exist "%%d" rmdir /s /q "%%d"
echo Clean completed.

echo.
echo Step 3: Building Executable...
echo Building LogLife Operations application...
pyinstaller main_ctk.spec --clean --noconfirm

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo ERROR: Build failed!
    echo Check the output above for error details.
    pause
    exit /b 1
)

echo.
echo =========================================================
echo Build completed successfully!
echo.
echo Executable location: dist\LogLife_Operacao.exe
echo.
echo You can now distribute the executable file.
echo =========================================================
echo.

:: Optional: Copy version file to dist
if exist "version.json" (
    copy version.json dist\version.json >nul
    echo Version file copied to dist folder.
)

:: Optional: Create release folder
if not exist "release" mkdir release
if exist "dist\LogLife_Operacao.exe" (
    copy dist\LogLife_Operacao.exe release\LogLife_Operacao.exe >nul
    echo.
    echo Executable also copied to release folder.
)

echo.
pause
