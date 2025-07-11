@echo off
REM Setup script for creating Python virtual environment on Windows

echo Setting up Python virtual environment for LogLife Automations...

REM Check if Python is installed
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo Python is not installed or not in PATH.
    echo Please install Python from https://python.org and try again.
    pause
    exit /b 1
)

REM Display Python version
echo Current Python version:
python --version

REM Create virtual environment
echo Creating virtual environment...
python -m venv venv

REM Activate virtual environment and install dependencies
echo Activating virtual environment...
call venv\Scripts\activate.bat

echo Upgrading pip...
python -m pip install --upgrade pip

echo Installing project dependencies...
pip install -r requirements.txt

echo.
echo ===================================
echo Virtual environment setup complete!
echo ===================================
echo.
echo To activate the virtual environment, run:
echo   venv\Scripts\activate.bat
echo.
echo To deactivate, run:
echo   deactivate
echo.
echo To run the automation system:
echo   python main.py
echo.

pause
