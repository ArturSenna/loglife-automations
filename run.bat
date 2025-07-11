@echo off
REM Run the automation system in the virtual environment

if not exist "venv\Scripts\activate.bat" (
    echo Virtual environment not found. Run setup_venv.bat first.
    pause
    exit /b 1
)

echo Activating virtual environment and running LogLife Automations...
call venv\Scripts\activate.bat
python main.py

echo.
echo Automation completed. Press any key to exit.
pause >nul
