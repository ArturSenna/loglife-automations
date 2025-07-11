@echo off
REM Quick activation script for the virtual environment

if not exist "venv\Scripts\activate.bat" (
    echo Virtual environment not found. Run setup_venv.bat first.
    pause
    exit /b 1
)

echo Activating LogLife Automations virtual environment...
call venv\Scripts\activate.bat

echo Virtual environment activated!
echo You can now run: python main.py
cmd /k
