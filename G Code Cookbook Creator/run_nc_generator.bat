@echo off
echo Checking for Python installation...

python --version >nul 2>&1
if errorlevel 1 (
    echo Python is not installed or not in PATH
    echo Please install Python from https://www.python.org/downloads/
    pause
    exit /b 1
)

echo Running NC Generator...
python nc_generator.py

if errorlevel 1 (
    echo Error running NC Generator
    pause
    exit /b 1
)

pause