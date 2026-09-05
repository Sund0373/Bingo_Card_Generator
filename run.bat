@echo off
REM Double-click launcher. Creates a local virtual environment on first run,
REM installs dependencies, then starts the web UI and opens a browser.

cd /d "%~dp0"

where python >nul 2>&1
if errorlevel 1 (
    echo.
    echo Python was not found on PATH.
    echo Install Python 3.10 or newer from https://www.python.org/downloads/
    echo and tick "Add python.exe to PATH" during setup.
    echo.
    pause
    exit /b 1
)

if not exist ".venv\Scripts\python.exe" (
    echo Creating virtual environment...
    python -m venv .venv
    if errorlevel 1 (
        echo Could not create the virtual environment.
        pause
        exit /b 1
    )
    echo Installing dependencies...
    ".venv\Scripts\python.exe" -m pip install --upgrade pip setuptools
    ".venv\Scripts\python.exe" -m pip install -r requirements.txt
    if errorlevel 1 (
        echo Dependency install failed. See the messages above.
        pause
        exit /b 1
    )
)

echo.
echo Starting the Bingo Card Generator...
".venv\Scripts\python.exe" app.py
pause
