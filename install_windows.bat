@echo off
cd /d "%~dp0"
echo === Installing Beeroniza Pinterest Generator ===
py -3 --version >nul 2>&1
if %errorlevel%==0 (
    py -3 -m venv .venv
) else (
    python -m venv .venv
)
if errorlevel 1 (
    echo ERROR: Python 3.10+ is required. Install Python and try again.
    pause
    exit /b 1
)
".venv\Scripts\python.exe" -m pip install -r requirements.txt
if errorlevel 1 (
    echo ERROR: Could not install dependencies.
    pause
    exit /b 1
)
echo.
echo Done. Now copy config.example.json to config.json and edit it.
pause
