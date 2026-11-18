@echo off
cd /d "%~dp0"
if not exist "config.json" (
    echo ERROR: Create config.json from config.example.json first.
    pause
    exit /b 1
)
if not exist ".venv\Scripts\python.exe" (
    echo ERROR: Run install_windows.bat first.
    pause
    exit /b 1
)
".venv\Scripts\python.exe" generate_pins.py %*
echo.
pause
