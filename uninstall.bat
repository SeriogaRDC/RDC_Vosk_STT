@echo off
echo Starting uninstallation...

:: Deactivate virtual environment if active
if defined VIRTUAL_ENV (
    call venv\Scripts\deactivate.bat
)

:: Remove virtual environment directory
if exist venv (
    echo Removing virtual environment...
    rmdir /s /q venv
    if errorlevel 1 (
        echo Failed to remove virtual environment!
        pause
        exit /b 1
    )
)

echo Uninstallation completed successfully!
pause
