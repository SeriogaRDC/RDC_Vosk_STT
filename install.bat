@echo off
echo Starting installation...

:: Check if Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo Python not found! Please install Python 3.8 or higher.
    pause
    exit /b 1
)

:: Create virtual environment
echo Creating virtual environment...
python -m venv venv
if errorlevel 1 (
    echo Failed to create virtual environment!
    pause
    exit /b 1
)

:: Activate virtual environment and install requirements
echo Installing dependencies...
call venv\Scripts\activate.bat
python -m pip install --upgrade pip
pip install -r requirements.txt

:: Verify installation
python -c "import vosk; import sounddevice; import numpy; import wave" >nul 2>&1
if errorlevel 1 (
    echo Installation verification failed!
    pause
    exit /b 1
)

:: Create launcher VBS script
echo Creating silent launcher...
(
echo Set WShell = CreateObject^("WScript.Shell"^)
echo strPath = CreateObject^("Scripting.FileSystemObject"^).GetParentFolderName^(WScript.ScriptFullName^)
echo WShell.CurrentDirectory = strPath
echo WShell.Run Chr^(34^) ^& strPath ^& "\venv\Scripts\pythonw.exe" ^& Chr^(34^) ^& " VoskSTT\RDC_Vosk_STT.py", 0, False
echo Set WShell = Nothing
) > "launch_silent.vbs"

:: Create desktop shortcut to VBS launcher
echo Creating desktop shortcut...
set SCRIPT_PATH=%~dp0
set SHORTCUT_PATH=%USERPROFILE%\Desktop\Speech To Text.lnk
powershell -Command "$WS = New-Object -ComObject WScript.Shell; $SC = $WS.CreateShortcut('%SHORTCUT_PATH%'); $SC.TargetPath = 'wscript.exe'; $SC.Arguments = '%SCRIPT_PATH%launch_silent.vbs'; $SC.WorkingDirectory = '%SCRIPT_PATH%'; $SC.Save()"

echo Installation completed successfully!
echo Please use the desktop shortcut to launch the application silently.
pause
