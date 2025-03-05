@echo off
cd /d "%~dp0"

:: Force activate the correct Conda environment
CALL C:\Users\serio\anaconda3\Scripts\activate.bat
CALL conda activate RDC_AI_Vision

:: Run the Python script with the correct interpreter
C:\Users\serio\anaconda3\envs\RDC_AI_Vision\python.exe RDC_Vosk_STT.py

:: Hold terminal open if errors occur
echo.
echo If errors occurred, read above. Press any key to close...
pause >nul
