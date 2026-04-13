@echo off
setlocal
cd /d "%~dp0"
for /f %%i in ('powershell -NoProfile -ExecutionPolicy Bypass -Command "Get-Date -Format yyyyMMdd_HHmmss"') do set "BUILD_VERSION=%%i"
set "APP_NAME=WorklogAutomation_%BUILD_VERSION%"

echo Building %APP_NAME%
python -m pip install -r requirements.txt
python -m pip install pyinstaller
python -m PyInstaller --clean --onedir --name "%APP_NAME%" --hidden-import tkinter --hidden-import tkinter.filedialog --collect-all streamlit --add-data "streamlit_app.py;." --add-data "automate_worklog.py;." streamlit_launcher.py
echo %BUILD_VERSION%> "dist\%APP_NAME%\version.txt"
if exist worklog_set1.xlsx copy /Y worklog_set1.xlsx "dist\%APP_NAME%\worklog_set1.xlsx"
if exist worklog_set2.xlsx copy /Y worklog_set2.xlsx "dist\%APP_NAME%\worklog_set2.xlsx"
if exist "dist\%APP_NAME%.zip" del /Q "dist\%APP_NAME%.zip"
powershell -NoProfile -ExecutionPolicy Bypass -Command "Compress-Archive -Path 'dist\%APP_NAME%\*' -DestinationPath 'dist\%APP_NAME%.zip' -Force"
echo.
echo Build complete.
echo Version: %BUILD_VERSION%
echo EXE: dist\%APP_NAME%\%APP_NAME%.exe
echo ZIP: dist\%APP_NAME%.zip
echo Send dist\%APP_NAME%.zip to another PC, unzip it, then run %APP_NAME%.exe.
echo Do not send the exe alone. The _internal folder must stay next to the exe.
pause
