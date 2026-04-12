@echo off
cd /d "%~dp0"
python -m pip install -r requirements.txt
python -m pip install pyinstaller
python -m PyInstaller --onedir --name WorklogAutomation web_app.py
if exist worklog_set.xlsx copy /Y worklog_set.xlsx dist\WorklogAutomation\worklog_set.xlsx
echo.
echo Build complete.
echo Run dist\WorklogAutomation\WorklogAutomation.exe
echo Keep dist\WorklogAutomation\worklog_set.xlsx next to the exe.
pause
