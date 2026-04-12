@echo off
cd /d "%~dp0"
python -m pip install -r requirements.txt
python -m pip install pyinstaller
python -m PyInstaller --onedir --name WorklogAutomation web_app.py
if exist worklog_set1.xlsx copy /Y worklog_set1.xlsx dist\WorklogAutomation\worklog_set1.xlsx
if exist worklog_set2.xlsx copy /Y worklog_set2.xlsx dist\WorklogAutomation\worklog_set2.xlsx
echo.
echo Build complete.
echo Run dist\WorklogAutomation\WorklogAutomation.exe
echo Keep worklog_set1.xlsx and worklog_set2.xlsx next to the exe.
pause
