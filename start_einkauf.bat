@echo off
rem Startet den Einkaufs-Abgleich. Fuer die Windows-Aufgabenplanung gedacht.
cd /d "%~dp0"
set PYTHONUTF8=1
if not exist logs mkdir logs
".venv\Scripts\python.exe" einkauf.py %* >> "logs\aufgabenplanung.log" 2>&1
exit /b %ERRORLEVEL%
