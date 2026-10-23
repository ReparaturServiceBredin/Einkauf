@echo off
rem Bedarfsliste per Drag ^& Drop auf diese Datei ziehen - oder doppelklicken,
rem dann werden alle Dateien im Ordner "eingang" verarbeitet. Ergebnis oeffnet sich in Excel.
cd /d "%~dp0"
set PYTHONUTF8=1
".venv\Scripts\python.exe" einkauf.py %* --oeffnen
echo.
pause
