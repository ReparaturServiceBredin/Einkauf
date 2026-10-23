@echo off
rem Bedarfsliste per Drag ^& Drop auf diese Datei ziehen - oder doppelklicken,
rem dann werden alle Dateien im Ordner "eingang" verarbeitet. Ergebnis oeffnet sich in Excel.
cd /d "%~dp0"
set PYTHONUTF8=1
if not exist ".venv\Scripts\python.exe" (
  echo FEHLER: .venv fehlt. Bitte Installation laut README.md, Schritt 3, ausfuehren.
  pause
  exit /b 1
)
".venv\Scripts\python.exe" einkauf.py %* --oeffnen
if errorlevel 1 (
  echo.
  echo ==== FEHLER - Details siehe oben bzw. im Ordner logs ====
) else (
  echo.
  echo ==== Fertig. Ergebnisse liegen im Ordner ausgabe ====
)
echo Fenster schliessen: beliebige Taste druecken.
pause >nul
