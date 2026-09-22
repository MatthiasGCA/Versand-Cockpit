@echo off
setlocal
chcp 65001 >nul
REM ===========================================================================
REM  Carrier-Dashboard_starten.bat
REM  Startet das Carrier-Dashboard per Doppelklick.
REM
REM  "%~dp0" ist immer der Ordner, in dem DIESE bat-Datei liegt - dadurch
REM  funktioniert die Datei unveraendert sowohl hier auf dem Laptop als auch
REM  spaeter auf dem Faktura-PC, egal wie der Ordner dort heisst, solange
REM  carrier_dashboard.py + carrier_regeln.py + carrier_export.py + packliste.py
REM  daneben liegen (das Dashboard sucht sie im selben Ordner).
REM ===========================================================================
cd /d "%~dp0"

py carrier_dashboard.py
if errorlevel 1 (
  echo.
  echo *** Das Dashboard wurde mit einem Fehler beendet - siehe oben. ***
  pause
)
