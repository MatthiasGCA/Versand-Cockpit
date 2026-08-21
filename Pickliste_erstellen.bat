@echo off
setlocal enabledelayedexpansion
chcp 65001 >nul
REM ===========================================================================
REM  Pickliste_erstellen.bat  -  Gasecenter
REM  Erzeugt aus den Amicron-Rechnungs-PDFs die Kommissionierliste, kopiert die
REM  vier Bruecken-CSVs in den Netzwerkordner (fuer scan_druck.py) und
REM  ARCHIVIERT die aufgenommenen Rechnungen, damit keine doppelt auf eine
REM  spaetere Packliste geraten.
REM
REM  Die Pickliste-PDF wird mit Zeitstempel benannt (Pickliste_JJJJ-MM-TT_HHMM.pdf)
REM  und daher NICHT mehr ueberschrieben - alle Laeufe bleiben im Ausgabeordner
REM  erhalten. Die vier CSVs behalten feste Namen (so erwartet sie scan_druck.py).
REM
REM  2026-08-21: Fehlerpruefung beim CSV-Kopieren ergaenzt. Vorher pruefte keiner
REM    der vier "copy /Y"-Befehle den Erfolg - war der Netzwerkordner kurz nicht
REM    erreichbar (kommt vor, siehe Vorfall vom 06.08.), meldete das Skript trotzdem
REM    "Fertig", waehrend scan_druck.py in Wirklichkeit mit den ALTEN CSVs weiterlief.
REM    Jetzt wird die Erreichbarkeit vorab geprueft und jeder Kopiervorgang einzeln
REM    auf Erfolg kontrolliert; bei einem Fehler erscheint eine deutliche Warnung
REM    statt der normalen "Fertig"-Meldung.
REM
REM  >>> Nur die Pfade unten pruefen/anpassen, dann die Datei doppelklicken. <<<
REM ===========================================================================

REM --- Python-Startbefehl (meist "py", sonst voller Pfad zur python.exe) ------
set "PYTHON=py"

REM --- Pfad zur packliste.py -------------------------------------------------
set "SKRIPT=C:\Packlisten\packliste.py"

REM --- Ordner mit den Rechnungs-PDFs ----------------------------------------
set "RECHNUNGEN=C:\Packlisten"

REM --- Ausgabeordner (Pickliste-PDF + CSVs) - MUSS ein Unterordner sein, ----
REM     damit die Pickliste nicht wieder als Rechnung eingelesen wird ---------
set "AUSGABE=C:\Packlisten\Pickliste"

REM --- Archivordner fuer die verarbeiteten Rechnungen -----------------------
REM     (Unterordner; das Skript legt darin je Tag einen Ordner an) ----------
set "ARCHIV=C:\Packlisten\Archiv"

REM --- Netzwerkordner, den scan_druck.py ueberwacht -------------------------
set "NETZWERK=\\DESKTOP-N2H75H\Netzwerk\Paketscheine"

REM ===========================================================================
echo.
echo === Pickliste wird erstellt ===
if not exist "%AUSGABE%" mkdir "%AUSGABE%"

REM --- Zeitstempel JJJJ-MM-TT_HHMM (locale-unabhaengig ueber PowerShell) ------
for /f %%i in ('powershell -NoProfile -Command "Get-Date -Format yyyy-MM-dd_HHmm"') do set "STAMP=%%i"
if not defined STAMP set "STAMP=unbenannt"
set "PICKPDF=%AUSGABE%\Pickliste_%STAMP%.pdf"

"%PYTHON%" "%SKRIPT%" "%RECHNUNGEN%" "" "%PICKPDF%" "%ARCHIV%"
if errorlevel 1 (
  echo.
  echo *** FEHLER bei der Erstellung - Abbruch, nichts kopiert/archiviert. ***
  pause
  exit /b 1
)

echo.
echo === CSVs in den Netzwerkordner kopieren ===
set "KOPIER_FEHLER=0"
if not exist "%NETZWERK%\" (
  echo.
  echo *** WARNUNG: Netzwerkordner "%NETZWERK%" nicht erreichbar - CSVs NICHT kopiert^! ***
  echo *** scan_druck.py arbeitet bis zum naechsten erfolgreichen Lauf mit den ALTEN   ***
  echo *** Daten weiter. Bitte Netzwerkfreigabe pruefen und diese Bat-Datei danach     ***
  echo *** erneut ausfuehren.                                                          ***
  set "KOPIER_FEHLER=1"
) else (
  copy /Y "%AUSGABE%\post_zuordnung.csv"   "%NETZWERK%\post_zuordnung.csv"   || set "KOPIER_FEHLER=1"
  copy /Y "%AUSGABE%\sammel_zuordnung.csv" "%NETZWERK%\sammel_zuordnung.csv" || set "KOPIER_FEHLER=1"
  copy /Y "%AUSGABE%\mengen_zuordnung.csv" "%NETZWERK%\mengen_zuordnung.csv" || set "KOPIER_FEHLER=1"
  copy /Y "%AUSGABE%\ean_zuordnung.csv"    "%NETZWERK%\ean_zuordnung.csv"    || set "KOPIER_FEHLER=1"
  REM "!KOPIER_FEHLER!" statt "%KOPIER_FEHLER%": innerhalb DESSELBEN runden-
  REM Klammer-Blocks wird %VAR% schon beim Einlesen des ganzen Blocks ersetzt
  REM (also mit dem Stand VOR den obigen copy-Befehlen) - ohne "setlocal
  REM enabledelayedexpansion" + "!VAR!" wuerde diese Pruefung nie anschlagen,
  REM selbst wenn ein copy wirklich fehlschlug.
  if "!KOPIER_FEHLER!"=="1" (
    echo.
    echo *** WARNUNG: mindestens eine CSV konnte NICHT kopiert werden - siehe oben^! ***
  )
)

echo.
if "!KOPIER_FEHLER!"=="1" (
  echo === FERTIG MIT WARNUNG - CSVs NICHT vollstaendig aktualisiert, siehe oben^! ===
) else (
  echo === Fertig ===
)
echo Pickliste:    %PICKPDF%
echo CSVs nach:    %NETZWERK%
echo Archiviert:   %ARCHIV%\(Datum)
echo.
REM Pickliste zum Drucken oeffnen (REM davor deaktiviert das Oeffnen)
start "" "%PICKPDF%"

pause
