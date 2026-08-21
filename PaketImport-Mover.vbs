' ===========================================================
'  PaketImport-Mover.vbs  (Gasecenter Augsburg / Faktura-PC)
'  Startet PaketImport-Mover.ps1 KOMPLETT unsichtbar - kein
'  Konsolenfenster, auch kein kurzes Aufblitzen.
'
'  Hintergrund: "-WindowStyle Hidden" bei PowerShell allein
'  verhindert das kurze Aufblitzen des Fensters NICHT zuver-
'  laessig (bekannter Windows-Effekt: das Fenster wird kurz
'  erzeugt, bevor der Hidden-Stil greift). Der Umweg ueber
'  WScript.Shell.Run mit Fensterstil 0 vermeidet das komplett
'  - derselbe Trick wie bei Versand-Cockpit.vbs.
'
'  KORRIGIERT: das eigentliche Skript liegt unter C:\Scripts\,
'  nicht unter C:\AfterSell\ (das ist nur das Move-Ziel der
'  Sendungsnummern-CSVs, ein anderer Ordner).
'
'  Diese Datei liegt in C:\Scripts\ (gleicher Ordner wie
'  PaketImport-Mover.ps1).
' ===========================================================

Set shell = CreateObject("WScript.Shell")

skript = "C:\Scripts\PaketImport-Mover.ps1"

shell.Run "powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File """ & skript & """", 0, False
