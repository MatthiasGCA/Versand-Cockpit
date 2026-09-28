' ===========================================================
'  wc_auftragsimport.vbs  (Gasecenter Augsburg / Faktura-PC)
'  Startet wc_auftragsimport.py --commit KOMPLETT unsichtbar und
'  haengt die Ausgabe an eine Log-Datei an - gleicher Aufbau wie
'  wc_sendungsnummer_sync.vbs.
'
'  Diese Datei liegt in C:\Scripts\ (gleicher Ordner wie
'  wc_auftragsimport.py und wc_auftragsimport_config.json).
' ===========================================================

Set shell = CreateObject("WScript.Shell")

befehl = "cmd /c py C:\Scripts\wc_auftragsimport.py --commit >> C:\Scripts\wc_auftragsimport.log 2>&1"

shell.Run befehl, 0, False
