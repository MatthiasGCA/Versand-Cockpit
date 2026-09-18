' ===========================================================
'  wc_sendungsnummer_sync.vbs  (Gasecenter Augsburg / Faktura-PC)
'  Startet wc_sendungsnummer_sync.py --commit KOMPLETT unsichtbar
'  (kein Konsolenfenster) und haengt die Ausgabe an eine Log-Datei
'  an - derselbe Trick wie bei PaketImport-Mover.vbs (WScript.Shell.Run
'  mit Fensterstil 0 statt "-WindowStyle Hidden", das kurz aufblitzen
'  kann).
'
'  Ohne Umweg ueber "cmd /c" mit Umleitung wuerde die Konsolen-
'  ausgabe des Python-Skripts bei einem unsichtbaren Lauf einfach
'  verloren gehen - dann liesse sich ein Fehlschlag nirgends
'  nachvollziehen.
'
'  Diese Datei liegt in C:\Scripts\ (gleicher Ordner wie
'  wc_sendungsnummer_sync.py und wc_sync_config.json).
' ===========================================================

Set shell = CreateObject("WScript.Shell")

befehl = "cmd /c py C:\Scripts\wc_sendungsnummer_sync.py --commit >> C:\Scripts\wc_sendungsnummer_sync.log 2>&1"

shell.Run befehl, 0, False
