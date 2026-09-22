' ===========================================================
'  Carrier-Dashboard_starten.vbs
'  Startet carrier_dashboard.py KOMPLETT unsichtbar - es geht NUR das
'  Dashboard-Fenster auf, kein zusaetzliches schwarzes Konsolenfenster
'  (derselbe Trick wie bei wc_sendungsnummer_sync.vbs/PaketImport-Mover.vbs:
'  WScript.Shell.Run mit Fensterstil 0 statt z.B. "-WindowStyle Hidden",
'  das kurz aufblitzen kann).
'
'  Ohne Umleitung waere eine Fehlermeldung bei einem unsichtbaren Lauf
'  nirgends nachvollziehbar - deshalb wird die Ausgabe an
'  Carrier-Dashboard.log (gleicher Ordner) angehaengt.
'
'  Ermittelt den eigenen Ordner selbst (wie %~dp0 im .bat) - funktioniert
'  daher unveraendert auch nach einem Umzug, z.B. auf den Faktura-PC.
' ===========================================================

Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
ordner = fso.GetParentFolderName(WScript.ScriptFullName)

befehl = "cmd /c py """ & ordner & "\carrier_dashboard.py"" >> """ & _
         ordner & "\Carrier-Dashboard.log"" 2>&1"

shell.Run befehl, 0, False
