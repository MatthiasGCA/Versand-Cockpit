' ===========================================================
'  live_upload.vbs  (Gasecenter Augsburg / Lager-PC DESKTOP-N2H75H)
'  Startet live_upload.py UNSICHTBAR (kein Konsolenfenster) im Dauerlauf.
'  Das Skript schickt den Live-Stand des Versandcockpits an die
'  Versandstatistik-Webseite (Reiter "Live"). Eine zweite Instanz beendet
'  sich selbst - doppeltes Starten schadet also nicht.
'
'  Einrichten: diese Datei (oder eine Verknuepfung darauf) in den
'  Autostart-Ordner legen (Win+R -> shell:startup), damit sie nach jeder
'  Anmeldung laeuft. Liegt im selben Ordner wie live_upload.py
'  (z.B. C:\Packlisten) und findet ihn selbst.
' ===========================================================

Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
ordner = fso.GetParentFolderName(WScript.ScriptFullName)

befehl = "cmd /c py """ & ordner & "\live_upload.py"""

shell.Run befehl, 0, False
