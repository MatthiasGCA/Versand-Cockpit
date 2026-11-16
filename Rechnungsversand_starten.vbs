' ===========================================================
'  Rechnungsversand_starten.vbs
'  Startet den taeglichen Rechnungsversand (rechnungsversand.py) UNSICHTBAR - ohne schwarzes
'  Konsolenfenster, das man versehentlich schliessen koennte (real 16.11.2026: Lauf mitten im
'  Versand beendet, 24 von ~127 Mails raus). Fuer die Windows-Aufgabenplanung:
'    schtasks /create /f /tn "Rechnungsversand" /tr "wscript.exe \"C:\Carrier-Dashboard\Rechnungsversand_starten.vbs\"" /sc daily /st 15:00
'  Das Protokoll steht wie immer in C:\Rechnungsversand\Ausgabe\rechnungsversand.log; zusaetzliche
'  Programmausgaben (z.B. Python-Fehler) in Rechnungsversand_konsole.log neben diesem Skript.
' ===========================================================

Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
ordner = fso.GetParentFolderName(WScript.ScriptFullName)

befehl = "cmd /c cd /d """ & ordner & """ && py """ & ordner & "\rechnungsversand.py"" >> """ & _
         ordner & "\Rechnungsversand_konsole.log"" 2>&1"

shell.Run befehl, 0, False
