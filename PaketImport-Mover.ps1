# === PaketImport-Mover.ps1 ===
# Verarbeitet den Downloads-Ordner in zwei unabhaengigen Schritten:
#
#   1) Sendungsnummern-Exporte (CSV, Namensmuster unten):
#        a) ZUERST nach C:\Scripts\Sendungsnummern_WC KOPIEREN
#           (Quelle fuer den WooCommerce-Sendungsnummer-Sync, gasecenter-
#           onlineshop.de - laeuft auf DEMSELBEN PC).
#        b) DANN wie bisher nach C:\AfterSell\PaketImport VERSCHIEBEN
#           (Amazon/eBay-Schnittstelle "AfterSell").
#      Die Reihenfolge ist wichtig: Downloads ist die einzige Quelle. Erst wenn
#      die Datei sicher im WC-Ordner liegt, wird sie aus Downloads wegbewegt -
#      sonst kann ein Export dauerhaft nie beim WC-Sync ankommen. Beide Ziele
#      liegen lokal auf C:, das Kopieren sollte also praktisch immer klappen;
#      die Absicherung unten bleibt trotzdem (Platte voll, Rechteproblem).
#
#   2) ECHTE Versandlabel-PDFs (per Inhalt erkannt ueber label_erkennen.py -
#      derselbe Erkennungscode wie in scan_druck.py, NICHT nur "*.pdf": im
#      Downloads-Ordner landen erfahrungsgemaess auch viele andere PDFs, die
#      NICHT nach Paketscheine sollen) -> nach
#      \\DESKTOP-N2H75H\Netzwerk\Paketscheine KOPIEREN, damit scan_druck.py
#      (Lager-PC) sie findet. Das Original wird danach aus Downloads in
#      einen Archiv-Unterordner VERSCHOBEN (nicht geloescht) - sonst wuerde
#      jeder folgende 1-Minuten-Lauf dieselbe Datei erneut kopieren, auch
#      wenn scan_druck.py sie laengst gedruckt hat. Alle anderen PDFs (kein
#      erkanntes Label) bleiben in Downloads unangetastet liegen.
#
# 2026-08-28: Schritt 1 um die WooCommerce-Kopie erweitert (s.o.). Damit die
#   fest benannte "DHL-VLS-Export.csv" bei einem fehlgeschlagenen Move nach
#   AfterSell nicht bei jedem Lauf erneut in den WC-Ordner dupliziert wird,
#   gibt es analog zur PDF-Merkliste eine kleine Statusdatei
#   ($CsvAusstehendDatei): Dateien, die bereits nach Sendungsnummern_WC
#   kopiert wurden, deren Move nach AfterSell aber (noch) nicht klappte,
#   stehen dort und werden beim naechsten Lauf NUR noch verschoben, nicht
#   erneut kopiert. Ein trotzdem entstehendes Duplikat im WC-Ordner ist
#   harmlos - der WC-Sync entdoppelt ueber Rechnungsnummer + Sendungsnummer.
#
# 2026-08-21: Absicherung gegen wiederholtes Doppel-Kopieren ergaenzt. Bisher
#   wurde bei einem fehlgeschlagenen Move-Item (Original -> Downloads-Archiv,
#   z.B. durch eine kurze Antivirus-/Browser-Sperre direkt nach dem Download)
#   NUR geloggt - die Datei blieb unveraendert in Downloads liegen UND wurde
#   NICHT in die Merkliste aufgenommen. Beim naechsten Lauf (alle paar Minuten)
#   wurde sie dadurch erneut als "neues" Label erkannt und ERNEUT nach
#   Paketscheine kopiert (mit neuem Zeitstempel-Namen wegen Namenskollision) -
#   theoretisch unbegrenzt oft, solange die Sperre bestand. Jetzt wird das
#   Archivieren bis zu 3x mit kurzer Pause wiederholt (deckt die uebliche
#   kurze Sperre ab); schlaegt es endgueltig fehl, wird die Datei TROTZDEM in
#   die Merkliste aufgenommen (sie wurde ja bereits erfolgreich nach
#   Paketscheine kopiert - ein erneutes Kopieren waere so oder so falsch) und
#   deutlich als "manuell aus Downloads entfernen" geloggt, statt bei jedem
#   Lauf erneut zu kopieren.

$Quelle          = "$env:USERPROFILE\Downloads"
$ZielAfterSell   = "C:\AfterSell\PaketImport"
$ZielWooCommerce = "C:\Scripts\Sendungsnummern_WC"
$ZielNetzwerk    = "\\DESKTOP-N2H75H\Netzwerk\Paketscheine"
$QuelleArchiv    = Join-Path $Quelle "Verarbeitete-Label"
$LogDatei        = "C:\Scripts\PaketImport-Mover.log"
$LabelErkennung  = "C:\Scripts\label_erkennen.py"
$PythonExe       = "py"    # ggf. auf vollen Pfad umstellen, siehe Hinweis im Chat
# Merkliste bereits geprueft-und-ABGELEHNTER PDFs (Pfad;Groesse;Aenderungszeit).
# Verhindert, dass fremde PDFs, die dauerhaft in Downloads liegen bleiben, bei
# JEDEM 1-Minuten-Lauf erneut per Python/pdfplumber geprueft werden - nur eine
# tatsaechlich veraenderte Datei (anderer Zeitstempel/Groesse) wird neu geprueft.
# Traegt seit 2026-08-21 AUCH Label ein, die zwar erfolgreich nach Paketscheine
# kopiert, aber deren Original NICHT archiviert werden konnte (s.o.) - sonst
# droht dieselbe Duplikat-Kaskade wie beim Vorfall vom 06.08.
$GeprueftDatei   = "C:\Scripts\PaketImport-Mover-geprueft.csv"
# Merkliste fuer Schritt 1: Sendungsnummern-CSVs, die bereits nach
# Sendungsnummern_WC kopiert wurden, deren Move nach AfterSell aber noch
# aussteht. Signatur = Pfad|Groesse|Aenderungszeit (wie $GeprueftDatei).
$CsvAusstehendDatei = "C:\Scripts\PaketImport-Mover-csv-ausstehend.csv"
# Wiederholversuche fuers Verschieben (Original -> AfterSell bzw. -> Label-
# Archiv) nach erfolgreichem Kopieren. Eine kurze Antivirus-/Browser-Sperre
# direkt nach Downloadende ist meist in ein paar Sekunden vorbei.
$ArchivVersuche      = 3
$ArchivWartenSekunden = 2

# Dateinamen-Muster fuer die Sendungsnummern-CSVs (unveraendert):
#   DHL-VLS*.csv          DHL Versandlabel-System, FESTER Name (Kollision haeufig)
#   *_Sendungsnummern.csv Deutsche Post ("..._mit_Sendungsnummern.csv")
#   EXPORT_*.csv          DPD (myDPD-Export, mit Zeitstempel im Namen)
$Muster = @(
    "DHL-VLS*.csv",
    "*_Sendungsnummern.csv",
    "EXPORT_*.csv"
)

function Log($msg) {
    $zeile = "{0}  {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $msg
    try { Add-Content -Path $LogDatei -Value $zeile -Encoding UTF8 } catch {}
}

# Prüft, ob der Download fertig ist (Datei nicht mehr vom Browser gesperrt)
function Test-FileReady($Pfad) {
    try {
        $fs = [System.IO.File]::Open($Pfad, 'Open', 'Read', 'None')
        $fs.Close(); $fs.Dispose(); return $true
    } catch { return $false }
}

# Zielpfad in $Ordner, bei Namenskollision mit Zeitstempel eindeutig gemacht.
function Get-EindeutigenZielpfad($Ordner, $Name) {
    $ZielPfad = Join-Path $Ordner $Name
    if (Test-Path $ZielPfad) {
        $Basis = [System.IO.Path]::GetFileNameWithoutExtension($Name)
        $Ext   = [System.IO.Path]::GetExtension($Name)
        $Stamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $ZielPfad = Join-Path $Ordner ("{0}_{1}{2}" -f $Basis, $Stamp, $Ext)
    }
    return $ZielPfad
}

# Eindeutige Signatur einer Datei (Pfad+Groesse+Aenderungszeit). Aendert sich
# die Datei (neuer Inhalt unter gleichem Namen), aendert sich die Signatur ->
# wird automatisch wieder neu geprueft statt dauerhaft uebersprungen.
function Get-DateiSignatur($DateiInfo) {
    return "{0}|{1}|{2}" -f $DateiInfo.FullName, $DateiInfo.Length, $DateiInfo.LastWriteTimeUtc.Ticks
}

# Selbstreinigend eine Signatur-Merkliste laden: nur Eintraege behalten, deren
# Datei (erstes Feld) noch existiert. Geloeschte/verschobene Downloads
# verschwinden damit automatisch wieder aus der Liste.
function Read-SignaturSet($Pfad) {
    $set = New-Object 'System.Collections.Generic.HashSet[string]'
    if (Test-Path $Pfad) {
        Get-Content $Pfad | ForEach-Object {
            $teile = $_ -split '\|'
            if ($teile.Count -eq 3 -and (Test-Path -LiteralPath $teile[0])) {
                [void]$set.Add($_)
            }
        }
    }
    return ,$set
}

function Write-SignaturSet($Pfad, $Set) {
    try {
        $Set | Set-Content -Path $Pfad -Encoding UTF8
    } catch {
        Log "FEHLER beim Schreiben der Merkliste $($Pfad): $($_.Exception.Message)"
    }
}

# Verschieben mit ein paar Wiederholversuchen bei einer kurzen (Antivirus-/
# Browser-)Sperre. True bei Erfolg.
function Move-ItemMitRetry($QuellPfad, $ZielPfad) {
    for ($versuch = 1; $versuch -le $ArchivVersuche; $versuch++) {
        try {
            Move-Item -LiteralPath $QuellPfad -Destination $ZielPfad -Force
            return $true
        } catch {
            if ($versuch -lt $ArchivVersuche) {
                Start-Sleep -Seconds $ArchivWartenSekunden
            } else {
                Log "FEHLER beim Verschieben nach $ArchivVersuche Versuch(en) $QuellPfad -> $($ZielPfad): $($_.Exception.Message)"
            }
        }
    }
    return $false
}

# Prueft per Inhalt (label_erkennen.py), ob eine PDF ein echtes Versandlabel
# ist. Rueckgabe: $true/$false. Bei technischem Fehler (Python/Skript fehlt
# o.ae.) wird NICHT kopiert und der Fehler geloggt - lieber einmal zu wenig
# automatisch kopieren als versehentlich fremde PDFs in Paketscheine ablegen.
#
# Lauscht ueber .NET ProcessStartInfo mit CreateNoWindow=$true statt dem
# einfachen "&"-Operator: Der einfache Aufruf wuerde fuer JEDEN Python-Prozess
# ein NEUES Konsolenfenster erzeugen (kurz aufblitzend), auch wenn PowerShell
# selbst schon unsichtbar laeuft - CreateNoWindow verhindert das an der
# Wurzel, statt ein Fenster nachtraeglich zu verstecken (derselbe Unterschied
# wie beim VBS-Umweg fuer PowerShell selbst, nur eine Ebene tiefer).
function Test-IstVersandlabel($Pfad) {
    try {
        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName = $PythonExe
        $psi.Arguments = ('"{0}" "{1}"' -f $LabelErkennung, $Pfad)
        $psi.UseShellExecute = $false
        $psi.CreateNoWindow = $true
        $psi.RedirectStandardOutput = $true
        $psi.RedirectStandardError = $true

        $proc = [System.Diagnostics.Process]::Start($psi)
        $ausgabe = $proc.StandardOutput.ReadToEnd().Trim()
        $fehlerausgabe = $proc.StandardError.ReadToEnd().Trim()
        $proc.WaitForExit()
        $code = $proc.ExitCode

        if ($code -eq 0) {
            Log "Label erkannt ($Pfad): $ausgabe"
            return $true
        } elseif ($code -eq 1) {
            return $false
        } else {
            Log "Erkennung FEHLGESCHLAGEN ($Pfad): $ausgabe $fehlerausgabe"
            return $false
        }
    } catch {
        Log "FEHLER beim Aufruf der Label-Erkennung ($Pfad): $($_.Exception.Message)"
        return $false
    }
}

# Alle drei Zielordner liegen lokal auf C: - bei Bedarf einfach anlegen.
foreach ($Ordner in @($ZielAfterSell, $QuelleArchiv, $ZielWooCommerce)) {
    if (-not (Test-Path $Ordner)) {
        try { New-Item -ItemType Directory -Path $Ordner -Force | Out-Null }
        catch { Log "FEHLER beim Anlegen von $($Ordner): $($_.Exception.Message)" }
    }
}

$CsvAusstehendSet = Read-SignaturSet $CsvAusstehendDatei

# --- 1) Sendungsnummern-CSVs -> erst nach Sendungsnummern_WC KOPIEREN, --------
#        dann nach AfterSell VERSCHIEBEN --------------------------------------
foreach ($m in $Muster) {
    Get-ChildItem -Path $Quelle -Filter $m -File -ErrorAction SilentlyContinue | ForEach-Object {
        # Datei-Objekt VOR try/catch in eine eigene Variable uebernehmen: $_
        # wird INNERHALB eines catch-Blocks von PowerShell auf den Fehler-
        # datensatz umgebogen, nicht mehr die Pipeline-Datei - "$_.Name" im
        # catch waere dann leer, egal was schiefging.
        $datei = $_
        if (-not (Test-FileReady $datei.FullName)) { return }   # Download noch nicht abgeschlossen

        $signatur = Get-DateiSignatur $datei

        # a) Kopie nach Sendungsnummern_WC - nur wenn nicht schon geschehen
        #    (siehe $CsvAusstehendSet: Move nach AfterSell stand beim letzten
        #    Lauf noch aus, kopiert wurde aber bereits).
        if (-not $CsvAusstehendSet.Contains($signatur)) {
            if (-not (Test-Path $ZielWooCommerce)) {
                # Sollte nach dem Anlegen oben nicht vorkommen (Platte voll /
                # Rechteproblem). Datei bleibt komplett in Downloads liegen und
                # wird beim naechsten Lauf erneut versucht - NICHT schon nach
                # AfterSell verschieben, sonst erreicht dieser Export den WC-Sync nie.
                Log "WARNUNG: WC-Ordner fehlt: $ZielWooCommerce - $($datei.Name) bleibt vorerst in Downloads."
                return
            }
            $WcZielPfad = Get-EindeutigenZielpfad $ZielWooCommerce $datei.Name
            try {
                Copy-Item -LiteralPath $datei.FullName -Destination $WcZielPfad -Force
                [void]$CsvAusstehendSet.Add($signatur)
                Log "Sendungsnummern kopiert nach WC: $($datei.Name) -> $WcZielPfad"
            } catch {
                Log "FEHLER beim Kopieren nach WC $($datei.Name): $($_.Exception.Message) - bleibt in Downloads."
                return
            }
        }

        # b) Verschieben nach AfterSell (lokal). Die Datei liegt jetzt sicher
        #    im WC-Ordner; schlaegt der Move fehl, bleibt sie in Downloads und
        #    der Move wird naechsten Lauf erneut versucht - dank $CsvAusstehendSet
        #    OHNE erneute WC-Kopie.
        $AsZielPfad = Get-EindeutigenZielpfad $ZielAfterSell $datei.Name
        if (Move-ItemMitRetry $datei.FullName $AsZielPfad) {
            [void]$CsvAusstehendSet.Remove($signatur)
            Log "Sendungsnummern verschoben nach AfterSell: $($datei.Name) -> $AsZielPfad"
        } else {
            Log ("WARNUNG: $($datei.Name) wurde nach WC kopiert, aber der Move nach AfterSell " +
                 "schlug fehl - Datei bleibt in Downloads, Move wird naechsten Lauf erneut " +
                 "versucht (keine erneute WC-Kopie).")
        }
    }
}

# --- 2) Versandlabel-PDFs -> ins Netzwerk kopieren, Original archivieren ---
$GeprueftSet = Read-SignaturSet $GeprueftDatei

if (-not (Test-Path $ZielNetzwerk)) {
    # Netzwerkordner gerade nicht erreichbar (z.B. Freigabe kurz weg) -> Label
    # bleibt in Downloads liegen und wird beim naechsten Lauf erneut versucht.
    Log "WARNUNG: Netzwerkordner nicht erreichbar: $ZielNetzwerk - PDFs bleiben vorerst in Downloads."
} else {
    Get-ChildItem -Path $Quelle -Filter "*.pdf" -File -ErrorAction SilentlyContinue | ForEach-Object {
        # s.o.: $_ VOR try/catch sichern, sonst zeigt ein Fehlerlog im catch
        # keinen Dateinamen mehr (PowerShell bindet $_ im catch auf den
        # Fehlerdatensatz um).
        $datei = $_
        if (-not (Test-FileReady $datei.FullName)) { return }
        $signatur = Get-DateiSignatur $datei
        if ($GeprueftSet.Contains($signatur)) { return }   # schon mal geprueft+abgelehnt/kopiert, unveraendert -> ueberspringen
        if (-not (Test-IstVersandlabel $datei.FullName)) {
            [void]$GeprueftSet.Add($signatur)              # kein Label -> merken, kuenftig ueberspringen
            return
        }
        $NetzZielPfad = Get-EindeutigenZielpfad $ZielNetzwerk $datei.Name
        try {
            Copy-Item -LiteralPath $datei.FullName -Destination $NetzZielPfad -Force
        } catch {
            Log "FEHLER bei Label $($datei.Name) (Kopieren nach Paketscheine): $($_.Exception.Message)"
            return   # Original bleibt in Downloads, naechster Lauf versucht es erneut
        }
        $ArchivZielPfad = Get-EindeutigenZielpfad $QuelleArchiv $datei.Name
        if (Move-ItemMitRetry $datei.FullName $ArchivZielPfad) {
            Log "Label kopiert + Original archiviert: $($datei.Name) -> $NetzZielPfad"
        } else {
            # Kopiert ist es bereits (s.o.) - ein erneutes Kopieren beim naechsten
            # Lauf waere in jedem Fall falsch, deshalb IN JEDEM FALL merken, auch
            # wenn das Archivieren des Originals nicht geklappt hat. Die Datei
            # bleibt sichtbar in Downloads liegen, statt bei jedem Lauf erneut
            # (mit neuem Zeitstempel-Namen) nach Paketscheine dupliziert zu werden.
            [void]$GeprueftSet.Add($signatur)
            $warnText = "WARNUNG: $($datei.Name) wurde nach Paketscheine kopiert, konnte aber " +
                        "NICHT aus Downloads archiviert werden - bitte manuell pruefen/entfernen: $($datei.FullName)"
            Log $warnText
        }
    }
}

# Merklisten aktualisiert zurueckschreiben (nur noch existierende Dateien, s.o.)
Write-SignaturSet $CsvAusstehendDatei $CsvAusstehendSet
Write-SignaturSet $GeprueftDatei $GeprueftSet
