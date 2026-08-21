# -*- coding: utf-8 -*-
"""
Versand-Cockpit - Scan & Druck Versandscheine (Netzwerk-/Dauerbetrieb)
======================================================================
Grafisches Cockpit (Tkinter). Scannt den Barcode (Rechnungsnummer) und druckt
den/die passenden Versandschein(e) aus einem (Netzwerk-)Ordner.
Cockpit zeigt:
  - Digitaluhr mit Datum.
  - Pro Versender eine Kachel: Countdown bis zur Abhol-Deadline, Kreisdiagramm
    (gedruckt/offen) und die Liste der noch offenen Auftraege (Rechnungsnr +
    Empfaenger).
  - EIGENE Kachel fuer die Deutsche Post (GoGreen): zusaetzlich Anzeige der
    Labels, die sich KEINER Rechnung zuordnen liessen (Adresse pruefen).
  - Gruener Hintergrund, sobald alles gedruckt ist; Status oben rechts.
  - Meldungszeile unten mit dem Ergebnis des letzten Scans.
Hintergrund (unveraendert bewaehrt): Live-Ueberwachung des Ordners,
Doppeldruck-Schutz, Tagesstatistik, Auto-Archiv fertiger Dateien.
Versender-Erkennung am INHALT der Seite (nicht am Dateinamen):
  - DPD  : Anker "Referenz 1"           Deadline 14:00
  - DHL  : Anker "Kostenstelle:"        Deadline 16:00
  - Deutsche Post (GoGreen): KEINE Rechnungsnummer auf dem Label -> Zuordnung
           ueber die Lieferadresse (PLZ + Hausnr + Name) via post_zuordnung.csv.
Voraussetzungen (einmalig):
  - Python 3 (Tkinter ist dabei), dann: py -m pip install pdfplumber pypdf
  - SumatraPDF (kostenlos) fuer lautloses Drucken.
Starten:  py scan_druck.py
"""
import os
import re
import io
import sys
import glob
import time
import shutil
import hashlib
import tempfile
import threading
import subprocess
from datetime import datetime
from collections import defaultdict, Counter
import pdfplumber
from pypdf import PdfReader, PdfWriter
# reportlab nur fuer das Sammeldruck-Deckblatt. Fehlt es, wird der Sammeldruck
# ohne Deckblatt ausgefuehrt (mit Hinweis). Installation: py -m pip install reportlab
try:
    from reportlab.pdfgen import canvas as rl_canvas
    from reportlab.lib.units import mm as rl_mm
    _HAS_REPORTLAB = True
except Exception:
    _HAS_REPORTLAB = False
# ============================ KONFIGURATION ============================
VERSION = "2026-08-21c"          # im Fenstertitel sichtbar -> Deployment pruefbar
# Versionsschema: JJJJ-MM-TT + Kleinbuchstabe je Aenderung am selben Tag (erste
# Aenderung des Tages = a, dann b, c ...; ein neuer Tag beginnt wieder bei a).
# 2026-08-21c: Nachbesserung zu 2026-08-21a/b nach erneuter Pruefung:
#   1) Sammeldruck-Teilerfolg wurde als voller Erfolg gemeldet (gruen, bei per
#      EAN bestaetigtem Sammeldruck sogar als grosses gruenes "GEDRUCKT"-
#      Overlay), wenn einer von mehreren Versendern erfolgreich war, ein
#      anderer aber nicht (drucke_sammel()-modus wurde nur bei KOMPLETTEM
#      Fehlschlag "fehler", nicht schon bei teilweisem). Jetzt "fehler"
#      sobald irgendein Versender fehlschlug - die EAN-Bestaetigungsstelle war
#      bereits korrekt auf "st == 'gedruckt'" verdrahtet und zeigt dadurch
#      automatisch keine gruene Bestaetigung mehr bei einem Teilerfolg.
#   2) DRUCK_TIMEOUT_SEKUNDEN von 20 auf 8 gesenkt: _sende_an_drucker() (neu
#      seit 2026-08-21a) laeuft SYNCHRON im Tkinter-Hauptthread (direkt im
#      <Return>-Handler des Scan-Eingabefelds) - jede Sekunde Wartezeit hier
#      friert das GESAMTE Cockpit ein (Uhr, Countdown, naechster Scan an JEDER
#      Station), nicht nur die aktuelle Bestellung. Ein Sammeldruck mit
#      mehreren ausgefallenen Versendern wartet zusaetzlich mehrfach
#      hintereinander - bei 20s theoretisch bis zu 40-60s Totalausfall der
#      Oberflaeche. 8s begrenzt das Risiko spuerbar, ohne einen echten
#      Spooler-Haenger vorschnell abzubrechen. Fuer wirklich nicht-
#      blockierendes Drucken waere ein Hintergrund-Thread noetig (groessere
#      Architektur-Aenderung, bewusst noch nicht umgesetzt).
#   3) drucke() und drucke_sammel() fangen jetzt OSError beim Lesen einer
#      Label-Datei ab (_reader()). Grund: die Race-Condition-Absicherung
#      VERSCHWINDEN_TOLERANZ (2026-08-21a) laesst eine geraede verschwindende
#      Datei bewusst noch bis zu ~2 Polls im Index stehen - wird sie GENAU in
#      diesem kurzen Fenster gescannt, schlug das Oeffnen der (bereits
#      verschobenen) Datei bisher mit einem rohen "[Errno 2] ..."-Text bis in
#      die Meldezeile durch. Jetzt eine klare "Drucken fehlgeschlagen"-
#      Meldung statt der Python-Fehlermeldung; bei drucke_sammel() betrifft
#      das nur den einzelnen betroffenen Versender, die anderen im selben
#      Sammeldruck laufen unbeeinflusst weiter.
#   4) Tagesmengen-Anzeige: Emoji "📦" (ausserhalb der Basic Multilingual
#      Plane, anders als die uebrigen Symbole ★/⚠/✓ im Cockpit) durch "Σ"
#      ersetzt - vermeidet ein moegliches Darstellungsproblem auf aelteren
#      Tcl/Tk-Versionen.
# 2026-08-21b: Tagesmenge in der Kopfzeile ergaenzt - Summe der heute
#   gedruckten Labels ueber ALLE Versender (DHL + DPD + Deutsche Post
#   zusammen) auf einen Blick, mit Aufschluesselung je Versender in Klammern.
#   Steht ueber dem Tagesrekord, wird bei jedem aktualisiere_anzeige()-Lauf
#   aus z.gedruckt_heute_pv neu berechnet (bleibt korrekt auch nach dem
#   Archivieren einzelner Dateien, da dieser Zaehler dafuer extra separat
#   gefuehrt wird).
# 2026-08-21a: Projektweite Fehlerpruefung, zwei Kernprobleme behoben:
#   1) Druck-Erfolg wird jetzt tatsaechlich geprueft. Bisher wurde SumatraPDF
#      per subprocess.Popen() nur GESTARTET, nie auf Erfolg geprueft - ein
#      falscher/offline Druckername oder ein SumatraPDF-Fehler blieb unbemerkt,
#      die Bestellung galt im Cockpit trotzdem als "gedruckt" (drucke()/
#      _drucke_pdf() ueber die neue _sende_an_drucker(), die per
#      subprocess.run()+Timeout wartet und den Exit-Code prueft). Alle vier
#      Druck-Aufrufstellen (Normalfall, Mehrfach-Scan-Sicherung, EAN-
#      Verifikation, Sammeldruck, Nachdruck) markieren eine Bestellung/einen
#      Sammelcode jetzt nur noch bei tatsaechlichem Druckerfolg als gedruckt;
#      bei Fehlschlag erscheint eine rote "FEHLER beim Drucken"-Meldung und
#      NICHTS wird geloggt/gezaehlt. Ein Papierstau NACH erfolgreicher
#      Uebergabe an den Windows-Spooler bleibt weiterhin unerkennbar.
#   2) Race Condition in der Duplikat-Erkennung vom 2026-07-30a behoben: fuer
#      NEU erscheinende Dateien gab es schon eine Settle-Wartezeit gegen
#      Fehlalarme, fuer VERSCHWINDENDE Dateien nicht - ein einzelner
#      transienter Aussetzer des Netzwerk-Listings (UNC-Haenger, AV-Scan)
#      loeschte sofort den Inhalts-Hash einer Datei; taucht im selben Moment
#      eine zweite, inhaltsgleiche Kopie auf, fand der Duplikat-Abgleich dann
#      keinen Treffer mehr -> moeglicher Doppeldruck trotz der Absicherung.
#      Eine Datei muss jetzt VERSCHWINDEN_TOLERANZ (2) aufeinanderfolgende
#      Polls fehlen, bevor sie als wirklich entfernt gilt (Zustand.fehlend_seit).
# 2026-07-30a: Zwei Absicherungen gegen Fehlbedienung beim Datei-Handling im
#   Netzwerkordner (ausgeloest durch einen doppelten DPD-Download UND eine
#   Nachsendung mit wiederverwendeter Rechnungsnummer):
#   1) Inhalts-Duplikat-Erkennung: landet dieselbe Label-Seite (identischer
#      Text) durch einen zweiten Download zweimal im Ordner (i.d.R. unter
#      leicht anderem Dateinamen, z.B. mit "(1)"), wird die zweite Kopie NICHT
#      mehr zusaetzlich in den Index aufgenommen -> kein doppelter Druck beim
#      naechsten Scan mehr. Echte Mehrpaket-Bestellungen (mehrere UNTERSCHIED-
#      LICHE Labels mit derselben Rechnungsnummer) sind davon nicht betroffen,
#      da deren Seiteninhalt sich unterscheidet.
#   2) Taucht eine Rechnungsnummer erneut in einer NEUEN Datei auf, obwohl sie
#      laut Druck-Log schon gedruckt wurde (z.B. Nachsendung mit derselben
#      Rechnungsnummer wie die urspruengliche Bestellung), wird die Datei NICHT
#      mehr automatisch stillschweigend ins Archiv verschoben. Stattdessen
#      erscheint eine rote Warnzeile im Cockpit, damit der Fall bewusst
#      geprueft wird (z.B. gezielt erneut scannen -> Nachfrage "erneut
#      drucken?"), statt dass das Label unbemerkt im Archiv verschwindet.
# 2026-07-20a: Packplatz "Verpackungshalle" aktiviert - Scanner-Praefix "L" leitet
#   DHL/DPD-Labels auf den Hallendrucker "VerpackungshalleZ". Vorher war der Eintrag
#   auskommentiert, weshalb Hallen-Scans auf den Verpackungsraum-Drucker fielen.
# 2026-07-17a: Post-Zuordnung nach einem ZWEITEN Rechnungslauf greift jetzt ohne
#   Neustart - aktualisiere_index erkennt die geaenderte post_zuordnung.csv und
#   parst bereits eingelesene PDFs neu (zuvor unzuordenbare Post-Sendungen werden
#   dadurch nachtraeglich zugeordnet).
#   Ausserdem: Nachdruck eines bereits archivierten Belegs friert die Oberflaeche
#   nicht mehr ein. Statt das ganze Archiv zu durchsuchen, merkt sich das Cockpit
#   beim Archivieren die Zuordnung (Archiv-Gedaechtnis) -> Nachdruck sofort. Liegt
#   das Label nur noch im Archiv einer frueheren Sitzung, kommt eine klare Meldung
#   statt eines Hangs.
# 2026-07-01: Deutsche-Post-INTERNETMARKE/Briefmarke wird jetzt erkannt (Text ist
#   zeichenweise gespiegelt und traegt kein "Deutsche Post"-Wort, nur das Logo) ->
#   Erkennung ueber die gespiegelte IM-Frankaturzeile. Auslands-PLZ (Oesterreich,
#   4-stellig) wird beim Post-Adressabgleich mit ausgewertet.
#   Ausserdem: Auto-Ausloesung bei Tipp-Pause standardmaessig AUS
#   (AUTO_ENTER_BEI_TIPPPAUSE=False) - gedruckt wird per Enter (Scanner sendet es).
# (Netzwerk-)Ordner mit den Label-PDFs. UNC erlaubt, z.B. r"\\server\Versand".
LABEL_ORDNER = r"\\DESKTOP-N2H75H\Netzwerk\Paketscheine"
DATEI_MUSTER = "*.pdf"            # alle PDFs pruefen; Versender wird am Inhalt erkannt
# Adress-Zuordnung fuer Deutsche-Post-Labels (wird von packliste.py erzeugt):
POST_ZUORDNUNG_CSV = r"\\DESKTOP-N2H75H\Netzwerk\Paketscheine\post_zuordnung.csv"
# Sammeldruck-Zuordnung (wird von packliste.py erzeugt): Sammelcode -> Bestellungen
SAMMEL_ZUORDNUNG_CSV = r"\\DESKTOP-N2H75H\Netzwerk\Paketscheine\sammel_zuordnung.csv"
# Mengen-Zuordnung (wird von packliste.py erzeugt): Rechnungsnummer -> Stueckzahl.
# Grundlage fuer die Mehrfach-Scan-Sicherung (s.u.).
MENGEN_ZUORDNUNG_CSV = r"\\DESKTOP-N2H75H\Netzwerk\Paketscheine\mengen_zuordnung.csv"
# --- Mehrfach-Scan-Sicherung -------------------------------------------------
# Gegen "ein Artikel verpackt, zweiten vergessen": Besteht eine Bestellung aus
# mehreren Stueck (Summe aller Mengen, z.B. 1x Schlauch + 2x Kartusche = 3),
# muss der Barcode entsprechend oft gescannt werden, BEVOR gedruckt wird. Die
# noetige Stueckzahl kommt aus MENGEN_ZUORDNUNG_CSV; fehlt sie (CSV nicht da /
# Rechnung nicht enthalten), gilt 1 -> es wird wie bisher sofort gedruckt.
MEHRFACH_SCAN_AKTIV = True        # False = alte Sofort-Druck-Logik
MAX_SCANS = 5                     # nie mehr als so viele Scans verlangen (Deckel)
# Mindestabstand (ms) zwischen zwei ZAEHL-Scans derselben Bestellung. Verhindert,
# dass jemand denselben Barcode in Sekundenbruchteilen mehrfach "durchrattert",
# um die Sicherung zu umgehen. 0 = aus. Empfehlung bei Missbrauch: 500-800.
MIN_SCAN_ABSTAND_MS = 0
# Strenger Modus: Wenn zu einer Bestellung KEINE Mengeninfo vorliegt (Rechnung
# fehlt in mengen_zuordnung.csv / CSV nicht da), dann NICHT drucken, sondern
# warnen. Verhindert, dass mehrteilige Bestellungen unbemerkt durchrutschen, weil
# die Mengenliste fehlt. Achtung: blockiert dann auch echte Einzelartikel ohne
# Eintrag -> erst einschalten, wenn der Datenfluss (packliste.py schreibt die CSV
# in den Netzwerkordner) sicher steht. False = alter, nie blockierender Fallback.
MENGE_PFLICHT = False
# --- EAN-Verifikationsscan ---------------------------------------------------
# Echte Inhaltspruefung: Bei Artikeln, die in der Pickliste eine EAN tragen, muss
# nach dem Rechnungs-Scan die EAN des PHYSISCHEN Artikels gescannt werden. Erst
# wenn alle erwarteten EANs (in der noetigen Anzahl) stimmen, wird das Label
# gedruckt. Falsche EAN -> rote Warnung, kein Druck. Bei EAN-Artikeln ERSETZT der
# EAN-Scan das Mengen-Mehrfachscannen. Bestellungen OHNE EAN-Position laufen
# unveraendert ueber die Stueckzahl-Sicherung (mengen_zuordnung.csv).
# Datenquelle: ean_zuordnung.csv (von packliste.py erzeugt).
EAN_VERIFIKATION_AKTIV = True     # False = EAN-Schritt komplett aus (alte Logik)
EAN_ZUORDNUNG_CSV = r"\\DESKTOP-N2H75H\Netzwerk\Paketscheine\ean_zuordnung.csv"
# Versender-Profile. Reihenfolge zaehlt: erstes passendes "erkennung" gewinnt.
#   name             Anzeigename
#   erkennung        Regex; passt sie auf den Seitentext -> Seite gehoert hierher
#   anker            Regex, Gruppe 1 = Rechnungsnummer (entfaellt bei post_adresse)
#   empfaenger_anker Regex, Gruppe 1 = Empfaengername (leer -> Zeile davor)
#   deadline         "HH:MM" Abhol-Deadline fuer den Countdown (oder None)
#   farbe            Akzentfarbe im Cockpit (Hex)
#   post_adresse     True -> Zuordnung ueber die Lieferadresse statt Rechnungsnummer
#   drucker          Druckername fuer diesen Versender (fehlt/None = Standarddrucker;
#                    erfordert SumatraPDF; exakter Windows-Name, siehe DRUCKER unten)
VERSENDER_PROFILE = [
    {
        "name": "DPD",
        "erkennung": r"DPD|PARCELLetter|Parcel",
        # Wert steht in der ZEILE NACH "Referenz 1" -> erste 5+-stellige Zahl danach
        "anker": r"Referenz\s*1\b\D*?(\d{5,})",
        # "Empfänger" (mit Umlaut) steht allein, Name in der naechsten Zeile
        "empfaenger_anker": r"(?:Empf[\wä]*nger|Name)\s*[:\s]+(.+)",
        "deadline": "14:00",
        "farbe": "#C62828",
        "drucker": "VerpackungsraumZ",   # DHL/DPD-Standard (wenn kein Scanner-Praefix)
    },
    {
        "name": "DHL",
        "erkennung": r"DHL\s*PAKET|PAKET INTERNATIONAL|Kostenstelle:",
        "anker": r"Kostenstelle:\s*(\d+)",
        "empfaenger_anker": r"(?:An|To)\s*[:\s]+(.*)",
        "deadline": "16:00",
        "farbe": "#B8860B",
        "drucker": "VerpackungsraumZ",   # DHL/DPD-Standard (wenn kein Scanner-Praefix)
    },
    {
        # Deutsche Post (GoGreen / Internetmarke): KEINE Rechnungsnummer auf dem
        # Label. Erkennung am (gespiegelten) Inhalt, Zuordnung ueber die Adresse.
        "name": "Deutsche Post",
        "erkennung": r"NEERGOG|GOGREEN|tsoP ehcstueD|Deutsche Post",
        "post_adresse": True,
        "deadline": "16:00",      # Abhol-Deadline der Post (= DHL)
        "farbe": "#FFCC00",       # Post-Gelb
        "drucker": "VerpackungsraumD",   # einziger Post-faehiger Drucker - IMMER hierhin
    },
]
SUMATRA_EXE = r"C:\Users\Gasecenter\AppData\Local\SumatraPDF\SumatraPDF.exe"
# Druckeinstellungen fuer SumatraPDF: "noscale" druckt 1:1 (fuer Versandlabel
# korrekt und i.d.R. schneller, weil keine Skalierung gerechnet wird). Leer = aus.
SUMATRA_PRINT_SETTINGS = "noscale"
DRUCKER = None                   # None = Standarddrucker
MIN_STELLEN = 5                  # so viele Ziffern muss ein Scan mindestens haben
# --- Packplatz-Routing (mehrere Scanner) --------------------------------------
# Jeder Scanner bekommt per Konfigurations-Barcode (siehe Scanner-Handbuch,
# "Preamble"/"Prefix") ein festes PRAEFIX, das er vor JEDEN Scan setzt. Daran
# erkennt das Programm, an welchem Packplatz gescannt wurde, und waehlt den
# passenden Drucker. Der Versender bestimmt weiterhin die Label-ART
# (Post -> DYMO), der Packplatz das GERAET.
#
# WICHTIG zum Praefix: nur Buchstaben verwenden (auf jedem Tastaturlayout gleich)
# und NICHT mit "S" beginnen (sonst Verwechslung mit "SAM-"-Codes). Ziffern sind
# tabu, weil Rechnungsnummern reine Ziffern sind. Der Verpackungsraum-Scanner
# sendet z.B. "H" (-> H1689095).
PACKPLATZ_PRAEFIX_AKTIV = True    # False = altes Verhalten (nur Versender entscheidet)
PACKPLATZ_PROFILE = {
    # Praefix : {name, drucker: {Versendername -> exakter Windows-Druckername}}
    #
    # WICHTIG: Deutsche Post hier BEWUSST NICHT eintragen. Post-Labels gehen
    # immer auf den Drucker aus dem Versender-Profil (VerpackungsraumD) -
    # unabhaengig davon, welcher Scanner gescannt hat. Nur DHL/DPD entscheidet
    # der Scanner.
    "H": {
        "name": "Verpackungsraum",
        "drucker": {
            "DHL": "VerpackungsraumZ",
            "DPD": "VerpackungsraumZ",
        },
    },
    # Verpackungshalle: der Hallen-Scanner sendet das Praefix "L" -> DHL/DPD-Labels
    # gehen auf den Hallendrucker "VerpackungshalleZ" (exakter Windows-Name).
    # Post bleibt ausgenommen (nicht eingetragen) -> immer VerpackungsraumD.
    "L": {
        "name": "Verpackungshalle",
        "drucker": {
            "DHL": "VerpackungshalleZ",
            "DPD": "VerpackungshalleZ",
        },
    },
}
# Scan OHNE erkanntes Praefix (z.B. von der normalen Tastatur) -> None bedeutet:
# Drucker allein nach Versender (Profilfeld "drucker") = Verpackungsraum-Drucker.
PACKPLATZ_STANDARD = None
DRUCK_LOG = r"\\DESKTOP-N2H75H\Netzwerk\Paketscheine\gedruckt.log"
STATISTIK_DATEI = r"\\DESKTOP-N2H75H\Netzwerk\Paketscheine\statistik.csv"
# Firmenlogo fuer die Cockpit-Kopfzeile. PNG oder GIF laufen ohne Zusatzpaket
# (Tkinter kann beides); mit installiertem Pillow zusaetzlich JPG und glattere
# Skalierung. Datei einfach als "logo.png" neben scan_druck.py (bzw. neben die
# .exe) legen. Fehlt sie, bleibt die Kopfzeile ohne Logo - kein Fehler.
_BASIS_ORDNER = (os.path.dirname(sys.executable) if getattr(sys, "frozen", False)
                 else os.path.dirname(os.path.abspath(__file__)))
LOGO_DATEI = os.path.join(_BASIS_ORDNER, "logo.png")
LOGO_HOEHE = 56                  # Zielhoehe des Logos in Pixeln
ARCHIV_UNTERORDNER = "Archiv"
POLL_SEKUNDEN = 3                # Ordner-Pruefintervall
SETTLE_SEKUNDEN = 3              # Datei erst lesen, wenn so lange unveraendert
SCAN_PAUSE_MS = 120              # Tipp-Pause, ab der ein Scan als beendet gilt
# Auto-Ausloesung bei Tipp-Pause (Druck ohne Enter). Standard AUS: die Scanner
# senden nach dem Scan ohnehin ein Enter, und bei HAENDISCHER Eingabe reicht die
# kurze Pause nicht -> es wuerde zu frueh ausgeloest. True = altes Verhalten.
AUTO_ENTER_BEI_TIPPPAUSE = False
DEADLINE_WARN_MIN = 30           # Countdown wird rot, wenn weniger Minuten offen
TESTMODUS = False                # True = Treffer als PDF in "test_ausgabe" statt Druck
# Max. Wartezeit auf SumatraPDF (Sekunden), bevor ein Druckauftrag als
# fehlgeschlagen gilt. WICHTIG: _sende_an_drucker() laeuft synchron im
# Tkinter-Hauptthread (direkt im <Return>-Handler des Scan-Eingabefelds) -
# jede Sekunde hier friert das GESAMTE Cockpit fuer ALLE Packplaetze ein
# (Uhr, Countdown, naechster Scan), nicht nur die aktuelle Station. Ein
# Sammeldruck mit mehreren ausgefallenen Versendern wartet zusaetzlich MEHRFACH
# hintereinander. Urspruenglich auf 20s gesetzt ("grosszuegig fuer einen
# langsamen Spooler-Start") - das war zu hoch angesetzt: SumatraPDF gibt bei
# einem ungueltigen Druckernamen oder fehlendem Sumatra sofort (Bruchteile
# einer Sekunde) einen Fehler zurueck, ein echter Haenger (z.B. eine
# Treiber-Dialogbox) ist der seltene Rand- statt der Normalfall. Deutlich
# kuerzer gesetzt, um das Einfrier-Risiko zu begrenzen; fuer wirklich
# nicht-blockierendes Drucken waere ein Hintergrund-Thread noetig (groessere
# Umbau-Aenderung, absichtlich noch nicht umgesetzt).
DRUCK_TIMEOUT_SEKUNDEN = 8
# --- Duplikat-Erkennung (Inhalt) ---------------------------------------------
# Mindestlaenge des extrahierten Seitentextes, ab der ein Inhalts-Hash als
# verlaesslich gilt (verhindert Fehlalarm bei fast leeren Seiten). Kuerzer ->
# keine Duplikatpruefung fuer diese Seite (wird immer normal aufgenommen).
DUPLIKAT_MIN_TEXTLAENGE = 20
# ======================================================================
# =====================================================================
#  BACKEND-LOGIK (ohne Tkinter, eigenstaendig testbar)
# =====================================================================
class Zustand:
    def __init__(self):
        self.lock = threading.RLock()
        self.index = defaultdict(list)   # nr -> [(pfad, seite, versender)]
        self.empfaenger = {}             # nr -> Empfaengername
        self.datei_nrs = {}              # pfad -> set(nr)
        self.datei_sig = {}              # pfad -> (mtime, size)
        self.gedruckt = {}               # nr -> zeitstempel
        # offen_pv: noch nicht gedruckte Labels (Seiten) je Versender, aus dem Index.
        # gedruckt_heute_pv: heute bereits gedruckte Labels je Versender - bleibt
        # erhalten, auch wenn die Datei nach dem Druck archiviert wird.
        self.offen_pv = Counter()
        self.gedruckt_heute_pv = Counter()
        self.heute_datum = None          # fuer Mitternachts-Reset des Tageszaehlers
        self.reader_cache = {}           # pfad -> PdfReader
        self.post_unzuordenbar = {}      # pfad -> [(name, plz), ...]
        self.archiv_index = {}           # nr -> [(archiv_pfad, seite, versender)] (Nachdruck)
        # Mehrfach-Scan-Sicherung: Rechnungsnummer -> bisher gescannte Anzahl
        # (Sitzungs-Speicher; nach Neustart leer, nach Druck/Reset entfernt).
        self.scan_fortschritt = {}
        self.letzter_scan_ts = {}        # Rechnungsnummer -> monotone Zeit des letzten Zaehl-Scans
        # EAN-Verifikation: Station(-Praefix) -> offener Verifikations-Kontext einer
        # Bestellung {rn, treffer, versender, station, soll, ist, rest:{ean:offen},
        # namen:{ean:bez}, eans_all:set}. Sitzungs-Speicher; nach Druck/Reset leer.
        self.aktive_verifikation = {}
        # Sammeldruck-EAN-Bestaetigung: Station(-Praefix) -> {code, ean, bez, art}.
        # Wird gesetzt, sobald ein Sammelcode gescannt wird, dessen Artikel eine
        # EAN traegt - gedruckt wird der Sammelblock erst nach passendem EAN-Scan.
        # Sitzungs-Speicher; nach Druck/Reset/Abbruch leer.
        self.aktive_sammel_verifikation = {}
        # --- Absicherung gegen Datei-Handling-Fehler im Netzwerkordner --------
        # seite_hashes: nr -> {inhalts_hash: pfad}. Traegt jede aktuell im Index
        # gefuehrte Seite mit ihrem Text-Hash ein -> Grundlage der Duplikat-
        # Erkennung (zweiter Download derselben Sendung).
        self.seite_hashes = defaultdict(dict)
        # duplikate: nr -> [(neue_datei, urspruengliche_datei), ...]. Protokoll
        # erkannter Inhalts-Duplikate dieser Sitzung (fuer die Warnanzeige).
        self.duplikate = {}
        # verdaechtig: nr -> [pfad, ...]. Dateien, deren Rechnungsnummer beim
        # Einlesen BEREITS im Druck-Log stand (z.B. Nachsendung mit wieder-
        # verwendeter Rechnungsnummer). Diese Dateien werden NICHT automatisch
        # archiviert, bis der Fall manuell geklaert ist.
        self.verdaechtig = defaultdict(list)
        # fehlend_seit: pfad -> Anzahl aufeinanderfolgender Polls, in denen eine
        # bisher bekannte Datei im Ordner-Listing NICHT mehr auftauchte. Fuer
        # NEU erscheinende Dateien gibt es die SETTLE_SEKUNDEN-Wartezeit gegen
        # Fehlalarme; fuer VERSCHWINDENDE Dateien gab es das bisher nicht -> ein
        # einzelner transienter Aussetzer des Netzwerk-Listings (UNC-Haenger,
        # AV-Scan) loeschte sofort den Inhalts-Hash der Datei (_entferne_datei).
        # Taucht im selben Moment ein zweiter, inhaltsgleicher Download auf,
        # fand der Duplikat-Abgleich dann keinen Hash-Treffer mehr -> doppelter
        # Druck trotz der Absicherung vom 2026-07-30a. Jetzt muss eine Datei
        # VERSCHWINDEN_TOLERANZ aufeinanderfolgende Polls fehlen, bevor sie
        # wirklich als entfernt gilt.
        self.fehlend_seit = {}
VERSCHWINDEN_TOLERANZ = 2         # so viele Polls darf eine Datei aus dem
                                   # Listing fehlen, bevor sie als geloescht/
                                   # verschoben gilt (siehe Zustand.fehlend_seit)
# ----------------- Deutsche-Post-Zuordnung (Adresse -> Rechnungsnummer) --------
POST_LOOKUP = {}            # "PLZ|Hausnr" -> [(Rechnungsnummer, Name), ...]
POST_BY_PLZ = {}            # "PLZ" -> [(Rechnungsnummer, Name, Hausnr), ...]
_POST_CSV_SIG = None
def _norm(s):
    s = (s or "").lower()
    s = s.replace("ä", "ae").replace("ö", "oe").replace("ü", "ue").replace("ß", "ss")
    return re.sub(r"[^a-z0-9]+", "", s)
def _name_tokens(s):
    s = (s or "").lower()
    s = s.replace("ä", "ae").replace("ö", "oe").replace("ü", "ue").replace("ß", "ss")
    return set(t for t in re.findall(r"[a-z0-9]+", s) if len(t) >= 2)
def _namen_passen(a, b):
    """Namensvergleich unabhaengig von der Wortreihenfolge
    ('hauer thomas' == 'thomas hauer')."""
    ta, tb = _name_tokens(a), _name_tokens(b)
    if not ta or not tb:
        return False
    return ta <= tb or tb <= ta or len(ta & tb) >= 2
def _hausnr(street):
    m = re.search(r"(\d+\s*[a-zA-Z]?)\s*$", (street or "").strip())
    return re.sub(r"\s+", "", m.group(1)) if m else ""
def _plz(zeile):
    m = re.search(r"\b(\d{5})\b", zeile or "")
    return m.group(1) if m else ""
# --- Auslands-Zustellung (v.a. Oesterreich) -----------------------------------
# Deutsche PLZ = 5-stellig, oesterreichische = 4-stellig. Damit AT-Briefmarken
# ueber PLZ+Hausnr / PLZ+Name genauso zugeordnet werden koennen, wird das Land
# aus dem Adressblock bestimmt und die PLZ passend dazu gelesen.
def _land_aus_zeilen(zeilen):
    """'AT' wenn eine Zeile auf Oesterreich deutet, sonst 'DE'.
    Erkannt werden 'ÖSTERREICH'/'AUSTRIA' sowie das ISO-Praefix vor der PLZ
    ('AT-8020', 'A-8020')."""
    txt = " ".join(zeilen or []).upper().replace("Ö", "OE")
    if re.search(r"OESTERREICH|AUSTRIA|\b(?:AT|A)-\s?\d{4}\b", txt):
        return "AT"
    return "DE"
def _plz_aus_zeilen(zeilen, land=None):
    """PLZ aus den Adresszeilen, laenderabhaengig.
    DE: 5-stellig (letzte passende Zeile). AT: 4-stellig, als 'AT-8020'/'A-8020'
    oder '8020 Ort' (Ort dahinter) -> keine Verwechslung mit Hausnummern."""
    zeilen = zeilen or []
    if land is None:
        land = _land_aus_zeilen(zeilen)
    if land == "AT":
        # AT-Format 'AT-8020'/'A-8020' zuerst, dann '8020 Ort'
        for z in zeilen:
            m = re.search(r"\b(?:AT|A)-?\s?(\d{4})\b", z)
            if m:
                return m.group(1)
        for z in zeilen:
            m = re.search(r"\b(\d{4})\b\s+\S", z)   # 4-stellig + Ort dahinter
            if m:
                return m.group(1)
        return ""
    # DE (Default): 5-stellig
    for z in reversed(zeilen):
        m = re.search(r"\b(\d{5})\b", z)
        if m:
            return m.group(1)
    return ""
def _ph_key(plz, hausnr):
    return f"{plz}|{hausnr}"
def lade_post_zuordnung(pfad):
    """post_zuordnung.csv -> POST_LOOKUP {"PLZ|Hausnr": [(nr, name), ...]} und
    POST_BY_PLZ {"PLZ": [(nr, name, hausnr), ...]} fuer den Namens-Fallback."""
    global POST_LOOKUP, POST_BY_PLZ, _POST_CSV_SIG
    try:
        st = os.stat(pfad)
    except OSError:
        changed = _POST_CSV_SIG is not None
        POST_LOOKUP, POST_BY_PLZ, _POST_CSV_SIG = {}, {}, None
        return changed              # True, wenn vorher eine Zuordnung geladen war
    sig = (st.st_mtime, st.st_size)
    if sig == _POST_CSV_SIG:
        return False                # unveraendert
    lookup = defaultdict(list)
    by_plz = defaultdict(list)
    try:
        with open(pfad, encoding="utf-8-sig") as f:
            f.readline()  # Kopfzeile
            for zeile in f:
                t = zeile.rstrip("\n").split(";")
                if len(t) < 5 or not t[0]:
                    continue
                nr, name, strasse, hausnr, plz = t[0], t[1], t[2], t[3], t[4]
                if not hausnr:
                    hausnr = _hausnr(strasse)
                lookup[_ph_key(plz, hausnr)].append((nr, name))
                by_plz[plz].append((nr, name, hausnr))
    except Exception:
        return False                # Sig NICHT uebernehmen -> naechster Lauf erneut
    POST_LOOKUP, POST_BY_PLZ = dict(lookup), dict(by_plz)
    _POST_CSV_SIG = sig
    return True                     # Zuordnung wurde neu eingelesen
# ----------------- Sammeldruck-Zuordnung (Sammelcode -> Bestellungen) ---------
SAMMEL_LOOKUP = {}          # "SAM-n" -> {"art","bez","menge","rnr":[...]}
_SAMMEL_CSV_SIG = None
def lade_sammel_zuordnung(pfad):
    """sammel_zuordnung.csv -> {Code: {art, bez, menge, rnr:[...], ean}}.
    Format: Code;Artikelnr;Bezeichnung;Menge;Rechnungsnummern(kommasepariert)
    [;EAN]. Die EAN-Spalte ist optional (aeltere packliste.py-Version ohne
    EAN-Spalte) - fehlt sie oder ist sie leer, wird fuer diesen Sammelcode
    KEINE EAN-Bestaetigung verlangt (alte Logik: direkter Druck)."""
    global SAMMEL_LOOKUP, _SAMMEL_CSV_SIG
    try:
        st = os.stat(pfad)
    except OSError:
        SAMMEL_LOOKUP, _SAMMEL_CSV_SIG = {}, None
        return
    sig = (st.st_mtime, st.st_size)
    if sig == _SAMMEL_CSV_SIG:
        return
    lookup = {}
    try:
        with open(pfad, encoding="utf-8-sig") as f:
            f.readline()  # Kopfzeile
            for zeile in f:
                t = zeile.rstrip("\n").split(";")
                if len(t) < 5 or not t[0]:
                    continue
                code, art, bez, menge, rnrs = t[0], t[1], t[2], t[3], t[4]
                ean = re.sub(r"\D", "", t[5]) if len(t) > 5 else ""
                rnr = [r.strip() for r in rnrs.split(",") if r.strip()]
                lookup[code.strip().upper()] = {
                    "art": art, "bez": bez, "menge": menge, "rnr": rnr, "ean": ean}
    except Exception:
        return
    SAMMEL_LOOKUP, _SAMMEL_CSV_SIG = lookup, sig
# ----------------- Mengen-Zuordnung (Rechnungsnummer -> Stueckzahl) -----------
MENGEN_LOOKUP = {}          # Rechnungsnummer(nur Ziffern) -> Stueckzahl (int >=1)
_MENGEN_CSV_SIG = None
def lade_mengen_zuordnung(pfad):
    """mengen_zuordnung.csv -> {Rechnungsnummer(nur Ziffern): Stueckzahl}.
    Format: Rechnungsnummer;Stueckzahl. Schluessel werden auf reine Ziffern
    normiert (wie der gescannte Wert), damit der Abgleich passt."""
    global MENGEN_LOOKUP, _MENGEN_CSV_SIG
    try:
        st = os.stat(pfad)
    except OSError:
        MENGEN_LOOKUP, _MENGEN_CSV_SIG = {}, None
        return
    sig = (st.st_mtime, st.st_size)
    if sig == _MENGEN_CSV_SIG:
        return
    lookup = {}
    try:
        with open(pfad, encoding="utf-8-sig") as f:
            f.readline()  # Kopfzeile
            for zeile in f:
                t = zeile.rstrip("\n").split(";")
                if len(t) < 2 or not t[0]:
                    continue
                key = re.sub(r"\D", "", t[0])
                if not key:
                    continue
                try:
                    menge = int(round(float((t[1] or "").replace(",", "."))))
                except ValueError:
                    continue
                if menge >= 1:
                    lookup[key] = menge
    except Exception:
        return
    MENGEN_LOOKUP, _MENGEN_CSV_SIG = lookup, sig
def artikelmenge(rn):
    """Echte Stueckzahl der Bestellung laut Pickliste (>=1). Fehlt die Angabe
    (CSV nicht vorhanden / Rechnung nicht enthalten), wird 1 angenommen -> die
    alte Sofort-Druck-Logik bleibt unveraendert."""
    return MENGEN_LOOKUP.get(rn, 1)
def soll_scans(rn):
    """Wie oft muss fuer rn gescannt werden, bevor gedruckt wird?
    = min(Stueckzahl, MAX_SCANS), mindestens 1. Bei abgeschalteter Sicherung 1."""
    if not MEHRFACH_SCAN_AKTIV:
        return 1
    return max(1, min(artikelmenge(rn), MAX_SCANS))
def menge_bekannt(rn):
    """True, wenn fuer rn eine Stueckzahl in der Mengenliste steht."""
    return rn in MENGEN_LOOKUP
def mengen_status():
    """Zustand der Mengenliste fuer die Anzeige/Diagnose.
    Rueckgabe: (ok: bool, anzahl: int, text: str). ok=False heisst: die
    Mehrfach-Scan-Sicherung kann gerade NICHT greifen (CSV fehlt/leer)."""
    if not MEHRFACH_SCAN_AKTIV:
        return (True, len(MENGEN_LOOKUP), "Mehrartikel-Sicherung AUS")
    da = os.path.exists(MENGEN_ZUORDNUNG_CSV)
    n = len(MENGEN_LOOKUP)
    if not da:
        return (False, 0, "Mengenliste fehlt - Sicherung inaktiv!")
    if n == 0:
        return (False, 0, "Mengenliste leer - Sicherung inaktiv!")
    return (True, n, f"Mengenliste: {n} Bestellungen")
# ----------------- EAN-Zuordnung (Rechnungsnummer -> erwartete EANs) ----------
# EAN_LOOKUP[rn] = {"eans": {ean: anzahl}, "namen": {ean: bez}, "soll": int}
# Nur Bestellungen mit mindestens einer EAN-Position stehen hier.
EAN_LOOKUP = {}
_EAN_CSV_SIG = None
def lade_ean_zuordnung(pfad):
    """ean_zuordnung.csv -> EAN_LOOKUP. Format je Zeile:
        Rechnungsnummer;EAN;Bezeichnung;Anzahl
    Mehrere Zeilen je Rechnung sind moeglich (mehrere Positionen). Gleiche EAN in
    einer Rechnung wird in der Anzahl aufsummiert. Schluessel (Rechnungsnr und EAN)
    werden auf reine Ziffern normiert - genau wie der gescannte Wert."""
    global EAN_LOOKUP, _EAN_CSV_SIG
    try:
        st = os.stat(pfad)
    except OSError:
        EAN_LOOKUP, _EAN_CSV_SIG = {}, None
        return
    sig = (st.st_mtime, st.st_size)
    if sig == _EAN_CSV_SIG:
        return
    lookup = {}
    try:
        with open(pfad, encoding="utf-8-sig") as f:
            f.readline()  # Kopfzeile
            for zeile in f:
                t = zeile.rstrip("\n").split(";")
                if len(t) < 4 or not t[0]:
                    continue
                rn = re.sub(r"\D", "", t[0])
                ean = re.sub(r"\D", "", t[1])
                if not rn or not ean:
                    continue
                bez = (t[2] or "").strip().strip('"')
                try:
                    anzahl = int(round(float((t[3] or "").replace(",", "."))))
                except ValueError:
                    anzahl = 1
                anzahl = max(1, anzahl)
                eintrag = lookup.setdefault(rn, {"eans": {}, "namen": {}, "soll": 0})
                eintrag["eans"][ean] = eintrag["eans"].get(ean, 0) + anzahl
                eintrag["namen"].setdefault(ean, bez)
                eintrag["soll"] += anzahl
    except Exception:
        return
    EAN_LOOKUP, _EAN_CSV_SIG = lookup, sig
def hat_ean_pflicht(rn):
    """True, wenn fuer rn mindestens eine EAN-Position vorliegt (und EAN-Scan an)."""
    if not EAN_VERIFIKATION_AKTIV:
        return False
    eintrag = EAN_LOOKUP.get(rn)
    return bool(eintrag and eintrag["soll"] > 0)
def ean_status():
    """(ok, anzahl, text) der EAN-Liste fuer die Statusanzeige."""
    if not EAN_VERIFIKATION_AKTIV:
        return (True, len(EAN_LOOKUP), "EAN-Verifikation AUS")
    da = os.path.exists(EAN_ZUORDNUNG_CSV)
    n = len(EAN_LOOKUP)
    if not da:
        return (True, 0, "EAN-Liste fehlt - keine EAN-Pruefung")
    return (True, n, f"EAN-Liste: {n} Bestellung(en) mit EAN")
def post_rechnungsnr(name, zeilen, plz, belegte=None):
    """(nr|None, status) fuer ein Post-Label.
    `zeilen`: alle Adresszeilen (Name..PLZ). Robuste Zuordnung:
      1. PLZ + Hausnummer (Hausnummer aus JEDER Adresszeile probiert),
      2. falls eindeutig -> nehmen; bei mehreren ueber den Namen eingrenzen
         (Reihenfolge egal); bei echten Doppeladressen der Reihe nach die
         naechste freie Rechnungsnummer,
      3. Fallback: PLZ + Name (wenn die Hausnummer nicht passt)."""
    belegt = belegte if belegte is not None else set()
    # Hausnummer-Kandidaten aus allen Adresszeilen + leere Hausnummer
    hnrs = []
    for z in zeilen or []:
        h = _hausnr(z)
        if h and h not in hnrs:
            hnrs.append(h)
    hnrs.append("")
    # 1) Kandidaten ueber PLZ+Hausnr sammeln (dedupe, Reihenfolge erhalten)
    seen, kand = set(), []
    for h in hnrs:
        for c in POST_LOOKUP.get(_ph_key(plz, h), []):
            if c[0] not in seen:
                seen.add(c[0])
                kand.append(c)
    # 3) Fallback ueber PLZ + Name, falls Hausnummer nichts (Passendes) liefert
    if not kand or not any(_namen_passen(name, c[1]) for c in kand):
        plz_kand = [(nr, nm) for nr, nm, _h in POST_BY_PLZ.get(plz, [])]
        namens = [c for c in plz_kand if _namen_passen(name, c[1])]
        if namens:
            kand = namens if not kand else kand + [c for c in namens
                                                    if c[0] not in seen]
    if not kand:
        return None, "kein_treffer"
    # Bei mehreren ueber den Namen eingrenzen (reihenfolgeunabhaengig)
    pool = kand
    if len(kand) > 1:
        namens = [c for c in kand if _namen_passen(name, c[1])]
        if namens:
            pool = namens
    # der Reihe nach die naechste noch freie Rechnungsnummer
    for nr, _nm in pool:
        if nr not in belegt:
            return nr, "ok"
    return None, "alle_vergeben"
def _post_adresse_aus_seite(seite):
    """(Name, Adresszeilen, PLZ) aus einem gedrehten Deutsche-Post-Label.
    Zeichen gespiegelt -> je Wort umkehren; Spalten (=Zeilen) nach x sortieren."""
    rcpt = [w for w in seite.extract_words(use_text_flow=False) if w["x0"] >= 85]
    if not rcpt:
        return None
    cols = {}
    for w in rcpt:
        cols.setdefault(round(w["x0"] / 5) * 5, []).append(w)
    zeilen = []
    for x in sorted(cols):
        toks = sorted(cols[x], key=lambda w: -w["top"])
        zeilen.append(" ".join(t["text"][::-1] for t in toks).strip())
    zeilen = [z for z in zeilen if z]
    if not zeilen:
        return None
    name = zeilen[0]
    plz = _plz_aus_zeilen(zeilen)      # laenderabhaengig: DE 5-stellig, AT 4-stellig
    return name, zeilen, plz
# --------------------------- Datei-Erkennung & Parsen -------------------------
def _ist_internetmarke(text):
    """Deutsche-Post-Internetmarke / Briefmarke.
    Diese Labels tragen KEIN 'Deutsche Post'-Wort (nur das Logo als Grafik) und
    ihr Text ist zeichenweise gespiegelt. Erkennung ueber die gespiegelte
    IM-Frankaturzeile ('IM <TT.MM.JJ> <Betrag>'): erst jedes Wort zurueckdrehen,
    dann nach dem 'IM'-Marker plus einem Datum suchen."""
    rev = "\n".join(" ".join(w[::-1] for w in ln.split())
                    for ln in (text or "").splitlines())
    return bool(re.search(r"\bIM\b", rev)) and \
           bool(re.search(r"\d{2}\.\d{2}\.\d{2}", rev))
def erkenne_profil(text):
    """Erstes Profil, dessen 'erkennung' auf den Seitentext passt (oder None).
    Zusaetzlich: Deutsche-Post-Internetmarke ueber die gespiegelte IM-Zeile."""
    for prof in VERSENDER_PROFILE:
        if re.search(prof["erkennung"], text):
            return prof
    # Internetmarke/Briefmarke: kein Text-Token vorhanden -> ueber die gespiegelte
    # IM-Zeile erkennen und dem Deutsche-Post-Profil (Adressabgleich) zuordnen.
    # Bewusst NACH den regulaeren Profilen, damit DHL/DPD immer Vorrang haben.
    if _ist_internetmarke(text):
        for prof in VERSENDER_PROFILE:
            if prof.get("post_adresse"):
                return prof
    return None
def extrahiere_empfaenger(text, prof):
    anker = prof.get("empfaenger_anker")
    if not anker:
        return ""
    zeilen = text.splitlines()
    for i, ln in enumerate(zeilen):
        m = re.search(anker, ln)
        if m:
            name = (m.group(1) if m.groups() else "").strip()
            if not name and i > 0:
                name = zeilen[i - 1].strip()
            return name[:40]
    return ""
def parse_datei(pfad, belegte=None):
    """[(rechnungsnr, seiten_index, versender, empfaenger, inhalts_hash)] oder
    None bei Fehler. Erkennt den Versender je Seite am Inhalt; Post wird ueber
    die Adresse zugeordnet. `belegte`: bereits vergebene Rechnungsnummern
    (eindeutige Post-Zuordnung bei mehreren identischen Adressen).
    `inhalts_hash`: md5 des extrahierten Seitentextes (oder None bei sehr
    kurzem Text) - Grundlage der Duplikat-Erkennung in aktualisiere_index."""
    out = []
    unzuordenbar = []
    belegte = set(belegte) if belegte else set()
    try:
        with pdfplumber.open(pfad) as pdf:
            for i, seite in enumerate(pdf.pages):
                t = seite.extract_text() or ""
                prof = erkenne_profil(t)
                if not prof:
                    continue
                phash = (hashlib.md5(t.strip().encode("utf-8")).hexdigest()
                         if len(t.strip()) >= DUPLIKAT_MIN_TEXTLAENGE else None)
                if prof.get("post_adresse"):
                    adr = _post_adresse_aus_seite(seite)
                    if not adr:
                        continue
                    name, zeilen, plz = adr
                    nr, status = post_rechnungsnr(name, zeilen, plz, belegte)
                    if nr:
                        out.append((nr, i, prof["name"], name, phash))
                        belegte.add(nr)
                    else:
                        unzuordenbar.append((name, plz))
                else:
                    m = re.search(prof["anker"], t)
                    if not m:
                        continue
                    out.append((m.group(1), i, prof["name"],
                                extrahiere_empfaenger(t, prof), phash))
    except Exception:
        return None
    POST_UNZUORDENBAR_TMP[pfad] = unzuordenbar
    return out
# parse_datei puffert die unzuordenbar-Liste je Datei zwischen; nur der
# Live-Index (aktualisiere_index) uebernimmt sie.
POST_UNZUORDENBAR_TMP = {}
def liste_pdfs(ordner):
    out = []
    for pfad in glob.glob(os.path.join(ordner, DATEI_MUSTER)):
        if os.path.isfile(pfad):
            out.append(pfad)
    return out
def _entferne_datei(z, pfad):
    nrs = z.datei_nrs.pop(pfad, set())
    z.datei_sig.pop(pfad, None)
    z.reader_cache.pop(pfad, None)
    z.post_unzuordenbar.pop(pfad, None)
    z.fehlend_seit.pop(pfad, None)
    for nr in nrs:
        if nr in z.index:
            z.index[nr] = [t for t in z.index[nr] if t[0] != pfad]
            if not z.index[nr]:
                del z.index[nr]
                z.empfaenger.pop(nr, None)
        # Duplikat-/Verdaechtig-Markierungen dieser Datei mit aufraeumen, damit
        # sie beim naechsten Auftauchen einer wirklich neuen Datei nicht
        # faelschlich als "schon gesehen" gelten.
        if nr in z.seite_hashes:
            z.seite_hashes[nr] = {h: p for h, p in z.seite_hashes[nr].items()
                                   if p != pfad}
            if not z.seite_hashes[nr]:
                del z.seite_hashes[nr]
        if nr in z.verdaechtig:
            z.verdaechtig[nr] = [p for p in z.verdaechtig[nr] if p != pfad]
            if not z.verdaechtig[nr]:
                del z.verdaechtig[nr]
def _recompute(z):
    """Zaehlt die noch OFFENEN Labels (Seiten) je Versender aus dem Index.
    Die gedruckten Labels werden separat in z.gedruckt_heute_pv gefuehrt, damit
    die Anzeige nach dem Archivieren der Dateien nicht auf 0 zurueckspringt."""
    offen = Counter()
    for nr, treffer in z.index.items():
        if nr in z.gedruckt:
            continue
        for _, _, v in treffer:
            offen[v] += 1
    z.offen_pv = offen
def archiviere(z, pfad, ordner):
    ziel_ordner = os.path.join(ordner, ARCHIV_UNTERORDNER)
    try:
        os.makedirs(ziel_ordner, exist_ok=True)
        name = os.path.basename(pfad)
        ziel = os.path.join(ziel_ordner, name)
        if os.path.exists(ziel):
            stamm, ext = os.path.splitext(name)
            ziel = os.path.join(ziel_ordner, f"{stamm}_{datetime.now():%Y%m%d_%H%M%S}{ext}")
        shutil.move(pfad, ziel)
    except Exception:
        return False
    # Zuordnungen dieser Datei ins Archiv-Gedaechtnis uebernehmen (mit dem NEUEN
    # Archivpfad), damit ein spaeterer Nachdruck sie sofort findet, ohne das
    # ganze Archiv mit pdfplumber zu durchsuchen (das fror die GUI ein).
    for nr in z.datei_nrs.get(pfad, set()):
        eintr = [(ziel, seite, vers) for (p, seite, vers) in z.index.get(nr, [])
                 if p == pfad]
        if eintr:
            z.archiv_index[nr] = eintr
    _entferne_datei(z, pfad)
    _recompute(z)
    return True
def aktualisiere_index(z, ordner):
    """Synchronisiert den Index mit dem Ordnerinhalt. True bei Aenderung."""
    with z.lock:
        post_geaendert = lade_post_zuordnung(POST_ZUORDNUNG_CSV)
        lade_sammel_zuordnung(SAMMEL_ZUORDNUNG_CSV)
        lade_mengen_zuordnung(MENGEN_ZUORDNUNG_CSV)
        lade_ean_zuordnung(EAN_ZUORDNUNG_CSV)
        if post_geaendert:
            # Die Post-Zuordnung hat sich geaendert (typisch: zweiter
            # Rechnungslauf). Bereits eingelesene PDFs neu bewerten, damit zuvor
            # unzuordenbare Post-Sendungen jetzt zugeordnet werden - bisher
            # geschah das erst nach einem Neustart.
            z.datei_sig.clear()
        aktuelle = set(liste_pdfs(ordner))
        geaendert = False
        for pfad in list(z.datei_nrs):
            if pfad not in aktuelle:
                # Nicht sofort entfernen - ein einzelner Poll, in dem eine
                # Datei transient im Netzwerk-Listing fehlt (UNC-Haenger,
                # AV-Scan), darf nicht sofort den Inhalts-Hash loeschen (siehe
                # Zustand.fehlend_seit). Erst nach VERSCHWINDEN_TOLERANZ
                # aufeinanderfolgenden Fehlversuchen gilt die Datei als
                # wirklich geloescht/verschoben.
                bisher = z.fehlend_seit.get(pfad, 0) + 1
                if bisher < VERSCHWINDEN_TOLERANZ:
                    z.fehlend_seit[pfad] = bisher
                    continue
                _entferne_datei(z, pfad)
                geaendert = True
            else:
                z.fehlend_seit.pop(pfad, None)
        for pfad in aktuelle:
            try:
                st = os.stat(pfad)
            except OSError:
                continue
            sig = (st.st_mtime, st.st_size)
            if z.datei_sig.get(pfad) == sig:
                continue
            if time.time() - st.st_mtime < SETTLE_SEKUNDEN:
                continue
            # bereits vergebene Rechnungsnummern (ohne die eigenen alten dieser
            # Datei) -> eindeutige Post-Zuordnung bei mehreren gleichen Adressen
            eigene_alt = z.datei_nrs.get(pfad, set())
            belegte = set(z.index.keys()) - eigene_alt
            eintraege = parse_datei(pfad, belegte)
            if eintraege is None:
                continue
            if pfad in z.datei_nrs:
                _entferne_datei(z, pfad)
            nrs = set()
            for nr, seite, versender, emp, phash in eintraege:
                # --- Inhalts-Duplikat: dieselbe Seite (identischer Text) steckt
                #     schon unter einer ANDEREN Datei im Index -> zweiter Download
                #     derselben Sendung. NICHT zusaetzlich aufnehmen, sonst wird
                #     beim naechsten Scan doppelt gedruckt. Echte Mehrpaket-
                #     Bestellungen sind nicht betroffen, da sich deren Seiten im
                #     Text unterscheiden (andere Sendungs-/Paketnummer o.ae.).
                if phash is not None:
                    quelle = z.seite_hashes[nr].get(phash)
                    if quelle is not None and quelle != pfad:
                        paar = (os.path.basename(pfad), os.path.basename(quelle))
                        liste = z.duplikate.setdefault(nr, [])
                        if paar not in liste:
                            liste.append(paar)
                        continue
                    z.seite_hashes[nr][phash] = pfad
                # --- Diese Rechnungsnummer stand BEREITS im Druck-Log, BEVOR
                #     diese Datei auftauchte (z.B. Nachsendung mit wieder-
                #     verwendeter Rechnungsnummer, oder versehentlich erneut
                #     abgelegte Datei). Nicht stillschweigend durchlaufen lassen
                #     -> markieren, damit sie unten NICHT automatisch archiviert
                #     wird, sondern als Warnung sichtbar bleibt.
                if nr in z.gedruckt and pfad not in z.verdaechtig[nr]:
                    z.verdaechtig[nr].append(pfad)
                z.index[nr].append((pfad, seite, versender))
                if emp:
                    z.empfaenger[nr] = emp
                nrs.add(nr)
            z.datei_nrs[pfad] = nrs
            z.datei_sig[pfad] = sig
            z.post_unzuordenbar[pfad] = POST_UNZUORDENBAR_TMP.get(pfad, [])
            geaendert = True
        if geaendert:
            _recompute(z)
        for pfad in list(z.datei_nrs):
            nrs = z.datei_nrs[pfad]
            verdaechtig_hier = any(pfad in z.verdaechtig.get(nr, []) for nr in nrs)
            if (nrs and all(nr in z.gedruckt for nr in nrs)
                    and not z.post_unzuordenbar.get(pfad) and not verdaechtig_hier):
                if archiviere(z, pfad, ordner):
                    geaendert = True
        return geaendert
def lade_gedruckt(pfad):
    gedruckt = {}
    if os.path.exists(pfad):
        with open(pfad, encoding="utf-8") as f:
            for zeile in f:
                teile = zeile.strip().split("\t")
                if teile and teile[0]:
                    gedruckt[teile[0]] = teile[1] if len(teile) > 1 else "?"
    return gedruckt
def lade_gedruckt_heute(pfad):
    """Summe der HEUTE gedruckten Labels (Pakete/Seiten) je Versender aus der
    Statistik-CSV. Quelle fuer die 'gedruckt'-Anzeige im Cockpit, damit sie nach
    dem Archivieren erhalten bleibt."""
    heute = datetime.now().strftime("%Y-%m-%d")
    pv = Counter()
    if os.path.exists(pfad):
        try:
            with open(pfad, encoding="utf-8") as f:
                next(f, None)  # Kopfzeile
                for zeile in f:
                    t = zeile.strip().split(";")
                    if len(t) >= 4 and t[0] == heute:
                        try:
                            pakete = int(t[3])
                        except ValueError:
                            pakete = 1
                        pv[t[1]] += pakete
        except OSError:
            pass
    return pv
def merke_gedruckt(pfad, nr):
    with open(pfad, "a", encoding="utf-8") as f:
        f.write(f"{nr}\t{datetime.now():%Y-%m-%d %H:%M:%S}\n")
def log_statistik(pfad, versender, nr, pakete):
    neu = not os.path.exists(pfad)
    with open(pfad, "a", encoding="utf-8") as f:
        if neu:
            f.write("Datum;Versender;Rechnungsnummer;Pakete;Uhrzeit\n")
        jetzt = datetime.now()
        f.write(f"{jetzt:%Y-%m-%d};{versender};{nr};{pakete};{jetzt:%H:%M:%S}\n")
def highscore_tag(pfad):
    """(datum_iso, anzahl) des Tages mit den meisten versendeten Bestellungen.
    Zaehlt EINDEUTIGE Rechnungsnummern je Datum aus der Statistik-CSV (nicht die
    Pakete). Der laufende Tag zaehlt mit, damit ein neuer Rekord sofort erscheint.
    Gibt (None, 0) zurueck, wenn keine Daten vorliegen."""
    if not os.path.exists(pfad):
        return None, 0
    tage = defaultdict(set)                 # datum -> {rechnungsnummern}
    try:
        with open(pfad, encoding="utf-8") as f:
            next(f, None)                   # Kopfzeile ueberspringen
            for zeile in f:
                t = zeile.rstrip("\n").split(";")
                if len(t) >= 3 and t[0] and t[2]:
                    tage[t[0]].add(t[2])
    except OSError:
        return None, 0
    if not tage:
        return None, 0
    best = max(tage, key=lambda d: len(tage[d]))
    return best, len(tage[best])
def highscore_text(pfad):
    """Zweizeiliger Text fuer die Tagesrekord-Anzeige in der Kopfzeile."""
    d, n = highscore_tag(pfad)
    if not d or n <= 0:
        return "★ TAGESREKORD\n– noch keiner –"
    try:
        dd = datetime.strptime(d, "%Y-%m-%d").strftime("%d.%m.%Y")
    except ValueError:
        dd = d
    wort = "Bestellung" if n == 1 else "Bestellungen"
    return f"★ TAGESREKORD\n{n} {wort} · {dd}"
def lade_logo(pfad, zielhoehe):
    """Logo laden und auf ~zielhoehe px Hoehe bringen -> PhotoImage oder None.
    Bevorzugt Pillow (glatte Skalierung, auch JPG). Ohne Pillow nur PNG/GIF ueber
    Tkinter mit ganzzahligem Verkleinern. Fehlt/defekt -> None (kein Fehler)."""
    import tkinter as tk
    if not pfad or not os.path.exists(pfad):
        return None
    try:                                    # Weg 1: Pillow (falls installiert)
        from PIL import Image, ImageTk
        try:
            _resample = Image.Resampling.LANCZOS      # Pillow >= 9.1
        except AttributeError:
            _resample = Image.LANCZOS                 # aeltere Pillow
        im = Image.open(pfad)
        breite, hoehe = im.size
        if hoehe > zielhoehe:
            faktor = zielhoehe / float(hoehe)
            im = im.resize((max(1, int(breite * faktor)), int(zielhoehe)),
                           _resample)
        return ImageTk.PhotoImage(im)
    except Exception:
        pass
    try:                                    # Weg 2: Tkinter pur (PNG/GIF)
        img = tk.PhotoImage(file=pfad)
        if img.height() > zielhoehe:
            img = img.subsample(max(1, round(img.height() / zielhoehe)))
        return img
    except Exception:
        return None
def offene_auftraege(z):
    """{versender: [(nr, empfaenger), ...]} fuer alle noch nicht gedruckten."""
    offen = defaultdict(list)
    with z.lock:
        for nr, treffer in z.index.items():
            if nr in z.gedruckt:
                continue
            versender = treffer[0][2] if treffer else "?"
            offen[versender].append((nr, z.empfaenger.get(nr, "")))
    for v in offen:
        offen[v].sort()
    return offen
def post_unzuordenbar_liste(z):
    with z.lock:
        return [x for liste in z.post_unzuordenbar.values() for x in liste]
def warnungen_liste(z):
    """Kurztexte zu erkannten Inhalts-Duplikaten und verdaechtigen (bereits
    gedruckten) Rechnungsnummern, fuer die Warnzeile im Cockpit."""
    with z.lock:
        dup = {nr: list(v) for nr, v in z.duplikate.items()}
        verd = {nr: list(v) for nr, v in z.verdaechtig.items()}
    texte = []
    for nr, paare in dup.items():
        neu, alt = paare[-1]
        texte.append(f"Duplikat erkannt: {nr} - {neu} = bereits {alt} (nicht gedruckt)")
    for nr, dateien in verd.items():
        texte.append(f"Rechnungsnr. {nr} war schon gedruckt, taucht erneut auf "
                     f"({os.path.basename(dateien[-1])}) - bitte pruefen, nicht "
                     f"automatisch archiviert")
    return texte
def _reader(z, pfad):
    with z.lock:
        r = z.reader_cache.get(pfad)
    if r is None:
        with open(pfad, "rb") as fh:
            r = PdfReader(io.BytesIO(fh.read()))
        with z.lock:
            z.reader_cache[pfad] = r
    return r
def _sende_an_drucker(pfad, drucker):
    """Schickt eine fertige PDF-Datei an den Drucker. Gibt True bei Erfolg
    zurueck, False bei jedem erkennbaren Fehlschlag (Aufrufer darf dann NICHT
    als gedruckt gelten). Frueher wurde SumatraPDF per Popen() nur GESTARTET,
    ohne je das Ergebnis zu pruefen - ein falscher/offline Druckername, ein
    fehlendes SumatraPDF oder ein sofortiger SumatraPDF-Fehler blieb dadurch
    unbemerkt: die Bestellung galt im Cockpit trotzdem als "gedruckt", obwohl
    physisch nichts herauskam. Jetzt wird auf das Prozessende gewartet
    ("-exit-when-done") und der Exit-Code geprueft; ein Timeout (haengender
    Prozess) gilt ebenfalls als Fehlschlag statt das Cockpit zu blockieren.
    Ein Papierstau NACH erfolgreicher Uebergabe an den Spooler bleibt weiterhin
    unerkennbar (das kann nur der Windows-Druckertreiber selbst feststellen) -
    das deckt diese Absicherung bewusst nicht ab."""
    if SUMATRA_EXE and os.path.exists(SUMATRA_EXE):
        cmd = [SUMATRA_EXE, "-silent", "-exit-when-done"]
        cmd += (["-print-to", drucker] if drucker else ["-print-to-default"])
        if SUMATRA_PRINT_SETTINGS:
            cmd += ["-print-settings", SUMATRA_PRINT_SETTINGS]
        cmd.append(pfad)
        try:
            ergebnis = subprocess.run(cmd, timeout=DRUCK_TIMEOUT_SEKUNDEN)
        except subprocess.TimeoutExpired:
            print(f"Drucken fehlgeschlagen (SumatraPDF haengt/Timeout): {pfad}")
            return False
        except Exception as e:
            print(f"Drucken fehlgeschlagen (SumatraPDF-Start): {e}")
            return False
        if ergebnis.returncode != 0:
            print(f"Drucken fehlgeschlagen (SumatraPDF Exit-Code "
                  f"{ergebnis.returncode}, Drucker: {drucker or 'Standard'}): {pfad}")
            return False
        return True
    if drucker:
        print(f"Hinweis: ohne SumatraPDF nur Standarddrucker moeglich "
              f"(gewuenscht war: {drucker}).")
    try:
        os.startfile(pfad, "print")  # type: ignore[attr-defined]
        return True
    except Exception as e:
        print(f"Drucken fehlgeschlagen: {e}")
        return False
def drucke(z, treffer, nr, station=None):
    """Gibt True bei erfolgreichem Druck zurueck, sonst False."""
    writer = PdfWriter()
    try:
        for pdf_pfad, idx, _ in treffer:
            writer.add_page(_reader(z, pdf_pfad).pages[idx])
    except OSError as e:
        # Die Datei stand noch im Index, ist aber gerade in dem kurzen Fenster
        # verschwunden, das VERSCHWINDEN_TOLERANZ absichtlich offen laesst
        # (Absicherung gegen die Race Condition vom 2026-08-21b) - z.B. wird
        # sie in genau diesem Moment archiviert/verschoben. Ohne dieses
        # Abfangen wuerde ein roher "[Errno 2] ..."-Text bis in die Melde-
        # zeile durchschlagen (generischer except-Handler in scan_ausfuehren).
        print(f"Drucken fehlgeschlagen (Label-Datei kurzzeitig nicht "
              f"verfuegbar, {nr}): {e}")
        return False
    if TESTMODUS:
        os.makedirs("test_ausgabe", exist_ok=True)
        with open(os.path.join("test_ausgabe", f"Rechnung_{nr}.pdf"), "wb") as f:
            writer.write(f)
        return True
    tmp = os.path.join(tempfile.gettempdir(), f"versandschein_{nr}.pdf")
    with open(tmp, "wb") as f:
        writer.write(f)
    # Drucker richtet sich nach dem Versender der Sendung (Post -> eigener Drucker)
    versender = treffer[0][2] if treffer else None
    drucker = drucker_von(versender, station)
    return _sende_an_drucker(tmp, drucker)
def _drucke_pdf(writer, drucker, name):
    """Einen fertigen PdfWriter an einen (optionalen) Drucker schicken.
    Gibt True bei erfolgreichem Druck zurueck, sonst False."""
    if TESTMODUS:
        os.makedirs("test_ausgabe", exist_ok=True)
        with open(os.path.join("test_ausgabe", f"{name}.pdf"), "wb") as f:
            writer.write(f)
        return True
    tmp = os.path.join(tempfile.gettempdir(), f"{name}.pdf")
    with open(tmp, "wb") as f:
        writer.write(f)
    return _sende_an_drucker(tmp, drucker)
def deckblatt_seite(breite, hoehe, code, art, bez, anzahl, versender):
    """Deckblatt in Label-Groesse als PdfReader-Seite (oder None ohne reportlab)."""
    if not _HAS_REPORTLAB:
        return None
    buf = io.BytesIO()
    c = rl_canvas.Canvas(buf, pagesize=(breite, hoehe))
    m = max(6.0, min(breite, hoehe) * 0.05)
    c.setLineWidth(2)
    c.rect(m, m, breite - 2 * m, hoehe - 2 * m)
    cx = breite / 2.0
    innen = breite - 2 * m - 10
    y = [hoehe - m - hoehe * 0.11]
    def zeile(text, groesse, bold=True, luecke=1.55):
        font = "Helvetica-Bold" if bold else "Helvetica"
        size = max(6.0, groesse)
        while size > 6 and c.stringWidth(text, font, size) > innen:
            size -= 0.5
        c.setFont(font, size)
        c.drawCentredString(cx, y[0], text)
        y[0] -= size * luecke
    zeile("SAMMELDRUCK", hoehe * 0.05, bold=False)
    y[0] -= hoehe * 0.015
    zeile(f"Artikel {art}", hoehe * 0.085)
    zeile(bez, hoehe * 0.038, bold=False)
    y[0] -= hoehe * 0.015
    zeile(f"{anzahl} x 1 Stueck", hoehe * 0.06)
    c.setFont("Helvetica", max(7.0, hoehe * 0.028))
    c.drawCentredString(cx, m + hoehe * 0.04, f"{versender}  -  {code}")
    c.showPage()
    c.save()
    buf.seek(0)
    return PdfReader(buf).pages[0]
def drucke_sammel(z, code, ordner, station=None):
    """Alle Labels eines Sammelcodes auf einmal drucken: je Versender ein
    Deckblatt (Artikel/Menge/Bezeichnung) + dessen Labels. (status, meldung)."""
    g = SAMMEL_LOOKUP.get(code)
    if not g:
        return ("unbekannt",
                f"Sammelcode {code} nicht gefunden (sammel_zuordnung.csv aktuell?)")
    gefunden, fehlend, schon = [], [], []
    for rnr in g["rnr"]:
        if rnr in z.gedruckt:
            schon.append(rnr)
            continue
        tr = finde_treffer(z, rnr, ordner)
        if tr:
            gefunden.append((rnr, tr))
        else:
            fehlend.append(rnr)
    if not gefunden:
        teile = []
        if schon:
            teile.append(f"{len(schon)} bereits gedruckt")
        if fehlend:
            teile.append(f"{len(fehlend)} ohne Label")
        zusatz = (" (" + ", ".join(teile) + ")") if teile else ""
        return ("unbekannt", f"Sammelcode {code}: nichts zu drucken{zusatz}")
    # Seiten je Versender (= je Drucker) buendeln, Deckblatt voranstellen
    nach_versender = defaultdict(list)   # versender -> [(pfad, idx)]
    for _rnr, tr in gefunden:
        for (pfad, idx, vers) in tr:
            nach_versender[vers].append((pfad, idx))
    # Erfolg je Versender einzeln merken: schlaegt der Druckauftrag fuer einen
    # Versender fehl (falscher/offline Drucker), duerfen dessen Rechnungen NICHT
    # als gedruckt gelten - frueher wurden hier ALLE `gefunden`-Rechnungen
    # unbedingt markiert, unabhaengig vom tatsaechlichen Druckergebnis.
    versender_ok = {}
    for versender, seiten in nach_versender.items():
        writer = PdfWriter()
        try:
            erste = _reader(z, seiten[0][0]).pages[seiten[0][1]]
            try:
                bseite, hseite = float(erste.mediabox.width), float(erste.mediabox.height)
            except Exception:
                bseite, hseite = 283.0, 425.0
            deck = deckblatt_seite(bseite, hseite, code, g["art"], g["bez"],
                                   len(gefunden), versender)
            if deck is not None:
                writer.add_page(deck)
            for (pfad, idx) in seiten:
                writer.add_page(_reader(z, pfad).pages[idx])
        except OSError as e:
            # s.o. (drucke()): Label-Datei kurzzeitig verschwunden (Race-
            # Condition-Toleranzfenster). Nur DIESEN Versender als fehlgeschlagen
            # werten, die anderen im selben Sammeldruck laufen unbeeinflusst
            # weiter - statt die gesamte Funktion mit einer rohen Exception
            # abzubrechen.
            print(f"Drucken fehlgeschlagen (Label-Datei kurzzeitig nicht "
                  f"verfuegbar, Sammelcode {code}/{versender}): {e}")
            versender_ok[versender] = False
            continue
        versender_ok[versender] = _drucke_pdf(
            writer, drucker_von(versender, station),
            f"sammel_{code}_{versender}".replace(" ", "_"))
    # je Rechnung als gedruckt markieren + Statistik - nur wenn der Druckauftrag
    # fuer ihren Versender tatsaechlich erfolgreich war.
    jetzt = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    ok_liste = [(rnr, tr) for rnr, tr in gefunden if versender_ok.get(tr[0][2])]
    druckfehler = sorted({vers for vers, ok in versender_ok.items() if not ok})
    for rnr, tr in ok_liste:
        with z.lock:
            z.gedruckt[rnr] = jetzt
            z.gedruckt_heute_pv[tr[0][2]] += len(tr)
        merke_gedruckt(DRUCK_LOG, rnr)
        log_statistik(STATISTIK_DATEI, tr[0][2], rnr, len(tr))
    with z.lock:
        _recompute(z)
    teile = [f"{len(ok_liste)} Label gedruckt"]
    if druckfehler:
        teile.append(f"FEHLER beim Drucken ({', '.join(druckfehler)}) - "
                      f"NICHT als gedruckt markiert")
    if schon:
        teile.append(f"{len(schon)} schon vorher")
    if fehlend:
        teile.append(f"{len(fehlend)} OHNE LABEL")
    deck_hint = "" if _HAS_REPORTLAB else " - OHNE Deckblatt (reportlab fehlt)"
    # "fehler" schon bei TEILWEISEM Fehlschlag (nicht erst wenn ALLES
    # fehlschlug): scheitert z.B. nur der DPD-Drucker, waren die anderen
    # Versender trotzdem erfolgreich (ok_liste nicht leer) - vorher blieb
    # modus dann "gedruckt" (gruen), obwohl druckfehler nicht leer war, und
    # die Fehlermeldung ging in der gruenen Erfolgsmeldung unter (bei einer
    # per EAN bestaetigten Sammeldruck-Bestellung sogar als grosses gruenes
    # "GEDRUCKT"-Overlay, siehe Aufrufer).
    modus = "fehler" if druckfehler else "gedruckt"
    return (modus, f"Sammeldruck {code} ({g['art']}): "
            + ", ".join(teile) + deck_hint)
def finde_treffer(z, rn, ordner):
    """Treffer fuer eine Rechnungsnummer: erst der Live-Index, dann das
    Archiv-Gedaechtnis (in dieser Sitzung archivierte Dateien). Es wird bewusst
    NICHT mehr das ganze Archiv von der Platte geparst - das lief im GUI-Thread
    und liess die Oberflaeche bei grossem Archiv einfrieren."""
    with z.lock:
        tr = list(z.index.get(rn, []))
        if not tr:
            tr = list(z.archiv_index.get(rn, []))
    return tr
def countdown_text(deadline_str, offen, jetzt=None):
    """(text, ist_rot) fuer den Abhol-Countdown eines Versenders."""
    if not deadline_str:
        return ("-", False)
    jetzt = jetzt or datetime.now()
    try:
        hh, mm = map(int, deadline_str.split(":"))
    except ValueError:
        return ("-", False)
    ziel = jetzt.replace(hour=hh, minute=mm, second=0, microsecond=0)
    if offen == 0:
        return ("fertig", False)
    delta = (ziel - jetzt).total_seconds()
    if delta < 0:
        return ("ueberfaellig", True)
    std, rest = divmod(int(delta), 3600)
    minu, sek = divmod(rest, 60)
    return (f"{std:02d}:{minu:02d}:{sek:02d}", delta <= DEADLINE_WARN_MIN * 60)
def aktive_versender():
    return [p["name"] for p in VERSENDER_PROFILE]
def deadline_von(versender):
    for p in VERSENDER_PROFILE:
        if p["name"] == versender:
            return p.get("deadline")
    return None
def farbe_von(versender):
    for p in VERSENDER_PROFILE:
        if p["name"] == versender:
            return p.get("farbe", "#444444")
    return "#444444"
def drucker_von(versender, station=None):
    """Druckername fuer Versender (+ optional Packplatz).
    Mit Packplatz: PACKPLATZ_PROFILE[station]['drucker'][versender]. Ohne Treffer
    (oder station=None) faellt es auf den Versender-Standard zurueck -> altes
    Verhalten bleibt unveraendert."""
    if station and station in PACKPLATZ_PROFILE:
        m = PACKPLATZ_PROFILE[station].get("drucker", {})
        if versender in m:
            return m[versender]
    for p in VERSENDER_PROFILE:
        if p["name"] == versender:
            return p.get("drucker", DRUCKER)
    return DRUCKER
def _zerlege_packplatz(eingabe):
    """(station, rest) aus einer Scan-Eingabe. Trennt ein vorangestelltes
    Packplatz-Praefix ab. Kein/keines erkannt -> (PACKPLATZ_STANDARD, eingabe)."""
    if PACKPLATZ_PRAEFIX_AKTIV and eingabe:
        s = eingabe.lstrip()
        for praefix in PACKPLATZ_PROFILE:
            if s[:len(praefix)].upper() == praefix.upper():
                return praefix.upper(), s[len(praefix):]
    return PACKPLATZ_STANDARD, eingabe
def _tag(station):
    """Kurzes Praefix fuer die Statusmeldung, z.B. '[Packplatz 1] '."""
    if station and station in PACKPLATZ_PROFILE:
        return f"[{PACKPLATZ_PROFILE[station]['name']}] "
    return ""
def verarbeite_scan(z, eingabe, ordner, erneut_callback=None):
    """Gemeinsame Scan-Logik. Rueckgabe: (status, meldung, info).
    status: 'gedruckt' | 'erneut' | 'teil' | 'unbekannt' | 'kurz' | 'abbruch'.
    info: dict mit Fortschrittsdaten fuer die Warnanzeige (oder None).
    erneut_callback(nr) -> bool: Rueckfrage bei bereits gedruckter Sendung."""
    # Packplatz-Praefix (welcher Scanner?) vorne abtrennen -> waehlt spaeter den Drucker
    station, eingabe = _zerlege_packplatz(eingabe)
    # --- Sammeldruck-EAN-Bestaetigung: ist fuer diese Station gerade ein
    #     Sammelcode auf die Artikel-EAN offen? Dann wird der naechste Scan
    #     hier ausgewertet, BEVOR er als neuer Sammelcode/neue Rechnung
    #     interpretiert wird:
    #       a) erwartete EAN -> jetzt WIRKLICH drucken, Kontext schliessen
    #       b) sieht wie eine (fremde) EAN aus, passt aber nicht -> rote
    #          Warnung, Kontext bleibt offen (Wiederholung moeglich)
    #       c) irgendetwas anderes (neuer/gleicher Sammelcode, Rechnungsnr.,
    #          ...) -> Kontext verwerfen (NICHTS gedruckt) und normal
    #          weiterverarbeiten; ein erneut gescannter gleicher Sammelcode
    #          landet so wieder im Sammelcode-Zweig unten und fordert die
    #          EAN erneut an.
    eingabe_digits_vor = re.sub(r"\D", "", eingabe or "")
    sctx = z.aktive_sammel_verifikation.get(station)
    if sctx is not None:
        code, erwartet, bez = sctx["code"], sctx["ean"], sctx.get("bez", "")
        if eingabe_digits_vor and erwartet and eingabe_digits_vor == erwartet:
            with z.lock:
                z.aktive_sammel_verifikation.pop(station, None)
            st, msg = drucke_sammel(z, code, ordner, station)
            # info nur bei tatsaechlichem Druckerfolg setzen - sonst wuerde
            # zeige_warnung() unten faelschlich die gruene "fertig"-Bestaetigung
            # zeigen, obwohl der Druck laut drucke_sammel() fehlgeschlagen ist.
            info = {"modus": "sam_ean_fertig", "code": code, "bez": bez} \
                if st == "gedruckt" else None
            return (st, f"{msg}  [EAN bestätigt: {bez}]", info)
        if re.fullmatch(r"\d{8}|\d{12,14}", eingabe_digits_vor):
            info = {"modus": "sam_ean_falsch", "code": code, "bez": bez,
                     "ean": eingabe_digits_vor}
            return ("sam_ean_falsch",
                    f"{_tag(station)}FALSCHER ARTIKEL: EAN {eingabe_digits_vor} "
                    f"gehört NICHT zu Sammeldruck {code} ({bez})", info)
        with z.lock:
            z.aktive_sammel_verifikation.pop(station, None)
        # -> faellt durch zur normalen Verarbeitung unten (auch: neuer Sammelcode)
    # Sammeldruck-Barcode? (z.B. "SAM-1") -> ganzen Block drucken
    # Hinweis: ist der Scanner auf US-Tastaturlayout konfiguriert, kommt das
    # Minus '-' der Sammelcodes auf einem deutschen Windows als 'ß' an
    # (gleiche Taste rechts neben der 0). Ausserdem macht "ß".upper() == "SS".
    # Wir erkennen den Sammelcode daher tolerant: "SAM", optional EIN
    # Trennzeichen (egal welches), dann die Nummer.
    roh = (eingabe or "").strip()
    roh = roh.replace("ẞ", "-").replace("ß", "-").upper()
    msam = re.match(r"^SAM[^0-9]?(\d+)$", roh)
    if msam:
        code = f"SAM-{int(msam.group(1))}"
        g = SAMMEL_LOOKUP.get(code)
        # EAN-Pflicht: nur wenn fuer den Artikel dieser Gruppe eine EAN
        # hinterlegt ist (sammel_zuordnung.csv) UND die EAN-Verifikation
        # global aktiv ist. Sonst wie bisher direkt drucken.
        if g and g.get("ean") and EAN_VERIFIKATION_AKTIV:
            with z.lock:
                z.aktive_sammel_verifikation[station] = {
                    "code": code, "ean": g["ean"], "bez": g.get("bez", ""),
                    "art": g.get("art", "")}
            info = {"modus": "sam_ean_start", "code": code,
                     "bez": g.get("bez", "")}
            return ("sam_ean_start",
                    f"{_tag(station)}Sammeldruck {code} ({g.get('bez','')}): "
                    f"bitte die Artikel-EAN zur Bestätigung scannen", info)
        st, msg = drucke_sammel(z, code, ordner, station)
        return (st, msg, None)
    eingabe_digits = re.sub(r"\D", "", eingabe or "")
    # --- EAN-Verifikation: ist fuer diese Station gerade eine Bestellung offen? -
    #     Dann sind die naechsten Scans die PRODUKT-EANs dieser Bestellung.
    ctx = z.aktive_verifikation.get(station)
    if ctx is not None:
        rn_ctx = ctx["rn"]
        # a) erwartete, noch offene Produkt-EAN -> zaehlen
        if eingabe_digits in ctx["rest"] and ctx["rest"][eingabe_digits] > 0:
            with z.lock:
                ctx["rest"][eingabe_digits] -= 1
                ctx["ist"] += 1
            bez = ctx["namen"].get(eingabe_digits, "")
            if ctx["ist"] >= ctx["soll"]:
                # vollstaendig geprueft -> JETZT drucken
                if not drucke(z, ctx["treffer"], rn_ctx, station):
                    # Druck fehlgeschlagen: NICHT als gedruckt markieren. Der
                    # EAN-Verifikationskontext bleibt bewusst NICHT stehen (die
                    # naechste Bedingung wuerde ohnehin nie wieder True werden,
                    # da "rest" schon leer ist) - Bestellung muss neu gescannt
                    # und die EANs erneut bestaetigt werden.
                    with z.lock:
                        z.aktive_verifikation.pop(station, None)
                    return ("fehler",
                            f"{_tag(station)}FEHLER beim Drucken von {rn_ctx} "
                            f"({ctx['versender']}) - bitte Drucker pruefen und "
                            f"Rechnung erneut scannen", None)
                with z.lock:
                    z.gedruckt[rn_ctx] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    z.gedruckt_heute_pv[ctx["versender"]] += ctx["pakete"]
                    z.aktive_verifikation.pop(station, None)
                    _recompute(z)
                merke_gedruckt(DRUCK_LOG, rn_ctx)
                log_statistik(STATISTIK_DATEI, ctx["versender"], rn_ctx, ctx["pakete"])
                info = {"modus": "ean_fertig", "rn": rn_ctx, "ist": ctx["soll"],
                        "soll": ctx["soll"], "versender": ctx["versender"], "bez": bez}
                return ("gedruckt",
                        f"{_tag(station)}{ctx['versender']} {rn_ctx} gedruckt "
                        f"({ctx['pakete']} Seite(n))  [EAN geprueft "
                        f"{ctx['soll']}/{ctx['soll']}]", info)
            info = {"modus": "ean_teil", "rn": rn_ctx, "ist": ctx["ist"],
                    "soll": ctx["soll"], "versender": ctx["versender"], "bez": bez}
            return ("ean_teil",
                    f"{_tag(station)}{rn_ctx}: {ctx['ist']}/{ctx['soll']} EAN geprueft "
                    f"- {bez} ✓", info)
        # b) gehoert zur Bestellung, aber schon vollstaendig gescannt -> Hinweis
        if eingabe_digits in ctx["eans_all"]:
            bez = ctx["namen"].get(eingabe_digits, "")
            info = {"modus": "ean_falsch", "rn": rn_ctx, "ist": ctx["ist"],
                    "soll": ctx["soll"], "versender": ctx["versender"],
                    "grund": "schon vollstaendig", "bez": bez}
            return ("ean_falsch",
                    f"{_tag(station)}{rn_ctx}: {bez} bereits vollstaendig geprueft "
                    f"({ctx['ist']}/{ctx['soll']})", info)
        # c) sieht aus wie eine EAN, gehoert aber NICHT zu dieser Bestellung
        if re.fullmatch(r"\d{8}|\d{12,14}", eingabe_digits):
            info = {"modus": "ean_falsch", "rn": rn_ctx, "ist": ctx["ist"],
                    "soll": ctx["soll"], "versender": ctx["versender"],
                    "grund": "fremde EAN", "ean": eingabe_digits}
            return ("ean_falsch",
                    f"{_tag(station)}FALSCHER ARTIKEL: EAN {eingabe_digits} gehoert "
                    f"NICHT zu Rechnung {rn_ctx} ({ctx['ist']}/{ctx['soll']} geprueft)",
                    info)
        # d) die laufende Rechnung erneut gescannt -> nur Hinweis "Artikel scannen"
        if eingabe_digits == rn_ctx:
            info = {"modus": "ean_start", "rn": rn_ctx, "ist": ctx["ist"],
                    "soll": ctx["soll"], "versender": ctx["versender"],
                    "namen": list(ctx["namen"].values())}
            return ("ean_teil",
                    f"{_tag(station)}{rn_ctx}: bitte ARTIKEL-EAN scannen "
                    f"({ctx['ist']}/{ctx['soll']})", info)
        # e) etwas anderes (vermutlich neue Rechnung) -> laufende Verifikation
        #    verwerfen (es wurde nichts gedruckt) und Scan normal weiterverarbeiten.
        with z.lock:
            z.aktive_verifikation.pop(station, None)
        # -> faellt durch zur normalen Rechnungs-Logik unten
    # Kein offener Auftrag, aber die Eingabe sieht stark nach einer Produkt-EAN
    # aus (12-14 Ziffern). Vermutlich wurde vergessen, ZUERST die Rechnung zu
    # scannen. Freundlicher Hinweis statt "unbekannte Sendung" - und es kann so
    # keine 13-stellige EAN versehentlich als Rechnungsnummer gedruckt werden.
    # (Laenge 8 bleibt zugelassen, weil Rechnungsnummern 8-stellig sein koennen.)
    if EAN_VERIFIKATION_AKTIV and re.fullmatch(r"\d{12,14}", eingabe_digits):
        return ("unbekannt",
                f"{_tag(station)}EAN {eingabe_digits} gescannt, aber kein Auftrag "
                f"offen - bitte ZUERST die Rechnung scannen.", None)
    rn = re.sub(r"\D", "", eingabe or "")
    if len(rn) < MIN_STELLEN:
        return ("kurz", f"Eingabe zu kurz: {eingabe!r}", None)
    treffer = finde_treffer(z, rn, ordner)
    if not treffer:
        if rn in z.gedruckt:
            # Bereits gedruckt und inzwischen archiviert (z.B. aus einer frueheren
            # Sitzung) -> Label nicht mehr im Schnellzugriff. Klare Meldung statt
            # Einfrieren; Nachdruck bei Bedarf direkt aus dem Archivordner.
            return ("archiviert",
                    f"{rn} wurde bereits gedruckt und ist archiviert - Label nicht "
                    f"mehr im Schnellzugriff. Bei Bedarf aus dem Archivordner drucken.",
                    None)
        return ("unbekannt", f"Keine Sendung zu {rn} gefunden", None)
    versender = treffer[0][2]
    pakete = len(treffer)
    bereits = rn in z.gedruckt
    if bereits:
        if erneut_callback is None or not erneut_callback(rn):
            return ("abbruch", f"{rn} bereits gedruckt - nicht erneut gedruckt", None)
        if not drucke(z, treffer, rn, station):      # Reprint zaehlt NICHT mit
            # Verdacht-Markierung bewusst STEHEN LASSEN - der Fall ist nicht
            # geklaert, wenn der Nachdruck gar nicht erst herauskam.
            return ("fehler",
                    f"{_tag(station)}FEHLER beim erneuten Drucken von {rn} "
                    f"({versender}) - bitte Drucker pruefen", None)
        with z.lock:
            z.scan_fortschritt.pop(rn, None)
            z.letzter_scan_ts.pop(rn, None)
            # Sobald bewusst erneut gedruckt wurde, ist der Fall geklaert ->
            # Verdacht-Markierung fuer diese Rechnungsnummer aufheben, damit die
            # Datei beim naechsten Ordner-Update wieder normal archiviert wird.
            z.verdaechtig.pop(rn, None)
        return ("erneut",
                f"{_tag(station)}{versender} {rn} ERNEUT gedruckt ({pakete} Seite(n))",
                None)
    # --- EAN-Verifikation: Hat die Bestellung EAN-Positionen? Dann ZUERST die
    #     Artikel-EANs scannen lassen. Das Label kommt erst bei Uebereinstimmung;
    #     der EAN-Scan ERSETZT hier das Mengen-Mehrfachscannen. Bestellungen ohne
    #     EAN-Position laufen unveraendert ueber die Stueckzahl-Sicherung unten.
    if hat_ean_pflicht(rn):
        eintrag = EAN_LOOKUP[rn]
        with z.lock:
            z.aktive_verifikation[station] = {
                "rn": rn, "treffer": treffer, "versender": versender,
                "pakete": pakete, "station": station,
                "soll": eintrag["soll"], "ist": 0,
                "rest": dict(eintrag["eans"]),        # ean -> noch offene Scans
                "namen": dict(eintrag["namen"]),
                "eans_all": set(eintrag["eans"].keys()),
            }
            z.scan_fortschritt.pop(rn, None)          # evtl. alten Mengen-Stand loeschen
            z.letzter_scan_ts.pop(rn, None)
        artikel = " + ".join(
            f"{n}x {eintrag['namen'].get(e, '').strip()}".strip()
            for e, n in eintrag["eans"].items())
        info = {"modus": "ean_start", "rn": rn, "ist": 0, "soll": eintrag["soll"],
                "versender": versender, "artikel": artikel,
                "namen": list(eintrag["namen"].values())}
        return ("ean_start",
                f"{_tag(station)}{versender} {rn}: bitte {eintrag['soll']} Artikel-EAN(s) "
                f"scannen – {artikel}", info)
    # --- Mehrfach-Scan-Sicherung: mehrteilige Bestellung muss mehrfach gescannt
    #     werden, bevor das Label kommt. menge = echte Stueckzahl, soll = noetige
    #     Scans (auf MAX_SCANS gedeckelt). Fehlt die Menge -> soll==1 -> sofort.
    menge = artikelmenge(rn)
    soll = soll_scans(rn)
    bekannt = menge_bekannt(rn)
    fertig_info = None
    # Strenger Modus: keine Mengeninfo -> NICHT drucken, laut warnen.
    if MEHRFACH_SCAN_AKTIV and MENGE_PFLICHT and not bekannt:
        info = {"modus": "fehlt", "rn": rn, "ist": 0, "soll": 0,
                "menge": 0, "versender": versender}
        return ("fehlt",
                f"{_tag(station)}{versender} {rn}: KEINE Mengeninfo - NICHT gedruckt. "
                f"Pickliste/mengen_zuordnung.csv aktualisieren!", info)
    if soll > 1:
        with z.lock:
            # Anti-"Durchrattern": zu schnelle Wiederholung NICHT zaehlen
            ist = z.scan_fortschritt.get(rn, 0)
            last = z.letzter_scan_ts.get(rn)
            jetzt_ts = time.monotonic()
            if (MIN_SCAN_ABSTAND_MS > 0 and last is not None
                    and (jetzt_ts - last) * 1000 < MIN_SCAN_ABSTAND_MS):
                info = {"modus": "teil", "rn": rn, "ist": ist, "soll": soll,
                        "menge": menge, "versender": versender}
                return ("teil",
                        f"{_tag(station)}{versender} {rn}: zu schnell - "
                        f"NICHT gezaehlt ({ist}/{soll}). Pro Artikel einzeln scannen!",
                        info)
            ist += 1
            z.scan_fortschritt[rn] = ist
            z.letzter_scan_ts[rn] = jetzt_ts
        if ist < soll:
            info = {"modus": "teil", "rn": rn, "ist": ist, "soll": soll,
                    "menge": menge, "versender": versender}
            return ("teil",
                    f"{_tag(station)}{versender} {rn}: {ist}/{soll} gescannt - "
                    f"noch {soll - ist}x scannen (MEHRERE ARTIKEL!)", info)
        # letzter noetiger Scan erreicht -> jetzt wirklich drucken
        with z.lock:
            z.scan_fortschritt.pop(rn, None)
            z.letzter_scan_ts.pop(rn, None)
        fertig_info = {"modus": "fertig", "rn": rn, "ist": soll, "soll": soll,
                       "menge": menge, "versender": versender}
    if not drucke(z, treffer, rn, station):
        # Weder als gedruckt markieren noch loggen/zaehlen - scan_fortschritt
        # ist zu diesem Zeitpunkt schon zurueckgesetzt (oben, vor dem Druck),
        # ein erneuter vollstaendiger Scan der Bestellung startet die
        # Mehrfach-Scan-Sicherung sauber neu.
        return ("fehler",
                f"{_tag(station)}FEHLER beim Drucken von {rn} ({versender}) - "
                f"bitte Drucker pruefen und erneut scannen", None)
    with z.lock:
        z.gedruckt[rn] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        z.gedruckt_heute_pv[versender] += pakete
        _recompute(z)
    merke_gedruckt(DRUCK_LOG, rn)
    log_statistik(STATISTIK_DATEI, versender, rn, pakete)
    meldung = f"{_tag(station)}{versender} {rn} gedruckt ({pakete} Seite(n))"
    if fertig_info:
        meldung += f"  [{soll}x bestaetigt]"
    elif MEHRFACH_SCAN_AKTIV and not bekannt:
        # Einzelscan-Druck, obwohl Sicherung an ist -> Grund sichtbar machen
        meldung += "  [keine Mengeninfo - Sicherung griff nicht]"
    return ("gedruckt", meldung, fertig_info)
def statistik_text(pfad):
    if not os.path.exists(pfad):
        return "Noch keine Statistik vorhanden."
    heute = datetime.now().strftime("%Y-%m-%d")
    heute_pv, gesamt_pv, tage = Counter(), Counter(), set()
    with open(pfad, encoding="utf-8") as f:
        next(f, None)
        for zeile in f:
            t = zeile.strip().split(";")
            if len(t) >= 4 and t[0]:
                try:
                    pakete = int(t[3])
                except ValueError:
                    pakete = 1
                tage.add(t[0])
                gesamt_pv[t[1]] += pakete
                if t[0] == heute:
                    heute_pv[t[1]] += pakete
    zeilen = [f"Heute ({heute}):"]
    if heute_pv:
        for v in sorted(heute_pv):
            zeilen.append(f"   {v}: {heute_pv[v]}")
        zeilen.append(f"   Summe: {sum(heute_pv.values())}")
    else:
        zeilen.append("   noch nichts versendet")
    zeilen.append(f"\nGesamt ueber {len(tage)} Tag(e):")
    for v in sorted(gesamt_pv):
        zeilen.append(f"   {v}: {gesamt_pv[v]}")
    zeilen.append(f"   Summe: {sum(gesamt_pv.values())}")
    return "\n".join(zeilen)
# =====================================================================
#  GUI (Tkinter)
# =====================================================================
def starte_cockpit():
    import tkinter as tk
    from tkinter import messagebox, font as tkfont
    BG = "#1E1E24"
    BG_FERTIG = "#1B5E20"
    CARD = "#2A2A33"
    FG = "#ECECEC"
    MUTED = "#9AA0A6"
    ROT = "#E53935"
    GRUEN = "#00E676"     # Kreis + Prozent, sobald alle Labels eines Versenders fertig
    GOLD = "#FFD54F"      # Tagesrekord-Anzeige in der Kopfzeile
    z = Zustand()
    z.gedruckt = lade_gedruckt(DRUCK_LOG)
    z.gedruckt_heute_pv = lade_gedruckt_heute(STATISTIK_DATEI)
    z.heute_datum = datetime.now().strftime("%Y-%m-%d")
    z.highscore_sig = None          # Signatur, damit der Rekord nur bei Aenderung neu liest
    lade_post_zuordnung(POST_ZUORDNUNG_CSV)
    lade_sammel_zuordnung(SAMMEL_ZUORDNUNG_CSV)
    lade_mengen_zuordnung(MENGEN_ZUORDNUNG_CSV)
    root = tk.Tk()
    root.title(f"Versand-Cockpit  -  Gasecenter Augsburg  -  v{VERSION}")
    root.geometry("1240x780")
    root.configure(bg=BG)
    f_uhr = tkfont.Font(family="Consolas", size=40, weight="bold")
    f_datum = tkfont.Font(family="Segoe UI", size=13)
    f_status = tkfont.Font(family="Segoe UI", size=15, weight="bold")
    f_kachel = tkfont.Font(family="Segoe UI", size=16, weight="bold")
    f_count = tkfont.Font(family="Consolas", size=26, weight="bold")
    f_klein = tkfont.Font(family="Segoe UI", size=10)
    f_liste = tkfont.Font(family="Consolas", size=11)
    f_scan = tkfont.Font(family="Consolas", size=20, weight="bold")
    f_meld = tkfont.Font(family="Segoe UI", size=13)
    f_high = tkfont.Font(family="Segoe UI", size=11, weight="bold")
    # ----- Kopfzeile -----
    kopf = tk.Frame(root, bg=BG)
    kopf.pack(fill="x", padx=16, pady=(12, 4))
    uhr_lbl = tk.Label(kopf, font=f_uhr, fg=FG, bg=BG)
    uhr_lbl.pack(side="left")
    datum_lbl = tk.Label(kopf, font=f_datum, fg=MUTED, bg=BG)
    datum_lbl.pack(side="left", padx=(14, 0), anchor="s", pady=(0, 10))
    # Firmenlogo (falls logo.png vorhanden) ganz rechts in der Kopfzeile
    logo_img = lade_logo(LOGO_DATEI, LOGO_HOEHE)
    if logo_img is not None:
        logo_lbl = tk.Label(kopf, image=logo_img, bg=BG)
        logo_lbl.image = logo_img           # Referenz halten (sonst weg-GC't)
        logo_lbl.pack(side="right", padx=(16, 0))
    # Rechte Info-Spalte: Status oben, Tagesrekord darunter - vertikal gestapelt,
    # damit ein langer Status-Text ("N Post-Label ohne Zuordnung") nie mit der
    # Rekord-Anzeige kollidiert.
    kopf_rechts = tk.Frame(kopf, bg=BG)
    kopf_rechts.pack(side="right", padx=(16, 0))
    status_lbl = tk.Label(kopf_rechts, font=f_status, fg=FG, bg=BG, anchor="e")
    status_lbl.pack(anchor="e")
    # Tagesmenge: Summe der heute gedruckten Labels ueber ALLE Versender
    # (DHL + DPD + Deutsche Post zusammen) - auf einen Blick, ohne die
    # einzelnen Kacheln von Hand zusammenzaehlen zu muessen.
    tagesmenge_lbl = tk.Label(kopf_rechts, font=f_high, fg="#4FC3F7", bg=BG,
                              justify="right", anchor="e")
    tagesmenge_lbl.pack(anchor="e", pady=(2, 0))
    # Tagesrekord: hoechste Zahl versendeter Bestellungen an einem einzelnen Tag
    highscore_lbl = tk.Label(kopf_rechts, font=f_high, fg=GOLD, bg=BG,
                             justify="right", anchor="e")
    highscore_lbl.pack(anchor="e", pady=(2, 0))
    # ----- Versender-Kacheln -----
    mitte = tk.Frame(root, bg=BG)
    mitte.pack(fill="both", expand=True, padx=12, pady=6)
    kacheln = {}
    versender = aktive_versender()
    for spalte, v in enumerate(versender):
        mitte.columnconfigure(spalte, weight=1, uniform="k")
        akz = farbe_von(v)
        card = tk.Frame(mitte, bg=CARD, highlightbackground=akz,
                        highlightthickness=2, bd=0)
        card.grid(row=0, column=spalte, sticky="nsew", padx=8, pady=4)
        mitte.rowconfigure(0, weight=1)
        tk.Label(card, text=v, font=f_kachel, fg=akz, bg=CARD).pack(pady=(10, 2))
        dl = deadline_von(v)
        dl_txt = f"Abhol-Deadline {dl}" if dl else "ohne Deadline"
        tk.Label(card, text=dl_txt, font=f_klein, fg=MUTED, bg=CARD).pack()
        count_lbl = tk.Label(card, text="-", font=f_count, fg=FG, bg=CARD)
        count_lbl.pack(pady=(2, 6))
        cv = tk.Canvas(card, width=150, height=150, bg=CARD, highlightthickness=0)
        cv.pack()
        prozent_lbl = tk.Label(card, text="", font=f_klein, fg=MUTED, bg=CARD)
        prozent_lbl.pack(pady=(2, 4))
        tk.Label(card, text="offen:", font=f_klein, fg=MUTED, bg=CARD,
                 anchor="w").pack(fill="x", padx=12)
        box = tk.Frame(card, bg=CARD)
        box.pack(fill="both", expand=True, padx=12, pady=(0, 10))
        liste = tk.Text(box, font=f_liste, fg=FG, bg="#22222A", bd=0,
                        height=8, wrap="none")
        liste.pack(fill="both", expand=True)
        liste.configure(state="disabled")
        warn_lbl = None
        if v == "Deutsche Post":
            warn_lbl = tk.Label(card, text="", font=f_klein, fg=ROT, bg=CARD,
                                justify="left", anchor="w", wraplength=300)
            warn_lbl.pack(fill="x", padx=12, pady=(0, 10))
        kacheln[v] = {"count": count_lbl, "canvas": cv, "prozent": prozent_lbl,
                      "liste": liste, "warn": warn_lbl, "akzent": akz}
    # ----- Scan-Zeile + Meldung + Buttons -----
    unten = tk.Frame(root, bg=BG)
    unten.pack(fill="x", padx=16, pady=(4, 12))
    tk.Label(unten, text="Scan:", font=f_scan, fg=FG, bg=BG).pack(side="left")
    scan_var = tk.StringVar()
    scan_entry = tk.Entry(unten, textvariable=scan_var, font=f_scan, width=18,
                          bg="#000000", fg="#00E676", insertbackground="#00E676")
    scan_entry.pack(side="left", padx=10)
    btns = tk.Frame(unten, bg=BG)
    btns.pack(side="right")
    meld_lbl = tk.Label(root, text="Bereit. Barcode scannen ...", font=f_meld,
                        fg=FG, bg=BG, anchor="w")
    meld_lbl.pack(fill="x", padx=16, pady=(0, 4))
    # Status der Mengenliste (Mehrartikel-Sicherung) - dauerhaft sichtbar
    menge_status_lbl = tk.Label(root, text="", font=tkfont.Font(family="Segoe UI",
                                size=11, weight="bold"), fg=MUTED, bg=BG, anchor="w")
    menge_status_lbl.pack(fill="x", padx=16, pady=(0, 8))
    # Warnzeile: Inhalts-Duplikate und Rechnungsnummern, die erneut auftauchen,
    # obwohl sie laut Log schon gedruckt wurden (z.B. Nachsendung mit wieder-
    # verwendeter Rechnungsnummer). Dauerhaft sichtbar, sobald etwas ansteht.
    warnungen_lbl = tk.Label(root, text="", font=tkfont.Font(family="Segoe UI",
                             size=11, weight="bold"), fg=ROT, bg=BG,
                             justify="left", anchor="w", wraplength=1180)
    warnungen_lbl.pack(fill="x", padx=16, pady=(0, 8))
    # ----- Grosses Warn-Overlay "MEHRERE ARTIKEL" (Gefahr-Dreieck) -----------
    # Schwebt als Overlay ueber den Kacheln (place), nicht im Layoutfluss; nimmt
    # keinen Tastaturfokus -> der Scanner schreibt weiter ins Scan-Feld.
    WARN_GELB = "#FFD400"
    WARN_BG = "#161611"
    warn_overlay = tk.Frame(root, bg=WARN_BG, highlightbackground=WARN_GELB,
                            highlightthickness=8, bd=0)
    warn_canvas = tk.Canvas(warn_overlay, width=150, height=140, bg=WARN_BG,
                            highlightthickness=0)
    warn_canvas.pack(side="left", padx=(28, 18), pady=24)
    warn_text = tk.Frame(warn_overlay, bg=WARN_BG)
    warn_text.pack(side="left", padx=(0, 34), pady=24)
    f_warn_titel = tkfont.Font(family="Segoe UI", size=30, weight="bold")
    f_warn_prog = tkfont.Font(family="Segoe UI", size=24, weight="bold")
    f_warn_sub = tkfont.Font(family="Segoe UI", size=14)
    warn_titel = tk.Label(warn_text, text="ACHTUNG – MEHRERE ARTIKEL",
                          font=f_warn_titel, fg=WARN_GELB, bg=WARN_BG)
    warn_titel.pack(anchor="w")
    warn_prog = tk.Label(warn_text, text="", font=f_warn_prog,
                         fg="#FFFFFF", bg=WARN_BG)
    warn_prog.pack(anchor="w", pady=(10, 0))
    warn_sub = tk.Label(warn_text, text="", font=f_warn_sub,
                        fg="#C9C9C9", bg=WARN_BG)
    warn_sub.pack(anchor="w", pady=(4, 0))
    def _zeichne_dreieck(farbe):
        warn_canvas.delete("all")
        # gleichseitiges Warndreieck mit dickem schwarzem Rand + grossem "!"
        warn_canvas.create_polygon(75, 10, 144, 130, 6, 130,
                                   fill=farbe, outline="#000000", width=5,
                                   joinstyle="round")
        warn_canvas.create_text(75, 78, text="!", fill="#000000",
                                font=("Segoe UI", 64, "bold"))
    _warn_timer = {"job": None}
    def verstecke_warnung():
        if _warn_timer["job"]:
            root.after_cancel(_warn_timer["job"])
            _warn_timer["job"] = None
        warn_overlay.place_forget()
    def zeige_warnung(info):
        """info['modus']: 'teil' (gelb, bleibt stehen), 'fertig' (gruen, blendet
        aus) oder 'fehlt' (rot, bleibt stehen - Mengeninfo fehlt, nicht gedruckt)."""
        if _warn_timer["job"]:
            root.after_cancel(_warn_timer["job"])
            _warn_timer["job"] = None
        rn = info.get("rn")
        soll = info.get("soll")
        menge = info.get("menge", soll)
        if info["modus"] == "teil":
            ist = info["ist"]
            _zeichne_dreieck(WARN_GELB)
            warn_overlay.configure(highlightbackground=WARN_GELB)
            warn_titel.configure(text="ACHTUNG – MEHRERE ARTIKEL", fg=WARN_GELB)
            warn_prog.configure(
                text=f"{ist} von {soll} gescannt   –   noch {soll - ist}× scannen",
                fg="#FFFFFF")
            zusatz = f"Rechnung {rn}"
            if menge > soll:               # echte Menge > Deckel: deutlich machen
                zusatz += f"      Bestellung: {menge} Stück  (max. {soll}× nötig)"
            warn_sub.configure(text=zusatz)
            warn_overlay.place(relx=0.5, rely=0.40, anchor="center")
            try:
                root.bell()
            except Exception:
                pass
        elif info["modus"] == "fehlt":     # rot -> Mengeninfo fehlt, NICHT gedruckt
            _zeichne_dreieck("#FF5252")
            warn_overlay.configure(highlightbackground="#FF5252")
            warn_titel.configure(text="MENGENINFO FEHLT", fg="#FF5252")
            warn_prog.configure(text="NICHT gedruckt", fg="#FFFFFF")
            warn_sub.configure(text=f"Rechnung {rn} – Pickliste/mengen_zuordnung.csv "
                                    f"aktualisieren")
            warn_overlay.place(relx=0.5, rely=0.40, anchor="center")
            try:
                root.bell()
            except Exception:
                pass
        elif info["modus"] in ("ean_start", "ean_teil"):
            # blau -> Bestellung offen, jetzt die ARTIKEL-EAN(s) scannen
            ist = info.get("ist", 0)
            EAN_BLAU = "#29B6F6"
            _zeichne_dreieck(EAN_BLAU)
            warn_overlay.configure(highlightbackground=EAN_BLAU)
            warn_titel.configure(text="ARTIKEL-EAN SCANNEN", fg=EAN_BLAU)
            warn_prog.configure(
                text=f"{ist} von {soll} geprüft   –   noch {soll - ist}× scannen",
                fg="#FFFFFF")
            if info["modus"] == "ean_teil" and info.get("bez"):
                warn_sub.configure(text=f"Rechnung {rn}      zuletzt: {info['bez']} ✓")
            else:
                artikel = info.get("artikel") or ", ".join(info.get("namen", []))
                warn_sub.configure(text=f"Rechnung {rn}      {artikel}")
            warn_overlay.place(relx=0.5, rely=0.40, anchor="center")
            try:
                root.bell()
            except Exception:
                pass
        elif info["modus"] == "ean_falsch":   # rot -> falscher/fremder Artikel
            ist = info.get("ist", 0)
            _zeichne_dreieck("#FF5252")
            warn_overlay.configure(highlightbackground="#FF5252")
            if info.get("grund") == "schon vollstaendig":
                warn_titel.configure(text="SCHON GEPRÜFT", fg="#FF5252")
                warn_prog.configure(text=f"{ist} von {soll} geprüft", fg="#FFFFFF")
                warn_sub.configure(text=f"Rechnung {rn} – {info.get('bez','')} "
                                        f"bereits vollständig")
            else:
                warn_titel.configure(text="FALSCHER ARTIKEL", fg="#FF5252")
                warn_prog.configure(text="EAN passt NICHT zur Bestellung", fg="#FFFFFF")
                warn_sub.configure(text=f"Rechnung {rn} – {ist} von {soll} geprüft")
            warn_overlay.place(relx=0.5, rely=0.40, anchor="center")
            try:
                root.bell()
            except Exception:
                pass
        elif info["modus"] == "ean_fertig":   # gruen -> alle EANs geprueft, gedruckt
            _zeichne_dreieck("#00E676")
            warn_overlay.configure(highlightbackground="#00E676")
            warn_titel.configure(text="EAN GEPRÜFT – GEDRUCKT", fg="#00E676")
            warn_prog.configure(text=f"Alle {soll} Artikel bestätigt", fg="#00E676")
            warn_sub.configure(text=f"Rechnung {rn}")
            warn_overlay.place(relx=0.5, rely=0.40, anchor="center")
            _warn_timer["job"] = root.after(2000, verstecke_warnung)
        elif info["modus"] == "sam_ean_start":  # blau -> Sammelcode offen, EAN scannen
            EAN_BLAU = "#29B6F6"
            _zeichne_dreieck(EAN_BLAU)
            warn_overlay.configure(highlightbackground=EAN_BLAU)
            warn_titel.configure(text="SAMMELDRUCK – EAN SCANNEN", fg=EAN_BLAU)
            warn_prog.configure(text=info.get("code", ""), fg="#FFFFFF")
            warn_sub.configure(text=info.get("bez", ""))
            warn_overlay.place(relx=0.5, rely=0.40, anchor="center")
            try:
                root.bell()
            except Exception:
                pass
        elif info["modus"] == "sam_ean_falsch":  # rot -> falscher Artikel fuer Sammelcode
            _zeichne_dreieck("#FF5252")
            warn_overlay.configure(highlightbackground="#FF5252")
            warn_titel.configure(text="FALSCHER ARTIKEL", fg="#FF5252")
            warn_prog.configure(text=f"passt nicht zu {info.get('code','')}", fg="#FFFFFF")
            warn_sub.configure(text=info.get("bez", ""))
            warn_overlay.place(relx=0.5, rely=0.40, anchor="center")
            try:
                root.bell()
            except Exception:
                pass
        elif info["modus"] == "sam_ean_fertig":  # gruen -> EAN bestaetigt, Sammelblock gedruckt
            _zeichne_dreieck("#00E676")
            warn_overlay.configure(highlightbackground="#00E676")
            warn_titel.configure(text="EAN BESTÄTIGT – GEDRUCKT", fg="#00E676")
            warn_prog.configure(text=info.get("code", ""), fg="#00E676")
            warn_sub.configure(text=info.get("bez", ""))
            warn_overlay.place(relx=0.5, rely=0.40, anchor="center")
            _warn_timer["job"] = root.after(2000, verstecke_warnung)
        else:                              # 'fertig' -> gruene Bestaetigung
            _zeichne_dreieck("#00E676")
            warn_overlay.configure(highlightbackground="#00E676")
            warn_titel.configure(text="VOLLSTÄNDIG – GEDRUCKT", fg="#00E676")
            warn_prog.configure(text=f"Alle {soll} Positionen bestätigt", fg="#00E676")
            warn_sub.configure(text=f"Rechnung {rn}")
            warn_overlay.place(relx=0.5, rely=0.40, anchor="center")
            _warn_timer["job"] = root.after(2000, verstecke_warnung)
    # ---- Scan-Auswertung (Auto-Druck ohne Enter) ----
    _pending = {"job": None}
    def erneut_frage(nr):
        # Vorauswahl bewusst auf NEIN: ein versehentliches Scanner-"Enter"
        # bestaetigt damit KEINEN Nachdruck.
        return messagebox.askyesno("Bereits gedruckt",
                                   f"Rechnung {nr} wurde bereits gedruckt.\n"
                                   f"Trotzdem ERNEUT drucken?",
                                   default=messagebox.NO, icon="warning")
    def scan_ausfuehren(_evt=None):
        if _pending["job"]:
            root.after_cancel(_pending["job"])
            _pending["job"] = None
        eingabe = scan_var.get().strip()
        scan_var.set("")
        scan_entry.focus_set()
        if not eingabe:
            return
        try:
            status, meldung, info = verarbeite_scan(z, eingabe, LABEL_ORDNER, erneut_frage)
        except Exception as e:
            import traceback
            traceback.print_exc()
            status, meldung, info = "unbekannt", f"FEHLER bei Eingabe {eingabe!r}: {e}", None
        farben = {"gedruckt": "#00E676", "erneut": "#FFD54F",
                  "unbekannt": ROT, "kurz": ROT, "abbruch": MUTED,
                  "archiviert": WARN_GELB,
                  "teil": WARN_GELB, "fehlt": ROT, "fehler": ROT,
                  "ean_start": "#29B6F6", "ean_teil": "#29B6F6", "ean_falsch": ROT,
                  "sam_ean_start": "#29B6F6", "sam_ean_falsch": ROT}
        meld_lbl.configure(text=meldung, fg=farben.get(status, FG))
        # Grosses Warn-Overlay steuern
        if status in ("teil", "fehlt", "ean_start", "ean_teil", "ean_falsch",
                       "sam_ean_start", "sam_ean_falsch"):
            zeige_warnung(info)                         # gelb/blau/rot - bleibt stehen
        elif status == "gedruckt" and info and info.get("modus") in (
                "fertig", "ean_fertig", "sam_ean_fertig"):
            zeige_warnung(info)                         # gruene Bestaetigung
        else:
            verstecke_warnung()
        aktualisiere_anzeige()
    def auf_tastendruck(_evt=None):
        if _pending["job"]:
            root.after_cancel(_pending["job"])
        _pending["job"] = root.after(SCAN_PAUSE_MS, scan_ausfuehren)
    scan_entry.bind("<Return>", scan_ausfuehren)
    # Auto-Ausloesung bei Tipp-Pause nur, wenn ausdruecklich aktiviert. Scanner
    # senden nach dem Scan ohnehin ein Enter -> Standard aus (siehe Konfig oben);
    # bei haendischer Eingabe wuerde die Pause sonst zu frueh drucken.
    if AUTO_ENTER_BEI_TIPPPAUSE:
        scan_entry.bind("<KeyRelease>", auf_tastendruck)
    # ---- Buttons ----
    def zeige_statistik():
        messagebox.showinfo("Versandstatistik", statistik_text(STATISTIK_DATEI))
    def ordner_aktualisieren():
        aktualisiere_index(z, LABEL_ORDNER)
        aktualisiere_anzeige()
        meld_lbl.configure(text="Ordner neu eingelesen.", fg=FG)
        scan_entry.focus_set()
    def log_zuruecksetzen():
        if not os.path.exists(DRUCK_LOG):
            messagebox.showinfo("Log", "Kein Druck-Log vorhanden.")
            return
        if messagebox.askyesno("Log zuruecksetzen",
                               "Druck-Log wirklich leeren? (Statistik bleibt)"):
            try:
                os.remove(DRUCK_LOG)
            except OSError:
                pass
            with z.lock:
                z.gedruckt = {}
                z.gedruckt_heute_pv = Counter()
                z.scan_fortschritt = {}
                z.letzter_scan_ts = {}
                z.aktive_verifikation = {}
                z.verdaechtig = defaultdict(list)
                z.duplikate = {}
                _recompute(z)
            aktualisiere_anzeige()
            meld_lbl.configure(text="Druck-Log geleert.", fg=FG)
        scan_entry.focus_set()
    def beenden():
        stop.set()
        root.destroy()
    for txt, cmd in [("Statistik", zeige_statistik),
                     ("Ordner aktualisieren", ordner_aktualisieren),
                     ("Log zuruecksetzen", log_zuruecksetzen),
                     ("Beenden", beenden)]:
        tk.Button(btns, text=txt, command=cmd, font=f_klein, bd=0,
                  bg="#3A3A44", fg=FG, activebackground="#4A4A55",
                  activeforeground=FG, padx=10, pady=6).pack(side="left", padx=4)
    # ---- Kreisdiagramm zeichnen ----
    def zeichne_kreis(cv, gedruckt, gesamt, akz, fertig=False):
        cv.delete("all")
        x0, y0, x1, y1 = 15, 15, 135, 135
        if gesamt == 0:
            cv.create_oval(x0, y0, x1, y1, outline="#444", width=14)
            cv.create_text(75, 75, text="-", fill=MUTED,
                           font=("Segoe UI", 18, "bold"))
            return
        frac = gedruckt / gesamt
        cv.create_oval(x0, y0, x1, y1, outline="#3A3A44", width=14)
        if frac > 0:
            cv.create_arc(x0, y0, x1, y1, start=90, extent=-359.999 * frac,
                          style="arc", outline=akz, width=14)
        cv.create_text(75, 70, text=f"{int(round(frac*100))}%",
                       fill=(akz if fertig else FG),
                       font=("Segoe UI", 20, "bold"))
        cv.create_text(75, 96, text=f"{gedruckt}/{gesamt}", fill=MUTED,
                       font=("Segoe UI", 11))
    # ---- Gesamt-Aktualisierung der Anzeige ----
    def aktualisiere_anzeige():
        # Status der Mengenliste (Mehrartikel-Sicherung) anzeigen
        ok, _n, mtxt = mengen_status()
        _eok, _en, etxt = ean_status()
        if EAN_VERIFIKATION_AKTIV:
            mtxt = f"{mtxt}   |   {etxt}"
        if not MEHRFACH_SCAN_AKTIV:
            menge_status_lbl.configure(text="⚠ " + mtxt, fg=MUTED)
        elif ok:
            menge_status_lbl.configure(text="✓ " + mtxt, fg=GRUEN)
        else:
            menge_status_lbl.configure(text="⚠ " + mtxt, fg=ROT)
        # Duplikat-/Verdaechtig-Warnungen (Datei-Handling-Absicherung)
        warn_texte = warnungen_liste(z)
        warnungen_lbl.configure(
            text=("⚠ " + "   |   ".join(warn_texte)) if warn_texte else "")
        with z.lock:
            gedruckt_pv = dict(z.gedruckt_heute_pv)
            offen_pv = dict(z.offen_pv)
            fortschritt = dict(z.scan_fortschritt)
            # offene EAN-Verifikationen je Rechnung -> {rn: (ist, soll)}
            ean_offen = {c["rn"]: (c["ist"], c["soll"])
                         for c in z.aktive_verifikation.values()}
        offen = offene_auftraege(z)
        unzuordenbar = post_unzuordenbar_liste(z)
        offen_gesamt = sum(len(v) for v in offen.values())
        for v, w in kacheln.items():
            gedr = gedruckt_pv.get(v, 0)        # heute gedruckte Labels (persistent)
            off = offen_pv.get(v, 0)            # noch offene Labels im Ordner
            ges = gedr + off                    # heutige Gesamtzahl je Versender
            txt, rot = countdown_text(deadline_von(v), off)
            w["count"].configure(text=txt, fg=(ROT if rot else FG))
            fertig_v = (ges > 0 and off == 0)   # alles dieses Versenders gedruckt
            akz = GRUEN if fertig_v else w["akzent"]
            zeichne_kreis(w["canvas"], gedr, ges, akz, fertig_v)
            w["prozent"].configure(text=f"{gedr} gedruckt / {ges} gesamt")
            eintraege = offen.get(v, [])
            w["liste"].configure(state="normal")
            w["liste"].delete("1.0", "end")
            if eintraege:
                for nr, emp in eintraege:
                    p = fortschritt.get(nr)
                    if nr in ean_offen:
                        ist, soll = ean_offen[nr]
                        extra = f"   [EAN {ist}/{soll}]"
                    elif p:
                        extra = f"   [{p}/{soll_scans(nr)} gescannt]"
                    else:
                        extra = ""
                    w["liste"].insert("end", f"{nr}  {emp}{extra}\n")
            else:
                w["liste"].insert("end", "  (nichts offen)\n")
            w["liste"].configure(state="disabled")
            if w["warn"] is not None:
                if unzuordenbar:
                    namen = "\n".join(f"  ! {n} ({p})" for n, p in unzuordenbar[:6])
                    mehr = "" if len(unzuordenbar) <= 6 else f"\n  (+{len(unzuordenbar)-6} weitere)"
                    w["warn"].configure(
                        text=f"OHNE Zuordnung ({len(unzuordenbar)}) - Adresse pruefen:\n{namen}{mehr}")
                else:
                    w["warn"].configure(text="")
        # ---- Tagesmenge (Summe ueber alle Versender) + Tagesrekord ----
        heute_summe = sum(gedruckt_pv.values())
        aufschluesselung = " · ".join(
            f"{v} {n}" for v, n in sorted(gedruckt_pv.items()) if n > 0)
        tagesmenge_lbl.configure(
            text=f"Σ Heute gesamt: {heute_summe}"
                 + (f"   ({aufschluesselung})" if aufschluesselung else ""))
        # Tagesrekord nur neu berechnen, wenn heute etwas dazukam
        if heute_summe != z.highscore_sig:
            z.highscore_sig = heute_summe
            highscore_lbl.configure(text=highscore_text(STATISTIK_DATEI))
        fertig = (offen_gesamt == 0 and not unzuordenbar)
        flaechen = [root, kopf, unten, btns]
        transparente = ([uhr_lbl, datum_lbl, meld_lbl]
                        + list(kopf.winfo_children())
                        + list(kopf_rechts.winfo_children()))
        neu_bg = BG_FERTIG if fertig else BG
        for w_ in flaechen:
            w_.configure(bg=neu_bg)
        for c in transparente:
            try:
                c.configure(bg=neu_bg)
            except tk.TclError:
                pass
        if fertig:
            txt = "Alles gedruckt" if sum(gedruckt_pv.values()) > 0 else "Bereit - nichts offen"
            status_lbl.configure(text=txt, fg="#A5D6A7", bg=neu_bg)
        elif unzuordenbar and offen_gesamt == 0:
            status_lbl.configure(
                text=f"{len(unzuordenbar)} Post-Label ohne Zuordnung", fg=ROT, bg=neu_bg)
        else:
            status_lbl.configure(
                text=f"In Arbeit - {offen_gesamt} offen", fg=FG, bg=neu_bg)
    # ---- Takt: Uhr + Countdown jede Sekunde, Anzeige alle 0,5 s ----
    def takt():
        jetzt = datetime.now()
        uhr_lbl.configure(text=jetzt.strftime("%H:%M:%S"))
        datum_lbl.configure(text=jetzt.strftime("%A, %d.%m.%Y"))
        heute = jetzt.strftime("%Y-%m-%d")
        if heute != z.heute_datum:          # Tageswechsel -> Tageszaehler neu laden
            with z.lock:
                z.heute_datum = heute
                z.gedruckt_heute_pv = lade_gedruckt_heute(STATISTIK_DATEI)
        aktualisiere_anzeige()
        root.after(500, takt)
    # ---- Hintergrund-Thread: Ordner-Ueberwachung ----
    stop = threading.Event()
    def beobachter():
        while not stop.is_set():
            try:
                aktualisiere_index(z, LABEL_ORDNER)
            except Exception:
                pass
            stop.wait(POLL_SEKUNDEN)
    if not os.path.isdir(LABEL_ORDNER):
        messagebox.showwarning("Ordner fehlt",
                               f"Label-Ordner nicht erreichbar:\n{LABEL_ORDNER}\n\n"
                               "Pfad in der Konfiguration pruefen.")
    aktualisiere_index(z, LABEL_ORDNER)
    threading.Thread(target=beobachter, daemon=True).start()
    scan_entry.focus_set()
    takt()
    # Nur fuer den automatisierten Oberflaechen-Test (im Produktivbetrieb inaktiv):
    if os.environ.get("COCKPIT_SELFCLOSE"):
        def _shot():
            try:
                root.update_idletasks(); root.update()
                subprocess.run(["import", "-window", "root",
                                os.environ["COCKPIT_SELFCLOSE"]], check=False)
            except Exception as e:
                print("Screenshot-Fehler:", e)
            root.after(200, root.destroy)
        root.after(1600, _shot)
    root.mainloop()
if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "selftest":
        sys.exit(0)
    # Diagnose:  py scan_druck.py mengentest [Rechnungsnummer]
    # Zeigt, ob die Mengenliste gefunden/geladen wird und was fuer eine Rechnung
    # an Scans verlangt wuerde. Hilft, "warum druckt das in 1 Scan?" zu klaeren.
    if len(sys.argv) > 1 and sys.argv[1] == "mengentest":
        lade_mengen_zuordnung(MENGEN_ZUORDNUNG_CSV)
        da = os.path.exists(MENGEN_ZUORDNUNG_CSV)
        print(f"MENGEN_ZUORDNUNG_CSV : {MENGEN_ZUORDNUNG_CSV}")
        print(f"Datei vorhanden      : {da}")
        print(f"Eintraege geladen    : {len(MENGEN_LOOKUP)}")
        print(f"MEHRFACH_SCAN_AKTIV  : {MEHRFACH_SCAN_AKTIV}   MAX_SCANS={MAX_SCANS}   "
              f"MENGE_PFLICHT={MENGE_PFLICHT}")
        if len(sys.argv) > 2:
            rn = re.sub(r"\D", "", sys.argv[2])
            print(f"\nRechnung {rn}:")
            print(f"  in Mengenliste     : {menge_bekannt(rn)}")
            print(f"  Stueckzahl (menge) : {artikelmenge(rn)}")
            print(f"  noetige Scans soll : {soll_scans(rn)}")
        else:
            bsp = list(MENGEN_LOOKUP.items())[:10]
            if bsp:
                print("\nBeispiele (erste 10):")
                for k, v in bsp:
                    print(f"  {k} -> {v} Stueck (soll {min(v, MAX_SCANS)})")
        sys.exit(0)
    # Diagnose:  py scan_druck.py eantest [Rechnungsnummer]
    # Zeigt, ob die EAN-Liste gefunden/geladen wird und welche EAN(s) eine
    # Rechnung zur Verifikation verlangt. Hilft zu klaeren, "warum will der die
    # EAN sehen / warum nicht?".
    if len(sys.argv) > 1 and sys.argv[1] == "eantest":
        lade_ean_zuordnung(EAN_ZUORDNUNG_CSV)
        da = os.path.exists(EAN_ZUORDNUNG_CSV)
        print(f"EAN_ZUORDNUNG_CSV    : {EAN_ZUORDNUNG_CSV}")
        print(f"Datei vorhanden      : {da}")
        print(f"Bestellungen mit EAN : {len(EAN_LOOKUP)}")
        print(f"EAN_VERIFIKATION_AKTIV: {EAN_VERIFIKATION_AKTIV}")
        if len(sys.argv) > 2:
            rn = re.sub(r"\D", "", sys.argv[2])
            print(f"\nRechnung {rn}:")
            print(f"  EAN-Pflicht        : {hat_ean_pflicht(rn)}")
            eintrag = EAN_LOOKUP.get(rn)
            if eintrag:
                print(f"  Soll-Scans gesamt  : {eintrag['soll']}")
                for e, n in eintrag["eans"].items():
                    print(f"    EAN {e} x{n}  ({eintrag['namen'].get(e,'')})")
        else:
            bsp = list(EAN_LOOKUP.items())[:10]
            if bsp:
                print("\nBeispiele (erste 10):")
                for k, v in bsp:
                    eans = ", ".join(f"{e}x{n}" for e, n in v["eans"].items())
                    print(f"  {k} -> soll {v['soll']}  [{eans}]")
        sys.exit(0)
    # Diagnose:  py scan_druck.py sammeltest [SAM-n]
    # Zeigt, ob die Sammeldruck-Zuordnung geladen ist und ob/welche EAN ein
    # Sammelcode zur Bestaetigung verlangt. Hilft zu klaeren, "warum will der
    # beim Sammeldruck eine EAN sehen / warum nicht?".
    if len(sys.argv) > 1 and sys.argv[1] == "sammeltest":
        lade_sammel_zuordnung(SAMMEL_ZUORDNUNG_CSV)
        da = os.path.exists(SAMMEL_ZUORDNUNG_CSV)
        print(f"SAMMEL_ZUORDNUNG_CSV : {SAMMEL_ZUORDNUNG_CSV}")
        print(f"Datei vorhanden      : {da}")
        print(f"Sammelcodes geladen  : {len(SAMMEL_LOOKUP)}")
        print(f"EAN_VERIFIKATION_AKTIV: {EAN_VERIFIKATION_AKTIV}")
        if len(sys.argv) > 2:
            code = sys.argv[2].strip().upper()
            if not code.startswith("SAM"):
                code = f"SAM-{code}"
            g = SAMMEL_LOOKUP.get(code)
            print(f"\nSammelcode {code}:")
            if not g:
                print("  nicht gefunden")
            else:
                print(f"  Artikel            : {g['art']} - {g['bez']}")
                print(f"  Bestellungen       : {len(g['rnr'])}")
                print(f"  EAN hinterlegt     : {g.get('ean') or '(keine)'}")
                pflicht = bool(g.get("ean")) and EAN_VERIFIKATION_AKTIV
                print(f"  EAN-Bestaetigung   : {'JA - vor Druck noetig' if pflicht else 'nein, druckt direkt'}")
        else:
            bsp = list(SAMMEL_LOOKUP.items())[:10]
            if bsp:
                print("\nBeispiele (erste 10):")
                for k, v in bsp:
                    ean = v.get("ean") or "-"
                    print(f"  {k} -> {v['art']} ({len(v['rnr'])} Best.)  EAN: {ean}")
        sys.exit(0)
    starte_cockpit()
