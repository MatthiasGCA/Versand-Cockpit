# -*- coding: utf-8 -*-
"""
packliste.py - Kommissionierliste aus Gasecenter-Ausgangsrechnungen (Amicron-PDF)
=================================================================================
Liest einen Ordner mit Ausgangsrechnungen (PDF, Amicron "entzerrtes" Layout mit
Einzelpreis/G-Preis/Rechnungsbetrag und Lagerort-Feld) und erzeugt:

  1. eine ZWEITEILIGE PDF-Kommissionierliste
     - Teil A "Sammelliste": je Lagerkategorie die Artikel aggregiert
       (Gesamtmenge ueber alle Rechnungen), mit Abhak-Kaestchen.
     - Teil B "Packuebersicht je Rechnung": Kopf mit Rechnungsnr., Kd-Nr.,
       Datum, Empfaenger und Code128-Barcode der Rechnungsnummer, darunter
       die Positionen mit Abhak-Kaestchen.
  2. die Datei "post_zuordnung.csv" (Bruecke fuer scan_druck.py):
       Rechnungsnummer;Name;Strasse;Hausnummer;PLZ
     Format exakt so, wie scan_druck.py es einliest (Semikolon, Kopfzeile,
     UTF-8). Damit ordnet das Cockpit Deutsche-Post-Labels (ohne Rechnungs-
     nummer) ueber die Lieferadresse zu.

 Aufruf:
    python packliste.py <rechnungs_ordner> "<stammdaten_oder_leer>" <ausgabe.pdf>
Beispiel:
    python packliste.py /mnt/user-data/uploads "" /mnt/user-data/outputs/Pickliste.pdf

Betragsabgleich (zwei Sicherheitspruefungen):
  - je Position:  Menge x Einzelpreis == G-Preis
  - je Rechnung:  Summe G-Preise == Rechnungsbetrag
Rechnungen, deren Abgleich nicht aufgeht, werden trotzdem aufgenommen, aber im
Protokoll und mit einem Hinweis in der Packuebersicht deutlich markiert.

Kategorie-Zuordnung allein aus dem Lagerort-Feld der Packliste:
  Topseller, Eigenfertigung, Schlauchlager, Palettenlager, Poolchemie,
  Kleinteile, Regallager, DPD / Warenpost, Allgemein  (Regeln unten in classify()).
"""

import os
import re
import sys
import csv
import glob
import io
import shutil
from datetime import datetime

import pdfplumber
from pypdf import PdfReader, PdfWriter
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Flowable,
    KeepTogether, PageBreak,
)
from reportlab.graphics.barcode import code128
from reportlab.graphics.shapes import Drawing

# ----------------------------------------------------------------------------
# KONFIGURATION
# ----------------------------------------------------------------------------

VERSION = "2026-08-20a"          # Versionsschema: JJJJ-MM-TT + Kleinbuchstabe je
# Aenderung am selben Tag (erste = a, dann b, c ...; neuer Tag beginnt wieder bei a).
# 2026-08-20a: Falsche PLZ bei Amazon-Business-Lieferadressen mit angehaengter
#   PO-Referenz behoben (Rechnung 1700723, "VDS Getriebe"/Wolfern: Namenszeile
#   "GmbHPO126-0287 VDS Getriebe" verschmilzt Firmenname und Bestellreferenz
#   ohne Trennzeichen). Der Oesterreich-Zweig in extrahiere_adresse() suchte die
#   PLZ-Zeile bisher VORWAERTS durch die Lieferadresse; die Bestellreferenz
#   "0287" erfuellte dabei zufaellig das Muster "4 Ziffern + Leerzeichen + Text"
#   und wurde als PLZ genommen (post_zuordnung.csv bekam "0287" statt der
#   echten PLZ "4493" -> Deutsche-Post-Zuordnung per PLZ+Hausnummer scheiterte).
#   Jetzt wird wie im DE-Zweig direkt darunter RUECKWAERTS gesucht (PLZ+Ort
#   steht so gut wie immer als letzte inhaltliche Zeile), das findet nun
#   zuverlaessig die echte PLZ-Zeile statt einer zufaellig passenden Ziffernfolge
#   weiter oben im Adressblock.
# 2026-08-11a: Fehlende PLZ bei abweichender Lieferadresse behoben (Rechnung
#   1699130, Amazon "Jan Glöer": Rechnungsadresse AT-1020 Wien, Lieferadresse
#   24404 Maasholm/Deutschland). extrahiere_adresse() bestimmte das Land bisher
#   IMMER aus Liefer- UND Rechnungsadresse zusammen (fuer den Fall 1695459, bei
#   dem nur die Rechnungsadresse das "AT-"-Kuerzel trug). Das riss hier den
#   umgekehrten Fall mit: eine oesterreichische Rechnungsadresse liess eine
#   voellig normale DEUTSCHE Lieferadresse faelschlich als "AT" gelten -> es
#   wurde nur noch nach einer 4-stelligen PLZ gesucht, die echte 5-stellige PLZ
#   passte auf kein Muster mehr und blieb leer (keine post_zuordnung.csv-
#   Zuordnung, keine PLZ auf der Sammeldruck-Pickliste). Jetzt wird das Land
#   zuerst NUR aus der Lieferadresse bestimmt; nur wenn die Lieferadresse WEDER
#   ein AT-Signal NOCH eine gueltige 5-stellige PLZ enthaelt, wird zusaetzlich
#   die Rechnungsadresse herangezogen (deckt weiterhin Fall 1695459 ab).
# 2026-07-20b: Erstellungsdatum + Uhrzeit werden als Fuss auf jeder Seite der
#   Pickliste abgedruckt. Ausserdem Fehltrigger der Gewichtsartikel-Erkennung
#   behoben: Artikel wie 'AutoPR-11' (Modellnummer 11 == '11kg' im Text = fuer
#   11-kg-Flaschen) galten faelschlich als 11-kg-Gewichtsartikel; jetzt zaehlt nur
#   noch ein Komma-Dezimal (0,5 / 1,0 / 3,0) als Gewichtsangabe.
# 2026-07-20a: Packungsartikel "<Zahl>/<N>" (z.B. 2248/10) mit "<N> Stück" im Text
#   werden in der Packuebersicht wie Gewichtsartikel behandelt: "<N> Stück" -> in
#   der Bezeichnung "GESAMTMENGE", und die Mengenspalte zeigt die tatsaechliche
#   Gesamtstueckzahl (Packungsgroesse x bestellte Menge, z.B. 2x -> 20 Stück).

# Reihenfolge der Kategorien in der Sammelliste (Laufweg im Lager).
KATEGORIE_REIHENFOLGE = [
    "Topseller", "Eigenfertigung", "Schlauchlager", "Palettenlager",
    "Poolchemie", "Kleinteile", "Regallager", "DPD / Warenpost", "Allgemein",
]

# Kleine Nachlaeufe kurz vor der Abhol-Deadline (meist nur ein paar Bestellungen)
# lohnen sich nicht auf mehrere Lagerort-Listen aufgeteilt - dann kommt ALLES auf
# EINE gemeinsame Liste (weiterhin sortiert nach Artikelnummer). Ab dieser Anzahl
# Bestellungen im Lauf wird wieder wie gewohnt nach Lagerort/Kategorie getrennt.
KOMBINIERT_SCHWELLE = 20
KOMBINIERT_LABEL = "Alle Bestellungen"

# Poolchemie wird ueber den Lagerort "Poolchemie" erkannt. Falls einzelne
# Artikelnummern fest zu Poolchemie gehoeren (ohne dass der Lagerort es sagt),
# hier eintragen:  POOLCHEMIE_ARTIKEL = {"12345", "67890"}
POOLCHEMIE_ARTIKEL = set()

# Sammeldruck: ab wie vielen gleichen Einzel-Artikel-Bestellungen (je 1 Stueck,
# nur ein Artikel in der Bestellung) ein gemeinsamer Sammel-Barcode erzeugt wird.
SAMMEL_MIN = 2

# Spaltenbaender nach x-Position.
# ---------------------------------------------------------------------------
# AMICRON-LAYOUT MIT EAN-SPALTE (Stand 26.06.2026):
# Amicron druckt seit der Layout-Umstellung zwischen Einheit und Menge eine
# zusaetzliche EAN-Spalte. Dadurch sind Menge/Einzelpreis/G-Preis je ~eine
# Spalte nach RECHTS gerutscht. Die Baender wurden an der neuen Anlage neu
# vermessen (Werte = Mittelpunkt-x der jeweiligen Spalte, mit Rand fuer breite
# Betraege, da diese rechtsbuendig nach links wachsen):
#   Einheit  'Stck.'          mid ~291
#   EAN      '4262419350831'  mid ~337   <-- NEU
#   Menge    '1,00'           mid ~415
#   EPreis   '6,89'           mid ~474
#   GPreis   '6,89'           mid ~508
# Faellt eine Layout-Variante auf (Warnung "Menge/Betrag" in der Konsole),
# hier nachjustieren. Die Selbstpruefung Menge*EP==GP und Summe==Rechnungsbetrag
# faengt grobe Fehlkalibrierungen ab.
ART_MAX = 112
BEZ_LO, BEZ_HI = 145, 281
EINH_LO, EINH_HI = 281, 305     # Einheit (EAN liegt rechts davon, eigenes Band)
EAN_LO, EAN_HI = 305, 380       # NEU: EAN-Spalte (zwischen Einheit und Menge)
MENGE_LO, MENGE_HI = 380, 440   # war 345..415 (vor EAN-Spalte)
EP_LO, EP_HI = 440, 487         # war 415..467
GP_LO, GP_HI = 487, 545         # war 467..540
TABLE_TOP, TABLE_BOT = 26, 388   # Tabellenbereich (unter Kopf, ueber Rechnungsbetrag)

# Eine EAN/GTIN ist eine reine Ziffernfolge mit 8 oder 12-14 Stellen. Das Muster
# unterscheidet sie zuverlaessig von Mengen/Betraegen (die ein Komma fuehren) und
# von Artikelnummern (die i.d.R. Buchstaben enthalten und ganz links stehen).
EAN_RE = re.compile(r"^\d{8}$|^\d{12,14}$")

CENT = 0.005  # Toleranz fuer Betragsabgleich


# ----------------------------------------------------------------------------
# HILFSFUNKTIONEN
# ----------------------------------------------------------------------------

def num(s):
    """Deutsche Zahl '1.234,56' -> float; None wenn keine Zahl."""
    s = (s or "").strip()
    if not re.search(r"\d", s):
        return None
    s = s.replace(".", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return None


def _hausnr(street):
    """Wie in scan_druck.py: letzte Zahl (+optionaler Buchstabe) am Ende."""
    m = re.search(r"(\d+\s*[a-zA-Z]?)\s*$", (street or "").strip())
    return re.sub(r"\s+", "", m.group(1)) if m else ""


def _plz(s, land="DE"):
    """PLZ aus einem String. DE = 5-stellig, AT = 4-stellig ('AT-8020', 'A-8020'
    oder '8020 Ort'). So passt die in die post_zuordnung.csv geschriebene PLZ zu
    der, die scan_druck.py vom (auslaendischen) Label liest."""
    if land == "AT":
        m = (re.search(r"\b(?:AT|A)-?\s?(\d{4})\b", s or "")
             or re.search(r"\b(\d{4})\b", s or ""))
        return m.group(1) if m else ""
    m = re.search(r"\b(\d{5})\b", s or "")
    return m.group(1) if m else ""


def _land_aus_zeilen(zeilen):
    """'AT' wenn eine Adresszeile auf Oesterreich deutet, sonst 'DE'.
    Erkannt werden 'ÖSTERREICH'/'AUSTRIA' sowie das ISO-Praefix vor der PLZ
    ('AT-8020', 'A-8020')."""
    txt = " ".join(zeilen or []).upper().replace("Ö", "OE")
    if re.search(r"OESTERREICH|AUSTRIA|\b(?:AT|A)-\s?\d{4}\b", txt):
        return "AT"
    return "DE"


def in_band(w, lo, hi):
    """Spaltenzuordnung ueber den MITTELPUNKT des Wortes (nicht die linke Kante).
    Betraege sind rechtsbuendig: ein breiter Wert wie '194,97' beginnt weit links
    (x0 ggf. unter der Bandgrenze) und wuerde per linker Kante ins falsche Band
    (Einzelpreis statt G-Preis) fallen. Der Mittelpunkt liegt stabil in der
    richtigen Spalte."""
    mid = (w["x0"] + w["x1"]) / 2.0
    return lo <= mid < hi


def cluster_lines(words, tol=5):
    """Woerter zu Zeilen gruppieren (top-Toleranz)."""
    ws = sorted(words, key=lambda w: (w["top"], w["x0"]))
    lines, cur, cy = [], [], None
    for w in ws:
        if cy is None or abs(w["top"] - cy) <= tol:
            cur.append(w)
            cy = w["top"] if cy is None else (cy + w["top"]) / 2
        else:
            lines.append(cur)
            cur = [w]
            cy = w["top"]
    if cur:
        lines.append(cur)
    return lines


def line_text(line):
    return " ".join(w["text"] for w in sorted(line, key=lambda w: w["x0"]))


def _lagerort_wert(txt):
    """Vollstaendigen Lagerort-Wert aus einer 'Lagerort:'-Zeile ziehen.
    Nimmt ALLES nach dem letzten ':' und normalisiert Mehrfach-Leerzeichen,
    damit mehrteilige Angaben ('R8 E2', 'A8 + A9') vollstaendig erhalten
    bleiben (frueher wurde nur das erste Token genommen -> 'R8 E2' landete auf
    'Allgemein' statt Kleinteile, und die Fachanzeige war unvollstaendig).
    Leere Zeile ('Lagerort:' ohne Wert) -> ''."""
    after = txt.rsplit(":", 1)[-1]
    return " ".join(after.split())


# ----------------------------------------------------------------------------
# KATEGORIE-ZUORDNUNG (allein aus dem Lagerort)
# ----------------------------------------------------------------------------

GRID_RE = re.compile(r"\b([A-K])\s?([1-8])\b")


def classify(lagerorte, art):
    """Lagerkategorie aus den Lagerort-Strings (Reihenfolge = Prioritaet)."""
    text = " ".join(lagerorte)
    low = text.lower()
    if art and art in POOLCHEMIE_ARTIKEL:
        return "Poolchemie"
    if not lagerorte:
        return "Allgemein"
    if "wapo" in low or "warenpost" in low:
        return "DPD / Warenpost"
    if "topseller" in low:
        return "Topseller"
    if "fertigung" in low:
        return "Eigenfertigung"
    if "schläuche" in low or "schlauch" in low or "schlaeuche" in low:
        return "Schlauchlager"
    if "schwere artikel" in low or "palette" in low:
        return "Palettenlager"
    if "poolchemie" in low:
        return "Poolchemie"
    if "kleinteile" in low:                 # Lagerort woertlich "Kleinteile"
        return "Kleinteile"
    mreg = re.search(r"regal\s*0*(\d+)", low)
    if mreg:
        nr = int(mreg.group(1))
        return "Kleinteile" if 21 <= nr <= 30 else "Regallager"
    if GRID_RE.search(text.upper()):
        return "Kleinteile"
    return "Allgemein"


def ist_versand(art, bez):
    """Versandkosten / nicht-koerperliche Position (nicht ins Lager)."""
    if "versandkosten" in (bez or "").lower():
        return True
    if not art and not (bez or "").strip():
        return True
    return False


def auftrag_kategorie(r):
    """Haupt-Lagerort einer ganzen Bestellung = die Kategorie mit der HOECHSTEN
    Prioritaet (= weitestVORNE in KATEGORIE_REIHENFOLGE) ueber alle echten
    Positionen. So landet jede Bestellung in genau EINER Kommissionierliste."""
    kats = set()
    for p in r["positionen"]:
        if ist_versand(p["art"], p["bez"]):
            continue
        kats.add(classify(p["lagerorte"], p["art"]))
    for k in KATEGORIE_REIHENFOLGE:
        if k in kats:
            return k
    return "Allgemein"


def finde_sammelgruppen(rechnungen, min_anzahl=SAMMEL_MIN):
    """Bestellungen, die genau EINEN Artikel mit Menge 1 enthalten (Versand-
    zeilen zaehlen nicht), nach Artikelnummer gruppieren. Gruppen ab
    `min_anzahl` Bestellungen bekommen einen Sammel-Barcode 'SAM-n'.
    Rueckgabe: Liste von dicts {code, art, bez, einh, rnr:[...]}.

    ACHTUNG: die hier vergebenen Codes sind nur VORLAEUFIG (reine
    Aufzaehlung "SAM-1, SAM-2, ..." innerhalb DIESES Laufs) - main() ruft
    danach IMMER merge_sammelgruppen() auf, um Kollisionen mit noch nicht
    abgearbeiteten Codes aus einem frueheren Lauf zu vermeiden, bevor die
    Codes gedruckt/geschrieben werden."""
    nach_artikel = {}
    for r in rechnungen:
        echte = [p for p in r["positionen"] if not ist_versand(p["art"], p["bez"])]
        if len(echte) != 1:
            continue
        p = echte[0]
        if p["menge"] is None or abs(p["menge"] - 1.0) > 1e-9:
            continue
        if not p["art"]:
            continue
        g = nach_artikel.setdefault(
            p["art"], {"art": p["art"], "bez": p["bez"], "einh": p["einh"],
                       "ean": "", "rnr": []})
        if not g["ean"] and p.get("ean"):
            g["ean"] = p["ean"]
        if r["rnr"]:
            g["rnr"].append(r["rnr"])
    gruppen = [g for g in nach_artikel.values() if len(g["rnr"]) >= min_anzahl]
    gruppen.sort(key=lambda g: (-len(g["rnr"]), g["art"]))
    for i, g in enumerate(gruppen, 1):
        g["code"] = f"SAM-{i}"
    return gruppen


def _lade_bestehende_sammelgruppen(csv_pfad):
    """sammel_zuordnung.csv -> Liste noch 'frischer' Gruppen aus einem
    frueheren Lauf (aeltere als BRUECKEN_CSV_MERGE_TAGE werden verworfen,
    siehe dort). Format je Zeile:
        Code;Artikelnr;Bezeichnung;Menge;Rechnungsnummern;EAN;Rechnungsdatum"""
    gruppen = []
    for t in _lies_bestehende_zeilen(csv_pfad, 5):
        datum = t[6] if len(t) > 6 else ""
        if not _bruecken_datum_frisch(datum):
            continue
        rnr = [x.strip() for x in t[4].split(",") if x.strip()] if len(t) > 4 else []
        if not rnr:
            continue
        gruppen.append({
            "code": t[0], "art": t[1], "bez": t[2],
            "ean": t[5] if len(t) > 5 else "", "rnr": rnr, "datum": datum,
        })
    return gruppen


def merge_sammelgruppen(neue_gruppen, csv_pfad):
    """Fuehrt frisch berechnete Sammeldruck-Gruppen (vorlaeufige Codes aus
    finde_sammelgruppen, reine Aufzaehlung INNERHALB dieses Laufs) mit noch
    frischen Gruppen aus einem FRUEHEREN Lauf zusammen - mit der Garantie,
    dass ein bereits vergebener Code NIE auf eine andere Bestellgruppe
    umgebogen wird.

    Hintergrund: Der Sammelcode wird auf Papier gedruckt (Packuebersicht-
    Deckblatt mit den Original-Kundennamen). Scannt jemand einen noch nicht
    abgearbeiteten Code aus einem frueheren Lauf, MUSS er weiterhin exakt
    dieselben Bestellungen ergeben wie auf dem Papier steht - sonst wuerden
    am Ende Label fuer die falschen Bestellungen gedruckt, ohne dass das beim
    Scannen auffaellt (es wird ja keine Adresse manuell gegengeprueft).
    Da archivierte Rechnungen in einem spaeteren Lauf nie wieder auftauchen,
    ueberschneiden sich alte und neue Gruppen nie in ihren Rechnungsnummern -
    das eigentliche Risiko ist rein die Wiederverwendung derselben Codenummer
    fuer eine andere Gruppe, und genau das verhindert diese Funktion.

    Rueckgabe: (fuer_pdf, fuer_csv)
      fuer_pdf: NUR die Gruppen DIESES Laufs, mit finalen kollisionsfreien
                Codes - das geht in die neu gedruckte Pickliste.
      fuer_csv: fuer_pdf PLUS alle noch frischen Alt-Gruppen unveraendert -
                das wird komplett nach sammel_zuordnung.csv geschrieben."""
    alt = _lade_bestehende_sammelgruppen(csv_pfad)
    belegt = set()
    for g in alt:
        m = re.match(r"SAM-(\d+)$", g["code"])
        if m:
            belegt.add(int(m.group(1)))

    naechste = 1
    fuer_pdf = []
    for g in neue_gruppen:
        while naechste in belegt:
            naechste += 1
        neu = dict(g)
        neu["code"] = f"SAM-{naechste}"
        belegt.add(naechste)
        fuer_pdf.append(neu)

    return fuer_pdf, (alt + fuer_pdf)


# ----------------------------------------------------------------------------
# PARSER
# ----------------------------------------------------------------------------

# Numerische Spalten (Menge/Einzelpreis/G-Preis) werden NICHT mehr ueber feste
# x-Baender gelesen, sondern layout-unabhaengig als die (bis zu drei) Komma-
# Dezimalzahlen RECHTS der Bezeichnung, in Lesereihenfolge Menge | EP | GP.
# Damit funktionieren alle Amicron-Spaltenlagen (mit und ohne EAN-Spalte, sowie
# die verschiedenen alten Vorlagen) und gemischte Rechnungen. Eine EAN ist eine
# reine Ziffernfolge OHNE Komma und wird so nie als Preis gezaehlt. Betraege in
# der Bezeichnung (z.B. "3,2mm") liegen links von NUM_MIN_X und stoeren nicht.
NUM_MIN_X = 305        # ab hier (rechts der Einheit) stehen Menge/EP/GP


def _preiszahlen(line):
    """Zahlen einer Zeile rechts der Bezeichnung (Menge/EP/GP), links -> rechts.
    Die Menge kann eine GANZE Zahl ohne Komma sein (z.B. '2'); Einzelpreis und
    G-Preis tragen immer ein Komma. Reine EAN/GTIN-Ziffernfolgen (8 oder 12-14
    Stellen) werden ausgeschlossen. Zahlen der Bezeichnung (z.B. '3,2mm', '8mm',
    '50 Stück') liefern entweder kein num() oder liegen links von NUM_MIN_X."""
    out = []
    for w in sorted(line, key=lambda w: w["x0"]):
        t = w["text"]
        if w["x0"] < NUM_MIN_X or EAN_RE.match(t) or num(t) is None:
            continue
        out.append(w)
    return out


def parse_block(words):
    """Ein Positions-Block (Woerter zwischen zwei Trenner-Punkten) -> dict.
    Layout-unabhaengig: Menge/EP/GP = die Komma-Zahlen rechts der Bezeichnung
    (rechtsbuendig verankert: GP ganz rechts). Artikelnr (x0<ART_MAX) und
    Bezeichnung (BEZ-Band) werden ueber ALLE Blockzeilen gesammelt, weil die
    Bezeichnung mehrzeilig sein kann und die Zahlen dann in einer anderen Zeile
    stehen als die Artikelnummer."""
    lines = cluster_lines(words)
    # Kopf- und Rechnungsbetrag-Zeile gehoeren nicht zur Position (koennen aber im
    # selben Block liegen, wenn die Tabelle knapp getrennt ist) -> raus.
    lines = [ln for ln in lines
             if not any(k in line_text(ln)
                        for k in ("Artikelnr", "Bezeichnung", "Rechnungsbetrag"))]

    # Hauptzeile = Zeile mit den meisten Preis-Zahlen (Menge/EP/GP).
    main, best = None, 0
    for ln in lines:
        n = len(_preiszahlen(ln))
        if n > best:
            main, best = ln, n

    art = einh = ean = ""
    menge = ep = gp = None
    bez_parts, lagerorte = [], []

    # Menge / Einzelpreis / G-Preis aus der Hauptzeile (rechts verankert).
    if main:
        pz = _preiszahlen(main)
        if len(pz) >= 3:
            menge, ep, gp = num(pz[-3]["text"]), num(pz[-2]["text"]), num(pz[-1]["text"])
        elif len(pz) == 2:
            menge, gp = num(pz[-2]["text"]), num(pz[-1]["text"])
        elif len(pz) == 1:
            gp = num(pz[-1]["text"])
        # Einheit (Mittelpunkt im Einheit-Band; EAN-Ziffernfolgen ausgeschlossen)
        einh_toks = [w["text"] for w in main
                     if in_band(w, EINH_LO, EINH_HI) and not EAN_RE.match(w["text"])]
        einh = einh_toks[0] if einh_toks else ""

    # Artikelnr, Bezeichnung, EAN und Lagerort ueber ALLE Zeilen sammeln.
    # Kopfzeile ('Artikelnr Bezeichnung ...') und die 'Lagerort:'-Zeile werden
    # dabei uebersprungen. Die Artikelnr wird NUR aus der ersten Positionszeile
    # gezogen (Folgezeilen einer umbrochenen Bezeichnung koennen links ebenfalls
    # etwas haben, z.B. '10 Stück' -> darf nicht in die Artnr wandern).
    art_toks = []
    art_done = False
    i = 0
    while i < len(lines):
        ln = lines[i]
        txt = line_text(ln)
        if "Artikelnr" in txt or "Bezeichnung" in txt:      # Tabellenkopf
            i += 1
            continue
        if "Lagerort" in txt:
            # VOLLSTAENDIGEN Lagerort-Wert ziehen (nicht nur das erste Token),
            # damit mehrteilige Angaben wie 'R8 E2' oder 'A8 + A9' erhalten
            # bleiben. Wichtig fuer die Fachanzeige UND die Einsortierung: bei
            # 'R8 E2' bringt erst das 'E2' die Bestellung aufs Kleinteile-Blatt
            # (classify -> GRID_RE); 'R8' allein liefe auf 'Allgemein'.
            wert = _lagerort_wert(txt)
            if not wert and i + 1 < len(lines):
                # 'Lagerort:'-Zeile ohne Wert -> der Wert steht umbrochen in der
                # Folgezeile. Diese nur uebernehmen, wenn sie KEINE Positions-/
                # Preiszeile ist (keine Menge/EP/GP-Zahlen); die Folgezeile wird
                # dann mitverbraucht, damit sie nicht als Bezeichnung/Artnr
                # fehlinterpretiert wird.
                nxt = lines[i + 1]
                if not _preiszahlen(nxt):
                    wert = line_text(nxt).strip()
                    i += 1
            if wert:
                lagerorte.append(wert)
            i += 1
            continue
        links = [w for w in sorted(ln, key=lambda w: w["x0"])
                 if w["x0"] < ART_MAX and w["text"] != "."]
        if links and not art_done:                          # erste Positionszeile
            art_toks = [w["text"] for w in links]
            art_done = True
        for w in sorted(ln, key=lambda w: w["x0"]):
            t = w["text"]
            if t == "." or w["x0"] < ART_MAX:
                continue
            if in_band(w, BEZ_LO, BEZ_HI):                  # Bezeichnungs-Spalte
                bez_parts.append((round(w["top"]), w["x0"], t))
            elif EAN_RE.match(t) and w["x0"] >= EINH_LO and not ean:  # reine EAN/GTIN
                ean = t
        i += 1

    art = "".join(art_toks)
    bez = " ".join(t for _, _, t in sorted(bez_parts)).strip()

    return {
        "art": art,
        "bez": bez,
        "einh": einh,
        "ean": ean,
        "menge": menge,
        "ep": ep,
        "gp": gp,
        "lagerorte": lagerorte,
    }


def extrahiere_adresse(words):
    """Lieferadresse -> (name, strasse, hausnr, plz)."""
    send = next((w["top"] for w in words if "Gasecenter" in w["text"]), 640)
    # Untergrenze des Adressfensters. Die kompaktere Rechnungsvorlage schiebt den
    # Lieferadressblock weiter nach unten (PLZ-Zeile jetzt bei top~775); die frueheren
    # 760 schnitten Strasse+PLZ der Lieferadresse ab (-> leere post_zuordnung.csv,
    # Deutsche-Post-Zuordnung per PLZ+Hausnummer scheiterte). 800 deckt den Block ab;
    # unterhalb 776 steht in den Rechnungen kein Inhalt (Seitenhoehe 842).
    ADR_BOT = 800
    left = [w for w in words if w["x0"] < 405 and send + 5 < w["top"] < ADR_BOT]
    right = [w for w in words if w["x0"] >= 405 and send + 5 < w["top"] < ADR_BOT]
    lefts = [line_text(ln) for ln in cluster_lines(left)]
    rights = [line_text(ln) for ln in cluster_lines(right)]
    right_real = [l for l in rights if "Lieferadresse" not in l and "Rechnungsadresse" not in l]
    same = any("Rechnungsadresse" in l for l in rights)
    deliv = lefts if (same or not right_real) else right_real
    deliv = [l for l in deliv if l.strip()]
    if not deliv:
        return "", "", "", ""
    name = deliv[0]
    # Bei abweichender Lieferadresse (Bestellung 1695459, Amazon "Bäck-Kuhn"/
    # "AT-8380 Jennersdorf"): das Länderkürzel "AT-" steht oft NUR in der
    # Rechnungsadresse ("KUHN Bernhard"/"AT-8380"), die separate Lieferadresse
    # nennt nur "8380 Jennersdorf" ohne AT-Praefix und ohne "ÖSTERREICH" ->
    # Land wurde faelschlich als DE erkannt, die 4-stellige PLZ passte dann
    # nicht auf die DE-Regel (5-stellig) und blieb leer (keine Zuordnung,
    # kein Druck). Daher zusaetzlich lefts (Rechnungsadresse) auswerten.
    land = _land_aus_zeilen(deliv)
    if land != "AT" and not any(re.search(r"\b\d{5}\b", l) for l in deliv):
        # Lieferadresse liefert weder ein AT-Signal noch eine gueltige
        # 5-stellige PLZ -> evtl. fehlt nur das AT-Praefix in der separaten
        # Lieferadresse (Fall 1695459) -> zusaetzlich Rechnungsadresse pruefen.
        land = _land_aus_zeilen(deliv + lefts)
    if land == "AT":
        # Oesterreich: 4-stellige PLZ ('AT-8020', 'A-8020' oder '8020 Ort'); die
        # Zeile 'ÖSTERREICH' steht meist darunter und ist nicht die PLZ-Zeile.
        # Rueckwaerts durchsuchen (wie der DE-Zweig unten) statt vorwaerts: bei
        # Amazon-Business-Bestellungen mit angehaengter PO-Referenz in der
        # Namenszeile (z.B. "GmbHPO126-0287 VDS Getriebe") passt "0287" bereits
        # auf das Muster "4 Ziffern + Leerzeichen + Text" und wurde faelschlich
        # als PLZ genommen, bevor die echte PLZ-Zeile ueberhaupt geprueft wurde
        # (Rechnung 1700723, VDS Getriebe/Wolfern: PLZ "0287" statt "4493").
        plzline = next((l for l in reversed(deliv)
                        if re.search(r"\b(?:AT|A)-?\s?\d{4}\b", l)
                        or re.search(r"\b\d{4}\b\s+\S", l)), "")
    else:
        plzline = next((l for l in reversed(deliv) if re.search(r"\b\d{5}\b", l)), "")
    idx = deliv.index(plzline) if plzline in deliv else len(deliv) - 1
    street = deliv[idx - 1] if idx - 1 >= 1 else (deliv[1] if len(deliv) > 1 else "")
    return name, street, _hausnr(street), _plz(plzline, land)


def parse_pdf(path):
    """Eine Rechnung -> dict mit Kopf, Positionen, Adresse, Abgleich-Status."""
    with pdfplumber.open(path) as pdf:
        page = pdf.pages[0]
        words = page.extract_words(use_text_flow=False)

    tops = sorted({round(w["top"]) for w in words})
    zeilen = [line_text([w for w in words if round(w["top"]) == t]) for t in tops]
    full = "\n".join(zeilen)

    m = re.search(r"Rechnung\s*Nr\.?\s*:?\s*(\d+)", full)
    rnr = m.group(1) if m else ""

    # Kd-Nr. = Zeile direkt nach der Rechnungsnummer-Zeile
    kdnr = ""
    for i, z in enumerate(zeilen):
        if re.search(r"Rechnung\s*Nr", z):
            for j in range(i + 1, min(i + 3, len(zeilen))):
                cand = zeilen[j].strip()
                if cand and "@" not in cand:
                    kdnr = cand
                    break
            break

    md = re.search(r"(\d{2}\.\d{2}\.\d{4})", full)
    datum = md.group(1) if md else ""

    # Positionen ueber Trenner-Punkte (einzelnes '.' am linken Rand) zerlegen.
    # Toleranter x-Bereich (20..40), da die Punkte je nach Vorlage bei ~28 oder
    # ~32 sitzen (z.B. 27.99 wurde vom alten Fenster 28..38 knapp verfehlt).
    def _ist_trenner(w):
        return w["text"] == "." and 20 <= w["x0"] <= 40

    # Tabellenende dynamisch bestimmen: Die Amicron-Seite wiederholt UNTER der
    # Positionstabelle einen 'Druckdatum:'-Block (Lieferschein-/Adressteil). Bis
    # kurz davor reicht die Tabelle. Fixe 388 waren zu niedrig -> bei mehreren
    # Artikeln / langen Bezeichnungen fielen Versand oder ganze Positionen weg,
    # weil deren Trennpunkte unterhalb 388 lagen.
    kopf_top = min((w["top"] for w in words if w["text"] == "Artikelnr"),
                   default=TABLE_TOP)
    wiederhol = [w["top"] for w in words
                 if w["text"].startswith("Druckdatum") and w["top"] > kopf_top + 5]
    if wiederhol:
        table_bot = min(wiederhol) - 2
    else:
        rb = [w["top"] for w in words if "Rechnungsbetrag" in w["text"]]
        table_bot = (max(rb) + 40) if rb else TABLE_BOT

    tab = [w for w in words if TABLE_TOP < w["top"] < table_bot]
    seps = sorted(w["top"] for w in tab if _ist_trenner(w))
    positionen = []
    lo = TABLE_TOP
    for s in seps:
        block = [w for w in tab if lo < w["top"] < s and not _ist_trenner(w)]
        if block:
            p = parse_block(block)
            if p["gp"] is not None or p["art"] or p["bez"]:
                positionen.append(p)
        lo = s

    mb = re.search(r"Rechnungsbetrag\s*€?\s*([\d.,]+)", full)
    total = num(mb.group(1)) if mb else None

    name, strasse, hausnr, plz = extrahiere_adresse(words)

    # Betragsabgleich
    zeilen_ok = True
    for p in positionen:
        if p["menge"] is not None and p["ep"] is not None and p["gp"] is not None:
            if abs(round(p["menge"] * p["ep"], 2) - p["gp"]) > CENT:
                zeilen_ok = False
    summe = round(sum(p["gp"] for p in positionen if p["gp"] is not None), 2)
    # Ein NICHT lesbarer Rechnungsbetrag (z.B. wenn eine sehr lange Bezeichnung ihn
    # im Quell-PDF ueberlagert) darf keinen falschen "Summe != 0,00"-Fehler ausloesen.
    total_lesbar = total is not None
    summe_ok = (not total_lesbar) or (abs(summe - total) <= CENT)

    # Vollstaendigkeitspruefung: jede ECHTE Position braucht Artikelnr + Menge +
    # Einzelpreis + Betrag(GP). Versandzeilen (ohne Artikel) brauchen Menge + Betrag.
    vollstaendig_ok = True
    unvollstaendig = []
    for p in positionen:
        if ist_versand(p["art"], p["bez"]):
            # Versand braucht nur den Betrag (Menge ist stets 1 und geht nicht in
            # die Summe ein). So warnt eine im Quell-PDF verschmolzene Versand-Menge
            # nicht unnoetig.
            fehlt = []
            if p["gp"] is None:
                fehlt.append("Betrag")
        else:
            fehlt = []
            if not p["art"]:
                fehlt.append("Artikelnr")
            if p["menge"] is None:
                fehlt.append("Menge")
            if p["ep"] is None:
                fehlt.append("Einzelpreis")
            if p["gp"] is None:
                fehlt.append("Betrag")
        if fehlt:
            vollstaendig_ok = False
            kennung = p["art"] or (p["bez"][:20] if p["bez"] else "?")
            unvollstaendig.append(f"{kennung} ({'/'.join(fehlt)})")

    return {
        "datei": os.path.basename(path),
        "rnr": rnr, "kdnr": kdnr, "datum": datum,
        "positionen": positionen, "total": total, "summe": summe,
        "name": name, "strasse": strasse, "hausnr": hausnr, "plz": plz,
        "zeilen_ok": zeilen_ok, "summe_ok": summe_ok,
        "total_lesbar": total_lesbar,
        "vollstaendig_ok": vollstaendig_ok, "unvollstaendig": unvollstaendig,
    }


# ----------------------------------------------------------------------------
# PDF-AUSGABE
# ----------------------------------------------------------------------------

class CheckBox(Flowable):
    """Kleines leeres Abhak-Kaestchen."""
    def __init__(self, size=10):
        super().__init__()
        self.size = size
        self.width = size
        self.height = size

    def draw(self):
        self.canv.setLineWidth(0.8)
        self.canv.rect(0, 1, self.size, self.size)


class MengeFlow(Flowable):
    """Menge-Anzeige. Bei Menge > 1: rot, fett, GROESSER und eingekreist.
    Die hervorgehobene Menge wird zusaetzlich um `HERVOR_PLUS` Punkt groesser
    gesetzt, damit ein Mehrfach-Artikel wirklich sofort ins Auge faellt."""
    HERVOR_PLUS = 6   # so viele pt groesser als die normale Schrift bei Menge>1

    def __init__(self, text, hervor=False, fs=9):
        super().__init__()
        self.text = text or ""
        self.hervor = hervor
        self.fs = (fs + self.HERVOR_PLUS) if hervor else fs

    def wrap(self, aw, ah):
        self.aw = aw
        self.height = self.fs + (12 if self.hervor else 4)
        return aw, self.height

    def draw(self):
        c = self.canv
        font = "Helvetica-Bold" if self.hervor else "Helvetica"
        c.setFont(font, self.fs)
        tw = c.stringWidth(self.text, font, self.fs)
        x = 9 if self.hervor else 1
        y = (self.height - self.fs) / 2.0 + 1
        if self.hervor:
            c.setFillColor(colors.red)
            c.drawString(x, y, self.text)
            cx = x + tw / 2.0
            cy = y + self.fs * 0.33
            hw = tw / 2.0 + 6
            hh = self.fs / 2.0 + 5
            c.setStrokeColor(colors.red)
            c.setLineWidth(1.6)
            c.ellipse(cx - hw, cy - hh, cx + hw, cy + hh)
        else:
            c.setFillColor(colors.black)
            c.drawString(x, y, self.text)


def mengentext(menge, einh):
    if menge is None:
        return ""
    if abs(menge - round(menge)) < 1e-9:
        m = str(int(round(menge)))
    else:
        m = ("%.2f" % menge).replace(".", ",")
    return f"{m} {einh}".strip()


# ---------------------------------------------------------------------------
# Gewichtsartikel (Schweissdraht/-staebe u.ae., Verkauf nach kg)
# ---------------------------------------------------------------------------
# Erkennung: im Bezeichnungstext steht ein kg-Token (z.B. "3kg", "0,5 kg",
# "1,00 kg") UND dieser Wert ist die LETZTE Zahl der Artikelnummer
# (z.B. V2A-308L-1,6-3,0 -> 3,0  /  WSG2-1,6-0,5 -> 0,5). Die Doppelpruefung
# verhindert Fehlausloeser bei Artikeln, die kg nur als Packungsgroesse im Text
# fuehren (z.B. "Chlorgranulat 5kg"), deren Artikelnummer aber nicht auf 5 endet.
GEWICHT_RE = re.compile(r"(\d+(?:[.,]\d+)?)\s*kg\b", re.IGNORECASE)


def artikel_gewicht(art, bez):
    """Stueckgewicht in kg, wenn die Position ein nach Gewicht verkaufter Artikel
    ist (kg-Token im Text passt zur letzten Zahl der Artikelnummer), sonst None."""
    m = GEWICHT_RE.search(bez or "")
    if not m:
        return None
    w = float(m.group(1).replace(",", "."))
    letzte = (art or "").split("-")[-1].strip()
    # Nur ein KOMMA-Dezimalwert als letztes Segment gilt als Gewichtsangabe
    # (Katalog-Konvention: 0,5 / 1,0 / 3,0). So triggert eine reine Modell-/
    # Groessennummer wie 'AutoPR-11' (11 == '11kg' im Text = fuer 11-kg-Flaschen)
    # NICHT faelschlich als Gewichtsartikel.
    if not re.fullmatch(r"\d+,\d+", letzte):
        return None
    try:
        if abs(float(letzte.replace(",", ".")) - w) < 1e-6:
            return w
    except ValueError:
        pass
    return None


# ---------------------------------------------------------------------------
# Packungsartikel (feste Stueck-Packung, z.B. Schlauchverbinder 2248/10)
# ---------------------------------------------------------------------------
# Erkennung analog zu den Gewichtsartikeln ueber eine Doppelpruefung: die
# Artikelnummer hat die Form "<Zahl>/<N>" (rein numerisch, z.B. 2248/10) UND im
# Bezeichnungstext steht "<N> Stück". Nur dann ist N die Packungsgroesse. Das
# verhindert Fehlausloeser bei Artikeln, deren "/N" etwas anderes bedeutet
# (z.B. 52107/4 "4x ...", 299/3 "3x ..." -> kein "<N> Stück" im Text).
PACK_ART_RE = re.compile(r"^(\d+)/(\d+)$")


def _pack_token_re(n):
    """Regex fuer ein '<n> Stück'-Token im Text (Stück/Stueck/Stck/Stk, Sg./Pl.)."""
    return re.compile(rf"(?<!\d){n}\s*St(?:ü|ue)?c?ke?\b", re.IGNORECASE)


def artikel_packung(art, bez):
    """Packungsgroesse (Stueck je Packung), wenn die Position ein in fester
    Stueck-Packung verkaufter Artikel ist: Artikelnummer '<Zahl>/<N>' UND '<N>
    Stück' im Text. Sonst None."""
    m = PACK_ART_RE.match((art or "").strip())
    if not m:
        return None
    n = int(m.group(2))
    if n <= 1:
        return None
    if _pack_token_re(n).search(bez or ""):
        return n
    return None


def pack_anzeige(p):
    """Anzeige (bez, mengentext, hervorheben) fuer die PACKUEBERSICHT.
    Gewichtsartikel: kg-Token im Text durch 'GESAMTMENGE' ersetzen, Menge =
    Stueckgewicht x bestellte Menge, Einheit kg. Packungsartikel (<Zahl>/<N> mit
    '<N> Stück' im Text): '<N> Stück' -> 'GESAMTMENGE', Menge = N x bestellte
    Menge in Stück. Beide immer hervorgehoben. Sonst unveraendert (Hervorhebung
    wie bisher bei Menge > 1)."""
    qty = p["menge"] if p["menge"] is not None else 1
    gw = artikel_gewicht(p["art"], p["bez"])
    if gw is not None:
        bez = GEWICHT_RE.sub("GESAMTMENGE", p["bez"], count=1)
        return (bez, mengentext(gw * qty, "kg"), True)
    pk = artikel_packung(p["art"], p["bez"])
    if pk is not None:
        bez = _pack_token_re(pk).sub("GESAMTMENGE", p["bez"], count=1)
        return (bez, mengentext(pk * qty, "Stück"), True)
    return (p["bez"], mengentext(p["menge"], p["einh"]),
            (p["menge"] or 0) > 1.0000001)


def scan_bedarf(eff):
    """Noetige Scans fuer EINE Position mit effektiver Menge eff:
    genau 1 -> 1, unter 1 (z.B. 0,5 kg) -> 2, sonst aufgerundet (gedeckelt 5)."""
    if eff is None or eff <= 0:
        return 1
    if abs(eff - 1.0) < 1e-9:
        return 1
    if eff < 1:
        return 2
    e = round(eff, 6)
    ganz = int(e) + (1 if e > int(e) else 0)
    return min(5, ganz)


def order_scananzahl(r):
    """Noetige Scans fuer eine ganze Bestellung = Summe der Positions-Scans,
    gedeckelt auf 5 (mind. 1). Gewichtsartikel zaehlen nach kg (Gewicht x Menge)."""
    total = 0
    for p in r["positionen"]:
        if ist_versand(p["art"], p["bez"]):
            continue
        gw = artikel_gewicht(p["art"], p["bez"])
        qty = p["menge"] if p["menge"] is not None else 1
        eff = (gw * qty) if gw is not None else qty
        total += scan_bedarf(eff)
    return max(1, min(5, total))


def _artikel_sortkey(s):
    """Natuerliche Sortierung: Zahlen numerisch (auf 8 Stellen aufgefuellt)."""
    return re.sub(r"\d+", lambda m: m.group().zfill(8), (s or "").lower())


def rahmen_block(flowables, breite, dashed=False):
    """Flowables in einen umrahmten 1x1-Tabellenblock packen (Rahmen optional
    gestrichelt)."""
    t = Table([[flowables]], colWidths=[breite])
    box = (("BOX", (0, 0), (-1, -1), 1.0, colors.HexColor("#555555"), None, (3, 2))
           if dashed else ("BOX", (0, 0), (-1, -1), 0.9, colors.HexColor("#37474F")))
    t.setStyle(TableStyle([
        box,
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
        ("RIGHTPADDING", (0, 0), (-1, -1), 8),
        ("TOPPADDING", (0, 0), (-1, -1), 7),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
        ("BACKGROUND", (0, 0), (-1, -1), colors.white),
    ]))
    return t


def _fuss_zeichnen(canvas, doc, text):
    """Kleiner grauer Fuss unten links auf jeder Seite (z.B. Erstellungsdatum)."""
    canvas.saveState()
    canvas.setFont("Helvetica", 7)
    canvas.setFillGray(0.5)
    canvas.drawString(14 * mm, 7 * mm, text)
    canvas.restoreState()


def _story_to_reader(story, fusstext=None):
    """Eine Flowable-Liste zu einem PdfReader rendern (fuer seitengenaues
    Zusammenfuehren). fusstext (optional) wird auf JEDER Seite unten abgedruckt."""
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=14 * mm, rightMargin=14 * mm,
                            topMargin=12 * mm, bottomMargin=12 * mm)
    if fusstext:
        cb = lambda c, d: _fuss_zeichnen(c, d, fusstext)
        doc.build(story, onFirstPage=cb, onLaterPages=cb)
    else:
        doc.build(story)
    buf.seek(0)
    return PdfReader(buf)


def baue_pdf(rechnungen, pdf_pfad, gruppen):
    styles = getSampleStyleSheet()
    st_titel = ParagraphStyle("titel", parent=styles["Title"], fontSize=16, spaceAfter=4)
    st_cell = ParagraphStyle("cell", parent=styles["Normal"], fontSize=9, leading=11)
    st_klein = ParagraphStyle("klein", parent=styles["Normal"], fontSize=8,
                              textColor=colors.grey)
    st_kopf = ParagraphStyle("kopf", parent=styles["Normal"], fontSize=10, leading=13)
    # Eigener Kopf-Stil fuer die Packuebersicht: grosszuegige Zeilenhoehe, damit
    # der vergroesserte Empfaengername (inline <font size=...>) nicht ueberlappt.
    st_pack_kopf = ParagraphStyle("packkopf", parent=styles["Normal"], fontSize=10, leading=18)
    st_warn = ParagraphStyle("warn", parent=styles["Normal"], fontSize=9,
                            textColor=colors.red)
    st_info = ParagraphStyle("info", parent=styles["Normal"], fontSize=8,
                            textColor=colors.HexColor("#7a6000"))

    SEITE = A4[0] - 28 * mm            # nutzbare Breite (= doc.width)
    anz_rechnungen = len(rechnungen)
    # Kleiner Nachlauf -> alles auf eine gemeinsame Liste statt auf mehrere
    # Lagerort-Kategorien zu verteilen (siehe KOMBINIERT_SCHWELLE oben).
    kombiniert = anz_rechnungen < KOMBINIERT_SCHWELLE
    kategorien = [KOMBINIERT_LABEL] if kombiniert else KATEGORIE_REIHENFOLGE

    def kat_von(r):
        return KOMBINIERT_LABEL if kombiniert else auftrag_kategorie(r)

    # Kopf-Tabellenstil fuer Listen
    def _liste_style():
        return TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#37474F")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONTSIZE", (0, 0), (-1, 0), 9),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F2F4F5")]),
            ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#CFD8DC")),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("ALIGN", (-1, 0), (-1, -1), "CENTER"),     # Kaestchen-Spalte zentriert
            ("LEFTPADDING", (0, 0), (-1, -1), 4),
            ("RIGHTPADDING", (0, 0), (-1, -1), 4),
            ("TOPPADDING", (0, 0), (-1, -1), 3),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ])

    # ----- Aggregation Teil A: Kategorie -> art -> {bez, einh, lager, menge} ---
    # Die Kommissionierliste eines Lagerorts listet ALLE Positionen der
    # Bestellungen, die DORT gepackt werden (Haupt-Lagerort = auftrag_kategorie),
    # nicht nur die dort gelagerten Artikel. So steht z.B. ein Wapo-Artikel, der
    # zusammen mit Topseller/Poolchemie verschickt wird, auf DEREN Liste - er muss
    # ja mitverpackt werden. Die Spalte "Lagerort" zeigt, wo der Artikel liegt.
    # Dadurch sind Kommissionierliste (Teil A) und Packuebersicht (Teil B)
    # konsistent, und jeder Artikel steht genau einmal (kein Doppel-Picken).
    kat_artikel = {k: {} for k in kategorien}
    for r in rechnungen:
        kat = kat_von(r)
        for p in r["positionen"]:
            if ist_versand(p["art"], p["bez"]):
                continue
            key = p["art"] or p["bez"]
            eintrag = kat_artikel[kat].setdefault(
                key, {"art": p["art"], "bez": p["bez"], "einh": p["einh"],
                      "lager": set(), "menge": 0.0})
            eintrag["menge"] += (p["menge"] or 0)
            for lo in p["lagerorte"]:
                eintrag["lager"].add(lo)

    def liste_story(artikel, titel, untertitel):
        """Eine (Kategorie-)Liste als Flowables: Kaestchen RECHTS, nach
        Artikelnummer sortiert, Menge>1 rot/fett/eingekreist."""
        story = [Paragraph(titel, st_titel)]
        if untertitel:
            story.append(Paragraph(untertitel, st_klein))
        story.append(Spacer(1, 6))
        head = ["Lagerort", "Artikelnr", "Bezeichnung", "Menge", ""]
        lw, aw, mw, cw = 66, 80, 80, 26
        bw = SEITE - lw - aw - mw - cw
        data = [head]
        for key in sorted(artikel, key=lambda k: _artikel_sortkey(artikel[k]["art"] or k)):
            a = artikel[key]
            lager = ", ".join(sorted(a["lager"])) if a["lager"] else "-"
            hv = (a["menge"] or 0) > 1.0000001
            data.append([
                Paragraph(lager, st_cell),
                Paragraph(a["art"] or "-", st_cell),
                Paragraph(a["bez"], st_cell),
                MengeFlow(mengentext(a["menge"], a["einh"]), hv),
                CheckBox(),
            ])
        t = Table(data, colWidths=[lw, aw, bw, mw, cw], repeatRows=1)
        t.setStyle(_liste_style())
        story.append(t)
        return story

    def einzel_block(r):
        """Eine Bestellung als solider Rahmen (Barcode, Empfaenger, Positionen).
        Ohne Kundennummer; Kaestchen rechts; Menge>1 hervorgehoben."""
        inner = SEITE - 16
        pos = [p for p in r["positionen"] if not ist_versand(p["art"], p["bez"])]
        try:
            bc = code128.Code128(r["rnr"], barHeight=14 * mm, barWidth=0.42 * mm)
        except Exception:
            bc = Paragraph(r["rnr"], st_kopf)
        kopf_rechts = (
            f"<font size=8>Rechnung Nr. {r['rnr']} &nbsp;&middot;&nbsp; {r['datum']}</font>"
            f"<br/><font size=15><b>{r['name']}</b></font>"
        )
        kopf = Table([[bc, Paragraph(kopf_rechts, st_pack_kopf)]],
                     colWidths=[150, inner - 150])
        kopf.setStyle(TableStyle([
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("LEFTPADDING", (0, 0), (0, 0), 0),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
        ]))
        inhalt = [kopf]
        if not (r["zeilen_ok"] and r["summe_ok"] and r.get("vollstaendig_ok", True)):
            details = []
            if not r["zeilen_ok"]:
                details.append("Zeile Menge×EP≠GP")
            if not r["summe_ok"]:
                details.append(f"Summe {('%.2f' % r['summe']).replace('.', ',')} ≠ "
                                f"Rechnungsbetrag {('%.2f' % (r['total'] or 0)).replace('.', ',')}")
            if not r.get("vollstaendig_ok", True):
                details.append("Position unvollständig: "
                               + ", ".join(r.get("unvollstaendig", [])))
            inhalt.append(Paragraph("⚠ Prüfung NICHT bestanden: " + "; ".join(details), st_warn))
        elif not r.get("total_lesbar", True):
            inhalt.append(Paragraph(
                "ℹ Rechnungsbetrag im PDF nicht maschinell lesbar (vermutlich sehr "
                "lange Bezeichnung) – Positionssumme "
                + ('%.2f' % r['summe']).replace('.', ',') + " €, Positionen geprüft",
                st_info))

        head2 = ["Artikelnr", "Bezeichnung", "Menge", "Lagerort", ""]
        aw, mw, lw, cw = 74, 80, 92, 26
        bw = inner - aw - mw - lw - cw
        data = [head2]
        for p in pos:
            lager = ", ".join(p["lagerorte"]) if p["lagerorte"] else "-"
            bez_disp, menge_disp, hv = pack_anzeige(p)
            data.append([
                Paragraph(p["art"] or "-", st_cell),
                Paragraph(bez_disp, st_cell),
                MengeFlow(menge_disp, hv),
                Paragraph(lager, st_cell),
                CheckBox(),
            ])
        t = Table(data, colWidths=[aw, bw, mw, lw, cw], repeatRows=1)
        t.setStyle(_liste_style())
        inhalt.append(Spacer(1, 4))
        inhalt.append(t)
        return KeepTogether([rahmen_block(inhalt, SEITE), Spacer(1, 8)])

    def gruppe_block(g, gruppe_rechnungen):
        """Sammeldruck-Gruppe: alle betroffenen Bestellungen in einem
        GESTRICHELTEN Rahmen, der Sammel-Barcode unten INNERHALB des Rahmens."""
        n = len(gruppe_rechnungen)
        titel = Paragraph(
            f"<b>Sammeldruck {g['code']}</b> &middot; Artikel <b>{g['art']}</b> &ndash; "
            f"{g['bez']} &middot; {n} Bestellungen à 1 Stück", st_kopf)
        try:
            bc = code128.Code128(g["code"], barHeight=18 * mm, barWidth=0.5 * mm)
        except Exception:
            bc = Paragraph(g["code"], st_kopf)
        hint = Paragraph(
            f"<para align='center'>Barcode <b>{g['code']}</b> scannen &rarr; "
            f"alle {n} Versandlabel auf einmal drucken</para>", st_klein)
        c_rnr, c_plz, c_chk = 90, 60, 26
        c_emp = SEITE - c_rnr - c_plz - c_chk
        # Reihenfolge im Rahmen: Titel -> Deckblatt (Barcode) -> Hinweis ->
        # dann erst die Auftragsliste. (Deckblatt VOR den Auftraegen.)
        data = [[titel, "", "", ""],
                [bc, "", "", ""],
                [hint, "", "", ""],
                [Paragraph("<b>Rechnung</b>", st_cell), Paragraph("<b>Empfänger</b>", st_cell),
                 Paragraph("<b>PLZ</b>", st_cell), ""]]
        kopf_rows = 4                 # Titel, Barcode, Hinweis, Tabellenkopf
        for r in gruppe_rechnungen:
            data.append([Paragraph(r["rnr"], st_cell), Paragraph(r["name"], st_cell),
                         Paragraph(r["plz"], st_cell), CheckBox()])
        last = len(data) - 1
        t = Table(data, colWidths=[c_rnr, c_emp, c_plz, c_chk], repeatRows=kopf_rows)
        t.setStyle(TableStyle([
            ("BOX", (0, 0), (-1, -1), 1.1, colors.HexColor("#444444"), None, (3, 2)),
            ("SPAN", (0, 0), (-1, 0)),                 # Titel
            ("SPAN", (0, 1), (-1, 1)),                 # Barcode (Deckblatt)
            ("SPAN", (0, 2), (-1, 2)),                 # Hinweis
            ("ALIGN", (0, 1), (-1, 2), "CENTER"),      # Barcode + Hinweis zentriert
            ("LINEBELOW", (0, 3), (-1, 3), 0.5, colors.HexColor("#999999")),
            ("ALIGN", (-1, kopf_rows), (-1, last), "CENTER"),   # Kaestchen-Spalte
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING", (0, 0), (-1, -1), 4),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ("LEFTPADDING", (0, 0), (-1, -1), 6),
            ("RIGHTPADDING", (0, 0), (-1, -1), 6),
        ]))
        # Eine Sammeldruck-Gruppe nie ueber einen Seitenumbruch trennen.
        return KeepTogether([t])

    # ============================ PDF zusammenbauen =========================
    writer = PdfWriter()

    # Erstellungszeitpunkt einmal festhalten -> auf jeder Seite als Fuss.
    erstellt = datetime.now()
    fusstext = f"Erstellt am {erstellt:%d.%m.%Y} um {erstellt:%H:%M} Uhr"

    def add_story(story, pad_even=False):
        reader = _story_to_reader(story, fusstext)
        for p in reader.pages:
            writer.add_page(p)
        if pad_even and len(reader.pages) % 2 == 1:
            writer.add_blank_page(width=float(A4[0]), height=float(A4[1]))

    # ---- Bestellungen den Lagerorten zuordnen --------------------------------
    # Jede Bestellung kommt KOMPLETT unter genau EINEN Lagerort (Haupt-Lagerort
    # = hoechste Prioritaet ueber ihre Positionen, siehe auftrag_kategorie()).
    # Sammeldruck-Gruppen (immer ein einzelner Artikel) ordnen sich ueber diesen
    # Artikel demselben Lagerort zu.
    # Die Kommissionierliste je Lagerort listet GENAU die Positionen dieser
    # Bestellungen (siehe kat_artikel oben) - also alles, was an diesem Packplatz
    # verpackt wird, inkl. Artikeln, die woanders gelagert sind (Spalte Lagerort).
    grouped_rnr = set(rnr for g in gruppen for rnr in g["rnr"])
    rnr_map = {r["rnr"]: r for r in rechnungen}
    einzel = [r for r in rechnungen if r["rnr"] not in grouped_rnr]

    einzel_nach_kat = {k: [] for k in kategorien}
    for r in einzel:
        einzel_nach_kat[kat_von(r)].append(r)

    gruppen_nach_kat = {k: [] for k in kategorien}
    for g in gruppen:
        gr = [rnr_map[x] for x in g["rnr"] if x in rnr_map]
        kat = kat_von(gr[0]) if gr else (KOMBINIERT_LABEL if kombiniert else "Allgemein")
        gruppen_nach_kat[kat].append((g, gr))

    def packuebersicht_flowables(einzel_k, gruppen_k):
        """Packuebersicht-Teil EINES Lagerorts: erst die Einzel-Bestellungen,
        dann die Sammeldruck-Gruppen. Jeder Block dank KeepTogether ungebrochen.
        Ohne zugeordnete Bestellung -> kurzer Hinweis."""
        fl = [Spacer(1, 10),
              Paragraph("Packübersicht je Rechnung", st_titel),
              Paragraph("Eine Bestellung je Rahmen &middot; Versandkosten ausgeblendet "
                        "&middot; Sammeldruck-Gruppen gestrichelt umrahmt", st_klein),
              Spacer(1, 6)]
        if not einzel_k and not gruppen_k:
            fl.append(Paragraph(
                "Keine Bestellung hat diesen Lagerort als Haupt-Lagerort &ndash; die "
                "oben gelisteten Artikel gehören zu Bestellungen, die bei deren "
                "Haupt-Lagerort gepackt werden.", st_klein))
            return fl
        for r in einzel_k:
            fl.append(einzel_block(r))
        for g, gr in gruppen_k:
            fl.append(gruppe_block(g, gr))
            fl.append(Spacer(1, 10))
        return fl

    # ---- Pro Lagerort: Kommissionierliste + DIREKT die Packuebersicht --------
    # Jeder Lagerort beginnt auf einem NEUEN Blatt (eigene Handout-Liste je
    # Packplatz). Nur Lagerorte mit Inhalt erscheinen.
    story = []
    erste = True
    for kat in kategorien:
        artikel = kat_artikel[kat]
        einzel_k = einzel_nach_kat[kat]
        gruppen_k = gruppen_nach_kat[kat]
        if not artikel and not einzel_k and not gruppen_k:
            continue
        if not erste:
            story.append(PageBreak())
        erste = False
        anz_best = len(einzel_k) + sum(len(gr) for _g, gr in gruppen_k)
        titel = "Kommissionierliste" if kombiniert else f"Kommissionierliste &ndash; {kat}"
        story += liste_story(
            artikel, titel,
            f"{len(artikel)} Artikel zu kommissionieren &nbsp;&middot;&nbsp; "
            f"{anz_best} Bestellung(en) zum Packen &nbsp;&middot;&nbsp; "
            f"sortiert nach Artikelnummer")
        story += packuebersicht_flowables(einzel_k, gruppen_k)

    if not story:
        story = [Paragraph("Keine auswertbaren Positionen.", st_titel)]
    add_story(story, pad_even=False)

    try:
        with open(pdf_pfad, "wb") as f:
            writer.write(f)
    except PermissionError:
        raise SystemExit(
            "\n*** FEHLER: Die Pickliste-PDF kann nicht geschrieben werden:\n"
            f"    {pdf_pfad}\n"
            "    Die Datei ist vermutlich noch GEOEFFNET (im PDF-Betrachter) oder\n"
            "    schreibgeschuetzt. Bitte das Pickliste-Fenster schliessen und die\n"
            "    Batch erneut starten. (Es wurde nichts kopiert/archiviert.)\n")


# ----------------------------------------------------------------------------
# post_zuordnung.csv
# ----------------------------------------------------------------------------

# ----------------------------------------------------------------------------
# post_zuordnung.csv / mengen_zuordnung.csv / ean_zuordnung.csv
# ----------------------------------------------------------------------------
# Diese drei Dateien werden bei JEDEM Lauf mit einer eventuell BESTEHENDEN
# Datei ZUSAMMENGEFUEHRT statt sie zu ueberschreiben. Grund: nach jedem Lauf
# werden die eingelesenen Rechnungs-PDFs archiviert (verschwinden aus dem
# Eingangsordner, siehe archiviere()) - ein zweiter Lauf saehe sonst nur noch
# die NEUEN Rechnungen, und die Zuordnung fuer noch nicht gescannte
# Bestellungen aus einem frueheren Lauf wuerde ersatzlos verschwinden. Das
# faellt bei der Deutschen Post am haertesten auf: dort steht KEINE
# Rechnungsnummer auf dem Label, der Abgleich laeuft ausschliesslich ueber
# post_zuordnung.csv (PLZ+Hausnummer) - fehlt der Eintrag, kann scan_druck.py
# das Label ueberhaupt nicht mehr zuordnen.
#
# Damit die Dateien dabei nicht unbegrenzt weiterwachsen, verfallen alte
# Eintraege nach BRUECKEN_CSV_MERGE_TAGE automatisch (Rechnungsdatum als
# LETZTE Spalte in jeder Datei angehaengt, damit eine aeltere scan_druck.py-
# Version, die diese Spalte nicht kennt, die Dateien weiterhin lesen kann).
BRUECKEN_CSV_MERGE_TAGE = 5


def _bruecken_datum_frisch(datum_str, max_tage=BRUECKEN_CSV_MERGE_TAGE):
    """True, wenn ein 'TT.MM.JJJJ'-Rechnungsdatum noch nicht aelter als
    max_tage ist. Unlesbares/leeres Datum wird vorsichtshalber als 'noch
    frisch' gewertet (lieber einen Alteintrag zu lange behalten als einen
    gueltigen zu frueh verlieren)."""
    if not datum_str:
        return True
    try:
        d = datetime.strptime(datum_str.strip(), "%d.%m.%Y")
    except ValueError:
        return True
    return (datetime.now() - d).days <= max_tage


def _lies_bestehende_zeilen(csv_pfad, mindest_spalten):
    """Bestehende Bruecken-CSV zeilenweise als Liste von Tokenlisten einlesen
    (roh, vor der eigentlichen Verarbeitung). Bei fehlender/kaputter Datei:
    leere Liste (verhaelt sich dann wie 'erster Lauf', nichts zum
    Zusammenfuehren)."""
    if not os.path.exists(csv_pfad):
        return []
    zeilen = []
    try:
        with open(csv_pfad, encoding="utf-8-sig") as f:
            next(f, None)  # Kopfzeile
            for zeile in f:
                t = zeile.rstrip("\n").split(";")
                if len(t) >= mindest_spalten and t[0]:
                    zeilen.append(t)
    except Exception:
        return []
    return zeilen


def schreibe_csv(rechnungen, csv_pfad):
    """post_zuordnung.csv fuer scan_druck.py - v.a. fuer das Deutsche-Post-
    Matching per PLZ+Hausnummer (auf dem Label selbst steht keine
    Rechnungsnummer). Format:
        Rechnungsnummer;Name;Strasse;Hausnummer;PLZ;Rechnungsdatum
    Wird mit einer eventuell bestehenden Datei zusammengefuehrt, siehe
    BRUECKEN_CSV_MERGE_TAGE weiter oben."""
    bestand = {}
    for t in _lies_bestehende_zeilen(csv_pfad, 5):
        datum = t[5] if len(t) > 5 else ""
        if _bruecken_datum_frisch(datum):
            bestand[t[0]] = [t[0], t[1], t[2], t[3], t[4], datum]
    for r in rechnungen:
        if not r["rnr"]:
            continue
        bestand[r["rnr"]] = [r["rnr"], r["name"], r["strasse"], r["hausnr"],
                              r["plz"], r.get("datum", "")]
    with open(csv_pfad, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f, delimiter=";", lineterminator="\n")
        w.writerow(["Rechnungsnummer", "Name", "Strasse", "Hausnummer",
                     "PLZ", "Rechnungsdatum"])
        for row in bestand.values():
            w.writerow(row)


def schreibe_sammel_csv(gruppen, csv_pfad):
    """Bruecke fuer den Sammeldruck in scan_druck.py.
    Format: Code;Artikelnr;Bezeichnung;Menge;Rechnungsnummern;EAN;Rechnungsdatum
    (Rechnungsnummern kommasepariert). Die EAN-Spalte ist nur gefuellt, wenn
    der Artikel tatsaechlich eine EAN traegt - scan_druck.py verlangt dann vor
    dem Sammeldruck zusaetzlich zum Sammelcode die Bestaetigungs-EAN des
    Artikels. 'Rechnungsdatum' = Datum, an dem der Code ERSTELLT wurde (fuer
    die Verfallsfrist beim Zusammenfuehren mit einem spaeteren Lauf, siehe
    merge_sammelgruppen/BRUECKEN_CSV_MERGE_TAGE). Beide Spalten bewusst ans
    ENDE angehaengt, damit eine aeltere scan_druck.py-Version die Datei
    weiter unveraendert lesen kann.

    `gruppen` ist hier bereits das GEMERGTE Gesamtergebnis aus
    merge_sammelgruppen() (alte, noch frische Gruppen + neue Gruppen dieses
    Laufs) - nicht nur die Gruppen des aktuellen Laufs."""
    with open(csv_pfad, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f, delimiter=";", lineterminator="\n")
        w.writerow(["Code", "Artikelnr", "Bezeichnung", "Menge",
                     "Rechnungsnummern", "EAN", "Rechnungsdatum"])
        for g in gruppen:
            w.writerow([g["code"], g["art"], g["bez"], len(g["rnr"]),
                        ",".join(g["rnr"]), g.get("ean", ""), g.get("datum", "")])


def schreibe_mengen_csv(rechnungen, csv_pfad):
    """Bruecke fuer die Mehrfach-Scan-Sicherung in scan_druck.py.
    Format: Rechnungsnummer;Scananzahl;Rechnungsdatum  (1..5). Die Scananzahl
    je Bestellung kommt aus order_scananzahl(): Summe der Positions-Scans
    (genau 1 -> 1, unter 1 wie 0,5 kg -> 2, sonst aufgerundete Menge),
    Gewichtsartikel zaehlen nach kg, gedeckelt auf 5. scan_druck.py liest die
    Zahl direkt als noetige Scananzahl - es muss dafuer NICHT geaendert/neu
    deployt werden. Wird mit einer eventuell bestehenden Datei
    zusammengefuehrt, siehe BRUECKEN_CSV_MERGE_TAGE weiter oben."""
    bestand = {}
    for t in _lies_bestehende_zeilen(csv_pfad, 2):
        datum = t[2] if len(t) > 2 else ""
        if _bruecken_datum_frisch(datum):
            bestand[t[0]] = [t[0], t[1], datum]
    for r in rechnungen:
        if not r["rnr"]:
            continue
        bestand[r["rnr"]] = [r["rnr"], order_scananzahl(r), r.get("datum", "")]
    with open(csv_pfad, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f, delimiter=";", lineterminator="\n")
        w.writerow(["Rechnungsnummer", "Scananzahl", "Rechnungsdatum"])
        for row in bestand.values():
            w.writerow(row)


def schreibe_ean_csv(rechnungen, csv_pfad):
    """Bruecke fuer den EAN-Verifikationsscan in scan_druck.py.
    Eine Zeile JE Position MIT EAN:
        Rechnungsnummer;EAN;Bezeichnung;Anzahl;Rechnungsdatum
    `Anzahl` = noetige Verifikations-Scans dieser Position (= scan_bedarf der
    bestellten Menge, gedeckelt auf MAX). Positionen OHNE EAN tauchen hier NICHT
    auf - fuer sie bleibt es bei der normalen Mehrfach-Scan-Sicherung ueber die
    Stueckzahl (mengen_zuordnung.csv). So wird der EAN-Scan ausschliesslich bei
    Artikeln verlangt, die in der Pickliste eine EAN tragen.

    Wird mit einer eventuell bestehenden Datei zusammengefuehrt (gruppiert
    nach Rechnungsnummer): Rechnungen aus dem AKTUELLEN Lauf ersetzen ihre
    alten Zeilen komplett; Rechnungen, die NUR im Bestand stehen (frueherer
    Lauf, noch nicht gescannt), bleiben unveraendert erhalten - siehe
    BRUECKEN_CSV_MERGE_TAGE weiter oben."""
    bestand = {}  # rnr -> Liste von Zeilen (inkl. Rechnungsdatum)
    for t in _lies_bestehende_zeilen(csv_pfad, 4):
        datum = t[4] if len(t) > 4 else ""
        if _bruecken_datum_frisch(datum):
            bestand.setdefault(t[0], []).append([t[0], t[1], t[2], t[3], datum])

    for r in rechnungen:
        if not r["rnr"]:
            continue
        zeilen = []
        for p in r["positionen"]:
            ean = (p.get("ean") or "").strip()
            if not ean:
                continue
            if ist_versand(p["art"], p["bez"]):
                continue
            qty = p["menge"] if p["menge"] is not None else 1
            anzahl = scan_bedarf(qty)
            zeilen.append([r["rnr"], ean, p["bez"], anzahl, r.get("datum", "")])
        # Rechnung aus dem aktuellen Lauf ERSETZT ihre alten Zeilen komplett
        # (auch mit einer leeren Liste, falls inzwischen keine EAN-Position
        # mehr vorhanden waere).
        bestand[r["rnr"]] = zeilen

    with open(csv_pfad, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f, delimiter=";", lineterminator="\n")
        w.writerow(["Rechnungsnummer", "EAN", "Bezeichnung", "Anzahl", "Rechnungsdatum"])
        for zeilen in bestand.values():
            for row in zeilen:
                w.writerow(row)


def archiviere(rechnungen, archiv_basis):
    """Verschiebt die AUFGENOMMENEN Rechnungs-PDFs nach
    <archiv_basis>/<JJJJ-MM-TT>/, damit sie nicht erneut auf eine Packliste
    geraten. Nur erfolgreich eingelesene Rechnungen werden verschoben -
    uebersprungene/fehlerhafte bleiben liegen. (verschoben, fehler, zielordner)."""
    ziel = os.path.join(archiv_basis, datetime.now().strftime("%Y-%m-%d"))
    os.makedirs(ziel, exist_ok=True)
    verschoben, fehler = 0, []
    for r in rechnungen:
        quelle = r.get("quelle")
        if not quelle or not os.path.exists(quelle):
            continue
        name = os.path.basename(quelle)
        zpfad = os.path.join(ziel, name)
        if os.path.exists(zpfad):           # Namenskollision vermeiden
            stamm, ext = os.path.splitext(name)
            i = 2
            while os.path.exists(os.path.join(ziel, f"{stamm}_{i}{ext}")):
                i += 1
            zpfad = os.path.join(ziel, f"{stamm}_{i}{ext}")
        try:
            shutil.move(quelle, zpfad)
            verschoben += 1
        except Exception as e:
            fehler.append(f"{name}: {e}")
    return verschoben, fehler, ziel


# ----------------------------------------------------------------------------
# MAIN
# ----------------------------------------------------------------------------

def main(argv):
    if len(argv) < 3:
        print("Aufruf: python packliste.py <rechnungs_ordner> "
              "\"<stammdaten_oder_leer>\" <ausgabe.pdf> [archiv_ordner]")
        return 2
    ordner = argv[0]
    _stammdaten = argv[1]  # optional, derzeit nicht fuer Sortierung benoetigt
    ausgabe = argv[2]
    archiv_ordner = argv[3].strip() if len(argv) >= 4 and argv[3].strip() else None

    pdfs = sorted(glob.glob(os.path.join(ordner, "*.pdf")))
    if not pdfs:
        print(f"Keine PDF-Rechnungen in {ordner} gefunden.")
        return 1

    rechnungen = []
    print(f"Lese {len(pdfs)} PDF(s) aus {ordner} ...")
    for p in pdfs:
        try:
            r = parse_pdf(p)
        except Exception as e:
            print(f"  FEHLER beim Lesen von {os.path.basename(p)}: {e}")
            continue
        if not r["positionen"]:
            print(f"  uebersprungen (keine Positionen erkannt): {os.path.basename(p)}")
            continue
        r["quelle"] = p          # Quellpfad fuer die spaetere Archivierung merken
        rechnungen.append(r)
        ok = r["zeilen_ok"] and r["summe_ok"] and r["vollstaendig_ok"]
        status = "OK" if ok else "WARNUNG"
        total_txt = f"{r['total']:.2f}" if r.get("total_lesbar", True) else "n.l."
        print(f"  {r['datei']}: Rg {r['rnr']}  {len(r['positionen'])} Pos  "
              f"Summe {r['summe']:.2f}/{total_txt}  [{status}]")
        if not r.get("total_lesbar", True):
            print(f"      Hinweis: Rechnungsbetrag nicht lesbar (lange Bezeichnung) "
                  f"-> Positionssumme {r['summe']:.2f} geprueft")
        if not r["vollstaendig_ok"]:
            print(f"      unvollstaendige Position(en): {', '.join(r['unvollstaendig'])}")

    if not rechnungen:
        print("Keine auswertbaren Rechnungen.")
        return 1

    os.makedirs(os.path.dirname(os.path.abspath(ausgabe)), exist_ok=True)
    out_dir = os.path.dirname(os.path.abspath(ausgabe))
    sammel_pfad = os.path.join(out_dir, "sammel_zuordnung.csv")

    gruppen_roh = finde_sammelgruppen(rechnungen)
    heute_str = datetime.now().strftime("%d.%m.%Y")
    for g in gruppen_roh:
        g["datum"] = heute_str
    # Codes kollisionsfrei mit einem eventuell frueheren, noch nicht
    # abgearbeiteten Lauf zusammenfuehren (siehe merge_sammelgruppen).
    gruppen, gruppen_fuer_csv = merge_sammelgruppen(gruppen_roh, sammel_pfad)
    baue_pdf(rechnungen, ausgabe, gruppen)

    csv_pfad = os.path.join(out_dir, "post_zuordnung.csv")
    schreibe_csv(rechnungen, csv_pfad)
    schreibe_sammel_csv(gruppen_fuer_csv, sammel_pfad)
    mengen_pfad = os.path.join(out_dir, "mengen_zuordnung.csv")
    schreibe_mengen_csv(rechnungen, mengen_pfad)
    ean_pfad = os.path.join(out_dir, "ean_zuordnung.csv")
    schreibe_ean_csv(rechnungen, ean_pfad)

    warn = [r for r in rechnungen
            if not (r["zeilen_ok"] and r["summe_ok"] and r["vollstaendig_ok"])]
    print("-" * 60)
    print(f"PDF geschrieben:  {ausgabe}")
    print(f"CSV geschrieben:  {csv_pfad}")
    print(f"Mengen-CSV:       {mengen_pfad}")
    # EAN-CSV: Anzahl Zeilen (= Positionen mit EAN) zur schnellen Kontrolle melden
    _ean_pos = sum(1 for r in rechnungen for p in r["positionen"]
                   if (p.get("ean") or "").strip() and not ist_versand(p["art"], p["bez"]))
    print(f"EAN-CSV:          {ean_pfad}  ({_ean_pos} Position(en) mit EAN)")
    if gruppen:
        print(f"Sammeldruck-CSV:  {sammel_pfad}  ({len(gruppen)} Sammelcode(s): "
              + ", ".join(f"{g['code']}={g['art']}×{len(g['rnr'])}" for g in gruppen) + ")")
    else:
        print(f"Sammeldruck-CSV:  {sammel_pfad}  (keine Gruppe ab {SAMMEL_MIN} gleichen Einzelartikeln)")
    if warn:
        print(f"ACHTUNG: {len(warn)} Rechnung(en) mit Warnung (Betrag/Vollständigkeit): "
              + ", ".join(r["rnr"] for r in warn))
    else:
        print("Alle Rechnungen: Abgleich UND Vollständigkeit OK.")

    # Aufgenommene Rechnungen archivieren (nur wenn Archivordner uebergeben)
    if archiv_ordner:
        verschoben, fehler, ziel = archiviere(rechnungen, archiv_ordner)
        print(f"Archiviert:       {verschoben} Rechnung(en) -> {ziel}")
        if fehler:
            print(f"  NICHT verschoben ({len(fehler)}): " + "; ".join(fehler))
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
