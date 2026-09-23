# -*- coding: utf-8 -*-
"""
carrier_statistik.py - Kg-Statistik je Carrier fuer das Carrier-Dashboard
================================================================================
Haengt bei jedem Schritt-2-Lauf (carrier_dashboard.exportiere_alles) eine Zeile
pro TATSAECHLICH exportierter Rechnung an eine CSV an (dieselbe Bedingung wie
carrier_export.exportiere(): status != "fehler" UND Carrier zugeordnet):
Datum, Carrier, Gewicht (kg), Land, Ausland-Flag, Rechnungsnummer. Reine
Schreib-/Auswertelogik ohne GUI - Selbsttest via py carrier_statistik.py.

Ziel: Jahres-/Monats-Auswertung "wieviel kg bewegen wir pro Carrier", z.B. als
Grundlage fuer Mengenrabatt-Verhandlungen mit DHL/DPD (siehe statistik_text()).
Liegt bewusst auf demselben Netzlaufwerk wie die bestehende statistik.csv aus
scan_druck.py (gleicher Ordner, andere Datei) - dauerhaft, nicht an einen
einzelnen PC (Faktura-PC) gebunden.

Das Loggen ist ein Nice-to-have und darf den eigentlichen Export/die
Archivierung NIE blockieren - log_lauf() faengt Schreibfehler (z.B.
Netzlaufwerk kurz nicht erreichbar) ab und wirft keine Exception.
"""

import os
from collections import defaultdict
from datetime import datetime

VERSION = "2026-09-23b"

STATISTIK_DATEI = r"\\DESKTOP-N2H75H\Netzwerk\Paketscheine\carrier_statistik.csv"
ARTIKEL_STATISTIK_DATEI = r"\\DESKTOP-N2H75H\Netzwerk\Paketscheine\artikel_statistik.csv"

_HEADER = "Datum;Carrier;Gewicht_kg;Land;Ausland;Rechnungsnummer\n"
_ARTIKEL_HEADER = "Datum;Rechnungsnummer;Artikelanzahl\n"


def _de(x):
    """Zahl mit deutschem Komma, 1 Nachkommastelle (nur fuer die Textanzeige,
    NICHT fuer die CSV - die CSV nutzt _gewicht_txt-Stil mit 4 Stellen)."""
    return ("%.1f" % x).replace(".", ",")


def log_lauf(ergebnisse, pfad=STATISTIK_DATEI, jetzt=None):
    """Haengt fuer jede exportierte Rechnung (status != 'fehler' UND Carrier
    zugeordnet) eine Zeile an pfad an. Legt Datei+Kopfzeile bei Bedarf an.
    Rueckgabe: Anzahl geloggter Zeilen (0 bei leerer Liste ODER Schreibfehler -
    Fehler werden bewusst verschluckt, siehe Modul-Kopf)."""
    zeilen = [b for b in ergebnisse if b["status"] != "fehler" and b["carrier"]]
    if not zeilen:
        return 0
    jetzt = jetzt or datetime.now()
    try:
        ordner = os.path.dirname(pfad)
        if ordner:
            os.makedirs(ordner, exist_ok=True)
        neu = not os.path.exists(pfad)
        with open(pfad, "a", encoding="utf-8", newline="") as f:
            if neu:
                f.write(_HEADER)
            for b in zeilen:
                gtxt = ("%.4f" % (b["gewicht"] or 0)).replace(".", ",")
                f.write("%s;%s;%s;%s;%s;%s\n" % (
                    jetzt.strftime("%Y-%m-%d"), b["carrier"], gtxt,
                    b["adresse"]["land"] or "", "1" if b["ausland"] else "0",
                    b["rnr"]))
    except OSError:
        return 0
    return len(zeilen)


def _artikelanzahl(r):
    """Summe der bestellten Menge ueber alle ECHTEN (Nicht-Versand-)Positionen
    einer Rechnung - z.B. 1x 10er-Pack Dichtungen + 2x 5kg-Schweissdraht +
    1x Schlauch = 4. Zaehlt die Menge WIE AUF DER RECHNUNG, OHNE Fach-Artikel-/
    Multipack-Multiplikator (effektive_menge() in packliste.py) - ein
    '2-Fach-Artikel' mit Menge 1 zaehlt hier als 1, nicht als 2 Stueck; dieser
    Multiplikator betrifft nur die Pickliste, nicht diese Verkaufsstatistik."""
    import packliste
    n = 0.0
    for p in r.get("positionen") or []:
        if packliste.ist_versand(p.get("art"), p.get("bez")):
            continue
        menge = p.get("menge")
        n += menge if menge is not None else 1
    return n


def log_artikel(rechnungen, pfad=ARTIKEL_STATISTIK_DATEI, jetzt=None):
    """Haengt fuer JEDE in Schritt 2 verarbeitete Rechnung (unabhaengig vom
    Carrier-Status - ein Adressfehler aendert nichts an der bestellten
    Artikelmenge) eine Zeile Datum;Rechnungsnummer;Artikelanzahl an pfad an.
    Legt Datei+Kopfzeile bei Bedarf an. Rueckgabe: Anzahl geloggter Zeilen (0
    bei leerer Liste ODER Schreibfehler - siehe log_lauf())."""
    if not rechnungen:
        return 0
    jetzt = jetzt or datetime.now()
    try:
        ordner = os.path.dirname(pfad)
        if ordner:
            os.makedirs(ordner, exist_ok=True)
        neu = not os.path.exists(pfad)
        with open(pfad, "a", encoding="utf-8", newline="") as f:
            if neu:
                f.write(_ARTIKEL_HEADER)
            for r in rechnungen:
                n = _artikelanzahl(r)
                ntxt = str(int(n)) if n == int(n) else ("%.2f" % n).replace(".", ",")
                f.write("%s;%s;%s\n" % (jetzt.strftime("%Y-%m-%d"), r.get("rnr") or "", ntxt))
    except OSError:
        return 0
    return len(rechnungen)


def _lies(pfad):
    """Rohe Zeilen (je eine Liste von Feldern) der Statistik-CSV, Kopfzeile
    uebersprungen. Leere Liste wenn die Datei fehlt oder nicht lesbar ist."""
    zeilen = []
    if not os.path.exists(pfad):
        return zeilen
    try:
        with open(pfad, encoding="utf-8") as f:
            next(f, None)                    # Kopfzeile
            for zeile in f:
                t = zeile.rstrip("\n").split(";")
                if len(t) >= 6 and t[0]:
                    zeilen.append(t)
    except OSError:
        pass
    return zeilen


def _lies_artikel(pfad):
    """Rohe Zeilen der Artikel-Statistik-CSV, Kopfzeile uebersprungen. Leere
    Liste wenn die Datei fehlt oder nicht lesbar ist."""
    zeilen = []
    if not os.path.exists(pfad):
        return zeilen
    try:
        with open(pfad, encoding="utf-8") as f:
            next(f, None)                    # Kopfzeile
            for zeile in f:
                t = zeile.rstrip("\n").split(";")
                if len(t) >= 3 and t[0]:
                    zeilen.append(t)
    except OSError:
        pass
    return zeilen


def artikel_je_monat(pfad=ARTIKEL_STATISTIK_DATEI, jahr=None):
    """{'YYYY-MM': (anzahl_bestellungen, artikel_summe)}. jahr (int) filtert
    optional auf ein Jahr, sonst alle vorhandenen Jahre."""
    nach_monat = defaultdict(lambda: [0, 0.0])   # [bestellungen, artikel_summe]
    for t in _lies_artikel(pfad):
        datum, artikelanzahl_txt = t[0], t[2]
        if jahr is not None and not datum.startswith("%d-" % jahr):
            continue
        try:
            n = float(artikelanzahl_txt.replace(",", "."))
        except ValueError:
            continue
        eintrag = nach_monat[datum[:7]]
        eintrag[0] += 1
        eintrag[1] += n
    return {monat: tuple(werte) for monat, werte in nach_monat.items()}


def artikel_jahr(pfad=ARTIKEL_STATISTIK_DATEI, jahr=None):
    """(anzahl_bestellungen, artikel_summe, durchschnitt_je_bestellung) eines
    Jahres (Default: laufendes Jahr)."""
    jahr = jahr or datetime.now().year
    daten = artikel_je_monat(pfad, jahr)
    bestellungen = sum(b for b, _ in daten.values())
    artikel = sum(a for _, a in daten.values())
    schnitt = (artikel / bestellungen) if bestellungen else 0.0
    return bestellungen, artikel, schnitt


def kg_je_carrier_monat(pfad=STATISTIK_DATEI, jahr=None):
    """{'YYYY-MM': {carrier: kg_summe}}. jahr (int) filtert optional auf ein
    Jahr, sonst alle vorhandenen Jahre."""
    nach_monat = defaultdict(lambda: defaultdict(float))
    for t in _lies(pfad):
        datum, carrier, gewicht_txt = t[0], t[1], t[2]
        if jahr is not None and not datum.startswith("%d-" % jahr):
            continue
        try:
            kg = float(gewicht_txt.replace(",", "."))
        except ValueError:
            continue
        nach_monat[datum[:7]][carrier] += kg
    return {monat: dict(carrier_kg) for monat, carrier_kg in nach_monat.items()}


def kg_jahr(pfad=STATISTIK_DATEI, jahr=None):
    """Gesamt-kg eines Jahres (Default: laufendes Jahr) UEBER ALLE Carrier."""
    jahr = jahr or datetime.now().year
    daten = kg_je_carrier_monat(pfad, jahr)
    return sum(kg for carrier_kg in daten.values() for kg in carrier_kg.values())


def _kg_block(pfad, jahr):
    daten = kg_je_carrier_monat(pfad, jahr)
    if not daten:
        return ["Kg-Statistik %d" % jahr, "", "  Noch keine Daten vorhanden."]
    je_carrier = defaultdict(float)
    for carrier_kg in daten.values():
        for carrier, kg in carrier_kg.items():
            je_carrier[carrier] += kg
    gesamt = sum(je_carrier.values())
    zeilen = ["Kg-Statistik %d" % jahr, "", "Gesamt:  %s kg" % _de(gesamt), "", "Je Carrier:"]
    for carrier in sorted(je_carrier, key=lambda c: -je_carrier[c]):
        zeilen.append("  %-20s %10s kg" % (carrier, _de(je_carrier[carrier])))
    zeilen += ["", "Je Monat:"]
    for monat in sorted(daten):
        teil = ", ".join("%s %s kg" % (c, _de(kg)) for c, kg in sorted(daten[monat].items()))
        zeilen.append("  %s:  %s" % (monat, teil))
    return zeilen


def _artikel_block(pfad, jahr):
    daten = artikel_je_monat(pfad, jahr)
    if not daten:
        return ["Artikel-Statistik %d" % jahr, "", "  Noch keine Daten vorhanden."]
    bestellungen = sum(b for b, _ in daten.values())
    artikel = sum(a for _, a in daten.values())
    schnitt = (artikel / bestellungen) if bestellungen else 0.0
    zeilen = ["Artikel-Statistik %d" % jahr, "",
              "Bestellungen:      %d" % bestellungen,
              "Artikel gesamt:    %s" % _de(artikel),
              "Ø Artikel/Bestellung: %s" % _de(schnitt), "", "Je Monat:"]
    for monat in sorted(daten):
        b, a = daten[monat]
        zeilen.append("  %s:  %d Bestellung(en), %s Artikel" % (monat, b, _de(a)))
    return zeilen


def statistik_text(pfad=STATISTIK_DATEI, artikel_pfad=ARTIKEL_STATISTIK_DATEI, jahr=None):
    """Mehrzeiliger Text fuer eine einfache messagebox-Anzeige im Dashboard:
    Kg-Statistik (Gesamt + je Carrier + je Monat) gefolgt von der Artikel-
    Statistik (Bestellungen/Artikel/Schnitt je Monat) desselben Jahres."""
    jahr = jahr or datetime.now().year
    return "\n".join(_kg_block(pfad, jahr) + ["", "─" * 40, ""] + _artikel_block(artikel_pfad, jahr))


# ==========================================================================
# Selbsttest:  py carrier_statistik.py
# ==========================================================================

def selftest():
    import shutil
    import tempfile

    n_ok, n_fail = 0, []

    def check(name, ist, soll):
        nonlocal n_ok
        if ist == soll:
            n_ok += 1
        else:
            n_fail.append("%s: erwartet %r, war %r" % (name, soll, ist))

    def _b(carrier, gewicht, rnr, status="ok", land="DE", ausland=False):
        return {"status": status, "carrier": carrier, "gewicht": gewicht, "rnr": rnr,
                "adresse": {"land": land}, "ausland": ausland}

    tmp = tempfile.mkdtemp(prefix="carrier_statistik_test_")
    try:
        pfad = os.path.join(tmp, "unterordner", "carrier_statistik.csv")

        n = log_lauf([], pfad)
        check("log_lauf(): leere Liste -> 0, keine Datei", (n, os.path.exists(pfad)), (0, False))

        lauf1 = [
            _b("DHL", 3.8, "1705334"),
            _b("DHL", 0.5, "1705335"),
            _b("DPD", 1.0, "1705336"),
            _b("Post Brief", 0.03, "1705337"),
            _b(None, 1.0, "1705338", status="fehler"),      # kein Carrier -> nicht geloggt
            _b("DHL", 2.0, "1705339", status="fehler"),     # Fehler-Status -> nicht geloggt
        ]
        n = log_lauf(lauf1, pfad, datetime(2026, 3, 15))
        check("log_lauf(): nur exportierbare Zeilen geloggt", n, 4)
        check("log_lauf(): legt Zielordner an", os.path.isdir(os.path.dirname(pfad)), True)

        lauf2 = [_b("DHL", 4.2, "1705400"), _b("Post Großbrief", 0.5, "1705401", ausland=True)]
        log_lauf(lauf2, pfad, datetime(2026, 8, 2))
        lauf3 = [_b("DHL", 10.0, "1600001")]
        log_lauf(lauf3, pfad, datetime(2025, 12, 1))       # anderes Jahr

        with open(pfad, encoding="utf-8") as f:
            inhalt = f.read()
        check("CSV: Header exakt einmal", inhalt.count("Datum;Carrier;"), 1)
        check("CSV: 7 Datenzeilen (4+2+1)", inhalt.count("\n") - 1, 7)

        daten_2026 = kg_je_carrier_monat(pfad, 2026)
        check("kg_je_carrier_monat: nur 2026", sorted(daten_2026), ["2026-03", "2026-08"])
        check("kg_je_carrier_monat: DHL Maerz = 4.3", round(daten_2026["2026-03"]["DHL"], 2), 4.3)
        check("kg_je_carrier_monat: DPD Maerz = 1.0", daten_2026["2026-03"]["DPD"], 1.0)
        check("kg_je_carrier_monat: 2025 nicht enthalten", "2025" in str(daten_2026), False)

        check("kg_jahr(2026) = Summe aller Carrier/Monate", round(kg_jahr(pfad, 2026), 2), 10.03)
        check("kg_jahr(2025) getrennt", round(kg_jahr(pfad, 2025), 2), 10.0)

        text = statistik_text(pfad, jahr=2026)
        check("statistik_text: enthaelt Jahr", "2026" in text, True)
        check("statistik_text: enthaelt DHL-Zeile", "DHL" in text, True)
        check("statistik_text: 2025-Zeile NICHT in 2026-Auswertung", "1600001" in text, False)

        check("statistik_text: leeres Jahr -> Hinweistext",
              "Noch keine" in statistik_text(pfad, jahr=1999), True)

        # --- Artikelanzahl (unabhaengig vom Carrier-Status geloggt) ---------
        r_beispiel = {"rnr": "1705500", "positionen": [
            {"art": "D10", "bez": "10er Pack Dichtungen", "menge": 1},
            {"art": "SD5", "bez": "5kg Schweissdraht", "menge": 2},
            {"art": "SCH1", "bez": "Schlauch", "menge": 1},
            {"art": "", "bez": "Versandkosten", "menge": 1},       # zaehlt NICHT mit
        ]}
        check("_artikelanzahl(): Beispiel aus der Anfrage (1+2+1, Versand ausgeklammert)",
              _artikelanzahl(r_beispiel), 4)

        artikel_pfad = os.path.join(tmp, "artikel_statistik.csv")
        n = log_artikel([], artikel_pfad)
        check("log_artikel(): leere Liste -> 0, keine Datei",
              (n, os.path.exists(artikel_pfad)), (0, False))

        r2 = {"rnr": "1705501", "positionen": [{"art": "X", "bez": "Sonstiges", "menge": 3}]}
        r3 = {"rnr": "1705600", "positionen": [{"art": "Y", "bez": "Ding", "menge": 5}]}
        r4 = {"rnr": "1600100", "positionen": [{"art": "Z", "bez": "Alt", "menge": 2}]}
        log_artikel([r_beispiel, r2], artikel_pfad, datetime(2026, 3, 10))
        log_artikel([r3], artikel_pfad, datetime(2026, 8, 1))
        log_artikel([r4], artikel_pfad, datetime(2025, 12, 1))     # anderes Jahr

        daten = artikel_je_monat(artikel_pfad, 2026)
        check("artikel_je_monat: Monate", sorted(daten), ["2026-03", "2026-08"])
        check("artikel_je_monat: Maerz = 2 Bestellungen, 7 Artikel", daten["2026-03"], (2, 7.0))
        check("artikel_je_monat: August = 1 Bestellung, 5 Artikel", daten["2026-08"], (1, 5.0))

        b, a, schnitt = artikel_jahr(artikel_pfad, 2026)
        check("artikel_jahr: Bestellungen", b, 3)
        check("artikel_jahr: Artikel gesamt", a, 12.0)
        check("artikel_jahr: Schnitt", round(schnitt, 2), 4.0)
        check("artikel_jahr: 2025 getrennt", artikel_jahr(artikel_pfad, 2025), (1, 2.0, 2.0))

        text2 = statistik_text(pfad, artikel_pfad, jahr=2026)
        check("statistik_text: enthaelt Artikel-Block", "Artikel-Statistik" in text2, True)
        check("statistik_text: Bestellungen im Text", "Bestellungen:      3" in text2, True)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)

    if n_fail:
        print("SELBSTTEST FEHLGESCHLAGEN (%d von %d):" % (len(n_fail), n_ok + len(n_fail)))
        for f in n_fail:
            print("  -", f)
        return False
    print("Selbsttest OK (%d Prüfungen)" % n_ok)
    return True


if __name__ == "__main__":
    import sys
    sys.exit(0 if selftest() else 1)
