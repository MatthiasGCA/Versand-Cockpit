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

VERSION = "2026-09-23a"

STATISTIK_DATEI = r"\\DESKTOP-N2H75H\Netzwerk\Paketscheine\carrier_statistik.csv"

_HEADER = "Datum;Carrier;Gewicht_kg;Land;Ausland;Rechnungsnummer\n"


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


def statistik_text(pfad=STATISTIK_DATEI, jahr=None):
    """Mehrzeiliger Text: Gesamt-kg des Jahres + Aufschluesselung je Carrier
    und Monat, fuer eine einfache messagebox-Anzeige im Dashboard."""
    jahr = jahr or datetime.now().year
    daten = kg_je_carrier_monat(pfad, jahr)
    if not daten:
        return "Noch keine Carrier-Statistik für %d vorhanden." % jahr
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
    return "\n".join(zeilen)


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

        text = statistik_text(pfad, 2026)
        check("statistik_text: enthaelt Jahr", "2026" in text, True)
        check("statistik_text: enthaelt DHL-Zeile", "DHL" in text, True)
        check("statistik_text: 2025-Zeile NICHT in 2026-Auswertung", "1600001" in text, False)

        check("statistik_text: leeres Jahr -> Hinweistext",
              "Noch keine" in statistik_text(pfad, 1999), True)
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
