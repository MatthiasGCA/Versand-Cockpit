# -*- coding: utf-8 -*-
"""
artikel_auswertung_archiv.py - Rueckwaerts-Auswertung der Bestseller aus dem Rechnungs-Archiv
============================================================================
Liest ALLE archivierten Rechnungs-PDFs (Standard: ARCHIV_ORDNER des Carrier-Dashboards, auf
dem Faktura-PC also C:\\Packlisten\\Archiv inkl. aller Datums-Unterordner) und traegt jede
Rechnungsposition in die Artikel-Verkaeufe-CSV ein (artikel_verkaeufe.csv im Netzordner
Paketscheine - dieselbe Datei, die das Dashboard ab jetzt bei jedem Schritt 2 fortschreibt).
Danach zeigt der Dashboard-Button "Top-Artikel" auch die Zeit VOR der Inbetriebnahme.

  * Idempotent: Rechnungsnummern, die schon in der CSV stehen, werden uebersprungen - das
    Skript darf beliebig oft laufen (z.B. nach neuen Archivordnern).
  * Doppelte PDFs (gleiche Rechnungsnummer) zaehlen einmal; Gutschriften/Auftragsbestaetigungen
    (siehe packliste.pruefe_belegnummer) werden ausgelassen.
  * Datum der Position = Rechnungsdatum aus der PDF.

Aufruf (Eingabeaufforderung, im Dashboard-Ordner):
    py artikel_auswertung_archiv.py                 (Archiv + Ziel aus der Dashboard-Konfiguration)
    py artikel_auswertung_archiv.py --trocken       (nur auswerten und zeigen, nichts schreiben)
    py artikel_auswertung_archiv.py --archiv "D:\\Archiv" --ziel "C:\\Temp\\verkaeufe.csv"
    py artikel_auswertung_archiv.py --von 2026-08-01   (nur Rechnungen ab diesem Datum)
"""

import argparse
import os
import sys
import time

VERSION = "2026-10-08a"
BLOCK = 100          # nach so vielen neuen Rechnungen wird in die CSV geschrieben (Fortschritt sichern)


def sammle_pdfs(archiv):
    pdfs = []
    for dp, dn, fn in os.walk(archiv):
        pdfs += [os.path.join(dp, f) for f in fn if f.lower().endswith(".pdf")]
    return sorted(pdfs)


def main(argv=None):
    import carrier_dashboard as dash
    import carrier_statistik as cs
    import packliste

    ap = argparse.ArgumentParser(description="Bestseller-Rueckwaertsauswertung aus dem Archiv")
    ap.add_argument("--archiv", default=dash.ARCHIV_ORDNER)
    ap.add_argument("--ziel", default=cs.ARTIKEL_VERKAEUFE_DATEI)
    ap.add_argument("--von", default=None, help="nur Rechnungen ab JJJJ-MM-TT")
    ap.add_argument("--trocken", action="store_true", help="nichts schreiben")
    arg = ap.parse_args(argv)

    if not os.path.isdir(arg.archiv):
        print("Archivordner nicht gefunden: %s" % arg.archiv)
        return 1
    pdfs = sammle_pdfs(arg.archiv)
    bekannt = {t[1] for t in cs._lies_verkaeufe(arg.ziel)}
    print("Archiv:  %s  (%d PDFs)" % (arg.archiv, len(pdfs)))
    print("Ziel:    %s  (%d Rechnungen schon vorhanden)%s" % (
        arg.ziel, len(bekannt), "   [TROCKENLAUF - es wird nichts geschrieben]" if arg.trocken else ""))

    gesehen, block, summe_zeilen = set(), [], 0
    neu_rg = doppelt = uebersprungen = fehler = vorhanden = zu_alt = 0
    alle_zeilen = []                    # fuer die Zusammenfassung (nur neue Rechnungen)
    t0 = time.time()
    for i, pfad in enumerate(pdfs, 1):
        if i % 100 == 0:
            print("  ... %d/%d PDFs gelesen (%.0f s)" % (i, len(pdfs), time.time() - t0))
        try:
            r = packliste.parse_pdf(pfad)
        except Exception:
            fehler += 1
            continue
        rnr = r.get("rnr") or ""
        if packliste.pruefe_belegnummer(rnr) or not r.get("positionen"):
            uebersprungen += 1                  # Gutschrift/AB/unlesbar
            continue
        if rnr in gesehen:
            doppelt += 1
            continue
        gesehen.add(rnr)
        if rnr in bekannt:
            vorhanden += 1
            continue
        if arg.von and (cs._iso_datum(r.get("datum")) or "9999") < arg.von:
            zu_alt += 1
            continue
        neu_rg += 1
        block.append(r)
        alle_zeilen += cs.verkaeufe_zeilen(r)
        if len(block) >= BLOCK and not arg.trocken:
            summe_zeilen += cs.log_verkaeufe(block, arg.ziel)
            block = []
    if block and not arg.trocken:
        summe_zeilen += cs.log_verkaeufe(block, arg.ziel)

    print("\nErgebnis: %d neue Rechnungen (%d Positionen) %s" % (
        neu_rg, len(alle_zeilen), "ausgewertet" if arg.trocken else "in die CSV geschrieben"))
    print("          %d schon vorhanden, %d doppelte PDFs, %d uebersprungen (Gutschrift/unlesbar), "
          "%d Lesefehler%s" % (vorhanden, doppelt, uebersprungen, fehler,
                               (", %d vor %s" % (zu_alt, arg.von)) if arg.von else ""))
    if not arg.trocken and summe_zeilen != len(alle_zeilen):
        print("WARNUNG: geschrieben %d, erwartet %d Zeilen - Ziel erreichbar/beschreibbar?"
              % (summe_zeilen, len(alle_zeilen)))
        return 2
    if alle_zeilen:
        for sortierung in ("bestellungen", "umsatz"):
            liste = cs.top_artikel(zeilen=alle_zeilen, sortierung=sortierung, n=10)
            print("\n" + cs.top_text(liste, "Top 10 dieser Auswertung nach %s" % cs.SORTIERUNG[sortierung]))
    return 0


if __name__ == "__main__":
    sys.exit(main())
