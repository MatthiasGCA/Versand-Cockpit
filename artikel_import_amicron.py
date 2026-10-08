# -*- coding: utf-8 -*-
"""
artikel_import_amicron.py - Amicron-Export der Ausgangsrechnungen -> artikel_historie.csv
============================================================================
Traegt FRUEHERE Zeitraeume (vor dem Carrier-Dashboard) monatsweise in die Artikel-Historie ein, damit die
Online-Auswertung "Top-Artikel" Vorjahresvergleiche und Saisonverlaeufe zeigen kann (siehe artikel_upload.py).

Eingabe: CSV-Export(e) der Ausgangsrechnungen aus Amicron (Semikolon, Windows-1252 oder UTF-8). Die Spalten
werden am Namen erkannt:
  Artikelnummer   "Artikel Nr. intern" | "Artikelnummer" | "Artikel Nr." | "Art.-Nr."
  Bezeichnung     "Titel" | "Bezeichnung" | "Artikelbezeichnung"
  Menge           "Menge" | "Anzahl"
  Betrag          "Betrag" | "Gesamtpreis" | "G-Preis" | "Gesamt" | "Summe"          (optional)
  Einzelpreis     "Einzelpreis" | "E-Preis" - ohne Betrag-Spalte gilt Umsatz = Menge x Einzelpreis
  Rechnungsnr.    "Rechnungsnummer" | "Rechnung Nr." | "Beleg Nr." | "Belegnummer"   (optional)
  Datum           "Rechnungsdatum" | "Belegdatum" | "Datum" (TT.MM.JJJJ, JJJJ-MM-TT, JJJJMMTT) (optional)
Passende Amicron-Exportdefinition: amicron/Exportdefinition_Artikelverkaeufe.XML (eine Zeile je Position).
Ein Semikolon im Titel verschiebt die Spalten - die Zusatzspalten werden dem Titel zugeschlagen.
Versandkosten-Positionen werden wie im Dashboard weggelassen (ist_versand).

Zuordnung zum Monat:
  * hat die Datei eine Datumsspalte, wird je Zeile der Monat aus dem Datum bestimmt (ein Export fuer ein
    ganzes Jahr ist moeglich); --monat JJJJ-MM begrenzt dann auf diesen Monat.
  * ohne Datumsspalte muss --monat JJJJ-MM angegeben werden (eine Datei = ein Monat).

Ergebnis: artikel_historie.csv (Monat;Artikelnummer;Bezeichnung;Bestellungen;Menge;Umsatz). Bereits vorhandene
Zeilen eines importierten Monats werden ERSETZT, andere Monate bleiben - der Import ist wiederholbar.
"Bestellungen" ist leer, wenn der Export keine Rechnungsnummer enthaelt.

Aufruf:
    py artikel_import_amicron.py export_2025_03.csv --monat 2025-03
    py artikel_import_amicron.py export_2025.csv                (mit Datumsspalte, mehrere Monate)
    py artikel_import_amicron.py datei.csv --zeige-spalten      (nur Spalten + erste Zeilen zeigen)
    py artikel_import_amicron.py datei.csv --monat 2025-03 --trocken      (nichts schreiben)
    py artikel_import_amicron.py --selftest
"""

import argparse
import csv
import io
import os
import re
import sys
from collections import Counter, defaultdict
from datetime import datetime

VERSION = "2026-10-08b"
STANDARD_ZIEL = r"\\DESKTOP-N2H75H\Netzwerk\Paketscheine\artikel_historie.csv"
HEADER = "Monat;Artikelnummer;Bezeichnung;Bestellungen;Menge;Umsatz\n"

SPALTEN = {
    "art": ["artikel nr. intern", "artikelnummer", "artikel nr.", "artikelnr", "art.-nr.", "art.nr."],
    "titel": ["titel", "bezeichnung", "artikelbezeichnung"],
    "menge": ["menge", "anzahl"],
    "betrag": ["betrag", "gesamtpreis", "g-preis", "gesamt", "summe"],
    "ep": ["einzelpreis", "e-preis", "epreis"],
    "rnr": ["rechnungsnummer", "rechnung nr.", "rechnungs-nr.", "beleg nr.", "belegnummer", "rechnung"],
    "datum": ["rechnungsdatum", "belegdatum", "datum"],
}


def lies_text(pfad):
    raw = open(pfad, "rb").read()
    for enc in ("utf-8-sig", "cp1252"):
        try:
            return raw.decode(enc), enc
        except UnicodeDecodeError:
            continue
    return raw.decode("latin-1"), "latin-1"


def finde_spalten(header):
    norm = [h.strip().lower() for h in header]
    gefunden = {}
    for key, namen in SPALTEN.items():
        for n in namen:
            if n in norm:
                gefunden[key] = norm.index(n)
                break
    return gefunden


def zahl(s):
    s = (s or "").strip().replace("\u20ac", "").replace(" ", "")
    if not s:
        return 0.0
    if "," in s:
        s = s.replace(".", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return 0.0


def monat_aus_datum(s):
    s = (s or "").strip()[:10]
    if re.fullmatch(r"\d{8}", s):
        s = s[:4] + "-" + s[4:6] + "-" + s[6:]
    for fmt in ("%d.%m.%Y", "%Y-%m-%d", "%d.%m.%y"):
        try:
            return datetime.strptime(s, fmt).strftime("%Y-%m")
        except ValueError:
            continue
    return None


def ist_versand(art, bez):
    try:
        import packliste
        return packliste.ist_versand(art, bez)
    except Exception:                                   # packliste/pdfplumber nicht installiert
        return "versandkosten" in (bez or "").lower() or not ((art or "").strip() or (bez or "").strip())


def lies_export(text, monat=None):
    """-> (zeilen {(monat, art): {...}}, info dict). Wirft ValueError bei fehlenden Pflichtspalten."""
    zeilen = list(csv.reader(io.StringIO(text), delimiter=";", quoting=csv.QUOTE_NONE))
    if not zeilen:
        raise ValueError("leere Datei")
    cols = finde_spalten(zeilen[0])
    fehlend = [k for k in ("art", "titel", "menge") if k not in cols]
    if fehlend:
        raise ValueError("Spalte(n) nicht gefunden: %s. Vorhandene Spalten: %s"
                         % (", ".join(fehlend), "; ".join(zeilen[0])))
    if "datum" not in cols and not monat:
        raise ValueError("Der Export hat keine Datumsspalte - bitte --monat JJJJ-MM angeben")
    agg = {}
    info = {"zeilen": 0, "versand": 0, "ohne_monat": 0, "anderer_monat": 0, "spalten": cols}
    for z in zeilen[1:]:
        extra = len(z) - len(zeilen[0])
        if extra > 0 and "titel" in cols:                  # Semikolon im Titel: Zusatzspalten dem Titel zuschlagen
            t = cols["titel"]
            z = z[:t] + [";".join(z[t:t + extra + 1])] + z[t + extra + 1:]
        if len(z) <= max(cols.values()):
            continue
        info["zeilen"] += 1
        art, titel = z[cols["art"]].strip(), z[cols["titel"]].strip()
        if ist_versand(art, titel):
            info["versand"] += 1
            continue
        mo = monat_aus_datum(z[cols["datum"]]) if "datum" in cols else monat
        if not mo:
            info["ohne_monat"] += 1
            continue
        if monat and mo != monat:
            info["anderer_monat"] += 1
            continue
        a = agg.setdefault((mo, art.lower()), {"art": art, "titel": Counter(), "rnr": set(),
                                               "menge": 0.0, "umsatz": 0.0, "rnr_da": "rnr" in cols})
        a["titel"][titel] += 1
        a["menge"] += zahl(z[cols["menge"]])
        if "betrag" in cols:
            a["umsatz"] += zahl(z[cols["betrag"]])
        elif "ep" in cols:
            a["umsatz"] += zahl(z[cols["menge"]]) * zahl(z[cols["ep"]])
        if "rnr" in cols and z[cols["rnr"]].strip():
            a["rnr"].add(z[cols["rnr"]].strip())
    return agg, info


def zeile_text(mo, a):
    def z(x):
        return str(int(x)) if x == int(x) else ("%.2f" % x).replace(".", ",")
    titel = a["titel"].most_common(1)[0][0].replace(";", ",")
    best = str(len(a["rnr"])) if a["rnr_da"] else ""
    return "%s;%s;%s;%s;%s;%s\n" % (mo, a["art"].replace(";", ","), " ".join(titel.split()), best, z(a["menge"]),
                                    ("%.2f" % a["umsatz"]).replace(".", ","))


def schreibe_historie(ziel, agg, trocken=False):
    """Ersetzt alle Zeilen der Monate in agg, behaelt die uebrigen. Rueckgabe Anzahl geschriebener Zeilen."""
    monate = {k[0] for k in agg}
    behalten = []
    if os.path.exists(ziel):
        with open(ziel, encoding="utf-8-sig") as f:
            next(f, None)
            behalten = [z for z in f if z.split(";", 1)[0] not in monate and z.strip()]
    neu = [zeile_text(mo, a) for (mo, _), a in sorted(agg.items())]
    if not trocken:
        ordner = os.path.dirname(ziel)
        if ordner:
            os.makedirs(ordner, exist_ok=True)
        zeilen = sorted(behalten + neu, key=lambda z: (z.split(";", 1)[0], z.split(";", 2)[1].lower()))
        with open(ziel, "w", encoding="utf-8", newline="") as f:
            f.write(HEADER)
            f.writelines(z if z.endswith("\n") else z + "\n" for z in zeilen)
    return len(neu)


def zusammenfassung(agg):
    je = defaultdict(lambda: [0, 0.0, 0.0])
    for (mo, _), a in agg.items():
        je[mo][0] += 1
        je[mo][1] += a["menge"]
        je[mo][2] += a["umsatz"]
    zeilen = []
    for mo in sorted(je):
        n, m, u = je[mo]
        zeilen.append("  %s: %d Artikel, %.0f Stück, Umsatz %.2f EUR" % (mo, n, m, u))
        top = sorted(((a["menge"], a["art"], a["titel"].most_common(1)[0][0]) for (k, _), a in agg.items() if k == mo),
                     reverse=True)[:3]
        zeilen += ["      %s  %-16s %s" % (("%.0f" % m_).rjust(6), a_[:16], t_[:44]) for m_, a_, t_ in top]
    return "\n".join(zeilen)


def selftest():
    import tempfile
    ok, fehler = 0, []

    def check(name, ist, soll):
        nonlocal ok
        if ist == soll:
            ok += 1
        else:
            fehler.append("%s: erwartet %r, war %r" % (name, soll, ist))
    amicron = ("Rechnung Nr.;Rechnungsdatum;Artikel Nr. intern;Titel;Menge;Betrag\n"
               "1001;05.03.2025;A1;Artikel Eins;2;10,00\n1001;05.03.2025;660;Versandkosten;1;4,90\n"
               "1002;07.03.2025;A1;Artikel Eins Zoll;1;5,00\n1002;07.03.2025;B2;Artikel Zwei;1,5;1.234,50\n"
               "1003;02.04.2025;A1;Artikel Eins;3;15,00\n")
    agg, info = lies_export(amicron)
    check("Monate aus Datum, Versand raus", sorted((k[0], k[1]) for k in agg), [("2025-03", "a1"), ("2025-03", "b2"), ("2025-04", "a1")])
    a1 = agg[("2025-03", "a1")]
    check("A1 Maerz: Menge 3, Umsatz 15, 2 Rechnungen", (a1["menge"], a1["umsatz"], len(a1["rnr"])), (3.0, 15.0, 2))
    check("Tausenderpunkt/Komma: 1.234,50", agg[("2025-03", "b2")]["umsatz"], 1234.5)
    check("Versandzeile gezaehlt", info["versand"], 1)
    agg2, _ = lies_export(amicron, monat="2025-03")
    check("--monat begrenzt", sorted({k[0] for k in agg2}), ["2025-03"])
    neu = ("Rechnungsnummer;Rechnungsdatum;Artikelnummer;Titel;Menge;Einzelpreis\n"
           "2001;20250310;X1;Teil 1/4\" x 8mm; lang;3;2,50\n2001;20250310;660;Versandkosten;1;4,90\n"
           "2002;20250311;X1;Teil;1;2,50\n")
    agg4, i4 = lies_export(neu)
    x1 = agg4[("2025-03", "x1")]
    check("Exportdefinition-Format: JJJJMMTT, Menge x Einzelpreis, Semikolon im Titel",
          (x1["menge"], round(x1["umsatz"], 2), len(x1["rnr"])), (4.0, 10.0, 2))
    agg5, _ = lies_export("Rechnungsnummer;Rechnungsdatum;Artikelnummer;Titel;Menge;Einzelpreis\n"
                          "3001;20250312;Q1;\"Quote am Anfang;1;1,00\n3002;20250312;Q2;Normal;2;1,00\n")
    check("Anfuehrungszeichen am Titelanfang verschluckt keine Folgezeilen", sorted(k[1] for k in agg5), ["q1", "q2"])
    zub = "Artikelnummer;Titel;Menge;Betrag\nA1;Artikel Eins;4;0,00\n"
    try:
        lies_export(zub)
        check("ohne Datum und --monat -> Fehler", True, False)
    except ValueError:
        check("ohne Datum und --monat -> Fehler", True, True)
    agg3, _ = lies_export(zub, monat="2025-05")
    check("Zubehoer-Format mit --monat, ohne Rechnungsnummer", (agg3[("2025-05", "a1")]["menge"], agg3[("2025-05", "a1")]["rnr_da"]), (4.0, False))
    tmp = tempfile.mkdtemp()
    try:
        ziel = os.path.join(tmp, "h.csv")
        schreibe_historie(ziel, agg)
        schreibe_historie(ziel, agg3)
        z1 = open(ziel, encoding="utf-8").read().strip().split("\n")
        check("Historie: Kopf + 4 Zeilen", len(z1), 5)
        check("Zeile mit Bestellungen", z1[1], "2025-03;A1;Artikel Eins;2;3;15,00")
        check("ohne Rechnungsnummer: Bestellungen leer", z1[4], "2025-05;A1;Artikel Eins;;4;0,00")
        schreibe_historie(ziel, agg2)                                   # Maerz erneut -> ersetzt, nicht doppelt
        check("wiederholbar: Maerz ersetzt", len(open(ziel, encoding="utf-8").read().strip().split("\n")), 5)
    finally:
        import shutil
        shutil.rmtree(tmp, ignore_errors=True)
    if fehler:
        print("SELBSTTEST FEHLGESCHLAGEN (%d von %d):" % (len(fehler), ok + len(fehler)))
        for f in fehler:
            print("  - " + f)
        return 1
    print("Selbsttest OK (%d Pruefungen)" % ok)
    return 0


def main(argv=None):
    ap = argparse.ArgumentParser(description="Amicron-Export -> artikel_historie.csv")
    ap.add_argument("dateien", nargs="*")
    ap.add_argument("--monat", default=None, help="JJJJ-MM")
    ap.add_argument("--ziel", default=STANDARD_ZIEL)
    ap.add_argument("--trocken", action="store_true")
    ap.add_argument("--zeige-spalten", action="store_true")
    ap.add_argument("--selftest", action="store_true")
    arg = ap.parse_args(argv)
    if arg.selftest:
        return selftest()
    if not arg.dateien:
        ap.print_help()
        return 1
    if arg.monat and not re.fullmatch(r"\d{4}-\d{2}", arg.monat):
        print("--monat muss JJJJ-MM sein, z.B. 2025-03")
        return 1
    gesamt = {}
    for pfad in arg.dateien:
        text, enc = lies_text(pfad)
        if arg.zeige_spalten:
            rows = list(csv.reader(io.StringIO(text), delimiter=";", quoting=csv.QUOTE_NONE))
            print("%s (%s, %d Zeilen)" % (pfad, enc, len(rows)))
            for z in rows[:4]:
                print("   ", " | ".join(z)[:200])
            print("    erkannte Spalten:", finde_spalten(rows[0]) if rows else {})
            continue
        try:
            agg, info = lies_export(text, arg.monat)
        except ValueError as e:
            print("%s: %s" % (pfad, e))
            return 1
        print("%s (%s): %d Zeilen, %d Versandpositionen weggelassen%s%s" % (
            pfad, enc, info["zeilen"], info["versand"],
            (", %d ohne lesbares Datum" % info["ohne_monat"]) if info["ohne_monat"] else "",
            (", %d aus anderem Monat" % info["anderer_monat"]) if info["anderer_monat"] else ""))
        print("    Spalten: %s" % ", ".join("%s=%d" % kv for kv in sorted(info["spalten"].items())))
        for k, a in agg.items():
            if k in gesamt:
                g = gesamt[k]
                g["menge"] += a["menge"]
                g["umsatz"] += a["umsatz"]
                g["rnr"] |= a["rnr"]
                g["titel"].update(a["titel"])
            else:
                gesamt[k] = a
    if arg.zeige_spalten:
        return 0
    if not gesamt:
        print("Nichts zu importieren.")
        return 1
    print("\n" + zusammenfassung(gesamt))
    if not any(a["umsatz"] for a in gesamt.values()):
        print("\nHINWEIS: Der Betrag ist ueberall 0 bzw. fehlt - die Auswertung zeigt dann keinen Umsatz "
              "(Zubehoer-Listen fuer den Amicron-Import haben absichtlich Betrag 0,00).")
    n = schreibe_historie(arg.ziel, gesamt, arg.trocken)
    print("\n%s %d Zeilen -> %s" % ("Trockenlauf:" if arg.trocken else "Geschrieben:", n, arg.ziel))
    return 0


if __name__ == "__main__":
    sys.exit(main())
