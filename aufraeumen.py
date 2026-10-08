# -*- coding: utf-8 -*-
"""
aufraeumen.py - raeumt die Ordner auf, die Carrier-Dashboard und Versand-Cockpit fuellen
============================================================================
Loescht NUR veraltete, nicht mehr benoetigte Dateien. Standard-Aufbewahrung (Tage, aenderbar):

  rechnungs_archiv  45  Rechnungs-PDFs im Dashboard-Archiv (<Archiv>\\JJJJ-MM-TT\\Packliste Nr ...pdf).
                        Eine PDF wird NUR geloescht, wenn ihre Rechnungsnummer schon in
                        artikel_verkaeufe.csv steht (Bestseller-Statistik) - sonst bleibt sie liegen
                        und das Protokoll sagt es (dann artikel_auswertung_archiv.py laufen lassen).
  pickliste_pdf     30  Pickliste_*.pdf im Ausgabeordner (die festen Bruecken-CSVs bleiben IMMER).
  carrier_csv       60  DHL_/DPD_/Post_*.csv im Carrier-Export-Ordner (bereits importiert).
  label_archiv      90  Label-/Briefmarken-PDFs im Netz-Ordner Paketscheine\\Archiv (scan_druck).
  sync_archiv       90  Datums-Ordner im Archiv des WooCommerce-Sendungsnummer-Syncs.
  dashboard_log          Carrier-Dashboard.log wird auf die letzten Zeilen gekuerzt, wenn > 1 MB.

NIE angefasst werden: carrier_statistik.csv, artikel_statistik.csv, artikel_verkaeufe.csv,
statistik.csv, gedruckt.log, sammel_/post_/mengen_/ean_zuordnung.csv, wc_bestellnummern*.csv,
Programm- und Konfigurationsdateien. Es wird nur in den ausdruecklich genannten Ordnern und nur
nach den genannten Dateimustern geloescht, nie ueber Verknuepfungen hinweg.

Aufruf:
    py aufraeumen.py                 (Trockenlauf: zeigt nur, was geloescht WUERDE)
    py aufraeumen.py --ausfuehren    (loescht wirklich)
    py aufraeumen.py --selftest
Das Dashboard ruft die Aufraeumung bei Bedarf einmal taeglich selbst auf, wenn in der
carrier_dashboard_config.json "aufraeumen": true steht (siehe carrier_dashboard.py).
"""

import os
import re
import shutil
import sys
from datetime import date, datetime, timedelta

VERSION = "2026-10-08a"

STANDARD_TAGE = {"rechnungs_archiv": 45, "pickliste_pdf": 30, "carrier_csv": 60,
                 "label_archiv": 90, "sync_archiv": 90}
MAX_LOESCHUNGEN_JE_REGEL = 20000      # Notbremse je Regel und Lauf
LOG_MAX_BYTES = 1000000
LOG_BEHALTEN_ZEILEN = 5000

GESCHUETZT = {"carrier_statistik.csv", "artikel_statistik.csv", "artikel_verkaeufe.csv",
              "statistik.csv", "gedruckt.log", "sammel_zuordnung.csv", "post_zuordnung.csv",
              "mengen_zuordnung.csv", "ean_zuordnung.csv", "wc_bestellnummern.csv",
              "wc_bestellnummern_dashboard.csv"}
_RE_PICKLISTE = re.compile(r"^Pickliste_.*\.pdf$", re.I)
_RE_CARRIER_CSV = re.compile(r"^(DHL|DPD|Post_Brief|Post_Grossbrief)_.*\.csv$", re.I)
_RE_RNR = re.compile(r"(?:Nr\.?\s*)(\d{7})", re.I)


def _datumsordner(name):
    """date, wenn name 'JJJJ-MM-TT' ist, sonst None."""
    if len(name) == 10 and name[4] == "-" and name[7] == "-":
        try:
            return datetime.strptime(name, "%Y-%m-%d").date()
        except ValueError:
            return None
    return None


def _root_ok(root):
    """Nur ein vorhandener, echter Ordner, nie ein Laufwerksstamm/zu kurzer Pfad/Verknuepfung."""
    if not root or not os.path.isdir(root) or os.path.islink(root):
        return False
    r = os.path.abspath(root)
    return len(r) > 12 and os.path.dirname(r) != r


def _unter(root, pfad):
    r = os.path.realpath(root)
    p = os.path.realpath(pfad)
    return p.startswith(r + os.sep)


class Bericht:
    def __init__(self):
        self.regeln = {}          # name -> dict(geloescht, bytes, behalten, notiz[])

    def regel(self, name):
        return self.regeln.setdefault(name, {"geloescht": 0, "bytes": 0, "behalten": 0, "notiz": []})

    def summe(self):
        return (sum(r["geloescht"] for r in self.regeln.values()),
                sum(r["bytes"] for r in self.regeln.values()))


def _loesche_datei(pfad, root, trocken, rg):
    if os.path.islink(pfad) or not _unter(root, pfad) or os.path.basename(pfad) in GESCHUETZT:
        rg["behalten"] += 1
        return False
    try:
        groesse = os.path.getsize(pfad)
        if not trocken:
            os.remove(pfad)
    except OSError as e:
        rg["notiz"].append("%s: %s" % (os.path.basename(pfad), e))
        return False
    rg["geloescht"] += 1
    rg["bytes"] += groesse
    return True


def regel_rechnungs_archiv(bericht, root, tage, jetzt, bekannte_rnr, trocken):
    rg = bericht.regel("rechnungs_archiv")
    if not _root_ok(root):
        rg["notiz"].append("Ordner nicht erreichbar: %s" % root)
        return
    if bekannte_rnr is None:
        rg["notiz"].append("artikel_verkaeufe.csv nicht lesbar - Rechnungs-Archiv wird NICHT aufgeraeumt")
        return
    grenze = jetzt.date() - timedelta(days=tage)
    for name in sorted(os.listdir(root)):
        d = _datumsordner(name)
        pfad = os.path.join(root, name)
        if d is None or d >= grenze or not os.path.isdir(pfad) or os.path.islink(pfad):
            continue
        offen = 0
        for datei in os.listdir(pfad):
            fp = os.path.join(pfad, datei)
            if not os.path.isfile(fp):
                offen += 1
                continue
            m = _RE_RNR.search(datei)
            if (datei.lower().endswith(".pdf") and m and m.group(1) in bekannte_rnr
                    and rg["geloescht"] < MAX_LOESCHUNGEN_JE_REGEL):
                _loesche_datei(fp, root, trocken, rg)
            else:
                offen += 1
                rg["behalten"] += 1
        if offen:
            rg["notiz"].append("%s: %d Datei(en) noch nicht in artikel_verkaeufe.csv bzw. unbekannter "
                               "Name - bleiben liegen" % (name, offen))
        elif not trocken:
            try:
                os.rmdir(pfad)
            except OSError:
                pass


def regel_dateien(bericht, name, root, muster, tage, jetzt, trocken):
    rg = bericht.regel(name)
    if not _root_ok(root):
        rg["notiz"].append("Ordner nicht erreichbar: %s" % root)
        return
    grenze = jetzt - timedelta(days=tage)
    for datei in sorted(os.listdir(root)):
        fp = os.path.join(root, datei)
        if not os.path.isfile(fp) or not muster.match(datei):
            continue
        try:
            alt = datetime.fromtimestamp(os.path.getmtime(fp)) < grenze
        except OSError:
            continue
        if alt and rg["geloescht"] < MAX_LOESCHUNGEN_JE_REGEL:
            _loesche_datei(fp, root, trocken, rg)
        else:
            rg["behalten"] += 1


def regel_datumsordner(bericht, name, root, tage, jetzt, trocken):
    """Ganze Datums-Unterordner (JJJJ-MM-TT) loeschen, die aelter als tage sind."""
    rg = bericht.regel(name)
    if not _root_ok(root):
        rg["notiz"].append("Ordner nicht vorhanden: %s" % root)
        return
    grenze = jetzt.date() - timedelta(days=tage)
    for dn in sorted(os.listdir(root)):
        d = _datumsordner(dn)
        pfad = os.path.join(root, dn)
        if d is None or d >= grenze or not os.path.isdir(pfad) or os.path.islink(pfad):
            continue
        for dp, _dirs, fns in os.walk(pfad):
            for f in fns:
                if rg["geloescht"] < MAX_LOESCHUNGEN_JE_REGEL:
                    _loesche_datei(os.path.join(dp, f), root, trocken, rg)
        if not trocken:
            shutil.rmtree(pfad, ignore_errors=True)


def kuerze_log(bericht, pfad, trocken):
    rg = bericht.regel("dashboard_log")
    try:
        if not pfad or not os.path.isfile(pfad) or os.path.getsize(pfad) <= LOG_MAX_BYTES:
            return
        with open(pfad, encoding="utf-8", errors="replace") as f:
            zeilen = f.readlines()
        alt = os.path.getsize(pfad)
        if not trocken:
            tmp = pfad + ".tmp"
            with open(tmp, "w", encoding="utf-8", newline="") as f:
                f.writelines(zeilen[-LOG_BEHALTEN_ZEILEN:])
            os.replace(tmp, pfad)
        rg["geloescht"] += max(0, len(zeilen) - LOG_BEHALTEN_ZEILEN)
        rg["bytes"] += max(0, alt - sum(len(z.encode("utf-8")) for z in zeilen[-LOG_BEHALTEN_ZEILEN:]))
    except OSError as e:
        rg["notiz"].append(str(e))


def lies_bekannte_rnr(verkaeufe_csv):
    """Menge der Rechnungsnummern aus artikel_verkaeufe.csv; None, wenn nicht lesbar/vorhanden."""
    if not verkaeufe_csv or not os.path.isfile(verkaeufe_csv):
        return None
    try:
        with open(verkaeufe_csv, encoding="utf-8") as f:
            next(f, None)
            return {z.split(";", 2)[1] for z in f if z.count(";") >= 2}
    except OSError:
        return None


def aufraeumen(pfade, tage=None, trocken=True, jetzt=None):
    """pfade: dict mit archiv, ausgabe, carrier_export, label_archiv, sync_archiv, verkaeufe_csv,
    dashboard_log (jeder Eintrag optional). tage ueberschreibt STANDARD_TAGE. Rueckgabe Bericht."""
    jetzt = jetzt or datetime.now()
    t = dict(STANDARD_TAGE)
    t.update(tage or {})
    b = Bericht()
    if pfade.get("archiv"):
        regel_rechnungs_archiv(b, pfade["archiv"], t["rechnungs_archiv"], jetzt,
                               lies_bekannte_rnr(pfade.get("verkaeufe_csv")), trocken)
    if pfade.get("ausgabe"):
        regel_dateien(b, "pickliste_pdf", pfade["ausgabe"], _RE_PICKLISTE, t["pickliste_pdf"], jetzt, trocken)
    if pfade.get("carrier_export"):
        regel_dateien(b, "carrier_csv", pfade["carrier_export"], _RE_CARRIER_CSV, t["carrier_csv"], jetzt, trocken)
    if pfade.get("label_archiv"):
        regel_dateien(b, "label_archiv", pfade["label_archiv"], re.compile(r".*\.pdf$", re.I),
                      t["label_archiv"], jetzt, trocken)
    if pfade.get("sync_archiv"):
        regel_datumsordner(b, "sync_archiv", pfade["sync_archiv"], t["sync_archiv"], jetzt, trocken)
    if pfade.get("dashboard_log"):
        kuerze_log(b, pfade["dashboard_log"], trocken)
    return b


def bericht_text(b, trocken):
    z = []
    for name, r in b.regeln.items():
        z.append("%-17s %s %5d Datei(en), %7.1f MB, %d behalten%s" % (
            name, "wuerde loeschen:" if trocken else "geloescht:      ", r["geloescht"],
            r["bytes"] / 1048576.0, r["behalten"], ("  | " + " | ".join(r["notiz"][:3])) if r["notiz"] else ""))
    n, byt = b.summe()
    z.append("Summe: %d Datei(en), %.1f MB %s" % (n, byt / 1048576.0, "wuerden geloescht" if trocken else "geloescht"))
    return "\n".join(z)


def pfade_aus_dashboard():
    import carrier_dashboard as d
    import carrier_statistik as cs
    base = os.path.dirname(os.path.abspath(d.__file__))
    return {
        "archiv": d.ARCHIV_ORDNER, "ausgabe": d.AUSGABE_ORDNER, "carrier_export": d.CARRIER_EXPORT_ORDNER,
        "label_archiv": os.path.join(d.BRUECKEN_ZIEL, "Archiv") if d.BRUECKEN_ZIEL else None,
        "sync_archiv": os.path.join(d.WC_SYNC_ORDNER, "Archiv") if d.WC_SYNC_ORDNER else None,
        "verkaeufe_csv": cs.ARTIKEL_VERKAEUFE_DATEI, "dashboard_log": os.path.join(base, "Carrier-Dashboard.log"),
    }


def _log(zeile, pfad):
    try:
        with open(pfad, "a", encoding="utf-8") as f:
            f.write("%s  %s\n" % (datetime.now().strftime("%Y-%m-%d %H:%M:%S"), zeile))
    except OSError:
        pass


def lauf_wenn_faellig(pfade, tage=None, jetzt=None, marker=None, log_pfad=None):
    """Fuer das Dashboard: hoechstens einmal je 20 Stunden WIRKLICH aufraeumen. Rueckgabe Bericht oder None."""
    jetzt = jetzt or datetime.now()
    if marker and os.path.isfile(marker):
        try:
            letzter = datetime.fromisoformat(open(marker, encoding="utf-8").read().strip())
            if jetzt - letzter < timedelta(hours=20):
                return None
        except (OSError, ValueError):
            pass
    b = aufraeumen(pfade, tage, trocken=False, jetzt=jetzt)
    if marker:
        try:
            with open(marker, "w", encoding="utf-8") as f:
                f.write(jetzt.isoformat())
        except OSError:
            pass
    if log_pfad:
        _log(bericht_text(b, False).replace("\n", " | "), log_pfad)
    return b


def selftest():
    import tempfile
    n_ok, fehler = 0, []

    def check(name, ist, soll):
        nonlocal n_ok
        if ist == soll:
            n_ok += 1
        else:
            fehler.append("%s: erwartet %r, war %r" % (name, soll, ist))

    tmp = tempfile.mkdtemp(prefix="aufraeumen_test_")
    jetzt = datetime(2026, 10, 8, 12, 0)
    try:
        arch = os.path.join(tmp, "Packlisten", "Archiv")
        for tag, rnrs in (("2026-07-01", ["1700001", "1700002"]), ("2026-08-20", ["1700003"]),
                          ("2026-09-20", ["1700004"]), ("2026-07-02", ["1700005"])):
            os.makedirs(os.path.join(arch, tag))
            for r in rnrs:
                open(os.path.join(arch, tag, "Packliste Nr %s.pdf" % r), "w").write("x")
        os.makedirs(os.path.join(arch, "kein-datum"))
        open(os.path.join(arch, "kein-datum", "Packliste Nr 1700009.pdf"), "w").write("x")
        verk = os.path.join(tmp, "artikel_verkaeufe.csv")
        open(verk, "w", encoding="utf-8").write("Datum;Rechnungsnummer;x\n2026-07-01;1700001;a\n"
                                                "2026-07-01;1700002;a\n2026-08-20;1700003;a\n")
        aus = os.path.join(tmp, "Packlisten", "Pickliste")
        os.makedirs(aus)
        for n in ("Pickliste_alt.pdf", "Pickliste_neu.pdf", "sammel_zuordnung.csv", "post_zuordnung.csv"):
            open(os.path.join(aus, n), "w").write("x")
        alt = (jetzt - timedelta(days=40)).timestamp()
        os.utime(os.path.join(aus, "Pickliste_alt.pdf"), (alt, alt))
        os.utime(os.path.join(aus, "sammel_zuordnung.csv"), (alt, alt))
        car = os.path.join(tmp, "Carrier_Export")
        os.makedirs(car)
        for n in ("DHL_2026-07-01_100000.csv", "DHL_2026-10-07_100000.csv", "wichtig.txt"):
            open(os.path.join(car, n), "w").write("x")
        for n in ("DHL_2026-07-01_100000.csv", "wichtig.txt"):
            a = (jetzt - timedelta(days=100)).timestamp()
            os.utime(os.path.join(car, n), (a, a))
        lab = os.path.join(tmp, "Paketscheine", "Archiv")
        os.makedirs(lab)
        for n, tagealt in (("Briefmarken.1Stk.pdf", 100), ("SHIPMENT_LABEL_1.pdf", 10), ("notiz.txt", 200)):
            open(os.path.join(lab, n), "w").write("x")
            a = (jetzt - timedelta(days=tagealt)).timestamp()
            os.utime(os.path.join(lab, n), (a, a))
        sy = os.path.join(tmp, "WC", "Archiv")
        os.makedirs(os.path.join(sy, "2026-05-01"))
        os.makedirs(os.path.join(sy, "2026-10-01"))
        open(os.path.join(sy, "2026-05-01", "x.csv"), "w").write("x")
        lg = os.path.join(tmp, "Carrier-Dashboard.log")
        open(lg, "w", encoding="utf-8").write("zeile\n" * 300000)
        pf = {"archiv": arch, "ausgabe": aus, "carrier_export": car, "label_archiv": lab,
              "sync_archiv": sy, "verkaeufe_csv": verk, "dashboard_log": lg}

        b = aufraeumen(pf, trocken=True, jetzt=jetzt)
        check("Trockenlauf loescht nichts", os.path.exists(os.path.join(arch, "2026-07-01", "Packliste Nr 1700001.pdf")), True)
        check("Trockenlauf: Rechnungs-Archiv 3 PDFs (1700001/2/3 bekannt+alt)", b.regeln["rechnungs_archiv"]["geloescht"], 3)

        b = aufraeumen(pf, trocken=False, jetzt=jetzt)
        r = b.regeln
        check("alte + bekannte Rechnungs-PDF geloescht",
              [os.path.exists(os.path.join(arch, d, n)) for d, n in (
                  ("2026-07-01", "Packliste Nr 1700001.pdf"), ("2026-08-20", "Packliste Nr 1700003.pdf"))], [False, False])
        check("alter Ordner ohne Rest entfernt", os.path.exists(os.path.join(arch, "2026-07-01")), False)
        check("alte PDF ohne Eintrag in artikel_verkaeufe bleibt (1700005)",
              os.path.exists(os.path.join(arch, "2026-07-02", "Packliste Nr 1700005.pdf")), True)
        check("junge PDF bleibt (1700004)", os.path.exists(os.path.join(arch, "2026-09-20", "Packliste Nr 1700004.pdf")), True)
        check("Ordner ohne Datumsnamen unberuehrt", os.path.exists(os.path.join(arch, "kein-datum", "Packliste Nr 1700009.pdf")), True)
        check("alte Pickliste weg, neue bleibt",
              [os.path.exists(os.path.join(aus, n)) for n in ("Pickliste_alt.pdf", "Pickliste_neu.pdf")], [False, True])
        check("Bruecken-CSVs (auch alt) bleiben IMMER",
              [os.path.exists(os.path.join(aus, n)) for n in ("sammel_zuordnung.csv", "post_zuordnung.csv")], [True, True])
        check("alte Carrier-CSV weg, neue + fremde Datei bleiben",
              [os.path.exists(os.path.join(car, n)) for n in ("DHL_2026-07-01_100000.csv", "DHL_2026-10-07_100000.csv", "wichtig.txt")],
              [False, True, True])
        check("Label-Archiv: nur alte PDF weg",
              [os.path.exists(os.path.join(lab, n)) for n in ("Briefmarken.1Stk.pdf", "SHIPMENT_LABEL_1.pdf", "notiz.txt")],
              [False, True, True])
        check("Sync-Archiv: alter Datumsordner weg, junger bleibt",
              [os.path.exists(os.path.join(sy, d)) for d in ("2026-05-01", "2026-10-01")], [False, True])
        check("Log gekuerzt auf letzte Zeilen", sum(1 for _ in open(lg, encoding="utf-8")), LOG_BEHALTEN_ZEILEN)
        check("verkaeufe.csv unberuehrt", os.path.exists(verk), True)
        # ohne artikel_verkaeufe: Rechnungs-Archiv wird NICHT angefasst
        os.makedirs(os.path.join(arch, "2026-06-01"))
        open(os.path.join(arch, "2026-06-01", "Packliste Nr 1700001.pdf"), "w").write("x")
        b2 = aufraeumen(dict(pf, verkaeufe_csv=os.path.join(tmp, "nix.csv")), trocken=False, jetzt=jetzt)
        check("ohne verkaeufe.csv bleibt das Archiv liegen",
              (os.path.exists(os.path.join(arch, "2026-06-01", "Packliste Nr 1700001.pdf")),
               b2.regeln["rechnungs_archiv"]["geloescht"]), (True, 0))
        # Laufwerksstamm/Fremdordner werden abgelehnt
        check("Laufwerksstamm abgelehnt", _root_ok("C:\\"), False)
        # lauf_wenn_faellig: zweiter Aufruf innerhalb 20 h tut nichts
        mk = os.path.join(tmp, "marker.txt")
        e1 = lauf_wenn_faellig(pf, jetzt=jetzt, marker=mk)
        e2 = lauf_wenn_faellig(pf, jetzt=jetzt + timedelta(hours=5), marker=mk)
        e3 = lauf_wenn_faellig(pf, jetzt=jetzt + timedelta(hours=21), marker=mk)
        check("faellig: 1. Lauf ja, nach 5 h nein, nach 21 h ja", (e1 is not None, e2 is None, e3 is not None), (True, True, True))
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
    if fehler:
        print("SELBSTTEST FEHLGESCHLAGEN (%d von %d):" % (len(fehler), n_ok + len(fehler)))
        for f in fehler:
            print("  - " + f)
        return 1
    print("Selbsttest OK (%d Pruefungen)" % n_ok)
    return 0


if __name__ == "__main__":
    if "--selftest" in sys.argv:
        sys.exit(selftest())
    ausfuehren = "--ausfuehren" in sys.argv
    pf = pfade_aus_dashboard()
    print("Aufraeumen (%s)  Version %s" % ("AUSFUEHREN" if ausfuehren else "TROCKENLAUF - nichts wird geloescht", VERSION))
    for k, v in pf.items():
        print("  %-15s %s" % (k, v))
    ber = aufraeumen(pf, trocken=not ausfuehren)
    print()
    print(bericht_text(ber, not ausfuehren))
    for name, r in ber.regeln.items():
        for n in r["notiz"][:10]:
            print("  [%s] %s" % (name, n))
    sys.exit(0)
