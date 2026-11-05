# -*- coding: utf-8 -*-
"""
artikel_upload.py  --  laeuft auf dem Lager-PC (DESKTOP-N2H75H), wie statistik_upload.py
============================================================================
Verdichtet die Artikel-Verkaeufe zu einer kompakten artikel.json. QUELLE: artikel_amicron.csv (Amicron-Export, tages-
genau, Stueck laut Stueckliste - siehe artikel_import_amicron.py); fuer Rechnungen, die dort noch NICHT stehen (neuer
als der letzte Export), springt die Auswertung auf die Dashboard-Daten (artikel_verkaeufe.csv) ein und rechnet
Mehrfachpackungen mit der Stueckfaktor-Tabelle auf den Einzelartikel um (BP-GF-24 -> 24 x BP-GF). Ergebnis:
artikel.json (Tageswerte je Artikel: Bestellungen / Menge / Umsatz) und laedt sie per HTTPS-POST an
vs-ingest.php (Ziel "artikel"). Die Seite versandstatistik/index.html zeigt daraus den Reiter
"Top-Artikel" (frei waehlbarer Zeitraum, Top 20, Vorjahresvergleich, Artikelverlauf).

Stueckfaktoren (artikel_stueckfaktor.csv: Verkaufseinheit -> Stueck, optional Produktgruppe) werden je
Artikel mitgeschickt; die Seite rechnet damit "Stueck" (z.B. 1 x BP-GF-24 = 24 Stueck).

Monatsdaten FRUEHERER Zeitraeume (Amicron-Export, siehe artikel_import_amicron.py) liegen in
artikel_historie.csv und werden fuer alle Monate VOR dem ersten Tag mit Tagesdaten dazugenommen.

Nur Standardbibliothek. Konfiguration (alles optional) in artikel_upload_config.json neben diesem Skript:
  { "url": "https://www.gasecenter-onlineshop.de/vs-ingest.php",
    "token": "...",                      (sonst wird UPLOAD_TOKEN aus statistik_upload.py gelesen)
    "verkaeufe_csv": "...", "historie_csv": "...", "log": "...",
    "komponenten_csv": "...", "zusatz_csv": "...",
    "kopie_ordner": "G:\\Mein Drive\\Versandstatistik" }

Ist "kopie_ordner" gesetzt (alter Name "onedrive_ordner" gilt weiter), wird artikel.json zusaetzlich dort abgelegt
(bei jeder Aenderung bzw. wenn die Kopie fehlt, unabhaengig vom Upload) - gedacht fuer einen synchronisierten Ordner
(Google Drive fuer Desktop). Daraus liest das woechentliche Briefing (Top-10 Vorwoche / Prognose 4 Wochen) ueber
Google Drive - ohne Zugangsdaten zum passwortgeschuetzten Webspace.

Umrechnung der Dashboard-Zeilen (fuer Rechnungen, die im Amicron-Export noch fehlen): Mehrfachpackungen und
Slave-Artikel werden mit den Tabellen artikel_komponenten.csv (gelernt) und artikel_komponenten_zusatz.csv (abgeleitet,
siehe artikel_sets_ableiten.py) auf den Einzelartikel/Master umgerechnet; Sets aus verschiedenen Teilen (Doppelschlauch,
Set 396) bleiben als Set stehen. Ohne die Tabellen gilt nur die kleine Stueckfaktor-Tabelle wie bisher.

Aufruf (Aufgabenplanung, z.B. alle 30 Minuten oder taeglich):
    py artikel_upload.py                 (verdichten + hochladen, nur wenn sich etwas geaendert hat)
    py artikel_upload.py --dry-run --out C:\\Temp\\artikel.json     (nur lokal erzeugen, kein Upload)
    py artikel_upload.py --force         (auch ohne Aenderung hochladen)
"""

import argparse
import hashlib
import json
import os
import re
import ssl
import sys
import urllib.error
import urllib.request
from collections import defaultdict
from datetime import datetime, timedelta

VERSION = "2026-11-05a"
BASIS = os.path.dirname(os.path.abspath(__file__))
STANDARD = {
    "url": "https://www.gasecenter-onlineshop.de/vs-ingest.php",
    "verkaeufe_csv": r"\\DESKTOP-N2H75H\Netzwerk\Paketscheine\artikel_verkaeufe.csv",
    "amicron_csv": r"\\DESKTOP-N2H75H\Netzwerk\Paketscheine\artikel_amicron.csv",
    "historie_csv": r"\\DESKTOP-N2H75H\Netzwerk\Paketscheine\artikel_historie.csv",
    "stueckfaktor_csv": r"\\DESKTOP-N2H75H\Netzwerk\Paketscheine\artikel_stueckfaktor.csv",
    "komponenten_csv": r"\\DESKTOP-N2H75H\Netzwerk\Paketscheine\artikel_komponenten.csv",
    "zusatz_csv": r"\\DESKTOP-N2H75H\Netzwerk\Paketscheine\artikel_komponenten_zusatz.csv",
    "log": r"C:\Packlisten\artikel_upload.log",
}
STATE = os.path.join(BASIS, "artikel_upload_state.json")


def lade_config():
    cfg = dict(STANDARD)
    pfad = os.path.join(BASIS, "artikel_upload_config.json")
    if os.path.exists(pfad):
        with open(pfad, encoding="utf-8-sig") as f:
            cfg.update({k: v for k, v in json.load(f).items() if isinstance(v, str) and v.strip()})
    if not cfg.get("token"):
        # Token aus dem bestehenden statistik_upload.py uebernehmen (gleicher Empfaenger)
        for kand in (os.path.join(BASIS, "statistik_upload.py"), r"C:\Packlisten\statistik_upload.py"):
            if os.path.exists(kand):
                m = re.search(r'^UPLOAD_TOKEN\s*=\s*"(.*)"\s*$', open(kand, encoding="utf-8").read(), re.M)
                if m:
                    cfg["token"] = m.group(1)
                    break
    return cfg


def num(txt):
    try:
        return float((txt or "").strip().replace(",", ".") or 0)
    except ValueError:
        return 0.0


def lies_verkaeufe(pfad):
    """Zeilen (datum, rnr, art, bez, menge, umsatz) aus artikel_verkaeufe.csv."""
    zeilen = []
    if not os.path.exists(pfad):
        return zeilen
    with open(pfad, encoding="utf-8-sig") as f:
        next(f, None)
        for z in f:
            t = z.rstrip("\n").split(";")
            if len(t) >= 8 and t[0]:
                zeilen.append((t[0], t[1], t[2], t[3], num(t[4]), num(t[7])))
    return zeilen


def lies_amicron(pfad):
    """Zeilen (datum, rnr, art, bez, menge, umsatz) aus artikel_amicron.csv (Datum;Rechnungsnummer;Artikelnummer;
    Bezeichnung;Menge;Umsatz)."""
    zeilen = []
    if not os.path.exists(pfad):
        return zeilen
    with open(pfad, encoding="utf-8-sig") as f:
        next(f, None)
        for z in f:
            t = z.rstrip("\n").split(";")
            if len(t) >= 6 and t[0]:
                zeilen.append((t[0], t[1], t[2], t[3], num(t[4]), num(t[5])))
    return zeilen


def lies_faktor_tabelle(pfad):
    """{art_klein: (faktor, gruppe, bezeichnung, artikelnummer_original)} aus artikel_stueckfaktor.csv."""
    out = {}
    if not os.path.exists(pfad):
        return out
    with open(pfad, encoding="utf-8-sig") as f:
        next(f, None)
        for z in f:
            t = z.rstrip("\r\n").split(";")
            if len(t) >= 3 and t[0].strip():
                try:
                    fk = float(t[2].strip().replace(",", "."))
                except ValueError:
                    continue
                if fk > 0:
                    out[t[0].strip().lower()] = (fk, t[3].strip() if len(t) > 3 else "", t[1].strip(), t[0].strip())
    return out


def nrm(a):
    return re.sub(r"[^a-z0-9]", "", (a or "").lower())


# Schwellen wie in artikel_import_amicron.py (gelernte Tabelle): mindestens so viele Rechnungen mit Set UND
# Bestandteil, und mindestens dieser Anteil der Rechnungen mit dem Set
MIN_STUETZE = 2
MIN_ANTEIL = 0.6


def lies_komponenten_tabelle(pfad):
    """{set: {bestandteil: faktor}} aus artikel_komponenten.csv bzw. artikel_komponenten_zusatz.csv
    (Set;Komponente;RechnungenKomponente;RechnungenSet;Verhaeltnis), nur Beziehungen, die die Schwellen erreichen."""
    n_set, zeilen = {}, []
    if not pfad or not os.path.exists(pfad):
        return {}
    with open(pfad, encoding="utf-8-sig") as f:
        next(f, None)
        for z in f:
            t = z.rstrip("\r\n").split(";")
            if len(t) < 5 or not t[0] or t[0] == "*":
                continue
            try:
                rx, ns, vh = int(t[2]), int(t[3]), float(t[4].replace(",", "."))
            except ValueError:
                continue
            n_set[t[0]] = max(n_set.get(t[0], 0), ns)
            zeilen.append((t[0], t[1], rx, vh))
    akt = {}
    for s, x, rx, vh in zeilen:
        if rx >= MIN_STUETZE and n_set[s] and rx / n_set[s] >= MIN_ANTEIL:
            akt.setdefault(s, {})[x] = vh
    return akt


def lade_einzel(komp_pfad, zusatz_pfad, faktoren):
    """{set: (einzelartikel, faktor)} fuer Sets mit genau EINEM Bestandteil = Mehrfachpackung bzw. Slave -> Master
    (Matthias: der Einzelartikel/Master zaehlt, z.B. BP-GF-24 = 24 x BP-GF). Sets mit mehreren verschiedenen
    Bestandteilen (Doppelschlauch, Set 396 ...) bleiben als Set stehen. Reihenfolge: gelernte Tabelle, Zusatz-Tabelle,
    Stueckfaktor-Tabelle (nur ergaenzend, gelernte Beziehungen haben Vorrang)."""
    komp = lies_komponenten_tabelle(komp_pfad)
    for s, v in lies_komponenten_tabelle(zusatz_pfad).items():
        for x, f in v.items():
            komp.setdefault(s, {}).setdefault(x, f)
    basis = {}
    for art, (fk, gr, bez, orig) in faktoren.items():
        if gr and fk == 1 and gr.lower() not in basis:
            basis[gr.lower()] = orig
    for art, (fk, gr, bez, orig) in faktoren.items():
        if fk > 1 and gr and gr.lower() in basis and nrm(basis[gr.lower()]) != nrm(orig):
            komp.setdefault(nrm(orig), {}).setdefault(nrm(basis[gr.lower()]), fk)
    return {s: next(iter(v.items())) for s, v in komp.items() if len(v) == 1}


def kombiniere(amicron, dashboard, faktoren, einzel=None):
    """Amicron-Zeilen (Stueck laut Stueckliste) plus Dashboard-Zeilen NUR fuer Rechnungen, die in amicron fehlen.
    Dashboard-Zeilen zeigen den Artikel, wie er auf der Rechnung steht (z.B. die Packung BP-GF-24). Sie werden
    - mit der Tabelle `einzel` (lade_einzel) auf den Einzelartikel/Master umgerechnet, soweit dessen Schreibweise aus
      den Daten bekannt ist (hoechstens zwei Stufen), danach
    - mit der Stueckfaktor-Tabelle (Faktor > 1 und Einzelartikel mit Faktor 1 in derselben Gruppe);
    sonst bleibt die Zeile wie sie ist."""
    einzel = einzel or {}
    bekannt = {z[1] for z in amicron}
    namen = {}                                           # normierte Nr. -> (Schreibweise, Bezeichnung)
    for z in amicron:
        namen.setdefault(nrm(z[2]), (z[2], z[3]))
    for z in dashboard:
        namen.setdefault(nrm(z[2]), (z[2], z[3]))
    for art, (fk, gr, bez, orig) in faktoren.items():
        namen.setdefault(nrm(orig), (orig, bez))
    basis = {}
    for art, (fk, gr, bez, orig) in faktoren.items():
        if gr and fk == 1 and gr.lower() not in basis:
            basis[gr.lower()] = (orig, bez)
    out = list(amicron)
    umgerechnet = 0
    for datum, rnr, art, bez, menge, umsatz in dashboard:
        if rnr in bekannt:
            continue
        s, m, ziel = nrm(art), menge, None
        for _ in range(2):
            e = einzel.get(s)
            if not e or e[0] == s or e[0] not in namen:
                break
            s, m, ziel = e[0], m * e[1], e[0]
        if ziel:
            out.append((datum, rnr, namen[ziel][0], namen[ziel][1] or bez, m, umsatz))
            umgerechnet += 1
            continue
        fk, gr = faktoren.get(art.lower(), (1, "", "", ""))[:2]
        if fk > 1 and gr and gr.lower() in basis and basis[gr.lower()][0].lower() != art.lower():
            b_art, b_bez = basis[gr.lower()]
            out.append((datum, rnr, b_art, b_bez or bez, menge * fk, umsatz))
            umgerechnet += 1
        else:
            out.append((datum, rnr, art, bez, menge, umsatz))
    return out, umgerechnet


def lies_historie(pfad):
    """Zeilen (monat 'JJJJ-MM', art, bez, bestellungen|None, menge, umsatz) aus artikel_historie.csv."""
    zeilen = []
    if not os.path.exists(pfad):
        return zeilen
    with open(pfad, encoding="utf-8-sig") as f:
        next(f, None)
        for z in f:
            t = z.rstrip("\n").split(";")
            if len(t) >= 6 and re.fullmatch(r"\d{4}-\d{2}", t[0]):
                best = int(t[3]) if t[3].strip().isdigit() else None
                zeilen.append((t[0], t[1], t[2], best, num(t[4]), num(t[5])))
    return zeilen


def lies_faktoren(pfad):
    """{artikelnummer_klein: (faktor, gruppe)} aus artikel_stueckfaktor.csv (Artikelnummer;Bezeichnung;Faktor;Gruppe;Quelle)."""
    out = {}
    if not os.path.exists(pfad):
        return out
    with open(pfad, encoding="utf-8-sig") as f:
        next(f, None)
        for z in f:
            t = z.rstrip("\r\n").split(";")
            if len(t) >= 3 and t[0].strip():
                try:
                    fk = float(t[2].strip().replace(",", "."))
                except ValueError:
                    continue
                if fk > 0:
                    out[t[0].strip().lower()] = (fk, t[3].strip() if len(t) > 3 else "")
    return out


TAGE_DETAIL = 430          # so viele Tage bleiben tagesgenau; Aelteres wird zu Monatswerten verdichtet


def baue(verkaeufe, historie, jetzt=None, faktoren=None):
    """Kompaktes JSON-Objekt: artikel [[nr, bez]], tage {datum: [[idx, best, menge, umsatz]]},
    monate {JJJJ-MM: [[idx, best|None, menge, umsatz]]} (nur Monate VOR dem ersten Tag mit Tagesdaten)."""
    jetzt = jetzt or datetime.now()
    idx, artikel, bez_datum = {}, [], {}

    def nr_idx(art, bez, datum):
        k = art.strip().lower()
        if k not in idx:
            idx[k] = len(artikel)
            artikel.append([art.strip(), bez])
            bez_datum[k] = datum
        elif datum >= bez_datum[k] and bez:
            artikel[idx[k]][1] = bez                  # aktuellste Bezeichnung gewinnt
            bez_datum[k] = datum
        return idx[k]

    tage = defaultdict(lambda: defaultdict(lambda: [set(), 0.0, 0.0]))
    grenze = (jetzt - timedelta(days=TAGE_DETAIL)).strftime("%Y-%m-%d")
    alt_monate = defaultdict(lambda: defaultdict(lambda: [set(), 0.0, 0.0]))     # Monatsverdichtung aelterer Tage
    for datum, rnr, art, bez, menge, umsatz in verkaeufe:
        i = nr_idx(art, bez, datum)
        a = alt_monate[datum[:7]][i] if datum < grenze else tage[datum][i]
        a[0].add(rnr)
        a[1] += menge
        a[2] += umsatz
    erster_tag = min(tage) if tage else None
    if erster_tag and erster_tag[8:] != "01":          # Tagesdaten fangen mitten im Monat an: diesen Monat noch als Monat fuehren
        for d in [d for d in tage if d[:7] == erster_tag[:7]]:
            for i, v in tage[d].items():
                a = alt_monate[d[:7]][i]
                a[0] |= v[0]
                a[1] += v[1]
                a[2] += v[2]
            del tage[d]
        erster_tag = min(tage) if tage else None
    # Faengt die Tagesdaten mitten im Monat an (z.B. 16.07., weil das Archiv dort beginnt), liefert die
    # Amicron-Monatshistorie den GANZEN Monat: dann zaehlen fuer diesen Monat NUR die Monatswerte.
    if erster_tag and erster_tag[8:] != "01" and any(h[0] == erster_tag[:7] for h in historie):
        for d in [d for d in tage if d[:7] == erster_tag[:7]]:
            del tage[d]
        erster_tag = min(tage) if tage else None
    monate = defaultdict(lambda: defaultdict(lambda: [None, 0.0, 0.0]))
    for mo, per_art in alt_monate.items():            # aus Tagesdaten verdichtete Monate
        for i, v in per_art.items():
            m = monate[mo][i]
            m[0] = len(v[0])
            m[1] += v[1]
            m[2] += v[2]
    for monat, art, bez, best, menge, umsatz in historie:
        if erster_tag and monat >= erster_tag[:7]:
            continue                                   # ab dem ersten Tagesdaten-Monat zaehlen die Tageswerte
        i = nr_idx(art, bez, monat + "-00")
        a = monate[monat][i]
        a[0] = (a[0] or 0) + best if best is not None else a[0]
        a[1] += menge
        a[2] += umsatz

    def r2(x):
        return round(x, 2) if x != int(x) else int(x)
    faktoren = faktoren or {}
    for a in artikel:                              # [nr, bez] -> [nr, bez, faktor, gruppe]
        fk, gr = faktoren.get(a[0].lower(), (1, ""))
        a.append(int(fk) if fk == int(fk) else fk)
        a.append(gr)
    return {
        "generated": jetzt.strftime("%d.%m.%Y %H:%M"),
        "erster_tag": erster_tag,
        "artikel": artikel,
        "tage": {d: [[i, len(v[0]), r2(v[1]), r2(v[2])] for i, v in sorted(m.items())]
                 for d, m in sorted(tage.items())},
        "monate": {mo: [[i, v[0], r2(v[1]), r2(v[2])] for i, v in sorted(m.items())]
                   for mo, m in sorted(monate.items())},
    }


def log(msg, pfad):
    zeile = "%s  %s" % (datetime.now().strftime("%Y-%m-%d %H:%M:%S"), msg)
    print(zeile)
    try:
        with open(pfad, "a", encoding="utf-8") as f:
            f.write(zeile + "\n")
    except OSError:
        pass


def kopie_ablegen(roh, ordner, lg):
    """artikel.json atomar in den Kopie-Ordner (z.B. Google-Drive-Ordner) schreiben (tmp + replace).
    Fehler stoeren den Upload nie."""
    try:
        os.makedirs(ordner, exist_ok=True)
        ziel = os.path.join(ordner, "artikel.json")
        tmp = ziel + ".tmp"
        with open(tmp, "wb") as f:
            f.write(roh)
        os.replace(tmp, ziel)
        log("Kopie OK -> %s" % ziel, lg)
    except OSError as e:
        log("FEHLER Kopie (%s): %s" % (ordner, e), lg)


def hochladen(roh, url, token):
    req = urllib.request.Request(url + ("&" if "?" in url else "?") + "ziel=artikel", data=roh, method="POST",
                                 headers={"Content-Type": "application/json; charset=utf-8",
                                          "X-Upload-Token": token})
    with urllib.request.urlopen(req, timeout=90, context=ssl.create_default_context()) as r:
        return r.status, r.read().decode("utf-8", "replace")[:200]


def selftest():
    import tempfile
    ok, fehler = 0, []

    def check(name, ist, soll):
        nonlocal ok
        if ist == soll:
            ok += 1
        else:
            fehler.append("%s: erwartet %r, war %r" % (name, soll, ist))
    tmp = tempfile.mkdtemp()
    try:
        kopf = "Set;Komponente;RechnungenKomponente;RechnungenSet;Verhaeltnis\n"
        gel = os.path.join(tmp, "k.csv")
        zus = os.path.join(tmp, "z.csv")
        with open(gel, "w", encoding="utf-8") as f:
            f.write(kopf + "bpgf24;bpgf;10;10;24,0000\nset396;t244;10;10;1,0000\nset396;s245;10;10;1,0000\n"
                    "selten;bpgf;1;10;5,0000\n*;rein;9;0;0\nab2;a;10;10;1,0000\n")
        with open(zus, "w", encoding="utf-8") as f:
            f.write(kopf + "slave1;master1;6;6;1,0000\nab2;b;5;5;1,0000\n")
        einzel = lade_einzel(gel, zus, {})
        check("Packung -> Einzelartikel (gelernte Tabelle)", einzel.get("bpgf24"), ("bpgf", 24.0))
        check("Slave -> Master aus der Zusatz-Tabelle", einzel.get("slave1"), ("master1", 1.0))
        check("Set aus 2 Bestandteilen bleibt Set", "set396" in einzel, False)
        check("zu wenig Stuetze (1 von 10) zaehlt nicht", "selten" in einzel, False)
        check("Zusatz ergaenzt einen zweiten Bestandteil -> wird Set", "ab2" in einzel, False)
        dash = [("2026-11-02", "R1", "BP-GF-24", "GermanFire 24", 2, 30.0), ("2026-11-02", "R2", "SET396", "Set", 1, 33.99),
                ("2026-11-02", "R3", "SLAVE1", "Slave", 3, 9.0), ("2026-11-02", "R4", "NEU-99", "Neu", 1, 5.0)]
        amic = [("2026-10-01", "A1", "BP-GF", "Sicherheitsbrennpaste", 24, 12.0),
                ("2026-10-01", "A2", "Master1", "Master Artikel", 1, 3.0)]
        out, n = kombiniere(amic, dash, {}, einzel)
        zeilen = {z[1]: z for z in out}
        check("BP-GF-24 x2 -> 48 x BP-GF, Umsatz bleibt", (zeilen["R1"][2], zeilen["R1"][4], zeilen["R1"][5]), ("BP-GF", 48, 30.0))
        check("Set 396 bleibt Set", zeilen["R2"][2], "SET396")
        check("Slave -> Master in der Schreibweise des Bestands", (zeilen["R3"][2], zeilen["R3"][4]), ("Master1", 3))
        check("unbekannter Artikel bleibt", zeilen["R4"][2], "NEU-99")
        check("Anzahl umgerechnet", n, 2)
        out2, _ = kombiniere([], dash, {}, {"bpgf24": ("bpgf", 24.0)})
        check("Schreibweise des Einzelartikels unbekannt -> keine Umrechnung", [z for z in out2 if z[1] == "R1"][0][2], "BP-GF-24")
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
    ap = argparse.ArgumentParser(description="Artikel-Statistik hochladen")
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--out", default=None)
    ap.add_argument("--force", action="store_true")
    ap.add_argument("--selftest", action="store_true")
    arg = ap.parse_args(argv)
    if arg.selftest:
        return selftest()
    cfg = lade_config()
    lg = cfg["log"]
    try:
        dash = lies_verkaeufe(cfg["verkaeufe_csv"])
        amic = lies_amicron(cfg["amicron_csv"])
        hist = lies_historie(cfg["historie_csv"])
        fak = lies_faktor_tabelle(cfg["stueckfaktor_csv"])
        einzel = lade_einzel(cfg["komponenten_csv"], cfg["zusatz_csv"], fak)
        verk, n_umgerechnet = kombiniere(amic, dash, fak, einzel)
    except OSError as e:
        log("FEHLER beim Lesen: %s" % e, lg)
        return 1
    if not verk and not hist:
        log("keine Daten gefunden (%s)" % cfg["verkaeufe_csv"], lg)
        return 1
    obj = baue(verk, hist)
    roh = json.dumps(obj, ensure_ascii=False, separators=(",", ":")).encode("utf-8")
    pruef = hashlib.sha256(json.dumps({k: v for k, v in obj.items() if k != "generated"},
                                      sort_keys=True).encode("utf-8")).hexdigest()
    info = "%d Artikel, %d Tage, %d Monate, %.0f KB (Amicron %d Zeilen, Dashboard-Ersatz fuer neuere Rechnungen: %d Zeilen, davon %d Packungen umgerechnet)" % (
        len(obj["artikel"]), len(obj["tage"]), len(obj["monate"]), len(roh) / 1024, len(amic),
        len(verk) - len(amic), n_umgerechnet)
    if arg.dry_run:
        if arg.out:
            with open(arg.out, "wb") as f:
                f.write(roh)
        print("Trockenlauf: %s%s" % (info, (" -> " + arg.out) if arg.out else ""))
        return 0
    try:
        alt = json.load(open(STATE, encoding="utf-8")).get("hash")
    except (OSError, ValueError):
        alt = None
    od = cfg.get("kopie_ordner") or cfg.get("onedrive_ordner")       # alter Name bleibt als Alias
    if od and (alt != pruef or arg.force or not os.path.exists(os.path.join(od, "artikel.json"))):
        kopie_ablegen(roh, od, lg)
    if alt == pruef and not arg.force:
        print("unveraendert - kein Upload (%s)" % info)
        return 0
    if not cfg.get("token"):
        log("FEHLER: kein Upload-Token (artikel_upload_config.json oder statistik_upload.py)", lg)
        return 1
    try:
        status, antwort = hochladen(roh, cfg["url"], cfg["token"])
    except (urllib.error.URLError, OSError) as e:
        log("FEHLER beim Upload: %s" % e, lg)
        return 1
    log("Upload OK (%s, HTTP %s: %s)" % (info, status, antwort), lg)
    try:
        json.dump({"hash": pruef}, open(STATE, "w", encoding="utf-8"))
    except OSError:
        pass
    return 0


if __name__ == "__main__":
    sys.exit(main())
