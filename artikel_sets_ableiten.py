# -*- coding: utf-8 -*-
"""
artikel_sets_ableiten.py - fehlende Set->Bestandteil-Beziehungen aus den Amicron-Exporten ableiten
=================================================================================================
Hintergrund: artikel_import_amicron.py entfernt Stuecklisten-Bestandteile anhand einer GELERNTEN Tabelle
(artikel_komponenten.csv, gelernt aus Rechnungen mit Dashboard-Wahrheit = erst ab 16.07.2026). Fuer 2025/H1 2026
kennt sie viele Sets nicht (z.B. MB25-08-10 = 10 x MB25-SD08). Dieses Skript findet solche Paare in den Exporten selbst:

  Set X (Verkaufszeile mit Preis) und Artikel Y stehen in derselben Rechnung, Menge(Y) = r x Menge(X) mit r >= 1,
  und X kommt (fast) nur zusammen mit Y vor (>= MIN_ANTEIL der Rechnungen mit X, mindestens MIN_STUETZE Rechnungen).

Nur Artikel Y, die nach dem normalen Import als "verdaechtig" uebrig bleiben (Titel mit "Master", nicht erkannt),
werden betrachtet - das haelt die Zahl falscher Treffer klein. Ergebnis: artikel_komponenten_zusatz.csv (gleiches
Format wie artikel_komponenten.csv), das der Importer zusaetzlich zur gelernten Tabelle liest. Die Datei darf von
Hand ergaenzt/bereinigt werden und ueberlebt ein neues Lernen.

Aufruf:
    py artikel_sets_ableiten.py export1.csv export2.csv ...            (Vorschau, nichts schreiben)
    py artikel_sets_ableiten.py export1.csv ... --schreibe             (artikel_komponenten_zusatz.csv schreiben)
    py artikel_sets_ableiten.py export1.csv ... --slave-master         (zusaetzlich 1:1 Slave -> Master: Master zaehlt)
Optionen: --zusatz DATEI, --min-anteil 0.8, --min-stuetze 3

Master/Slave (Matthias 2026-11-05): Mehrere Artikelnummern derselben Ware; nur der Master hat Bestandsfuehrung und
Bestelldetails, ein verkaufter Slave leitet den Umsatz per Stueckliste auf den Master. In der Statistik zaehlt der
MASTER (Variante A). Zusammengesetzte Produkte (Doppelschlauch, Set 396 ...) zaehlen dagegen als Set.
"""

import argparse
import collections
import os
import re
import sys

import artikel_import_amicron as imp

VERSION = "2026-11-05a"
MIN_STUETZE = 3
MIN_ANTEIL = 0.8
ZUSATZ_HEADER = "Set;Komponente;RechnungenKomponente;RechnungenSet;Verhaeltnis\n"


def lade_rechnungen(dateien):
    """-> {rnr: [Zeilen]} ohne Gutschriften/AB, Versand und Zeilen ohne Artikelnummer; Mengen je Artikel addiert."""
    rg = {}
    for pfad in dateien:
        text, _ = imp.lies_text(pfad)
        zeilen, _info = imp.parse_export(text)
        je = collections.defaultdict(dict)
        for l in zeilen:
            if not l["rnr"]:
                continue
            k = imp.nrm(l["art"])
            a = je[l["rnr"]].setdefault(k, {"art": l["art"], "titel": l["titel"], "menge": 0.0, "betrag": 0.0})
            a["menge"] += l["menge"]
            a["betrag"] += l["betrag"]
        rg.update(je)
    return rg


def verdaechtige(dateien, komp):
    """normalisierte Artikelnummern, die nach dem normalen Import als verdaechtig uebrig bleiben (+ Stueck)."""
    erg = collections.Counter()
    for pfad in dateien:
        text, _ = imp.lies_text(pfad)
        _agg, info = imp.lies_export(text, None, komp)
        for l in info["zeilen_final"]:
            if not l.get("fest") and (l["betrag"] == 0 or re.search(r"\bmaster\b", l["titel"], re.I)):
                erg[imp.nrm(l["art"])] += l["menge"]
    return erg


def ist_vielfachpackung(r):
    """Ganzzahliges Vielfaches ab 2 (5er-, 10er-, 24er-Packung ...). Faktor 1 und krumme Faktoren sind unsicher:
    dort kann es sich auch um zwei oft zusammen gekaufte Artikel handeln -> nur zur Pruefung anzeigen."""
    return r >= 2 - 1e-9 and abs(r - round(r)) < 1e-6


def titel_aehnlichkeit(a, b):
    """0..1: wie aehnlich sind zwei Artikeltitel (ohne 'Master', Gross-/Kleinschreibung egal)."""
    import difflib

    def norm(t):
        return " ".join(re.sub(r"\bmaster\b", " ", t.lower()).split())

    def woerter(t):
        return set(re.findall(r"[a-zäöüß0-9]{3,}", norm(t)))
    wa, wb = woerter(a), woerter(b)
    # Wortmengen-Anteil (Reihenfolge egal: "1 kg WSG2 Schweissstaebe" ~ "Schweissstaebe WSG2 ... kg") oder Zeichenfolge
    wort = len(wa & wb) / min(len(wa), len(wb)) if wa and wb else 0.0
    return max(wort, difflib.SequenceMatcher(None, norm(a), norm(b)).ratio())


def ist_slave_master(t, anzahl_je_x, min_stuetze=5, min_anteil=0.95, min_aehnlich=0.45):
    """1:1-Beziehung Slave X -> Master Y (dieselbe Ware, andere Artikelnummer): Faktor 1, fast immer zusammen,
    X hat nur diesen einen Bestandteil und die Titel sind aehnlich (sonst ist es ein Set aus verschiedenen Produkten)."""
    _s, x, _y, r, n_paar, n_x, tx, ty = t
    return (abs(r - 1.0) < 1e-6 and n_paar >= min_stuetze and n_paar / n_x >= min_anteil
            and anzahl_je_x[x] == 1 and titel_aehnlichkeit(tx, ty) >= min_aehnlich)


def leite_ab(rg, verd, min_stuetze=MIN_STUETZE, min_anteil=MIN_ANTEIL, bekannt=None):
    """-> Liste (stueck, X, Y, r, n_paar, n_x, titel_x, titel_y) sortiert nach Auswirkung."""
    bekannt = bekannt or {}
    idx = collections.defaultdict(set)
    for rnr, arts in rg.items():
        for k in arts:
            idx[k].add(rnr)
    erg = []
    for y in verd:
        paare = collections.defaultdict(list)             # (x, r) -> [rnr]
        for rnr in idx.get(y, ()):
            ay = rg[rnr][y]
            if ay["menge"] <= 0:
                continue
            for x, ax in rg[rnr].items():
                if x == y or ax["betrag"] <= 0 or ax["menge"] <= 0:
                    continue
                if imp.ist_versand(ax["art"], ax["titel"]) or ax["art"] == "660":
                    continue
                r = ay["menge"] / ax["menge"]
                if r < 1 - 1e-9:
                    continue
                paare[(x, round(r, 3))].append(rnr)
        for (x, r), rl in paare.items():
            n_x = len(idx[x])
            if len(rl) >= min_stuetze and len(rl) / n_x >= min_anteil and y not in bekannt.get(x, ()):
                stueck = sum(rg[rn][y]["menge"] for rn in rl)
                erg.append((stueck, x, y, r, len(rl), n_x, rg[rl[0]][x]["titel"], rg[rl[0]][y]["titel"]))
    # gegenseitige Paare (X enthaelt Y UND Y enthaelt X, z.B. Bier- und CO2-Schlauch als Zweierset) sind keine
    # Stueckliste, sondern zwei Artikel, die oft zusammen gekauft werden -> beide verwerfen
    paare = {(e[1], e[2]) for e in erg}
    erg = [e for e in erg if (e[2], e[1]) not in paare]
    # je Y nur das stärkste X behalten
    beste = {}
    for e in sorted(erg, reverse=True):
        beste.setdefault(e[2], e)
    return sorted(beste.values(), reverse=True)


def schreibe_zusatz(pfad, treffer):
    alt = {}
    if os.path.exists(pfad):
        with open(pfad, encoding="utf-8-sig") as f:
            next(f, None)
            for z in f:
                t = z.rstrip("\r\n").split(";")
                if len(t) >= 5:
                    alt[(t[0], t[1])] = z.rstrip("\r\n")
    for _s, x, y, r, n_paar, n_x, _tx, _ty in treffer:
        alt[(x, y)] = "%s;%s;%d;%d;%s" % (x, y, n_paar, n_x, ("%.4f" % r).replace(".", ","))
    with open(pfad, "w", encoding="utf-8", newline="") as f:
        f.write(ZUSATZ_HEADER)
        for k in sorted(alt):
            f.write(alt[k] + "\n")
    return len(alt)


def main(argv=None):
    ap = argparse.ArgumentParser(description="Set->Bestandteil-Beziehungen aus Amicron-Exporten ableiten")
    ap.add_argument("dateien", nargs="+")
    ap.add_argument("--zusatz", default=imp.STANDARD_ZUSATZ)
    ap.add_argument("--min-anteil", type=float, default=MIN_ANTEIL)
    ap.add_argument("--min-stuetze", type=int, default=MIN_STUETZE)
    ap.add_argument("--slave-master", action="store_true",
                    help="auch 1:1-Beziehungen Slave -> Master uebernehmen (Master zaehlt in der Statistik)")
    ap.add_argument("--schreibe", action="store_true")
    arg = ap.parse_args(argv)
    komp = imp.lade_komponenten(imp.STANDARD_KOMPONENTEN, imp.STANDARD_FAKTOREN, arg.zusatz)
    rg = lade_rechnungen(arg.dateien)
    verd = verdaechtige(arg.dateien, komp)
    treffer = leite_ab(rg, verd, arg.min_stuetze, arg.min_anteil, komp)
    sicher = [t for t in treffer if ist_vielfachpackung(t[3])]
    slaves = []
    if arg.slave_master:
        je_x = collections.Counter(t[1] for t in treffer)
        slaves = [t for t in treffer if not ist_vielfachpackung(t[3]) and ist_slave_master(t, je_x)]
        sicher = sicher + slaves
    unsicher = [t for t in treffer if t not in sicher]
    print("%d Rechnungen, %d verdaechtige Artikel, %d Beziehungen abgeleitet (%d Vielfachpackungen, %d Slave->Master 1:1, %d zur Pruefung)\n" % (
        len(rg), len(verd), len(treffer), len(sicher) - len(slaves), len(slaves), len(unsicher)))

    def zeige(liste):
        for s, x, y, r, n_paar, n_x, tx, ty in liste:
            print("%7.0f Stk  %-16s -> %-14s x%-7s (%d von %d Rechnungen)  %s | %s" % (s, x, y, r, n_paar, n_x, tx[:30], ty[:30]))
    print("VIELFACHPACKUNGEN (werden uebernommen):")
    zeige([t for t in sicher if t not in slaves])
    if arg.slave_master:
        print("\nSLAVE -> MASTER 1:1 (Master zaehlt, Slave-Zeile faellt weg; werden uebernommen):")
        zeige(slaves)
    print("\nZUR PRUEFUNG (Faktor 1 / krumm - NICHT uebernommen; bei Bedarf von Hand in die Zusatz-Datei):")
    zeige(unsicher)
    if arg.schreibe:
        n = schreibe_zusatz(arg.zusatz, sicher)
        print("\nGeschrieben: %s (%d Beziehungen)" % (arg.zusatz, n))
    else:
        print("\nVorschau - nichts geschrieben (--schreibe zum Speichern).")
    return 0


if __name__ == "__main__":
    sys.exit(main())
