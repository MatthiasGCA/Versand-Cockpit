# -*- coding: utf-8 -*-
"""
artikel_stueckfaktor.py - Stueckfaktor-Tabelle pflegen (Verkaufseinheit VE -> Stueck)
============================================================================
Die Online-Auswertung "Top-Artikel" zeigt STUECK statt Verkaufseinheiten: Eine Packung "GermanFire 24 Stueck"
(BP-GF-24) zaehlt 24 Stueck, "4x Chlorgranulat" 4 Stueck usw. Die Faktoren stehen in
artikel_stueckfaktor.csv (Paketscheine):  Artikelnummer;Bezeichnung;Faktor;Gruppe;Quelle
  * Faktor  Stueck je Verkaufseinheit (Standard 1, auch ohne Eintrag).
  * Gruppe  optional - gleiche Gruppe = gleiches Produkt in verschiedenen Packungsgroessen (z.B. BP-GF fuer
            BP-GF-6/-12/-24/-48); die Auswertung kann Gruppen zu EINER Zeile zusammenfassen.
  * Quelle  "auto" = aus den Rechnungen uebernommen, "hand" oder leer = von Hand gepflegt.
Von Hand eingetragene/geaenderte Zeilen haben IMMER Vorrang und werden von den Programmen nie ueberschrieben.
Das Carrier-Dashboard traegt neue Packungsartikel bei jedem Schritt 2 selbst ein ("N-Fach-Artikel"-Marker).

Aufruf:
    py artikel_stueckfaktor.py --aus-archiv            Faktoren aus den archivierten Rechnungen lesen (einmalig/Nachtrag)
    py artikel_stueckfaktor.py --ergaenze-basis        Einzelartikel (Faktor 1) zu erkannten Gruppen ergaenzen, wenn sie
                                                       in den Verkaeufen/der Historie vorkommen (z.B. BP-GF zu BP-GF-24)
    py artikel_stueckfaktor.py --pruefliste            Artikel zeigen, die wie Mehrfachpackungen aussehen, aber keinen
                                                       Faktor haben (schreibt artikel_stueckfaktor_vorschlaege.csv)
    py artikel_stueckfaktor.py --zeige                 Tabelle anzeigen
Optionen: --archiv ORDNER, --tabelle DATEI, --verkaeufe DATEI, --historie DATEI
"""

import argparse
import os
import re
import sys
from collections import defaultdict

VERSION = "2026-10-09a"
PAKETSCHEINE = r"\\DESKTOP-N2H75H\Netzwerk\Paketscheine"

# Hinweise im Titel, dass ein Artikel eine Mehrfachpackung sein koennte (nur fuer die Pruefliste!)
_RE_VORSCHLAG = [
    (re.compile(r"^\s*(\d+)\s*x\s", re.I), "Nx am Anfang"),
    (re.compile(r"\b(\d+)\s*x\s*\d"), "N x Menge"),
    (re.compile(r"\b(\d+)\s*(?:Stück|Stueck|Stck\.?|Stk\.?|St\.)(?!\w)", re.I), "N Stück"),
    (re.compile(r"\b(\d+)\s*(?:Rollen|Packungen|Dosen|Paar|Beutel|Flaschen|Kartuschen|Päckchen)\b", re.I), "N Rollen/Packungen/..."),
    (re.compile(r"\b(\d+)er[- ]?(?:Pack|Set|Karton)\b", re.I), "Ner-Pack"),
]


def tabelle_pfad(arg):
    return arg or os.path.join(PAKETSCHEINE, "artikel_stueckfaktor.csv")


def aus_archiv(archiv, tabelle):
    import carrier_statistik as cs
    import packliste
    pdfs = []
    for dp, dn, fn in os.walk(archiv):
        pdfs += [os.path.join(dp, f) for f in fn if f.lower().endswith(".pdf")]
    print("Archiv: %s (%d PDFs) -> %s" % (archiv, len(pdfs), tabelle))
    gesehen, neu_gesamt, block = {}, 0, []
    for i, pfad in enumerate(sorted(pdfs), 1):
        if i % 500 == 0:
            print("  ... %d/%d" % (i, len(pdfs)))
        try:
            r = packliste.parse_pdf(pfad)
        except Exception:
            continue
        for p in r["positionen"]:
            art = (p.get("art") or "").strip()
            if art and not packliste.ist_versand(p.get("art"), p.get("bez")):
                gesehen.setdefault(art.lower(), (art, " ".join((p.get("bez") or "").replace(";", ",").split())))
        block.append(r)
        if len(block) >= 300:
            neu_gesamt += cs.merke_stueckfaktoren(block, tabelle)
            block = []
    if block:
        neu_gesamt += cs.merke_stueckfaktoren(block, tabelle)
    gesehen.update(_bekannte_artikel(cs.ARTIKEL_VERKAEUFE_DATEI, os.path.join(os.path.dirname(tabelle), "artikel_historie.csv")))
    n_basis = ergaenze_basis(tabelle, gesehen)
    print("Neu eingetragen: %d Packungsartikel + %d Basisartikel" % (neu_gesamt, n_basis))
    return 0


def _bekannte_artikel(verkaeufe, historie):
    import carrier_statistik as cs
    gesehen = {}
    for datum, rnr, art, bez, m, einh, ep, gp in cs._lies_verkaeufe(verkaeufe):
        gesehen[art.lower()] = (art, " ".join(bez.replace(";", ",").split()))
    if historie and os.path.exists(historie):
        with open(historie, encoding="utf-8-sig") as f:
            next(f, None)
            for z in f:
                t = z.rstrip("\n").split(";")
                if len(t) >= 6:
                    gesehen.setdefault(t[1].lower(), (t[1], " ".join(t[2].split())))
    return gesehen


def ergaenze_basis(tabelle, gesehen):
    """Fuer jede Gruppe der Tabelle den Einzelartikel (Nummer = Gruppe, Gruppe-1 oder Gruppe/1) mit Faktor 1
    eintragen, wenn er in gesehen vorkommt und noch keine Zeile hat. Rueckgabe Anzahl."""
    import carrier_statistik as cs
    tab = cs.lies_stueckfaktoren(tabelle)
    gruppen = {}
    for v in tab.values():
        if v["gruppe"]:
            gruppen.setdefault(v["gruppe"].lower(), v["gruppe"])
    neu = []
    for g, gname in sorted(gruppen.items()):
        for kand in (g, g + "-1", g + "/1"):
            if kand in gesehen and kand not in tab:
                neu.append((gesehen[kand][0], gesehen[kand][1], gname))
                break
    if neu:
        anlegen = not os.path.exists(tabelle)
        with open(tabelle, "a", encoding="utf-8-sig" if anlegen else "utf-8", newline="") as f:
            if anlegen:
                f.write(cs._FAKTOR_HEADER)
            for art, bez, g in neu:
                f.write("%s;%s;1;%s;auto\n" % (art, bez, g))
    return len(neu)


def pruefliste(tabelle, verkaeufe, historie):
    import carrier_statistik as cs
    tab = cs.lies_stueckfaktoren(tabelle)
    menge = defaultdict(float)
    titel = {}
    for datum, rnr, art, bez, m, einh, ep, gp in cs._lies_verkaeufe(verkaeufe):
        menge[art.lower()] += m or 0
        titel[art.lower()] = (art, bez)
    if historie and os.path.exists(historie):
        with open(historie, encoding="utf-8-sig") as f:
            next(f, None)
            for z in f:
                t = z.rstrip("\n").split(";")
                if len(t) >= 6:
                    try:
                        menge[t[1].lower()] += float(t[4].replace(",", "."))
                    except ValueError:
                        pass
                    titel.setdefault(t[1].lower(), (t[1], t[2]))
    kand = []
    for k, (art, bez) in titel.items():
        if k in tab:
            continue
        for rx, was in _RE_VORSCHLAG:
            m = rx.search(bez)
            if m and int(m.group(1)) > 1:
                kand.append((menge[k], art, bez, int(m.group(1)), was))
                break
    kand.sort(reverse=True)
    out = os.path.join(os.path.dirname(tabelle) or ".", "artikel_stueckfaktor_vorschlaege.csv")
    with open(out, "w", encoding="utf-8-sig", newline="") as f:
        f.write("Artikelnummer;Bezeichnung;Faktor;Gruppe;Quelle\r\n")
        for m, art, bez, fk, was in kand:
            f.write("%s;%s;%d;%s;vorschlag (%s)\r\n" % (art, bez.replace(";", ","), fk, cs.gruppe_vorschlag(art, fk), was))
    print("%d Artikel ohne Faktor, die wie Mehrfachpackungen aussehen -> %s" % (len(kand), out))
    print("Pruefen und passende Zeilen in artikel_stueckfaktor.csv kopieren (Quelle auf 'hand' setzen).\n")
    print("%8s  %-18s %-6s %s" % ("Menge", "Artikel", "Faktor", "Bezeichnung (Hinweis)"))
    for m, art, bez, fk, was in kand[:40]:
        print("%8.0f  %-18s %-6d %s (%s)" % (m, art[:18], fk, bez[:50], was))
    return 0


def main(argv=None):
    ap = argparse.ArgumentParser(description="Stueckfaktor-Tabelle pflegen")
    ap.add_argument("--aus-archiv", action="store_true")
    ap.add_argument("--pruefliste", action="store_true")
    ap.add_argument("--ergaenze-basis", action="store_true")
    ap.add_argument("--zeige", action="store_true")
    ap.add_argument("--archiv", default=None)
    ap.add_argument("--tabelle", default=None)
    ap.add_argument("--verkaeufe", default=None)
    ap.add_argument("--historie", default=None)
    arg = ap.parse_args(argv)
    import carrier_statistik as cs
    tabelle = tabelle_pfad(arg.tabelle)
    if arg.aus_archiv:
        import carrier_dashboard as dash
        return aus_archiv(arg.archiv or dash.ARCHIV_ORDNER, tabelle)
    if arg.ergaenze_basis:
        gesehen = _bekannte_artikel(arg.verkaeufe or cs.ARTIKEL_VERKAEUFE_DATEI,
                                    arg.historie or os.path.join(os.path.dirname(tabelle), "artikel_historie.csv"))
        print("Basisartikel ergaenzt: %d" % ergaenze_basis(tabelle, gesehen))
        return 0
    if arg.pruefliste:
        return pruefliste(tabelle, arg.verkaeufe or cs.ARTIKEL_VERKAEUFE_DATEI,
                          arg.historie or os.path.join(os.path.dirname(tabelle), "artikel_historie.csv"))
    if arg.zeige:
        tab = cs.lies_stueckfaktoren(tabelle)
        print("%d Eintraege in %s" % (len(tab), tabelle))
        for v in sorted(tab.values(), key=lambda v: (v["gruppe"].lower(), v["art"].lower())):
            print("  %-18s x%-6g %-16s %-6s %s" % (v["art"][:18], v["faktor"], v["gruppe"][:16], v["quelle"], v["bez"][:40]))
        return 0
    ap.print_help()
    return 1


if __name__ == "__main__":
    sys.exit(main())
