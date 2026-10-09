# -*- coding: utf-8 -*-
"""
artikel_import_amicron.py - Amicron-Export der Ausgangsrechnungen -> artikel_historie.csv
============================================================================
Traegt Rechnungspositionen aus dem Amicron-Export TAGESGENAU in artikel_amicron.csv ein (Datum;Rechnungsnummer;
Artikelnummer;Bezeichnung;Menge;Umsatz) - Quelle der Online-Auswertung "Top-Artikel" (siehe artikel_upload.py):
Vorjahresvergleiche, Saisonverlaeufe, Top-Artikel in STUECK.

WAS GEZAEHLT WIRD (Regel von Matthias, 2026-10-09):
  * Mehrfachpackung (Set = N x EIN Einzelartikel, z.B. BP-GF-6 = 6 x BP-GF): es zaehlt der EINZELARTIKEL mit seiner
    Menge aus der Stueckliste (BP-GF, 6 Stueck); der Umsatz der Packung wird dem Einzelartikel zugeschlagen.
  * Zusammengesetzter Artikel (Set aus verschiedenen Teilen, z.B. Paella-Schlauch Pr1509mm, Bunsenbrenner + Kartuschen):
    es zaehlt der Artikel selbst, seine Bestandteile fallen weg.
  * normale Artikel: wie auf der Rechnung. Es muss nichts mit Faktoren umgerechnet werden.

Eingabe: CSV-Export(e) der Ausgangsrechnungen aus Amicron (Semikolon, Windows-1252 oder UTF-8), erzeugt mit
amicron/Exportdefinition_Artikelverkaeufe.XML (oder _v2). Die Spalten werden am Namen erkannt:
  Artikelnummer   "Artikel Nr. intern" | "Artikelnummer" | "Artikel Nr." | "Art.-Nr."
  Bezeichnung     "Titel" | "Bezeichnung" | "Artikelbezeichnung"
  Menge           "Menge" | "Anzahl"
  Betrag          "Betrag" | "Gesamtpreis" | "G-Preis" | "Gesamt" | "Summe"          (optional)
  Einzelpreis     "Einzelpreis" | "E-Preis" - ohne Betrag-Spalte gilt Umsatz = Menge x Einzelpreis
  Rechnungsnr.    "Rechnungsnummer" | "Rechnung Nr." | "Beleg Nr." | "Belegnummer"   (optional)
  Datum           "Rechnungsdatum" | "Belegdatum" | "Datum" (TT.MM.JJJJ, JJJJ-MM-TT, JJJJMMTT) (optional)
Versandkosten-Positionen werden wie im Dashboard weggelassen (ist_versand). Belege mit Gutschrift-/
Auftragsbestaetigungs-Nummern (111..., 444..., siehe packliste.RNR_GESPERRT) werden uebersprungen.

STUECKLISTEN-KOMPONENTEN: Der Amicron-Export enthaelt zusaetzlich zu den Rechnungszeilen auch die Komponenten von
Set-Artikeln (z.B. neben "BP-GF-6" eine Zeile "BP-GF" Menge 6; oft mit Listenpreis -> Scheinumsatz). Sie stehen nicht
auf der Rechnung. Der Importer entfernt sie anhand einer GELERNTEN Tabelle artikel_komponenten.csv (Set -> Komponente
-> Verhaeltnis), die aus Rechnungen mit bekannter Wahrheit (artikel_verkaeufe.csv des Dashboards = Hauptzeilen laut
Rechnungs-PDF) gelernt wird:
    py artikel_import_amicron.py --lerne august.csv juli.csv         (lernen/ergaenzen, wiederholbar)
Ohne Tabelle warnt der Importer, dass die Zahlen Komponenten enthalten.

Zuordnung zum Monat:
  * hat die Datei eine Datumsspalte, wird je Zeile der Monat aus dem Datum bestimmt (ein Export fuer ein
    ganzes Jahr ist moeglich); --monat JJJJ-MM begrenzt dann auf diesen Monat.
  * ohne Datumsspalte muss --monat JJJJ-MM angegeben werden (eine Datei = ein Monat).

Ergebnis: artikel_amicron.csv. Bereits vorhandene Zeilen der importierten Rechnungen werden ERSETZT, alle anderen
bleiben - der Import ist beliebig wiederholbar (z.B. woechentlich ein neuer Export). Ohne Datumsspalte (--monat)
steht der 1. des Monats als Datum.

Aufruf:
    py artikel_import_amicron.py export_2025.csv --trocken               (Auswertung ansehen, nichts schreiben)
    py artikel_import_amicron.py export_2025.csv                         (schreiben)
    py artikel_import_amicron.py export_2025_03.csv --monat 2025-03      (Datei ohne Datumsspalte)
    py artikel_import_amicron.py datei.csv --zeige-spalten               (nur Spalten + erste Zeilen zeigen)
    py artikel_import_amicron.py --lerne august.csv [weitere.csv]        (Komponenten-Tabelle lernen)
    py artikel_import_amicron.py --selftest
Optionen: --ziel DATEI, --komponenten DATEI, --verkaeufe DATEI, --ohne-komponentenfilter
"""

import argparse
import csv
import io
import os
import re
import sys
from collections import Counter, defaultdict
from datetime import datetime

VERSION = "2026-10-09d"
PAKET = r"\\DESKTOP-N2H75H\Netzwerk\Paketscheine"
STANDARD_ZIEL = PAKET + r"\artikel_amicron.csv"
TAGE_HEADER = "Datum;Rechnungsnummer;Artikelnummer;Bezeichnung;Menge;Umsatz\n"
STANDARD_KOMPONENTEN = PAKET + r"\artikel_komponenten.csv"
STANDARD_VERKAEUFE = PAKET + r"\artikel_verkaeufe.csv"
STANDARD_FAKTOREN = PAKET + r"\artikel_stueckfaktor.csv"
HEADER = "Monat;Artikelnummer;Bezeichnung;Bestellungen;Menge;Umsatz\n"
KOMP_HEADER = "Set;Komponente;RechnungenKomponente;RechnungenSet;Verhaeltnis\n"
MIN_STUETZE = 2          # mindestens so viele Rechnungen, in denen Set UND Komponente zusammen vorkamen
MIN_ANTEIL = 0.6         # ... und in mindestens so vielen % der Rechnungen mit dem Set
MIN_REINE_KOMP = 5       # "reine Komponente": Artikel, der in der Lernbasis NIE als Hauptzeile, aber mind. so oft als Extra vorkam

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


def tag_aus_datum(s):
    """'TT.MM.JJJJ' / 'JJJJ-MM-TT' / 'JJJJMMTT' -> 'JJJJ-MM-TT' (None bei unlesbarem Datum)."""
    s = (s or "").strip()[:10]
    if re.fullmatch(r"\d{8}", s):
        s = s[:4] + "-" + s[4:6] + "-" + s[6:]
    for fmt in ("%d.%m.%Y", "%Y-%m-%d", "%d.%m.%y"):
        try:
            return datetime.strptime(s, fmt).strftime("%Y-%m-%d")
        except ValueError:
            continue
    return None


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


def beleg_gesperrt(rnr):
    """True fuer Gutschriften/Auftragsbestaetigungen (Nummernpraefix wie packliste.RNR_GESPERRT)."""
    try:
        import packliste
        praefixe = tuple(packliste.RNR_GESPERRT)
    except Exception:
        praefixe = ("111", "444")
    return bool(rnr) and rnr.startswith(praefixe)


def ist_versand(art, bez):
    try:
        import packliste
        return packliste.ist_versand(art, bez)
    except Exception:                                   # packliste/pdfplumber nicht installiert
        return "versandkosten" in (bez or "").lower() or not ((art or "").strip() or (bez or "").strip())


def nrm(a):
    return re.sub(r"[^a-z0-9]", "", (a or "").lower())


# ---------------------------------------------------------------------------------------------------------
# Einlesen
# ---------------------------------------------------------------------------------------------------------

def parse_export(text, monat=None):
    """-> (zeilen [dict rnr, mo, art, titel, menge, ep, betrag], info). Wirft ValueError bei fehlenden Pflichtspalten.
    Gutschriften/AB, Versandzeilen und Zeilen ausserhalb von monat werden schon hier aussortiert."""
    rows = list(csv.reader(io.StringIO(text), delimiter=";", quoting=csv.QUOTE_NONE))
    if not rows:
        raise ValueError("leere Datei")
    cols = finde_spalten(rows[0])
    fehlend = [k for k in ("art", "titel", "menge") if k not in cols]
    if fehlend:
        raise ValueError("Spalte(n) nicht gefunden: %s. Vorhandene Spalten: %s"
                         % (", ".join(fehlend), "; ".join(rows[0])))
    if "datum" not in cols and not monat:
        raise ValueError("Der Export hat keine Datumsspalte - bitte --monat JJJJ-MM angeben")
    out = []
    info = {"zeilen": 0, "versand": 0, "ohne_monat": 0, "anderer_monat": 0, "gesperrt": 0, "spalten": cols}
    for z in rows[1:]:
        extra = len(z) - len(rows[0])
        if extra > 0 and "titel" in cols:                  # Semikolon im Titel: Zusatzspalten dem Titel zuschlagen
            t = cols["titel"]
            z = z[:t] + [";".join(z[t:t + extra + 1])] + z[t + extra + 1:]
        if len(z) <= max(cols.values()):
            continue
        info["zeilen"] += 1
        art, titel = z[cols["art"]].strip(), z[cols["titel"]].strip()
        rnr = z[cols["rnr"]].strip() if "rnr" in cols else ""
        if beleg_gesperrt(rnr):
            info["gesperrt"] += 1
            continue
        if not art or art == "660" or ist_versand(art, titel):      # Versand-/Dienstleistungszeilen ("DHL Paket", 660)
            info["versand"] += 1
            continue
        tag = tag_aus_datum(z[cols["datum"]]) if "datum" in cols else (monat + "-01" if monat else None)
        mo = tag[:7] if tag else None
        if not mo:
            info["ohne_monat"] += 1
            continue
        if monat and mo != monat:
            info["anderer_monat"] += 1
            continue
        menge = zahl(z[cols["menge"]])
        ep = zahl(z[cols["ep"]]) if "ep" in cols else 0.0
        betrag = zahl(z[cols["betrag"]]) if "betrag" in cols else menge * ep
        out.append({"rnr": rnr, "mo": mo, "tag": tag, "art": art, "titel": titel, "menge": menge, "ep": ep, "betrag": betrag,
                    "rnr_da": "rnr" in cols})
    return out, info


# ---------------------------------------------------------------------------------------------------------
# Stuecklisten-Komponenten: lernen, speichern, anwenden
# ---------------------------------------------------------------------------------------------------------

def lies_wahrheit(verkaeufe_csv):
    """{rnr: Counter(artikelnr_normiert)} der Hauptzeilen aus artikel_verkaeufe.csv (Dashboard, laut Rechnungs-PDF)."""
    wahr = defaultdict(Counter)
    if not os.path.exists(verkaeufe_csv):
        return wahr
    with open(verkaeufe_csv, encoding="utf-8-sig") as f:
        next(f, None)
        for z in f:
            t = z.rstrip("\n").split(";")
            if len(t) >= 8 and t[1]:
                wahr[t[1]][nrm(t[2])] += 1
    return wahr


def teile_haupt_extra(zeilen_der_rechnung, wahr_counter):
    """Zeilen einer Rechnung -> (haupt, extra): Hauptzeilen = Artikel laut Dashboard-Wahrheit (je Zeile einmal),
    alles Uebrige = Extra (Komponenten)."""
    frei = Counter(wahr_counter)
    haupt, extra = [], []
    for l in zeilen_der_rechnung:
        k = nrm(l["art"])
        if frei[k] > 0:
            frei[k] -= 1
            haupt.append(l)
        else:
            extra.append(l)
    return haupt, extra


def lerne_komponenten(zeilen, wahrheit):
    """Zaehlt je (Set, Komponente) die Rechnungen, in denen beide vorkommen (Set = Hauptzeile, Komponente = Extra).
    Rueckgabe {set: {"n": Rechnungen mit Set, "komp": {x: [Rechnungen, Verhaeltnissumme]}}} - nur Rechnungen,
    die in der Wahrheit vorkommen."""
    je = defaultdict(list)
    for l in zeilen:
        if l["rnr"] in wahrheit:
            je[l["rnr"]].append(l)
    erg = {}
    rein = {"n": 0, "komp": {}}          # Sonderschluessel "*": je Artikel [Extra-Zeilen, Haupt-Zeilen]
    for rnr, ls in je.items():
        haupt, extra = teile_haupt_extra(ls, wahrheit[rnr])
        for l in haupt:
            rein["komp"].setdefault(nrm(l["art"]), [0, 0.0])[1] += 1
        for l in extra:
            rein["komp"].setdefault(nrm(l["art"]), [0, 0.0])[0] += 1
        ex = defaultdict(float)
        for l in extra:
            ex[nrm(l["art"])] += l["menge"]
        sets = defaultdict(float)
        for l in haupt:
            sets[nrm(l["art"])] += l["menge"]
        for s, ms in sets.items():
            e = erg.setdefault(s, {"n": 0, "komp": {}})
            e["n"] += 1
            for x, mq in ex.items():
                k = e["komp"].setdefault(x, [0, 0.0])
                k[0] += 1
                k[1] += (mq / ms) if ms else 0.0
    erg["*"] = rein
    return erg


def lies_tabelle(pfad):
    """{set: {"n": nSet, "komp": {x: [rechnungen, verhaeltnissumme]}}} aus artikel_komponenten.csv."""
    erg = {}
    if not pfad or not os.path.exists(pfad):
        return erg
    with open(pfad, encoding="utf-8-sig") as f:
        next(f, None)
        for z in f:
            t = z.rstrip("\r\n").split(";")
            if len(t) < 5 or not t[0]:
                continue
            try:
                n, rx, ns, vh = int(t[2]), int(t[2]), int(t[3]), float(t[4].replace(",", "."))
            except ValueError:
                continue
            e = erg.setdefault(t[0], {"n": ns, "komp": {}})
            if t[0] == "*":                          # reine Komponenten: [Extra-Zeilen, Haupt-Zeilen]
                e["komp"][t[1]] = [rx, float(ns)]
                continue
            e["n"] = max(e["n"], ns)
            e["komp"][t[1]] = [rx, vh * rx]
    return erg


def schreibe_tabelle(pfad, tab):
    ordner = os.path.dirname(pfad)
    if ordner:
        os.makedirs(ordner, exist_ok=True)
    with open(pfad, "w", encoding="utf-8", newline="") as f:
        f.write(KOMP_HEADER)
        for s in sorted(tab):
            for x, (rx, vs) in sorted(tab[s]["komp"].items()):
                if s == "*":
                    f.write("*;%s;%d;%d;0\n" % (x, rx, int(vs)))
                else:
                    f.write("%s;%s;%d;%d;%s\n" % (s, x, rx, tab[s]["n"], ("%.4f" % (vs / rx)).replace(".", ",")))


def faktor_komponenten(pfad):
    """Zusaetzliche Komponenten-Beziehungen aus der Stueckfaktor-Tabelle: Packung A (Faktor f > 1, Gruppe G) enthaelt
    f Stueck des Einzelartikels B derselben Gruppe (Faktor 1). {set: {komponente: f}} (normierte Artikelnummern)."""
    if not pfad or not os.path.exists(pfad):
        return {}
    zeilen = []
    with open(pfad, encoding="utf-8-sig") as fh:
        next(fh, None)
        for z in fh:
            t = z.rstrip("\r\n").split(";")
            if len(t) >= 4 and t[0].strip() and t[3].strip():
                try:
                    zeilen.append((t[0].strip(), float(t[2].strip().replace(",", ".")), t[3].strip().lower()))
                except ValueError:
                    pass
    basis = {}
    for art, f, g in zeilen:
        if f == 1 and g not in basis:
            basis[g] = art
    out = {}
    for art, f, g in zeilen:
        if f > 1 and g in basis and nrm(basis[g]) != nrm(art):
            out.setdefault(nrm(art), {})[nrm(basis[g])] = f
    return out


def vereinige(akt, zusatz):
    """Gelernte Beziehungen haben Vorrang; zusaetzliche (aus Faktoren) werden nur ergaenzt."""
    out = {s: dict(v) for s, v in akt.items()}
    for s, v in zusatz.items():
        for x, f in v.items():
            out.setdefault(s, {}).setdefault(x, f)
    return out


def merge_tabellen(alt, neu):
    """Addiert die Zaehler (Rechnungen) beider Tabellen."""
    out = {s: {"n": e["n"], "komp": {x: list(v) for x, v in e["komp"].items()}} for s, e in alt.items()}
    for s, e in neu.items():
        o = out.setdefault(s, {"n": 0, "komp": {}})
        o["n"] += e["n"]
        for x, (rx, vs) in e["komp"].items():
            k = o["komp"].setdefault(x, [0, 0.0])
            k[0] += rx
            k[1] += vs
    return out


def aktive_komponenten(tab):
    """{set: {komponente: verhaeltnis}} nach den Schwellen (MIN_STUETZE, MIN_ANTEIL)."""
    akt = {}
    for s, e in tab.items():
        for x, (rx, vs) in e["komp"].items():
            if s == "*":
                if vs == 0 and rx >= MIN_REINE_KOMP:     # nie Hauptzeile, oft Extra -> reine Komponente
                    akt.setdefault("*", {})[x] = 1.0
            elif rx >= MIN_STUETZE and e["n"] and rx / e["n"] >= MIN_ANTEIL:
                akt.setdefault(s, {})[x] = vs / rx
    return akt


def wende_stueckliste_an(zeilen, akt):
    """Wendet die Stuecklisten-Regel je Rechnung an (siehe Modul-Doku):
      * Set mit mind. 2 verschiedenen Bestandteilen (zusammengesetzter Artikel): Bestandteile entfernen, Set bleibt.
      * Set mit genau EINEM Bestandteil (Mehrfachpackung): Einzelartikel behalten, Set-Zeile entfernen, Umsatz des
        Sets dem Einzelartikel zuschlagen; fehlt die Bestandteil-Zeile, bleibt das Set unveraendert stehen.
      * "reine Komponenten" (nie Hauptzeile), die zu keinem erkannten Set gehoeren: entfernen.
    Rueckgabe (behalten, n_zeilen, n_stueck, n_umsatz, pack) - n_* = entfernte Bestandteile/reine Komponenten,
    pack = {"sets": Zahl aufgeloester Packungen, "stueck": Stueck der Einzelartikel, "umsatz": zugeschlagen}."""
    je = defaultdict(list)
    order = []
    for l in zeilen:
        key = l["rnr"] or id(l)
        if key not in je:
            order.append(key)
        je[key].append(l)
    behalten, n_zeilen, n_stueck, n_umsatz = [], 0, 0.0, 0.0
    pack = {"sets": 0, "stueck": 0.0, "umsatz": 0.0}
    rein = akt.get("*", {}) if akt else {}
    for key in order:
        ls = [dict(l) for l in je[key]]
        if not ls[0]["rnr"] or not akt:
            behalten += ls
            continue
        weg = set()                                     # Indizes entfernter Zeilen
        fest = set()                                    # Indizes bewusst behaltener Einzelartikel (Packungen)
        sets = defaultdict(list)
        for ix, l in enumerate(ls):
            k = nrm(l["art"])
            if k in akt and k != "*":
                sets[k].append(ix)
        # --- zusammengesetzte Artikel: Bestandteile entfernen ---
        erwartet = Counter()
        for s_, ixs in sets.items():
            if len(akt[s_]) >= 2:
                ms = sum(ls[ix]["menge"] for ix in ixs)
                for x, rt in akt[s_].items():
                    erwartet[x] += rt * ms
        for ix, l in enumerate(ls):
            k = nrm(l["art"])
            if erwartet[k] > 1e-9 and k not in sets:
                rest = l["menge"] - erwartet[k]
                erwartet[k] = max(0.0, erwartet[k] - l["menge"])
                gewicht = l["menge"] - max(rest, 0.0)
                n_zeilen += 1
                n_stueck += gewicht
                if l["menge"]:
                    n_umsatz += l["betrag"] * (gewicht / l["menge"])
                if rest > 1e-9:
                    l["betrag"] = l["betrag"] * (rest / l["menge"]) if l["menge"] else l["betrag"]
                    l["menge"] = rest
                else:
                    weg.add(ix)
        # --- Mehrfachpackungen: Einzelartikel zaehlt, Set-Zeile faellt weg ---
        for s_, ixs in sets.items():
            if len(akt[s_]) != 1:
                continue
            x, rt = next(iter(akt[s_].items()))
            q = sum(ls[ix]["menge"] for ix in ixs)
            b = sum(ls[ix]["betrag"] for ix in ixs)
            erw = rt * q
            xl = [ix for ix, l in enumerate(ls) if nrm(l["art"]) == x and ix not in weg and ix not in sets.get(x, [])]
            if not xl or sum(ls[ix]["menge"] for ix in xl) < 0.5 * erw:
                continue                                  # Bestandteil fehlt -> Set unveraendert lassen
            rest = erw
            for ix in xl:
                if rest <= 1e-9:
                    break
                l = ls[ix]
                t = min(l["menge"], rest)
                anteil = b * (t / erw) if erw else 0.0
                eigen = l["betrag"] * ((l["menge"] - t) / l["menge"]) if l["menge"] else 0.0
                l["betrag"] = anteil + eigen
                rest -= t
                l["fest"] = True
                fest.add(ix)
                pack["stueck"] += t
                pack["umsatz"] += anteil
            for ix in ixs:
                weg.add(ix)
            pack["sets"] += 1
        # --- reine Komponenten ohne erkanntes Set ---
        for ix, l in enumerate(ls):
            if ix not in weg and ix not in fest and nrm(l["art"]) in rein:
                weg.add(ix)
                n_zeilen += 1
                n_stueck += l["menge"]
                n_umsatz += l["betrag"]
        behalten += [l for ix, l in enumerate(ls) if ix not in weg]
    return behalten, n_zeilen, n_stueck, n_umsatz, pack


def entferne_komponenten(zeilen, akt):
    """Kompatibilitaet: wie wende_stueckliste_an, ohne das pack-Ergebnis."""
    return wende_stueckliste_an(zeilen, akt)[:4]


# ---------------------------------------------------------------------------------------------------------
# Aggregation / Ausgabe
# ---------------------------------------------------------------------------------------------------------

def aggregiere(zeilen):
    agg = {}
    for l in zeilen:
        a = agg.setdefault((l["mo"], l["art"].lower()), {"art": l["art"], "titel": Counter(), "rnr": set(),
                                                         "menge": 0.0, "umsatz": 0.0, "rnr_da": l["rnr_da"]})
        a["titel"][l["titel"]] += 1
        a["menge"] += l["menge"]
        a["umsatz"] += l["betrag"]
        if l["rnr"]:
            a["rnr"].add(l["rnr"])
    return agg


def lies_export(text, monat=None, komponenten=None):
    """-> (agg {(monat, art): {...}}, info dict). komponenten = aktive Tabelle (aktive_komponenten) oder None."""
    zeilen, info = parse_export(text, monat)
    info["komp_zeilen"] = info["komp_stueck"] = info["komp_umsatz"] = 0
    info["pack"] = {"sets": 0, "stueck": 0.0, "umsatz": 0.0}
    if komponenten:
        (zeilen, info["komp_zeilen"], info["komp_stueck"], info["komp_umsatz"],
         info["pack"]) = wende_stueckliste_an(zeilen, komponenten)
    info["zeilen_final"] = zeilen
    # Verdaechtige Reste: Nullpreis-Zeilen und "Master"-Artikel, die nicht als Komponente erkannt wurden
    verd = [l for l in zeilen if not l.get("fest") and (l["betrag"] == 0 or re.search(r"\bmaster\b", l["titel"], re.I))]
    info["verdaechtig"] = len(verd)
    info["verdaechtig_stueck"] = sum(l["menge"] for l in verd)
    return aggregiere(zeilen), info


def schreibe_tage(ziel, zeilen, trocken=False):
    """Schreibt die Positionen TAGESGENAU nach artikel_amicron.csv (je Rechnung+Artikel eine Zeile). Zeilen bereits
    vorhandener Rechnungen werden ersetzt. Rueckgabe (neue_rechnungen, geschriebene_zeilen)."""
    agg = {}
    for l in zeilen:
        if not l["rnr"]:
            continue
        a = agg.setdefault((l["rnr"], l["art"].lower()), {"tag": l["tag"], "art": l["art"], "titel": Counter(),
                                                         "menge": 0.0, "umsatz": 0.0})
        a["titel"][l["titel"]] += 1
        a["menge"] += l["menge"]
        a["umsatz"] += l["betrag"]
    rnrs = {k[0] for k in agg}
    behalten = []
    if os.path.exists(ziel):
        with open(ziel, encoding="utf-8-sig") as f:
            next(f, None)
            behalten = [z for z in f if z.strip() and z.split(";", 2)[1] not in rnrs]

    def zt(x):
        return str(int(x)) if x == int(x) else ("%.2f" % x).replace(".", ",")
    neu = []
    for (rnr, _), a in sorted(agg.items(), key=lambda kv: (kv[1]["tag"], kv[0])):
        titel = " ".join(a["titel"].most_common(1)[0][0].replace(";", ",").split())
        neu.append("%s;%s;%s;%s;%s;%s\n" % (a["tag"], rnr, a["art"].replace(";", ","), titel, zt(a["menge"]),
                                           ("%.2f" % a["umsatz"]).replace(".", ",")))
    if not trocken:
        ordner = os.path.dirname(ziel)
        if ordner:
            os.makedirs(ordner, exist_ok=True)
        alle = sorted(behalten + neu, key=lambda z: (z.split(";", 1)[0], z.split(";", 2)[1], z.split(";", 3)[2].lower()))
        with open(ziel, "w", encoding="utf-8", newline="") as f:
            f.write(TAGE_HEADER)
            f.writelines(z if z.endswith("\n") else z + "\n" for z in alle)
    return len(rnrs), len(neu)


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
    je = defaultdict(lambda: [0, 0.0, 0.0, set()])
    for (mo, _), a in agg.items():
        je[mo][0] += 1
        je[mo][1] += a["menge"]
        je[mo][2] += a["umsatz"]
        je[mo][3] |= a["rnr"]
    zeilen = []
    for mo in sorted(je):
        n, m, u, rg = je[mo]
        zeilen.append("  %s: %d Rechnungen, %d Artikel, %.0f Stück, Umsatz %.2f EUR%s" % (
            mo, len(rg), n, m, u, (" (%.2f EUR je Rechnung)" % (u / len(rg))) if rg else ""))
        top = sorted(((a["menge"], a["art"], a["titel"].most_common(1)[0][0]) for (k, _), a in agg.items() if k == mo),
                     reverse=True)[:3]
        zeilen += ["      %s  %-16s %s" % (("%.0f" % m_).rjust(6), a_[:16], t_[:44]) for m_, a_, t_ in top]
    return "\n".join(zeilen)


# ---------------------------------------------------------------------------------------------------------
# Selbsttest
# ---------------------------------------------------------------------------------------------------------

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
    g, gi = lies_export("Rechnungsnummer;Rechnungsdatum;Artikelnummer;Titel;Menge;Einzelpreis\n"
                        "1118209;20250310;G1;Gutschrift;1;5,00\n4440035;20250310;G2;AB;1;5,00\n"
                        "1700001;20250310;N1;Normal;1;5,00\n")
    check("Gutschrift (111...) und AB (444...) uebersprungen", (sorted(k[1] for k in g), gi["gesperrt"]), (["n1"], 2))
    agg2, _ = lies_export(amicron, monat="2025-03")
    check("--monat begrenzt", sorted({k[0] for k in agg2}), ["2025-03"])
    zub = "Artikelnummer;Titel;Menge;Betrag\nA1;Artikel Eins;4;0,00\n"
    try:
        lies_export(zub)
        check("ohne Datum und --monat -> Fehler", True, False)
    except ValueError:
        check("ohne Datum und --monat -> Fehler", True, True)
    agg3, _ = lies_export(zub, monat="2025-05")
    check("Zubehoer-Format mit --monat, ohne Rechnungsnummer", (agg3[("2025-05", "a1")]["menge"], agg3[("2025-05", "a1")]["rnr_da"]), (4.0, False))
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

    # --- Komponenten lernen und anwenden ------------------------------------------------------------------
    kopf = "Rechnungsnummer;Rechnungsdatum;Artikelnummer;Titel;Menge;Einzelpreis\n"
    lern = kopf
    wahr_zeilen = ["Datum;Rechnungsnummer;Artikelnummer;Bezeichnung;Menge;Einheit;Einzelpreis;Gesamtpreis"]
    for i in range(1, 6):                          # 5 Rechnungen: Set S1 (14,99) + Komponente K1 (6 Stk, 0,00) + K2 (1 Stk, 3,00)
        lern += "50%02d;20260805;S1;Set eins;1;14,99\n50%02d;20260805;K1;Komp eins;6;0,00\n50%02d;20260805;K2;Komp zwei;1;3,00\n" % (i, i, i)
        wahr_zeilen.append("2026-08-05;50%02d;S1;Set eins;1;Stck.;14,99;14,99" % i)
    lern += "5100;20260805;K1;Komp eins als Einzelartikel;2;4,00\n"      # K1 auch einzeln verkauft
    wahr_zeilen.append("2026-08-05;5100;K1;Komp eins als Einzelartikel;2;Stck.;4,00;8,00")
    tmp = tempfile.mkdtemp(prefix="komp_test_")
    try:
        vk = os.path.join(tmp, "artikel_verkaeufe.csv")
        open(vk, "w", encoding="utf-8").write("\n".join(wahr_zeilen) + "\n")
        zl, _ = parse_export(lern)
        tab = lerne_komponenten(zl, lies_wahrheit(vk))
        akt = aktive_komponenten(tab)
        check("gelernt: S1 hat Komponenten K1 (6x) und K2 (1x)", {x: round(v, 2) for x, v in akt.get("s1", {}).items()}, {"k1": 6.0, "k2": 1.0})
        check("K1 (Einzelartikel) ist selbst kein Set", "k1" in akt, False)
        tp = os.path.join(tmp, "artikel_komponenten.csv")
        schreibe_tabelle(tp, tab)
        akt2 = aktive_komponenten(lies_tabelle(tp))
        check("Tabelle speichern/lesen", {x: round(v, 2) for x, v in akt2["s1"].items()}, {"k1": 6.0, "k2": 1.0})
        # --- Regel Mehrfachpackung / zusammengesetzter Artikel (Matthias 2026-10-09) ---
        pk = {"bpgf6": {"bpgf": 6.0}, "pr1509mm": {"1419": 1.0, "1695": 1.5}, "*": {"bpgf": 1.0}}
        kp = kopf
        pa, pi = lies_export(kp + "1;20260805;BP-GF-6;GermanFire 6 Stück;1;14,99\n1;20260805;BP-GF;Brennpaste Dose;6;0,00\n", komponenten=pk)
        check("Packung: BP-GF-6 faellt weg, BP-GF zaehlt 6 Stueck mit dem Umsatz der Packung (auch wenn BP-GF 'reine Komponente' ist)",
              sorted((k[1], a_["menge"], round(a_["umsatz"], 2)) for k, a_ in pa.items()), [("bp-gf", 6.0, 14.99)])
        check("Packungsinfo", (pi["pack"]["sets"], pi["pack"]["stueck"], pi["pack"]["umsatz"]), (1, 6.0, 14.99))
        pb, _ = lies_export(kp + "2;20260805;BP-GF-6;GermanFire 6 Stück;1;14,99\n2;20260805;BP-GF;Brennpaste Dose;6;0,00\n"
                            "2;20260805;BP-GF;Brennpaste Dose einzeln;2;5,00\n", komponenten={"bpgf6": {"bpgf": 6.0}})
        check("Packung + Einzelverkauf desselben Artikels in einer Rechnung: 6 (Packung) + 2 (einzeln)",
              (round(sum(a_["menge"] for a_ in pb.values()), 2), round(sum(a_["umsatz"] for a_ in pb.values()), 2)), (8.0, 24.99))
        pc, _ = lies_export(kp + "3;20260805;BP-GF-6;GermanFire 6 Stück;1;14,99\n", komponenten={"bpgf6": {"bpgf": 6.0}})
        check("Packung ohne Bestandteil-Zeile bleibt unveraendert", sorted(k[1] for k in pc), ["bp-gf-6"])
        pd_, _ = lies_export(kp + "4;20260805;Pr1509mm;Paella-Schlauch;1;12,50\n4;20260805;1419;Ueberwurfmutter;1;0,00\n"
                             "4;20260805;1695;Schlauch Meterware;1,5;0,00\n", komponenten=pk)
        check("zusammengesetzter Artikel: Paella-Schlauch bleibt, Bestandteile fallen weg", sorted(k[1] for k in pd_), ["pr1509mm"])
        # Tagesausgabe
        zl4, _ = parse_export(kp + "10;20260805;A;Artikel A;2;3,00\n10;20260805;A;Artikel A;1;3,00\n11;20260806;B;Artikel B;1;7,50\n")
        tg = os.path.join(tmp, "t.csv")
        check("schreibe_tage: 2 Rechnungen, A je Rechnung zusammengefasst", schreibe_tage(tg, zl4), (2, 2))
        t1 = open(tg, encoding="utf-8").read().strip().split("\n")
        check("Tageszeile", t1[1], "2026-08-05;10;A;Artikel A;3;9,00")
        schreibe_tage(tg, parse_export(kp + "10;20260805;A;Artikel A;5;3,00\n12;20260807;C;Artikel C;1;1,00\n")[0])
        t2 = open(tg, encoding="utf-8").read().strip().split("\n")
        check("wiederholbar: Rechnung 10 ersetzt, 11 bleibt, 12 neu", [z.split(";")[1:5] for z in t2[1:]],
              [["10", "A", "Artikel A", "5"], ["11", "B", "Artikel B", "1"], ["12", "C", "Artikel C", "1"]])
        # Beziehungen aus der Stueckfaktor-Tabelle (BP-GF-6 = 6 x BP-GF)
        ft = os.path.join(tmp, "f.csv")
        open(ft, "w", encoding="utf-8").write("Artikelnummer;Bezeichnung;Faktor;Gruppe;Quelle\nBP-GF;Dose;1;BP-GF;auto\n"
                                              "BP-GF-6;6er;6;BP-GF;auto\nBP-GF-24;24er;24;BP-GF;auto\nOHNE;x;4;;auto\n")
        fz = faktor_komponenten(ft)
        check("Stueckfaktor-Tabelle -> Komponenten", {k: dict(v) for k, v in fz.items()}, {"bpgf6": {"bpgf": 6.0}, "bpgf24": {"bpgf": 24.0}})
        vz = vereinige({"bpgf6": {"bpgf": 5.0}}, fz)
        check("gelernte Beziehung hat Vorrang, fehlende kommen dazu", (vz["bpgf6"]["bpgf"], vz["bpgf24"]["bpgf"]), (5.0, 24.0))
        # reine Komponente: K9 nie Hauptzeile, 5x als Extra -> immer entfernt
        lern2 = kopf
        wz2 = ["Datum;Rechnungsnummer;Artikelnummer;Bezeichnung;Menge;Einheit;Einzelpreis;Gesamtpreis"]
        for i in range(1, 7):
            lern2 += "80%02d;20260805;S2;Set zwei;1;9,99\n80%02d;20260805;K9;Komp neun;3;1,00\n" % (i, i)
            wz2.append("2026-08-05;80%02d;S2;Set zwei;1;Stck.;9,99;9,99" % i)
        vk2 = os.path.join(tmp, "v2.csv")
        open(vk2, "w", encoding="utf-8").write("\n".join(wz2) + "\n")
        akt3 = aktive_komponenten(lerne_komponenten(parse_export(lern2)[0], lies_wahrheit(vk2)))
        check("reine Komponente gelernt (K9), S2 hat ebenfalls Beziehung", (sorted(akt3.get("*", {})), "k9" in akt3.get("s2", {})), (["k9"], True))
        weg, wi = lies_export(kopf + "9001;20260902;K9;Komp neun einzeln;4;1,00\n9001;20260902;N1;Normal;1;5,00\n", komponenten=akt3)
        check("reine Komponente auch OHNE Set entfernt", (sorted(k[1] for k in weg), wi["komp_zeilen"]), (["n1"], 1))
        schreibe_tabelle(tp, merge_tabellen(tab, lerne_komponenten(parse_export(lern2)[0], lies_wahrheit(vk2))))
        check("Tabelle mit '*'-Zeilen speichern/lesen (K2 aus dem ersten Lauf + K9)", sorted(aktive_komponenten(lies_tabelle(tp)).get("*", {})), ["k2", "k9"])
        check("Dienstleistungszeile ohne Artikelnummer und 660 werden weggelassen",
              len(lies_export(kopf + "9100;20260902;;DHL Paket;1;3,90\n9100;20260902;660;Versand- und Verpackungskosten;1;3,50\n9100;20260902;N1;Normal;1;5,00\n")[0]), 1)
        neu_agg, ni = lies_export(kopf + "6001;20260901;S1;Set eins;2;14,99\n6001;20260901;K1;Komp eins;12;0,00\n"
                                  "6001;20260901;K2;Komp zwei;2;3,00\n6001;20260901;K1;Komp eins als Einzelartikel;3;4,00\n"
                                  "6002;20260902;Z9;Unbekannt;1;9,00\n", komponenten=akt2)
        check("Komponenten entfernt: nur S1, K1 (Einzelartikel 3 Stk), Z9 bleiben",
              sorted((k[1], a["menge"]) for k, a in neu_agg.items()), [("k1", 3.0), ("s1", 2.0), ("z9", 1.0)])
        check("entfernte Komponenten gezaehlt", (ni["komp_zeilen"], ni["komp_stueck"], round(ni["komp_umsatz"], 2)), (2, 14.0, 6.0))
        check("Umsatz der verbleibenden K1-Zeile = Einzelartikel", round(neu_agg[("2026-09", "k1")]["umsatz"], 2), 12.0)
        check("ohne Tabelle: nichts entfernt", len(lies_export(kopf + "7001;20260901;S1;Set eins;1;14,99\n7001;20260901;K1;Komp;6;0,00\n")[0]), 2)
        # Zaehler mehrerer Laeufe addieren
        gem = merge_tabellen(tab, tab)
        check("merge: Rechnungszahlen addiert", (gem["s1"]["n"], gem["s1"]["komp"]["k1"][0]), (tab["s1"]["n"] * 2, tab["s1"]["komp"]["k1"][0] * 2))
        # Historie schreiben/ersetzen
        ziel = os.path.join(tmp, "h.csv")
        schreibe_historie(ziel, agg)
        schreibe_historie(ziel, agg3)
        z1 = open(ziel, encoding="utf-8").read().strip().split("\n")
        check("Historie: Kopf + 4 Zeilen", len(z1), 5)
        check("Zeile mit Bestellungen", z1[1], "2025-03;A1;Artikel Eins;2;3;15,00")
        check("ohne Rechnungsnummer: Bestellungen leer", z1[4], "2025-05;A1;Artikel Eins;;4;0,00")
        schreibe_historie(ziel, agg2)
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


# ---------------------------------------------------------------------------------------------------------
# Programmstart
# ---------------------------------------------------------------------------------------------------------

def main(argv=None):
    ap = argparse.ArgumentParser(description="Amicron-Export -> artikel_historie.csv")
    ap.add_argument("dateien", nargs="*")
    ap.add_argument("--monat", default=None, help="JJJJ-MM")
    ap.add_argument("--ziel", default=STANDARD_ZIEL, help="artikel_amicron.csv (tagesgenau)")
    ap.add_argument("--komponenten", default=STANDARD_KOMPONENTEN)
    ap.add_argument("--verkaeufe", default=STANDARD_VERKAEUFE)
    ap.add_argument("--faktoren", default=STANDARD_FAKTOREN)
    ap.add_argument("--ohne-komponentenfilter", action="store_true")
    ap.add_argument("--trocken", action="store_true")
    ap.add_argument("--zeige-spalten", action="store_true")
    ap.add_argument("--lerne", action="store_true", help="Komponenten-Tabelle aus den Dateien lernen")
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

    if arg.lerne:
        wahr = lies_wahrheit(arg.verkaeufe)
        if not wahr:
            print("Keine Dashboard-Rechnungen in %s gefunden - ohne Wahrheit kann nichts gelernt werden." % arg.verkaeufe)
            return 1
        gesamt = {}
        for pfad in arg.dateien:
            text, enc = lies_text(pfad)
            try:
                zl, info = parse_export(text, arg.monat)
            except ValueError as e:
                print("%s: %s" % (pfad, e))
                return 1
            gemeinsam = {l["rnr"] for l in zl if l["rnr"] in wahr}
            print("%s: %d Zeilen, %d Rechnungen mit Dashboard-Wahrheit" % (pfad, len(zl), len(gemeinsam)))
            gesamt = merge_tabellen(gesamt, lerne_komponenten(zl, wahr))
        alt = lies_tabelle(arg.komponenten)
        neu = merge_tabellen(alt, gesamt)
        akt = aktive_komponenten(neu)
        print("\nSets mit Komponenten: %d (vorher %d) | Beziehungen: %d (Schwelle: >= %d Rechnungen und >= %d %% der Set-Rechnungen)" % (
            len(akt), len(aktive_komponenten(alt)), sum(len(v) for v in akt.values()), MIN_STUETZE, int(MIN_ANTEIL * 100)))
        if not arg.trocken:
            schreibe_tabelle(arg.komponenten, neu)
            print("Gespeichert: %s" % arg.komponenten)
        else:
            print("Trockenlauf - nichts gespeichert.")
        return 0

    komp = None
    if not arg.ohne_komponentenfilter and not arg.zeige_spalten:
        tab = lies_tabelle(arg.komponenten)
        komp = aktive_komponenten(tab)
        zus = faktor_komponenten(arg.faktoren)
        if zus:
            komp = vereinige(komp, zus)
            print("Zusaetzlich %d Packungsartikel aus der Stueckfaktor-Tabelle (%s)" % (len(zus), arg.faktoren))
        if komp:
            print("Komponenten-Tabelle: %d Sets, %d Beziehungen (%s)" % (len(komp), sum(len(v) for v in komp.values()), arg.komponenten))
        else:
            print("WARNUNG: keine Komponenten-Tabelle (%s) - der Export enthaelt Stuecklisten-Komponenten als eigene Zeilen "
                  "(oft mit Listenpreis); Stueck und Umsatz sind dann zu hoch. Erst lernen: --lerne august.csv ..." % arg.komponenten)

    gesamt = {}
    alle_zeilen = []
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
            agg, info = lies_export(text, arg.monat, komp)
        except ValueError as e:
            print("%s: %s" % (pfad, e))
            return 1
        print("%s (%s): %d Zeilen, %d Versandpositionen weggelassen%s%s%s" % (
            pfad, enc, info["zeilen"], info["versand"],
            (", %d Gutschrift/AB-Zeilen uebersprungen" % info["gesperrt"]) if info["gesperrt"] else "",
            (", %d ohne lesbares Datum" % info["ohne_monat"]) if info["ohne_monat"] else "",
            (", %d aus anderem Monat" % info["anderer_monat"]) if info["anderer_monat"] else ""))
        if info["komp_zeilen"]:
            print("    Komponentenzeilen entfernt: %d Zeilen, %.0f Stück, %.0f EUR Scheinumsatz" % (
                info["komp_zeilen"], info["komp_stueck"], info["komp_umsatz"]))
        if info["pack"]["sets"]:
            print("    Mehrfachpackungen aufgeloest: %d Packungen -> %.0f Stück Einzelartikel (%.0f EUR Umsatz zugeschlagen)" % (
                info["pack"]["sets"], info["pack"]["stueck"], info["pack"]["umsatz"]))
        if info["verdaechtig"]:
            print("    Verbleibend verdächtig (Preis 0 oder 'Master' im Titel, nicht als Komponente gelernt): "
                  "%d Zeilen, %.0f Stück" % (info["verdaechtig"], info["verdaechtig_stueck"]))
        print("    Spalten: %s" % ", ".join("%s=%d" % kv for kv in sorted(info["spalten"].items())))
        alle_zeilen.extend(info["zeilen_final"])
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
    n_rg, n_z = schreibe_tage(arg.ziel, alle_zeilen, arg.trocken)
    print("\n%s %d Rechnungen, %d Zeilen -> %s" % ("Trockenlauf:" if arg.trocken else "Geschrieben:", n_rg, n_z, arg.ziel))
    return 0


if __name__ == "__main__":
    sys.exit(main())
