# -*- coding: utf-8 -*-
"""
carrier_regeln.py - Regeln fuer die Carrier-Zuordnung (DHL / DPD / Deutsche Post)
==================================================================================
Reine Logik ohne Abhaengigkeiten (kein PDF, keine GUI) - damit sie sich
einzeln testen laesst:   py carrier_regeln.py     (fuehrt den Selbsttest aus)

Genutzt vom Carrier-Dashboard (carrier_dashboard.py). Eingabe ist das Ergebnis
von packliste.parse_pdf() (Felder: rnr, positionen[*].kennungen/lagerorte,
sendungsgewicht, adresse_zeilen, zeilen_ok, summe_ok).

REGELN (Stand 2026-09-21, mit Matthias abgestimmt)
--------------------------------------------------
Kennung aus dem Artikelstamm (Gross-/Kleinschreibung egal):
    Pax1 = DHL             (immer, egal wie leicht)
    Wapo = DPD             (nur Deutschland)
    Pox1 = Deutsche Post Grossbrief
    Brx1 = Deutsche Post Brief
Das Sendungsgewicht (kg) kann eine Sendung nur in die NAECHSTE Versandklasse
"aufsteigen" lassen, nie absteigen:
    unter 0,05 kg  Brief  ->  unter 0,6 kg  Grossbrief  ->  unter 1,1 kg  DPD
    ->  ab 1,1 kg  DHL
Beispiel: Pox1 (Grossbrief) mit 0,7 kg (Kunde bestellt 2x) wird DPD; Brx1 mit
0,3 kg wird Grossbrief; Wapo mit 1,5 kg wird DHL.
Mehrere Kennungen auf einer Rechnung: die HOECHSTE Klasse gewinnt
(Pax1 > Wapo > Pox1 > Brx1).
DPD faehrt nur innerhalb Deutschlands - ein Auslands-Paket, das nach den Regeln
DPD waere, geht per DHL. Deutsche Post Brief/Grossbrief gibt es fuer Inland UND
Ausland (getrennte CSV-Dateien, siehe carrier_dashboard.py).
Grenzwert selbst (z.B. genau 1,1 kg) gehoert zur HOEHEREN Klasse ("unter" gilt
strikt). Alle Grenzen stehen unten als Konstanten.

DHL-Maximalgewicht G_DHL_MAX (31,5 kg): darueber blockiert ein Fehler den
Export. "<N>-je-Paket"-markierte Artikel (packliste.JE_PAKET_RE, 2026-09-23)
werden davon automatisch ausgenommen: je_paket_aufteilung() teilt eine
Rechnung mit GENAU EINEM so markierten Artikel automatisch in mehrere
DHL-Pakete auf (gleichmaessig verteilt, inkl. Fach-Artikel-Faktor fuer z.B.
einen schweren "2-Fach-Artikel"-Doppelpack mit "1-je-Paket").
"""

import html
import re

VERSION = "2026-09-23g"

# --- Gewichtsgrenzen in kg (Klasse gilt bei Gewicht STRIKT UNTER der Grenze) -----
G_BRIEF = 0.05
G_GROSSBRIEF = 0.6
G_DPD = 1.1
# DHL-Paket-Maximalgewicht (Matthias bestaetigt) - darueber kann DHL die
# Sendung nicht annehmen, blockiert also den Export (siehe bestimme_carrier()).
G_DHL_MAX = 31.5

BRIEF = "Post Brief"
GROSSBRIEF = "Post Großbrief"
DPD = "DPD"
DHL = "DHL"
KLASSEN = [BRIEF, GROSSBRIEF, DPD, DHL]             # aufsteigend
RANG = {k: i for i, k in enumerate(KLASSEN)}
KENNUNG_KLASSE = {"brx1": BRIEF, "pox1": GROSSBRIEF, "wapo": DPD, "pax1": DHL}
MAX_ZEILENLAENGE = 50                                # DHL: max. 50 Zeichen je Namens-/Adressfeld

# ISO2 -> (ISO3, Ländernamen in Grossschreibung ohne Umlaute)
LAENDER = {
    "DE": ("DEU", ["DEUTSCHLAND", "GERMANY", "ALLEMAGNE"]),
    "AT": ("AUT", ["OESTERREICH", "AUSTRIA", "AUTRICHE"]),
    "CH": ("CHE", ["SCHWEIZ", "SWITZERLAND", "SUISSE"]),
    "LI": ("LIE", ["LIECHTENSTEIN"]),
    "NL": ("NLD", ["NIEDERLANDE", "NETHERLANDS", "HOLLAND", "NEDERLAND"]),
    "BE": ("BEL", ["BELGIEN", "BELGIUM", "BELGIQUE", "BELGIE"]),
    "LU": ("LUX", ["LUXEMBURG", "LUXEMBOURG"]),
    "FR": ("FRA", ["FRANKREICH", "FRANCE"]),
    "IT": ("ITA", ["ITALIEN", "ITALY", "ITALIA"]),
    "ES": ("ESP", ["SPANIEN", "SPAIN", "ESPANA"]),
    "PT": ("PRT", ["PORTUGAL"]),
    "PL": ("POL", ["POLEN", "POLAND", "POLSKA"]),
    "CZ": ("CZE", ["TSCHECHIEN", "CZECHIA", "CZECH REPUBLIC"]),
    "SK": ("SVK", ["SLOWAKEI", "SLOVAKIA"]),
    "HU": ("HUN", ["UNGARN", "HUNGARY"]),
    "SI": ("SVN", ["SLOWENIEN", "SLOVENIA"]),
    "HR": ("HRV", ["KROATIEN", "CROATIA"]),
    "DK": ("DNK", ["DAENEMARK", "DENMARK"]),
    "SE": ("SWE", ["SCHWEDEN", "SWEDEN"]),
    "NO": ("NOR", ["NORWEGEN", "NORWAY"]),
    "FI": ("FIN", ["FINNLAND", "FINLAND"]),
    "IE": ("IRL", ["IRLAND", "IRELAND"]),
    "GB": ("GBR", ["GROSSBRITANNIEN", "UNITED KINGDOM", "VEREINIGTES KOENIGREICH",
                   "ENGLAND"]),
    "GR": ("GRC", ["GRIECHENLAND", "GREECE"]),
    "RO": ("ROU", ["RUMAENIEN", "ROMANIA"]),
    "BG": ("BGR", ["BULGARIEN", "BULGARIA"]),
    "LT": ("LTU", ["LITAUEN", "LITHUANIA"]),
    "LV": ("LVA", ["LETTLAND", "LATVIA"]),
    "EE": ("EST", ["ESTLAND", "ESTONIA"]),
}
_LAND_NAME = {}
for _iso2, (_iso3, _namen) in LAENDER.items():
    for _n in _namen:
        _LAND_NAME[_n] = _iso2
# Ein-Buchstaben-Kuerzel im Postleitzahl-Praefix ("A-8020", "D-12345")
_PRAEFIX_SONDER = {"A": "AT", "D": "DE"}


def iso3(land_iso2):
    """DHL/DPD-Laendercode (ISO 3166 alpha-3), z.B. 'DE' -> 'DEU'; '' wenn unbekannt."""
    return LAENDER.get(land_iso2, ("", []))[0]


# ==========================================================================
# Text-Bereinigung
# ==========================================================================

def bereinige(s):
    """Text fuer einen CSV-Export saeubern: HTML-Entities aufloesen (real im
    Amicron-Export gesehen: '&#34;' im Firmennamen - das enthaltene Semikolon
    zerriss die Zeile und verschob alle Folgespalten), Semikolons und
    Steuerzeichen entfernen, Leerraum zusammenfassen. Beginnt das Ergebnis mit
    einem Formel-Ausloeser (=, +, -, @), wird ein Apostroph vorangestellt -
    schuetzt vor CSV-/Formel-Injection, falls jemand die Export-Datei zur
    Kontrolle in Excel oeffnet (Namen/Adressen stammen letztlich aus
    Kundeneingaben im Onlineshop)."""
    s = html.unescape(s or "")
    s = s.replace(";", ",")
    s = re.sub(r"[\x00-\x1f\x7f]", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    if s and s[0] in "=+-@":
        s = "'" + s
    return s


def _norm_land(txt):
    t = (txt or "").upper()
    for a, b in (("Ä", "AE"), ("Ö", "OE"), ("Ü", "UE"), ("ß", "SS"), ("ẞ", "SS")):
        t = t.replace(a, b)
    return re.sub(r"[^A-Z ]", "", t).strip()


# ==========================================================================
# Adresse
# ==========================================================================

# Zeile, die mit (optionalem Laenderpraefix +) PLZ beginnt: "DE-94081 Fuerstenzell",
# "AT-8330 Feldbach", "42651 Solingen", "A-8020 Graz"
_PLZ_ZEILE = re.compile(
    r"^\s*(?:(?P<pre>[A-Za-z]{1,3})\s?-?\s?)?(?P<plz>\d{4,5})(?!\d)\s*(?P<ort>.*?)\s*$")
# Strasse + PLZ + Ort in EINER Zeile ("Laerchenweg 7 15518 Rauen")
_PLZ_IN_ZEILE = re.compile(r"^(?P<vor>.*\S)\s+(?P<plz>\d{5})\s+(?P<ort>\S.*)$")
# Strasse mit Hausnummer am Ende: "Hauptstr. 12", "Weg 15 B", "Ring 36/6",
# "Allee 12-14". Trenner ist normalerweise ein Leerzeichen, real beobachtet
# aber auch ein Doppelpunkt statt Leerzeichen ("Peheimer Str :11", Rechnung
# 1705571) oder GAR KEIN Trenner - die Hausnummer direkt angeklebt
# ("Rohrwangstr.3", Rechnung 1705548; "Heidelstein str.21", Rechnung 1705551) -
# [\s:]* (0 bis n Zeichen) deckt alle drei Faelle ab. ".*?" ist bewusst
# NICHT-gierig, damit bei mehreren Ziffern in der Zeile trotzdem nur die am
# Ende als Hausnummer erkannt wird.
_STRASSE_HNR = re.compile(
    r"^(?P<str>.*?\S)[\s:]*(?P<nr>\d+\s*[A-Za-z]?(?:\s*[-/]\s*\d+\s*[A-Za-z]?)?)[\s,.;]*$")
# Eine Zeile, die NUR aus der Hausnummer besteht (Strasse und Hausnummer auf
# zwei eigenen Zeilen, real beobachtet an Rechnung 1705611/1705631: "Wiesenweg"
# / "4", "lindenstrasse" / "8").
_NUR_HAUSNUMMER = re.compile(r"^\d+\s*[A-Za-z]?$")


def _finde_strasse(kandidaten):
    """Ermittelt aus den Adresszeilen zwischen Name und PLZ-Zeile, welche die
    Strasse(+Hausnummer) ist - der Rest wird Zusatz. NICHT einfach "die Zeile
    direkt vor der PLZ" (der frueheren Annahme), weil Amicron dort manchmal
    eine ZUSAeTZLICHE Zeile einschiebt: doppelte Strasse ohne Hausnummer,
    Firmenname, Landkreis-Name (real beobachtet an Rechnung 1705605/1705563/
    1705565 - dort stand die zusaetzliche Zeile jeweils NACH der echten
    Strassenzeile, direkt vor der PLZ) - oder Strasse und Hausnummer auf ZWEI
    eigene Zeilen aufteilt (Rechnung 1705611/1705631: "Wiesenweg" / "4").
    Rueckgabe: (strasse_roh, zusatz_liste)."""
    if not kandidaten:
        return "", []
    # 1) Rueckwaerts (naeher an der PLZ zuerst, das ist die ueblichere Position
    #    der echten Strassenzeile) nach der ersten Zeile suchen, die schon fuer
    #    sich allein wie "Strasse + Hausnummer" aussieht.
    for k in range(len(kandidaten) - 1, -1, -1):
        if _STRASSE_HNR.match(kandidaten[k]):
            return kandidaten[k], kandidaten[:k] + kandidaten[k + 1:]
    # 2) Keine Zeile passt einzeln - evtl. Strasse und Hausnummer auf zwei
    #    eigenen Zeilen: ist die letzte Zeile NUR die Hausnummer, mit der Zeile
    #    davor zur Strassenzeile verbinden.
    if len(kandidaten) >= 2 and _NUR_HAUSNUMMER.match(kandidaten[-1]):
        return "%s %s" % (kandidaten[-2], kandidaten[-1]), kandidaten[:-2]
    # 3) Nichts erkannt (z.B. Hausnummer mit einem noch nicht abgedeckten
    #    Trenner) - wie bisher die letzte Zeile nehmen, der Hausnummer-Hinweis
    #    in analysiere_adresse() greift dann weiterhin.
    return kandidaten[-1], kandidaten[:-1]


def analysiere_adresse(zeilen):
    """Zerlegt die Adresszeilen einer Rechnung fuer den CSV-Export.

    Erwartet: Zeile 1 = Name, dann optional Firma/c-o/Zusatzzeilen, dann Strasse,
    dann 'PLZ Ort' (mit optionalem Laenderpraefix), optional eine Laenderzeile.
    Rueckgabe dict: name, zusatz (Liste), strasse, hausnr, plz, ort, land (ISO2 oder
    ''), fehler (Liste, blockiert den Export), hinweise (Liste, nur Warnung)."""
    zl = [bereinige(z) for z in (zeilen or [])]
    zl = [z for z in zl if z]
    out = {"name": "", "zusatz": [], "strasse": "", "hausnr": "", "plz": "",
           "ort": "", "land": "", "fehler": [], "hinweise": [], "packstation": False}
    if not zl:
        out["fehler"].append("Keine Adresse auf der Rechnung erkannt")
        return out
    out["name"] = zl[0]

    # --- PLZ-Zeile suchen (von hinten; nie die Namenszeile) ------------------
    idx, m, kombiniert = None, None, False
    for i in range(len(zl) - 1, 0, -1):
        mm = _PLZ_ZEILE.match(zl[i])
        if mm:
            idx, m = i, mm
            break
    if m is None:
        for i in range(len(zl) - 1, 0, -1):
            mm = _PLZ_IN_ZEILE.match(zl[i])
            if mm:
                idx, m, kombiniert = i, mm, True
                break
    if m is None:
        out["fehler"].append("Keine Postleitzahl gefunden")
        out["zusatz"] = zl[1:]
        return out

    out["plz"] = m.group("plz")
    out["ort"] = m.group("ort").strip()
    if kombiniert:
        strasse_roh = m.group("vor")
        zusatz = zl[1:idx]
        out["hinweise"].append("PLZ/Ort stehen in der Strassenzeile - bitte prüfen")
    else:
        strasse_roh, zusatz = _finde_strasse(zl[1:idx])
    out["zusatz"] = zusatz

    # --- Land -----------------------------------------------------------------
    land = ""
    pre = (m.group("pre") or "") if not kombiniert else ""
    if pre:
        p = pre.upper()
        land = _PRAEFIX_SONDER.get(p, p if p in LAENDER else "")
        if not land:
            out["fehler"].append("Länderkürzel '%s' vor der PLZ unbekannt" % pre)
    rest = zl[idx + 1:]
    if not land and rest:
        cand = _LAND_NAME.get(_norm_land(rest[-1]))
        if cand:
            land = cand
            rest = rest[:-1]
    if rest:
        out["hinweise"].append("Unerwartete Zeile(n) nach der PLZ: %s" % " | ".join(rest))
    if not land:
        # Ohne Laenderkuerzel/-zeile: 5-stellige PLZ -> Deutschland, 4-stellige
        # PLZ -> Oesterreich (beide Defaults, nur als Hinweis - real beobachtet:
        # etliche AT-Rechnungen tragen KEIN "AT-"/keine Laenderzeile, obwohl
        # andere das tun, z.B. Rechnung 1705088 "4595 Waldneukirchen" ohne
        # jeden Hinweis auf Oesterreich). CH/LI haben ebenfalls 4-stellige PLZ,
        # sind hier aber die deutliche Ausnahme - daher AT als Default, mit
        # Hinweis statt stillschweigend, damit es auffaellt.
        if len(out["plz"]) == 5:
            land = "DE"
        elif len(out["plz"]) == 4:
            land = "AT"
            out["hinweise"].append(
                "Kein Länderkürzel/-zeile - 4-stellige PLZ als Österreich angenommen, bitte prüfen")
        else:
            out["fehler"].append("Land unklar (%d-stellige PLZ ohne Länderkürzel/-zeile)"
                                 % len(out["plz"]))
    out["land"] = land
    if land == "DE" and len(out["plz"]) != 5:
        out["fehler"].append("Deutsche PLZ muss 5-stellig sein: %s" % out["plz"])
    if land == "AT" and len(out["plz"]) != 4:
        out["fehler"].append("Österreichische PLZ muss 4-stellig sein: %s" % out["plz"])
    if land and land not in LAENDER:
        out["fehler"].append("Land '%s' nicht in der Länderliste" % land)
    if not out["ort"]:
        out["fehler"].append("Ort fehlt")

    # --- Strasse / Hausnummer -------------------------------------------------
    if not strasse_roh:
        out["fehler"].append("Keine Straße gefunden")
    else:
        ms = _STRASSE_HNR.match(strasse_roh)
        if ms:
            out["strasse"] = ms.group("str").strip()
            out["hausnr"] = re.sub(r"\s+", " ", ms.group("nr")).strip()
        else:
            out["strasse"] = strasse_roh
            out["hinweise"].append("Keine Hausnummer erkannt (%s)" % strasse_roh)
        if re.search(r"packstation|postfach|postnummer", strasse_roh, re.I):
            out["packstation"] = True
            out["hinweise"].append("Packstation/Postfach - bitte manuell prüfen")

    # --- Feld-Pruefungen ------------------------------------------------------
    felder = [out["name"], out["strasse"], out["ort"]] + list(zusatz)
    if any(len(f) > MAX_ZEILENLAENGE for f in felder):
        out["hinweise"].append("Ein Adressfeld ist länger als %d Zeichen" % MAX_ZEILENLAENGE)
    for f in felder:
        try:
            f.encode("cp1252")
        except UnicodeEncodeError:
            out["hinweise"].append("Sonderzeichen in der Adresse nicht darstellbar (%s)" % f)
            break
    return out


# ==========================================================================
# Carrier-Regeln
# ==========================================================================

def klasse_nach_gewicht(gewicht):
    g = round(gewicht, 5)
    if g < G_BRIEF:
        return BRIEF
    if g < G_GROSSBRIEF:
        return GROSSBRIEF
    if g < G_DPD:
        return DPD
    return DHL


def bestimme_carrier(kennungen, gewicht, land, packstation=False):
    """(carrier|None, grund, fehler, hinweise). carrier ist einer aus KLASSEN."""
    kenn = sorted({k.lower() for k in (kennungen or []) if k and k.lower() in KENNUNG_KLASSE})
    fehler, hinweise = [], []
    if not kenn:
        fehler.append("Keine Kennung (Pax1/Pox1/Brx1/Wapo) gefunden")
    if gewicht is None:
        fehler.append("Kein Sendungsgewicht auf der Rechnung")
    elif gewicht <= 0:
        fehler.append("Sendungsgewicht ist 0")
    if fehler:
        return None, "", fehler, hinweise
    # Mehrere Kennungen auf einer Rechnung sind der Normalfall (fast jede
    # Bestellung mischt Artikel mit unterschiedlicher Kennung) und die
    # Hoechste-Klasse-gewinnt-Regel loest das zuverlaessig auf - dafuer KEIN
    # Hinweis mehr (frueher hier, stufte den Status unnoetig auf "warn" hoch
    # und musste ohne echten Mehrwert regelmaessig weggeklickt werden). Die
    # gefundenen Kennungen bleiben trotzdem sichtbar, siehe "grund" unten.
    start = max((KENNUNG_KLASSE[k] for k in kenn), key=RANG.get)
    nach_gew = klasse_nach_gewicht(gewicht)
    klasse = start if RANG[start] >= RANG[nach_gew] else nach_gew
    gtxt = ("%.3f" % gewicht).replace(".", ",")
    grund = "Kennung %s → %s" % ("+".join(kenn), start)
    if klasse != start:
        grund += ", Gewicht %s kg → %s" % (gtxt, klasse)
    if klasse == DPD and packstation:
        # DPD liefert nicht an Packstationen/Postfaecher - Matthias bestaetigt
        # (Rechnung 1705540, Wapo -> waere DPD, Packstation -> auf DHL um).
        klasse = DHL
        grund += ", Packstation/Postfach - DPD nicht möglich → DHL"
    if klasse == DPD and land != "DE":
        klasse = DHL
        grund += ", DPD nur Deutschland → DHL"
    if klasse == DHL and packstation:
        # Gilt unabhaengig davon, WARUM es DHL wurde (Kennung Pax1 direkt oder
        # DPD->DHL-Umschreibung oben) - bei DHL-Packstationszustellung muss das
        # Produkt im DHL-System manuell auf "Kleinpaket" umgestellt werden.
        hinweise.append("Achtung bei DHL auf Kleinpaket abändern")
    if klasse == DHL and gewicht > G_DHL_MAX:
        # DHL nimmt schwerere Pakete nicht an (Matthias bestaetigt) - blockiert
        # den Export, carrier/grund bleiben aber informativ auf DHL stehen
        # (zeigt, wohin es OHNE das Gewichtsproblem gegangen waere); manuelle
        # Nachbearbeitung noetig (z.B. Spedition oder Aufteilen der Sendung).
        fehler.append("Sendungsgewicht %s kg überschreitet das DHL-Maximalgewicht von %s kg"
                      % (gtxt, ("%.1f" % G_DHL_MAX).replace(".", ",")))
    return klasse, grund, fehler, hinweise


def kennungen_der_rechnung(r):
    """Alle Kennungen einer geparsten Rechnung: Kennungszeilen der Positionen
    plus jede Kennung, die in einem Lagerort-Wert steht ('wapo', 'Regal 8, E0,
    Wapo' - Wapo steht bei euch teils nur dort; ausserdem schluckt der Parser
    eine Kennungszeile direkt unter einem leeren 'Lagerort:' als dessen Wert)."""
    kenn = set()
    for p in r.get("positionen") or []:
        for k in p.get("kennungen") or []:
            kenn.add(k.lower())
        for lo in p.get("lagerorte") or []:
            for m in re.finditer(r"\b(pax1|pox1|brx1|wapo)\b", lo or "", re.IGNORECASE):
                kenn.add(m.group(1).lower())
    return kenn


def _ist_versandposition(p):
    """Lokale Kopie von packliste.ist_versand() - carrier_regeln.py bleibt
    bewusst frei von PDF-Abhaengigkeiten (siehe Modulkopf), deshalb hier
    dupliziert statt importiert. Bei Aenderungen an packliste.ist_versand()
    HIER NACHZIEHEN."""
    art, bez = p.get("art"), p.get("bez")
    if "versandkosten" in (bez or "").lower():
        return True
    if not art and not (bez or "").strip():
        return True
    return False


def je_paket_aufteilung(r, gewicht):
    """(pakete|None, grund_zusatz|None) - automatische Paketaufteilung fuer
    eine Rechnung mit GENAU EINER echten Position, die mit "<N>-je-Paket"
    markiert ist (siehe packliste.JE_PAKET_RE, mit Matthias abgestimmt
    2026-09-23). N bezieht sich auf die EFFEKTIVE Stueckzahl (Menge *
    Fach-Artikel-Faktor). Zwei Faelle, je nachdem ob N kleiner oder
    groesser/gleich dem Fach-Artikel-Faktor ist:

    1) N < Fach-Faktor (z.B. "1-je-Paket" bei "2-Fach-Artikel"): der
       Fach-Artikel-Satz selbst wird aufgebrochen und STUECKWEISE verteilt -
       z.B. ein schwerer Doppelpack wird IMMER auf 2 Einzelpakete aufgeteilt.
    2) N >= Fach-Faktor, als Vielfaches gedacht (z.B. "72-je-Paket" bei
       "24-Fach-Artikel" = 3 ganze Kartons je Paket): ganze Kartons werden
       NIE aufgebrochen, nur gebuendelt. Verteilt wird deshalb in ganzen
       Mengen-Einheiten (Kartons), nicht in effektiven Einzelstuecken -
       sonst koennte bei einer ungeraden Kartonzahl ein rechnerischer
       "2,5-Kartons"-Split herauskommen, der physisch nicht packbar ist
       (z.B. 5 Kartons bei "3-je-Paket" -> 3+2 Kartons, NICHT 2,5+2,5).

    In beiden Faellen wird moeglichst GLEICHMAESSIG verteilt (z.B. 4 Kartons
    bei "3-je-Paket" -> 2+2, nicht 3+1). Das Gewicht je Paket wird aus dem
    Rechnungs-Sendungsgewicht abgeleitet - setzt voraus, dass die Rechnung
    AUSSCHLIESSLICH diesen einen Artikel enthaelt (Versandkosten-Zeilen
    zaehlen nicht mit), sonst laesst sich das Gewicht nicht zuverlaessig
    zuordnen.

    (None, None), wenn keine automatische Aufteilung anwendbar/noetig ist -
    u.a. wenn ein Einzelpaket dabei selbst ueber dem DHL-Maximalgewicht laege
    (echter Speditionsfall, keine automatische Loesung moeglich). Dann greift
    weiterhin die normale Gewichtspruefung in bestimme_carrier() plus ggf.
    "Paket aufteilen" von Hand im Dashboard."""
    positionen = [p for p in (r.get("positionen") or []) if not _ist_versandposition(p)]
    if len(positionen) != 1:
        return None, None
    p = positionen[0]
    je_paket = p.get("je_paket")
    menge = p.get("menge")
    if not je_paket or je_paket < 1 or not menge or menge <= 0:
        return None, None
    if gewicht is None or gewicht <= 0:
        return None, None
    if abs(menge - round(menge)) > 1e-6:
        return None, None            # keine ganzzahlige Menge - Sonderfall
    menge_i = int(round(menge))
    fach = p.get("fach") or 1
    effektiv = menge_i * fach

    if je_paket < fach:
        # Fall 1: der Fach-Artikel-Satz wird selbst aufgebrochen -
        # auf effektiver Einzelstueck-Ebene verteilen.
        n_pakete = (effektiv + je_paket - 1) // je_paket
        if n_pakete <= 1:
            return None, None        # passt ohnehin in ein Paket
        basis, rest = divmod(effektiv, n_pakete)
        stueckzahlen = [basis + 1] * rest + [basis] * (n_pakete - rest)
    else:
        # Fall 2: "<N>-je-Paket" = N/Fach ganze Mengen-Einheiten (Kartons)
        # je Paket - in ganzen Kartons verteilen, nie aufbrechen.
        kartons_je_paket = je_paket // fach
        if kartons_je_paket < 1:
            return None, None
        n_pakete = (menge_i + kartons_je_paket - 1) // kartons_je_paket
        if n_pakete <= 1:
            return None, None        # passt ohnehin in ein Paket
        basis, rest = divmod(menge_i, n_pakete)
        kartons_verteilung = [basis + 1] * rest + [basis] * (n_pakete - rest)
        stueckzahlen = [k * fach for k in kartons_verteilung]

    stueckgewicht = gewicht / effektiv
    pakete = [s * stueckgewicht for s in stueckzahlen]
    if any(g > G_DHL_MAX for g in pakete):
        return None, None            # selbst aufgeteilt zu schwer fuer DHL
    grund_zusatz = ("%d-je-Paket (%s) → automatisch %d Pakete"
                    % (je_paket, p.get("art") or "?", n_pakete))
    return pakete, grund_zusatz


def bewerte_rechnung(r):
    """Ergebnis-dict fuer eine geparste Rechnung (siehe carrier_dashboard.py).
    status: 'ok' | 'warn' (Hinweise, Export moeglich) | 'fehler' (blockiert)."""
    fehler, hinweise = [], []
    rnr = r.get("rnr") or ""
    if not rnr:
        fehler.append("Keine Rechnungsnummer erkannt")
    adr = analysiere_adresse(r.get("adresse_zeilen"))
    fehler += adr["fehler"]
    hinweise += adr["hinweise"]
    kenn = kennungen_der_rechnung(r)
    gewicht = r.get("sendungsgewicht")
    carrier, grund, f2, h2 = bestimme_carrier(kenn, gewicht, adr["land"], adr["packstation"])
    fehler += f2
    hinweise += h2
    # "<N>-je-Paket"-markierte Artikel (siehe je_paket_aufteilung()) automatisch
    # in mehrere DHL-Pakete aufteilen - loest dabei ggf. den Uebergewichts-
    # Fehler, den bestimme_carrier() oben anhand des GESAMTgewichts gesetzt
    # hat (der einzelne Pakete koennten ja problemlos unter dem Limit liegen).
    pakete = None
    if carrier == DHL:
        pakete, je_paket_grund = je_paket_aufteilung(r, gewicht)
        if pakete is not None:
            fehler = [f for f in fehler if "DHL-Maximalgewicht" not in f]
            grund += ", " + je_paket_grund
    if not (r.get("zeilen_ok", True) and r.get("summe_ok", True)
            and r.get("vollstaendig_ok", True)):
        hinweise.append("Rechnungsprüfung (Beträge/Vollständigkeit) nicht bestanden")
    kdnr = (r.get("kdnr") or "").strip()
    if carrier == DPD and not kdnr:
        # DPD-Export braucht die Kundennummer als Empfaengerreferenz (siehe
        # carrier_export.py) - fehlt sie, blockiert das den Export NICHT
        # (Feld bleibt leer), wird aber angezeigt, damit es auffaellt.
        hinweise.append("Keine Kundennummer erkannt - Empfängerreferenz bleibt beim DPD-Export leer")
    status = "fehler" if fehler else ("warn" if hinweise else "ok")
    return {
        "rnr": rnr, "datei": r.get("datei", ""), "adresse": adr,
        "kennungen": sorted(kenn), "gewicht": gewicht, "kdnr": kdnr,
        "carrier": carrier, "ausland": bool(adr["land"]) and adr["land"] != "DE",
        "grund": grund, "fehler": fehler, "hinweise": hinweise, "status": status,
        "pakete": pakete,
    }


# ==========================================================================
# Selbsttest:  py carrier_regeln.py
# ==========================================================================

def selftest():
    n_ok, n_fail = 0, []

    def check(name, ist, soll):
        nonlocal n_ok
        if ist == soll:
            n_ok += 1
        else:
            n_fail.append("%s: erwartet %r, war %r" % (name, soll, ist))

    def carrier(kenn, g, land="DE"):
        return bestimme_carrier(kenn, g, land)[0]

    # Kennung/Gewicht
    check("pax1 leicht", carrier(["pax1"], 0.34), DHL)
    check("Pax1 gross geschrieben", carrier(["Pax1"], 0.34), DHL)
    check("wapo 0,5", carrier(["wapo"], 0.5), DPD)
    check("wapo 1,0", carrier(["wapo"], 1.0), DPD)
    check("wapo genau 1,1", carrier(["wapo"], 1.1), DHL)
    check("wapo 1,5", carrier(["wapo"], 1.5), DHL)
    check("wapo Ausland", carrier(["wapo"], 0.5, "AT"), DHL)
    check("pox1 0,3", carrier(["pox1"], 0.3), GROSSBRIEF)
    check("pox1 genau 0,6", carrier(["pox1"], 0.6), DPD)
    check("pox1 0,7 (2x bestellt)", carrier(["pox1"], 0.7), DPD)
    check("pox1 0,7 Ausland", carrier(["pox1"], 0.7, "AT"), DHL)
    check("pox1 1,5", carrier(["pox1"], 1.5), DHL)
    check("brx1 0,04", carrier(["brx1"], 0.04), BRIEF)
    check("brx1 genau 0,05", carrier(["brx1"], 0.05), GROSSBRIEF)
    check("brx1 0,4", carrier(["brx1"], 0.4), GROSSBRIEF)
    check("brx1 0,7", carrier(["brx1"], 0.7), DPD)
    check("brx1 Ausland 0,03", carrier(["brx1"], 0.03, "AT"), BRIEF)
    check("mehrere: pax1+wapo", carrier(["pax1", "wapo"], 0.2), DHL)
    check("mehrere: wapo+pox1", carrier(["wapo", "pox1"], 0.2), DPD)
    check("mehrere Kennungen -> KEIN Hinweis mehr (Normalfall, Rechnung 1705553)",
          bestimme_carrier(["brx1", "pox1", "pax1"], 0.3746, "DE")[3], [])
    check("mehrere Kennungen -> Kennungen trotzdem im Grund sichtbar",
          "brx1" in bestimme_carrier(["brx1", "pox1", "pax1"], 0.3746, "DE")[1], True)
    check("keine Kennung", carrier([], 0.2), None)
    check("kein Gewicht", carrier(["pax1"], None), None)
    check("Gewicht 0", carrier(["pax1"], 0), None)
    check("Grund pox1 aufgestiegen",
          "Gewicht" in bestimme_carrier(["pox1"], 0.7, "DE")[1], True)

    # DHL-Maximalgewicht 31,5 kg (Kundenvorgabe) - darueber blockiert Fehler
    # den Export, Carrier bleibt informativ auf DHL stehen (real getestet an
    # Rechnung 1705548, Pax1 70,4 kg).
    carrier_schwer, _, fehler_schwer, _ = bestimme_carrier(["pax1"], 70.4, "DE")
    check("pax1 70,4 kg -> ueber DHL-Maximalgewicht -> Fehler, Carrier bleibt DHL",
          (carrier_schwer, bool(fehler_schwer)), (DHL, True))
    check("DHL-Maximalgewicht genau 31,5 kg -> noch OK (Grenze gilt NICHT strikt ueberschritten)",
          bestimme_carrier(["pax1"], 31.5, "DE")[2], [])
    check("DHL-Maximalgewicht 31,6 kg -> Fehler",
          bool(bestimme_carrier(["pax1"], 31.6, "DE")[2]), True)
    check("DPD/Post unterhalb 31,5 kg unbetroffen (kein DHL)",
          bestimme_carrier(["wapo"], 1.0, "DE")[2], [])

    # Packstation/Postfach: DPD kann dort nicht zustellen -> DHL, mit Hinweis
    # auf "Kleinpaket" (real getestet an Rechnung 1705540, Wapo 0,2512 kg)
    carrier_ps, grund_ps, _, hinweise_ps = bestimme_carrier(["wapo"], 0.2512, "DE", packstation=True)
    check("wapo Packstation -> DHL statt DPD", carrier_ps, DHL)
    check("wapo Packstation -> Grund nennt Umschreibung", "Packstation" in grund_ps, True)
    check("wapo Packstation -> Kleinpaket-Hinweis", hinweise_ps,
          ["Achtung bei DHL auf Kleinpaket abändern"])
    check("wapo OHNE Packstation -> weiterhin DPD, kein Hinweis",
          bestimme_carrier(["wapo"], 0.2512, "DE", packstation=False), (DPD, "Kennung wapo → DPD", [], []))
    carrier_pax_ps, _, _, hinweise_pax_ps = bestimme_carrier(["pax1"], 3.0, "DE", packstation=True)
    check("pax1 Packstation -> bleibt DHL, aber trotzdem Kleinpaket-Hinweis",
          (carrier_pax_ps, hinweise_pax_ps), (DHL, ["Achtung bei DHL auf Kleinpaket abändern"]))

    # Adresse (Faelle aus echten Rechnungen)
    a = analysiere_adresse(["Evi Schmid", "Jägerwirth 122", "DE-94081 Fürstenzell"])
    check("adr1", (a["name"], a["strasse"], a["hausnr"], a["plz"], a["ort"], a["land"],
                   a["fehler"]),
          ("Evi Schmid", "Jägerwirth", "122", "94081", "Fürstenzell", "DE", []))
    a = analysiere_adresse(["Michael Höfler", "Mühldorf 414", "AT-8330 Feldbach"])
    check("adr AT", (a["plz"], a["ort"], a["land"], a["fehler"]),
          ("8330", "Feldbach", "AT", []))
    a = analysiere_adresse(["Nomi Atelier", "Ebbelicher Weg 70", "45701 Herten"])
    check("adr ohne Praefix", (a["land"], a["plz"], a["hausnr"]), ("DE", "45701", "70"))
    a = analysiere_adresse(["Andreas Denteler", "Denteler Präzisionsteile GmbH",
                            "Feldstraße 16", "DE-86738 Deiningen"])
    check("adr Zusatz", (a["zusatz"], a["strasse"], a["hausnr"]),
          (["Denteler Präzisionsteile GmbH"], "Feldstraße", "16"))
    a = analysiere_adresse(["Jan Peeters", "Rue Haute 5", "1000 Bruxelles", "BELGIEN"])
    check("adr Laenderzeile", (a["land"], a["plz"], a["ort"], a["fehler"]),
          ("BE", "1000", "Bruxelles", []))
    a = analysiere_adresse(["Mirko Krämer", "Lärchenweg 7 15518 Rauen"])
    check("adr PLZ in Strassenzeile", (a["strasse"], a["hausnr"], a["plz"], a["ort"]),
          ("Lärchenweg", "7", "15518", "Rauen"))
    a = analysiere_adresse(["Elke Meier", "Clara Zetkin Straße 358,", "06464 Nachterstedt"])
    check("adr Komma nach Hausnr", (a["strasse"], a["hausnr"]), ("Clara Zetkin Straße", "358"))
    a = analysiere_adresse(["Karl Nummer", "Ring 36/6", "AT-9991 Kals"])
    check("adr Hausnr 36/6", a["hausnr"], "36/6")
    a = analysiere_adresse(["X Y", "Hauptstr. 12 B", "12345 Ort"])
    check("adr Hausnr 12 B", a["hausnr"], "12 B")
    # Zusaetzliche Zeile NACH der echten Strassenzeile (anders als "adr Zusatz"
    # oben, wo die Zusatzzeile VOR der Strasse steht) - real beobachtet, siehe
    # _finde_strasse()-Docstring.
    a = analysiere_adresse(["Vötsch Thomas", "Wilhelmstraße 100", "Wilhelmstraße",
                            "DE-72461 Albstadt"])
    check("adr real 1705605 (doppelte Strasse ohne Hausnr danach)",
          (a["strasse"], a["hausnr"], a["zusatz"], a["hinweise"]),
          ("Wilhelmstraße", "100", ["Wilhelmstraße"], []))
    a = analysiere_adresse(["Roman Schönfeld", "Gerhart-Hauptmann-Str. 13", "Braumanufaktur",
                            "DE-35321 Laubach"])
    check("adr real 1705563 (Zusatzzeile NACH der Strasse)",
          (a["strasse"], a["hausnr"], a["zusatz"], a["hinweise"]),
          ("Gerhart-Hauptmann-Str.", "13", ["Braumanufaktur"], []))
    # Strasse und Hausnummer auf zwei eigenen Zeilen.
    a = analysiere_adresse(["Kerekes Zoltan", "Wiesenweg", "4", "DE-85290 Geisenfeld"])
    check("adr real 1705611 (Strasse/Hausnr auf zwei Zeilen)",
          (a["strasse"], a["hausnr"], a["zusatz"], a["hinweise"]),
          ("Wiesenweg", "4", [], []))
    # Hausnummer mit Doppelpunkt statt Leerzeichen angeklebt.
    a = analysiere_adresse(["Johannes Wanke", "Peheimer Str :11", "DE-49699 Lindern"])
    check("adr real 1705571 (Hausnr mit Doppelpunkt angeklebt)",
          (a["strasse"], a["hausnr"], a["hinweise"]), ("Peheimer Str", "11", []))
    # Hausnummer OHNE jeden Trenner angeklebt.
    a = analysiere_adresse(["Bahittin Doener", "Rohrwangstr.3", "DE-73430 AALEN"])
    check("adr real 1705548 (Hausnr ohne Trenner angeklebt)",
          (a["strasse"], a["hausnr"], a["hinweise"]), ("Rohrwangstr.", "3", []))
    a = analysiere_adresse(["Gregor Worringen", "Heidelstein str.21", "DE-36043 Fulda"])
    check("adr real 1705551 (Hausnr ohne Trenner angeklebt, mit Leerzeichen im Strassennamen)",
          (a["strasse"], a["hausnr"], a["hinweise"]), ("Heidelstein str.", "21", []))
    a = analysiere_adresse(["X Y", "Hauptstr. 1"])
    check("adr ohne PLZ -> Fehler", bool(a["fehler"]), True)
    a = analysiere_adresse(["X Y", "Hauptstr. 1", "8330 Feldbach"])
    check("adr 4-stellig ohne Land -> AT-Default mit Hinweis, kein Fehler",
          (a["land"], a["fehler"], bool(a["hinweise"])), ("AT", [], True))
    a = analysiere_adresse(["Gegenleitner Johannes", "Bad Haller Straße 56",
                            "4595 Waldneukirchen"])
    check("adr real 1705088 (AT ohne Länderzeile)",
          (a["land"], a["plz"], a["ort"], a["fehler"]), ("AT", "4595", "Waldneukirchen", []))
    a = analysiere_adresse(["X Y", "Hauptstr. 1", "123 Ort"])
    check("adr 3-stellig ohne Land -> weiterhin Fehler", bool(a["fehler"]), True)
    a = analysiere_adresse(["Dienstleistungs-GmbH der WG &#34;Kohle Geiseltal&#34;",
                            "Am Stadion 1", "06242 Braunsbedra"])
    check("HTML-Entity/Semikolon bereinigt", ";" in a["name"] or "&#" in a["name"], False)
    check("bereinige Semikolon", bereinige("A; B"), "A, B")
    check("bereinige Formel-Injection =", bereinige("=HYPERLINK(\"x\")"), "'=HYPERLINK(\"x\")")
    check("bereinige Formel-Injection @", bereinige("@SUM(1)"), "'@SUM(1)")
    check("bereinige normaler Name unveraendert", bereinige("Müller GmbH"), "Müller GmbH")
    a = analysiere_adresse(["Christopher Beck", "867653498", "Packstation 105",
                            "DE-68307 Mannheim"])
    check("adr real 1705540 (Packstation) -> packstation=True",
          (a["land"], a["plz"], a["ort"], a["packstation"], a["fehler"]),
          ("DE", "68307", "Mannheim", True, []))
    a = analysiere_adresse(["X Y", "Postfach 12", "12345 Ort"])
    check("adr Postfach -> packstation=True", a["packstation"], True)
    a = analysiere_adresse(["Evi Schmid", "Jägerwirth 122", "DE-94081 Fürstenzell"])
    check("adr normale Strasse -> packstation=False", a["packstation"], False)

    # Rechnung komplett (wie parse_pdf sie liefert)
    r = {"rnr": "1705050", "datei": "x.pdf", "sendungsgewicht": 0.34,
         "adresse_zeilen": ["Felix Cox", "Heidelberger Str. 80a", "DE-12435 Berlin"],
         "positionen": [{"kennungen": ["pax1"], "lagerorte": ["Regal 8, E1", "Topseller"]}],
         "zeilen_ok": True, "summe_ok": True, "vollstaendig_ok": True}
    b = bewerte_rechnung(r)
    check("Rechnung 1705050", (b["carrier"], b["status"], b["ausland"]), (DHL, "ok", False))
    r2 = dict(r, positionen=[{"kennungen": [], "lagerorte": ["wapo"]}], sendungsgewicht=0.4)
    check("Wapo nur als Lagerort", bewerte_rechnung(r2)["carrier"], DPD)
    r3 = dict(r, positionen=[{"kennungen": [], "lagerorte": ["Pax1"]}])
    check("Pax1 als Lagerort-Wert verschluckt", bewerte_rechnung(r3)["carrier"], DHL)
    r4 = dict(r, kdnr="674982", positionen=[{"kennungen": ["wapo"], "lagerorte": []}],
              sendungsgewicht=0.5)
    b4 = bewerte_rechnung(r4)
    check("DPD mit Kundennummer -> ok, kdnr durchgereicht",
          (b4["status"], b4["kdnr"]), ("ok", "674982"))
    r5 = dict(r4, kdnr="")
    b5 = bewerte_rechnung(r5)
    check("DPD ohne Kundennummer -> warn, kdnr leer", (b5["status"], b5["kdnr"]), ("warn", ""))

    r6 = {"rnr": "1705540", "datei": "x.pdf", "sendungsgewicht": 0.2512,
          "adresse_zeilen": ["Christopher Beck", "867653498", "Packstation 105",
                             "DE-68307 Mannheim"],
          "positionen": [{"kennungen": ["wapo"], "lagerorte": ["Regal 8, E0"]}],
          "zeilen_ok": True, "summe_ok": True, "vollstaendig_ok": True}
    b6 = bewerte_rechnung(r6)
    check("Rechnung 1705540 (Packstation) -> DHL statt DPD, warn", (b6["carrier"], b6["status"]),
          (DHL, "warn"))
    check("Rechnung 1705540 -> Kleinpaket-Hinweis dabei",
          "Achtung bei DHL auf Kleinpaket abändern" in b6["hinweise"], True)

    # --- "<N>-je-Paket" automatische Paketaufteilung (2026-09-23, mit
    # Matthias abgestimmt: Doppelpack-Artikel IMMER einzeln, Kartons-Artikel
    # zu mehreren buendelbar, aber gleichmaessig statt gierig aufgeteilt) ---
    def _pos(art, menge, je_paket=None, fach=None, bez="Artikel"):
        return {"art": art, "bez": bez, "menge": menge, "kennungen": ["pax1"],
                "lagerorte": [], "fach": fach, "je_paket": je_paket}

    r_doppelpack = {"rnr": "1700900", "datei": "x.pdf", "sendungsgewicht": 40.0,
                    "adresse_zeilen": ["A B", "Weg 1", "12345 Ort"],
                    "positionen": [_pos("DP1", 1.0, je_paket=1, fach=2, bez="Schwerer Doppelpack")],
                    "zeilen_ok": True, "summe_ok": True, "vollstaendig_ok": True}
    b_dp = bewerte_rechnung(r_doppelpack)
    check("Doppelpack (1-je-Paket, 2-Fach-Artikel) -> automatisch 2 Pakete",
          b_dp.get("pakete"), [20.0, 20.0])
    check("Doppelpack -> kein Uebergewichts-Fehler mehr (jedes Paket 20 kg)",
          b_dp["fehler"], [])
    check("Doppelpack -> Grund nennt die Aufteilung", "je-Paket" in b_dp["grund"], True)

    # Kartons-Artikel, "3-je-Paket": 4 Kartons -> 2+2 (nicht 3+1), auch wenn
    # das Gesamtgewicht (hier klein) gar keinen Fehler ausgeloest haette.
    r_4kartons = {"rnr": "1700901", "datei": "x.pdf", "sendungsgewicht": 8.0,
                 "adresse_zeilen": ["A B", "Weg 1", "12345 Ort"],
                 "positionen": [_pos("KA1", 4.0, je_paket=3, bez="Karton-Artikel")],
                 "zeilen_ok": True, "summe_ok": True, "vollstaendig_ok": True}
    # sendungsgewicht=8,0 kg / 4 Kartons = 2,0 kg/Karton -> Stueckzahlen [2,2]
    # (gleichmaessig, NICHT [3,1]) -> Paketgewichte je 2x2,0 kg = [4.0, 4.0]
    check("4 Kartons bei 3-je-Paket -> 2+2 Stueck (gleichmaessig, nicht 3+1)",
          bewerte_rechnung(r_4kartons)["pakete"], [4.0, 4.0])

    r_6kartons = dict(r_4kartons, rnr="1700902",
                      positionen=[_pos("KA1", 6.0, je_paket=3, bez="Karton-Artikel")],
                      sendungsgewicht=12.0)
    # 12,0 kg / 6 Kartons = 2,0 kg/Karton, Stueckzahlen [3,3] -> [6.0, 6.0]
    check("6 Kartons bei 3-je-Paket -> 3+3 Stueck", bewerte_rechnung(r_6kartons)["pakete"],
          [6.0, 6.0])

    r_5kartons = dict(r_4kartons, rnr="1700903",
                      positionen=[_pos("KA1", 5.0, je_paket=3, bez="Karton-Artikel")],
                      sendungsgewicht=10.0)
    # 10,0 kg / 5 Kartons = 2,0 kg/Karton, Stueckzahlen [3,2] -> [6.0, 4.0]
    check("5 Kartons bei 3-je-Paket -> 3+2 Stueck", bewerte_rechnung(r_5kartons)["pakete"],
          [6.0, 4.0])

    r_3kartons = dict(r_4kartons, rnr="1700904",
                      positionen=[_pos("KA1", 3.0, je_paket=3, bez="Karton-Artikel")],
                      sendungsgewicht=6.0)
    check("3 Kartons bei 3-je-Paket -> passt in 1 Paket, keine Aufteilung",
          bewerte_rechnung(r_3kartons)["pakete"], None)

    # Kombination Fach-Artikel + je-Paket (Brennpasten-Beispiel, mit Matthias
    # verifiziert 2026-09-23: 1 Karton = 24-Fach-Artikel, 3 Kartons pro Paket
    # buendelbar -> "72-je-Paket"). Wichtig: bei 5 bestellten Kartons MUSS in
    # ganzen Kartons (3+2) verteilt werden, NICHT in effektiven Einzelstuecken
    # (das ergaebe rechnerisch 2,5+2,5 Kartons - physisch nicht packbar; das
    # war ein Bug in der ersten Fassung dieser Funktion).
    r_brennpaste = {"rnr": "1700907", "datei": "x.pdf", "sendungsgewicht": 50.0,
                    "adresse_zeilen": ["A B", "Weg 1", "12345 Ort"],
                    "positionen": [_pos("BP24", 5.0, je_paket=72, fach=24, bez="Brennpaste 24er")],
                    "zeilen_ok": True, "summe_ok": True, "vollstaendig_ok": True}
    # 50,0 kg / 120 Stueck = 0,41(6) kg/Stueck; Kartons [3,2] -> [72,48] Stueck
    # -> Paketgewichte [30.0, 20.0] (3/5 bzw. 2/5 von 50 kg, ganze Kartons)
    check("5 Kartons Brennpaste (24-Fach-Artikel) bei 72-je-Paket -> 3+2 Kartons",
          bewerte_rechnung(r_brennpaste)["pakete"], [30.0, 20.0])

    r_brennpaste_3 = dict(r_brennpaste, rnr="1700908",
                          positionen=[_pos("BP24", 3.0, je_paket=72, fach=24, bez="Brennpaste 24er")],
                          sendungsgewicht=30.0)
    check("3 Kartons Brennpaste bei 72-je-Paket -> passt genau in 1 Paket",
          bewerte_rechnung(r_brennpaste_3)["pakete"], None)

    # Mischbestellung (Einzelversand-Artikel + weiterer echter Artikel) -> KEINE
    # automatische Aufteilung, da sich das Gesamtgewicht nicht zuverlaessig
    # zuordnen laesst.
    r_misch = {"rnr": "1700905", "datei": "x.pdf", "sendungsgewicht": 45.0,
              "adresse_zeilen": ["A B", "Weg 1", "12345 Ort"],
              "positionen": [_pos("DP1", 1.0, je_paket=1, fach=2, bez="Doppelpack"),
                            _pos("K2", 1.0, bez="Kleinteil")],
              "zeilen_ok": True, "summe_ok": True, "vollstaendig_ok": True}
    b_misch = bewerte_rechnung(r_misch)
    check("Mischbestellung mit je-Paket-Artikel -> KEINE Auto-Aufteilung",
          b_misch.get("pakete"), None)
    check("Mischbestellung -> bleibt reeller Uebergewichts-Fehler (45 kg > 31,5 kg)",
          bool(b_misch["fehler"]), True)

    # Selbst aufgeteilt noch zu schwer fuer DHL -> keine automatische Loesung,
    # bleibt Fehler (echter Speditionsfall).
    r_zuschwer = {"rnr": "1700906", "datei": "x.pdf", "sendungsgewicht": 100.0,
                 "adresse_zeilen": ["A B", "Weg 1", "12345 Ort"],
                 "positionen": [_pos("DP1", 1.0, je_paket=1, fach=2, bez="Riesending")],
                 "zeilen_ok": True, "summe_ok": True, "vollstaendig_ok": True}
    b_zuschwer = bewerte_rechnung(r_zuschwer)
    check("Auch aufgeteilt (50+50 kg) noch ueber DHL-Maximalgewicht -> Fehler bleibt",
          (b_zuschwer.get("pakete"), bool(b_zuschwer["fehler"])), (None, True))

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
