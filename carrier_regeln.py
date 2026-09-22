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
"""

import html
import re

VERSION = "2026-09-21a"

# --- Gewichtsgrenzen in kg (Klasse gilt bei Gewicht STRIKT UNTER der Grenze) -----
G_BRIEF = 0.05
G_GROSSBRIEF = 0.6
G_DPD = 1.1

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
    Steuerzeichen entfernen, Leerraum zusammenfassen."""
    s = html.unescape(s or "")
    s = s.replace(";", ",")
    s = re.sub(r"[\x00-\x1f\x7f]", " ", s)
    return re.sub(r"\s+", " ", s).strip()


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
# Strasse mit Hausnummer am Ende: "Hauptstr. 12", "Weg 15 B", "Ring 36/6", "Allee 12-14"
_STRASSE_HNR = re.compile(
    r"^(?P<str>.*?\S)\s+(?P<nr>\d+\s*[A-Za-z]?(?:\s*[-/]\s*\d+\s*[A-Za-z]?)?)[\s,.;]*$")


def analysiere_adresse(zeilen):
    """Zerlegt die Adresszeilen einer Rechnung fuer den CSV-Export.

    Erwartet: Zeile 1 = Name, dann optional Firma/c-o/Zusatzzeilen, dann Strasse,
    dann 'PLZ Ort' (mit optionalem Laenderpraefix), optional eine Laenderzeile.
    Rueckgabe dict: name, zusatz (Liste), strasse, hausnr, plz, ort, land (ISO2 oder
    ''), fehler (Liste, blockiert den Export), hinweise (Liste, nur Warnung)."""
    zl = [bereinige(z) for z in (zeilen or [])]
    zl = [z for z in zl if z]
    out = {"name": "", "zusatz": [], "strasse": "", "hausnr": "", "plz": "",
           "ort": "", "land": "", "fehler": [], "hinweise": []}
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
        strasse_roh = zl[idx - 1] if idx - 1 >= 1 else ""
        zusatz = zl[1:idx - 1] if idx - 1 >= 1 else []
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
        if len(out["plz"]) == 5:
            land = "DE"
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


def bestimme_carrier(kennungen, gewicht, land):
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
    if len(kenn) > 1:
        hinweise.append("Mehrere Kennungen (%s) - höchste Klasse gilt" % ", ".join(kenn))
    start = max((KENNUNG_KLASSE[k] for k in kenn), key=RANG.get)
    nach_gew = klasse_nach_gewicht(gewicht)
    klasse = start if RANG[start] >= RANG[nach_gew] else nach_gew
    gtxt = ("%.3f" % gewicht).replace(".", ",")
    grund = "Kennung %s → %s" % ("+".join(kenn), start)
    if klasse != start:
        grund += ", Gewicht %s kg → %s" % (gtxt, klasse)
    if klasse == DPD and land != "DE":
        klasse = DHL
        grund += ", DPD nur Deutschland → DHL"
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
    carrier, grund, f2, h2 = bestimme_carrier(kenn, gewicht, adr["land"])
    fehler += f2
    hinweise += h2
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
    check("keine Kennung", carrier([], 0.2), None)
    check("kein Gewicht", carrier(["pax1"], None), None)
    check("Gewicht 0", carrier(["pax1"], 0), None)
    check("Grund pox1 aufgestiegen",
          "Gewicht" in bestimme_carrier(["pox1"], 0.7, "DE")[1], True)

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
    a = analysiere_adresse(["X Y", "Hauptstr. 1"])
    check("adr ohne PLZ -> Fehler", bool(a["fehler"]), True)
    a = analysiere_adresse(["X Y", "Hauptstr. 1", "8330 Feldbach"])
    check("adr 4-stellig ohne Land -> Fehler", bool(a["fehler"]), True)
    a = analysiere_adresse(["Dienstleistungs-GmbH der WG &#34;Kohle Geiseltal&#34;",
                            "Am Stadion 1", "06242 Braunsbedra"])
    check("HTML-Entity/Semikolon bereinigt", ";" in a["name"] or "&#" in a["name"], False)
    check("bereinige Semikolon", bereinige("A; B"), "A, B")

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
