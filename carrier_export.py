# -*- coding: utf-8 -*-
"""
carrier_export.py - CSV-Export fuer DHL, DPD und Deutsche Post
================================================================================
Carrier-Dashboard Schritt 2: schreibt aus den von carrier_regeln.bewerte_rechnung()
bewerteten Rechnungen die Import-CSVs fuer die Versanddienstleister nach
C:\\Carrier_Export. Reine Schreiblogik ohne GUI - Selbsttest via
py carrier_export.py.

Exportiert wird NUR, was status == "ok" UND eine Carrier-Zuordnung hat - siehe
exportiere(). Sowohl ein Fehler als auch ein noch nicht quittierter Hinweis
(status "fehler" bzw. "warn") blockieren den Export, bis sie behoben bzw. im
Dashboard ueber "Hinweis quittieren" bestaetigt wurden.

Dateinamen (mit Matthias abgestimmt, 2026-09-22): <Praefix>_<JJJJ-MM-TT>_<HHMM>.csv,
z.B. DHL_2026-09-22_161000.csv, Post_Brief_Ausland_2026-09-22_161000.csv. Jeder Lauf
legt neue Dateien an (Zeitstempel im Namen) - nichts wird ueberschrieben oder
gemergt. Eine leere Gruppe erzeugt keine Datei.

FORMATE (anhand der Amicron/Faktura-Musterdateien nachgebildet, mit Matthias
abgeglichen am 2026-09-21/22 - siehe \\DESKTOP-N2H75H\\Netzwerk\\KI Training\\
DHL_RG1703056.csv, DPD_RG1705331.csv, Brief_RG1705321.csv, Großbrief_RG1705349.csv):

  DHL/DPD: identisches 32-Spalten-Layout, Semikolon, ISO-8859-1, CRLF.
    - Sendungsreferenz UND Empfängerreferenz tragen bei DHL BEIDE die
      Rechnungsnummer (DHL druckt das beim Import als "Kostenstelle" aufs
      Label und gibt es im eigenen Rueckexport wieder genau so aus).
    - Bei DPD ist die Empfängerreferenz dagegen die Amicron-Kundennummer;
      die vier "Service - Nachnahme"-Felder IBAN/BIC/Zahlungsempfänger/
      Bankname sind zweckentfremdet und tragen IMMER die festen Werte
      34/23/5/16 (von Matthias bestaetigt, unabhaengig von der Rechnung).
    - Empfänger Telefonnummer: bei DHL zweckentfremdet zu " "+Rechnungsnummer
      (fuehrendes Leerzeichen), bei DPD leer - beides aus den Mustern uebernommen.
    - DHL-Produkt/Abrechnungsnummer nach Empfängerland: DE -> V01PAK /
      52148008630101, alle anderen Laender (nur AT im Muster belegt) ->
      V53WPAK.V53VV / 52148008635303 (Default fuer nicht belegte Laender,
      bis ein echter Fall etwas anderes zeigt). DPD faehrt nur Inland, daher
      immer das deutsche Produkt.
  Post (Brief/Grossbrief, je Inland/Ausland): 9 Spalten, Semikolon,
    ISO-8859-1, CRLF. ERSTE Zeile jeder Datei ist IMMER die Absenderadresse
    (ADRESS_TYP "HOUSE", keine REFERENZ), danach je eine Zeile pro Rechnung.
"""

import csv
import os
from datetime import datetime

import carrier_regeln as regeln

VERSION = "2026-09-25b"

# ---------------------------------------------------------------------------
# Absenderdaten (fest - aus den Musterdateien uebernommen; DHL/DPD nutzen
# eine andere Schreibweise der Strasse als Post, bewusst je Format beibehalten)
# ---------------------------------------------------------------------------
ABSENDER_DHL_DPD = {
    "name1": "Augsburger Gase- und Handelsgesellschaft mbH",
    "strasse": "Karlsruher Straße", "hausnr": "3", "plz": "86156",
    "ort": "Augsburg", "land": "DEU",
    "email": "onlinerechnung@Gasecenter-onlineshop.de",
    "telefon": "0821/21 86 56 7",
}
ABSENDER_POST = {
    "name": "Gasecenter Augsburg", "zusatz": "", "strasse": "Karlsruher Str.",
    "hausnr": "3", "plz": "86156", "ort": "Augsburg", "land": "DE",
}

DHL_DPD_SPALTEN = [
    "Sendungsreferenz", "Sendungsdatum", "Absender Name 1", "Absender Name 2",
    "Absender Name 3", "Absender Straße", "Absender Hausnummer",
    "Absender PLZ", "Absender Ort", "Absender Land",
    "Absender E-Mail-Adresse", "Absender Telefonnummer", "Empfänger Name 1",
    "Empfänger Name 2 / Postnummer", "Empfänger Name 3", "Empfänger Straße",
    "Empfänger Hausnummer", "Empfänger PLZ", "Empfänger Ort",
    "Empfänger Land", "Empfänger E-Mail-Adresse", "Empfänger Telefonnummer",
    "Gewicht", "Empfängerreferenz", "Produkt- und Servicedetails",
    "Abrechnungsnummer", "Service - Nachnahme - Betrag",
    "Service - Nachnahme - IBAN", "Service - Nachnahme - BIC",
    "Service - Nachnahme - Zahlungsempfänger", "Service - Nachnahme - Bankname",
    "Service - Nachnahme - Verwendungszweck 1",
]
assert len(DHL_DPD_SPALTEN) == 32

POST_SPALTEN = ["NAME", "ZUSATZ", "STRASSE", "NUMMER", "PLZ", "STADT", "LAND",
                "ADRESS_TYP", "REFERENZ"]

# DHL-Produkt + Abrechnungsnummer je Empfaengerland (ISO2). DE ist real belegt,
# alle anderen Laender fallen auf den AT-Fall zurueck (s.o.).
DHL_PRODUKT = {"DE": ("V01PAK", "52148008630101")}
DHL_PRODUKT_DEFAULT = ("V53WPAK.V53VV", "52148008635303")
DPD_PRODUKT = ("V01PAK", "52148008630101")          # DPD: immer Inland
# DHL Kleinpaket Inland (per Test-Import bei DHL bestaetigt 2026-09-25) -
# aktuell von keiner Regel genutzt, aber fuer spaeter bereit.
DHL_KLEINPAKET = ("V62KP", "52148008636201")
# Auslands-"Kleinpaket" (b["kleinpaket"], siehe carrier_regeln.
# ist_auslands_kleinpaket) ist bei DHL das Produkt "Warenpost International":
# der Import mit V62KP wurde abgelehnt, Matthias hat das Label im Portal
# korrigiert -> Abrechnungsnummer 52148008636601 (Verfahren 66). Produktcode
# V66WPI = DHL-Standardcode fuer Warenpost International; per Test-Import zu
# bestaetigen.
DHL_KLEINPAKET_AUSLAND = ("V66WPI", "52148008636601")
DPD_FESTWERTE = ("34", "23", "5", "16")             # IBAN/BIC/Zahlungsempf./Bankname


def _gewicht_txt(g):
    """Gewicht wie in den Mustern: deutsches Komma, 4 Nachkommastellen."""
    return ("%.4f" % (g or 0)).replace(".", ",")


def _name2_name3(a):
    """Zusatzzeilen (Firma, c/o, Postnummer ...) + Ortsteil -> (Name 2, Name 3).
    Der Ortsteil steht IMMER in Name 3 (nie im Ort - damit hat DHL oft
    Probleme). 1. Zusatzzeile -> Name 2, weitere mit ", " in Name 3 (vor dem
    Ortsteil; carrier_regeln setzt bei >2 Zeilen einen Hinweis)."""
    zusatz = [z for z in (a.get("zusatz") or []) if z]
    rest = zusatz[1:] + ([a["ortsteil"]] if a.get("ortsteil") else [])
    return (zusatz[0] if zusatz else "", ", ".join(rest))


def _dhl_dpd_zeile(b, ist_dpd, gewicht=None):
    """gewicht ueberschreibt optional b["gewicht"] (Gesamtgewicht laut
    Rechnung) - fuer eine manuell in zwei Pakete aufgeteilte DHL-Sendung
    (b["pakete"]) wird diese Funktion einmal je Einzelgewicht aufgerufen,
    siehe _dhl_zeilen()."""
    a = b["adresse"]
    name2, name3 = _name2_name3(a)
    if ist_dpd:
        produkt, abrechnung = DPD_PRODUKT
        empf_ref = b.get("kdnr") or ""
        empf_telefon = ""
        nachnahme = ("", *DPD_FESTWERTE, "")
    else:
        if b.get("kleinpaket"):
            produkt, abrechnung = (DHL_KLEINPAKET if a["land"] == "DE"
                                   else DHL_KLEINPAKET_AUSLAND)
        else:
            produkt, abrechnung = DHL_PRODUKT.get(a["land"], DHL_PRODUKT_DEFAULT)
        empf_ref = b["rnr"]                     # DHL: Referenz = Rechnungsnummer
        empf_telefon = " " + b["rnr"]            # zweckentfremdet, siehe Modul-Kopf
        nachnahme = ("", "", "", "", "", "")
    ab = ABSENDER_DHL_DPD
    zeile = [
        b["rnr"], "",
        ab["name1"], "", "",
        ab["strasse"], ab["hausnr"], ab["plz"], ab["ort"], ab["land"],
        ab["email"], ab["telefon"],
        a["name"], name2, name3,                 # Name nicht getrennt (Kundenwunsch)
        a["strasse"], a["hausnr"], a["plz"], a["ort"], regeln.iso3(a["land"]),
        "", empf_telefon,
        _gewicht_txt(gewicht if gewicht is not None else b["gewicht"]),
        empf_ref, produkt, abrechnung,
        *nachnahme,
    ]
    assert len(zeile) == 32
    return zeile


def _dhl_zeilen(b):
    """Eine ODER MEHRERE DHL-Zeilen fuer b: normalerweise eine mit dem
    Gesamtgewicht, bei aufgeteilten Sendungen (b["pakete"] - entweder manuell
    via carrier_dashboard.paket_aufteilen() bei Uebergewicht, oder automatisch
    via carrier_regeln.je_paket_aufteilung() fuer "<N>-je-Paket"-markierte
    Artikel) eine Zeile PRO PAKET mit dem jeweiligen Einzelgewicht - keine
    Obergrenze fuer die Paketanzahl (auch sehr schwere Sendungen, die selbst
    in zwei Pakete a 31,5 kg nicht mehr
    passen, z.B. Rechnung 1705548 mit 70,4 kg -> 3 Pakete noetig). Alle
    Pakete tragen bewusst DIESELBE Sendungsreferenz (Rechnungsnummer) - mit
    Matthias abgestimmt 2026-09-23, sein DHL-Geschaeftskundenportal
    akzeptiert doppelte Sendungsreferenzen innerhalb eines Import-Laufs."""
    pakete = b.get("pakete")
    if not pakete:
        return [_dhl_dpd_zeile(b, False)]
    return [_dhl_dpd_zeile(b, False, gewicht=g) for g in pakete]


def _post_zeile(a, rnr):
    # ", " statt "; " - ein rohes Semikolon im Feldwert ist genau das Muster,
    # das bereinige() im Rest der Datei bewusst vermeidet (siehe dortiger
    # Kommentar zur zerrissenen Zeile durch ein Semikolon im Firmennamen).
    zusatz = ", ".join(list(a.get("zusatz") or [])
                       + ([a["ortsteil"]] if a.get("ortsteil") else []))
    return [a["name"], zusatz, a["strasse"], a["hausnr"], a["plz"], a["ort"],
            a["land"], "HOUSE", rnr]


def _post_absenderzeile():
    ap = ABSENDER_POST
    return [ap["name"], ap["zusatz"], ap["strasse"], ap["hausnr"], ap["plz"],
            ap["ort"], ap["land"], "HOUSE", ""]


def dateiname(ziel_ordner, praefix, jetzt=None, ext="csv"):
    """Eindeutiger Pfad <Praefix>_<JJJJ-MM-TT>_<HHMMSS>.<ext> in ziel_ordner.
    Sekundengenauer Stempel PLUS Kollisions-Suffix (_2, _3, ...), falls im
    Ordner schon eine Datei mit demselben Stempel liegt (z.B. zwei Laeufe
    binnen derselben Sekunde) - vorher reichte die Minute (HHMM) allein nicht,
    ein zweiter Lauf in derselben Minute hat eine bereits geschriebene Datei
    (Carrier-CSV oder Pickliste) still ueberschrieben."""
    stamp = (jetzt or datetime.now()).strftime("%Y-%m-%d_%H%M%S")
    basis = "%s_%s" % (praefix, stamp)
    pfad = os.path.join(ziel_ordner, "%s.%s" % (basis, ext))
    n = 2
    while os.path.exists(pfad):
        pfad = os.path.join(ziel_ordner, "%s_%d.%s" % (basis, n, ext))
        n += 1
    return pfad


def _schreibe_csv(pfad, header, zeilen):
    # ISO-8859-1/CRLF wie die Amicron-Musterdateien; errors="replace" als
    # letzte Sicherung, falls ein Zeichen ausserhalb Latin-1 durchrutscht -
    # der eigentliche Hinweis dazu kommt schon aus analysiere_adresse().
    # "x" statt "w": dateiname() liefert zwar bereits einen freien Pfad, aber
    # "x" schlaegt hart fehl statt eine inzwischen doch vorhandene Datei
    # stillschweigend zu ueberschreiben (Restrisiko bei zwei Laeufen exakt
    # zeitgleich).
    with open(pfad, "x", encoding="iso-8859-1", errors="replace", newline="") as f:
        w = csv.writer(f, delimiter=";", lineterminator="\r\n")
        w.writerow(header)
        for z in zeilen:
            w.writerow(z)


def exportiere(ergebnisse, ziel_ordner, jetzt=None):
    """Schreibt alle Carrier-CSVs fuer die exportierbaren Ergebnisse (status
    == 'ok' UND Carrier zugeordnet) nach ziel_ordner (wird ggf. angelegt).
    status == 'ok' schliesst sowohl 'fehler' (blockiert) als auch offene,
    NICHT quittierte Hinweise ('warn') aus - ein Hinweis (z.B. unsichere
    Laendererkennung) blockiert den Export also genau wie ein Fehler, bis er
    im Dashboard ueber "Hinweis quittieren" bestaetigt wurde (das setzt
    status auf 'ok', siehe carrier_dashboard.quittiere_auswahl()). Rueckgabe:
    {dateipfad: anzahl_rechnungen}, nur tatsaechlich geschriebene Dateien
    (leere Gruppen erzeugen keine Datei)."""
    os.makedirs(ziel_ordner, exist_ok=True)
    exportierbar = [b for b in ergebnisse if b["status"] == "ok" and b["carrier"]]
    geschrieben = {}

    dhl = [b for b in exportierbar if b["carrier"] == regeln.DHL]
    if dhl:
        pfad = dateiname(ziel_ordner, "DHL", jetzt)
        zeilen_dhl = [z for b in dhl for z in _dhl_zeilen(b)]
        _schreibe_csv(pfad, DHL_DPD_SPALTEN, zeilen_dhl)
        # anzahl_rechnungen bleibt die Rechnungsanzahl (nicht die Zeilenzahl) -
        # konsistent mit den anderen Gruppen; eine aufgeteilte Sendung liefert
        # zwei Zeilen, zaehlt hier aber weiterhin als eine Rechnung.
        geschrieben[pfad] = len(dhl)

    dpd = [b for b in exportierbar if b["carrier"] == regeln.DPD]
    if dpd:
        pfad = dateiname(ziel_ordner, "DPD", jetzt)
        _schreibe_csv(pfad, DHL_DPD_SPALTEN, [_dhl_dpd_zeile(b, True) for b in dpd])
        geschrieben[pfad] = len(dpd)

    for carrier, praefix in ((regeln.BRIEF, "Post_Brief"), (regeln.GROSSBRIEF, "Post_Grossbrief")):
        for ausland, suffix in ((False, "Inland"), (True, "Ausland")):
            gruppe = [b for b in exportierbar
                      if b["carrier"] == carrier and b["ausland"] == ausland]
            if not gruppe:
                continue
            pfad = dateiname(ziel_ordner, "%s_%s" % (praefix, suffix), jetzt)
            zeilen = [_post_absenderzeile()] + [_post_zeile(b["adresse"], b["rnr"])
                                                for b in gruppe]
            _schreibe_csv(pfad, POST_SPALTEN, zeilen)
            geschrieben[pfad] = len(gruppe)

    return geschrieben


# ==========================================================================
# Selbsttest:  py carrier_export.py
# ==========================================================================

def _bsp(rnr, kennung, gewicht, adresse_zeilen, kdnr="", lagerorte=None):
    r = {"rnr": rnr, "datei": rnr + ".pdf", "sendungsgewicht": gewicht,
         "adresse_zeilen": adresse_zeilen, "kdnr": kdnr,
         "positionen": [{"kennungen": [kennung] if kennung else [],
                        "lagerorte": lagerorte or []}],
         "zeilen_ok": True, "summe_ok": True, "vollstaendig_ok": True}
    return regeln.bewerte_rechnung(r)


def selftest():
    import io
    import shutil
    import tempfile

    n_ok, n_fail = 0, []

    def check(name, ist, soll):
        nonlocal n_ok
        if ist == soll:
            n_ok += 1
        else:
            n_fail.append("%s: erwartet %r, war %r" % (name, soll, ist))

    b_dhl_de = _bsp("1705334", "pax1", 3.8, ["Karl-Heinz Kuril", "Stormstr. 98", "47445 Moers"])
    z = _dhl_dpd_zeile(b_dhl_de, False)
    check("DHL DE: Sendungsreferenz=Empfängerreferenz=Rnr",
          (z[0], z[23]), ("1705334", "1705334"))
    check("DHL DE: Produkt/Abrechnung", (z[24], z[25]), ("V01PAK", "52148008630101"))
    check("DHL DE: Land ISO3", z[19], "DEU")
    check("DHL DE: Telefon-Feld = Leerzeichen+Rnr", z[21], " 1705334")
    check("DHL DE: Name ungetrennt, ohne Zusatz Name 2/3 leer", (z[12], z[13], z[14]),
          ("Karl-Heinz Kuril", "", ""))
    check("DHL DE: Gewicht Komma-Format", z[22], "3,8000")
    b_z1 = _bsp("1700010", "pax1", 2.0, ["Max Muster", "Muster GmbH", "Weg 1", "12345 Ort"])
    z = _dhl_dpd_zeile(b_z1, False)
    check("DHL: 1 Zusatzzeile -> Name 2", (z[12], z[13], z[14]),
          ("Max Muster", "Muster GmbH", ""))
    b_z3 = _bsp("1700011", "wapo", 1.0,
                ["Max Muster", "Muster GmbH", "c/o Lager", "Tor 3", "Weg 1", "12345 Ort"])
    z = _dhl_dpd_zeile(b_z3, True)
    check("DPD: 3 Zusatzzeilen -> Name 2 + Name 3 zusammengefasst", (z[13], z[14]),
          ("Muster GmbH", "c/o Lager, Tor 3"))
    b_ot = _bsp("1705888", "pax1", 5.2, ["Scheffert Kornelia", "Dorf 17 OT Quitzerow",
                                         "17111 Kletzin"])
    z = _dhl_dpd_zeile(b_ot, False)
    check("DHL: Ortsteil ohne Zusatz -> Name 3, Ort sauber",
          (z[13], z[14], z[15], z[16], z[18]), ("", "OT Quitzerow", "Dorf", "17", "Kletzin"))
    b_ot2 = _bsp("1700012", "pax1", 2.0, ["A B", "Firma X", "Dorf 17 OT Quitzerow",
                                          "17111 Kletzin"])
    z = _dhl_dpd_zeile(b_ot2, False)
    check("DHL: Zusatz + Ortsteil -> Name 2 / Name 3", (z[13], z[14]),
          ("Firma X", "OT Quitzerow"))
    check("DHL DE: 32 Spalten", len(z), 32)

    # Manuell aufgeteilte Sendung (b["pakete"], siehe carrier_dashboard.
    # paket_aufteilen()) - zwei Beispiele mit Einzelgewichten, die auch die
    # GUI-Obergrenze (<=31,5 kg PRO Paket) einhalten wuerden.
    b_dhl_geteilt2 = _bsp("1705500", "pax1", 50.0, ["A B", "Weg 1", "12345 Ort"])
    b_dhl_geteilt2["pakete"] = [25.0, 25.0]
    zeilen2 = _dhl_zeilen(b_dhl_geteilt2)
    check("Aufgeteilte Sendung (2 Pakete): 2 Zeilen", len(zeilen2), 2)
    check("Aufgeteilte Sendung (2 Pakete): beide gleiche Sendungsreferenz",
          (zeilen2[0][0], zeilen2[1][0]), ("1705500", "1705500"))
    check("Aufgeteilte Sendung (2 Pakete): Einzelgewichte",
          (zeilen2[0][22], zeilen2[1][22]), ("25,0000", "25,0000"))

    # Keine Obergrenze fuer die Paketanzahl - real beobachtet an Rechnung
    # 1705548 (70,4 kg): passt selbst in ZWEI Pakete a 31,5 kg nicht mehr
    # (2x31,5=63 < 70,4), braucht also mindestens DREI.
    b_dhl_geteilt3 = _bsp("1705548", "pax1", 70.4, ["Bahittin Doener", "Rohrwangstr.3", "73430 AALEN"])
    check("Sendung vor Aufteilung -> Fehler (ueber DHL-Maximalgewicht)",
          b_dhl_geteilt3["status"], "fehler")
    # Wie carrier_dashboard.paket_aufteilen(): Aufteilung eintragen, den
    # Gewichts-Fehler damit als geloest markieren.
    b_dhl_geteilt3["pakete"] = [25.0, 25.0, 20.4]
    b_dhl_geteilt3["fehler"] = [f for f in b_dhl_geteilt3["fehler"] if "DHL-Maximalgewicht" not in f]
    b_dhl_geteilt3["status"] = "ok" if not b_dhl_geteilt3["fehler"] else "fehler"
    zeilen3 = _dhl_zeilen(b_dhl_geteilt3)
    check("Aufgeteilte Sendung (3 Pakete, 70,4 kg passt in keine 2): 3 Zeilen", len(zeilen3), 3)
    check("Aufgeteilte Sendung (3 Pakete): alle gleiche Sendungsreferenz",
          [z[0] for z in zeilen3], ["1705548"] * 3)
    check("Aufgeteilte Sendung (3 Pakete): Einzelgewichte, keins ueber 31,5 kg",
          [z[22] for z in zeilen3], ["25,0000", "25,0000", "20,4000"])

    check("Unaufgeteilte Sendung: weiterhin genau 1 Zeile", len(_dhl_zeilen(b_dhl_de)), 1)

    b_kp = _bsp("1700991", "wapo", 0.8, ["A B", "Weg 1", "AT-8330 Feldbach"])
    b_kp["kleinpaket"] = True
    z_kp = _dhl_dpd_zeile(b_kp, False)
    check("DHL Kleinpaket Ausland (= Warenpost International): Produkt/Abrechnung",
          (z_kp[24], z_kp[25]), ("V66WPI", "52148008636601"))
    check("DHL Kleinpaket Ausland: Land AUT", z_kp[19], "AUT")
    check("DHL ohne Kleinpaket-Flag: unveraendert Auslandsprodukt",
          (_dhl_dpd_zeile(dict(b_kp, kleinpaket=False), False)[24],), ("V53WPAK.V53VV",))

    b_dhl_at = _bsp("1703056", "pax1", 2.4, ["Michael Höfler", "Mühldorf 414", "AT-8330 Feldbach"])
    z = _dhl_dpd_zeile(b_dhl_at, False)
    check("DHL AT: Produkt/Abrechnung", (z[24], z[25]), ("V53WPAK.V53VV", "52148008635303"))
    check("DHL AT: Land ISO3", z[19], "AUT")

    b_dhl_fr = _bsp("1700001", "pax1", 1.0, ["Jan Dubois", "Rue Haute 5", "1000 Bruxelles", "BELGIEN"])
    z = _dhl_dpd_zeile(b_dhl_fr, False)
    check("DHL BE (unbelegtes Land): Default wie AT", (z[24], z[25]),
          ("V53WPAK.V53VV", "52148008635303"))

    b_dpd = _bsp("1705331", "wapo", 1.0, ["Nomi Atelier", "Ebbelicher Weg 70", "45701 Herten"],
                 kdnr="674982")
    z = _dhl_dpd_zeile(b_dpd, True)
    check("DPD: Empfängerreferenz=Kundennummer", z[23], "674982")
    check("DPD: Telefon-Feld leer", z[21], "")
    check("DPD: Produkt/Abrechnung", (z[24], z[25]), ("V01PAK", "52148008630101"))
    check("DPD: feste Nachnahme-Werte", tuple(z[26:32]), ("", "34", "23", "5", "16", ""))

    b_dpd_ohne_kdnr = _bsp("1705332", "wapo", 1.0, ["X Y", "Weg 1", "12345 Ort"])
    z = _dhl_dpd_zeile(b_dpd_ohne_kdnr, True)
    check("DPD ohne Kundennummer: Feld bleibt leer", z[23], "")

    b_brief = _bsp("1705321", "brx1", 0.03,
                    ["Petar Bekavac", "Hr. Bekavac / Produktion", "Balinger Straße 23",
                     "72415 Grosselfingen"])
    check("Brief: Ausland-Flag", b_brief["ausland"], False)
    zp = _post_zeile(b_brief["adresse"], b_brief["rnr"])
    check("Post: NAME/ZUSATZ getrennt", (zp[0], zp[1]), ("Petar Bekavac", "Hr. Bekavac / Produktion"))
    check("Post: Straße/Hausnr/PLZ/Ort/Land/Typ/Referenz", tuple(zp[2:]),
          ("Balinger Straße", "23", "72415", "Grosselfingen", "DE", "HOUSE", "1705321"))
    check("Post-Absenderzeile", _post_absenderzeile(),
          ["Gasecenter Augsburg", "", "Karlsruher Str.", "3", "86156", "Augsburg", "DE",
           "HOUSE", ""])

    b_grossbrief_aus = _bsp("1705399", "pox1", 0.4,
                             ["Jan Peeters", "Rue Haute 5", "1000 Bruxelles", "BELGIEN"])
    check("Großbrief Ausland: Flag", b_grossbrief_aus["ausland"], True)

    # Fehlerhafte Rechnung (keine PLZ) darf NICHT exportiert werden
    b_fehler = _bsp("1700099", "pax1", 1.0, ["X Y", "Hauptstr. 1"])
    check("Fehlerhafte Rechnung -> status fehler", b_fehler["status"], "fehler")

    # status "warn" (offener, NICHT quittierter Hinweis) blockiert den Export
    # GENAUSO wie "fehler" - b_dpd_ohne_kdnr hat wegen der fehlenden
    # Kundennummer status "warn" (siehe check oben) und darf hier NICHT
    # mitgezaehlt werden.
    check("DPD ohne Kundennummer -> status warn (offener Hinweis)",
          b_dpd_ohne_kdnr["status"], "warn")

    # --- exportiere(): echte Dateien in ein Temp-Verzeichnis schreiben ------
    tmp = tempfile.mkdtemp(prefix="carrier_export_test_")
    try:
        alle = [b_dhl_de, b_dhl_at, b_dpd, b_dpd_ohne_kdnr, b_brief, b_grossbrief_aus, b_fehler]
        jetzt = datetime(2026, 9, 22, 16, 10)
        geschrieben = exportiere(alle, tmp, jetzt)
        erwartet_dateien = {
            os.path.join(tmp, "DHL_2026-09-22_161000.csv"): 2,
            os.path.join(tmp, "DPD_2026-09-22_161000.csv"): 1,
            os.path.join(tmp, "Post_Brief_Inland_2026-09-22_161000.csv"): 1,
            os.path.join(tmp, "Post_Grossbrief_Ausland_2026-09-22_161000.csv"): 1,
        }
        check("exportiere(): erzeugte Dateien+Zeilenzahl (warn-Rechnung fehlt, siehe oben)",
              geschrieben, erwartet_dateien)
        with open(os.path.join(tmp, "DPD_2026-09-22_161000.csv"), encoding="iso-8859-1") as f:
            check("exportiere(): warn-Rechnung (offener Hinweis) nicht exportiert",
                  "1705332" not in f.read(), True)

        # Quittieren (Dashboard setzt status manuell auf "ok", siehe
        # carrier_dashboard.quittiere_auswahl()) macht die Rechnung
        # exportierbar - exportiere() selbst kennt den Begriff "quittiert"
        # nicht, prueft nur status.
        b_dpd_quittiert = dict(b_dpd_ohne_kdnr, status="ok")
        geschrieben_quittiert = exportiere([b_dpd_quittiert], tmp, datetime(2026, 9, 22, 16, 11))
        check("exportiere(): nach 'Quittieren' (status=ok) doch exportiert",
              len(geschrieben_quittiert), 1)

        # exportiere() mit einer in DREI Pakete aufgeteilten Sendung: DHL-Datei
        # bekommt DREI Datenzeilen fuer die eine Rechnung, "anzahl_rechnungen"
        # bleibt 1.
        geschrieben_geteilt = exportiere([b_dhl_geteilt3], tmp, datetime(2026, 9, 22, 16, 12))
        check("exportiere(): aufgeteilte Sendung zaehlt als 1 Rechnung",
              list(geschrieben_geteilt.values()), [1])
        [pfad_geteilt] = geschrieben_geteilt
        with open(pfad_geteilt, encoding="iso-8859-1", newline="") as f:
            rows_geteilt = list(csv.reader(f, delimiter=";"))
        check("exportiere(): aufgeteilte Sendung schreibt 3 Datenzeilen", len(rows_geteilt), 4)
        check("exportiere(): alle drei Zeilen dieselbe Rechnungsnummer",
              [r[0] for r in rows_geteilt[1:]], ["1705548"] * 3)
        check("exportiere(): keine Post_Brief_Ausland-Datei (leere Gruppe)",
              os.path.exists(os.path.join(tmp, "Post_Brief_Ausland_2026-09-22_161000.csv")), False)
        with open(os.path.join(tmp, "DHL_2026-09-22_161000.csv"), encoding="iso-8859-1") as f:
            check("exportiere(): fehlerhafte Rechnung nicht exportiert",
                  "1700099" not in f.read(), True)
        with open(os.path.join(tmp, "DHL_2026-09-22_161000.csv"), encoding="iso-8859-1", newline="") as f:
            rows = list(csv.reader(f, delimiter=";"))
        check("DHL-Datei: Header + 2 Datenzeilen", len(rows), 3)
        check("DHL-Datei: Header exakt", rows[0], DHL_DPD_SPALTEN)
        with open(os.path.join(tmp, "Post_Grossbrief_Ausland_2026-09-22_161000.csv"),
                  encoding="iso-8859-1", newline="") as f:
            rows = list(csv.reader(f, delimiter=";"))
        check("Post-Datei: Absenderzeile zuerst", rows[1][0], "Gasecenter Augsburg")
        check("Post-Datei: 1 Absender + 1 Rechnung", len(rows), 3)

        # dateiname(): zwei Aufrufe mit demselben Zeitstempel (zwei Laeufe binnen
        # derselben Sekunde) duerfen NIE denselben Pfad liefern - siehe Kollisions-
        # Suffix in dateiname(). Vorher (nur HHMM) waere das ueberschrieben worden.
        tmp2 = os.path.join(tmp, "kollision")
        os.makedirs(tmp2)
        p1 = dateiname(tmp2, "DHL", jetzt)
        open(p1, "w").close()
        p2 = dateiname(tmp2, "DHL", jetzt)
        check("dateiname(): Kollision bekommt eigenen Pfad", p1 != p2, True)
        check("dateiname(): Kollisions-Suffix", os.path.basename(p2), "DHL_2026-09-22_161000_2.csv")
        check("dateiname(): ext-Parameter (z.B. fuer Pickliste-PDF)",
              os.path.basename(dateiname(tmp2, "Pickliste", jetzt, ext="pdf")),
              "Pickliste_2026-09-22_161000.pdf")

        # _post_zeile(): Zusatz-Trenner darf KEIN Semikolon enthalten (Finding:
        # "; " zerriss die Post-CSV-Zeile genau wie das &#34;-Beispiel oben).
        a_zwei_zusatz = regeln.analysiere_adresse(
            ["Firma AB", "c/o Mueller", "Zweigstelle Nord", "Hauptstr. 1", "12345 Ort"])
        zp2 = _post_zeile(a_zwei_zusatz, "1700100")
        check("Post-Zusatz: kein Semikolon im Feld", ";" in zp2[1], False)
        check("Post-Zusatz: mit Komma verbunden", zp2[1], "c/o Mueller, Zweigstelle Nord")
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
