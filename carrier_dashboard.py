# -*- coding: utf-8 -*-
"""
carrier_dashboard.py - Carrier-Dashboard
============================================================================
Liest alle Rechnungs-PDFs eines Pool-Ordners, ordnet jede Bestellung einem
Versanddienstleister zu (DHL / DPD / Deutsche Post Brief / Grossbrief, je
Inland und Ausland) und zeigt das Ergebnis samt allen Fehlern an. Erst nach
einem zweiten, manuellen Schritt werden Pickliste, Bruecken-CSVs und
Carrier-CSVs tatsaechlich geschrieben und die eingelesenen PDFs archiviert.

  Schritt 1 "1. Bestellungen zuordnen": nur LESEN + ZUORDNEN, zeigt das
      Ergebnis inkl. aller Fehler/Hinweise. Es wird NICHTS geschrieben oder
      verschoben - die PDFs bleiben liegen, beliebig oft wiederholbar.
  Schritt 2 "2. Pickliste + CSV erstellen" (diese Version): baut aus dem
      Ergebnis von Schritt 1 die Pickliste (Layout unveraendert, wie
      packliste.baue_pdf) + die fuenf Bruecken-CSVs (wie packliste.main()),
      schreibt die Carrier-CSVs (carrier_export.exportiere) und archiviert
      die eingelesenen Rechnungs-PDFs (packliste.archiviere). NUR Rechnungen
      mit status == "ok" UND zugeordnetem Carrier landen in einer Carrier-CSV
      - Fehler UND noch nicht quittierte Hinweise blockieren das gleichermassen
      (siehe Button "Hinweis quittieren"); alle anderen werden trotzdem
      gepackt (Pickliste), aber NICHT automatisch exportiert (manuelle
      Nachbearbeitung noetig). Bei
      jedem Schritt-2-Lauf werden zusaetzlich automatisch eine Kg-Statistik je
      Carrier (carrier_statistik.log_lauf) UND eine Artikelanzahl-Statistik
      je Bestellung (carrier_statistik.log_artikel) mitgeschrieben - Anzeige
      ueber den Button "Statistik". Sicherung gegen Doppel-Verarbeitung:
      kommt eine Rechnungsnummer mehrfach im aktuellen Pool ODER laut
      Statistik-Historie schon in einem frueheren Lauf vor, verlangt Schritt 2
      eine EXPLIZITE Bestaetigung (Vorbelegung "Nein"), bevor irgendetwas
      gepackt/exportiert/archiviert wird. Filter: die Kacheln in der
      Zusammenfassung (DHL/DPD/.../Fehler/Hinweise) sind klickbar und
      schraenken die Tabelle ein ("Alle anzeigen" setzt zurueck). Ist ein
      Filter aktiv, verarbeitet Schritt 2 NUR die sichtbaren Zeilen (mit
      eigener Bestaetigung, Vorbelegung "Nein", da das die Ausnahme sein
      soll) - alles andere bleibt unveraendert in der Tabelle/im Pool stehen
      und kann spaeter in einem weiteren Lauf verarbeitet werden. Button
      "Hinweis quittieren": markiert die ausgewaehlte Zeile mit offenem
      Hinweis als geprueft (setzt status auf "ok") - sie zaehlt danach nicht
      mehr in der Hinweise-Kachel/im Hinweise-Filter (Meldungstext bleibt mit
      "[Quittiert]"-Praefix sichtbar) UND wird dadurch erst carrier-
      exportierbar (siehe carrier_export.py). Gilt nur fuer die laufende
      GUI-Sitzung, bleibt aber ueber einen erneuten Schritt-1-Lauf hinweg
      erhalten (rnr-basiert). Button "Paket aufteilen": bei einer DHL-
      Sendung ueber dem Maximalgewicht (siehe G_DHL_MAX in carrier_regeln.py)
      koennen beliebig viele (mindestens zwei) manuell gewogene Einzelgewichte
      eingetragen werden - loest den Gewichts-Fehler auf (macht daraus einen
      quittierbaren Hinweis) und erzeugt beim Export je EINE DHL-CSV-Zeile pro
      Paket mit dem gleichen Sendungsbezug (Matthias bestaetigt: alle Pakete
      tragen dieselbe Rechnungsnummer als Sendungsreferenz). Die Pickliste
      bleibt unveraendert (EIN Packvorgang) - die Pakete werden vorab manuell
      gepackt/gewogen, die Software muss nicht wissen, welcher Artikel in
      welches Paket kommt. Zusaetzlich AUTOMATISCH (kein Knopfdruck noetig):
      Artikel mit der Markierung "<N>-je-Paket" (packliste.JE_PAKET_RE) werden
      von carrier_regeln.je_paket_aufteilung() eigenstaendig auf mehrere
      DHL-Pakete verteilt (gleichmaessig, inkl. Fach-Artikel-Faktor), sofern
      die Rechnung ausschliesslich diesen einen Artikel enthaelt - siehe
      carrier_regeln.py Modulkopf.
  Schritt 4 (spaeter, optional): Warnung in scan_druck.py bei Carrier-
      Abweichung zwischen Label und dieser Zuordnung.

Die Regeln stehen in carrier_regeln.py, der CSV-Export in carrier_export.py,
die Kg-Statistik in carrier_statistik.py, das Auslesen/die Pickliste in
packliste.py - alle fuenf Dateien muessen im selben Ordner liegen.

Darkmode (auf Kundenwunsch, 2026-09-23): Fenster + Windows-Titelleiste dunkel
(angelehnt an die Farbpalette aus scan_druck.py fuer ein einheitliches Bild
zwischen Carrier-Dashboard und Versand-Cockpit), Ladebalken in Orange
(#F38808, aus Carrier-Dashboard.ico ausgezaehlt). Siehe _dunkles_theme()/
_dunkle_titelleiste() am Modulkopf.

Start:  py carrier_dashboard.py  (oder per Carrier-Dashboard_starten.vbs)

Die Ordner (Pool/Ausgabe/Archiv/Carrier-Export) sind bewusst FEST im Code
unten (POOL_ORDNER usw.) und NICHT im Dashboard waehlbar/aenderbar - sie
werden nicht von Lauf zu Lauf gewechselt. Zum Anpassen (z.B. beim Umzug auf
den Faktura-PC) einfach die Konstanten unten im Quelltext bearbeiten, genau
wie bei Pickliste_erstellen.bat.
"""

import glob
import os
import queue
import sys
import threading
import time
import tkinter as tk
from datetime import datetime
from tkinter import messagebox, simpledialog, ttk

import carrier_regeln as regeln

VERSION = "2026-09-25b"

# Fenster-/Taskleisten-Symbol (siehe gui() unten) - liegt im selben Ordner
# wie dieses Skript, damit es unveraendert auch nach einem Umzug funktioniert.
ICON_PFAD = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Carrier-Dashboard.ico")

# ----------------------------------------------------------------------------
# DARKMODE-Farben - angelehnt an die Palette aus scan_druck.py (BG/CARD/FG/
# MUTED/ROT), damit Carrier-Dashboard und Versand-Cockpit einheitlich
# aussehen. ORANGE ist die tatsaechliche Farbe aus Carrier-Dashboard.ico
# (per Pixel-Auszaehlung ermittelt, 2026-09-23), fuer den Ladebalken.
# ----------------------------------------------------------------------------
BG = "#1E1E24"
CARD = "#2A2A33"
FG = "#ECECEC"
MUTED = "#9AA0A6"
BORDER = "#3A3A45"
ROT = "#E53935"
ROT_ZEILE = "#4A2020"
GOLD = "#FFD54F"
GOLD_ZEILE = "#4A3A12"
ORANGE = "#F38808"


def _dunkle_titelleiste(root):
    """Dunkler Windows-Fenstertitel (Windows 10 1809+/11) - rein kosmetisch,
    schlaegt auf aelteren Windows-Versionen/anderen Betriebssystemen still
    fehl, das Fenster laeuft dann trotzdem (nur mit hellem Titelbalken)."""
    try:
        import ctypes
        hwnd = ctypes.windll.user32.GetParent(root.winfo_id())
        wert = ctypes.c_int(1)
        for attr in (20, 19):        # DWMWA_USE_IMMERSIVE_DARK_MODE: neu/alt
            if ctypes.windll.dwmapi.DwmSetWindowAttribute(
                    hwnd, attr, ctypes.byref(wert), ctypes.sizeof(wert)) == 0:
                break
    except Exception:
        pass


def _dunkles_theme(root):
    """ttk-Style auf Darkmode umstellen (Theme "clam", einziges eingebautes
    Theme, das Farb-Overrides auf Windows tatsaechlich anwendet - "vista"/
    "winnative" ignorieren die meisten davon). Rueckgabe: die Style-Instanz,
    falls weitere Styles gebraucht werden (siehe ORANGE-Ladebalken unten)."""
    style = ttk.Style(root)
    style.theme_use("clam")
    style.configure(".", background=BG, foreground=FG, fieldbackground=CARD,
                    bordercolor=BORDER, lightcolor=BG, darkcolor=BG)
    style.configure("TFrame", background=BG)
    style.configure("TLabel", background=BG, foreground=FG)
    style.configure("TLabelframe", background=BG, foreground=FG, bordercolor=BORDER)
    style.configure("TLabelframe.Label", background=BG, foreground=FG)
    style.configure("TButton", background=CARD, foreground=FG, bordercolor=BORDER,
                    focuscolor=BG)
    style.map("TButton", background=[("active", BORDER), ("disabled", BG)],
              foreground=[("disabled", MUTED)])
    style.configure("TPanedwindow", background=BG)
    style.configure("Treeview", background=CARD, foreground=FG, fieldbackground=CARD,
                    bordercolor=BORDER, rowheight=22)
    style.configure("Treeview.Heading", background=BG, foreground=FG, bordercolor=BORDER)
    style.map("Treeview.Heading", background=[("active", BORDER)])
    style.map("Treeview", background=[("selected", ORANGE)], foreground=[("selected", BG)])
    style.configure("TScrollbar", background=CARD, troughcolor=BG, bordercolor=BORDER,
                    arrowcolor=FG)
    style.configure("Orange.Horizontal.TProgressbar", troughcolor=CARD, background=ORANGE,
                    bordercolor=BORDER, lightcolor=ORANGE, darkcolor=ORANGE)
    return style

# ----------------------------------------------------------------------------
# ORDNER - hier fest eintragen, nicht im Dashboard waehlbar (die Ordner
# aendern sich nicht von Lauf zu Lauf). Beim Umzug auf den Faktura-PC einfach
# diese vier Zeilen anpassen.
# ----------------------------------------------------------------------------
POOL_ORDNER = r"C:\Carrier-Dashboard\Pool"
AUSGABE_ORDNER = r"C:\Carrier-Dashboard\Pool\Pickliste"
ARCHIV_ORDNER = r"C:\Carrier-Dashboard\Pool\Archiv"
CARRIER_EXPORT_ORDNER = r"C:\Carrier_Export"

REFRESH_MS = 5000

# Reihenfolge und Beschriftung der Zusammenfassung
GRUPPEN = [
    ("DHL", "DHL"),
    ("DPD", "DPD"),
    ("Post Großbrief|DE", "Großbrief Inland"),
    ("Post Großbrief|AUS", "Großbrief Ausland"),
    ("Post Brief|DE", "Brief Inland"),
    ("Post Brief|AUS", "Brief Ausland"),
]
# Zusaetzliche, nicht-carrier-spezifische Filter/Zusammenfassungs-Kacheln -
# zusammen mit GRUPPEN die Basis fuer die klickbaren Filter im Dashboard.
FILTER_TITEL = dict(GRUPPEN, FEHLER="Fehler", HINWEISE="Hinweise")


def gruppen_schluessel(b):
    """Zusammenfassungs-Schluessel eines Ergebnisses (None bei Fehler)."""
    c = b["carrier"]
    if c is None or b["status"] == "fehler":
        return None
    if c in (regeln.BRIEF, regeln.GROSSBRIEF):
        return "%s|%s" % (c, "AUS" if b["ausland"] else "DE")
    return c


def hat_offenen_hinweis(b):
    """True, wenn b einen Hinweis hat, der noch NICHT quittiert wurde (siehe
    Button 'Hinweis quittieren' im Dashboard). Ein quittierter Hinweis bleibt
    im Ergebnis-dict erhalten (Meldungstext/Detailansicht) - er zaehlt nur
    nicht mehr fuer die Hinweise-Kachel/den Hinweise-Filter."""
    return bool(b["hinweise"]) and not b.get("quittiert")


def braucht_paketaufteilung(b):
    """True, wenn b eine DHL-Sendung ueber dem Maximalgewicht ist (siehe
    regeln.G_DHL_MAX) - unabhaengig davon, ob schon eine Aufteilung
    (b["pakete"]) eingetragen wurde, damit der Button "Paket aufteilen" auch
    zum NACHTRAEGLICHEN Anpassen einer bereits eingetragenen Aufteilung
    nutzbar bleibt. b["gewicht"] ist immer das GESAMTgewicht laut Rechnung,
    das bleibt beim Aufteilen unveraendert (siehe paket_aufteilen())."""
    return b.get("carrier") == regeln.DHL and (b.get("gewicht") or 0) > regeln.G_DHL_MAX


def filter_treffer(schluessel, b):
    """True, wenn Ergebnis b zum Filter-Schluessel passt (fuer die klickbaren
    Kacheln in der Zusammenfassung). schluessel None = kein Filter, alles
    passt. 'HINWEISE' ist unabhaengig von 'FEHLER' (eine Rechnung kann beides
    gleichzeitig haben)."""
    if schluessel is None:
        return True
    if schluessel == "FEHLER":
        return b["status"] == "fehler"
    if schluessel == "HINWEISE":
        return hat_offenen_hinweis(b)
    return gruppen_schluessel(b) == schluessel


def lese_pool(ordner, melde=None):
    """Liest alle PDFs des Ordners und bewertet sie. Rueckgabe: (rechnungen_roh,
    ergebnisse) - zwei gleich lange, gemeinsam nach Rechnungsnummer sortierte
    Listen (Fehler ohne Nummer am Ende). rechnungen_roh[i] ist das komplette
    parse_pdf()-Ergebnis (inkl. "quelle" fuer die spaetere Archivierung) oder
    None, wenn diese Zeile nur ein Fehler ist (PDF nicht lesbar/keine
    Positionen) - dann ist ergebnisse[i] eine _fehlerzeile() ohne Carrier."""
    import packliste
    pdfs = sorted(glob.glob(os.path.join(ordner, "*.pdf")))
    treffer = []                      # Liste von (r_oder_None, b)
    for i, pfad in enumerate(pdfs, 1):
        name = os.path.basename(pfad)
        if melde:
            melde(i, len(pdfs), name)
        try:
            r = packliste.parse_pdf(pfad)
        except Exception as e:
            treffer.append((None, _fehlerzeile(name, "PDF nicht lesbar: %s" % e)))
            continue
        if not r.get("positionen"):
            treffer.append((None, _fehlerzeile(
                name, "Keine Positionen erkannt (keine Rechnung?)", r.get("rnr", ""))))
            continue
        r["datei"] = name
        r["quelle"] = pfad
        b = regeln.bewerte_rechnung(r)
        b["quelle"] = pfad
        treffer.append((r, b))
    treffer.sort(key=lambda t: (not t[1]["rnr"], t[1]["rnr"], t[1]["datei"]))
    ergebnisse = [t[1] for t in treffer]
    _markiere_pool_duplikate(ergebnisse)
    return [t[0] for t in treffer], ergebnisse


def _markiere_pool_duplikate(ergebnisse):
    """Haengt an jede Rechnung, deren Rechnungsnummer MEHRFACH (mit
    unterschiedlichen Dateien) im selben Pool vorkommt, einen Hinweis an -
    schon in der Schritt-1-Tabelle sichtbar, statt erst bei der Schritt-2-
    Bestaetigung zu ueberraschen (siehe dort: starte_export())."""
    import carrier_statistik
    doppelt = carrier_statistik.doppelte_im_lauf([b["rnr"] for b in ergebnisse])
    if not doppelt:
        return
    for b in ergebnisse:
        if b["rnr"] in doppelt:
            b["hinweise"].append(
                "ACHTUNG: Rechnungsnummer kommt %dx im Pool vor (mögliche "
                "Doppel-Verarbeitung)" % doppelt[b["rnr"]])
            if b["status"] == "ok":
                b["status"] = "warn"


def exportiere_alles(rechnungen, ergebnisse, ausgabe_pfad, archiv_ordner, carrier_ordner):
    """Schritt 2: rechnungen (gueltige parse_pdf()-Dicts, "quelle" gesetzt) UND
    ergebnisse (dazu bewertete Carrier-Ergebnisse, GLEICHE Reihenfolge/Laenge)
    -> Pickliste-PDF + fuenf Bruecken-CSVs (wie packliste.main(), Layout/Logik
    unveraendert) im Ordner von ausgabe_pfad, Carrier-CSVs nach carrier_ordner
    (carrier_export.exportiere - exportiert nur status == 'ok' MIT
    Carrier), danach Archivierung der eingelesenen PDFs nach archiv_ordner
    (packliste.archiviere - verschiebt nur, was erfolgreich verarbeitet wurde).
    Gibt einen Berichts-dict zurueck; einzelne Archiv-Fehler werfen KEINE
    Exception, sondern stehen im Bericht (siehe packliste.archiviere)."""
    import carrier_export
    import carrier_statistik
    import packliste

    out_dir = os.path.dirname(os.path.abspath(ausgabe_pfad))
    os.makedirs(out_dir, exist_ok=True)

    gruppen_roh = packliste.finde_sammelgruppen(rechnungen)
    heute_str = datetime.now().strftime("%d.%m.%Y")
    for g in gruppen_roh:
        g["datum"] = heute_str
    sammel_pfad = os.path.join(out_dir, "sammel_zuordnung.csv")
    gruppen, gruppen_fuer_csv = packliste.merge_sammelgruppen(gruppen_roh, sammel_pfad)
    packliste.baue_pdf(rechnungen, ausgabe_pfad, gruppen)

    packliste.schreibe_csv(rechnungen, os.path.join(out_dir, "post_zuordnung.csv"))
    packliste.schreibe_sammel_csv(gruppen_fuer_csv, sammel_pfad)
    packliste.schreibe_mengen_csv(rechnungen, os.path.join(out_dir, "mengen_zuordnung.csv"))
    packliste.schreibe_ean_csv(rechnungen, os.path.join(out_dir, "ean_zuordnung.csv"))
    wc_neu = packliste.schreibe_wc_bestellnummern_csv(
        rechnungen, os.path.join(out_dir, "wc_bestellnummern.csv"))

    carrier_dateien = carrier_export.exportiere(ergebnisse, carrier_ordner)
    # Kg-/Artikel-Statistik im Hintergrund mitschreiben (Nice-to-have, blockiert
    # bei Schreibfehlern - z.B. Netzlaufwerk kurz weg - NIE den eigentlichen
    # Export, siehe carrier_statistik.log_lauf()/log_artikel()). Kg nur fuer die
    # carrier-zugeordneten Rechnungen (dieselbe Basis wie die Carrier-CSVs),
    # Artikelanzahl fuer ALLE verarbeiteten Rechnungen (ein Adressfehler
    # aendert nichts an der bestellten Menge).
    kg_geloggt = carrier_statistik.log_lauf(ergebnisse)
    artikel_geloggt = carrier_statistik.log_artikel(rechnungen)

    verschoben, archiv_fehler, archiv_ziel = 0, [], None
    if archiv_ordner:
        verschoben, archiv_fehler, archiv_ziel = packliste.archiviere(rechnungen, archiv_ordner)

    return {
        "pickliste": ausgabe_pfad, "anzahl": len(rechnungen), "gruppen": gruppen,
        "wc_neu": wc_neu, "carrier_dateien": carrier_dateien, "kg_geloggt": kg_geloggt,
        "artikel_geloggt": artikel_geloggt,
        "archiviert": verschoben, "archiv_fehler": archiv_fehler, "archiv_ziel": archiv_ziel,
    }


def _fehlerzeile(datei, text, rnr=""):
    adr = regeln.analysiere_adresse([])
    return {"rnr": rnr, "datei": datei, "adresse": adr, "kennungen": [],
            "gewicht": None, "carrier": None, "ausland": False, "grund": "",
            "fehler": [text], "hinweise": [], "status": "fehler", "quelle": "",
            "pakete": None}


def zaehle(ergebnisse):
    """(z, fehler, hinweise) - z ist {gruppen_schluessel: anzahl}, fehler/
    hinweise sind Gesamtzahlen. Eine Rechnung kann in fehler UND hinweise
    gleichzeitig gezaehlt werden (siehe filter_treffer()); quittierte
    Hinweise (siehe hat_offenen_hinweis()) zaehlen nicht mehr mit."""
    z = {k: 0 for k, _ in GRUPPEN}
    fehler = hinweise = 0
    for b in ergebnisse:
        k = gruppen_schluessel(b)
        if k is None:
            fehler += 1
        else:
            z[k] = z.get(k, 0) + 1
        if hat_offenen_hinweis(b):
            hinweise += 1
    return z, fehler, hinweise


def detailtext(b):
    a = b["adresse"]
    zeilen = ["Rechnung:  %s   (%s)" % (b["rnr"] or "?", b["datei"])]
    zeilen.append("Kennung:   %s" % (", ".join(b["kennungen"]) or "-"))
    g = b["gewicht"]
    zeilen.append("Gewicht:   %s" % (("%.3f kg" % g).replace(".", ",") if g is not None else "-"))
    if b.get("pakete"):
        zeilen.append("Pakete:    %s (%d DHL-Sendungen, gleiche Rechnungsnummer als Referenz)"
                      % (" + ".join(("%.3f kg" % p).replace(".", ",") for p in b["pakete"]),
                         len(b["pakete"])))
    zeilen.append("Carrier:   %s%s" % (
        b["carrier"] or "-", "  (Ausland)" if b["carrier"] and b["ausland"] else ""))
    if b["grund"]:
        zeilen.append("Grund:     %s" % b["grund"])
    zeilen.append("")
    zeilen.append("Adresse (manuell bearbeitet):" if b.get("adresse_bearbeitet")
                  else "Adresse laut Auswertung:")
    zeilen.append("  Name:     %s" % (a["name"] or "-"))
    for z in a["zusatz"]:
        zeilen.append("  Zusatz:   %s" % z)
    zeilen.append("  Straße:   %s   Hausnr.: %s" % (a["strasse"] or "-", a["hausnr"] or "-"))
    zeilen.append("  PLZ/Ort:  %s %s   Land: %s" % (a["plz"], a["ort"], a["land"] or "?"))
    if b["fehler"]:
        zeilen.append("")
        zeilen.append("FEHLER (blockiert den Export):")
        zeilen += ["  - " + f for f in b["fehler"]]
    if b["hinweise"]:
        zeilen.append("")
        zeilen.append("Hinweise (quittiert - vom Nutzer bestätigt):"
                      if b.get("quittiert") else "Hinweise:")
        zeilen += ["  - " + h for h in b["hinweise"]]
    return "\n".join(zeilen)


def _frage_gewicht(root, titel, prompt, minvalue, maxvalue, initialvalue=None):
    """Fragt ein Gewicht in kg ab - anders als simpledialog.askfloat() (das
    NUR den englischen Punkt als Dezimaltrennzeichen versteht und bei einem
    deutschen Komma die kryptische Meldung "Not a floating point value"
    zeigt) akzeptiert dieser Dialog sowohl Komma als auch Punkt. Fragt bei
    ungueltiger Eingabe (kein Zahlenformat ODER ausserhalb min/max) mit einer
    klaren deutschen Fehlermeldung erneut, bis eine gueltige Zahl eingegeben
    oder abgebrochen wird (dann None)."""
    vorgabe = ("%.3f" % initialvalue).replace(".", ",") if initialvalue is not None else None
    while True:
        text = simpledialog.askstring(titel, prompt, parent=root, initialvalue=vorgabe)
        if text is None:
            return None
        text_bereinigt = text.strip().replace(",", ".")
        try:
            wert = float(text_bereinigt)
            if wert != wert:                    # "nan" passiert sonst beide Grenzvergleiche
                raise ValueError
        except ValueError:
            messagebox.showerror(titel,
                                 "\"%s\" ist keine gültige Zahl. Bitte z.B. 25,5 oder "
                                 "25.5 eingeben." % text)
            vorgabe = text
            continue
        if wert < minvalue or wert > maxvalue:
            messagebox.showerror(
                titel, "Wert muss zwischen %s und %s kg liegen."
                % (("%.3f" % minvalue).replace(".", ","), ("%.1f" % maxvalue).replace(".", ",")))
            vorgabe = text
            continue
        return wert


# ==========================================================================
# GUI
# ==========================================================================

def _adress_dialog(root, a, rnr):
    """Modaler Dialog 'Adresse bearbeiten' - Felder mit den aktuell ausgewerteten
    Werten vorbelegt. Rueckgabe: dict der Felder (name, zusatz[Liste], strasse,
    hausnr, ortsteil, plz, ort, land) oder None bei Abbruch."""
    dlg = tk.Toplevel(root)
    dlg.title("Adresse bearbeiten - Rechnung %s" % (rnr or "?"))
    dlg.configure(bg=BG)
    dlg.transient(root)
    dlg.resizable(False, False)
    _dunkle_titelleiste(dlg)
    zusatz = list(a.get("zusatz") or [])
    felder = [
        ("name", "Name", a.get("name", "")),
        ("z1", "Zusatz 1 (Firma/c/o)", zusatz[0] if len(zusatz) > 0 else ""),
        ("z2", "Zusatz 2", zusatz[1] if len(zusatz) > 1 else ""),
        ("strasse", "Straße", a.get("strasse", "")),
        ("hausnr", "Hausnummer", a.get("hausnr", "")),
        ("ortsteil", "Ortsteil (optional)", a.get("ortsteil", "")),
        ("plz", "PLZ", a.get("plz", "")),
        ("ort", "Ort", a.get("ort", "")),
        ("land", "Land (2 Buchstaben, z.B. DE/AT)", a.get("land", "")),
    ]
    if len(zusatz) > 2:                       # weitere Zusatzzeilen nicht verlieren
        felder[2] = ("z2", "Zusatz 2", " | ".join(zusatz[1:]))
    vars_ = {}
    for i, (key, label, wert) in enumerate(felder):
        ttk.Label(dlg, text=label).grid(row=i, column=0, sticky="w", padx=(12, 8), pady=3)
        v = tk.StringVar(value=wert)
        vars_[key] = v
        e = tk.Entry(dlg, textvariable=v, width=42, bg=CARD, fg=FG, insertbackground=FG,
                     relief="flat", highlightthickness=1, highlightbackground=BORDER,
                     highlightcolor=ORANGE)
        e.grid(row=i, column=1, padx=(0, 12), pady=3)
        if key == "hausnr":
            e.focus_set()
            e.select_range(0, "end")
    ergebnis = {"werte": None}

    def ok(_evt=None):
        w = {k: v.get().strip() for k, v in vars_.items()}
        ergebnis["werte"] = {
            "name": w["name"], "zusatz": [w["z1"], w["z2"]], "strasse": w["strasse"],
            "hausnr": w["hausnr"], "ortsteil": w["ortsteil"], "plz": w["plz"],
            "ort": w["ort"], "land": w["land"]}
        dlg.destroy()

    knoepfe_f = ttk.Frame(dlg)
    knoepfe_f.grid(row=len(felder), column=0, columnspan=2, pady=10)
    ttk.Button(knoepfe_f, text="Übernehmen", command=ok).pack(side="left", padx=6)
    ttk.Button(knoepfe_f, text="Abbrechen", command=dlg.destroy).pack(side="left", padx=6)
    dlg.bind("<Return>", ok)
    dlg.bind("<Escape>", lambda _e: dlg.destroy())
    dlg.grab_set()
    root.wait_window(dlg)
    return ergebnis["werte"]


def gui():
    root = tk.Tk()
    root.title("Carrier-Dashboard  (Version %s / Regeln %s)" % (VERSION, regeln.VERSION))
    root.geometry("1100x720")
    root.configure(bg=BG)
    _dunkle_titelleiste(root)
    _dunkles_theme(root)
    if os.path.exists(ICON_PFAD):
        try:
            root.iconbitmap(ICON_PFAD)
        except Exception:
            pass                      # z.B. falsches Format - Fenster laeuft trotzdem

    pool_anz = tk.StringVar(value="")
    status_var = tk.StringVar(value="Noch nicht zugeordnet.")
    ergebnisse = []
    rechnungen_roh = []
    q = queue.Queue()
    laeuft = {"an": False}
    export_info = {"uebersprungen": 0, "verarbeitete_indizes": set()}
    filter_state = {"schluessel": None}
    # rnr -> Hinweistexte, die der Nutzer per Button quittiert hat (nur genau
    # DIESE Texte gelten nach einem neuen Lauf weiter als quittiert; ein neu
    # hinzugekommener Hinweis macht die Zeile wieder offen) -
    # bleibt ueber einen erneuten Schritt-1-Lauf hinweg erhalten (wird nach
    # jedem Lese-Lauf erneut angewendet, siehe abfrage()), damit ein Nachlade-
    # Lauf mit neuen PDFs nicht bereits quittierte Hinweise wieder aufleben
    # laesst. Nur fuer die aktuelle GUI-Sitzung (kein Speichern auf Platte).
    quittiert = {}
    # rnr -> manuell korrigierte Adresszeilen (Button "Adresse bearbeiten"); wird
    # nach jedem Schritt-1-Lauf erneut angewendet, damit die Nachtragung (z.B.
    # Hausnummer nach Rueckfrage beim Kunden) nicht verloren geht.
    adress_korrekturen = {}

    # --- Kopf: Pool-Ordner (fest, nur zur Information) + Zaehler -------------
    # Die Ordner sind bewusst NICHT hier waehlbar, siehe Modul-Kopf/Konstanten
    # oben - der Platz bleibt frei fuer spaetere Erweiterungen.
    kopf = ttk.Frame(root, padding=8)
    kopf.pack(fill="x")
    ttk.Label(kopf, text="Pool-Ordner: %s" % POOL_ORDNER,
              foreground=MUTED).pack(anchor="w")

    zaehler = tk.Label(kopf, textvariable=pool_anz, font=("Segoe UI", 18, "bold"),
                       anchor="w", bg=BG, fg=FG)
    zaehler.pack(anchor="w", pady=(8, 0))

    def aktualisiere_zaehler():
        if os.path.isdir(POOL_ORDNER):
            n = len(glob.glob(os.path.join(POOL_ORDNER, "*.pdf")))
            pool_anz.set("Rechnungen im Pool: %d" % n)
        else:
            pool_anz.set("Pool-Ordner nicht gefunden: %s" % POOL_ORDNER)

    def tick():
        if not laeuft["an"]:
            aktualisiere_zaehler()
        root.after(REFRESH_MS, tick)

    # --- Buttons -------------------------------------------------------------
    knoepfe = ttk.Frame(root, padding=(8, 0))
    knoepfe.pack(fill="x")
    btn_zuordnen = ttk.Button(knoepfe, text="1. Bestellungen zuordnen")
    btn_zuordnen.pack(side="left")
    btn_export = ttk.Button(knoepfe, text="2. Pickliste + CSV erstellen", state="disabled")
    btn_export.pack(side="left", padx=8)

    wartet = {"an": False}

    def im_hintergrund(fn, timeout=20.0):
        """Fuehrt fn (z.B. Lesezugriff auf das Netzlaufwerk mit den Statistik-
        Dateien) in einem Thread aus, waehrend die GUI weiter Ereignisse
        verarbeitet - ein nicht erreichbarer UNC-Pfad kann sonst das ganze
        Fenster minutenlang einfrieren. Rueckgabe (True, wert) oder (False,
        None) bei Fehler/Zeitueberschreitung. Waehrenddessen sind die
        Hauptbuttons gesperrt (kein erneuter Klick moeglich)."""
        erg = {}

        def lauf():
            try:
                erg["wert"] = fn()
            except Exception as e:
                erg["fehler"] = e

        t = threading.Thread(target=lauf, daemon=True)
        t.start()
        fertig = tk.BooleanVar(value=False)
        start = time.monotonic()

        def poll():
            if not t.is_alive() or time.monotonic() - start > timeout:
                fertig.set(True)
            else:
                root.after(50, poll)

        alt = (str(btn_zuordnen["state"]), str(btn_export["state"]))
        alter_status = status_var.get()
        wartet["an"] = True
        btn_zuordnen.configure(state="disabled")
        btn_export.configure(state="disabled")
        status_var.set("Lese Statistik-Dateien ...")
        root.after(50, poll)
        root.wait_variable(fertig)
        wartet["an"] = False
        btn_zuordnen.configure(state=alt[0])
        btn_export.configure(state=alt[1])
        status_var.set(alter_status)
        if t.is_alive() or "fehler" in erg:
            return False, None
        return True, erg["wert"]

    def zeige_statistik():
        import carrier_statistik
        if wartet["an"]:
            return
        ok, text = im_hintergrund(carrier_statistik.statistik_text)
        if not ok:
            messagebox.showwarning("Carrier-Dashboard - Statistik",
                                   "Die Statistik-Dateien sind gerade nicht erreichbar "
                                   "(Netzlaufwerk?). Bitte spaeter erneut versuchen.")
            return
        messagebox.showinfo("Carrier-Dashboard - Kg-Statistik", text)

    ttk.Button(knoepfe, text="Statistik", command=zeige_statistik).pack(side="left", padx=(0, 8))

    def filter_zuruecksetzen():
        wende_filter(None)

    ttk.Button(knoepfe, text="Alle anzeigen", command=filter_zuruecksetzen).pack(
        side="left", padx=(0, 8))

    btn_quittieren = ttk.Button(knoepfe, text="Hinweis quittieren", state="disabled")
    btn_quittieren.pack(side="left", padx=(0, 8))
    btn_aufteilen = ttk.Button(knoepfe, text="Paket aufteilen", state="disabled")
    btn_aufteilen.pack(side="left", padx=(0, 8))
    btn_adresse = ttk.Button(knoepfe, text="Adresse bearbeiten", state="disabled")
    btn_adresse.pack(side="left", padx=(0, 8))
    ttk.Label(knoepfe, textvariable=status_var).pack(side="left", padx=12)
    prog = ttk.Progressbar(root, mode="determinate", style="Orange.Horizontal.TProgressbar")
    prog.pack(fill="x", padx=8, pady=(6, 0))

    # --- Zusammenfassung (Kacheln sind klickbar -> Filter, siehe fuelle()) --
    zf = ttk.LabelFrame(root, text="Zuordnung", padding=6)
    zf.pack(fill="x", padx=8, pady=8)
    zlabels = {}
    for i, (k, titel) in enumerate(GRUPPEN + [("FEHLER", "Fehler"), ("HINWEISE", "Hinweise")]):
        titel_lbl = ttk.Label(zf, text=titel, cursor="hand2")
        titel_lbl.grid(row=0, column=i, padx=10)
        lb = tk.Label(zf, text="-", font=("Segoe UI", 16, "bold"), cursor="hand2",
                      bg=BG, fg=FG)
        lb.grid(row=1, column=i, padx=10)
        for w in (titel_lbl, lb):
            w.bind("<Button-1>", lambda _evt, s=k: wende_filter(s))
        zlabels[k] = lb
        zf.columnconfigure(i, weight=1)

    # --- Tabelle + Detail ---------------------------------------------------
    pan = ttk.PanedWindow(root, orient="vertical")
    pan.pack(fill="both", expand=True, padx=8, pady=(0, 8))
    spalten = ("rnr", "carrier", "land", "gewicht", "kennung", "status", "meldung")
    tv = ttk.Treeview(pan, columns=spalten, show="headings", selectmode="browse")
    for sp, titel, br in (("rnr", "Rechnung", 90), ("carrier", "Carrier", 130),
                          ("land", "Land", 50), ("gewicht", "kg", 70),
                          ("kennung", "Kennung", 90), ("status", "Status", 70),
                          ("meldung", "Grund / Meldung", 560)):
        tv.heading(sp, text=titel)
        tv.column(sp, width=br, anchor="w", stretch=(sp == "meldung"))
    tv.tag_configure("fehler", background=ROT_ZEILE, foreground=FG)
    tv.tag_configure("warn", background=GOLD_ZEILE, foreground=FG)
    tv.tag_configure("ok", background=CARD, foreground=FG)
    sb = ttk.Scrollbar(tv, orient="vertical", command=tv.yview)
    tv.configure(yscrollcommand=sb.set)
    sb.pack(side="right", fill="y")
    pan.add(tv, weight=3)
    detail = tk.Text(pan, height=12, font=("Consolas", 10), wrap="word", state="disabled",
                     bg=CARD, fg=FG, insertbackground=FG, relief="flat",
                     selectbackground=ORANGE, selectforeground=BG)
    pan.add(detail, weight=1)

    def zeige_detail(_evt=None):
        sel = tv.selection()
        if not sel:
            btn_quittieren.configure(state="disabled")
            btn_adresse.configure(state="disabled")
            return
        idx = int(sel[0])
        b = ergebnisse[idx]
        btn_adresse.configure(state=("normal" if rechnungen_roh[idx] is not None else "disabled"))
        detail.configure(state="normal")
        detail.delete("1.0", "end")
        detail.insert("1.0", detailtext(b))
        detail.configure(state="disabled")
        # Nur quittierbar, wenn die Zeile gerade einen offenen (noch nicht
        # quittierten) Hinweis hat - ein Fehler blockiert weiterhin den
        # Export und wird hier bewusst NICHT wegquittierbar gemacht.
        btn_quittieren.configure(state=("normal" if hat_offenen_hinweis(b) else "disabled"))
        btn_aufteilen.configure(state=("normal" if braucht_paketaufteilung(b) else "disabled"))

    tv.bind("<<TreeviewSelect>>", zeige_detail)

    def quittiere_auswahl():
        sel = tv.selection()
        if not sel:
            return
        idx = int(sel[0])
        b = ergebnisse[idx]
        if not hat_offenen_hinweis(b):
            return
        b["quittiert"] = True
        if b["status"] == "warn":
            b["status"] = "ok"
        if b["rnr"]:
            quittiert[b["rnr"]] = frozenset(b["hinweise"])
        fuelle()

    btn_quittieren.configure(command=quittiere_auswahl)

    def bewerte_mit_korrektur(r, b_alt):
        """Bewertet r mit der gemerkten Adresskorrektur neu (gleiche Validierung
        wie im Original) und uebernimmt die Pool-Duplikat-Hinweise aus b_alt."""
        r["adresse_zeilen"] = adress_korrekturen[r["rnr"]]
        b = regeln.bewerte_rechnung(r)
        b["quelle"] = r.get("quelle", "")
        b["adresse_bearbeitet"] = True
        dup = [h for h in b_alt["hinweise"] if h.startswith("ACHTUNG: Rechnungsnummer kommt")]
        if dup:
            b["hinweise"] += dup
            if b["status"] == "ok":
                b["status"] = "warn"
        return b

    def adresse_bearbeiten():
        sel = tv.selection()
        if not sel:
            return
        idx = int(sel[0])
        r, b = rechnungen_roh[idx], ergebnisse[idx]
        if r is None:
            return
        if not r.get("rnr"):
            messagebox.showinfo("Carrier-Dashboard", "Ohne Rechnungsnummer nicht bearbeitbar.")
            return
        w = _adress_dialog(root, b["adresse"], r["rnr"])
        if w is None:
            return
        adress_korrekturen[r["rnr"]] = regeln.adresse_zeilen_aus_feldern(
            w["name"], w["zusatz"], w["strasse"], w["hausnr"], w["ortsteil"],
            w["plz"], w["ort"], w["land"])
        ergebnisse[idx] = bewerte_mit_korrektur(r, b)
        quittiert.pop(r["rnr"], None)                 # neue Adresse -> Hinweise erneut pruefen
        fuelle()

    btn_adresse.configure(command=adresse_bearbeiten)

    def paket_aufteilen():
        sel = tv.selection()
        if not sel:
            return
        idx = int(sel[0])
        b = ergebnisse[idx]
        if not braucht_paketaufteilung(b):
            return
        gesamt = b["gewicht"]
        gtxt = ("%.3f" % gesamt).replace(".", ",")
        maxtxt = ("%.1f" % regeln.G_DHL_MAX).replace(".", ",")
        vorgabe = b.get("pakete") or []
        # Kein fest verdrahtetes Limit auf zwei Pakete - carrier_export.
        # _dhl_zeilen()/carrier_statistik.log_lauf() verarbeiten b["pakete"]
        # ohnehin als beliebig lange Liste (z.B. bei sehr schweren Sendungen,
        # die selbst in zwei Pakete a 31,5 kg nicht mehr passen, real
        # beobachtet an Rechnung 1705548 mit 70,4 kg -> 2x31,5=63 reicht nicht).
        # Mindestens zwei Pakete werden immer verlangt (ein einzelnes Paket
        # waere keine Aufteilung).
        pakete = []
        n = 1
        while True:
            if n == 1:
                prompt = ("Rechnung %s: Gesamtgewicht laut Rechnung %s kg "
                          "(DHL-Maximalgewicht %s kg PRO Paket).\n\nAlle Pakete "
                          "vorab packen und wiegen, dann hier eintragen.\n\n"
                          "Gewicht Paket 1 (kg):" % (b["rnr"], gtxt, maxtxt))
            else:
                prompt = "Gewicht Paket %d (kg):" % n
            vorgabe_n = vorgabe[n - 1] if len(vorgabe) >= n else None
            g = _frage_gewicht(
                root, "Carrier-Dashboard - Paket aufteilen", prompt,
                minvalue=0.001, maxvalue=regeln.G_DHL_MAX, initialvalue=vorgabe_n)
            if g is None:
                return                          # Abbruch - nichts wird uebernommen
            pakete.append(g)
            summe = sum(pakete)
            if n == 1:
                n += 1
                continue                        # immer mindestens 2 Pakete verlangen
            reicht_aus = summe >= gesamt - 0.001
            weiter = messagebox.askyesno(
                "Carrier-Dashboard - Paket aufteilen",
                "Bisher %d Pakete erfasst, Summe %s kg von %s kg laut Rechnung.\n\n"
                "Noch ein weiteres Paket hinzufügen?"
                % (n, ("%.3f" % summe).replace(".", ","), gtxt),
                default=(messagebox.NO if reicht_aus else messagebox.YES))
            if not weiter:
                break
            n += 1

        summe = sum(pakete)
        summe_txt = ("%.3f" % summe).replace(".", ",")
        # Grobe Plausibilitaetspruefung (kein hartes Blockieren - laut Matthias
        # ist eine exakt gleiche/passende Aufteilung nicht immer moeglich,
        # z.B. wegen Verpackungsmaterial) - faengt aber Tippfehler ab (z.B.
        # Gramm statt Kilogramm eingetragen).
        abweichung = abs(summe - gesamt)
        if abweichung > max(2.0, gesamt * 0.1):
            if not messagebox.askyesno(
                "Carrier-Dashboard - Gewicht prüfen",
                "Summe der %d Pakete (%s kg) weicht deutlich vom Rechnungsgewicht "
                "(%s kg) ab. Trotzdem übernehmen?" % (len(pakete), summe_txt, gtxt),
                icon="warning", default=messagebox.NO):
                return
        b["pakete"] = pakete
        b["fehler"] = [f for f in b["fehler"] if "DHL-Maximalgewicht" not in f]
        b["hinweise"] = [h for h in b["hinweise"] if not h.startswith("Sendung manuell in")]
        pakete_txt = " + ".join(("%.3f" % g).replace(".", ",") for g in pakete)
        b["hinweise"].append(
            "Sendung manuell in %d Pakete aufgeteilt: %s kg (Summe %s kg, "
            "Rechnung: %s kg)" % (len(pakete), pakete_txt, summe_txt, gtxt))
        b["quittiert"] = False                  # neue Aufteilung -> erneut quittieren
        quittiert.pop(b["rnr"], None)
        b["status"] = "fehler" if b["fehler"] else ("warn" if b["hinweise"] else "ok")
        fuelle()

    btn_aufteilen.configure(command=paket_aufteilen)

    def wende_filter(schluessel):
        filter_state["schluessel"] = schluessel
        fuelle()

    def fuelle():
        tv.delete(*tv.get_children())
        btn_quittieren.configure(state="disabled")     # Auswahl ist durch delete() weg
        btn_aufteilen.configure(state="disabled")
        schluessel = filter_state["schluessel"]
        angezeigt = 0
        for i, b in enumerate(ergebnisse):
            if not filter_treffer(schluessel, b):
                continue
            angezeigt += 1
            if b["fehler"]:
                meldung = "; ".join(b["fehler"])
            elif b["hinweise"]:
                meldung = b["grund"] + "  |  " + "; ".join(b["hinweise"])
                if b.get("quittiert"):
                    meldung = "[Quittiert] " + meldung
            else:
                meldung = b["grund"]
            g = b["gewicht"]
            gtxt = ("%.3f" % g).replace(".", ",") if g is not None else "-"
            if b.get("pakete"):
                gtxt += " (%d Pakete)" % len(b["pakete"])
            c = b["carrier"] or "-"
            if b["carrier"] in (regeln.BRIEF, regeln.GROSSBRIEF):
                c += " Ausland" if b["ausland"] else " Inland"
            tv.insert("", "end", iid=str(i), tags=(b["status"],), values=(
                b["rnr"] or "?", c, b["adresse"]["land"] or "?", gtxt,
                ",".join(b["kennungen"]) or "-",
                {"ok": "OK", "warn": "Hinweis", "fehler": "FEHLER"}[b["status"]], meldung))
        z, fehler, hinweise = zaehle(ergebnisse)
        for k, _ in GRUPPEN:
            zlabels[k].configure(text=str(z.get(k, 0)), fg=FG)
        zlabels["FEHLER"].configure(text=str(fehler), fg=(ROT if fehler else FG))
        zlabels["HINWEISE"].configure(text=str(hinweise), fg=(GOLD if hinweise else FG))
        if schluessel is None:
            zf.configure(text="Zuordnung")
        else:
            zf.configure(text="Zuordnung  -  Filter: %s (%d von %d angezeigt) - "
                              "\"Alle anzeigen\" setzt zurueck"
                         % (FILTER_TITEL.get(schluessel, schluessel), angezeigt, len(ergebnisse)))

    # --- Schritt 1 im Hintergrund --------------------------------------------
    def arbeite(ordner):
        try:
            roh, erg = lese_pool(ordner, lambda i, n, name: q.put(("fortschritt", i, n, name)))
            q.put(("fertig", roh, erg))
        except Exception as e:                       # z.B. reportlab/pdfplumber fehlt
            q.put(("abbruch", "%s: %s" % (type(e).__name__, e)))

    def starte():
        if laeuft["an"]:
            return
        if not os.path.isdir(POOL_ORDNER):
            status_var.set("Pool-Ordner nicht gefunden: %s" % POOL_ORDNER)
            return
        laeuft["an"] = True
        btn_zuordnen.configure(state="disabled")
        btn_export.configure(state="disabled")
        prog.configure(value=0)
        status_var.set("Lese Rechnungen ...")
        threading.Thread(target=arbeite, args=(POOL_ORDNER,), daemon=True).start()
        root.after(100, abfrage)

    def abfrage():
        try:
            while True:
                m = q.get_nowait()
                if m[0] == "fortschritt":
                    _, i, n, name = m
                    prog.configure(maximum=max(n, 1), value=i)
                    status_var.set("Lese %d/%d: %s" % (i, n, name))
                elif m[0] == "fertig":
                    rechnungen_roh[:] = m[1]
                    ergebnisse[:] = m[2]
                    for i, (r, b) in enumerate(zip(rechnungen_roh, ergebnisse)):
                        if r is not None and r.get("rnr") in adress_korrekturen:
                            ergebnisse[i] = bewerte_mit_korrektur(r, b)
                    # Frueher (in dieser Sitzung) quittierte Hinweise wieder
                    # anwenden, falls die betroffene Rechnung erneut auftaucht
                    # (z.B. weil zwischenzeitlich neue PDFs dazukamen und
                    # Schritt 1 nochmal gelaufen ist) - sonst muesste man
                    # dieselbe Rechnung jedes Mal neu quittieren.
                    for b in ergebnisse:
                        if b["hinweise"] and set(b["hinweise"]) <= quittiert.get(b["rnr"], frozenset()):
                            b["quittiert"] = True
                            if b["status"] == "warn":
                                b["status"] = "ok"
                    filter_state["schluessel"] = None     # frischer Lauf -> kein alter Filter
                    fuelle()
                    laeuft["an"] = False
                    btn_zuordnen.configure(state="normal")
                    btn_export.configure(state=("normal" if ergebnisse else "disabled"))
                    z, fehler, _ = zaehle(ergebnisse)
                    status_var.set("Fertig: %d Rechnung(en), %d Fehler. Es wurde nichts "
                                   "geschrieben oder verschoben." % (len(ergebnisse), fehler))
                    aktualisiere_zaehler()
                    return
                elif m[0] == "abbruch":
                    laeuft["an"] = False
                    btn_zuordnen.configure(state="normal")
                    status_var.set("ABBRUCH: %s" % m[1])
                    return
        except queue.Empty:
            pass
        root.after(100, abfrage)

    # --- Schritt 2 im Hintergrund ---------------------------------------------
    def arbeite_export(rechnungen, ergebnisse_gueltig, ausgabe_pfad, archiv_ordner, carrier_ordner):
        try:
            bericht = exportiere_alles(rechnungen, ergebnisse_gueltig, ausgabe_pfad,
                                       archiv_ordner, carrier_ordner)
            q.put(("export_fertig", bericht))
        except Exception as e:
            q.put(("export_abbruch", "%s: %s" % (type(e).__name__, e)))

    def starte_export():
        if laeuft["an"]:
            return
        import carrier_export
        import carrier_statistik
        # Schritt 2 verarbeitet NUR, was gerade in der Tabelle sichtbar ist -
        # bei aktivem Filter (siehe wende_filter()) also nur der Ausschnitt.
        # Alles andere bleibt unveraendert in rechnungen_roh/ergebnisse liegen
        # (siehe export_fertig-Behandlung unten), nicht nur im Pool-Ordner.
        sichtbare_indizes = {int(iid) for iid in tv.get_children()}
        alle = list(enumerate(zip(rechnungen_roh, ergebnisse)))
        paare = [(r, b) for i, (r, b) in alle if r is not None and i in sichtbare_indizes]
        # Rechnungen, deren PDF gar nicht erst gelesen werden konnte (kein Positionen
        # erkannt / PDF nicht lesbar) - die werden von Schritt 2 komplett uebersprungen
        # (nicht gepackt, nicht archiviert) und bleiben unveraendert im Pool liegen.
        uebersprungen = [b for i, (r, b) in alle if r is None and i in sichtbare_indizes]
        verarbeitete_indizes = {i for i, (r, _) in alle if r is not None and i in sichtbare_indizes}
        if not paare:
            messagebox.showinfo("Carrier-Dashboard",
                                "Keine gueltigen Rechnungen zum Verarbeiten - bitte zuerst "
                                "Schritt 1 erneut ausfuehren.")
            return

        # Sicherung gegen Teilverarbeitung: ist ein Filter aktiv, wird nur der
        # sichtbare Ausschnitt verarbeitet - das ist laut Matthias die
        # Ausnahme, deshalb eine eigene, explizite Bestaetigung (Vorbelegung
        # "Nein") VOR der normalen Bestaetigung weiter unten.
        if len(sichtbare_indizes) < len(ergebnisse):
            rest = len(ergebnisse) - len(sichtbare_indizes)
            frage_teil = (
                "Es ist ein Filter aktiv: \"%s\".\n\n"
                "Nur %d von %d Rechnung(en) werden JETZT verarbeitet. Die "
                "restlichen %d bleiben UNVERAENDERT liegen (weder gepackt "
                "noch exportiert noch archiviert) und muessen spaeter separat "
                "verarbeitet werden.\n\n"
                "Das ist normalerweise NICHT gewuenscht - trotzdem nur diesen "
                "Ausschnitt jetzt verarbeiten?"
                % (FILTER_TITEL.get(filter_state["schluessel"], filter_state["schluessel"]),
                   len(sichtbare_indizes), len(ergebnisse), rest))
            if not messagebox.askyesno("Carrier-Dashboard - Nur Teil des Laufs verarbeiten?",
                                       frage_teil, icon="warning", default=messagebox.NO):
                return

        # Sicherung gegen Doppel-Verarbeitung: dieselbe Rechnungsnummer zweimal
        # im aktuellen Pool ODER laut Statistik-Historie schon einmal in einem
        # frueheren Schritt-2-Lauf verarbeitet - beides nur mit EXPLIZITER
        # Bestaetigung (Vorbelegung "Nein"), sonst kompletter Abbruch (nichts
        # wird gepackt/exportiert/archiviert).
        rnr_liste = [r.get("rnr") or "" for r, _ in paare]
        doppelt_intern = carrier_statistik.doppelte_im_lauf(rnr_liste)
        ok_hist, doppelt_historie = im_hintergrund(
            lambda: carrier_statistik.bereits_verarbeitet(rnr_liste))
        if not ok_hist:
            if not messagebox.askyesno(
                    "Carrier-Dashboard - Statistik nicht erreichbar",
                    "Die Statistik-Historie ist gerade nicht erreichbar (Netzlaufwerk?) - "
                    "eine Doppel-Verarbeitung frueherer Laeufe kann NICHT geprueft werden "
                    "(Doppelte im aktuellen Pool werden weiterhin erkannt).\n\n"
                    "Trotzdem mit Schritt 2 fortfahren?", icon="warning",
                    default=messagebox.NO):
                return
            doppelt_historie = {}
        if doppelt_intern or doppelt_historie:
            warnung = ["MÖGLICHE DOPPEL-VERARBEITUNG ERKANNT:", ""]
            if doppelt_intern:
                warnung.append("Rechnungsnummer(n) mehrfach im aktuellen Pool:")
                warnung += ["  %s (%dx)" % (rnr, n) for rnr, n in sorted(doppelt_intern.items())]
                warnung.append("")
            if doppelt_historie:
                warnung.append("Rechnungsnummer(n) laut Statistik bereits früher verarbeitet:")
                warnung += ["  %s (zuletzt %s)" % (rnr, max(daten))
                            for rnr, daten in sorted(doppelt_historie.items())]
                warnung.append("")
            warnung.append("Trotzdem mit Schritt 2 fortfahren?")
            if not messagebox.askyesno("Carrier-Dashboard - Doppel-Verarbeitung?",
                                       "\n".join(warnung), icon="warning",
                                       default=messagebox.NO):
                return

        ausgabe_ordner = AUSGABE_ORDNER
        archiv_ordner = ARCHIV_ORDNER
        carrier_ordner = CARRIER_EXPORT_ORDNER
        if os.path.abspath(ausgabe_ordner) == os.path.abspath(POOL_ORDNER):
            # Defensive Pruefung falls die Konstanten oben im Quelltext mal
            # falsch abgeaendert werden - die Pickliste wuerde sonst beim
            # naechsten Lauf als Rechnung mit eingelesen.
            messagebox.showerror("Carrier-Dashboard",
                                 "AUSGABE_ORDNER darf nicht gleich POOL_ORDNER sein - bitte "
                                 "im Quelltext (carrier_dashboard.py, Kopf) korrigieren.")
            return
        n_fehler = sum(1 for _, b in paare if b["status"] == "fehler")
        # Offener (NICHT quittierter) Hinweis blockiert den Carrier-Export
        # GENAUSO wie ein Fehler (siehe carrier_export.exportiere(): nur
        # status "ok" wird exportiert) - Rechnung wird trotzdem gepackt.
        n_hinweis_offen = sum(1 for _, b in paare if hat_offenen_hinweis(b))
        os.makedirs(ausgabe_ordner, exist_ok=True)
        # Sekundengenauer, kollisionssicherer Dateiname (dateiname() haengt bei
        # Bedarf _2/_3/... an) - eine Minute allein reichte nicht aus und liess
        # zwei Laeufe binnen derselben Minute die vorige Pickliste ueberschreiben.
        ausgabe_pfad = carrier_export.dateiname(ausgabe_ordner, "Pickliste", ext="pdf")
        hinweis_fehler = ("\n\n%d davon werden zwar gepackt, aber NICHT in eine "
                          "Carrier-CSV geschrieben (Zuordnungsfehler, siehe Tabelle) - "
                          "diese muessen manuell nachbearbeitet werden." % n_fehler
                          ) if n_fehler else ""
        hinweis_offen_txt = ("\n\n%d davon haben einen noch NICHT quittierten Hinweis und "
                             "werden zwar gepackt, aber ebenfalls NICHT in eine Carrier-CSV "
                             "geschrieben - erst \"Hinweis quittieren\" (oder beheben) und "
                             "danach erneut verarbeiten." % n_hinweis_offen
                             ) if n_hinweis_offen else ""
        hinweis_uebersprungen = ("\n\n%d Rechnung(en) konnten gar nicht gelesen werden "
                                 "(siehe Fehlermeldung in der Tabelle) und werden JETZT NICHT "
                                 "gepackt oder verschoben - sie bleiben unveraendert im "
                                 "Pool-Ordner liegen und muessen manuell geprueft werden."
                                 % len(uebersprungen)) if uebersprungen else ""
        frage = ("%d Rechnung(en) werden verarbeitet:%s%s%s\n\n"
                "Pickliste + Bruecken-CSVs -> %s\n"
                "Carrier-CSVs -> %s\n"
                "Die eingelesenen PDFs werden anschliessend NACH %s VERSCHOBEN "
                "(nicht kopiert).\n\nJetzt ausfuehren?"
                % (len(paare), hinweis_fehler, hinweis_offen_txt, hinweis_uebersprungen,
                   ausgabe_ordner, carrier_ordner, archiv_ordner or "(nicht archiviert)"))
        if not messagebox.askyesno("Carrier-Dashboard", frage):
            return
        export_info["uebersprungen"] = len(uebersprungen)
        export_info["verarbeitete_indizes"] = verarbeitete_indizes
        laeuft["an"] = True
        btn_zuordnen.configure(state="disabled")
        btn_export.configure(state="disabled")
        status_var.set("Erstelle Pickliste + CSVs ...")
        rechnungen_g = [r for r, _ in paare]
        ergebnisse_g = [b for _, b in paare]
        threading.Thread(target=arbeite_export,
                         args=(rechnungen_g, ergebnisse_g, ausgabe_pfad, archiv_ordner,
                               carrier_ordner),
                         daemon=True).start()
        root.after(100, abfrage_export)

    def abfrage_export():
        try:
            while True:
                m = q.get_nowait()
                if m[0] == "export_fertig":
                    bericht = m[1]
                    laeuft["an"] = False
                    btn_zuordnen.configure(state="normal")
                    # NUR die tatsaechlich verarbeiteten Indizes entfernen - alles
                    # andere (uebersprungene PDFs UND durch einen Filter nicht
                    # sichtbare Zeilen) bleibt in der Tabelle stehen, weil es
                    # weder gepackt noch archiviert wurde (siehe starte_export()).
                    verarbeitet = export_info["verarbeitete_indizes"]
                    rechnungen_roh[:] = [r for i, r in enumerate(rechnungen_roh)
                                         if i not in verarbeitet]
                    ergebnisse[:] = [b for i, b in enumerate(ergebnisse) if i not in verarbeitet]
                    for k in set(adress_korrekturen) - {b["rnr"] for b in ergebnisse}:
                        del adress_korrekturen[k]
                    for k in set(quittiert) - {b["rnr"] for b in ergebnisse}:
                        del quittiert[k]
                    filter_state["schluessel"] = None    # Indizes verschoben -> Filter zuruecksetzen
                    fuelle()
                    # Bleiben Zeilen stehen (Filter/uebersprungen), kann direkt ein
                    # weiterer Teil-Lauf gestartet werden, ohne erst Schritt 1 neu
                    # auszufuehren - Button muss dafuer wieder aktiv sein.
                    btn_export.configure(state=("normal" if ergebnisse else "disabled"))
                    aktualisiere_zaehler()
                    verbleibend = len(ergebnisse)
                    status_var.set("Fertig: %d Rechnung(en) verarbeitet, %d archiviert%s." %
                                   (bericht["anzahl"], bericht["archiviert"],
                                    (", %d verbleiben in der Tabelle" % verbleibend)
                                    if verbleibend else ""))
                    zeilen = ["Pickliste:        %s" % bericht["pickliste"],
                             "Rechnungen:       %d" % bericht["anzahl"],
                             "Archiviert:       %d -> %s" % (
                                 bericht["archiviert"], bericht["archiv_ziel"] or "-")]
                    if bericht["archiv_fehler"]:
                        zeilen.append("NICHT archiviert (%d): %s" %
                                      (len(bericht["archiv_fehler"]),
                                       "; ".join(bericht["archiv_fehler"])))
                    if export_info["uebersprungen"]:
                        zeilen.append("UEBERSPRUNGEN (PDF nicht lesbar, liegen noch im "
                                      "Pool): %d" % export_info["uebersprungen"])
                    if verbleibend:
                        zeilen.append("Verbleibend in der Tabelle (nicht verarbeitet, z.B. "
                                      "wegen aktivem Filter oder Lesefehler): %d" % verbleibend)
                    if bericht["kg_geloggt"]:
                        zeilen.append("Kg-Statistik: %d Rechnung(en) erfasst" %
                                      bericht["kg_geloggt"])
                    if bericht["artikel_geloggt"]:
                        zeilen.append("Artikel-Statistik: %d Rechnung(en) erfasst" %
                                      bericht["artikel_geloggt"])
                    if bericht["carrier_dateien"]:
                        zeilen.append("")
                        zeilen.append("Carrier-CSVs:")
                        zeilen += ["  %s (%d Rechnung(en))" % (os.path.basename(p), n)
                                   for p, n in bericht["carrier_dateien"].items()]
                    else:
                        zeilen.append("")
                        zeilen.append("Keine Carrier-CSV geschrieben (keine exportierbare "
                                      "Rechnung dabei).")
                    messagebox.showinfo("Carrier-Dashboard - Fertig", "\n".join(zeilen))
                    return
                elif m[0] == "export_abbruch":
                    laeuft["an"] = False
                    btn_zuordnen.configure(state="normal")
                    btn_export.configure(state=("normal" if ergebnisse else "disabled"))
                    status_var.set("ABBRUCH: %s" % m[1])
                    messagebox.showerror("Carrier-Dashboard - Fehler", m[1])
                    return
        except queue.Empty:
            pass
        root.after(100, abfrage_export)

    btn_zuordnen.configure(command=starte)
    btn_export.configure(command=starte_export)
    aktualisiere_zaehler()
    root.after(REFRESH_MS, tick)
    root.mainloop()


if __name__ == "__main__":
    gui()
    sys.exit(0)
