# -*- coding: utf-8 -*-
"""
carrier_dashboard.py - Carrier-Dashboard (Schritt 1: nur LESEN und ZUORDNEN)
============================================================================
Liest alle Rechnungs-PDFs eines Pool-Ordners, ordnet jede Bestellung einem
Versanddienstleister zu (DHL / DPD / Deutsche Post Brief / Grossbrief, je
Inland und Ausland) und zeigt das Ergebnis samt allen Fehlern an.

  Schritt 1 (diese Version): Zuordnung + Anzeige. Es wird NICHTS geschrieben,
      nichts verschoben, keine Pickliste erzeugt - die PDFs bleiben liegen.
  Schritt 2 (folgt): CSV-Dateien fuer DHL/DPD/Post nach C:\\Carrier_Export.
  Schritt 3 (folgt): Pickliste + Archivierung aus demselben Lauf.

Die Regeln stehen in carrier_regeln.py, das Auslesen der Rechnungen in
packliste.py (parse_pdf) - beide muessen im selben Ordner liegen.

Start:  py carrier_dashboard.py
Der zuletzt benutzte Pool-Ordner wird in carrier_config.json neben dem Skript
gemerkt.
"""

import glob
import json
import os
import queue
import sys
import threading
import tkinter as tk
from tkinter import filedialog, ttk

import carrier_regeln as regeln

VERSION = "2026-09-21a"

HIER = os.path.dirname(os.path.abspath(__file__))
CONFIG_PFAD = os.path.join(HIER, "carrier_config.json")
POOL_STANDARD = os.path.join(HIER, "Pool")
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


def gruppen_schluessel(b):
    """Zusammenfassungs-Schluessel eines Ergebnisses (None bei Fehler)."""
    c = b["carrier"]
    if c is None or b["status"] == "fehler":
        return None
    if c in (regeln.BRIEF, regeln.GROSSBRIEF):
        return "%s|%s" % (c, "AUS" if b["ausland"] else "DE")
    return c


def lade_config():
    try:
        with open(CONFIG_PFAD, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def speichere_config(cfg):
    try:
        with open(CONFIG_PFAD, "w", encoding="utf-8") as f:
            json.dump(cfg, f, indent=2, ensure_ascii=False)
    except Exception:
        pass


def lese_pool(ordner, melde=None):
    """Liest alle PDFs des Ordners und bewertet sie. Rueckgabe: Liste von
    Ergebnis-dicts (sortiert nach Rechnungsnummer, Fehler ohne Nummer am Ende)."""
    import packliste
    pdfs = sorted(glob.glob(os.path.join(ordner, "*.pdf")))
    ergebnisse = []
    for i, pfad in enumerate(pdfs, 1):
        name = os.path.basename(pfad)
        if melde:
            melde(i, len(pdfs), name)
        try:
            r = packliste.parse_pdf(pfad)
        except Exception as e:
            ergebnisse.append(_fehlerzeile(name, "PDF nicht lesbar: %s" % e))
            continue
        if not r.get("positionen"):
            ergebnisse.append(_fehlerzeile(
                name, "Keine Positionen erkannt (keine Rechnung?)", r.get("rnr", "")))
            continue
        r["datei"] = name
        b = regeln.bewerte_rechnung(r)
        b["quelle"] = pfad
        ergebnisse.append(b)
    ergebnisse.sort(key=lambda b: (not b["rnr"], b["rnr"], b["datei"]))
    return ergebnisse


def _fehlerzeile(datei, text, rnr=""):
    adr = regeln.analysiere_adresse([])
    return {"rnr": rnr, "datei": datei, "adresse": adr, "kennungen": [],
            "gewicht": None, "carrier": None, "ausland": False, "grund": "",
            "fehler": [text], "hinweise": [], "status": "fehler", "quelle": ""}


def zaehle(ergebnisse):
    z = {k: 0 for k, _ in GRUPPEN}
    fehler = 0
    for b in ergebnisse:
        k = gruppen_schluessel(b)
        if k is None:
            fehler += 1
        else:
            z[k] = z.get(k, 0) + 1
    return z, fehler


def detailtext(b):
    a = b["adresse"]
    zeilen = ["Rechnung:  %s   (%s)" % (b["rnr"] or "?", b["datei"])]
    zeilen.append("Kennung:   %s" % (", ".join(b["kennungen"]) or "-"))
    g = b["gewicht"]
    zeilen.append("Gewicht:   %s" % (("%.3f kg" % g).replace(".", ",") if g is not None else "-"))
    zeilen.append("Carrier:   %s%s" % (
        b["carrier"] or "-", "  (Ausland)" if b["carrier"] and b["ausland"] else ""))
    if b["grund"]:
        zeilen.append("Grund:     %s" % b["grund"])
    zeilen.append("")
    zeilen.append("Adresse laut Auswertung:")
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
        zeilen.append("Hinweise:")
        zeilen += ["  - " + h for h in b["hinweise"]]
    return "\n".join(zeilen)


# ==========================================================================
# GUI
# ==========================================================================

def gui():
    cfg = lade_config()
    root = tk.Tk()
    root.title("Carrier-Dashboard  (Version %s / Regeln %s)" % (VERSION, regeln.VERSION))
    root.geometry("1100x720")

    pool_var = tk.StringVar(value=cfg.get("pool_ordner") or POOL_STANDARD)
    pool_anz = tk.StringVar(value="")
    status_var = tk.StringVar(value="Noch nicht zugeordnet.")
    ergebnisse = []
    q = queue.Queue()
    laeuft = {"an": False}

    # --- Kopf: Pool-Ordner + Zaehler ---------------------------------------
    kopf = ttk.Frame(root, padding=8)
    kopf.pack(fill="x")
    ttk.Label(kopf, text="Pool-Ordner:").grid(row=0, column=0, sticky="w")
    ttk.Entry(kopf, textvariable=pool_var).grid(row=0, column=1, sticky="ew", padx=6)
    kopf.columnconfigure(1, weight=1)

    def waehle():
        d = filedialog.askdirectory(initialdir=pool_var.get() or HIER)
        if d:
            pool_var.set(os.path.normpath(d))
            cfg["pool_ordner"] = pool_var.get()
            speichere_config(cfg)
            aktualisiere_zaehler()

    ttk.Button(kopf, text="Durchsuchen ...", command=waehle).grid(row=0, column=2)

    zaehler = tk.Label(kopf, textvariable=pool_anz, font=("Segoe UI", 18, "bold"), anchor="w")
    zaehler.grid(row=1, column=0, columnspan=3, sticky="w", pady=(8, 0))

    def aktualisiere_zaehler():
        ordner = pool_var.get()
        if os.path.isdir(ordner):
            n = len(glob.glob(os.path.join(ordner, "*.pdf")))
            pool_anz.set("Rechnungen im Pool: %d" % n)
        else:
            pool_anz.set("Pool-Ordner nicht gefunden")

    def tick():
        if not laeuft["an"]:
            aktualisiere_zaehler()
        root.after(REFRESH_MS, tick)

    # --- Buttons -------------------------------------------------------------
    knoepfe = ttk.Frame(root, padding=(8, 0))
    knoepfe.pack(fill="x")
    btn_zuordnen = ttk.Button(knoepfe, text="1. Bestellungen zuordnen")
    btn_zuordnen.pack(side="left")
    btn_export = ttk.Button(knoepfe, text="2. Pickliste + CSV erstellen  (folgt in Schritt 2/3)",
                            state="disabled")
    btn_export.pack(side="left", padx=8)
    ttk.Label(knoepfe, textvariable=status_var).pack(side="left", padx=12)
    prog = ttk.Progressbar(root, mode="determinate")
    prog.pack(fill="x", padx=8, pady=(6, 0))

    # --- Zusammenfassung ----------------------------------------------------
    zf = ttk.LabelFrame(root, text="Zuordnung", padding=6)
    zf.pack(fill="x", padx=8, pady=8)
    zlabels = {}
    for i, (k, titel) in enumerate(GRUPPEN + [("FEHLER", "Fehler")]):
        ttk.Label(zf, text=titel).grid(row=0, column=i, padx=10)
        lb = tk.Label(zf, text="-", font=("Segoe UI", 16, "bold"))
        lb.grid(row=1, column=i, padx=10)
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
    tv.tag_configure("fehler", background="#f8d0d0")
    tv.tag_configure("warn", background="#fff2c2")
    tv.tag_configure("ok", background="#ffffff")
    sb = ttk.Scrollbar(tv, orient="vertical", command=tv.yview)
    tv.configure(yscrollcommand=sb.set)
    sb.pack(side="right", fill="y")
    pan.add(tv, weight=3)
    detail = tk.Text(pan, height=12, font=("Consolas", 10), wrap="word", state="disabled")
    pan.add(detail, weight=1)

    def zeige_detail(_evt=None):
        sel = tv.selection()
        if not sel:
            return
        idx = int(sel[0])
        detail.configure(state="normal")
        detail.delete("1.0", "end")
        detail.insert("1.0", detailtext(ergebnisse[idx]))
        detail.configure(state="disabled")

    tv.bind("<<TreeviewSelect>>", zeige_detail)

    def fuelle():
        tv.delete(*tv.get_children())
        for i, b in enumerate(ergebnisse):
            if b["fehler"]:
                meldung = "; ".join(b["fehler"])
            elif b["hinweise"]:
                meldung = b["grund"] + "  |  " + "; ".join(b["hinweise"])
            else:
                meldung = b["grund"]
            g = b["gewicht"]
            c = b["carrier"] or "-"
            if b["carrier"] in (regeln.BRIEF, regeln.GROSSBRIEF):
                c += " Ausland" if b["ausland"] else " Inland"
            tv.insert("", "end", iid=str(i), tags=(b["status"],), values=(
                b["rnr"] or "?", c, b["adresse"]["land"] or "?",
                ("%.3f" % g).replace(".", ",") if g is not None else "-",
                ",".join(b["kennungen"]) or "-",
                {"ok": "OK", "warn": "Hinweis", "fehler": "FEHLER"}[b["status"]], meldung))
        z, fehler = zaehle(ergebnisse)
        for k, _ in GRUPPEN:
            zlabels[k].configure(text=str(z.get(k, 0)), fg="black")
        zlabels["FEHLER"].configure(text=str(fehler), fg=("#b00000" if fehler else "black"))

    # --- Lauf im Hintergrund -------------------------------------------------
    def arbeite(ordner):
        try:
            erg = lese_pool(ordner, lambda i, n, name: q.put(("fortschritt", i, n, name)))
            q.put(("fertig", erg))
        except Exception as e:                       # z.B. reportlab/pdfplumber fehlt
            q.put(("abbruch", "%s: %s" % (type(e).__name__, e)))

    def starte():
        ordner = pool_var.get()
        if laeuft["an"]:
            return
        if not os.path.isdir(ordner):
            status_var.set("Pool-Ordner nicht gefunden: %s" % ordner)
            return
        cfg["pool_ordner"] = ordner
        speichere_config(cfg)
        laeuft["an"] = True
        btn_zuordnen.configure(state="disabled")
        prog.configure(value=0)
        status_var.set("Lese Rechnungen ...")
        threading.Thread(target=arbeite, args=(ordner,), daemon=True).start()
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
                    ergebnisse[:] = m[1]
                    fuelle()
                    laeuft["an"] = False
                    btn_zuordnen.configure(state="normal")
                    z, fehler = zaehle(ergebnisse)
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

    btn_zuordnen.configure(command=starte)
    aktualisiere_zaehler()
    root.after(REFRESH_MS, tick)
    root.mainloop()


if __name__ == "__main__":
    gui()
    sys.exit(0)
