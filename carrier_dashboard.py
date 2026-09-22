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
      mit carrier-status != "fehler" UND zugeordnetem Carrier landen in einer
      Carrier-CSV - alle anderen werden trotzdem gepackt (Pickliste), aber
      NICHT automatisch exportiert (manuelle Nachbearbeitung noetig).
  Schritt 4 (spaeter, optional): Warnung in scan_druck.py bei Carrier-
      Abweichung zwischen Label und dieser Zuordnung.

Die Regeln stehen in carrier_regeln.py, der CSV-Export in carrier_export.py,
das Auslesen/die Pickliste in packliste.py - alle vier Dateien muessen im
selben Ordner liegen.

Start:  py carrier_dashboard.py
Pool-/Ausgabe-/Archiv-/Carrier-Export-Ordner werden in carrier_config.json
neben dem Skript gemerkt.
"""

import glob
import json
import os
import queue
import sys
import threading
import tkinter as tk
from datetime import datetime
from tkinter import filedialog, messagebox, ttk

import carrier_regeln as regeln

VERSION = "2026-09-22a"

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
    return [t[0] for t in treffer], [t[1] for t in treffer]


def exportiere_alles(rechnungen, ergebnisse, ausgabe_pfad, archiv_ordner, carrier_ordner):
    """Schritt 2: rechnungen (gueltige parse_pdf()-Dicts, "quelle" gesetzt) UND
    ergebnisse (dazu bewertete Carrier-Ergebnisse, GLEICHE Reihenfolge/Laenge)
    -> Pickliste-PDF + fuenf Bruecken-CSVs (wie packliste.main(), Layout/Logik
    unveraendert) im Ordner von ausgabe_pfad, Carrier-CSVs nach carrier_ordner
    (carrier_export.exportiere - exportiert nur status != 'fehler' MIT
    Carrier), danach Archivierung der eingelesenen PDFs nach archiv_ordner
    (packliste.archiviere - verschiebt nur, was erfolgreich verarbeitet wurde).
    Gibt einen Berichts-dict zurueck; einzelne Archiv-Fehler werfen KEINE
    Exception, sondern stehen im Bericht (siehe packliste.archiviere)."""
    import carrier_export
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

    verschoben, archiv_fehler, archiv_ziel = 0, [], None
    if archiv_ordner:
        verschoben, archiv_fehler, archiv_ziel = packliste.archiviere(rechnungen, archiv_ordner)

    return {
        "pickliste": ausgabe_pfad, "anzahl": len(rechnungen), "gruppen": gruppen,
        "wc_neu": wc_neu, "carrier_dateien": carrier_dateien,
        "archiviert": verschoben, "archiv_fehler": archiv_fehler, "archiv_ziel": archiv_ziel,
    }


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
    ausgabe_var = tk.StringVar(
        value=cfg.get("ausgabe_ordner") or os.path.join(pool_var.get(), "Pickliste"))
    archiv_var = tk.StringVar(
        value=cfg.get("archiv_ordner") or os.path.join(pool_var.get(), "Archiv"))
    carrier_var = tk.StringVar(value=cfg.get("carrier_export_ordner") or r"C:\Carrier_Export")
    pool_anz = tk.StringVar(value="")
    status_var = tk.StringVar(value="Noch nicht zugeordnet.")
    ergebnisse = []
    rechnungen_roh = []
    q = queue.Queue()
    laeuft = {"an": False}

    # --- Kopf: Ordner + Zaehler ---------------------------------------------
    kopf = ttk.Frame(root, padding=8)
    kopf.pack(fill="x")

    def _pfadzeile(row, label, var, merk_schluessel):
        ttk.Label(kopf, text=label).grid(row=row, column=0, sticky="w")
        ttk.Entry(kopf, textvariable=var).grid(row=row, column=1, sticky="ew", padx=6)

        def waehle():
            d = filedialog.askdirectory(initialdir=var.get() or HIER)
            if d:
                var.set(os.path.normpath(d))
                cfg[merk_schluessel] = var.get()
                speichere_config(cfg)
                aktualisiere_zaehler()

        ttk.Button(kopf, text="Durchsuchen ...", command=waehle).grid(row=row, column=2)

    _pfadzeile(0, "Pool-Ordner (Rechnungen):", pool_var, "pool_ordner")
    _pfadzeile(1, "Ausgabe (Pickliste + Bruecken-CSVs):", ausgabe_var, "ausgabe_ordner")
    _pfadzeile(2, "Archiv-Ordner:", archiv_var, "archiv_ordner")
    _pfadzeile(3, "Carrier-Export-Ordner:", carrier_var, "carrier_export_ordner")
    kopf.columnconfigure(1, weight=1)

    zaehler = tk.Label(kopf, textvariable=pool_anz, font=("Segoe UI", 18, "bold"), anchor="w")
    zaehler.grid(row=4, column=0, columnspan=3, sticky="w", pady=(8, 0))

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
    btn_export = ttk.Button(knoepfe, text="2. Pickliste + CSV erstellen", state="disabled")
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

    # --- Schritt 1 im Hintergrund --------------------------------------------
    def arbeite(ordner):
        try:
            roh, erg = lese_pool(ordner, lambda i, n, name: q.put(("fortschritt", i, n, name)))
            q.put(("fertig", roh, erg))
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
        btn_export.configure(state="disabled")
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
                    rechnungen_roh[:] = m[1]
                    ergebnisse[:] = m[2]
                    fuelle()
                    laeuft["an"] = False
                    btn_zuordnen.configure(state="normal")
                    btn_export.configure(state=("normal" if ergebnisse else "disabled"))
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
        paare = [(r, b) for r, b in zip(rechnungen_roh, ergebnisse) if r is not None]
        if not paare:
            messagebox.showinfo("Carrier-Dashboard",
                                "Keine gueltigen Rechnungen zum Verarbeiten - bitte zuerst "
                                "Schritt 1 erneut ausfuehren.")
            return
        ausgabe_ordner = ausgabe_var.get().strip()
        pool_abs = os.path.abspath(pool_var.get() or "")
        if not ausgabe_ordner or os.path.abspath(ausgabe_ordner) == pool_abs:
            messagebox.showerror("Carrier-Dashboard",
                                 "Der Ausgabe-Ordner darf nicht der Pool-Ordner selbst sein "
                                 "(die Pickliste wuerde sonst beim naechsten Lauf als "
                                 "Rechnung mit eingelesen). Bitte einen Unterordner waehlen, "
                                 "z.B. %s." % os.path.join(pool_var.get(), "Pickliste"))
            return
        archiv_ordner = archiv_var.get().strip()
        carrier_ordner = carrier_var.get().strip()
        n_fehler = sum(1 for _, b in paare if b["status"] == "fehler")
        stamp = datetime.now().strftime("%Y-%m-%d_%H%M")
        ausgabe_pfad = os.path.join(ausgabe_ordner, "Pickliste_%s.pdf" % stamp)
        hinweis_fehler = ("\n\n%d davon werden zwar gepackt, aber NICHT in eine "
                          "Carrier-CSV geschrieben (Zuordnungsfehler, siehe Tabelle) - "
                          "diese muessen manuell nachbearbeitet werden." % n_fehler
                          ) if n_fehler else ""
        frage = ("%d Rechnung(en) werden verarbeitet:%s\n\n"
                "Pickliste + Bruecken-CSVs -> %s\n"
                "Carrier-CSVs -> %s\n"
                "Die eingelesenen PDFs werden anschliessend NACH %s VERSCHOBEN "
                "(nicht kopiert).\n\nJetzt ausfuehren?"
                % (len(paare), hinweis_fehler, ausgabe_ordner, carrier_ordner,
                   archiv_ordner or "(nicht archiviert)"))
        if not messagebox.askyesno("Carrier-Dashboard", frage):
            return
        cfg["ausgabe_ordner"] = ausgabe_ordner
        cfg["archiv_ordner"] = archiv_ordner
        cfg["carrier_export_ordner"] = carrier_ordner
        speichere_config(cfg)
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
                    rechnungen_roh.clear()
                    ergebnisse.clear()
                    fuelle()
                    aktualisiere_zaehler()
                    status_var.set("Fertig: %d Rechnung(en) verarbeitet, %d archiviert." %
                                   (bericht["anzahl"], bericht["archiviert"]))
                    zeilen = ["Pickliste:        %s" % bericht["pickliste"],
                             "Rechnungen:       %d" % bericht["anzahl"],
                             "Archiviert:       %d -> %s" % (
                                 bericht["archiviert"], bericht["archiv_ziel"] or "-")]
                    if bericht["archiv_fehler"]:
                        zeilen.append("NICHT archiviert (%d): %s" %
                                      (len(bericht["archiv_fehler"]),
                                       "; ".join(bericht["archiv_fehler"])))
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
