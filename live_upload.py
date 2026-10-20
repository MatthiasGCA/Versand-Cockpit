# -*- coding: utf-8 -*-
"""
live_upload.py  --  laeuft auf dem Lager-PC (DESKTOP-N2H75H), neben statistik_upload.py
============================================================================
Schickt den Live-Stand des Versandcockpits an die Versandstatistik-Webseite
(Reiter "Live"). Das Cockpit (scan_druck.py) schreibt seinen Stand regelmaessig
als cockpit_live.json in den Paketscheine-Ordner; dieses Skript beobachtet die
Datei und laedt sie per HTTPS-POST an vs-ingest.php (Ziel "live") hoch - sobald
sie sich geaendert hat (das Cockpit schreibt auch ohne Aenderung alle ~30 s ein
Lebenszeichen, die Seite erkennt daran, ob das Cockpit noch laeuft).

Das Upload-Token liegt NUR auf dem Lager-PC (wie bei statistik_upload.py /
artikel_upload.py) und wird von dort uebernommen - die Scanner-PCs brauchen es
nicht.

Dauerlauf (Endlosschleife, alle POLL_S Sekunden pruefen) - einmal beim Anmelden
starten, am bequemsten ueber live_upload.vbs (unsichtbar, Autostart-Ordner).
Eine zweite Instanz beendet sich sofort (Sperre ueber einen lokalen Port).

Nur Standardbibliothek. Konfiguration (alles optional) in live_upload_config.json
neben diesem Skript:
  { "url": "https://www.gasecenter-onlineshop.de/vs-ingest.php",
    "token": "...",                 (sonst wird UPLOAD_TOKEN aus statistik_upload.py gelesen)
    "quelle": "...cockpit_live.json", "log": "..." }

Aufruf:
    py live_upload.py            (Dauerlauf)
    py live_upload.py --once     (einmal pruefen/hochladen, dann beenden - zum Testen)
"""

import json
import os
import re
import socket
import sys
import time
import urllib.error
import urllib.request
from datetime import datetime

VERSION = "2026-10-20a"
BASIS = os.path.dirname(os.path.abspath(__file__))
STANDARD = {
    "url": "https://www.gasecenter-onlineshop.de/vs-ingest.php",
    "quelle": r"\\DESKTOP-N2H75H\Netzwerk\Paketscheine\cockpit_live.json",
    "log": os.path.join(BASIS, "live_upload.log"),     # immer neben dem Skript - kein fester Ordner
}
# Wo nach statistik_upload.py (Token) gesucht wird: zuerst neben diesem Skript,
# dann die ueblichen Ordner auf dem Lager-PC.
TOKEN_SUCHORTE = [BASIS, r"C:\VersandStatistik", r"C:\Packlisten"]
POLL_S = 5                    # so oft nach einer geaenderten Datei sehen
MAX_ALTER_S = 120             # aelter als das wird NICHT hochgeladen (Cockpit tot -> Seite zeigt das selbst)
LOG_MAX_BYTES = 300_000       # Log wird bei Ueberschreitung auf die jungen Haelfte gekuerzt
SPERR_PORT = 47655            # lokaler Port als Einzelinstanz-Sperre
FEHLER_LOG_ABSTAND_S = 600    # gleiche Fehlerart hoechstens alle 10 Minuten loggen


def lade_config():
    cfg = dict(STANDARD)
    pfad = os.path.join(BASIS, "live_upload_config.json")
    if os.path.exists(pfad):
        with open(pfad, encoding="utf-8-sig") as f:
            cfg.update({k: v for k, v in json.load(f).items() if isinstance(v, str) and v.strip()})
    if not cfg.get("token"):
        for ordner in TOKEN_SUCHORTE:
            kand = os.path.join(ordner, "statistik_upload.py")
            if os.path.exists(kand):
                m = re.search(r'^UPLOAD_TOKEN\s*=\s*["\'](.*)["\']\s*(#.*)?$',
                              open(kand, encoding="utf-8-sig").read(), re.M)
                if m:
                    cfg["token"] = m.group(1)
                    break
    return cfg


def log(text, pfad):
    zeile = "%s  %s\n" % (datetime.now().strftime("%Y-%m-%d %H:%M:%S"), text)
    print(zeile, end="")
    try:
        if os.path.exists(pfad) and os.path.getsize(pfad) > LOG_MAX_BYTES:
            with open(pfad, encoding="utf-8", errors="replace") as f:
                zeilen = f.readlines()
            with open(pfad, "w", encoding="utf-8") as f:
                f.writelines(zeilen[len(zeilen) // 2:])
        with open(pfad, "a", encoding="utf-8") as f:
            f.write(zeile)
    except OSError:
        pass


def hochladen(roh, url, token):
    ziel = url + ("&" if "?" in url else "?") + "ziel=live"
    req = urllib.request.Request(ziel, data=roh, method="POST", headers={
        "Content-Type": "application/json; charset=utf-8", "X-Upload-Token": token,
        "User-Agent": "live_upload/" + VERSION})
    try:
        with urllib.request.urlopen(req, timeout=20) as r:
            return r.status, r.read().decode("utf-8", "replace").strip()
    except urllib.error.HTTPError as e:
        # Antworttext des Servers mitnehmen und die haeufigen Ursachen benennen
        try:
            body = e.read().decode("utf-8", "replace").strip()[:200]
        except OSError:
            body = ""
        tipps = {400: "Webspace hat noch die ALTE vs-ingest.php (kennt Ziel 'live' nicht)",
                 403: "Token stimmt nicht mit dem Geheimwort in vs-ingest.php ueberein",
                 404: "vs-ingest.php nicht unter dieser URL gefunden",
                 413: "Datei zu gross", 500: "Server kann live.json nicht schreiben "
                      "(Ordner versandstatistik/ fehlt oder nicht beschreibbar)"}
        raise OSError("HTTP %s %s%s" % (e.code, ("- '%s' " % body) if body else "",
                                        ("-> " + tipps[e.code]) if e.code in tipps else ""))


def einmal(cfg, state):
    """Prueft die Quelldatei und laedt hoch, wenn neu. Rueckgabe: kurzer Status-Text."""
    q = cfg["quelle"]
    try:
        mtime = os.path.getmtime(q)
        groesse = os.path.getsize(q)
    except OSError:
        return "Quelle nicht erreichbar"
    if (mtime, groesse) == state.get("sig"):
        return "unveraendert"
    try:
        with open(q, "rb") as f:
            roh = f.read()
        daten = json.loads(roh.decode("utf-8"))
    except (OSError, ValueError):
        return "Quelle unlesbar (wird gerade geschrieben?)"
    # Alter ueber die Datei-mtime (setzt der Dateiserver beim Schreiben - hier der Lager-PC selbst),
    # NICHT ueber das epoch-Feld: das stammt von der Uhr des Scanner-PCs - laeuft die mehr als
    # MAX_ALTER_S gegenueber dem Lager-PC falsch, wuerde sonst nie (oder zu lange) hochgeladen.
    alter = time.time() - mtime
    if alter > MAX_ALTER_S:
        state["sig"] = (mtime, groesse)
        return "Quelle zu alt (%d s) - Cockpit laeuft nicht" % alter
    status, antwort = hochladen(roh, cfg["url"], cfg["token"])
    state["sig"] = (mtime, groesse)
    state["ok_ts"] = time.time()
    return "Upload OK (HTTP %s: %s)" % (status, antwort)


def main(argv=None):
    argv = sys.argv[1:] if argv is None else argv
    nur_einmal = "--once" in argv
    cfg = lade_config()
    lg = cfg["log"]
    if not cfg.get("token"):
        log("FEHLER: kein Upload-Token gefunden. Gesucht wurde statistik_upload.py in: %s - "
            "bitte live_upload.py in denselben Ordner wie statistik_upload.py legen oder "
            "live_upload_config.json mit {\"token\": \"...\"} anlegen." % ", ".join(TOKEN_SUCHORTE), lg)
        return 1
    sperre = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    if not nur_einmal:        # --once (Test) darf auch neben einem laufenden Dauerlauf starten
        try:
            sperre.bind(("127.0.0.1", SPERR_PORT))
        except OSError:
            print("live_upload.py laeuft bereits im Hintergrund (Dauerlauf) - das ist in Ordnung, "
                  "diese zweite Instanz beendet sich. Zum Testen: py live_upload.py --once")
            return 0
    log("live_upload %s gestartet%s, Quelle: %s" % (VERSION, " (einmalig)" if nur_einmal else "",
                                                     cfg["quelle"]), lg)
    state, letzter_fehler = {}, {}
    while True:
        try:
            text = einmal(cfg, state)
            if text.startswith("Upload OK") and not letzter_fehler.get("lief_ok"):
                log(text, lg)                      # nur den ERSTEN Erfolg nach einem Fehler/Start loggen
                letzter_fehler.clear()
                letzter_fehler["lief_ok"] = True
            elif not text.startswith(("Upload OK", "unveraendert")):
                if time.time() - letzter_fehler.get(text, 0) > FEHLER_LOG_ABSTAND_S:
                    log(text, lg)
                    letzter_fehler[text] = time.time()
                letzter_fehler["lief_ok"] = False
            if nur_einmal:
                print(text)
                return 0
        except (urllib.error.URLError, OSError) as e:
            t = "FEHLER beim Upload: %s" % e
            if time.time() - letzter_fehler.get(t, 0) > FEHLER_LOG_ABSTAND_S:
                log(t, lg)
                letzter_fehler[t] = time.time()
            letzter_fehler["lief_ok"] = False
            if nur_einmal:
                return 1
        time.sleep(POLL_S)


if __name__ == "__main__":
    sys.exit(main())
