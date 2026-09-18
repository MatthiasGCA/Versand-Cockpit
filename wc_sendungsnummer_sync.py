# -*- coding: utf-8 -*-
"""
wc_sendungsnummer_sync.py - Carrier-Sendungsnummern in den WooCommerce-Shop
==========================================================================
Traegt die Sendungsnummern aus den drei Carrier-Export-CSVs (DHL, DPD,
Deutsche Post) ueber die WooCommerce-/Shiptastic-REST-API in die passende
Bestellung des Onlineshops (gasecenter-onlineshop.de) ein.

Pendant zum bestehenden Amazon/eBay-Tool "AfterSell", technisch unabhaengig
von Versand-Cockpit / packliste.py. Laeuft auf DEMSELBEN PC wie Amicron /
AfterSell / packliste.py.

ABLAUF
------
1. Eingangsordner einlesen (Standard: C:\\Scripts\\Sendungsnummern_WC):
   - wc_bestellnummern.csv  (von packliste.py erzeugt)
       Rechnungsnummer -> WooCommerce-Bestellnummer  (nur Onlineshop-Rechnungen)
   - DHL-VLS*.csv            ISO-8859-1 ; "Empfaengerreferenz" -> Rg-Nr,
                                          "Sendungsnummer"     -> Tracking
   - *_Sendungsnummern.csv   ISO-8859-1 ; "REFERENZ"           -> Rg-Nr,
                                          "SENDUNGSNUMMKER"    -> Tracking (Dt. Post)
   - EXPORT_*.csv            UTF-8-BOM  ; "Sendungsreferenz 1" -> Rg-Nr,
                                          "Paketnummer"        -> Tracking (DPD, 14-stellig,
                                          fuehrende 0 - NUR als Text behandeln, nie als Zahl)
2. Join ueber die Rechnungsnummer -> (WC-Bestellnummer, Sendungsnummer(n), Carrier).
3. Je Bestellung:
   - GET /wc/v3/shipments?order_id=<nr>   (Shiptastic legt je Bestellung
     automatisch eine Sendung an - meist mit leerer tracking_id)
   - vorhandene freie Sendung per PUT /wc/v3/shipments/<id> mit tracking_id +
     shipping_provider (+ status="shipped", konfigurierbar) befuellen
   - nur falls KEINE freie Sendung existiert: POST /wc/v3/shipments
   - Mehrere Pakete je Bestellung -> erste Nummer auf die vorhandene Sendung,
     weitere als zusaetzliche Sendungen (POST).
4. Statusdatei fuehren, damit jede Bestellung/Sendungsnummer nur einmal
   geschrieben wird. Die Rechnungsnummer->Bestellnummer-Zuordnung wird KUMULATIV
   in der Statusdatei mitgefuehrt - ein spaeterer Carrier-Export kann eine
   Rechnung referenzieren, deren Zeile in wc_bestellnummern.csv inzwischen
   verfallen ist.
5. Nach einem --commit-Lauf: verarbeitete Carrier-CSVs, die aelter als
   "archive_after_days" sind (Standard 14), nach
   <Eingangsordner>\Archiv\<Datum>\ verschieben - sonst wuerde der Ordner
   unbegrenzt weiterwachsen (PaketImport-Mover.ps1 haengt bei einer Namens-
   kollision wie "DHL-VLS-Export.csv" einen Zeitstempel an, ueberschreibt
   also nie) UND jeder Lauf wuerde jede Datei fuer immer erneut einlesen.
   Eine Datei mit einer NOCH offenen Sendungsnummer (Grund "api_fehler" -
   soll beim naechsten Lauf erneut versucht werden) wird NICHT archiviert,
   auch wenn sie schon aelter ist. Eine Datei, deren Spalten nicht erkannt
   wurden (Parse-Fehler), wird ebenfalls nie automatisch archiviert - die
   braucht einen manuellen Blick. "keine WC-Bestellnummer bekannt" gilt
   dagegen NICHT als offen - das sind i.d.R. Nicht-Onlineshop-Sendungen
   (Zaehltheke/ERP), die nie eine Zuordnung bekommen werden; die Datei
   verfaellt nach der Frist ganz normal, analog zu BRUECKEN_CSV_MERGE_TAGE
   in packliste.py. Nur bei einem vollstaendigen Lauf (kein --only/--limit)
   aktiv, damit ein Testlauf nichts verschiebt.

SICHERHEIT
----------
- Ohne --commit laeuft alles als TROCKENLAUF (nur GETs, zeigt was es taete).
  Erst --commit schreibt tatsaechlich (PUT/POST) - und kann je nach
  Shiptastic-Einstellung eine Versandbestaetigung mit Trackinglink an den
  Kunden ausloesen (status="shipped").
- Eine bereits gefuellte, ABWEICHENDE tracking_id wird NIE ueberschrieben,
  sondern als Konflikt gemeldet.
- Eine WC-Bestellnummer, die sich per API nicht als echte Bestellung aufloesen
  laesst, wird uebersprungen (faengt eine faelschlich als Onlineshop erkannte
  Kd-Nr ab).

AUFRUF
------
  python wc_sendungsnummer_sync.py [<eingangsordner>] [Optionen]
    --config PATH     Konfig-JSON (Standard: wc_sync_config.json neben dem Skript)
    --state  PATH     Statusdatei (Standard: aus Konfig, sonst wc_sync_state.json neben dem Skript)
    --commit          tatsaechlich schreiben (sonst Trockenlauf)
    --only RGNR       nur diese eine Rechnungsnummer verarbeiten (Test)
    --limit N         hoechstens N Bestellungen schreiben (Test)

KONFIG (wc_sync_config.json - Vorlage: wc_sync_config.example.json):
{
  "base_url": "https://www.gasecenter-onlineshop.de",
  "consumer_key": "ck_...",
  "consumer_secret": "cs_...",
  "input_dir": "C:\\Scripts\\Sendungsnummern_WC",
  "state_path": "C:\\Scripts\\wc_sync_state.json",
  "set_status_shipped": true,
  "archive_after_days": 14,
  "provider_slugs": { "DHL": "dhl", "DPD": "dpd", "Deutsche Post": "deutsche_post" }
}
"""

import argparse
import base64
import csv
import glob
import json
import os
import shutil
import sys
import time
import urllib.error
import urllib.request
from datetime import datetime

SKRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# --- Carrier-Definitionen ---------------------------------------------------
# Die Fallback-Indizes stammen aus der manuellen Spaltenzaehlung an echten
# Dateien (2026-08-28), 0-basiert:
#   DHL           Rg = Sp. 24  "Empfaengerreferenz"   Tracking = Sp. 106 "Sendungsnummer"
#   Deutsche Post Rg = Sp. 10  "REFERENZ"             Tracking = Sp. 1   "SENDUNGSNUMMKER"
#   DPD           Rg = Sp. 22  "Sendungsreferenz 1"   Tracking = Sp. 15  "Paketnummer"
CARRIERS = [
    {
        "name": "DHL",
        "muster": "DHL-VLS*.csv",
        "encoding": "latin-1",
        "ref_spalten": ["Empf\xe4ngerreferenz"],
        "track_spalten": ["Sendungsnummer"],
        "ref_fallback_idx": 23,
        "track_fallback_idx": 105,
    },
    {
        "name": "Deutsche Post",
        "muster": "*_Sendungsnummern.csv",
        "encoding": "latin-1",
        "ref_spalten": ["REFERENZ"],
        "track_spalten": ["SENDUNGSNUMMKER", "SENDUNGSNUMMER"],
        "ref_fallback_idx": 9,
        "track_fallback_idx": 0,
    },
    {
        "name": "DPD",
        "muster": "EXPORT_*.csv",
        "encoding": "utf-8-sig",
        "ref_spalten": ["Sendungsreferenz 1"],
        # Sp. 15 "Paketnummer" ist die auf dpd.de trackbare 14-stellige Nummer
        # (beginnt mit 0186). NICHT Sp. 17 "Sendungsnummer" (MPS-Master).
        "track_spalten": ["Paketnummer"],
        "ref_fallback_idx": 21,
        "track_fallback_idx": 14,
    },
]

DEFAULT_PROVIDER_SLUGS = {"DHL": "dhl", "DPD": "dpd", "Deutsche Post": "deutsche_post"}
# Alles laeuft auf einem PC (Amicron / AfterSell / packliste.py) - lokaler Ordner,
# kein Netzwerkpfad. packliste.py und PaketImport-Mover.ps1 legen hier ab.
DEFAULT_INPUT_DIR = r"C:\Scripts\Sendungsnummern_WC"


# ==========================================================================
# Konfig / State
# ==========================================================================

def lade_konfig(pfad):
    if not os.path.exists(pfad):
        sys.exit("FEHLER: Konfigdatei nicht gefunden: %s\n"
                 "       Vorlage: wc_sync_config.example.json" % pfad)
    with open(pfad, "r", encoding="utf-8") as f:
        cfg = json.load(f)
    for pflicht in ("base_url", "consumer_key", "consumer_secret"):
        if not cfg.get(pflicht):
            sys.exit("FEHLER: '%s' fehlt in %s" % (pflicht, pfad))
    cfg["base_url"] = cfg["base_url"].rstrip("/")
    cfg.setdefault("set_status_shipped", True)
    cfg.setdefault("archive_after_days", 14)
    slugs = dict(DEFAULT_PROVIDER_SLUGS)
    slugs.update(cfg.get("provider_slugs") or {})
    cfg["provider_slugs"] = slugs
    return cfg


def lade_state(pfad):
    try:
        with open(pfad, "r", encoding="utf-8") as f:
            st = json.load(f)
    except FileNotFoundError:
        st = {}
    except Exception as e:
        sys.exit("FEHLER: Statusdatei %s unlesbar (%s). "
                 "NICHT von Hand loeschen - im Zweifel Ruecksprache." % (pfad, e))
    st.setdefault("mapping", {})   # rgnr -> {"bestellnr","email","name"}
    st.setdefault("synced", {})    # rgnr -> {"order_id","trackings":[...],"carrier","provider","ts"}
    st.setdefault("offen", {})     # rgnr -> {"grund","carrier","tracking",...}
    return st


def speichere_state(pfad, st):
    st["last_run"] = datetime.now().isoformat(timespec="seconds")
    os.makedirs(os.path.dirname(os.path.abspath(pfad)), exist_ok=True)
    tmp = pfad + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(st, f, ensure_ascii=False, indent=2)
    os.replace(tmp, pfad)


# ==========================================================================
# CSV-Einlesen
# ==========================================================================

def _finde_spalte(header, kandidaten, fallback_idx):
    """Spaltenindex ueber den Kopfzeilen-Namen (exakt, dann case-insensitiv),
    sonst ueber den fixen Fallback-Index. (index_oder_None, methode)."""
    norm = [h.strip().lstrip("\ufeff") for h in header]
    for k in kandidaten:
        if k in norm:
            return norm.index(k), "Kopf '%s'" % k
    low = [h.lower() for h in norm]
    for k in kandidaten:
        if k.lower() in low:
            return low.index(k.lower()), "Kopf '%s' (case-insensitiv)" % k
    if 0 <= fallback_idx < len(header):
        return fallback_idx, "Fallback-Spalte %d" % (fallback_idx + 1)
    return None, "nicht gefunden"


def lies_carrier_csv(pfad, carrier, log):
    """Liest eine Carrier-CSV -> (Liste (rgnr, tracking), ok). Leere Werte
    werden verworfen. Tracking bleibt IMMER ein String (DPD-Paketnummer hat
    eine fuehrende 0).

    ok=False nur bei einem ECHTEN Problem (Datei nicht lesbar oder Spalten
    nicht gefunden) - so eine Datei wird von archiviere_carrier_dateien()
    NIE automatisch archiviert, sondern bleibt fuer einen manuellen Blick
    liegen. Eine Datei ohne verwertbare Zeilen (z.B. leer) gilt dagegen als
    ok=True - die darf ganz normal verfallen."""
    try:
        with open(pfad, "r", encoding=carrier["encoding"], newline="") as f:
            rows = list(csv.reader(f, delimiter=";"))
    except Exception as e:
        log("  FEHLER beim Lesen von %s: %s" % (os.path.basename(pfad), e))
        return [], False
    if len(rows) < 2:
        return [], True
    header = rows[0]
    ref_i, ref_m = _finde_spalte(header, carrier["ref_spalten"], carrier["ref_fallback_idx"])
    trk_i, trk_m = _finde_spalte(header, carrier["track_spalten"], carrier["track_fallback_idx"])
    if ref_i is None or trk_i is None:
        log("  FEHLER %s: Spalte nicht gefunden (Rg=%s, Tracking=%s) - Datei uebersprungen."
            % (os.path.basename(pfad), ref_m, trk_m))
        return [], False
    log("  %s [%s]: Rg-Nr via %s, Tracking via %s, %d Datenzeile(n)"
        % (os.path.basename(pfad), carrier["name"], ref_m, trk_m, len(rows) - 1))
    out = []
    for r in rows[1:]:
        if len(r) <= max(ref_i, trk_i):
            continue
        ref = r[ref_i].strip()
        trk = r[trk_i].strip()
        if not ref or not trk:
            continue
        # Rechnungsnummern sind bei Gasecenter reine Ziffern (7-stellig, 170xxxx).
        if not ref.isdigit():
            continue
        out.append((ref, trk))
    return out, True


def lies_bestellnummern_csv(pfad, log):
    """wc_bestellnummern.csv -> dict rgnr -> {bestellnr,email,name}."""
    m = {}
    if not os.path.exists(pfad):
        log("  HINWEIS: %s nicht im Eingangsordner - es zaehlt nur die kumulative "
            "Zuordnung aus der Statusdatei." % os.path.basename(pfad))
        return m
    try:
        with open(pfad, "r", encoding="utf-8-sig", newline="") as f:
            rd = csv.reader(f, delimiter=";")
            next(rd, None)  # Kopf
            for t in rd:
                if len(t) < 2 or not t[0].strip() or not t[1].strip():
                    continue
                m[t[0].strip()] = {
                    "bestellnr": t[1].strip(),
                    "email": (t[2].strip() if len(t) > 2 else ""),
                    "name": (t[3].strip() if len(t) > 3 else ""),
                }
    except Exception as e:
        log("  FEHLER beim Lesen von %s: %s" % (os.path.basename(pfad), e))
    log("  %s: %d Rechnungsnummer->Bestellnummer-Zuordnung(en)"
        % (os.path.basename(pfad), len(m)))
    return m


def sammle_eingang(eingang_dir, log):
    """Liest den Eingangsordner. Rueckgabe:
       mapping_neu : dict rgnr -> {bestellnr,email,name}
       sendungen   : dict rgnr -> list[(tracking, carrier_name)]  (dedupliziert, Reihenfolge stabil)
       carrier_info: dict dateipfad -> {"rgnr": set(...), "ok": bool, "carrier": name}
                     (fuer archiviere_carrier_dateien() - welche Rechnungsnummern
                     kamen aus welcher Datei, und ob sie sauber gelesen wurde)
    """
    mapping_neu = lies_bestellnummern_csv(
        os.path.join(eingang_dir, "wc_bestellnummern.csv"), log)

    sendungen = {}
    carrier_info = {}
    gesehen_dateien = set()
    for carrier in CARRIERS:
        for pfad in sorted(glob.glob(os.path.join(eingang_dir, carrier["muster"]))):
            key = os.path.basename(pfad).lower()
            if key in gesehen_dateien:
                continue
            gesehen_dateien.add(key)
            paare, ok = lies_carrier_csv(pfad, carrier, log)
            rgnr_set = set()
            for ref, trk in paare:
                rgnr_set.add(ref)
                liste = sendungen.setdefault(ref, [])
                if not any(t == trk for t, _ in liste):
                    liste.append((trk, carrier["name"]))
            carrier_info[pfad] = {"rgnr": rgnr_set, "ok": ok, "carrier": carrier["name"]}
    return mapping_neu, sendungen, carrier_info


# ==========================================================================
# REST-API
# ==========================================================================

class Api:
    def __init__(self, cfg):
        self.base = cfg["base_url"] + "/wp-json/wc/v3"
        tok = base64.b64encode(
            ("%s:%s" % (cfg["consumer_key"], cfg["consumer_secret"])).encode("utf-8")
        ).decode("ascii")
        self.auth = "Basic " + tok

    def _call(self, method, pfad, payload=None):
        url = self.base + pfad
        data = json.dumps(payload).encode("utf-8") if payload is not None else None
        req = urllib.request.Request(url, data=data, method=method)
        req.add_header("Authorization", self.auth)
        req.add_header("Accept", "application/json")
        if data is not None:
            req.add_header("Content-Type", "application/json")
        for versuch in range(1, 4):
            try:
                with urllib.request.urlopen(req, timeout=30) as resp:
                    body = resp.read().decode("utf-8")
                    return resp.getcode(), (json.loads(body) if body else None)
            except urllib.error.HTTPError as e:
                raw = e.read().decode("utf-8", "replace")
                try:
                    parsed = json.loads(raw)
                except Exception:
                    parsed = {"raw": raw[:300]}
                return e.code, parsed
            except (urllib.error.URLError, OSError) as e:
                if versuch == 3:
                    return None, {"fehler": str(e)}
                time.sleep(2 * versuch)
        return None, {"fehler": "unerreichbar"}

    def order_existiert(self, oid):
        st, body = self._call("GET", "/orders/%s" % oid)
        if st == 200 and isinstance(body, dict) and body.get("id"):
            return body
        return None

    def shipments_der_order(self, oid):
        st, body = self._call("GET", "/shipments?order_id=%s" % oid)
        if st != 200 or not isinstance(body, list):
            return None
        out = []
        for el in body:
            if isinstance(el, dict) and isinstance(el.get("data"), dict):
                out.append(el["data"])
            elif isinstance(el, dict):
                out.append(el)
        return out

    def shipment_put(self, sid, payload):
        return self._call("PUT", "/shipments/%s" % sid, payload)

    def shipment_post(self, payload):
        return self._call("POST", "/shipments", payload)


# ==========================================================================
# Kernlogik
# ==========================================================================

def _kurz(body):
    if isinstance(body, dict):
        return str(body.get("message") or body.get("fehler") or body.get("raw") or body)[:200]
    return str(body)[:200]


def _payload(trk, slug, carrier_name, cfg):
    p = {"tracking_id": trk, "shipping_provider": slug,
         "shipping_provider_title": carrier_name}
    if cfg["set_status_shipped"]:
        p["status"] = "shipped"
    return p


def _merke_sync(state, rgnr, bestellnr, trackings, carrier_name, slug):
    eintrag = state["synced"].get(rgnr, {})
    alle = list(dict.fromkeys((eintrag.get("trackings") or []) + list(trackings)))
    state["synced"][rgnr] = {
        "order_id": bestellnr, "trackings": alle,
        "carrier": carrier_name, "provider": slug,
        "ts": datetime.now().isoformat(timespec="seconds"),
    }


def _offen(state, rgnr, grund, bestellnr, carrier_name, trackings):
    state["offen"][rgnr] = {
        "grund": grund, "bestellnr": bestellnr, "carrier": carrier_name,
        "tracking": list(trackings),
        "zuletzt": datetime.now().isoformat(timespec="seconds"),
    }


def verarbeite(rgnr, bestellnr, trackings, carrier_name, api, cfg, state, dry, log):
    """Traegt die Sendungsnummer(n) einer Bestellung ein.
    Rueckgabe: 'sync' | 'schon' | 'konflikt' | 'order_fehlt' | 'api_fehler'."""
    slug = cfg["provider_slugs"].get(carrier_name, "")
    if not slug:
        log("  %s: kein Provider-Slug fuer '%s' in der Konfig - uebersprungen." % (rgnr, carrier_name))
        _offen(state, rgnr, "kein Provider-Slug", bestellnr, carrier_name, trackings)
        return "konflikt"

    order = api.order_existiert(bestellnr)
    if not order:
        log("  %s: Bestellnummer %s existiert nicht im Shop - uebersprungen "
            "(vermutlich keine echte Onlineshop-Bestellung)." % (rgnr, bestellnr))
        _offen(state, rgnr, "Bestellnummer nicht im Shop", bestellnr, carrier_name, trackings)
        return "order_fehlt"

    vorhandene = api.shipments_der_order(bestellnr)
    if vorhandene is None:
        log("  %s: GET /shipments?order_id=%s fehlgeschlagen - spaeter erneut." % (rgnr, bestellnr))
        return "api_fehler"

    belegt = {(s.get("tracking_id") or "").strip(): s
              for s in vorhandene if (s.get("tracking_id") or "").strip()}
    frei = [s for s in vorhandene if not (s.get("tracking_id") or "").strip()]

    offene_trackings = [t for t in trackings if t not in belegt]
    if not offene_trackings:
        log("  %s: alle %d Sendungsnummer(n) bereits im Shop - nichts zu tun." % (rgnr, len(trackings)))
        if not dry:
            _merke_sync(state, rgnr, bestellnr, trackings, carrier_name, slug)
            state["offen"].pop(rgnr, None)
        return "schon"

    # Konflikt-Schutz: es gibt schon Sendung(en) MIT einer fremden Nummer, aber
    # KEINE freie Sendung fuer unsere Nummer. Dann nicht blind eine zweite
    # Sendung anlegen - melden und den Menschen entscheiden lassen.
    if belegt and not frei and len(vorhandene) >= len(trackings):
        fremd = ", ".join(sorted(belegt.keys()))
        log("  %s: Bestellung %s hat bereits Sendung(en) mit anderer Nummer (%s) und keine "
            "freie Sendung - NICHT ueberschrieben/ergaenzt." % (rgnr, bestellnr, fremd))
        _offen(state, rgnr, "Konflikt: fremde Nummer %s, keine freie Sendung" % fremd,
               bestellnr, carrier_name, trackings)
        return "konflikt"

    ergebnis = "sync"
    geschrieben = []
    for trk in offene_trackings:
        ziel = frei.pop(0) if frei else None
        if ziel is not None:
            sid = ziel.get("id")
            if dry:
                log("  [TROCKEN] %s: PUT /shipments/%s  tracking_id='%s' provider='%s'%s"
                    % (rgnr, sid, trk, slug, "  status='shipped'" if cfg["set_status_shipped"] else ""))
                geschrieben.append(trk)
                continue
            st, body = api.shipment_put(sid, _payload(trk, slug, carrier_name, cfg))
            if st in (200, 201):
                log("  %s: Sendung %s aktualisiert -> '%s' (%s/%s)" % (rgnr, sid, trk, carrier_name, slug))
                geschrieben.append(trk)
            else:
                log("  %s: PUT /shipments/%s fehlgeschlagen HTTP %s: %s" % (rgnr, sid, st, _kurz(body)))
                ergebnis = "api_fehler"
        else:
            if dry:
                log("  [TROCKEN] %s: POST /shipments  order_id=%s tracking_id='%s' provider='%s'%s"
                    % (rgnr, bestellnr, trk, slug, "  status='shipped'" if cfg["set_status_shipped"] else ""))
                geschrieben.append(trk)
                continue
            pl = _payload(trk, slug, carrier_name, cfg)
            pl["order_id"] = int(bestellnr) if str(bestellnr).isdigit() else bestellnr
            st, body = api.shipment_post(pl)
            if st in (200, 201):
                neu_id = body.get("id") if isinstance(body, dict) else "?"
                log("  %s: neue Sendung %s angelegt -> '%s' (%s/%s)" % (rgnr, neu_id, trk, carrier_name, slug))
                geschrieben.append(trk)
            else:
                log("  %s: POST /shipments fehlgeschlagen HTTP %s: %s" % (rgnr, st, _kurz(body)))
                ergebnis = "api_fehler"

    if geschrieben and ergebnis != "api_fehler" and not dry:
        # nur die tatsaechlich geschriebenen Nummern vermerken; bereits zuvor
        # vorhandene (belegt) muessen nicht erneut in den State.
        _merke_sync(state, rgnr, bestellnr, geschrieben, carrier_name, slug)
        state["offen"].pop(rgnr, None)
    return ergebnis


# ==========================================================================
# Archivierung
# ==========================================================================

def _eindeutiger_archiv_pfad(ordner, name):
    """Zielpfad in `ordner`, bei Namenskollision mit Zaehler eindeutig
    gemacht (kaeme in der Praxis so gut wie nie vor, da PaketImport-Mover.ps1
    schon beim Kopieren nach Sendungsnummern_WC eindeutige Namen vergibt)."""
    ziel = os.path.join(ordner, name)
    if not os.path.exists(ziel):
        return ziel
    stamm, ext = os.path.splitext(name)
    i = 1
    while True:
        kandidat = os.path.join(ordner, "%s_%d%s" % (stamm, i, ext))
        if not os.path.exists(kandidat):
            return kandidat
        i += 1


def archiviere_carrier_dateien(eingang_dir, carrier_info, ergebnis_je_rgnr, archive_after_days, log):
    """Verschiebt vollstaendig abgearbeitete, genuegend alte Carrier-CSVs nach
    <eingang_dir>\\Archiv\\<heutiges Datum>\\ - sonst waechst der Eingangs-
    ordner unbegrenzt weiter und jeder Lauf liest jede Datei fuer immer erneut
    (siehe Docstring-Punkt 5 oben). Rueckgabe: Anzahl verschobener Dateien."""
    heute = datetime.now()
    verschoben = 0
    for pfad, info in sorted(carrier_info.items()):
        name = os.path.basename(pfad)
        if not info["ok"]:
            continue  # Parse-Fehler - liegen lassen, braucht einen manuellen Blick
        if not os.path.exists(pfad):
            continue  # zwischenzeitlich schon anderweitig verschwunden

        wartet_auf_retry = any(ergebnis_je_rgnr.get(r) == "api_fehler" for r in info["rgnr"])
        if wartet_auf_retry:
            log("  %s: noch nicht archiviert - mind. eine Sendungsnummer wartet auf einen "
                "erneuten API-Versuch." % name)
            continue

        alter_tage = (heute - datetime.fromtimestamp(os.path.getmtime(pfad))).days
        if alter_tage < archive_after_days:
            continue

        ziel_ordner = os.path.join(eingang_dir, "Archiv", heute.strftime("%Y-%m-%d"))
        try:
            os.makedirs(ziel_ordner, exist_ok=True)
            ziel = _eindeutiger_archiv_pfad(ziel_ordner, name)
            shutil.move(pfad, ziel)
            verschoben += 1
            log("  %s: archiviert (%d Tage alt) -> %s" % (name, alter_tage, ziel))
        except Exception as e:
            log("  %s: FEHLER beim Archivieren: %s" % (name, e))
    return verschoben


# ==========================================================================
# main
# ==========================================================================

def main(argv):
    ap = argparse.ArgumentParser(description="Carrier-Sendungsnummern in WooCommerce/Shiptastic eintragen.")
    ap.add_argument("eingang", nargs="?", help="Eingangsordner (sonst aus Konfig / Standard)")
    ap.add_argument("--config", default=os.path.join(SKRIPT_DIR, "wc_sync_config.json"))
    ap.add_argument("--state", default=None)
    ap.add_argument("--commit", action="store_true", help="tatsaechlich schreiben (sonst Trockenlauf)")
    ap.add_argument("--only", help="nur diese Rechnungsnummer")
    ap.add_argument("--limit", type=int, default=0, help="hoechstens N Bestellungen schreiben")
    args = ap.parse_args(argv)

    cfg = lade_konfig(args.config)
    eingang = args.eingang or cfg.get("input_dir") or DEFAULT_INPUT_DIR
    state_pfad = args.state or cfg.get("state_path") or os.path.join(SKRIPT_DIR, "wc_sync_state.json")
    dry = not args.commit

    def log(msg):
        print(msg)

    log("=" * 70)
    log("wc_sendungsnummer_sync  %s  %s"
        % (datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
           "TROCKENLAUF (kein --commit)" if dry else "COMMIT - schreibt in den Shop"))
    log("Eingangsordner: %s" % eingang)
    log("Statusdatei:    %s" % state_pfad)
    if not os.path.isdir(eingang):
        sys.exit("FEHLER: Eingangsordner nicht erreichbar: %s" % eingang)

    state = lade_state(state_pfad)
    log("Letzter Lauf:   %s" % state.get("last_run", "(keiner)"))

    log("-" * 70)
    log("Eingangsdateien:")
    mapping_neu, sendungen, carrier_info = sammle_eingang(eingang, log)

    neu_map = 0
    for rgnr, info in mapping_neu.items():
        if state["mapping"].get(rgnr) != info:
            state["mapping"][rgnr] = info
            neu_map += 1
    log("Zuordnung: %d gesamt in der Statusdatei (%d neu/geaendert aus wc_bestellnummern.csv)"
        % (len(state["mapping"]), neu_map))

    zaehler = {"sync": 0, "schon": 0, "konflikt": 0, "order_fehlt": 0,
               "api_fehler": 0, "keine_zuordnung": 0, "uebersprungen_state": 0}
    # Ergebnis je Rechnungsnummer, unabhaengig vom zaehler-dict - archiviere_
    # carrier_dateien() braucht das, um eine Datei mit einer noch offenen
    # "api_fehler"-Sendungsnummer von der Archivierung auszunehmen.
    ergebnis_je_rgnr = {}

    if not sendungen:
        log("Keine Sendungsnummern in den Carrier-CSVs gefunden.")
    else:
        api = Api(cfg)
        log("-" * 70)
        geschrieben = 0

        for rgnr in sorted(sendungen):
            if args.only and rgnr != args.only:
                continue
            trackings_carrier = sendungen[rgnr]
            trackings = [t for t, _ in trackings_carrier]
            carrier_name = trackings_carrier[0][1]
            andere = sorted({c for _, c in trackings_carrier if c != carrier_name})
            if andere:
                log("  %s: HINWEIS - Sendungsnummern von mehreren Carriern (%s + %s); "
                    "alle werden mit Provider '%s' eingetragen."
                    % (rgnr, carrier_name, ", ".join(andere), carrier_name))

            vorher = state["synced"].get(rgnr)
            if vorher and set(trackings).issubset(set(vorher.get("trackings") or [])):
                zaehler["uebersprungen_state"] += 1
                ergebnis_je_rgnr[rgnr] = "uebersprungen_state"
                continue

            info = state["mapping"].get(rgnr)
            if not info:
                zaehler["keine_zuordnung"] += 1
                ergebnis_je_rgnr[rgnr] = "keine_zuordnung"
                _offen(state, rgnr, "keine WC-Bestellnummer bekannt", "", carrier_name, trackings)
                log("  %s: keine WC-Bestellnummer bekannt (weder wc_bestellnummern.csv noch "
                    "Statusdatei) - Carrier %s, Tracking %s" % (rgnr, carrier_name, ", ".join(trackings)))
                continue

            if args.limit and geschrieben >= args.limit:
                log("  %s: --limit %d erreicht - Rest uebersprungen." % (rgnr, args.limit))
                break

            res = verarbeite(rgnr, info["bestellnr"], trackings, carrier_name,
                             api, cfg, state, dry, log)
            zaehler[res] = zaehler.get(res, 0) + 1
            ergebnis_je_rgnr[rgnr] = res
            if res == "sync":
                geschrieben += 1
            time.sleep(0.3)

    speichere_state(state_pfad, state)

    # Archivierung: nur bei einem vollstaendigen, echten Lauf (kein Trockenlauf,
    # kein --only/--limit) - siehe Docstring-Punkt 5 und archiviere_carrier_
    # dateien() oben.
    if not dry and not args.only and not args.limit and carrier_info:
        log("-" * 70)
        log("Archivierung (aelter als %d Tage, vollstaendig verarbeitet):" % cfg["archive_after_days"])
        n = archiviere_carrier_dateien(eingang, carrier_info, ergebnis_je_rgnr,
                                       cfg["archive_after_days"], log)
        if not n:
            log("  keine faellig.")

    log("-" * 70)
    log("ZUSAMMENFASSUNG%s" % ("  (Trockenlauf - nichts geschrieben)" if dry else ""))
    log("  eingetragen/aktualisiert     : %d" % zaehler["sync"])
    log("  bereits im Shop              : %d" % zaehler["schon"])
    log("  schon laut Statusdatei       : %d" % zaehler["uebersprungen_state"])
    log("  keine WC-Bestellnummer       : %d" % zaehler["keine_zuordnung"])
    log("  Bestellnummer nicht im Shop  : %d" % zaehler["order_fehlt"])
    log("  Konflikt (nicht ueberschrieben): %d" % zaehler["konflikt"])
    log("  API-Fehler (spaeter erneut)  : %d" % zaehler["api_fehler"])
    if state["offen"]:
        log("  --> %d offene Rechnungsnummer(n) in der Statusdatei (Feld 'offen')." % len(state["offen"]))
    if dry:
        log("  Mit --commit erneut ausfuehren, um die Aenderungen wirklich zu schreiben.")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
