# -*- coding: utf-8 -*-
"""
wc_storno.py - Sperre fuer stornierte Onlineshop-Bestellungen vor der Pickliste
================================================================================
Hintergrund (2026-11-16): wc_auftragsimport.py uebergibt bezahlte Shop-Bestellungen
an Amicron. Wird eine Bestellung danach im Shop storniert, kann der Dateiimport den
Amicron-Auftrag NICHT mehr aendern (Datei mit vorhandener Nummer wird still
uebersprungen, getestet 2026-09-29) - es kommt nur ein Warnauftrag <nr>-ST. Der
Originalauftrag kann also trotzdem fakturiert werden. Damit daraus nie ein Paket
wird, prueft das Carrier-Dashboard (und packliste.py) jede Onlineshop-Rechnung hier,
bevor sie auf die Pickliste kommt.

Quellen, in dieser Reihenfolge:
  1. Statusdatei von wc_auftragsimport.py (wc_auftragsimport_state.json): enthaelt
     jede Bestellung, fuer die schon ein Warnauftrag geschrieben wurde. Lokal, ohne
     Netz - greift auch, wenn der Shop gerade nicht erreichbar ist.
  2. Live-Abfrage GET /wc/v3/orders/<nr> (HTTPS, Lese-Schluessel aus
     wc_auftragsimport_config.json) - faengt auch eine Stornierung von vor fuenf
     Minuten ab, die der Auftragsimport noch nicht gesehen hat.

Gesperrt wird bei: storniert/erstattet/fehlgeschlagen/widerrufen/geloescht, wieder
unbezahlt (on-hold/pending) und jedem UNBEKANNTEN Status. Nicht gesperrt: processing,
completed, Teilerstattung (die Ware geht ja trotzdem raus).

Gegenprobe gegen Kd-Nr-Kollision: die 4-6-stellige Kd-Nr auf der Rechnung ist bei
Shop-Bestellungen die Shop-Bestellnummer - bei anderen Kunden koennte sie zufaellig
eine fremde Shop-Bestellung treffen. Gesperrt wird nur, wenn E-Mail ODER PLZ der
Rechnung zur Shop-Bestellung passen (oder die Bestellung laut Statusdatei vom
Auftragsimport kam).

Ist der Shop nicht erreichbar, wird NICHT gesperrt (der Versand soll nicht am Shop
haengen), aber die Rechnung landet in "nicht_geprueft" - das Dashboard fragt dann
nach.
"""

import base64
import json
import os
import time
import urllib.error
import urllib.request

VERSION = "2026-11-16a"

STANDARD_CONFIG = r"C:\Scripts\wc_auftragsimport_config.json"

OK_STATUS = {"processing", "completed"}
STATUS_TEXT = {
    "cancelled": "storniert", "refunded": "erstattet", "failed": "fehlgeschlagen",
    "withdrawn": "widerrufen", "trash": "geloescht", "geloescht": "geloescht",
    "on-hold": "wieder unbezahlt (on-hold)", "pending": "unbezahlt (pending)",
    "checkout-draft": "unbezahlt (Entwurf)", "draft": "unbezahlt (Entwurf)",
}
TIMEOUT_S = 8
OFFLINE_PAUSE_S = 60       # nach einem Netzfehler so lange keine weiteren Abrufe

_offline_bis = [0.0]


def _ist_shop_rechnung(r):
    import packliste
    return packliste._ist_onlineshop_bestellung(r.get("kdnr"), r.get("email"))


class StornoPruefung:
    def __init__(self, config_pfad=None):
        self.config_pfad = config_pfad or STANDARD_CONFIG
        self.fehler = ""
        self.cfg = None
        try:
            with open(self.config_pfad, "r", encoding="utf-8-sig") as f:
                cfg = json.load(f)
            if not str(cfg.get("base_url", "")).lower().startswith("https://"):
                raise ValueError("base_url fehlt oder ist kein https://")
            if not cfg.get("consumer_key") or not cfg.get("consumer_secret"):
                raise ValueError("Schluessel fehlt")
            self.cfg = cfg
        except (OSError, ValueError) as e:
            self.fehler = "%s: %s" % (self.config_pfad, e)

    @property
    def verfuegbar(self):
        return self.cfg is not None

    # -- Quellen -------------------------------------------------------------

    def _warnungen_aus_state(self):
        """{bestellnr: grund} aller Bestellungen mit Warnauftrag laut Statusdatei
        des Auftragsimports, plus Menge aller von dort uebergebenen Nummern."""
        warn, bekannt = {}, set()
        pfad = (self.cfg or {}).get("state_path") or r"C:\Scripts\wc_auftragsimport_state.json"
        try:
            with open(pfad, "r", encoding="utf-8") as f:
                st = json.load(f)
        except (OSError, ValueError):
            return warn, bekannt
        for oid, e in (st.get("exportiert") or {}).items():
            nr = str(e.get("nummer") or oid)
            bekannt.add(nr)
            erl = str(e.get("erledigt") or "")
            if erl.startswith("warnung:"):
                status = erl.split(":", 1)[1]
                warn[nr] = "im Shop %s (Warnauftrag %s-ST)" % (STATUS_TEXT.get(status, status), nr)
        return warn, bekannt

    def _hole(self, nr):
        """-> ("ok", order) | ("fehlt", None) | ("netz", fehlertext)"""
        if time.monotonic() < _offline_bis[0]:
            return "netz", "Shop vor kurzem nicht erreichbar"
        tok = base64.b64encode(("%s:%s" % (self.cfg["consumer_key"], self.cfg["consumer_secret"]))
                               .encode("utf-8")).decode("ascii")
        req = urllib.request.Request("%s/wp-json/wc/v3/orders/%s" % (self.cfg["base_url"].rstrip("/"), nr))
        req.add_header("Authorization", "Basic " + tok)
        req.add_header("Accept", "application/json")
        try:
            with urllib.request.urlopen(req, timeout=TIMEOUT_S) as resp:
                body = json.loads(resp.read().decode("utf-8-sig") or "null")
            if isinstance(body, dict) and body.get("id"):
                return "ok", body
            return "netz", "unerwartete Antwort"
        except urllib.error.HTTPError as e:
            if e.code == 404:
                return "fehlt", None
            return "netz", "HTTP %s" % e.code
        except (urllib.error.URLError, OSError, ValueError) as e:
            _offline_bis[0] = time.monotonic() + OFFLINE_PAUSE_S
            return "netz", str(e)

    # -- Pruefung ------------------------------------------------------------

    def pruefe(self, rechnungen, live=True):
        """-> (gesperrt, nicht_geprueft)
        live=False: nur die Statusdatei (schnell, ohne Netz - fuer Schritt 1).
        live=True:  ist der Shop erreichbar, entscheidet der LIVE-Status (auch gegen
                    einen alten Warnvermerk der Statusdatei - eine wieder auf
                    "processing" gesetzte Bestellung wird so wieder frei); nur wenn der
                    Shop nicht antwortet, sperrt der Warnvermerk.
        gesperrt:       {rnr: Meldungstext}  - NICHT packen/versenden
        nicht_geprueft: {rnr: Grund}         - Shop-Rechnung, Live-Abfrage scheiterte
        self.live_frei: Menge der rnr, die live geprueft und in Ordnung sind.
        Rechnungen ohne Shop-Bestellnummer werden nicht angefasst."""
        gesperrt, nicht_geprueft = {}, {}
        self.live_frei = set()
        if not self.verfuegbar:
            if not live:
                return gesperrt, nicht_geprueft
            for r in rechnungen:
                if r and r.get("rnr") and _ist_shop_rechnung(r):
                    nicht_geprueft[r["rnr"]] = "Storno-Pruefung nicht eingerichtet (%s)" % self.fehler
            return gesperrt, nicht_geprueft
        warn, bekannt = self._warnungen_aus_state()
        for r in rechnungen:
            if not r or not r.get("rnr") or not _ist_shop_rechnung(r):
                continue
            rnr, nr = r["rnr"], str(r.get("kdnr") or "").strip()
            if not live:
                if nr in warn:
                    gesperrt[rnr] = self._text(nr, warn[nr])
                continue
            art, o = self._hole(nr)
            if art == "netz":
                if nr in warn:
                    gesperrt[rnr] = self._text(nr, warn[nr])
                else:
                    nicht_geprueft[rnr] = "Shop nicht erreichbar (%s)" % o
                continue
            if art == "fehlt":
                if nr in bekannt:      # kam vom Auftragsimport, ist jetzt weg
                    gesperrt[rnr] = self._text(nr, "im Shop geloescht")
                else:                  # sonst: Kd-Nr ist keine Shop-Bestellung
                    self.live_frei.add(rnr)
                continue
            if not self._passt(r, o) and nr not in bekannt:
                self.live_frei.add(rnr)
                continue               # Kd-Nr-Kollision mit fremder Bestellung
            status = o.get("status") or ""
            if status in OK_STATUS:
                self.live_frei.add(rnr)
                continue
            text = STATUS_TEXT.get(status, "mit UNBEKANNTEM Status '%s'" % status)
            gesperrt[rnr] = self._text(nr, "im Shop %s" % text)
        return gesperrt, nicht_geprueft

    @staticmethod
    def _passt(r, o):
        mail = (r.get("email") or "").strip().lower()
        plz = (r.get("plz") or "").strip()
        b, s = o.get("billing") or {}, o.get("shipping") or {}
        if mail and mail == (b.get("email") or "").strip().lower():
            return True
        if plz and plz in {(b.get("postcode") or "").strip(), (s.get("postcode") or "").strip()}:
            return True
        return not mail and not plz      # nichts zum Vergleichen -> im Zweifel sperren

    @staticmethod
    def _text(nr, grund):
        return ("STORNIERT: Shop-Bestellung %s %s - NICHT versenden! Rechnung per Gutschrift "
                "ausbuchen und PDF aus dem Pool entfernen." % (nr, grund))
