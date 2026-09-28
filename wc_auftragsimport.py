# -*- coding: utf-8 -*-
"""
wc_auftragsimport.py - WooCommerce-Bestellungen als XML fuer den Amicron-Auftragsimport
======================================================================================
Ersetzt den Shop-Bestellimport ueber AfterSell (af_wp_api.php auf api-works.de).
Gruende fuer den eigenen Weg:
  1. AfterSell importiert ALLE Bestellungen, auch unbezahlte (on-hold/pending).
     Hier werden nur Bestellungen mit Status "processing" (bezahlt, noch nicht
     versendet) uebergeben.
  2. Die Versandkosten kamen ohne Artikelnummer in Amicron an -> Fehler in der
     E-Rechnung. Hier werden sie als eigene Auftragsposition mit der
     Versandkosten-Artikelnummer (Standard 660) ausgegeben.
  3. af_wp_api.php laeuft nur ueber http:// - Kundendaten und Login gingen
     unverschluesselt ueber einen fremden Server. Dieses Skript spricht direkt
     per HTTPS mit dem Shop (WooCommerce-REST-API, Lese-Schluessel genuegt).

ABLAUF
------
1. GET /wc/v3/orders?status=processing (paginiert, orderby=id - siehe
   API-Eigenheiten unten).
2. Jede Bestellung, die laut Statusdatei noch nicht uebergeben wurde, wird als
   eigene XML-Datei WC_<Bestellnr>.xml in den Ausgabeordner geschrieben.
   Aufbau = Importdefinition "Woo-Shop-Script.XML" (liegt daneben).
3. In Amicron: Extras -> Datenimport -> Auftraege (Datei), Ordner = Ausgabeordner,
   Importdefinition = Woo-Shop-Script.XML.

Die Statusdatei merkt sich jede uebergebene Bestellung - eine Bestellung wird
nur EINMAL als Datei geschrieben, auch wenn sie noch tagelang "processing" ist.
Wird eine Datei in Amicron versehentlich nicht importiert, bleibt sie im
Ausgabeordner liegen; neu erzeugen: --only <nr> --neu.

UMSTIEG VON AFTERSELL
---------------------
Beim Umstieg sind die aktuell offenen processing-Bestellungen meist schon ueber
AfterSell in Amicron. Einmalig "--baseline --commit" ausfuehren: markiert alle
aktuell offenen Bestellungen als uebergeben, OHNE Dateien zu schreiben. Danach
kommen nur noch neue Bestellungen.

API-EIGENHEITEN DIESES SHOPS (siehe Skill woocommerce-vorkasse-datev-export)
---------------------------------------------------------------------------
- kein after=, kein _fields (URL-Laenge), order=asc unzuverlaessig
- ohne orderby=id ist die Seitenreihenfolge instabil -> immer orderby=id&order=desc
- per_page=10 hat sich bewaehrt
- eine leere Antwort ist nicht zwingend "keine Treffer" -> Abgleich mit X-WP-Total,
  bei Abweichung einmal mit leicht veraendertem Query wiederholen

AUFRUF
------
  python wc_auftragsimport.py [Optionen]
    --config PATH   Konfig-JSON (Standard: wc_auftragsimport_config.json neben dem Skript)
    --commit        Dateien schreiben + Statusdatei fortschreiben (sonst Trockenlauf)
    --only NR       nur diese Bestellnummer
    --neu           mit --only: Datei auch dann erneut schreiben, wenn schon uebergeben
    --baseline      alle aktuell offenen processing-Bestellungen als uebergeben markieren,
                    ohne Dateien zu schreiben (einmalig beim Umstieg, mit --commit)
    --vorschau DIR  XML-Dateien zur Ansicht nach DIR schreiben (Statusdatei bleibt unberuehrt)

KONFIG (wc_auftragsimport_config.json - Vorlage: wc_auftragsimport_config.example.json)
"""

import argparse
import base64
import json
import os
import sys
import time
import urllib.error
import urllib.request
from datetime import datetime
from xml.sax.saxutils import escape

SKRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

DEFAULTS = {
    "output_dir": r"C:\Scripts\WC_Auftragsimport",
    "state_path": r"C:\Scripts\wc_auftragsimport_state.json",
    "status": "processing",
    "versand_artikelnr": "660",
    "versand_bezeichnung": "Versandkosten",
    "bestellart": "Onlineshop",
    "per_page": 10,
    "max_seiten": 30,
}


# ==========================================================================
# Konfig / State
# ==========================================================================

def lade_konfig(pfad):
    if not os.path.exists(pfad):
        sys.exit("FEHLER: Konfigdatei nicht gefunden: %s\n"
                 "       Vorlage: wc_auftragsimport_config.example.json" % pfad)
    with open(pfad, "r", encoding="utf-8") as f:
        cfg = json.load(f)
    for pflicht in ("base_url", "consumer_key", "consumer_secret"):
        if not cfg.get(pflicht):
            sys.exit("FEHLER: '%s' fehlt in %s" % (pflicht, pfad))
    if not cfg["base_url"].lower().startswith("https://"):
        sys.exit("FEHLER: base_url muss mit https:// beginnen - der Schluessel darf nie "
                 "unverschluesselt uebertragen werden.")
    cfg["base_url"] = cfg["base_url"].rstrip("/")
    for k, v in DEFAULTS.items():
        cfg.setdefault(k, v)
    return cfg


def lade_state(pfad):
    try:
        with open(pfad, "r", encoding="utf-8") as f:
            st = json.load(f)
    except FileNotFoundError:
        st = {}
    except Exception as e:
        sys.exit("FEHLER: Statusdatei %s unlesbar (%s). "
                 "NICHT von Hand loeschen - sonst werden Bestellungen doppelt uebergeben." % (pfad, e))
    st.setdefault("exportiert", {})   # order_id -> {"nummer","datei","ts","kunde","total"}
    return st


def speichere_state(pfad, st):
    st["last_run"] = datetime.now().isoformat(timespec="seconds")
    os.makedirs(os.path.dirname(os.path.abspath(pfad)), exist_ok=True)
    tmp = pfad + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(st, f, ensure_ascii=False, indent=2)
    os.replace(tmp, pfad)


# ==========================================================================
# REST-API (nur GET)
# ==========================================================================

class Api:
    def __init__(self, cfg):
        self.base = cfg["base_url"] + "/wp-json/wc/v3"
        tok = base64.b64encode(
            ("%s:%s" % (cfg["consumer_key"], cfg["consumer_secret"])).encode("utf-8")
        ).decode("ascii")
        self.auth = "Basic " + tok

    def get(self, pfad):
        """-> (status, body, headers). GET ist idempotent, bis zu 3 Versuche."""
        req = urllib.request.Request(self.base + pfad, method="GET")
        req.add_header("Authorization", self.auth)
        req.add_header("Accept", "application/json")
        for versuch in range(1, 4):
            try:
                with urllib.request.urlopen(req, timeout=30) as resp:
                    raw = resp.read().decode("utf-8-sig")
                    return resp.getcode(), (json.loads(raw) if raw.strip() else None), resp.headers
            except urllib.error.HTTPError as e:
                return e.code, {"raw": e.read().decode("utf-8", "replace")[:300]}, e.headers
            except (urllib.error.URLError, OSError, ValueError) as e:
                if versuch == 3:
                    return None, {"fehler": str(e)}, {}
                time.sleep(2 * versuch)

    def _seite(self, status, seite, per_page, variante):
        # variante 1 = leicht veraenderter Query fuer den Wiederholungsversuch
        # bei einer verdaechtig leeren Antwort (bekannte Shop-Eigenheit).
        q = "/orders?status=%s&per_page=%d&page=%d&orderby=id&order=desc" % (status, per_page, seite)
        if variante:
            q = "/orders?status=%s&page=%d&per_page=%d&order=desc&orderby=id" % (status, seite, per_page)
        return self.get(q)

    def bestellungen(self, status, per_page, max_seiten, log):
        """Alle Bestellungen eines Status. Bricht mit sys.exit ab, wenn der
        Abruf nicht vollstaendig ist - lieber kein Lauf als ein stilles Loch."""
        out, gesamt, seite = [], None, 1
        while seite <= max_seiten:
            st, body, hdr = self._seite(status, seite, per_page, 0)
            if st == 200 and isinstance(body, list) and not body and gesamt != 0:
                st2, body2, hdr2 = self._seite(status, seite, per_page, 1)
                if st2 == 200 and isinstance(body2, list) and body2:
                    log("  Hinweis: Seite %d war leer, Wiederholung lieferte %d Bestellung(en)."
                        % (seite, len(body2)))
                    st, body, hdr = st2, body2, hdr2
            if st != 200 or not isinstance(body, list):
                sys.exit("FEHLER: Abruf Seite %d (status=%s) fehlgeschlagen: HTTP %s %s"
                         % (seite, status, st, str(body)[:200]))
            if gesamt is None:
                try:
                    gesamt = int(hdr.get("X-WP-Total"))
                except (TypeError, ValueError):
                    gesamt = None
            out.extend(o for o in body if isinstance(o, dict) and o.get("id"))
            if not body or len(body) < per_page:
                break
            if gesamt is not None and len(out) >= gesamt:
                break
            seite += 1
            time.sleep(0.3)
        # Duplikate (falls sich zwischen zwei Seiten etwas verschoben hat) entfernen
        eindeutig = {}
        for o in out:
            eindeutig[o["id"]] = o
        out = sorted(eindeutig.values(), key=lambda o: o["id"])
        if gesamt is not None and len(out) < gesamt:
            sys.exit("FEHLER: Shop meldet %d Bestellungen mit Status '%s', abgerufen wurden nur %d. "
                     "Abbruch, damit keine Bestellung verloren geht." % (gesamt, status, len(out)))
        return out


# ==========================================================================
# XML
# ==========================================================================

def _geld(wert):
    try:
        return "%.2f" % round(float(wert or 0), 2)
    except (TypeError, ValueError):
        return "0.00"


def _preis(wert):
    """Einzelpreis: 2 Nachkommastellen, bis zu 4 nur wenn noetig (Rabatt auf
    Menge > 1 kann krumme Einzelpreise ergeben)."""
    s = "%.4f" % wert
    return s[:-2] if s.endswith("00") else s.rstrip("0")


def _num(wert):
    try:
        return float(wert or 0)
    except (TypeError, ValueError):
        return 0.0


def _datum(iso):
    """'2026-09-26T12:58:03' -> ('2026-09-26 12:58:03', '26.09.2026')."""
    try:
        dt = datetime.fromisoformat(iso)
    except (TypeError, ValueError):
        return iso or "", iso or ""
    return dt.strftime("%Y-%m-%d %H:%M:%S"), dt.strftime("%d.%m.%Y")


def _brutto_einzelpreis(li):
    """Einzelpreis brutto nach Rabatt. WooCommerce fuehrt line 'total' immer
    netto (ohne Steuer) - deshalb total + total_tax, geteilt durch die Menge."""
    menge = _num(li.get("quantity")) or 1
    return (_num(li.get("total")) + _num(li.get("total_tax"))) / menge


def _t(tag, wert, einzug):
    return "%s<%s>%s</%s>" % (" " * einzug, tag, escape(str(wert if wert is not None else "")), tag)


def _adresse(tag, a, mit_kontakt, einzug):
    z = ["%s<%s>" % (" " * einzug, tag)]
    felder = ["first_name", "last_name", "company", "address_1", "address_2",
              "city", "postcode", "country"]
    if mit_kontakt:
        felder += ["phone", "email"]
    for f in felder:
        z.append(_t(f, (a or {}).get(f, ""), einzug + 4))
    if mit_kontakt:
        z.append(_t("birthday", "", einzug + 4))
    z.append("%s</%s>" % (" " * einzug, tag))
    return z


def bestellung_zu_xml(o, cfg):
    """Eine Bestellung -> XML-Text im Aufbau von Woo-Shop-Script.XML."""
    datum, datum_kurz = _datum(o.get("date_created"))
    nummer = str(o.get("number") or o["id"])
    z = ['<?xml version="1.0" encoding="iso-8859-1"?>',
         "<!-- erzeugt von wc_auftragsimport.py am %s - WooCommerce-Bestellung %s -->"
         % (datetime.now().strftime("%d.%m.%Y %H:%M"), escape(nummer)),
         "<orders>",
         "    <item>"]
    z.append(_t("number", nummer, 8))
    z.append(_t("date_created", datum, 8))
    z.append(_t("date_created_formatted", datum_kurz, 8))
    z.append(_t("freifeld3", cfg["bestellart"], 8))
    z.append(_t("freifeld4", "", 8))
    z += _adresse("shipping", o.get("shipping"), False, 8)
    z += _adresse("billing", o.get("billing"), True, 8)

    z.append("        <line_items>")
    for li in o.get("line_items") or []:
        z.append("            <item>")
        z.append(_t("sku", li.get("sku", ""), 16))
        z.append(_t("name", li.get("name", ""), 16))
        z.append(_t("discription", "", 16))
        z.append(_t("quantity", li.get("quantity", 0), 16))
        z.append(_t("price_incl_tax", _preis(_brutto_einzelpreis(li)), 16))
        z.append("            </item>")

    # Versandkosten als eigene Position MIT Artikelnummer (E-Rechnung verlangt
    # fuer jede Position eine Artikelnummer). Die Summe aller Versandzeilen,
    # brutto. Bei 0,00 keine Position.
    versand = sum(_num(s.get("total")) + _num(s.get("total_tax"))
                  for s in (o.get("shipping_lines") or []))
    if round(versand, 2) > 0:
        z.append("            <item>")
        z.append(_t("sku", cfg["versand_artikelnr"], 16))
        z.append(_t("name", cfg["versand_bezeichnung"], 16))
        z.append(_t("discription", "", 16))
        z.append(_t("quantity", 1, 16))
        z.append(_t("price_incl_tax", _geld(versand), 16))
        z.append("            </item>")
    z.append("        </line_items>")

    z.append(_t("payment_method_title", o.get("payment_method_title", ""), 8))
    lieferart = ", ".join(s.get("method_title", "") for s in (o.get("shipping_lines") or [])
                          if s.get("method_title"))
    z.append("        <shipping_lines>")
    z.append("            <item>")
    z.append(_t("method_title", lieferart, 16))
    z.append("            </item>")
    z.append("        </shipping_lines>")

    gebuehren = sum(_num(f.get("total")) + _num(f.get("total_tax"))
                    for f in (o.get("fee_lines") or []))
    z.append("        <fee_lines>")
    z.append("            <item>")
    z.append(_t("total", _geld(gebuehren) if round(gebuehren, 2) else "", 16))
    z.append("            </item>")
    z.append("        </fee_lines>")

    z.append(_t("customer_note", o.get("customer_note", ""), 8))
    z.append(_t("footer_note", "", 8))
    z.append(_t("total", _geld(o.get("total")), 8))
    z.append(_t("tax_flag", "J" if o.get("prices_include_tax", True) else "N", 8))
    z.append("    </item>")
    z.append("</orders>")
    return "\r\n".join(z) + "\r\n"


def pruefsumme(o):
    """Positionen + Versand + Gebuehren gegen den Bestell-Endbetrag."""
    summe = sum(_brutto_einzelpreis(li) * (_num(li.get("quantity")) or 1)
                for li in (o.get("line_items") or []))
    summe += sum(_num(s.get("total")) + _num(s.get("total_tax")) for s in (o.get("shipping_lines") or []))
    summe += sum(_num(f.get("total")) + _num(f.get("total_tax")) for f in (o.get("fee_lines") or []))
    return round(summe, 2), round(_num(o.get("total")), 2)


def schreibe_xml(ordner, o, cfg):
    os.makedirs(ordner, exist_ok=True)
    nummer = str(o.get("number") or o["id"])
    pfad = os.path.join(ordner, "WC_%s.xml" % nummer)
    daten = bestellung_zu_xml(o, cfg).encode("iso-8859-1", "xmlcharrefreplace")
    tmp = pfad + ".tmp"
    with open(tmp, "wb") as f:
        f.write(daten)
    os.replace(tmp, pfad)
    return pfad


def _kunde(o):
    b = o.get("billing") or {}
    return ("%s %s" % (b.get("first_name", ""), b.get("last_name", ""))).strip()


# ==========================================================================
# Hauptprogramm
# ==========================================================================

def main(argv):
    ap = argparse.ArgumentParser(description="WooCommerce-Bestellungen als XML fuer den Amicron-Auftragsimport.")
    ap.add_argument("--config", default=os.path.join(SKRIPT_DIR, "wc_auftragsimport_config.json"))
    ap.add_argument("--commit", action="store_true", help="Dateien schreiben (sonst Trockenlauf)")
    ap.add_argument("--only", help="nur diese Bestellnummer")
    ap.add_argument("--neu", action="store_true", help="mit --only: auch schon uebergebene erneut schreiben")
    ap.add_argument("--baseline", action="store_true",
                    help="offene Bestellungen nur als uebergeben markieren (Umstieg von AfterSell)")
    ap.add_argument("--vorschau", help="XML zur Ansicht in diesen Ordner schreiben, Statusdatei unberuehrt")
    args = ap.parse_args(argv)
    if args.neu and not args.only:
        sys.exit("FEHLER: --neu nur zusammen mit --only <Bestellnummer>.")
    if args.baseline and (args.only or args.vorschau):
        sys.exit("FEHLER: --baseline nicht mit --only/--vorschau kombinieren.")

    cfg = lade_konfig(args.config)
    dry = not args.commit

    def log(msg):
        print(msg)

    log("=" * 70)
    modus = "BASELINE (nur markieren)" if args.baseline else "Dateien schreiben"
    log("wc_auftragsimport  %s  %s"
        % (datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
           "TROCKENLAUF (kein --commit)" if dry else "COMMIT - " + modus))
    log("Ausgabeordner: %s" % cfg["output_dir"])
    log("Statusdatei:   %s" % cfg["state_path"])

    state = lade_state(cfg["state_path"])
    log("Letzter Lauf:  %s   (%d Bestellung(en) bisher uebergeben)"
        % (state.get("last_run", "(keiner)"), len(state["exportiert"])))

    api = Api(cfg)
    orders = api.bestellungen(cfg["status"], int(cfg["per_page"]), int(cfg["max_seiten"]), log)
    log("Shop: %d Bestellung(en) mit Status '%s'." % (len(orders), cfg["status"]))
    log("-" * 70)

    zaehler = {"neu": 0, "schon": 0, "abweichung": 0}
    for o in orders:
        oid = str(o["id"])
        nummer = str(o.get("number") or oid)
        if args.only and nummer != args.only and oid != args.only:
            continue
        vorher = state["exportiert"].get(oid)
        if vorher and not args.neu:
            zaehler["schon"] += 1
            continue

        summe, total = pruefsumme(o)
        warnung = ""
        if abs(summe - total) > 0.02:
            zaehler["abweichung"] += 1
            warnung = "  !! Positionssumme %.2f <> Endbetrag %.2f - in Amicron pruefen" % (summe, total)
        zeile = "  %s  %s  %-28s %8s EUR  %s" % (nummer, (o.get("date_created") or "")[:10],
                                                 _kunde(o)[:28], _geld(o.get("total")),
                                                 o.get("payment_method_title", ""))
        if args.vorschau:
            pfad = schreibe_xml(args.vorschau, o, cfg)
            log(zeile + "  -> Vorschau %s" % pfad + warnung)
            zaehler["neu"] += 1
            continue
        if dry:
            log(zeile + ("  (wuerde markiert)" if args.baseline else "  (wuerde geschrieben)") + warnung)
            zaehler["neu"] += 1
            continue

        datei = ""
        if not args.baseline:
            datei = schreibe_xml(cfg["output_dir"], o, cfg)
        state["exportiert"][oid] = {
            "nummer": nummer,
            "datei": os.path.basename(datei) if datei else "(baseline - kam ueber AfterSell)",
            "ts": datetime.now().isoformat(timespec="seconds"),
            "kunde": _kunde(o),
            "total": _geld(o.get("total")),
        }
        # nach JEDER Datei speichern - bricht der Lauf ab, wird nichts doppelt geschrieben
        speichere_state(cfg["state_path"], state)
        log(zeile + ("  markiert" if args.baseline else "  -> %s" % os.path.basename(datei)) + warnung)
        zaehler["neu"] += 1

    if args.only and not zaehler["neu"] and not zaehler["schon"]:
        log("  Bestellung %s ist nicht im Status '%s'." % (args.only, cfg["status"]))

    if not dry and not args.vorschau and not args.only:
        speichere_state(cfg["state_path"], state)

    log("-" * 70)
    log("ZUSAMMENFASSUNG%s" % ("  (Trockenlauf - nichts geschrieben)" if dry and not args.vorschau else ""))
    log("  %s: %d" % ("als uebergeben markiert" if args.baseline else
                      "Vorschau-Dateien" if args.vorschau else "neue Auftragsdateien", zaehler["neu"]))
    log("  schon frueher uebergeben : %d" % zaehler["schon"])
    if zaehler["abweichung"]:
        log("  Summen-Abweichungen      : %d  (siehe '!!' oben)" % zaehler["abweichung"])
    if dry and not args.vorschau:
        log("  Mit --commit erneut ausfuehren, um die Dateien wirklich zu schreiben.")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
