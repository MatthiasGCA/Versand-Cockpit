# -*- coding: utf-8 -*-
"""
rechnungsversand.py - taeglicher, SICHERER Rechnungsversand per E-Mail (ohne KI, reines Programm)
===========================================================================================
Versendet die Kundenrechnungen (PDF aus dem Ordner "Ausgangsrechnungen") gesammelt per E-Mail - aber NUR
an Onlineshop- und eBay-Kunden. Amazon-Kunden duerfen nie angeschrieben werden und werden deshalb auf
mehreren, voneinander unabhaengigen Wegen ausgeschlossen.

Ablauf (taeglich, z.B. 15:00 per Windows-Aufgabenplanung):
  1. Alle Rechnungs-PDFs im Rechnungsordner lesen (die Rechnungsnummer kommt aus dem INHALT, nicht aus dem
     - irrefuehrenden - Dateinamen). Gutschriften, Storno, Duplikate, Auftragsbestaetigungen u.a. werden ignoriert.
  2. Zu jeder Rechnungsnummer die zugehoerige PACKLISTE suchen (dort stehen Fremdbeleg und E-Mail-Adresse des
     Kunden; auf der Kundenrechnung steht keine Adresse).
  3. Klassifizieren (siehe klassifiziere()) - NUR POSITIV: gesendet wird ausschliesslich, wenn die Rechnung
     eindeutig als Onlineshop (4-6-stellige Shop-Nummer + normale Kundenadresse) oder eBay (eBay-Relay-Adresse
     @members.ebay.com + eBay-Name) erkannt wird UND nirgends ein Amazon-Merkmal vorkommt. Alles andere
     (unklar, keine Packliste, Zweifel) wird NICHT gesendet und im Bericht aufgelistet.
  4. Modus "bericht": nur Bericht schreiben. "entwurf": zusaetzlich fertige .eml-Dateien ablegen (in Outlook
     zu oeffnen). "senden": per SMTP versenden und protokollieren.

Absicherungen: Pflicht-Startnummer (ab_rechnungsnummer) gegen Versand alter Rechnungen, Maximalzahl Mails je
Lauf, Sperrdatei nie_senden.txt, Sendeprotokoll gegen Doppelversand (inkl. "Versuch ohne Bestaetigung" ->
nie blind wiederholen), Einzelinstanz-Sperre, bei JEDEM Zweifel/Fehler: nicht senden.

Aufruf:   py rechnungsversand.py [--config rechnungsversand_config.json] [--trocken]
Selbsttest: py rechnungsversand.py selftest
"""
import os
import re
import sys
import csv
import json
import glob
import time
import smtplib
import argparse
import hashlib
from datetime import datetime, timedelta
from email.message import EmailMessage
from email.utils import formatdate, make_msgid, formataddr

VERSION = "2026-10-22a"

RE_AMAZON_NR = re.compile(r"\b\d{3}-\d{7}-\d{7}\b")
RE_MAIL = re.compile(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}")
RE_RNR_KOPF = re.compile(r"^\s*Rechnung\s*Nr\.?\s*:?\s*(\d{6,8})\b", re.M)
FIRMEN_DOMAINS = ("gasecenter-augsburg.de", "gasecenter-onlineshop.de")
GESPERRTE_PRAEFIXE = ("111", "444")                 # Gutschriften / Auftragsbestaetigungen
# Belege, die KEINE Kundenrechnung sind (Textmarker im PDF)
KEINE_RECHNUNG = ("gutschrift", "storno", "wareneingang", "auftragsbestätigung", "auftragsbestaetigung",
                  "duplikat", "lieferschein", "angebot", "mahnung")
# interne Angaben, die NICHT auf einer Kundenrechnung stehen duerfen (sonst keine Mail)
INTERN_MARKER = re.compile(r"Lagerort\s*:|Sendungsgewicht\s*:|^\s*(pax1|pox1|brx1|wapo)\s*$", re.I | re.M)
EBAY_DOMAINS = ("members.ebay.com",)
MARKTPLATZ_WOERTER = ("amazon", "ebay", "marketplace", "members.")

KANAL_SHOP, KANAL_EBAY, KANAL_AMAZON, KANAL_UNKLAR = "SHOP", "EBAY", "AMAZON", "UNKLAR"


# ----------------------------------------------------------------------------------------------
# PDF lesen
# ----------------------------------------------------------------------------------------------
def pdf_text(pfad):
    """Gesamter Text aller Seiten (leer bei Fehler)."""
    import pdfplumber
    try:
        with pdfplumber.open(pfad) as pdf:
            return "\n".join((pg.extract_text() or "") for pg in pdf.pages)
    except Exception:
        return ""


def rechnungsnummer_aus_text(text):
    m = RE_RNR_KOPF.search(text or "")
    return m.group(1) if m else ""


def ist_kundenrechnung(text):
    """(True, "") wenn der Text eine normale Kundenrechnung ist; sonst (False, Grund)."""
    rnr = rechnungsnummer_aus_text(text)
    if not rnr:
        return False, "keine Rechnungsnummer im PDF"
    if rnr.startswith(GESPERRTE_PRAEFIXE):
        return False, "Gutschrift/Auftragsbestätigung (Nr. %s…)" % rnr[:3]
    kopf = (text or "")[:400].lower()
    for w in KEINE_RECHNUNG:
        if w in kopf:
            return False, "kein Rechnungsbeleg (%s)" % w
    return True, ""


def ist_firmenadresse(m):
    m = m.lower()
    return any(m.endswith("@" + d) or m.endswith("." + d) for d in FIRMEN_DOMAINS)


def fremdbeleg_und_mails(text):
    """Aus dem Kopf einer Rechnung/Packliste: (fremdbeleg, [mails]).
    Fremdbeleg = die erste Zeile NACH 'Rechnung Nr.', die weder Mail, Rechnungsnummer noch Feldname ist."""
    zeilen = (text or "").split("\n")
    fremd, mails = "", []
    idx = next((i for i, z in enumerate(zeilen) if re.match(r"^\s*Rechnung\s*Nr\.?", z)), None)
    if idx is not None:
        for z in zeilen[idx + 1: idx + 7]:
            z = z.strip()
            if not z:
                continue
            ms = RE_MAIL.findall(z)
            if ms:
                mails += [m for m in ms if not ist_firmenadresse(m)]
                continue
            if re.match(r"^(Sendungsgewicht|Kunden\s*Nr|Rechnungsdatum|Druckdatum|Gasecenter)", z, re.I):
                continue
            if not fremd:
                fremd = z
    return fremd, mails


def alle_kundenmails(text):
    """Alle E-Mail-Adressen im Text ausser den eigenen Firmenadressen (kleingeschrieben, ohne Dubletten)."""
    out = []
    for m in RE_MAIL.findall(text or ""):
        m = m.lower()
        if any(m.endswith("@" + d) or m.endswith("." + d) for d in FIRMEN_DOMAINS):
            continue
        if m not in out:
            out.append(m)
    return out


def lies_packliste(pfad):
    """Packliste -> dict(rnr, fremdbeleg, mails, alle_mails, amazon_signal, text) oder None."""
    text = pdf_text(pfad)
    rnr = rechnungsnummer_aus_text(text)
    if not rnr:
        return None
    fremd, mails = fremdbeleg_und_mails(text)
    return {"rnr": rnr, "fremdbeleg": fremd, "mails": [m.lower() for m in mails],
            "alle_mails": alle_kundenmails(text), "text": text, "pfad": pfad}


# ----------------------------------------------------------------------------------------------
# Klassifizierung - das Herzstueck. NUR POSITIVE Erkennung, jeder Zweifel -> UNKLAR.
# ----------------------------------------------------------------------------------------------
def amazon_signale(*texte):
    """Alle Amazon-Merkmale in den uebergebenen Texten (Liste von Kurzbeschreibungen)."""
    gef = []
    for t in texte:
        t = t or ""
        if RE_AMAZON_NR.search(t):
            gef.append("Amazon-Bestellnummer")
        if "amazon" in t.lower():
            gef.append("Wort 'Amazon'")
    return sorted(set(gef))


def klassifiziere(pack, rechnung_text):
    """(kanal, grund, empfaenger). kanal: SHOP | EBAY | AMAZON | UNKLAR. Nur SHOP und EBAY duerfen gesendet werden.

    Regeln (alle muessen erfuellt sein):
      Hart AMAZON:  irgendwo (Packliste, Rechnung, Adresse) Amazon-Bestellnummer, das Wort 'amazon' oder eine
                    Marketplace-Adresse -> AMAZON, nie senden.
      SHOP:  Fremdbeleg ist genau eine 4-6-stellige Zahl UND genau eine Kundenadresse UND die Adresse ist weder
             Marktplatz- noch Firmenadresse.
      EBAY:  genau eine Adresse @members.ebay.com UND Fremdbeleg ist ein Name (kein reiner Zahlen-/Amazon-Wert).
      Sonst: UNKLAR (Mehrdeutigkeit, fehlende/mehrere Adressen, fehlende Packliste, ...).
    """
    if not pack:
        return KANAL_UNKLAR, "keine zugehörige Packliste gefunden (E-Mail-Adresse unbekannt)", ""
    ptext, rtext = pack.get("text", ""), rechnung_text or ""
    sig = amazon_signale(ptext, rtext)
    if sig:
        return KANAL_AMAZON, "Amazon-Merkmal: " + ", ".join(sig), ""
    alle = pack.get("alle_mails") or []
    if any(any(w in m for w in ("amazon", "marketplace")) for m in alle):
        return KANAL_AMAZON, "Amazon-/Marketplace-Adresse", ""
    fremd = (pack.get("fremdbeleg") or "").strip()
    mails = pack.get("mails") or []
    if len(alle) != 1 or len(set(mails)) != 1 or mails[0] != alle[0]:
        return KANAL_UNKLAR, "keine eindeutige einzelne Kunden-E-Mail-Adresse (%d gefunden)" % len(alle), ""
    mail = alle[0]
    dom = mail.split("@", 1)[1]
    if not fremd:
        return KANAL_UNKLAR, "kein Fremdbeleg (Shop-Nummer/eBay-Name) erkannt", ""
    # Kundenrechnung darf ihren Fremdbeleg nur identisch zeigen
    r_fremd, _ = fremdbeleg_und_mails_rechnung(rtext)
    if r_fremd and r_fremd != fremd:
        return KANAL_UNKLAR, "Fremdbeleg auf Rechnung (%s) weicht von Packliste (%s) ab" % (r_fremd, fremd), ""
    if dom in EBAY_DOMAINS:
        if re.fullmatch(r"\d+", fremd):
            return KANAL_UNKLAR, "eBay-Adresse, aber Fremdbeleg ist eine Zahl", ""
        return KANAL_EBAY, "eBay-Relay-Adresse + eBay-Name '%s'" % fremd, mail
    if re.fullmatch(r"\d{4,6}", fremd):
        if any(w in dom for w in MARKTPLATZ_WOERTER):
            return KANAL_UNKLAR, "Shop-Nummer, aber Marktplatz-Adresse (%s)" % dom, ""
        return KANAL_SHOP, "Onlineshop-Nummer %s + Kundenadresse" % fremd, mail
    return KANAL_UNKLAR, "Fremdbeleg '%s' / Adresse '%s' passen zu keinem sicher erkannten Kanal" % (fremd, dom), ""


def fremdbeleg_und_mails_rechnung(text):
    """Fremdbeleg der KUNDENRECHNUNG (steht nach 'Rechnungsdatum:'); '' wenn nicht erkennbar."""
    zeilen = (text or "").split("\n")
    j = next((i for i, z in enumerate(zeilen) if z.strip().startswith("Rechnungsdatum:")), None)
    if j is None:
        return "", []
    for z in zeilen[j + 1: j + 4]:
        z = z.strip()
        if z.startswith("Rechnungsdatum =") or z.lower() == "rechnung":
            break                                   # Fremdbeleg-Zeile fehlt (Vorlagentext beginnt)
        if not z:
            continue
        return z, []
    return "", []


# ----------------------------------------------------------------------------------------------
# Hilfsfunktionen: Packlisten-Index, Protokoll, Sperren
# ----------------------------------------------------------------------------------------------
def rnr_aus_dateiname(name):
    m = re.search(r"Nr\.?\s*(\d{6,8})", name)
    return m.group(1) if m else ""


def suche_packlisten(ordnerliste, rnr_menge):
    """{rnr: [pfade]} - Packlisten per Dateiname (Nr.) in den Ordnern und deren Unterordnern (1 Ebene)."""
    gef = {}
    for o in ordnerliste:
        if not os.path.isdir(o):
            continue
        muster = [os.path.join(o, "*.pdf")] + [os.path.join(o, "*", "*.pdf")]
        for mu in muster:
            for p in glob.glob(mu):
                n = os.path.basename(p)
                if not n.lower().startswith("packliste"):
                    continue
                r = rnr_aus_dateiname(n)
                if r in rnr_menge:
                    gef.setdefault(r, []).append(p)
    return gef


def lade_protokoll(pfad):
    """{rnr: letzter Status} aus gesendet.csv (Spalten: rnr;zeit;status;kanal;empfaenger;pdf;modus)."""
    d = {}
    if not os.path.exists(pfad):
        return d
    with open(pfad, encoding="utf-8-sig", newline="") as f:
        for t in csv.reader(f, delimiter=";"):
            if len(t) >= 3 and t[0].isdigit():
                d[t[0]] = t[2]
    return d


def protokolliere(pfad, rnr, status, kanal, empfaenger, pdf, modus):
    neu = not os.path.exists(pfad)
    with open(pfad, "a", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f, delimiter=";", lineterminator="\n")
        if neu:
            w.writerow(["Rechnungsnummer", "Zeit", "Status", "Kanal", "Empfaenger", "PDF", "Modus"])
        w.writerow([rnr, datetime.now().strftime("%Y-%m-%d %H:%M:%S"), status, kanal, empfaenger,
                    os.path.basename(pdf), modus])
        f.flush()
        os.fsync(f.fileno())


def lade_sperre(pfad):
    s = set()
    if os.path.exists(pfad):
        with open(pfad, encoding="utf-8") as f:
            for z in f:
                z = z.split("#")[0].strip()
                if z:
                    s.add(re.sub(r"\D", "", z))
    return s


# ----------------------------------------------------------------------------------------------
# Mail bauen / senden
# ----------------------------------------------------------------------------------------------
def baue_mail(cfg, rnr, empfaenger, pdf_pfad):
    msg = EmailMessage()
    absender = cfg["absender"]
    msg["From"] = formataddr((cfg.get("absender_name", ""), absender)) if cfg.get("absender_name") else absender
    msg["To"] = empfaenger
    if cfg.get("bcc"):
        msg["Bcc"] = cfg["bcc"]
    if cfg.get("antwort_an"):
        msg["Reply-To"] = cfg["antwort_an"]
    msg["Subject"] = cfg.get("betreff", "Ihre Rechnung Nr. {rnr}").format(rnr=rnr)
    msg["Date"] = formatdate(localtime=True)
    msg["Message-ID"] = make_msgid(domain=absender.split("@", 1)[-1])
    msg.set_content(cfg.get("text", "Guten Tag,\n\nanbei erhalten Sie Ihre Rechnung Nr. {rnr}.\n").format(rnr=rnr))
    with open(pdf_pfad, "rb") as f:
        msg.add_attachment(f.read(), maintype="application", subtype="pdf", filename="Rechnung_%s.pdf" % rnr)
    return msg


def sende_smtp(cfg, msg):
    s = cfg["smtp"]
    sicher = (s.get("sicherheit") or "starttls").lower()
    host, port = s["host"], int(s.get("port", 587))
    if sicher == "ssl":
        srv = smtplib.SMTP_SSL(host, port, timeout=60)
    else:
        srv = smtplib.SMTP(host, port, timeout=60)
    try:
        srv.ehlo()
        if sicher == "starttls":
            srv.starttls()
            srv.ehlo()
        if s.get("benutzer"):
            pw = s.get("passwort") or os.environ.get(s.get("passwort_env", ""), "")
            if not pw:
                raise RuntimeError("SMTP-Passwort fehlt (passwort oder passwort_env)")
            srv.login(s["benutzer"], pw)
        empf = [a for a in (msg["To"], msg["Bcc"]) if a]
        fehler = srv.send_message(msg, to_addrs=[x.strip() for e in empf for x in e.split(",")])
        if fehler:
            raise RuntimeError("Empfaenger abgelehnt: %s" % fehler)
    finally:
        try:
            srv.quit()
        except Exception:
            pass


# ----------------------------------------------------------------------------------------------
# Lauf
# ----------------------------------------------------------------------------------------------
class EinzelInstanz:
    def __init__(self, pfad):
        self.pfad = pfad
        self.ok = False

    def __enter__(self):
        try:
            if os.path.exists(self.pfad) and time.time() - os.path.getmtime(self.pfad) < 3 * 3600:
                return self
            with open(self.pfad, "w") as f:
                f.write(str(os.getpid()))
            self.ok = True
        except OSError:
            pass
        return self

    def __exit__(self, *a):
        if self.ok:
            try:
                os.remove(self.pfad)
            except OSError:
                pass


def lauf(cfg, modus=None, jetzt=None, log=print):
    """Ein kompletter Durchlauf. Rueckgabe: dict mit Zaehlern und Ergebnisliste."""
    jetzt = jetzt or datetime.now()
    modus = modus or cfg.get("modus", "bericht")
    if modus not in ("bericht", "entwurf", "senden"):
        raise ValueError("modus muss bericht, entwurf oder senden sein")
    if not cfg.get("ab_rechnungsnummer"):
        raise ValueError("ab_rechnungsnummer fehlt - Pflichtangabe (verhindert Versand alter Rechnungen)")
    ab = int(cfg["ab_rechnungsnummer"])
    maxalter = int(cfg.get("max_alter_tage", 4))
    maxmails = int(cfg.get("max_mails_pro_lauf", 400))
    rdir = cfg["rechnungen_ordner"]
    aus = cfg["ausgabe_ordner"]
    os.makedirs(aus, exist_ok=True)
    prot_pfad = os.path.join(aus, "gesendet.csv")
    if not os.path.isdir(rdir):
        raise RuntimeError("Rechnungsordner nicht erreichbar: %s - nichts gesendet" % rdir)
    protokoll = lade_protokoll(prot_pfad)
    sperre = lade_sperre(os.path.join(aus, "nie_senden.txt"))

    # 1) Kandidaten-PDFs (nur junge Dateien), je Rechnungsnummer die NEUESTE Datei
    kandidaten = {}
    ignoriert = []
    grenze = jetzt - timedelta(days=maxalter)
    try:
        namen = os.listdir(rdir)
    except OSError as e:
        raise RuntimeError("Rechnungsordner nicht lesbar: %s" % e)
    for n in namen:
        if not n.lower().endswith(".pdf"):
            continue
        p = os.path.join(rdir, n)
        try:
            mt = datetime.fromtimestamp(os.path.getmtime(p))
        except OSError:
            continue
        if mt < grenze:
            continue
        nr_name = rnr_aus_dateiname(n)
        if nr_name and nr_name.isdigit() and int(nr_name) < ab:
            continue                                   # alt (vor Startnummer) - gar nicht erst lesen
        kandidaten.setdefault("?", []).append((mt, p))
    pdfs = sorted(kandidaten.get("?", []), key=lambda x: x[0])
    je_rnr = {}
    texte = {}
    for mt, p in pdfs:
        t = pdf_text(p)
        ok, grund = ist_kundenrechnung(t)
        if not ok:
            ignoriert.append((os.path.basename(p), grund))
            continue
        rnr = rechnungsnummer_aus_text(t)
        if int(rnr) < ab:
            continue
        je_rnr[rnr] = p                                # spaetere (neuere) Datei gewinnt
        texte[rnr] = t

    # 2) Packlisten suchen und lesen
    pl_pfade = suche_packlisten(cfg.get("packlisten_ordner", []), set(je_rnr))
    ergebnisse = []
    gesendet = entworfen = 0
    for rnr in sorted(je_rnr):
        pdf = je_rnr[rnr]
        rtext = texte[rnr]
        zeile = {"rnr": rnr, "pdf": os.path.basename(pdf), "kanal": "", "entscheidung": "", "grund": "", "empfaenger": ""}
        ergebnisse.append(zeile)
        if protokoll.get(rnr) == "gesendet":
            zeile.update(entscheidung="bereits gesendet", grund="laut Protokoll schon versendet")
            continue
        if protokoll.get(rnr) == "versuch":
            zeile.update(entscheidung="NICHT gesendet", grund="früherer Versuch ohne Bestätigung - bitte von Hand prüfen "
                                                              "(wurde die Mail evtl. schon verschickt?)")
            continue
        if rnr in sperre:
            zeile.update(entscheidung="NICHT gesendet", grund="steht in nie_senden.txt")
            continue
        interne = INTERN_MARKER.search(rtext)
        packs = []
        for pp in pl_pfade.get(rnr, []):
            pk = lies_packliste(pp)
            if pk and pk["rnr"] == rnr:
                packs.append(pk)
        pack = None
        if packs:
            # mehrere Packlisten derselben Rechnung muessen dieselben Angaben haben
            sig = {(p_["fremdbeleg"], tuple(p_["alle_mails"])) for p_ in packs}
            pack = packs[-1]
            if len(sig) > 1:
                pack = dict(pack, alle_mails=sorted({m for p_ in packs for m in p_["alle_mails"]}), fremdbeleg="")
                zeile.update(kanal=KANAL_UNKLAR, entscheidung="NICHT gesendet",
                             grund="mehrere Packlisten mit unterschiedlichen Angaben")
                continue
        kanal, grund, empf = klassifiziere(pack, rtext)
        zeile.update(kanal=kanal, grund=grund, empfaenger=empf)
        if kanal == KANAL_AMAZON:
            zeile["entscheidung"] = "NICHT gesendet (Amazon)"
            continue
        if kanal == KANAL_UNKLAR:
            zeile["entscheidung"] = "NICHT gesendet (unklar)"
            continue
        if interne:
            zeile.update(entscheidung="NICHT gesendet", grund="Rechnungs-PDF enthält interne Angaben (%s) - "
                                                              "falsches Formular?" % interne.group(0).strip())
            continue
        if gesendet + entworfen >= maxmails:
            zeile.update(entscheidung="NICHT gesendet", grund="Höchstzahl Mails je Lauf (%d) erreicht" % maxmails)
            continue
        # --- Versand / Entwurf ---
        try:
            msg = baue_mail(cfg, rnr, empf, pdf)
        except Exception as e:
            zeile.update(entscheidung="NICHT gesendet", grund="Mail konnte nicht gebaut werden: %s" % e)
            continue
        if modus == "bericht":
            zeile["entscheidung"] = "würde senden"
        elif modus == "entwurf":
            ed = os.path.join(aus, "Entwuerfe", jetzt.strftime("%Y-%m-%d"))
            os.makedirs(ed, exist_ok=True)
            with open(os.path.join(ed, "Rechnung_%s_%s.eml" % (rnr, kanal)), "wb") as f:
                f.write(bytes(msg))
            zeile["entscheidung"] = "Entwurf erstellt"
            entworfen += 1
        else:
            protokolliere(prot_pfad, rnr, "versuch", kanal, empf, pdf, modus)     # VOR dem Senden (fail-closed)
            try:
                sende_smtp(cfg, msg)
            except Exception as e:
                zeile.update(entscheidung="FEHLER beim Senden", grund=str(e))
                protokolliere(prot_pfad, rnr, "fehler", kanal, empf, pdf, modus)
                log("FEHLER Rechnung %s: %s" % (rnr, e))
                if cfg.get("bei_fehler_abbrechen", True):
                    break
                continue
            protokolliere(prot_pfad, rnr, "gesendet", kanal, empf, pdf, modus)
            zeile["entscheidung"] = "gesendet"
            gesendet += 1
            time.sleep(float(cfg.get("pause_sekunden", 1.0)))
    # Bericht schreiben
    bericht = os.path.join(aus, "Bericht_%s.csv" % jetzt.strftime("%Y-%m-%d_%H%M"))
    with open(bericht, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f, delimiter=";", lineterminator="\n")
        w.writerow(["Rechnungsnummer", "Kanal", "Entscheidung", "Grund", "Empfaenger", "PDF", "Modus"])
        for z in ergebnisse:
            w.writerow([z["rnr"], z["kanal"], z["entscheidung"], z["grund"], z["empfaenger"], z["pdf"], modus])
    zaehler = {}
    for z in ergebnisse:
        zaehler[z["entscheidung"]] = zaehler.get(z["entscheidung"], 0) + 1
    return {"modus": modus, "ergebnisse": ergebnisse, "zaehler": zaehler, "bericht": bericht,
            "ignoriert": ignoriert, "gesendet": gesendet, "entworfen": entworfen}


def schreibe_zusammenfassung(res, log=print):
    log("Modus: %s  |  Bericht: %s" % (res["modus"], res["bericht"]))
    for k, v in sorted(res["zaehler"].items()):
        log("  %-32s %d" % (k, v))
    pruef = [z for z in res["ergebnisse"] if z["entscheidung"].startswith(("NICHT gesendet (unklar", "FEHLER"))
             or "früherer Versuch" in z["grund"]]
    if pruef:
        log("Bitte prüfen (nicht automatisch versendbar):")
        for z in pruef:
            log("  %s: %s" % (z["rnr"], z["grund"]))


STATUS_DATEI = "letzter_lauf.json"


def lade_status(ausgabe_ordner):
    """Letzter Lauf (vom Programm geschrieben) fuer die Anzeige im Carrier-Dashboard; None wenn keiner/unlesbar."""
    try:
        with open(os.path.join(ausgabe_ordner, STATUS_DATEI), encoding="utf-8") as f:
            return json.load(f)
    except (OSError, ValueError):
        return None


def schreibe_status(ausgabe_ordner, status):
    """Atomar schreiben (Temp-Datei + os.replace)."""
    pfad = os.path.join(ausgabe_ordner, STATUS_DATEI)
    tmp = pfad + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(status, f, ensure_ascii=False, indent=1)
    os.replace(tmp, pfad)


def _status_aus_ergebnis(res, quelle, fehler=""):
    erg = res["ergebnisse"] if res else []
    n_unklar = sum(1 for z in erg if z["entscheidung"].startswith("NICHT gesendet") and "Amazon" not in z["entscheidung"]
                   and "bereits" not in z["entscheidung"])
    n_fehler = sum(1 for z in erg if z["entscheidung"].startswith("FEHLER"))
    return {"zeit": datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "quelle": quelle, "version": VERSION,
            "modus": res["modus"] if res else "", "gesendet": res["gesendet"] if res else 0,
            "entworfen": res["entworfen"] if res else 0, "zaehler": res["zaehler"] if res else {},
            "unklar": n_unklar, "fehlerzahl": n_fehler, "fehler": fehler,
            "bericht": res["bericht"] if res else "", "ergebnisse": erg}


def fuehre_aus(cfg, modus=None, quelle="geplant", log=print):
    """Kompletter Lauf MIT Einzelinstanz-Sperre und Statusdatei (fuer Aufgabenplanung UND Dashboard-Knopf).
    Rueckgabe: Status-dict (siehe _status_aus_ergebnis); bei laufendem anderen Lauf {'laeuft_bereits': True}.
    Jeder Fehler wird als Status mit 'fehler' festgehalten, damit er im Dashboard sichtbar wird."""
    aus = cfg["ausgabe_ordner"]
    os.makedirs(aus, exist_ok=True)
    logpfad = os.path.join(aus, "rechnungsversand.log")

    def schreib_log(t):
        z = "%s  %s" % (datetime.now().strftime("%Y-%m-%d %H:%M:%S"), t)
        try:
            log(z)
        except Exception:
            pass
        try:
            with open(logpfad, "a", encoding="utf-8") as f:
                f.write(z + "\n")
        except OSError:
            pass
    with EinzelInstanz(os.path.join(aus, "lauf.lock")) as lock:
        if not lock.ok:
            schreib_log("Abbruch: läuft bereits (lauf.lock vorhanden)")
            return {"laeuft_bereits": True}
        try:
            res = lauf(cfg, modus, log=schreib_log)
        except Exception as e:
            schreib_log("ABBRUCH: %s" % e)
            st = _status_aus_ergebnis(None, quelle, fehler=str(e))
            st["modus"] = modus or cfg.get("modus", "")
            schreibe_status(aus, st)
            return st
        schreibe_zusammenfassung(res, schreib_log)
        st = _status_aus_ergebnis(res, quelle)
        schreibe_status(aus, st)
        return st


def main(argv=None):
    ap = argparse.ArgumentParser(description="Sicherer Rechnungsversand (nur Onlineshop und eBay, nie Amazon)")
    ap.add_argument("--config", default=os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                                     "rechnungsversand_config.json"))
    ap.add_argument("--trocken", action="store_true", help="nur Bericht, nichts senden (überschreibt modus)")
    ap.add_argument("--quelle", default="geplant", help="geplant (Aufgabenplanung) oder manuell")
    a = ap.parse_args(argv)
    with open(a.config, encoding="utf-8") as f:
        cfg = json.load(f)
    st = fuehre_aus(cfg, "bericht" if a.trocken else None, a.quelle)
    if st.get("laeuft_bereits"):
        return 2
    return 1 if st.get("fehler") else 0


# ----------------------------------------------------------------------------------------------
# Selbsttest
# ----------------------------------------------------------------------------------------------
def _selftest():
    import tempfile
    import threading
    import socketserver
    n_ok, fails = 0, []

    def check(name, got, want):
        nonlocal n_ok
        if got == want:
            n_ok += 1
        else:
            fails.append("%s: erwartet %r, war %r" % (name, want, got))

    def pack(fremd, mails, ex="", rnr="1710008"):
        t = "Artikelnr Bezeichnung\n...\nRechnung Nr. : %s\n%s\n%s\nSendungsgewicht : 1,0\nGasecenter, Karlsruher Straße 3\n%s" % (
            rnr, fremd, "\n".join(mails), ex)
        return {"rnr": rnr, "fremdbeleg": fremd_(t), "mails": [m.lower() for m in fremdbeleg_und_mails(t)[1]],
                "alle_mails": alle_kundenmails(t), "text": t, "pfad": "x"}

    def fremd_(t):
        return fremdbeleg_und_mails(t)[0]

    rechn = "Gasecenter\nRechnung Nr. : 1710008\nKunden Nr. : 1\nRechnungsdatum: 20.10.2026\n%s\nRechnungsdatum = Lieferdatum falls\nRechnung\n"
    # --- echte Muster aus 1741 Packlisten ---
    k = lambda fr, ml, rf=None, ex="": klassifiziere(pack(fr, ml, ex), rechn % (rf if rf is not None else fr))[0]
    check("Amazon (Nr + marketplace)", k("305-9853395-4797960", ["j5nv1s6ywkx9ywl@marketplace.amazon.de"]), KANAL_AMAZON)
    check("eBay (Name + members.ebay.com)", k("clerprem_de", ["0090fee3be64c33e1a7c@members.ebay.com"]), KANAL_EBAY)
    check("Shop (Nummer + Kundenmail)", k("25623", ["thorsten.gottlieb@st-tiefbau.de"]), KANAL_SHOP)
    # --- Sicherheit: Amazon-Merkmal ueberall blockiert ---
    check("Amazon-Nr auf der Rechnung obwohl Packliste nach eBay aussieht",
          klassifiziere(pack("clerprem_de", ["x@members.ebay.com"]), rechn % "305-9853395-4797960")[0], KANAL_AMAZON)
    check("Amazon-Nr irgendwo in der Packliste", k("clerprem_de", ["x@members.ebay.com"], ex="Ref 305-9853395-4797960"), KANAL_AMAZON)
    check("Wort Amazon in Packliste", k("25623", ["a@b.de"], ex="per Amazon bestellt"), KANAL_AMAZON)
    check("Shop-Nr aber Amazon-Adresse", k("25623", ["a@marketplace.amazon.de"]), KANAL_AMAZON)
    check("Amazon-Adresse anderer Domain amazon.com", k("25623", ["a@amazon.com"]), KANAL_AMAZON)
    check("eBay-Name aber Amazon-Marketplace-Mail", k("clerprem_de", ["a@marketplace.amazon.de"]), KANAL_AMAZON)
    # --- Zweifel -> UNKLAR ---
    check("keine Mail", k("25623", []), KANAL_UNKLAR)
    check("zwei Mails", k("25623", ["a@b.de", "c@d.de"]), KANAL_UNKLAR)
    check("kein Fremdbeleg", klassifiziere(pack("", ["a@b.de"]), rechn % "")[0], KANAL_UNKLAR)
    check("normaler Name + normale Mail (kein eBay) -> unklar", k("hans_mueller", ["h@gmx.de"]), KANAL_UNKLAR)
    check("Zahl zu kurz (3 Stellen) + Mail -> unklar", k("123", ["h@gmx.de"]), KANAL_UNKLAR)
    check("Zahl zu lang (7 Stellen) + Mail -> unklar", k("1234567", ["h@gmx.de"]), KANAL_UNKLAR)
    check("eBay-Mail aber Zahl als Fremdbeleg -> unklar", k("25623", ["x@members.ebay.com"]), KANAL_UNKLAR)
    check("Shop-Mail mit 'members.' Domain -> unklar", k("25623", ["x@members.example.org"]), KANAL_UNKLAR)
    check("Fremdbeleg Rechnung weicht ab -> unklar", klassifiziere(pack("25623", ["a@b.de"]), rechn % "25999")[0], KANAL_UNKLAR)
    check("keine Packliste -> unklar", klassifiziere(None, "x")[0], KANAL_UNKLAR)
    check("Rechnung ohne Fremdbeleg-Zeile bleibt ok (Vorlagentext wird nicht als Fremdbeleg gelesen)",
          klassifiziere(pack("25623", ["a@b.de"]),
                        "Rechnungsdatum: 20.10.2026\nRechnungsdatum = Lieferdatum falls\nnicht anders angegeben\n")[0], KANAL_SHOP)
    check("eBay-Name + normale Mail + Shop-Nr (Historie, 2 Fremdbeleg-Zeilen) -> unklar (manuell)",
          klassifiziere(pack("svestei_0", ["x@gmail.com"], ex="25662"), rechn % "svestei_0")[0], KANAL_UNKLAR)
    check("eigene Firmenadresse zaehlt nicht als Kunde",
          klassifiziere(pack("25623", ["a@b.de"], ex="Email@Gasecenter-Augsburg.de"), rechn % "25623")[0], KANAL_SHOP)
    # --- Rechnungsart ---
    check("Rechnung ok", ist_kundenrechnung("Gasecenter\nRechnung Nr. : 1710008\n")[0], True)
    check("Gutschrift (Nr 111...)", ist_kundenrechnung("Rechnung Nr. : 1110008")[0], False)
    check("Auftragsbestätigung (444)", ist_kundenrechnung("Rechnung Nr. : 4440008")[0], False)
    check("ohne Nummer", ist_kundenrechnung("Hallo")[0], False)
    check("Stornobeleg", ist_kundenrechnung("Bestell-Storno\nRechnung Nr. : 1710008")[0], False)
    check("Interne Marker erkannt", bool(INTERN_MARKER.search("Pos\nLagerort: F5\n")), True)
    check("Kundenrechnung ohne interne Marker", bool(INTERN_MARKER.search("Artikel 1 Stück 3,00\nVielen Dank")), False)

    # --- Gesamtlauf mit PDFs + Fake-SMTP ---
    try:
        from reportlab.pdfgen import canvas
    except Exception:
        print("reportlab fehlt - Lauftest uebersprungen")
        canvas = None
    if canvas:
        tmp = tempfile.mkdtemp()
        rd, pd_, aus = (os.path.join(tmp, x) for x in ("rechnungen", "packlisten", "aus"))
        for d in (rd, pd_, aus):
            os.makedirs(d)

        def mkpdf(pfad, zeilen):
            c = canvas.Canvas(pfad)
            y = 800
            for z in zeilen:
                c.drawString(40, y, z)
                y -= 16
            c.save()

        def rechnung(rnr, fremd, name=None):
            mkpdf(os.path.join(rd, name or "Packliste Nr. %s 20261021080000.pdf" % rnr),
                  ["Gasecenter, Karlsruher Strasse 3", "Rechnung Nr. : %s" % rnr, "Kunden Nr. : 55",
                   "Rechnungsdatum: 21.10.2026", fremd, "Rechnungsdatum = Lieferdatum falls", "Rechnung",
                   "Artikel 1 Stueck 5,00", "Vielen Dank fuer Ihre Bestellung"])

        def packl(rnr, fremd, mail):
            mkpdf(os.path.join(pd_, "Packliste Nr %s.pdf" % rnr),
                  ["Artikelnr Bezeichnung", "Rechnung Nr. : %s" % rnr, fremd, mail, "Sendungsgewicht : 1,0",
                   "Gasecenter, Karlsruher Strasse 3"])
        for rnr, fremd, mail in (("1710010", "25623", "kunde@example.org"), ("1710011", "clerprem_de", "a1@members.ebay.com"),
                                 ("1710012", "305-9853395-4797960", "zz@marketplace.amazon.de"),
                                 ("1710013", "hans_x", "h@gmx.de"), ("1710015", "25700", "k2@example.org")):
            rechnung(rnr, fremd)
            packl(rnr, fremd, mail)
        rechnung("1710014", "25623")                      # ohne Packliste
        rechnung("1709999", "25623")                      # unter Startnummer
        rechnung("1710016", "25623")
        packl("1710016", "25623", "kunde@example.org")
        mkpdf(os.path.join(rd, "Gutschrift Nr. 1110001 2026.pdf"), ["Gutschrift", "Rechnung Nr. : 1110001"])
        mkpdf(os.path.join(rd, "irrefuehrend.pdf"), ["Rechnung Nr. : 1710020", "Rechnungsdatum: 21.10.2026", "25623",
                                                     "Lagerort: F5"])
        packl("1710020", "25623", "kunde@example.org")
        open(os.path.join(aus, "nie_senden.txt"), "w").write("1710015  # Testsperre\n")
        gefangen = []

        class H(socketserver.StreamRequestHandler):
            def handle(self):
                self.wfile.write(b"220 fake\r\n")
                daten = False
                buf = []
                while True:
                    l = self.rfile.readline()
                    if not l:
                        return
                    u = l.decode("latin-1").strip()
                    if daten:
                        if u == ".":
                            gefangen.append("\n".join(buf))
                            buf, daten = [], False
                            self.wfile.write(b"250 ok\r\n")
                        else:
                            buf.append(u)
                        continue
                    cmd = u.upper()
                    if cmd.startswith("EHLO") or cmd.startswith("HELO"):
                        self.wfile.write(b"250 fake\r\n")
                    elif cmd.startswith("DATA"):
                        daten = True
                        self.wfile.write(b"354 go\r\n")
                    elif cmd.startswith("QUIT"):
                        self.wfile.write(b"221 bye\r\n")
                        return
                    else:
                        self.wfile.write(b"250 ok\r\n")
        srv = socketserver.ThreadingTCPServer(("127.0.0.1", 0), H)
        srv.daemon_threads = True
        threading.Thread(target=srv.serve_forever, daemon=True).start()
        cfg = {"modus": "bericht", "rechnungen_ordner": rd, "packlisten_ordner": [pd_], "ausgabe_ordner": aus,
               "ab_rechnungsnummer": 1710000, "absender": "rechnung@example.org", "absender_name": "Test",
               "smtp": {"host": "127.0.0.1", "port": srv.server_address[1], "sicherheit": "none"},
               "pause_sekunden": 0}
        # ohne Startnummer -> Abbruch
        try:
            lauf(dict(cfg, ab_rechnungsnummer=0))
            check("Pflicht ab_rechnungsnummer", "kein Fehler", "ValueError")
        except ValueError:
            check("Pflicht ab_rechnungsnummer", "ValueError", "ValueError")
        res = lauf(cfg, "bericht")
        je = {z["rnr"]: z for z in res["ergebnisse"]}
        check("Bericht: Shop wuerde senden", je["1710010"]["entscheidung"], "würde senden")
        check("Bericht: eBay wuerde senden", je["1710011"]["entscheidung"], "würde senden")
        check("Bericht: Amazon gesperrt", je["1710012"]["entscheidung"], "NICHT gesendet (Amazon)")
        check("Bericht: unklarer Name/Mail", je["1710013"]["entscheidung"], "NICHT gesendet (unklar)")
        check("Bericht: ohne Packliste unklar", je["1710014"]["entscheidung"], "NICHT gesendet (unklar)")
        check("Bericht: nie_senden.txt greift", je["1710015"]["entscheidung"], "NICHT gesendet")
        check("Bericht: unter Startnummer fehlt", "1709999" in je, False)
        check("Bericht: Gutschrift ignoriert", "1110001" in je, False)
        check("Bericht: internes Formular (Lagerort) blockiert", je["1710020"]["entscheidung"], "NICHT gesendet")
        check("Bericht: nichts gesendet", len(gefangen), 0)
        res = lauf(cfg, "entwurf")
        check("Entwurf: 3 .eml", sorted(os.path.basename(p) for p in glob.glob(os.path.join(aus, "Entwuerfe", "*", "*.eml"))),
              ["Rechnung_1710010_SHOP.eml", "Rechnung_1710011_EBAY.eml", "Rechnung_1710016_SHOP.eml"])
        res = lauf(cfg, "senden")
        check("Senden: 3 Mails", len(gefangen), 3)
        check("Senden: Zaehler", res["gesendet"], 3)
        text_alle = "\n".join(gefangen)
        check("Senden: Anhang Rechnung_<nr>.pdf", "Rechnung_1710010.pdf" in text_alle, True)
        check("Senden: kein Amazon-Empfaenger", "amazon" in text_alle.lower(), False)
        check("Protokoll gesendet", sorted(k for k, v in lade_protokoll(os.path.join(aus, "gesendet.csv")).items() if v == "gesendet"),
              ["1710010", "1710011", "1710016"])
        res = lauf(cfg, "senden")
        check("Zweiter Lauf: keine Doppelversendung", (len(gefangen), res["gesendet"]), (3, 0))
        st = fuehre_aus(cfg, "bericht", "manuell", log=lambda t: None)
        st2 = lade_status(aus)
        check("Statusdatei: Quelle/Modus/Zaehler", (st2["quelle"], st2["modus"], st2["fehlerzahl"], st2["fehler"]),
              ("manuell", "bericht", 0, ""))
        check("Statusdatei: unklar-Zaehler (1710013, 1710014, 1710015, 1710020 ...)", st2["unklar"] >= 3, True)
        check("Statusdatei: Ergebnisliste vorhanden", any(z["rnr"] == "1710012" for z in st2["ergebnisse"]), True)
        st3 = fuehre_aus(dict(cfg, ab_rechnungsnummer=0), "bericht", "geplant", log=lambda t: None)
        check("Fehler (fehlende Startnummer) landet im Status", "ab_rechnungsnummer" in (lade_status(aus) or {}).get("fehler", ""), True)
        # Versuch ohne Bestaetigung -> nie blind wiederholen
        protokolliere(os.path.join(aus, "gesendet.csv"), "1710099", "versuch", "SHOP", "x@y.de", "x.pdf", "senden")
        rechnung("1710099", "25623")
        packl("1710099", "25623", "kunde@example.org")
        res = lauf(cfg, "senden")
        je = {z["rnr"]: z for z in res["ergebnisse"]}
        check("Versuch ohne Bestaetigung wird NICHT wiederholt", (je["1710099"]["entscheidung"], len(gefangen)), ("NICHT gesendet", 3))
        # Maximalzahl
        if os.path.exists(os.path.join(aus, "gesendet.csv")):
            os.remove(os.path.join(aus, "gesendet.csv"))
        res = lauf(dict(cfg, max_mails_pro_lauf=1), "entwurf")
        check("Maximalzahl Mails", res["entworfen"], 1)
        # SMTP-Fehler -> Abbruch und Protokoll 'fehler'
        srv.shutdown()
        srv.server_close()
        if os.path.exists(os.path.join(aus, "gesendet.csv")):
            os.remove(os.path.join(aus, "gesendet.csv"))
        res = lauf(cfg, "senden")
        check("SMTP weg: nichts als gesendet protokolliert",
              [k for k, v in lade_protokoll(os.path.join(aus, "gesendet.csv")).items() if v == "gesendet"], [])
        check("SMTP weg: FEHLER im Bericht", any(z["entscheidung"] == "FEHLER beim Senden" for z in res["ergebnisse"]), True)
    if fails:
        print("SELBSTTEST FEHLGESCHLAGEN (%d von %d):" % (len(fails), n_ok + len(fails)))
        for f in fails:
            print("  -", f)
        return False
    print("Selbsttest OK (%d Prüfungen)" % n_ok)
    return True


if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "selftest":
        sys.exit(0 if _selftest() else 1)
    sys.exit(main())
