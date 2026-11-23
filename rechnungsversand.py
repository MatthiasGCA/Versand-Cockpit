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
  3. Klassifizieren (siehe klassifiziere()): es entscheidet die E-Mail-Adresse. Gesendet wird, wenn genau EINE echte
     Kundenadresse vorliegt (Onlineshop, eBay-Relay @members.ebay.com oder sonstige) - nie an eine Amazon-Adresse
     (marketplace.amazon.de usw.). Alles andere (keine Packliste, keine/mehrere Adressen, Fremdbeleg weicht ab)
     wird NICHT gesendet und im Bericht aufgelistet.
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
import base64
import html as _html
from datetime import datetime, timedelta
from email.message import EmailMessage
from email.utils import formatdate, make_msgid, formataddr

VERSION = "2026-11-23b"
# 2026-11-23b: Offene Faelle werden bei JEDEM Lauf (auch "Pruefen") ausgetragen, wenn die Rechnung in nie_senden.txt
#   steht oder laut Protokoll gesendet ist - auch wenn sie aelter als max_alter_tage ist und gar nicht mehr gelesen
#   wird (real 1714151 vom 19.11.: frueheres "Erledigt" wurde vom Lauf wieder aufgenommen, danach zu alt -> blieb stehen).
# 2026-11-23a: "Erledigt" wirkte bei "frueherer Versuch ohne Bestaetigung" nicht - der Versuch-Check kam VOR der
#   Sperrliste, der naechste Lauf nahm den Fall wieder in die offenen Faelle auf (real 1714151). Sperrliste jetzt zuerst.
# 2026-11-16f: Sperrdatei lauf.lock enthaelt die Prozessnummer; laeuft dieser Prozess nicht mehr (Lauf abgebrochen,
#   z.B. Konsolenfenster geschlossen - real 16.11. 15:03, danach "laeuft bereits" bis zu 3 Stunden), wird die Sperre
#   sofort als verwaist erkannt und ersetzt. Dazu Rechnungsversand_starten.vbs: Start ohne Konsolenfenster.
# 2026-11-16e: Belege an eigene Kundennummern (Konfig "ignorierte_kundennummern", Standard 483911 = Ladengeschaeft/
#   Thekenrechner) werden still ignoriert - auch ohne Packliste (Theke-Verkaeufe haben keine). Pruefung an 1085
#   Rechnungen 1713500-1714699: einziger solcher Beleg war 1713932 (Kunde 483911).
# 2026-11-16d: THEKENRECHNER-AUSBUCHUNG (Wunsch Matthias): steht als Kunden-E-Mail der Packliste eine EIGENE Adresse
#   (z.B. email@gasecenter-augsburg.de), wird die Rechnung still ignoriert ("ignoriert (eigene Adresse)", grau, kein
#   offener Fall). Geprueft wird nur der Kopfblock unter "Rechnung Nr." - die Firmenadresse in der Fusszeile jeder
#   Rechnung zaehlt nicht.
# 2026-11-16c: OFFENE FAELLE (Wunsch Matthias): jede nicht versendete Problem-Rechnung (unklar, interne Angaben, Fehler,
#   Versuch ohne Bestaetigung ...) kommt in offene_rechnungen.json und bleibt dort, bis sie im Dashboard von Hand
#   "Erledigt" (-> nie_senden.txt) oder "Doch senden" (an die bekannte oder eine eingegebene Adresse, nie an Amazon)
#   bekommt - sie verschwindet nicht mehr unbemerkt, wenn sie aelter als max_alter_tage wird.
# 2026-11-16b: Rechnungen aus nie_senden.txt heissen "ausgeschlossen" (grau) und zaehlen nicht mehr als "unklar" - so kann
#   man eine von Hand erledigte Rechnung (z.B. aus Amicron verschickt) still stellen.
# 2026-11-16a: (1) Interne Belege an die eigene Firma (z.B. "MONATLICHE VERKAEUFE LADENGESCHAEFT", Empfaenger
#   "Gasecenter Augsburg", real 1713932) werden still ignoriert statt jeden Tag als "unklar" gemeldet. (2) Amazon-
#   Bestellnummer auf der Rechnung und keine Packliste (E-Mail-Adresse unbekannt, real 1714238) -> Amazon (ruhig).
# 2026-11-11a: Rechnungen ueber der Hoechstzahl je Lauf heissen "zurueckgestellt" (grau) statt "NICHT gesendet" und
#   zaehlen nicht mehr als "unklar" (Testlauf mit max 5 zeigte sonst "224 unklar" und loeste die Warnung aus).
# 2026-11-06b: Mischfall (Matthias): bestellt ein eBay-Kunde spaeter direkt im Shop, steht auf der Packliste noch der alte eBay-Name,
#   auf der Rechnung aber die Shop-Nummer (real Rg 1712303). Das wird jetzt als Onlineshop GESENDET - nur wenn genau eine echte
#   Kundenadresse da ist (keine eBay-Relay-, keine Amazon-Adresse) und die Packliste einen Namen (keine Zahl/Amazon-Nr) nennt.
#   Der Mailtext nennt dann die Shop-Nummer der Rechnung. Alle anderen Abweichungen bleiben unklar.
# 2026-11-06a: Regel (Matthias): Es entscheidet die E-MAIL-ADRESSE. Gesendet wird an jede einzelne echte Kundenadresse,
#   nur nie an eine Amazon-Adresse (amazon/marketplace), weil die ueber das Amazon-Nachrichtensystem laeuft. Amazon-
#   Bestellnummer/Wort 'Amazon' im Text blockieren nicht mehr, wenn eine echte Adresse da ist. Neuer Kanal SONSTIGE
#   (echte Adresse, weder eBay-Relay noch Shop-Nummer). Weiter UNKLAR (nie gesendet): keine Packliste, keine/mehrere
#   Adressen, Fremdbeleg Rechnung <> Packliste.
# 2026-10-27b: Testversand: ist "test_empfaenger" gesetzt, gehen im Modus "senden" ALLE Mails an diese Adresse statt an
#   die Kunden (Betreff "[TEST] ..."). Protokoll-Status "testversuch"/"test" - blockiert den spaeteren echten Versand nicht.
# 2026-10-27a: HTML-Mail mit Firmensignatur (Vorlage rechnungsmail_vorlage.html, Konfig "html_vorlage") plus
#   Klartext-Teil; data:-Bilder der Vorlage (Logo) werden als eingebettete Bilder (CID) verschickt. Platzhalter
#   {rnr} {rechnungsdatum} {bestellung} in Betreff, Text und Vorlage ({bestellung} je Kanal: "Ihre Bestellung
#   Nr. 25623 in unserem Onlineshop" / "Ihre Bestellung bei eBay"). Fehlt die Vorlage -> Mail wird NICHT gebaut.

RE_AMAZON_NR = re.compile(r"\b\d{3}-\d{7}-\d{7}\b")
RE_MAIL = re.compile(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}")
# "Rechnung Nr. : 1710008" - steht bei vierzeiligen Adressen teils HINTER dem Ortstext derselben Zeile
RE_RNR_KOPF = re.compile(r"\bRechnung\s*Nr\.?\s*:\s*(\d{6,8})\b")
FIRMEN_DOMAINS = ("gasecenter-augsburg.de", "gasecenter-onlineshop.de")
GESPERRTE_PRAEFIXE = ("111", "444")                 # Gutschriften / Auftragsbestaetigungen
# Belege, die KEINE Kundenrechnung sind (Textmarker im PDF)
KEINE_RECHNUNG = ("gutschrift", "storno", "wareneingang", "auftragsbestätigung", "auftragsbestaetigung",
                  "duplikat", "lieferschein", "angebot", "mahnung")
# interne Angaben, die NICHT auf einer Kundenrechnung stehen duerfen (sonst keine Mail)
INTERN_MARKER = re.compile(r"Lagerort\s*:|Sendungsgewicht\s*:|^\s*(pax1|pox1|brx1|wapo)\s*$", re.I | re.M)
EBAY_DOMAINS = ("members.ebay.com",)
MARKTPLATZ_WOERTER = ("amazon", "ebay", "marketplace", "members.")

KANAL_SHOP, KANAL_EBAY, KANAL_AMAZON, KANAL_UNKLAR, KANAL_SONST = "SHOP", "EBAY", "AMAZON", "UNKLAR", "SONSTIGE"


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
    zeilen = [z.strip().lower() for z in (text or "").split("\n")[:4]]
    if "verkäufe ladengeschäft" in kopf or "verkaeufe ladengeschaeft" in kopf or \
            any(z.startswith("gasecenter augsburg") for z in zeilen[1:3]):
        return False, "interner Beleg an die eigene Firma (z.B. Ladengeschäft)"
    return True, ""


def ist_firmenadresse(m):
    m = m.lower()
    return any(m.endswith("@" + d) or m.endswith("." + d) for d in FIRMEN_DOMAINS)


def fremdbeleg_und_mails(text):
    """Aus dem Kopf einer Rechnung/Packliste: (fremdbeleg, [mails]).
    Fremdbeleg = die erste Zeile NACH 'Rechnung Nr.', die weder Mail, Rechnungsnummer noch Feldname ist."""
    zeilen = (text or "").split("\n")
    fremd, mails = "", []
    idx = next((i for i, z in enumerate(zeilen) if re.search(r"\bRechnung\s*Nr\.?\s*:", z)), None)
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


def eigene_adresse_im_kopf(text):
    """Eigene Firmenadresse (z.B. email@gasecenter-augsburg.de) als KUNDEN-Adresse im Kopfblock unter 'Rechnung Nr.'
    (Thekenrechner-Ausbuchung). Die Firmenadresse in der Fusszeile zaehlt nicht."""
    zeilen = (text or "").split("\n")
    idx = next((i for i, z in enumerate(zeilen) if re.search(r"\bRechnung\s*Nr\.?\s*:", z)), None)
    if idx is None:
        return ""
    for z in zeilen[idx + 1: idx + 7]:
        for m in RE_MAIL.findall(z):
            if ist_firmenadresse(m):
                return m.lower()
    return ""


def lies_packliste(pfad):
    """Packliste -> dict(rnr, fremdbeleg, mails, alle_mails, eigene_adresse, text) oder None."""
    text = pdf_text(pfad)
    rnr = rechnungsnummer_aus_text(text)
    if not rnr:
        return None
    fremd, mails = fremdbeleg_und_mails(text)
    return {"rnr": rnr, "fremdbeleg": fremd, "mails": [m.lower() for m in mails],
            "alle_mails": alle_kundenmails(text), "eigene_adresse": eigene_adresse_im_kopf(text),
            "text": text, "pfad": pfad}


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


def ist_amazon_adresse(m):
    """Amazon-Kaeuferadressen (marketplace.amazon.de u.a.): ueber sie laeuft das Amazon-Nachrichtensystem - dorthin
    darf NIE eine Rechnungsmail gehen."""
    m = (m or "").lower()
    return "amazon" in m or "marketplace" in m


def klassifiziere(pack, rechnung_text):
    """(kanal, grund, empfaenger). kanal: SHOP | EBAY | SONSTIGE | AMAZON | UNKLAR. AMAZON und UNKLAR werden nie gesendet.

    Regel (seit 2026-11-06a, Matthias): ENTSCHEIDEND IST DIE E-MAIL-ADRESSE. Gesendet wird an jede einzelne echte
    Kundenadresse - nur nicht an eine Amazon-Adresse (marketplace.amazon.de usw.), denn die laeuft ueber das
    Amazon-Nachrichtensystem. Hat eine Amazon-Bestellung ausnahmsweise eine echte Kundenadresse, ist das erlaubt.
      AMAZON:   eine Adresse mit 'amazon'/'marketplace' kommt in Packliste oder Rechnung vor - ODER es gibt ein
                Amazon-Merkmal (Bestellnummer/Wort) und keine einzelne echte Kundenadresse -> nie senden.
      UNKLAR:   keine Packliste, keine oder mehrere Kundenadressen, oder der Fremdbeleg der Rechnung weicht von
                dem der Packliste ab (dann passt die Packliste evtl. nicht zur Rechnung) -> nie senden.
      EBAY:     Adresse @members.ebay.com.
      SHOP:     Fremdbeleg ist genau eine 4-6-stellige Zahl (Onlineshop-Bestellnummer).
      SONSTIGE: jede andere echte Kundenadresse.
    """
    if not pack:
        if amazon_signale(rechnung_text or ""):    # Amazon-Bestellung ohne Packliste: Adresse unbekannt -> ruhig nie senden
            return KANAL_AMAZON, "Amazon-Bestellung ohne Packliste (E-Mail-Adresse unbekannt)", ""
        return KANAL_UNKLAR, "keine zugehörige Packliste gefunden (E-Mail-Adresse unbekannt)", ""
    ptext, rtext = pack.get("text", ""), rechnung_text or ""
    sig = amazon_signale(ptext, rtext)
    alle = pack.get("alle_mails") or []
    # Adressen auch aus dem Rechnungstext pruefen (dort steht normalerweise keine, aber sicher ist sicher)
    if any(ist_amazon_adresse(m) for m in list(alle) + alle_kundenmails(rtext)):
        return KANAL_AMAZON, "Amazon-/Marketplace-Adresse", ""
    fremd = (pack.get("fremdbeleg") or "").strip()
    mails = pack.get("mails") or []
    if len(alle) != 1 or len(set(mails)) != 1 or mails[0] != alle[0]:
        if sig:                        # Amazon-Merkmal und keine einzelne echte Adresse -> Amazon (nie senden, ruhig)
            return KANAL_AMAZON, "Amazon-Merkmal: " + ", ".join(sig), ""
        return KANAL_UNKLAR, "keine eindeutige einzelne Kunden-E-Mail-Adresse (%d gefunden)" % len(alle), ""
    mail = alle[0]
    dom = mail.split("@", 1)[1]
    # Kundenrechnung darf ihren Fremdbeleg nur identisch zeigen
    r_fremd, _ = fremdbeleg_und_mails_rechnung(rtext)
    # (sehr lange eBay-Namen kuerzt Amicron auf der Rechnung - Abkuerzung eines identischen Namens gilt als gleich)
    kurz, lang = sorted((r_fremd, fremd), key=len)
    gleich = r_fremd == fremd or (len(kurz) >= 20 and lang.startswith(kurz))
    if r_fremd and not gleich:
        # Mischfall (Matthias 06.11.): bestellt ein eBay-Kunde spaeter direkt im Shop, zieht Amicron den alten eBay-Namen
        # aus dem Kundensatz in die Packliste, die Rechnung zeigt aber die Shop-Nummer. Die Rechnung gehoert dann
        # eindeutig zur Shop-Bestellung -> senden, ABER nur an eine echte Adresse (keine eBay-Relay-Adresse: die
        # koennte im alten Kundensatz veraltet sein) und nur wenn die Packliste einen Namen (keine Zahl/Amazon-Nr) nennt.
        if (re.fullmatch(r"\d{4,6}", r_fremd) and fremd and not re.fullmatch(r"\d+", fremd)
                and not RE_AMAZON_NR.search(fremd) and dom not in EBAY_DOMAINS):
            zus = (" [Amazon-Merkmal: %s, aber echte Kundenadresse]" % ", ".join(sig)) if sig else ""
            return (KANAL_SHOP, "Onlineshop-Nummer %s (Rechnung) + Kundenadresse; Packliste nennt noch den "
                                "eBay-Namen '%s' (Mischfall)%s" % (r_fremd, fremd, zus), mail)
        return KANAL_UNKLAR, "Fremdbeleg auf Rechnung (%s) weicht von Packliste (%s) ab" % (r_fremd, fremd), ""
    zusatz = (" [Amazon-Merkmal: %s, aber echte Kundenadresse]" % ", ".join(sig)) if sig else ""
    if dom in EBAY_DOMAINS:
        return KANAL_EBAY, ("eBay-Relay-Adresse, Fremdbeleg '%s'" % fremd) + zusatz, mail
    if re.fullmatch(r"\d{4,6}", fremd):
        return KANAL_SHOP, "Onlineshop-Nummer %s + Kundenadresse%s" % (fremd, zusatz), mail
    return KANAL_SONST, "Kundenadresse, Fremdbeleg '%s'%s" % (fremd or "-", zusatz), mail


def fremdbeleg_fuer_mail(kanal, pack, rtext):
    """Fremdbeleg fuer den Mailtext. Beim Kanal SHOP kommt die Shop-Nummer von der RECHNUNG (im Mischfall nennt die
    Packliste noch den alten eBay-Namen), sonst die der Packliste."""
    fb = (pack or {}).get("fremdbeleg", "")
    if kanal == KANAL_SHOP:
        rf, _ = fremdbeleg_und_mails_rechnung(rtext)
        if re.fullmatch(r"\d{4,6}", rf or ""):
            fb = rf
    return fb


def fremdbeleg_und_mails_rechnung(text):
    """Fremdbeleg der KUNDENRECHNUNG (steht nach 'Rechnungsdatum:'); '' wenn nicht erkennbar."""
    zeilen = (text or "").split("\n")
    j = next((i for i, z in enumerate(zeilen) if "Rechnungsdatum:" in z), None)
    if j is None:
        return "", []
    for z in zeilen[j + 1: j + 4]:
        z = z.strip()
        if "Rechnungsdatum =" in z or z.lower() == "rechnung":
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
    """Eine Zeile ins Sendeprotokoll. Mehrere Versuche (Datei evtl. kurz in Excel offen); scheitert es ganz, wird
    die Ausnahme weitergereicht - VOR dem Senden heisst das: es wird nicht gesendet (fail-closed)."""
    for i in range(5):
        try:
            return _protokolliere_einmal(pfad, rnr, status, kanal, empfaenger, pdf, modus)
        except OSError:
            if i == 4:
                raise
            time.sleep(0.5 * (i + 1))


def _protokolliere_einmal(pfad, rnr, status, kanal, empfaenger, pdf, modus):
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
RE_RECHNUNGSDATUM = re.compile(r"Rechnungsdatum:\s*(\d{2}\.\d{2}\.\d{4})")
RE_DATA_BILD = re.compile(r'src="data:image/([A-Za-z0-9.+-]+);base64,([A-Za-z0-9+/=\s]+)"')


def rechnungsdatum_aus_text(text):
    m = RE_RECHNUNGSDATUM.search(text or "")
    return m.group(1) if m else ""


def platzhalter(rnr, kanal="", fremdbeleg="", rechnungsdatum=""):
    if kanal == KANAL_SHOP and fremdbeleg:
        best = "Ihre Bestellung Nr. %s in unserem Onlineshop" % fremdbeleg
    elif kanal == KANAL_EBAY:
        best = "Ihre Bestellung bei eBay"
    else:
        best = "Ihre Bestellung"
    return {"{rnr}": rnr, "{rechnungsdatum}": rechnungsdatum or datetime.now().strftime("%d.%m.%Y"),
            "{bestellung}": best}


def ersetze(text, werte, html_esc=False):
    for k, v in werte.items():
        text = text.replace(k, _html.escape(v) if html_esc else v)
    return text


def vorlage_pfad(cfg):
    p = cfg.get("html_vorlage") or ""
    if p and not os.path.isabs(p):
        p = os.path.join(os.path.dirname(os.path.abspath(__file__)), p)
    return p


def html_mit_cid(html_text, domain="gasecenter-augsburg.de"):
    """data:-Bilder -> cid:-Verweise. Rueckgabe (html, [(bytes, subtype, cid_mit_klammern)])."""
    bilder = []

    def ers(m):
        cid = make_msgid(domain=domain)
        bilder.append((base64.b64decode(re.sub(r"\s+", "", m.group(2))), m.group(1).lower(), cid))
        return 'src="cid:%s"' % cid[1:-1]
    return RE_DATA_BILD.sub(ers, html_text), bilder


def baue_mail(cfg, rnr, empfaenger, pdf_pfad, kanal="", fremdbeleg="", rechnungsdatum=""):
    werte = platzhalter(rnr, kanal, fremdbeleg, rechnungsdatum)
    msg = EmailMessage()
    absender = cfg["absender"]
    msg["From"] = formataddr((cfg.get("absender_name", ""), absender)) if cfg.get("absender_name") else absender
    msg["To"] = empfaenger
    if cfg.get("bcc"):
        msg["Bcc"] = cfg["bcc"]
    if cfg.get("antwort_an"):
        msg["Reply-To"] = cfg["antwort_an"]
    msg["Subject"] = ersetze(cfg.get("betreff", "Ihre Rechnung Nr. {rnr}"), werte)
    msg["Date"] = formatdate(localtime=True)
    msg["Message-ID"] = make_msgid(domain=absender.split("@", 1)[-1])
    msg.set_content(ersetze(cfg.get("text", "Guten Tag,\n\nanbei erhalten Sie Ihre Rechnung Nr. {rnr}.\n"), werte))
    vp = vorlage_pfad(cfg)
    if vp:
        with open(vp, encoding="utf-8") as f:            # fehlt die Vorlage -> Fehler -> keine Mail (fail-closed)
            h = ersetze(f.read(), werte, html_esc=True)
        h, bilder = html_mit_cid(h, absender.split("@", 1)[-1] or "gasecenter-augsburg.de")
        msg.add_alternative(h, subtype="html")
        html_teil = msg.get_payload()[1]
        for daten, subtyp, cid in bilder:
            html_teil.add_related(daten, maintype="image", subtype=subtyp, cid=cid)
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

    @staticmethod
    def _prozess_laeuft(pid):
        """True, wenn der Prozess pid noch laeuft (Windows ueber OpenProcess/GetExitCodeProcess). Im Zweifel True."""
        try:
            pid = int(pid)
        except (TypeError, ValueError):
            return True
        if pid == os.getpid():
            return True
        if os.name != "nt":
            try:
                os.kill(pid, 0)
                return True
            except OSError:
                return False
        try:
            import ctypes
            k = ctypes.windll.kernel32
            h = k.OpenProcess(0x1000, False, pid)            # PROCESS_QUERY_LIMITED_INFORMATION
            if not h:
                return False
            code = ctypes.c_ulong()
            ok = k.GetExitCodeProcess(h, ctypes.byref(code))
            k.CloseHandle(h)
            return bool(ok) and code.value == 259            # STILL_ACTIVE
        except Exception:
            return True

    def __enter__(self):
        try:
            if os.path.exists(self.pfad) and time.time() - os.path.getmtime(self.pfad) < 3 * 3600:
                try:
                    with open(self.pfad) as f:
                        pid_alt = f.read().strip()
                except OSError:
                    pid_alt = ""
                if self._prozess_laeuft(pid_alt):
                    return self                                # laeuft wirklich noch
                # verwaiste Sperre (Prozess beendet) -> uebernehmen
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
        if ok:
            mk = RE_KUNDENNR.search(t)
            if mk and mk.group(1) in {str(k) for k in cfg.get("ignorierte_kundennummern", IGNORIERTE_KUNDEN)}:
                ok, grund = False, "eigene Kundennummer %s (Ladengeschäft/Theke)" % mk.group(1)
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
        zeile = {"rnr": rnr, "pdf": os.path.basename(pdf), "kanal": "", "entscheidung": "", "grund": "", "empfaenger": "",
                 "pfad": pdf, "fremdbeleg": "", "rechnungsdatum": rechnungsdatum_aus_text(rtext)}
        ergebnisse.append(zeile)
        if protokoll.get(rnr) == "gesendet":
            zeile.update(entscheidung="bereits gesendet", grund="laut Protokoll schon versendet")
            continue
        if rnr in sperre:
            zeile.update(entscheidung="ausgeschlossen", grund="steht in nie_senden.txt (von Hand erledigt)")
            continue
        if protokoll.get(rnr) in ("versuch",):
            zeile.update(entscheidung="NICHT gesendet", grund="früherer Versuch ohne Bestätigung - bitte von Hand prüfen "
                                                              "(wurde die Mail evtl. schon verschickt?)")
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
        eigene = next((p_.get("eigene_adresse") for p_ in packs if p_.get("eigene_adresse")), "")
        if eigene:
            zeile.update(entscheidung="ignoriert (eigene Adresse)",
                         grund="Kunden-E-Mail ist die eigene Adresse %s (z.B. Thekenrechner-Ausbuchung)" % eigene)
            continue
        kanal, grund, empf = klassifiziere(pack, rtext)
        zeile.update(kanal=kanal, grund=grund, empfaenger=empf, fremdbeleg=fremdbeleg_fuer_mail(kanal, pack, rtext))
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
            zeile.update(entscheidung="zurückgestellt", grund="Höchstzahl Mails je Lauf (%d) erreicht - kommt im nächsten Lauf dran" % maxmails)
            continue
        # --- Versand / Entwurf ---
        try:
            test_an = (cfg.get("test_empfaenger") or "").strip()
            msg = baue_mail(cfg, rnr, test_an or empf, pdf, kanal, fremdbeleg_fuer_mail(kanal, pack, rtext),
                            rechnungsdatum_aus_text(rtext))
            if test_an:
                msg.replace_header("Subject", "[TEST für %s] %s" % (empf, msg["Subject"]))
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
            st_v, st_ok = ("testversuch", "test") if test_an else ("versuch", "gesendet")
            protokolliere(prot_pfad, rnr, st_v, kanal, test_an or empf, pdf, modus)   # VOR dem Senden (fail-closed)
            try:
                sende_smtp(cfg, msg)
            except Exception as e:
                zeile.update(entscheidung="FEHLER beim Senden", grund=str(e))
                protokolliere(prot_pfad, rnr, "fehler", kanal, empf, pdf, modus)
                log("FEHLER Rechnung %s: %s" % (rnr, e))
                if cfg.get("bei_fehler_abbrechen", True):
                    break
                continue
            protokolliere(prot_pfad, rnr, st_ok, kanal, test_an or empf, pdf, modus)
            zeile["entscheidung"] = ("TEST gesendet an %s" % test_an) if test_an else "gesendet"
            gesendet += 1
            time.sleep(float(cfg.get("pause_sekunden", 1.0)))
    # Offene Faelle pflegen: Problemfaelle aufnehmen, versendete/ausgeschlossene austragen
    try:
        offen = lade_offen(aus)
        for z in ergebnisse:
            e = z["entscheidung"]
            if ist_problemfall(e):
                alt = offen.get(z["rnr"], {})
                offen[z["rnr"]] = {"rnr": z["rnr"], "pdf": z.get("pfad", ""), "kanal": z["kanal"], "grund": z["grund"],
                                   "entscheidung": e, "empfaenger": z["empfaenger"],
                                   "fremdbeleg": z.get("fremdbeleg", ""), "rechnungsdatum": z.get("rechnungsdatum", ""),
                                   "erstmals": alt.get("erstmals") or jetzt.strftime("%Y-%m-%d %H:%M"),
                                   "zuletzt": jetzt.strftime("%Y-%m-%d %H:%M")}
            elif e in ("gesendet", "bereits gesendet", "ausgeschlossen") or e.startswith("ignoriert"):
                offen.pop(z["rnr"], None)
        # auch Faelle ausserhalb des Zeitfensters: erledigt (Sperrliste) oder inzwischen gesendet -> raus
        sperre_jetzt = lade_sperre(os.path.join(aus, "nie_senden.txt"))
        prot_jetzt = lade_protokoll(prot_pfad)
        for rnr_ in list(offen):
            if rnr_ in sperre_jetzt or prot_jetzt.get(rnr_) == "gesendet":
                offen.pop(rnr_, None)
        speichere_offen(aus, offen)
    except (OSError, ValueError) as e:
        log("Offene Fälle konnten nicht gespeichert werden: %s" % e)
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
RE_KUNDENNR = re.compile(r"Kunden\s*Nr\.?\s*:\s*(\d+)")
IGNORIERTE_KUNDEN = ("483911",)          # Ladengeschaeft / Thekenrechner (Sammelbelege an die eigene Firma)
OFFEN_DATEI = "offene_rechnungen.json"


def ist_problemfall(entscheidung):
    """Nicht versendet und braucht einen Menschen (kein Amazon, nicht zurueckgestellt/ausgeschlossen/schon gesendet)."""
    e = entscheidung or ""
    return e.startswith("FEHLER") or (e.startswith("NICHT gesendet") and "Amazon" not in e)


def _schreibe_json(pfad, daten):
    tmp = pfad + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(daten, f, ensure_ascii=False, indent=1)
    for i in range(6):
        try:
            os.replace(tmp, pfad)
            return
        except OSError:
            time.sleep(0.3 * (i + 1))
    with open(tmp, encoding="utf-8") as q, open(pfad, "w", encoding="utf-8") as z:
        z.write(q.read())
    os.remove(tmp)


def lade_offen(ausgabe_ordner):
    """{rnr: Fall} - offene (nicht versendete) Problem-Rechnungen, bis sie von Hand erledigt/gesendet werden."""
    try:
        with open(os.path.join(ausgabe_ordner, OFFEN_DATEI), encoding="utf-8") as f:
            d = json.load(f)
        return d if isinstance(d, dict) else {}
    except (OSError, ValueError):
        return {}


def speichere_offen(ausgabe_ordner, offen):
    _schreibe_json(os.path.join(ausgabe_ordner, OFFEN_DATEI), offen)


def markiere_erledigt(cfg, rnr, notiz="von Hand erledigt"):
    """Fall schliessen: Rechnungsnummer in nie_senden.txt (wird nie automatisch versendet) und aus den offenen Faellen."""
    aus = cfg["ausgabe_ordner"]
    with open(os.path.join(aus, "nie_senden.txt"), "a", encoding="utf-8") as f:
        f.write("%s  # %s %s\n" % (rnr, notiz, datetime.now().strftime("%Y-%m-%d %H:%M")))
    offen = lade_offen(aus)
    offen.pop(rnr, None)
    speichere_offen(aus, offen)


def pruefe_adresse(adresse):
    """'' wenn die Adresse fuer 'Doch senden' zulaessig ist, sonst der Grund."""
    a = (adresse or "").strip()
    if not re.fullmatch(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}", a):
        return "keine gültige E-Mail-Adresse"
    if ist_amazon_adresse(a.lower()) or "amazon" in a.lower():
        return "an Amazon-Adressen wird nie gesendet"
    if ist_firmenadresse(a):
        return "das ist eine eigene Firmenadresse"
    return ""


def sende_einzeln(cfg, rnr, empfaenger):
    """'Doch senden' aus dem Dashboard: eine offene Rechnung an empfaenger schicken (mit Sperre, Protokoll, BCC und
    Testadresse wie im normalen Lauf). Rueckgabe (ok, text)."""
    aus = cfg["ausgabe_ordner"]
    grund = pruefe_adresse(empfaenger)
    if grund:
        return False, grund
    with EinzelInstanz(os.path.join(aus, "lauf.lock")) as lock:
        if not lock.ok:
            return False, "es läuft gerade ein anderer Lauf - bitte gleich nochmal versuchen"
        fall = lade_offen(aus).get(rnr)
        if not fall:
            return False, "Rechnung %s ist kein offener Fall (mehr)" % rnr
        prot_pfad = os.path.join(aus, "gesendet.csv")
        if lade_protokoll(prot_pfad).get(rnr) == "gesendet":
            return False, "Rechnung %s wurde laut Protokoll schon versendet" % rnr
        pdf = fall.get("pdf") or ""
        if not pdf or not os.path.exists(pdf):
            return False, "Rechnungs-PDF nicht gefunden: %s" % pdf
        if rechnungsnummer_aus_text(pdf_text(pdf)) != rnr:
            return False, "PDF enthält nicht die Rechnung %s" % rnr
        kanal = fall.get("kanal") if fall.get("kanal") in (KANAL_SHOP, KANAL_EBAY, KANAL_SONST) else ""
        test_an = (cfg.get("test_empfaenger") or "").strip()
        msg = baue_mail(cfg, rnr, test_an or empfaenger.strip(), pdf, kanal, fall.get("fremdbeleg", ""),
                        fall.get("rechnungsdatum", ""))
        if test_an:
            msg.replace_header("Subject", "[TEST für %s] %s" % (empfaenger, msg["Subject"]))
        st_v, st_ok = ("testversuch", "test") if test_an else ("versuch", "gesendet")
        protokolliere(prot_pfad, rnr, st_v, kanal or "MANUELL", test_an or empfaenger, pdf, "manuell")
        try:
            sende_smtp(cfg, msg)
        except Exception as e:
            protokolliere(prot_pfad, rnr, "fehler", kanal or "MANUELL", empfaenger, pdf, "manuell")
            return False, "Senden fehlgeschlagen: %s" % e
        protokolliere(prot_pfad, rnr, st_ok, kanal or "MANUELL", test_an or empfaenger, pdf, "manuell")
        if not test_an:
            offen = lade_offen(aus)
            offen.pop(rnr, None)
            speichere_offen(aus, offen)
        return True, ("TEST gesendet an %s" % test_an) if test_an else "gesendet an %s" % empfaenger


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
    # Das Dashboard liest die Datei alle paar Sekunden - beim Ersetzen kann sie kurz offen sein (WinError 5/32)
    for i in range(6):
        try:
            os.replace(tmp, pfad)
            return
        except OSError:
            time.sleep(0.3 * (i + 1))
    try:
        with open(tmp, encoding="utf-8") as q, open(pfad, "w", encoding="utf-8") as z:
            z.write(q.read())
        os.remove(tmp)
    except OSError:
        pass                                  # Status ist nur Anzeige - der Versand selbst ist protokolliert


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
    # --- Sicherheit: Amazon-ADRESSE blockiert immer (Regel seit 2026-11-06a: die Adresse entscheidet) ---
    check("Shop-Nr aber Amazon-Adresse", k("25623", ["a@marketplace.amazon.de"]), KANAL_AMAZON)
    check("Amazon-Adresse anderer Domain amazon.com", k("25623", ["a@amazon.com"]), KANAL_AMAZON)
    check("eBay-Name aber Amazon-Marketplace-Mail", k("clerprem_de", ["a@marketplace.amazon.de"]), KANAL_AMAZON)
    check("Amazon-Adresse + zweite echte Adresse -> Amazon (nie senden)",
          k("305-9853395-4797960", ["a@marketplace.amazon.de", "k@gmx.de"]), KANAL_AMAZON)
    check("Amazon-Adresse nur im Rechnungstext", klassifiziere(pack("25623", ["a@b.de"]),
          (rechn % "25623") + "\nkontakt: x@marketplace.amazon.de\n")[0], KANAL_AMAZON)
    # --- Amazon-Merkmal OHNE Amazon-Adresse ---
    check("Amazon-Nr im Fremdbeleg, keine Adresse -> Amazon (ruhig, nie senden)",
          k("305-9853395-4797960", []), KANAL_AMAZON)
    check("Amazon-Nr im Fremdbeleg, zwei echte Adressen -> Amazon (nie senden)",
          k("305-9853395-4797960", ["a@b.de", "c@d.de"]), KANAL_AMAZON)
    check("Amazon-Nr im Fremdbeleg, EINE echte Adresse -> darf gesendet werden (Matthias 06.11.)",
          k("305-9853395-4797960", ["kunde@gmx.de"]), KANAL_SONST)
    check("Wort Amazon in Packliste, echte Adresse -> SHOP", k("25623", ["a@b.de"], ex="per Amazon bestellt"), KANAL_SHOP)
    check("Amazon-Nr irgendwo in der Packliste, eBay-Adresse -> EBAY",
          k("clerprem_de", ["x@members.ebay.com"], ex="Ref 305-9853395-4797960"), KANAL_EBAY)
    check("Amazon-Nr nur auf der Rechnung, Packliste anders -> unklar (Fremdbeleg weicht ab)",
          klassifiziere(pack("clerprem_de", ["x@members.ebay.com"]), rechn % "305-9853395-4797960")[0], KANAL_UNKLAR)
    check("Grund nennt Amazon-Merkmal bei echter Adresse",
          "Amazon-Merkmal" in klassifiziere(pack("305-9853395-4797960", ["kunde@gmx.de"]),
                                            rechn % "305-9853395-4797960")[1], True)
    # --- Zweifel -> UNKLAR ---
    check("keine Mail", k("25623", []), KANAL_UNKLAR)
    check("zwei Mails", k("25623", ["a@b.de", "c@d.de"]), KANAL_UNKLAR)
    check("kein Fremdbeleg, echte Adresse -> SONSTIGE", klassifiziere(pack("", ["a@b.de"]), rechn % "")[0], KANAL_SONST)
    check("normaler Name + normale Mail (kein eBay) -> SONSTIGE", k("hans_mueller", ["h@gmx.de"]), KANAL_SONST)
    check("Zahl zu kurz (3 Stellen) + Mail -> SONSTIGE", k("123", ["h@gmx.de"]), KANAL_SONST)
    check("Zahl zu lang (7 Stellen) + Mail -> SONSTIGE", k("1234567", ["h@gmx.de"]), KANAL_SONST)
    check("eBay-Mail mit Zahl als Fremdbeleg -> EBAY", k("25623", ["x@members.ebay.com"]), KANAL_EBAY)
    check("Shop-Nr + Mail anderer Domain mit 'members.' -> SHOP", k("25623", ["x@members.example.org"]), KANAL_SHOP)
    check("Fremdbeleg Rechnung weicht ab -> unklar", klassifiziere(pack("25623", ["a@b.de"]), rechn % "25999")[0], KANAL_UNKLAR)
    check("keine Packliste -> unklar", klassifiziere(None, "x")[0], KANAL_UNKLAR)
    check("keine Packliste + Amazon-Nr auf der Rechnung -> Amazon (ruhig)",
          klassifiziere(None, "Rechnungsdatum: 16.11.2026\n304-0312378-5789914\n")[0], KANAL_AMAZON)
    check("interner Beleg Ladengeschaeft wird ignoriert",
          ist_kundenrechnung("Gasecenter, Karlsruher Strasse 3, 86156 Augsburg (GVZ)\nGasecenter Augsburg\n"
                             "MONATLICHE VERKÄUFE LADENGESCHÄFT\n02.11 - 12.11.2026\nRechnung Nr. : 1713932\n")[0], False)
    check("normale Rechnung bleibt Rechnung",
          ist_kundenrechnung("Gasecenter, Karlsruher Strasse 3, 86156 Augsburg (GVZ)\nFritz Muster\nHauptstr. 1\n"
                             "Rechnung Nr. : 1713966\n")[0], True)
    check("Rechnung ohne Fremdbeleg-Zeile bleibt ok (Vorlagentext wird nicht als Fremdbeleg gelesen)",
          klassifiziere(pack("25623", ["a@b.de"]),
                        "Rechnungsdatum: 20.10.2026\nRechnungsdatum = Lieferdatum falls\nnicht anders angegeben\n")[0], KANAL_SHOP)
    check("eBay-Name + normale Mail + Shop-Nr (Historie, 2 Fremdbeleg-Zeilen) -> echte Adresse -> SONSTIGE",
          klassifiziere(pack("svestei_0", ["x@gmail.com"], ex="25662"), rechn % "svestei_0")[0], KANAL_SONST)
    # --- Mischfall: eBay-Kunde bestellt spaeter im Shop (Rechnung: Shop-Nr, Packliste: alter eBay-Name) ---
    check("Mischfall Rechnung Shop-Nr + Packliste eBay-Name + echte Adresse -> SHOP",
          klassifiziere(pack("mihen2013_7", ["k@gmx.de"]), rechn % "25956")[0], KANAL_SHOP)
    check("Mischfall: Grund nennt Mischfall + Rechnungsnummer des Shops",
          ("Mischfall" in klassifiziere(pack("mihen2013_7", ["k@gmx.de"]), rechn % "25956")[1],
           "25956" in klassifiziere(pack("mihen2013_7", ["k@gmx.de"]), rechn % "25956")[1]), (True, True))
    check("Mischfall: Empfaenger ist die Kundenadresse",
          klassifiziere(pack("mihen2013_7", ["k@gmx.de"]), rechn % "25956")[2], "k@gmx.de")
    check("Mischfall aber eBay-Relay-Adresse -> unklar (koennte veraltet sein)",
          klassifiziere(pack("mihen2013_7", ["x@members.ebay.com"]), rechn % "25956")[0], KANAL_UNKLAR)
    check("Mischfall aber zwei Adressen -> unklar",
          klassifiziere(pack("mihen2013_7", ["k@gmx.de", "m@web.de"]), rechn % "25956")[0], KANAL_UNKLAR)
    check("Mischfall aber Amazon-Adresse -> Amazon",
          klassifiziere(pack("mihen2013_7", ["x@marketplace.amazon.de"]), rechn % "25956")[0], KANAL_AMAZON)
    check("umgekehrt: Rechnung eBay-Name, Packliste Shop-Nr -> unklar",
          klassifiziere(pack("25956", ["k@gmx.de"]), rechn % "mihen2013_7")[0], KANAL_UNKLAR)
    check("Rechnung Shop-Nr, Packliste Amazon-Nr als Fremdbeleg -> unklar",
          klassifiziere(pack("305-9853395-4797960", ["k@gmx.de"]), rechn % "25956")[0], KANAL_UNKLAR)
    check("Rechnung Shop-Nr, Packliste ohne Fremdbeleg -> unklar",
          klassifiziere(pack("", ["k@gmx.de"]), rechn % "25956")[0], KANAL_UNKLAR)
    check("Rechnung andere Zahl als Packliste (beide Shop-Nr) -> unklar",
          klassifiziere(pack("25623", ["k@gmx.de"]), rechn % "25956")[0], KANAL_UNKLAR)
    check("Mailtext im Mischfall nennt die Shop-Nummer der RECHNUNG",
          fremdbeleg_fuer_mail(KANAL_SHOP, pack("mihen2013_7", ["k@gmx.de"]), rechn % "25956"), "25956")
    check("Mailtext normaler Shop-Fall nennt die Packliste-Nummer",
          fremdbeleg_fuer_mail(KANAL_SHOP, pack("25623", ["k@gmx.de"]), rechn % "25623"), "25623")
    check("Mailtext eBay nennt den eBay-Namen",
          fremdbeleg_fuer_mail(KANAL_EBAY, pack("clerprem_de", ["x@members.ebay.com"]), rechn % "clerprem_de"), "clerprem_de")
    check("SONSTIGE: Platzhalter {bestellung} neutral", platzhalter("1", KANAL_SONST, "abc")["{bestellung}"], "Ihre Bestellung")
    check("eigene Firmenadresse zaehlt nicht als Kunde",
          klassifiziere(pack("25623", ["a@b.de"], ex="Email@Gasecenter-Augsburg.de"), rechn % "25623")[0], KANAL_SHOP)
    # --- HTML-Mail mit Vorlage ---
    import tempfile as _tf
    _d = _tf.mkdtemp()
    _pdf = os.path.join(_d, "r.pdf")
    open(_pdf, "wb").write(b"%PDF-1.4 test")
    _vor = os.path.join(_d, "v.html")
    open(_vor, "w", encoding="utf-8").write('<p>Danke f&uuml;r {bestellung}, Rechnung {rnr} vom {rechnungsdatum}</p>'
                                            '<img src="data:image/png;base64,iVBORw0KGgo=">')
    _cfg = {"absender": "r@example.org", "betreff": "Rechnung {rnr}", "text": "Rechnung {rnr} / {bestellung}",
            "html_vorlage": _vor}
    _m = baue_mail(_cfg, "1710010", "k@example.org", _pdf, KANAL_SHOP, "25623", "26.10.2026")
    _roh = _m.as_string()
    check("HTML-Mail: Betreff", _m["Subject"], "Rechnung 1710010")
    check("HTML-Mail: Teile (Text, HTML, Logo, PDF)",
          [t.get_content_type() for t in _m.walk() if not t.is_multipart()],
          ["text/plain", "text/html", "image/png", "application/pdf"])
    _htm = next(t for t in _m.walk() if t.get_content_type() == "text/html").get_content()
    check("HTML-Mail: Platzhalter ersetzt, Logo als cid",
          ("Bestellung Nr. 25623 in unserem Onlineshop" in _htm, "1710010" in _htm, "26.10.2026" in _htm,
           "cid:" in _htm, "data:image" in _htm, "{" in _htm), (True, True, True, True, False, False))
    check("HTML-Mail: eBay-Text", "Ihre Bestellung bei eBay" in
          baue_mail(_cfg, "1", "k@members.ebay.com", _pdf, KANAL_EBAY, "name_x").get_body(("plain",)).get_content(), True)
    try:
        baue_mail(dict(_cfg, html_vorlage=os.path.join(_d, "fehlt.html")), "1", "k@x.de", _pdf, KANAL_SHOP, "1")
        check("Fehlende Vorlage -> keine Mail", "gebaut", "Fehler")
    except OSError:
        check("Fehlende Vorlage -> keine Mail", "Fehler", "Fehler")
    # --- Rechnungsart ---
    check("Rechnung ok", ist_kundenrechnung("Gasecenter\nRechnung Nr. : 1710008\n")[0], True)
    check("Rechnungsnummer hinter Ortstext derselben Zeile (vierzeilige Adresse, real 1706130)",
          rechnungsnummer_aus_text("Herrn\nA B\n851309770\nPackstation 150\nDE-28719 Bremen Rechnung Nr. : 1706130\nKunden Nr. : 1\n"), "1706130")
    check("Fremdbeleg-Block bei verschmolzener Kopfzeile",
          fremdbeleg_und_mails("DE-28719 Bremen Rechnung Nr. : 1706130\n25623\nkunde@example.org\nSendungsgewicht : 1\n"),
          ("25623", ["kunde@example.org"]))
    check("gekuerzter langer eBay-Name auf der Rechnung gilt als gleich",
          klassifiziere(pack("abenteuer--in--mecklenburg--vorpommern", ["x@members.ebay.com"]),
                        rechn % "abenteuer--in--mecklenburg--vorpomm")[0], KANAL_EBAY)
    check("kurzer Namensanfang gilt NICHT als gleich",
          klassifiziere(pack("abenteuer", ["x@members.ebay.com"]), rechn % "abent")[0], KANAL_UNKLAR)
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
                                 ("1710013", "hans_x", "h@gmx.de c@gmx.de"),    # zwei Adressen -> unklar
                                 ("1710015", "25700", "k2@example.org")):
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
        check("Bericht: nie_senden.txt greift (ausgeschlossen, nicht unklar)", je["1710015"]["entscheidung"], "ausgeschlossen")
        check("Bericht: unter Startnummer fehlt", "1709999" in je, False)
        check("Bericht: Gutschrift ignoriert", "1110001" in je, False)
        check("Bericht: internes Formular (Lagerort) blockiert", je["1710020"]["entscheidung"], "NICHT gesendet")
        check("Bericht: nichts gesendet", len(gefangen), 0)
        res = lauf(cfg, "entwurf")
        check("Entwurf: 3 .eml", sorted(os.path.basename(p) for p in glob.glob(os.path.join(aus, "Entwuerfe", "*", "*.eml"))),
              ["Rechnung_1710010_SHOP.eml", "Rechnung_1710011_EBAY.eml", "Rechnung_1710016_SHOP.eml"])
        res = lauf(dict(cfg, test_empfaenger="pruef@example.org"), "senden")
        check("Testversand: 3 Mails, alle an die Testadresse, Betreff [TEST",
              (len(gefangen), all("To: pruef@example.org" in g for g in gefangen),
               all("[TEST" in g for g in gefangen)), (3, True, True))
        check("Testversand blockiert den echten Versand nicht",
              sorted(set(lade_protokoll(os.path.join(aus, "gesendet.csv")).values())), ["test"])
        del gefangen[:]
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
        # --- offene Faelle ---
        off = lade_offen(aus)
        check("Offene Fälle: unklare Rechnungen drin, Amazon/gesendet/ausgeschlossen nicht",
              (("1710013" in off, "1710014" in off), "1710012" in off, "1710010" in off, "1710015" in off),
              ((True, True), False, False, False))
        check("Doch senden: Amazon-Adresse abgelehnt", sende_einzeln(cfg, "1710013", "x@marketplace.amazon.de")[0], False)
        check("Doch senden: ungültige Adresse abgelehnt", sende_einzeln(cfg, "1710013", "kein-at-zeichen")[0], False)
        n_vor = len(gefangen)
        ok_ds, txt_ds = sende_einzeln(cfg, "1710013", "h@gmx.de")
        check("Doch senden: gesendet, Mail raus, Fall weg, Protokoll gesendet",
              (ok_ds, len(gefangen) - n_vor, "1710013" in lade_offen(aus),
               lade_protokoll(os.path.join(aus, "gesendet.csv")).get("1710013")), (True, 1, False, "gesendet"))
        check("Doch senden: zweites Mal abgelehnt", sende_einzeln(cfg, "1710013", "h@gmx.de")[0], False)
        markiere_erledigt(cfg, "1710014")
        check("Erledigt: aus offenen Fällen raus und in nie_senden.txt", ("1710014" in lade_offen(aus),
              "1710014" in lade_sperre(os.path.join(aus, "nie_senden.txt"))), (False, True))
        res = lauf(cfg, "bericht")
        check("Erledigt: im nächsten Lauf 'ausgeschlossen', nicht wieder offen",
              ({z["rnr"]: z["entscheidung"] for z in res["ergebnisse"]}.get("1710014"), "1710014" in lade_offen(aus)),
              ("ausgeschlossen", False))
        # Versuch ohne Bestaetigung + "Erledigt" -> bleibt erledigt (real 1714151)
        check("Versuch ohne Bestaetigung ist offener Fall", "1710099" in lade_offen(aus), True)
        markiere_erledigt(cfg, "1710099")
        res = lauf(cfg, "bericht")
        check("Versuch + Erledigt: im nächsten Lauf 'ausgeschlossen', nicht wieder offen",
              ({z["rnr"]: z["entscheidung"] for z in res["ergebnisse"]}.get("1710099"), "1710099" in lade_offen(aus)),
              ("ausgeschlossen", False))
        # alter offener Fall ausserhalb des Zeitfensters, steht schon in nie_senden.txt -> naechster Lauf traegt aus
        off_ = lade_offen(aus)
        off_["1700001"] = {"rnr": "1700001", "pdf": "", "kanal": "", "grund": "alt", "entscheidung": "NICHT gesendet",
                           "empfaenger": "", "fremdbeleg": "", "rechnungsdatum": "", "erstmals": "2026-11-19 14:19",
                           "zuletzt": "2026-11-19 14:19"}
        off_["1700002"] = dict(off_["1700001"], rnr="1700002")
        speichere_offen(aus, off_)
        with open(os.path.join(aus, "nie_senden.txt"), "a", encoding="utf-8") as f_:
            f_.write("1700001  # Test alt erledigt\n")
        lauf(cfg, "bericht")
        off_ = lade_offen(aus)
        check("Alter offener Fall in nie_senden.txt: beim Prüfen ausgetragen, anderer alter Fall bleibt",
              ("1700001" in off_, "1700002" in off_), (False, True))
        off_.pop("1700002", None)
        speichere_offen(aus, off_)
        # eigene Kundennummer (Ladengeschaeft/Theke, ohne Packliste) -> still ignorieren
        mkpdf(os.path.join(rd, "Packliste Nr. 1710050 20261116.pdf"),
              ["Gasecenter, Karlsruher Strasse 3", "Max Theke", "Rechnung Nr. : 1710050", "Kunden Nr. : 483911",
               "Rechnungsdatum: 16.11.2026", "Rechnung"])
        res = lauf(cfg, "bericht")
        check("Kundennummer 483911 -> ignoriert (nicht in der Liste, kein offener Fall)",
              ("1710050" in {z["rnr"] for z in res["ergebnisse"]}, "1710050" in lade_offen(aus),
               any("483911" in g for _, g in res["ignoriert"])), (False, False, True))
        # Thekenrechner-Ausbuchung: eigene Adresse als Kunden-E-Mail -> still ignorieren
        rechnung("1710040", "Theke")
        packl("1710040", "Theke", "email@gasecenter-augsburg.de")
        res = lauf(cfg, "bericht")
        z40 = {z["rnr"]: z for z in res["ergebnisse"]}.get("1710040", {})
        check("Eigene Adresse (Theke) -> ignoriert, kein offener Fall, nicht unklar",
              (z40.get("entscheidung"), "1710040" in lade_offen(aus), _status_aus_ergebnis(res, "x")["unklar"] >= 0),
              ("ignoriert (eigene Adresse)", False, True))
        check("Firmenadresse in der Fusszeile zaehlt nicht als eigene Kundenadresse",
              eigene_adresse_im_kopf("Rechnung Nr. : 1\n25623\nkunde@example.org\nSendungsgewicht : 1\nA\nB\nC\n"
                                     "Gerichtsstand Augsburg Email@Gasecenter-Augsburg.de"), "")
        # verwaiste Sperre (Prozess beendet) wird uebernommen, eine lebende nicht
        lockp = os.path.join(aus, "lauf.lock")
        open(lockp, "w").write("999999")                       # PID, die es nicht gibt
        with EinzelInstanz(lockp) as l1:
            check("Verwaiste lauf.lock wird übernommen", l1.ok, True)
        open(lockp, "w").write(str(os.getpid()))               # lebender Prozess
        with EinzelInstanz(lockp) as l2:
            check("Lebende Sperre bleibt bestehen", l2.ok, False)
        os.remove(lockp)
        # Maximalzahl
        if os.path.exists(os.path.join(aus, "gesendet.csv")):
            os.remove(os.path.join(aus, "gesendet.csv"))
        res = lauf(dict(cfg, max_mails_pro_lauf=1), "entwurf")
        check("Maximalzahl Mails", res["entworfen"], 1)
        check("Maximalzahl: Rest 'zurückgestellt', nicht unklar",
              (sum(1 for z in res["ergebnisse"] if z["entscheidung"] == "zurückgestellt") > 0,
               _status_aus_ergebnis(res, "manuell")["unklar"] == sum(1 for z in res["ergebnisse"]
                                                                       if z["entscheidung"].startswith("NICHT gesendet")
                                                                       and "Amazon" not in z["entscheidung"])),
              (True, True))
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
