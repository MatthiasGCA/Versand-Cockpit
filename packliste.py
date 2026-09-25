# -*- coding: utf-8 -*-
"""
packliste.py - Kommissionierliste aus Gasecenter-Ausgangsrechnungen (Amicron-PDF)
=================================================================================
Liest einen Ordner mit Ausgangsrechnungen (PDF, Amicron "entzerrtes" Layout mit
Einzelpreis/G-Preis/Rechnungsbetrag und Lagerort-Feld) und erzeugt:

  1. eine ZWEITEILIGE PDF-Kommissionierliste
     - Teil A "Sammelliste": je Lagerkategorie die Artikel aggregiert
       (Gesamtmenge ueber alle Rechnungen), mit Abhak-Kaestchen.
     - Teil B "Packuebersicht je Rechnung": Kopf mit Rechnungsnr., Kd-Nr.,
       Datum, Empfaenger und Code128-Barcode der Rechnungsnummer, darunter
       die Positionen mit Abhak-Kaestchen.
  2. fuenf Bruecken-CSVs (Semikolon, Kopfzeile, UTF-8; jeweils mit dem
     Rechnungsdatum als LETZTER Spalte fuer den Verfall, siehe
     BRUECKEN_CSV_MERGE_TAGE; werden mit dem Bestand frueherer Laeufe
     zusammengefuehrt):
       post_zuordnung.csv    Rechnungsnummer;Name;Strasse;Hausnummer;PLZ;Datum
                             -> scan_druck.py ordnet Deutsche-Post-Labels (ohne
                                Rechnungsnummer) ueber die Lieferadresse zu
       sammel_zuordnung.csv  Sammelcodes (SAM-n) -> Bestellungen, EAN, Menge je
                             Bestellung (Sammeldruck-Deckblatt)
       mengen_zuordnung.csv  Rechnungsnummer;Scananzahl;Datum (Mehrfach-Scan)
       ean_zuordnung.csv     Rechnungsnummer;EAN;Bezeichnung;Anzahl;Datum
       wc_bestellnummern.csv Rechnungsnummer;WC_Bestellnummer;... (nur fuer
                             wc_sendungsnummer_sync.py, nicht fuer scan_druck.py)

 Aufruf:
    python packliste.py <rechnungs_ordner> "<stammdaten_oder_leer>" <ausgabe.pdf>
Beispiel:
    python packliste.py /mnt/user-data/uploads "" /mnt/user-data/outputs/Pickliste.pdf

Betragsabgleich (zwei Sicherheitspruefungen):
  - je Position:  Menge x Einzelpreis == G-Preis
  - je Rechnung:  Summe G-Preise == Rechnungsbetrag
Rechnungen, deren Abgleich nicht aufgeht, werden trotzdem aufgenommen, aber im
Protokoll und mit einem Hinweis in der Packuebersicht deutlich markiert.

Kategorie-Zuordnung allein aus dem Lagerort-Feld der Packliste:
  Topseller, Eigenfertigung, Schlauchlager, Palettenlager, Poolchemie,
  Kleinteile, Regallager, DPD / Warenpost, Allgemein  (Regeln unten in classify()).
"""

import os
import re
import sys
import csv
import glob
import io
import shutil
from datetime import datetime

import pdfplumber
from pypdf import PdfReader, PdfWriter
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Flowable,
    KeepTogether, PageBreak,
)
from reportlab.graphics.barcode import code128
from reportlab.graphics.shapes import Drawing

# ----------------------------------------------------------------------------
# KONFIGURATION
# ----------------------------------------------------------------------------

VERSION = "2026-09-25a"          # Versionsschema: JJJJ-MM-TT + Kleinbuchstabe je
# 2026-09-23b: Neue Markierungszeile "<N>-je-Paket" (z.B. "1-je-Paket",
#   "3-je-Paket", JE_PAKET_RE) - wie Lagerort/Fach-Artikel/Kennung eine eigene
#   linksbuendige Zeile im Positionsblock, NICHT Teil der Bezeichnung. Neues
#   Feld "je_paket" je Position. Rein additiv (Pickliste/Bruecken-CSVs
#   unveraendert) - ausgewertet vom Carrier-Dashboard (carrier_regeln.
#   je_paket_aufteilung()) fuer automatische DHL-Paketaufteilung bei
#   Artikeln, die aus Gewichtsgruenden nicht beliebig gebuendelt werden
#   duerfen. Mit Matthias abgestimmt.
# 2026-09-23a: kdnr-Erkennung priorisiert jetzt WC-Bestellnummer (4-6 Ziffern)
#   bzw. Amazon-Bestellnummer ("nnn-nnnnnnn-nnnnnnn") vor einem sonstigen
#   Marktplatz-Namen (z.B. eBay-Name), falls mehrere Fremdbeleg-Zeilen unter
#   "Rechnung Nr." stehen. Anlass: Rechnung 1705582 - Kunde hat Historie auf
#   Amazon UND eBay, Amicron druckt dann beide Zeilen (erst eBay-Name, dann
#   Amazon-Bestellnummer); die alte Logik nahm einfach die erste Zeile und
#   waehlte damit den weniger eindeutigen eBay-Namen statt der Bestellnummer.
#   Ohne mehrere Kandidatenzeilen (Normalfall) unveraendertes Verhalten -
#   gegen alle 145 echten Test-Rechnungen im Carrier-Dashboard-Pool
#   gegengeprueft (111 Amazon-Bestellnummern, 34 sonstige Namen, 0 Fehler).
# 2026-09-21d: Carrier-Dashboard (neu: carrier_dashboard.py/carrier_regeln.py). parse_pdf()
#   liefert zusaetzlich "sendungsgewicht" (kg, Kopfzeile "Sendungsgewicht"), "adresse_zeilen"
#   (alle Zeilen der Lieferadresse, sonst Rechnungsadresse) und je Position "kennungen"
#   (Pax1/Pox1/Brx1/Wapo als eigene Zeile, KENNUNG_RE). Rein additiv - Pickliste und
#   Bruecken-CSVs unveraendert (Kennungszeile wird wie Lagerort/Fach-Marker aus der
#   Bezeichnung ferngehalten, ist aber sonst ohne Wirkung auf die Pickliste).
# Aenderung am selben Tag (erste = a, dann b, c ...; neuer Tag beginnt wieder bei a).
# 2026-09-21c: Menge-Spalte breiter (80 -> 100 pt, Kommissionierliste UND
#   Packuebersicht) und MengeFlow verkleinert die hervorgehobene Menge
#   automatisch, wenn Text + Kreis trotzdem nicht in die Spalte passen. Anlass:
#   Rechnung 1705307, "2 Schuber" ragte mit dem roten Kreis ueber die Spalte
#   und ueberdeckte Lagerort bzw. Checkbox. Nur Layout, keine Mengenlogik.
#   (Hinweis: bei N-Fach-Artikeln ersetzt effektive_menge() die Einheit weiter
#   bewusst durch "Stück", z.B. Schuber -> "12 Stück" bei 4-Fach x 3.)
# 2026-09-21b: Projektpruefung, Nachzuegler: _fach_token_re()/_gewicht_token_re()
#   pruefen jetzt (?<![\d,.]) statt (?<!\d) - eine Dezimalzahl in der
#   Bezeichnung ("1,5 kg", "1,2 Stück") wird sonst hinter dem Komma angebrochen
#   ("1,GESAMTMENGE"). Ausserdem veraltete Kommentare zu den seit 2026-08-21i
#   entfernten Packungsartikeln bereinigt und den Modul-Kopf um alle fuenf
#   Bruecken-CSVs ergaenzt. Rein kosmetisch/haertend, keine Logikaenderung fuer
#   die bisherigen echten Bezeichnungen (von Hand gegen 5,00 kg / 10 Stück / 6x
#   / 2 Packungen gegengeprueft).
# 2026-09-21a: Bei der Projektpruefung gefunden: _hausnr() ignoriert jetzt
#   Satzzeichen hinter der Hausnummer ("Clara Zetkin Strasse 358," -> "358",
#   real in post_zuordnung.csv beobachtet, Hausnummer war dort leer und das
#   Post-Matching lief nur ueber den Namen). Gleiche Aenderung in scan_druck.py
#   (beide Dateien gemeinsam deployen, sonst weichen Schluessel voneinander ab).
# 2026-09-18a: WooCommerce-Anbindung (Pause seit 2026-08-28, s.u.) wieder
#   aufgenommen - reine Neudatierung dieses Eintrags, am Code selbst hat sich
#   seit der Pause nichts mehr geaendert (gegen HEAD/2026-09-04a gegengeprueft,
#   Diff bleibt sauber auf die unten beschriebenen Stellen beschraenkt).
# 2026-08-28a: Neue fuenfte Bruecken-CSV "wc_bestellnummern.csv"
#   (schreibe_wc_bestellnummern_csv): Rechnungsnummer -> WooCommerce-
#   Bestellnummer, NUR fuer Onlineshop-Rechnungen (Ziffernblock 4-6 Stellen
#   unter "Rechnung Nr." + Kundenmail nicht von Amazon/eBay). Quelle fuer das
#   separate Skript wc_sendungsnummer_sync.py, das die Carrier-Sendungsnummern
#   per REST-API in Shiptastic eintraegt. parse_pdf() liefert dazu neu das
#   Feld "email"; die kdnr-Erkennung ueberspringt jetzt zusaetzlich eine Zeile,
#   die exakt der Rechnungsnummer entspricht (Schutz vor falscher kdnr, falls
#   "Rechnung Nr." und Nummer je auf eigener Zeile stehen). Rein additiv -
#   Pickliste-PDF und die vier bestehenden CSVs unveraendert.
#   [Diese Aenderung ist weiterhin NICHT committet - bewusst getrennt von den
#   produktiven Bugfixes unten, bis die WC-Anbindung im Echtbetrieb getestet ist.]
# 2026-09-04a: KRITISCH - kg-Multipack-Erkennung (2026-08-21j) hat in der
#   Praxis nie gegriffen (real gemeldet und an Rechnung 1703025/WSG2-1,6-5
#   nachvollzogen). Ursache: die Markierungszeile "5kg-Multipack" steht -
#   genau wie "Lagerort:"/"N-Fach-Artikel" - links-buendig (x0 < ART_MAX) im
#   Positionsblock und wurde von parse_block() beim Zusammensetzen der
#   Bezeichnung schlicht VERWORFEN, noch bevor artikel_gewicht() den Text
#   je zu sehen bekam - kein Formatierungsproblem der Regex, sondern eine
#   fehlende Zeilen-Erkennung. Jetzt wird "<X>,<Y>kg-Multipack" in
#   parse_block() wie "N-Fach-Artikel" als eigene Markierungszeile erkannt
#   und in einem neuen Feld "gewicht" gemerkt; artikel_gewicht(p) liest das
#   jetzt direkt aus der Position statt (erfolglos) im fertigen
#   Bezeichnungstext zu suchen. Die tatsaechlich sichtbare Gewichtsangabe im
#   Bezeichnungstext (z.B. "5,00 kg", eigenstaendig neben dem Marker - analog
#   zu "10 Stück" neben "10-Fach-Artikel") wird ueber den neuen Helfer
#   _gewicht_token_re() gefunden und durch GESAMTMENGE ersetzt.
# 2026-09-01b: Auf Kundenwunsch die zweite Ursache fuer unbegrenzt wachsende
#   Bruecken-CSVs beseitigt (die aus 2026-09-01a bekannte, aber eigenstaendige
#   "immer frisch"-Schwaeche): _bruecken_datum_frisch() wertete ein leeres/
#   unlesbares Rechnungsdatum bisher als dauerhaft frisch ("lieber zu lange
#   behalten als zu frueh verlieren"). Da eine gemergte Zeile ihr Datum nie
#   aendert, blieb ein Eintrag ohne Datum dadurch fuer IMMER in der Datei -
#   genau die Bedingung, unter der der 2026-09-01a-Quotierungsbug ungebremst
#   exponentiell eskalieren konnte. Jetzt gilt ein leeres/unlesbares Datum
#   als NICHT frisch (wird verworfen). Damit das keinen echten Auftrag zu
#   frueh verliert, tragen alle Schreibfunktionen jetzt IMMER ein gueltiges
#   Datum ein (neue Helfer _csv_datum() bzw. in merge_sammelgruppen() fuer
#   Sammelgruppen, die kein einzelnes Rechnungsdatum haben) - notfalls das
#   heutige, statt es leer zu lassen.
# 2026-09-01a: KRITISCH - Absturz mit MemoryError beim Schreiben von
#   ean_zuordnung.csv behoben (real beobachtet, Datei war auf ~940 MB
#   angewachsen). Ursache: _lies_bestehende_zeilen() las mit einem naiven
#   zeile.split(";") statt echtem CSV-Parsing - ein Bezeichnungstext mit
#   Semikolon UND Anfuehrungszeichen (Rechnung 1700713: 'Regler 8,0kg/h;
#   0,5-4bar; Kombi x G 3/8"LH-KN; ...') wurde dadurch bei jedem Lese-/
#   Schreibzyklus erneut falsch requotiert - die Anfuehrungszeichen
#   verdoppelten sich exponentiell ueber ca. 30 Laeufe hinweg. Jetzt echtes
#   CSV-Parsing (csv.reader). Betrifft auch scan_druck.py (VERSION
#   2026-09-01a dort), das dieselbe Schwaeche in seinen vier CSV-Lesefunk-
#   tionen hatte. Die korrupte Live-Datei wurde manuell bereinigt.
# 2026-08-21n: Auf Kundenwunsch "Rolle"/"Rollen" als weiteres Mengen-Token in
#   _fach_token_re aufgenommen (zusaetzlich zu Kartusche(n)/Flasche(n)/
#   Dose(n)/Eimer aus 2026-08-21m).
# 2026-08-21m: Auf Kundenwunsch weitere Gebinde-Einheiten in _fach_token_re
#   aufgenommen, die im Katalog als Mengen-Token in der Bezeichnung eines
#   Fach-Artikels vorkommen koennen: Kartusche(n), Flasche(n), Dose(n), Eimer
#   (zusaetzlich zu x/Stück/Packung(en) aus 2026-08-21c/f/l).
# 2026-08-21l: Bei der Generalprobe mit echten Rechnungen gefunden (1701663/
#   BP-TILL-2, "2-Fach-Artikel"): die Menge wurde korrekt auf 2 umgerechnet,
#   im Bezeichnungstext blieb aber "2 Packungen" stehen statt durch
#   GESAMTMENGE ersetzt zu werden - _fach_token_re kannte bisher nur die
#   Muster "<n>x" und "<n> Stück", nicht "<n> Packung(en)". Jetzt zusaetzlich
#   erkannt, siehe Docstring von _fach_token_re.
# 2026-08-21k: Bei einer erneuten Projektpruefung gefundener, verifizierter
#   Aufraeum-Punkt behoben: die vier Konstanten EAN_LO/EAN_HI, MENGE_LO/
#   MENGE_HI, EP_LO/EP_HI, GP_LO/GP_HI (Spaltenbaender fuer Menge/Einzelpreis/
#   Gesamtpreis) wurden nirgends mehr gelesen - seit der Umstellung auf
#   _preiszahlen() werden diese drei Spalten layout-unabhaengig ueber
#   NUM_MIN_X erkannt, nicht mehr ueber feste x-Baender. Der Kommentar direkt
#   darueber sagte aber weiterhin, man solle bei einer neuen Layout-Variante
#   "hier nachjustieren" - das haette ins Leere gezielt, waere aber nicht
#   sofort aufgefallen. Tote Konstanten entfernt, Kommentar auf den
#   tatsaechlich aktiven Mechanismus (NUM_MIN_X/EINH_LO/EINH_HI) umgestellt.
# 2026-08-21j: Gewichtsartikel-Erkennung (artikel_gewicht) von Kundenwunsch
#   21i (kg-Text muss zur letzten Artikelnr-Zahl passen) auf einen rein
#   EXPLIZITEN Text-Marker umgestellt: "<X>,<Y>kg-Multipack" im
#   Bezeichnungstext (z.B. "5,0kg-Multipack" fuer 5 kg Verkaufsmenge,
#   "0,25kg-Multipack" fuer 0,25 kg) - analog zur "<N>-Fach-Artikel"-
#   Markierung bei Set-Artikeln. Grund: die Artikelnr-basierte Inferenz war
#   trotz der 21i-Absicherung weiterhin ein grundsaetzliches Fehlausloeser-
#   Risiko (z.B. "AutoPR-11" mit "...fuer 11-kg-Flaschen" im Text), das der
#   Kunde lieber ganz vermeiden wollte, da korrekte Mengen oberste Prioritaet
#   haben. Ab sofort daher KEINE Ableitung mehr aus der Bezeichnung/
#   Artikelnummer bei Gewichtsartikeln - nur der explizite "-Multipack"-
#   Marker (durch den Kunden gezielt in die betroffenen Artikeltexte
#   eingepflegt) loest die Erkennung noch aus. GEWICHT_RE/PACK_ART_RE-
#   Artikelnr-Logik aus 21i komplett entfernt; artikel_gewicht(bez) nimmt
#   jetzt nur noch die Bezeichnung entgegen (kein "art"-Parameter mehr).
# 2026-08-21i: Zwei auf Kundenwunsch praezisierte Regeln fuer effektive_menge()
#   (welche Positionen die Pickliste-Menge/-Bezeichnung veraendern duerfen):
#   1) Der Packungsartikel-Mechanismus (artikel_packung/PACK_ART_RE/
#      _pack_token_re: Artikelnr "<Zahl>/<N>" + "<N> Stück" im Text) wurde
#      KOMPLETT ENTFERNT. Grund: irrefuehrende Artikelbezeichnungen wurden
#      dadurch faelschlich korrigiert (real beobachtet an Rechnung 1701626,
#      299/6: "6 Flaschenkappen" wurde als 6er-Packung erkannt, obwohl keine).
#      Ab sofort aendert AUSSCHLIESSLICH die explizite "<N>-Fach-Artikel"-
#      Markierung (Feld "fach", FACH_ARTIKEL_RE) die Menge/Bezeichnung eines
#      Set-Artikels - keine Ableitung mehr aus Artikelnr-Form oder Text.
#   2) Die Gewichtsartikel-Erkennung (artikel_gewicht, unveraendert als
#      Ausnahme mit Inferenz erhalten, "funktioniert sehr gut") wurde
#      erweitert: Artikel wie "WSG2-1,6-5" (letztes Artikelnr-Segment eine
#      REINE GANZZAHL statt Komma-Dezimalwert) werden jetzt ebenfalls erkannt,
#      sofern der Text dazu passend "5,00kg" (mit Komma-Nachkommastelle)
#      enthaelt. Die Absicherung gegen den historischen Fehlausloeser
#      "AutoPR-11" (bare "11kg" im Text = Flaschengroesse, keine
#      Gewichtsangabe) sitzt jetzt in GEWICHT_RE selbst (Komma-Nachkommastelle
#      im Text zwingend erforderlich) statt in der Artikelnr-Pruefung - siehe
#      Kommentare bei GEWICHT_RE/artikel_gewicht fuer Details und Beispiele.
# 2026-08-21h: Sammeldruck-Deckblatt zeigt jetzt die ECHTE Packmenge statt
#   hartcodiert "1 Stück". Hintergrund (auf Wunsch behoben): Beim Sammeldruck
#   sieht der Packer haeufig NUR das gedruckte Deckblatt, nicht die Pickliste
#   selbst - bei einem Gewichts-/Packungs-/Set-Artikel (z.B. "3 Bestellungen
#   à 1 Stück" bei einem 3-Fach-Artikel) war das bisher schlicht falsch.
#   finde_sammelgruppen() berechnet die Menge je Bestellung jetzt ueber
#   effektive_menge() (neues Feld "menge_je_text", z.B. "3 Stück"), die
#   Bezeichnung hat das Mengen-Token dabei ebenfalls durch "GESAMTMENGE"
#   ersetzt. Neue CSV-Spalte "MengeJeBestellung" (ans Ende angehaengt, siehe
#   schreibe_sammel_csv/_lade_bestehende_sammelgruppen) transportiert das nach
#   scan_druck.py, das es sowohl auf dem live gedruckten Deckblatt
#   (deckblatt_seite()) als auch in der Sammeldruck-Vorschau auf der
#   Pickliste selbst (gruppe_block()) anzeigt.
# 2026-08-21g: Bei erneuter projektweiter Pruefung gefundener, von den
#   heutigen Aenderungen UNABHAENGIGER Bug in extrahiere_adresse() behoben:
#   Bei einer Lieferadresse mit nur 2 Zeilen (Name + "PLZ Ort", keine eigene
#   Strassenzeile) griff der alte Fallback "deliv[1] if len(deliv) > 1 else
#   ''" faelschlich, weil deliv[1] bei genau 2 Elementen identisch mit der
#   PLZ/Ort-Zeile selbst ist - strasse wurde faelschlich mit dem PLZ/Ort-Text
#   belegt statt leer zu bleiben (Hausnummer blieb dadurch leer, Deutsche-
#   Post-Zuordnung per PLZ+Hausnummer konnte fuer solche Bestellungen
#   fehlschlagen). Fallback jetzt korrekt "" statt einer falschen Zeile.
# 2026-08-21f: Set-Artikel-Feature (2026-08-21c) gegen zwei echte Rechnungen
#   getestet (1701528/52107-4 "2-Fach-Artikel", 1701529/MB15-GD12-10
#   "10-Fach-Artikel"). Menge-Berechnung war in beiden Faellen korrekt, aber
#   die GESAMTMENGE-Textersetzung (_fach_token_re) suchte bisher NUR nach dem
#   '<N>x'-Muster aus dem urspruenglichen Beispiel ('2x Flaschenkappe') - bei
#   1701529 steht die Set-Groesse aber als eigene Zeile '10 Stück' in der
#   Bezeichnung (wie bei den bereits laenger bestehenden Packungsartikeln),
#   wurde also nicht gefunden und blieb sichtbar stehen. Jetzt zusaetzlich
#   das '<N> Stück'-Muster erkannt (dieselbe Regex-Basis wie bei
#   Packungsartikeln/_pack_token_re). Bei 1701528 bleibt die Bezeichnung
#   bewusst unveraendert ('4x' beschreibt dort den Karton-Inhalt der
#   Basis-Artikelnummer 52107/4, nicht die unabhaengige '2-Fach-Artikel'-
#   Mehrfachverkaufs-Markierung dieser Bestellung - keine Zahl in der
#   Bezeichnung entspricht der Fach-Groesse, die Ersetzung findet also
#   zurecht nichts) - die Menge-Spalte zeigt trotzdem korrekt "2 Stück".
# 2026-08-21e: EAN-Verifikationsscan (ean_zuordnung.csv, schreibe_ean_csv) fuer
#   Set-Artikel (N-Fach-Artikel) auf Wunsch ergaenzt: verlangt jetzt N x Menge
#   Scans statt nur der rohen Bestellmenge - seit 2026-08-21c steht am Etikett
#   "GESAMTMENGE" statt "<N>x", es ist also am Label nicht mehr erkennbar, ob
#   es sich um ein Set oder Einzelbestellungen handelt; der Verifikationsscan
#   muss deshalb selbst die richtige (hoehere) Anzahl einfordern. Gewichts-
#   artikel bleiben hier bewusst unveraendert (laufen laut vorheriger
#   Rueckfrage im Chat ueber mengen_zuordnung.csv, nicht ueber EAN).
# 2026-08-21d: Die beiden in 2026-08-21c offen gelassenen Einschraenkungen auf
#   Wunsch nachgezogen:
#   1) Sammelliste (Teil A) zeigt jetzt fuer Gewichts-, Packungs- UND Set-
#      Artikel dieselbe tatsaechliche Stueckzahl/Einheit wie die Packuebersicht
#      (Teil B), statt der rohen Bestellmenge. Neue gemeinsame Funktion
#      effektive_menge() buendelt die Multiplikator-Logik (bisher nur inline in
#      pack_anzeige()); die Aggregation in der Sammelliste nutzt sie jetzt
#      ebenso wie pack_anzeige(). Hervorhebung (rot/fett) folgt konsistent mit.
#   2) order_scananzahl() (-> mengen_zuordnung.csv, Mehrfach-Scan-Sicherung in
#      scan_druck.py) zaehlt jetzt auch bei Set-Artikeln nach der tatsaechlichen
#      Stueckzahl (Set-Groesse x Menge), nicht mehr nur bei Gewichtsartikeln.
#      Packungsartikel zaehlen dort bewusst WEITERHIN nach der rohen
#      Bestellmenge (unveraendert, nicht Teil dieser Anfrage).
# 2026-08-21c: Set-Artikel ("N-Fach-Artikel") auf der Packuebersicht wie
#   Gewichts-/Packungsartikel behandelt. Neue explizite Markierung im PDF
#   ("<N>-Fach-Artikel" als eigene Zeile im selben Bereich wie "Lagerort:",
#   z.B. "2-Fach-Artikel" bei einem Artikel "2x Flaschenkappe") -> N wird in
#   parse_block() erfasst (neues Feld "fach"), pack_anzeige() ersetzt das
#   "<N>x"-Mengen-Praefix in der Bezeichnung durch "GESAMTMENGE" und zeigt in
#   der Menge-Spalte die tatsaechlich zu verpackende Stueckzahl (N x bestellte
#   Menge), analog zu den bestehenden Gewichts-/Packungsartikeln. Bewusst NUR
#   auf der Packuebersicht (Teil B) - die aggregierte Sammelliste (Teil A)
#   zeigt bei Gewichts-/Packungsartikeln schon bisher die rohe Bestellmenge
#   ohne Multiplikator (bekannte, bisher nicht angefragte Einschraenkung) und
#   bleibt fuer Set-Artikel aus Konsistenzgruenden gleich behandelt.
# 2026-08-21b: NACHBESSERUNG zu 2026-08-21a - nach dem Deployment fielen zwei
#   Regressionen auf (Rechnungslauf 21.08., u.a. fast alle Deutsche-Post-
#   Briefmarken nicht zuordenbar + "Prüfung NICHT bestanden" auf fast jeder
#   Packuebersicht-Seite):
#   1) Die Zeilenanfang-Verankerung der PLZ-Suche (Punkt 2 unten) brach bei
#      Lieferadressen mit explizitem Laenderpraefix VOR der PLZ, z.B.
#      "DE-97262 Hausen b. Würzburg" (Rechnung 1700861, Michael Caesar) -
#      genau wie "AT-8020" beim oesterreichischen Pendant, nur dass das bisher
#      niemand fuer Deutschland selbst erwartet hatte. Die Zeile beginnt dann
#      mit Buchstaben statt einer Ziffer und wurde nicht mehr gefunden -> PLZ
#      blieb leer, Post-Zuordnung per PLZ+Hausnummer schlug reihenweise fehl.
#      Jetzt zusaetzlich ein optionales kurzes Grossbuchstaben-Praefix (+
#      Trenner) vor der PLZ zugelassen; der urspruengliche Fall 1700723
#      ("GmbHPO126-0287 VDS Getriebe") bleibt weiterhin ausgeschlossen, da
#      "Gmb..." nach dem ersten Grossbuchstaben auf Kleinbuchstaben trifft.
#   2) Die neue Markierung "Artikelnr/Bezeichnung nicht erkannt" (Punkt 3
#      unten) wieder entfernt: "weder Artikelnr noch Bezeichnung, nur ein
#      Betrag" ist in diesem Amicron-Layout die NORMALE Form der Versand-
#      kosten-Zeile (oft ganz ohne das Wort "Versandkosten" im PDF-Text) und
#      steht auf so gut wie jeder Rechnung - die Warnung schlug dadurch
#      praktisch immer an und machte die Pruefung durch die Flut an
#      Falschmeldungen wertlos, statt echte Faelle sichtbar zu machen.
# 2026-08-21a: Projektweite Fehlerpruefung, drei Punkte behoben:
#   1) _land_aus_zeilen() verlangte bei "AT-1234" zwingend den Bindestrich,
#      waehrend _plz() ihn schon immer optional behandelte - "AT 1234"
#      (Leerzeichen statt Bindestrich) ohne das Wort "ÖSTERREICH" im Text
#      wurde dadurch weiterhin faelschlich als "DE" erkannt (derselbe
#      Ausfall wie bei Rechnung 1699130). Bindestrich bei "AT" jetzt auch
#      hier optional; bei der einbuchstabigen Form "A-" bleibt er Pflicht,
#      um ein einzelnes "A" nicht faelschlich als Oesterreich-Kuerzel zu werten.
#   2) Die PLZ-Zeilensuche (AT- wie DE-Zweig) verlangt jetzt zusaetzlich, dass
#      die Ziffernfolge am ZEILENANFANG steht (vorher: irgendwo in der Zeile).
#      Ohne diese Verankerung haette z.B. eine Telefonzeile im Adressfenster
#      ("Tel. 0664 1234567") die PLZ-Suche erneut in die Irre fuehren koennen -
#      strukturell derselbe Fehlertyp wie der Fix vom 2026-08-20a, nur von der
#      anderen Seite.
#   3) Eine Position mit weder erkannter Artikelnr NOCH Bezeichnung, aber
#      vorhandenem Betrag, wurde bisher lautlos als "Versand" durchgewunken -
#      die Vollstaendigkeitspruefung verlangte dafuer nur den (ja vorhandenen)
#      Betrag, und der Betrag stimmte auch in der Summenkontrolle. Ein echter
#      Artikel, dessen Artikelnr/Bezeichnung im PDF nicht extrahiert werden
#      konnte, fiel dadurch OHNE JEDE WARNUNG aus Pickliste, EAN-Liste und
#      Sammeldruck heraus. Dieser Fall wird jetzt IMMER als "Artikelnr/
#      Bezeichnung nicht erkannt" markiert und erscheint als sichtbare
#      "Pruefung NICHT bestanden"-Warnung auf der gedruckten Packuebersicht.
# 2026-08-20a: Falsche PLZ bei Amazon-Business-Lieferadressen mit angehaengter
#   PO-Referenz behoben (Rechnung 1700723, "VDS Getriebe"/Wolfern: Namenszeile
#   "GmbHPO126-0287 VDS Getriebe" verschmilzt Firmenname und Bestellreferenz
#   ohne Trennzeichen). Der Oesterreich-Zweig in extrahiere_adresse() suchte die
#   PLZ-Zeile bisher VORWAERTS durch die Lieferadresse; die Bestellreferenz
#   "0287" erfuellte dabei zufaellig das Muster "4 Ziffern + Leerzeichen + Text"
#   und wurde als PLZ genommen (post_zuordnung.csv bekam "0287" statt der
#   echten PLZ "4493" -> Deutsche-Post-Zuordnung per PLZ+Hausnummer scheiterte).
#   Jetzt wird wie im DE-Zweig direkt darunter RUECKWAERTS gesucht (PLZ+Ort
#   steht so gut wie immer als letzte inhaltliche Zeile), das findet nun
#   zuverlaessig die echte PLZ-Zeile statt einer zufaellig passenden Ziffernfolge
#   weiter oben im Adressblock.
# 2026-08-11a: Fehlende PLZ bei abweichender Lieferadresse behoben (Rechnung
#   1699130, Amazon "Jan Glöer": Rechnungsadresse AT-1020 Wien, Lieferadresse
#   24404 Maasholm/Deutschland). extrahiere_adresse() bestimmte das Land bisher
#   IMMER aus Liefer- UND Rechnungsadresse zusammen (fuer den Fall 1695459, bei
#   dem nur die Rechnungsadresse das "AT-"-Kuerzel trug). Das riss hier den
#   umgekehrten Fall mit: eine oesterreichische Rechnungsadresse liess eine
#   voellig normale DEUTSCHE Lieferadresse faelschlich als "AT" gelten -> es
#   wurde nur noch nach einer 4-stelligen PLZ gesucht, die echte 5-stellige PLZ
#   passte auf kein Muster mehr und blieb leer (keine post_zuordnung.csv-
#   Zuordnung, keine PLZ auf der Sammeldruck-Pickliste). Jetzt wird das Land
#   zuerst NUR aus der Lieferadresse bestimmt; nur wenn die Lieferadresse WEDER
#   ein AT-Signal NOCH eine gueltige 5-stellige PLZ enthaelt, wird zusaetzlich
#   die Rechnungsadresse herangezogen (deckt weiterhin Fall 1695459 ab).
# 2026-07-20b: Erstellungsdatum + Uhrzeit werden als Fuss auf jeder Seite der
#   Pickliste abgedruckt. Ausserdem Fehltrigger der Gewichtsartikel-Erkennung
#   behoben: Artikel wie 'AutoPR-11' (Modellnummer 11 == '11kg' im Text = fuer
#   11-kg-Flaschen) galten faelschlich als 11-kg-Gewichtsartikel; jetzt zaehlt nur
#   noch ein Komma-Dezimal (0,5 / 1,0 / 3,0) als Gewichtsangabe.
# 2026-07-20a: Packungsartikel "<Zahl>/<N>" (z.B. 2248/10) mit "<N> Stück" im Text
#   werden in der Packuebersicht wie Gewichtsartikel behandelt: "<N> Stück" -> in
#   der Bezeichnung "GESAMTMENGE", und die Mengenspalte zeigt die tatsaechliche
#   Gesamtstueckzahl (Packungsgroesse x bestellte Menge, z.B. 2x -> 20 Stück).

# Reihenfolge der Kategorien in der Sammelliste (Laufweg im Lager).
KATEGORIE_REIHENFOLGE = [
    "Topseller", "Eigenfertigung", "Schlauchlager", "Palettenlager",
    "Poolchemie", "Kleinteile", "Regallager", "DPD / Warenpost", "Allgemein",
]

# Kleine Nachlaeufe kurz vor der Abhol-Deadline (meist nur ein paar Bestellungen)
# lohnen sich nicht auf mehrere Lagerort-Listen aufgeteilt - dann kommt ALLES auf
# EINE gemeinsame Liste (weiterhin sortiert nach Artikelnummer). Ab dieser Anzahl
# Bestellungen im Lauf wird wieder wie gewohnt nach Lagerort/Kategorie getrennt.
KOMBINIERT_SCHWELLE = 20
KOMBINIERT_LABEL = "Alle Bestellungen"

# Poolchemie wird ueber den Lagerort "Poolchemie" erkannt. Falls einzelne
# Artikelnummern fest zu Poolchemie gehoeren (ohne dass der Lagerort es sagt),
# hier eintragen:  POOLCHEMIE_ARTIKEL = {"12345", "67890"}
POOLCHEMIE_ARTIKEL = set()

# Sammeldruck: ab wie vielen gleichen Einzel-Artikel-Bestellungen (je 1 Stueck,
# nur ein Artikel in der Bestellung) ein gemeinsamer Sammel-Barcode erzeugt wird.
SAMMEL_MIN = 2

# Spaltenbaender nach x-Position.
# ---------------------------------------------------------------------------
# AMICRON-LAYOUT MIT EAN-SPALTE (Stand 26.06.2026):
# Amicron druckt seit der Layout-Umstellung zwischen Einheit und Menge eine
# zusaetzliche EAN-Spalte. Nur Artikelnr (ART_MAX), Bezeichnung (BEZ_LO/HI)
# und Einheit (EINH_LO/HI) werden noch ueber feste x-Baender gelesen - Menge,
# Einzelpreis und Gesamtpreis werden seit der Umstellung auf _preiszahlen()
# NICHT mehr ueber eigene feste Baender pro Spalte erkannt, sondern
# layout-unabhaengig als die numerischen Tokens rechts von NUM_MIN_X
# (s.u., in der Reihenfolge Menge/EP/GP von links nach rechts, siehe
# _preiszahlen()/parse_block()). Ein fruehes MENGE_LO/EP_LO/GP_LO-Bandsystem
# (mit den festen Werten 380..440/440..487/487..545) wurde dadurch
# ueberfluessig und ist entfernt (2026-08-21j-Aufraeumung) - diese drei
# Spalten NICHT mehr hier nachjustieren, sondern bei einer neuen Layout-
# Variante an NUM_MIN_X (und ggf. EINH_LO/EINH_HI fuer die EAN-Erkennung).
# Die Selbstpruefung Menge*EP==GP und Summe==Rechnungsbetrag faengt grobe
# Fehlkalibrierungen ab.
ART_MAX = 112
BEZ_LO, BEZ_HI = 145, 281
EINH_LO, EINH_HI = 281, 305     # Einheit (EAN liegt rechts davon, eigenes Band)
TABLE_TOP, TABLE_BOT = 26, 388   # Tabellenbereich (unter Kopf, ueber Rechnungsbetrag)

# Eine EAN/GTIN ist eine reine Ziffernfolge mit 8 oder 12-14 Stellen. Das Muster
# unterscheidet sie zuverlaessig von Mengen/Betraegen (die ein Komma fuehren) und
# von Artikelnummern (die i.d.R. Buchstaben enthalten und ganz links stehen).
EAN_RE = re.compile(r"^\d{8}$|^\d{12,14}$")

CENT = 0.005  # Toleranz fuer Betragsabgleich


# ----------------------------------------------------------------------------
# HILFSFUNKTIONEN
# ----------------------------------------------------------------------------

def num(s):
    """Deutsche Zahl '1.234,56' -> float; None wenn keine Zahl."""
    s = (s or "").strip()
    if not re.search(r"\d", s):
        return None
    s = s.replace(".", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return None


def _hausnr(street):
    """Wie in scan_druck.py: letzte Zahl (+optionaler Buchstabe) am Ende.
    Satzzeichen dahinter (Komma/Punkt/Semikolon, z.B. 'Clara Zetkin Strasse
    358,' aus einer real beobachteten Lieferadresse) werden ignoriert - vorher
    blieb die Hausnummer dann leer und das Matching lief nur ueber den Namen."""
    m = re.search(r"(\d+\s*[a-zA-Z]?)[\s,.;]*$", (street or "").strip())
    return re.sub(r"\s+", "", m.group(1)) if m else ""


def _plz(s, land="DE"):
    """PLZ aus einem String. DE = 5-stellig, AT = 4-stellig ('AT-8020', 'A-8020'
    oder '8020 Ort'). So passt die in die post_zuordnung.csv geschriebene PLZ zu
    der, die scan_druck.py vom (auslaendischen) Label liest."""
    if land == "AT":
        m = (re.search(r"\b(?:AT|A)-?\s?(\d{4})\b", s or "")
             or re.search(r"\b(\d{4})\b", s or ""))
        return m.group(1) if m else ""
    m = re.search(r"\b(\d{5})\b", s or "")
    return m.group(1) if m else ""


def _land_aus_zeilen(zeilen):
    """'AT' wenn eine Adresszeile auf Oesterreich deutet, sonst 'DE'.
    Erkannt werden 'ÖSTERREICH'/'AUSTRIA' sowie das ISO-Praefix vor der PLZ
    ('AT-8020', 'AT 8020', 'A-8020'). Bindestrich bei 'AT' OPTIONAL (anders
    als frueher zwingend) - passte sonst nicht zu _plz(), die 'AT-' und
    'AT ' gleich behandelt: eine Lieferadresse mit Leerzeichen statt
    Bindestrich ('AT 1020 Wien') und ohne das Wort 'ÖSTERREICH' wurde hier
    faelschlich als 'DE' erkannt (derselbe Ausfall wie bei Rechnung
    1699130). Bei der einbuchstabigen Form 'A-8020' bleibt der Bindestrich
    Pflicht, um ein einzelnes 'A' (z.B. Gebaeudebezeichnung) nicht faelschlich
    als Oesterreich-Kuerzel zu werten."""
    txt = " ".join(zeilen or []).upper().replace("Ö", "OE")
    if re.search(r"OESTERREICH|AUSTRIA|\bAT-?\s?\d{4}\b|\bA-\s?\d{4}\b", txt):
        return "AT"
    return "DE"


def in_band(w, lo, hi):
    """Spaltenzuordnung ueber den MITTELPUNKT des Wortes (nicht die linke Kante).
    Betraege sind rechtsbuendig: ein breiter Wert wie '194,97' beginnt weit links
    (x0 ggf. unter der Bandgrenze) und wuerde per linker Kante ins falsche Band
    (Einzelpreis statt G-Preis) fallen. Der Mittelpunkt liegt stabil in der
    richtigen Spalte."""
    mid = (w["x0"] + w["x1"]) / 2.0
    return lo <= mid < hi


def cluster_lines(words, tol=5):
    """Woerter zu Zeilen gruppieren (top-Toleranz)."""
    ws = sorted(words, key=lambda w: (w["top"], w["x0"]))
    lines, cur, cy = [], [], None
    for w in ws:
        if cy is None or abs(w["top"] - cy) <= tol:
            cur.append(w)
            cy = w["top"] if cy is None else (cy + w["top"]) / 2
        else:
            lines.append(cur)
            cur = [w]
            cy = w["top"]
    if cur:
        lines.append(cur)
    return lines


def line_text(line):
    return " ".join(w["text"] for w in sorted(line, key=lambda w: w["x0"]))


def _lagerort_wert(txt):
    """Vollstaendigen Lagerort-Wert aus einer 'Lagerort:'-Zeile ziehen.
    Nimmt ALLES nach dem letzten ':' und normalisiert Mehrfach-Leerzeichen,
    damit mehrteilige Angaben ('R8 E2', 'A8 + A9') vollstaendig erhalten
    bleiben (frueher wurde nur das erste Token genommen -> 'R8 E2' landete auf
    'Allgemein' statt Kleinteile, und die Fachanzeige war unvollstaendig).
    Leere Zeile ('Lagerort:' ohne Wert) -> ''."""
    after = txt.rsplit(":", 1)[-1]
    return " ".join(after.split())


# "<N>-Fach-Artikel" (z.B. "2-Fach-Artikel", "5-Fach-Artikel"): steht wie die
# Lagerort-Zeile als eigene Zeile im Positionsblock und markiert einen im Set
# verkauften Artikel - N = Anzahl der im Set/in der Verpackungseinheit
# enthaltenen Einzelteile. Bewusst eine EXPLIZITE Markierung statt einer
# Ableitung aus der Bezeichnung (frueher auch bei Gewichts-/Packungsartikeln
# ueblich, dort inzwischen abgeschafft, siehe MULTIPACK_RE): die
# Bezeichnung selbst enthaelt haeufig ein "<N>x "-Praefix (z.B. "2x
# Flaschenkappe", "12x campcooga Schraubkartusche"), das aber NICHT immer
# bedeutet, dass N Einzelteile zu verpacken sind (z.B. ein Karton mit 12
# Kartuschen wird ggf. als EIN Stueck gepackt) - ohne expliziten Marker waere
# das nicht zuverlaessig unterscheidbar.
FACH_ARTIKEL_RE = re.compile(r"(\d+)-Fach-Artikel", re.IGNORECASE)

# "<X>,<Y>kg-Multipack" (z.B. "5kg-Multipack", "0,25kg-Multipack"): steht wie
# die Lagerort-/Fach-Artikel-Zeile als EIGENE Zeile im Positionsblock und
# markiert einen nach Gewicht verkauften Artikel (Gewichtsartikel) - X,Y = das
# Stueckgewicht in kg. Analog zu FACH_ARTIKEL_RE eine EXPLIZITE Markierung
# statt einer Ableitung aus der Bezeichnung (siehe dortiger Kommentar zu den
# historischen Fehlausloesern "AutoPR-11"/"Chlorgranulat 5kg").
MULTIPACK_RE = re.compile(r"(\d+(?:,\d+)?)\s*kg-Multipack", re.IGNORECASE)

# Versand-Kennung aus dem Artikelstamm (Amicron): steht als EIGENE Zeile im
# Positionsblock (wie "Lagerort:"/"N-Fach-Artikel", linksbuendig) - Pax1 = DHL,
# Pox1 = Deutsche Post Grossbrief, Brx1 = Deutsche Post Brief, Wapo = DPD.
# Gross-/Kleinschreibung ist egal. Wird von parse_block() als Feld "kennungen"
# gemerkt und NICHT in die Bezeichnung uebernommen; ausgewertet wird sie vom
# Carrier-Dashboard (carrier_dashboard.py / carrier_regeln.py), die Pickliste
# selbst nutzt sie nicht. ("Wapo" steht teils zusaetzlich als Lagerort - das
# wertet carrier_regeln.py ueber die Lagerort-Werte aus.)
KENNUNG_RE = re.compile(r"^\s*(pax1|pox1|brx1|wapo)\s*$", re.IGNORECASE)

# "<N>-je-Paket" (z.B. "1-je-Paket", "3-je-Paket"): wie Lagerort/Fach-Artikel/
# Kennung eine EIGENE, linksbuendige Zeile im Positionsblock - markiert einen
# Artikel, der aus Gewichts-/Groessengruenden NICHT beliebig in einem
# DHL-Paket gebuendelt werden darf. N = maximale (effektive, also inkl.
# Fach-Artikel-Faktor) Stueckzahl je Paket - "1-je-Paket" fuer Artikel, die
# IMMER einzeln verschickt werden muessen (z.B. ein schwerer 2-Fach-Artikel-
# Doppelpack -> 2 Pakete), "3-je-Paket" fuer Artikel, die zu bis zu drei
# Kartons gebuendelt werden koennen. Wird NICHT in die Bezeichnung
# uebernommen; ausgewertet vom Carrier-Dashboard (carrier_regeln.
# je_paket_aufteilung()), die Pickliste selbst nutzt sie nicht - mit
# Matthias abgestimmt 2026-09-23.
JE_PAKET_RE = re.compile(r"^\s*(\d+)-je-Paket[\s.]*$", re.IGNORECASE)   # ganze Zeile, wie KENNUNG_RE


# ----------------------------------------------------------------------------
# KATEGORIE-ZUORDNUNG (allein aus dem Lagerort)
# ----------------------------------------------------------------------------

GRID_RE = re.compile(r"\b([A-K])\s?([1-8])\b")


def classify(lagerorte, art):
    """Lagerkategorie aus den Lagerort-Strings (Reihenfolge = Prioritaet)."""
    text = " ".join(lagerorte)
    low = text.lower()
    if art and art in POOLCHEMIE_ARTIKEL:
        return "Poolchemie"
    if not lagerorte:
        return "Allgemein"
    if "wapo" in low or "warenpost" in low:
        return "DPD / Warenpost"
    if "topseller" in low:
        return "Topseller"
    if "fertigung" in low:
        return "Eigenfertigung"
    if "schläuche" in low or "schlauch" in low or "schlaeuche" in low:
        return "Schlauchlager"
    if "schwere artikel" in low or "palette" in low:
        return "Palettenlager"
    if "poolchemie" in low:
        return "Poolchemie"
    if "kleinteile" in low:                 # Lagerort woertlich "Kleinteile"
        return "Kleinteile"
    mreg = re.search(r"regal\s*0*(\d+)", low)
    if mreg:
        nr = int(mreg.group(1))
        return "Kleinteile" if 21 <= nr <= 30 else "Regallager"
    if GRID_RE.search(text.upper()):
        return "Kleinteile"
    return "Allgemein"


def ist_versand(art, bez):
    """Versandkosten / nicht-koerperliche Position (nicht ins Lager)."""
    if "versandkosten" in (bez or "").lower():
        return True
    if not art and not (bez or "").strip():
        return True
    return False


def auftrag_kategorie(r):
    """Haupt-Lagerort einer ganzen Bestellung = die Kategorie mit der HOECHSTEN
    Prioritaet (= weitestVORNE in KATEGORIE_REIHENFOLGE) ueber alle echten
    Positionen. So landet jede Bestellung in genau EINER Kommissionierliste."""
    kats = set()
    for p in r["positionen"]:
        if ist_versand(p["art"], p["bez"]):
            continue
        kats.add(classify(p["lagerorte"], p["art"]))
    for k in KATEGORIE_REIHENFOLGE:
        if k in kats:
            return k
    return "Allgemein"


def finde_sammelgruppen(rechnungen, min_anzahl=SAMMEL_MIN):
    """Bestellungen, die genau EINEN Artikel mit Menge 1 enthalten (Versand-
    zeilen zaehlen nicht), nach Artikelnummer gruppieren. Gruppen ab
    `min_anzahl` Bestellungen bekommen einen Sammel-Barcode 'SAM-n'.
    Rueckgabe: Liste von dicts {code, art, bez, einh, menge_je_text, rnr:[...]}.

    "bez" und "menge_je_text" kommen ueber effektive_menge() - bei einem
    Gewichts-/Set-Artikel (2026-08-21h) steht hier NICHT mehr
    hartcodiert "1 Stück", sondern die tatsaechlich je Bestellung zu
    verpackende Menge (z.B. "3 Stück" bei einem 3-Fach-Artikel), und die
    Bezeichnung hat das Mengen-Token durch "GESAMTMENGE" ersetzt. Der Packer
    sieht bei einem Sammeldruck haeufig NUR das gedruckte Deckblatt (nicht die
    Pickliste selbst) - das muss deshalb exakt und eindeutig die reale
    Packmenge zeigen, siehe drucke_sammel()/deckblatt_seite() in scan_druck.py.

    ACHTUNG: die hier vergebenen Codes sind nur VORLAEUFIG (reine
    Aufzaehlung "SAM-1, SAM-2, ..." innerhalb DIESES Laufs) - main() ruft
    danach IMMER merge_sammelgruppen() auf, um Kollisionen mit noch nicht
    abgearbeiteten Codes aus einem frueheren Lauf zu vermeiden, bevor die
    Codes gedruckt/geschrieben werden."""
    nach_artikel = {}
    for r in rechnungen:
        echte = [p for p in r["positionen"] if not ist_versand(p["art"], p["bez"])]
        if len(echte) != 1:
            continue
        p = echte[0]
        if p["menge"] is None or abs(p["menge"] - 1.0) > 1e-9:
            continue
        if not p["art"]:
            continue
        eff_menge, eff_einh, eff_bez, _ = effektive_menge(p)
        g = nach_artikel.setdefault(
            p["art"], {"art": p["art"], "bez": eff_bez, "einh": p["einh"],
                       "menge_je_text": mengentext(eff_menge, eff_einh),
                       "ean": "", "rnr": []})
        if not g["ean"] and p.get("ean"):
            g["ean"] = p["ean"]
        if r["rnr"]:
            g["rnr"].append(r["rnr"])
    gruppen = [g for g in nach_artikel.values() if len(g["rnr"]) >= min_anzahl]
    gruppen.sort(key=lambda g: (-len(g["rnr"]), g["art"]))
    for i, g in enumerate(gruppen, 1):
        g["code"] = f"SAM-{i}"
    return gruppen


def _lade_bestehende_sammelgruppen(csv_pfad):
    """sammel_zuordnung.csv -> Liste noch 'frischer' Gruppen aus einem
    frueheren Lauf (aeltere als BRUECKEN_CSV_MERGE_TAGE werden verworfen,
    siehe dort). Format je Zeile:
        Code;Artikelnr;Bezeichnung;Menge;Rechnungsnummern;EAN;Rechnungsdatum;
        MengeJeBestellung (letzte Spalte seit 2026-08-21h, siehe
        finde_sammelgruppen/schreibe_sammel_csv)."""
    gruppen = []
    for t in _lies_bestehende_zeilen(csv_pfad, 5):
        datum = t[6] if len(t) > 6 else ""
        if not _bruecken_datum_frisch(datum):
            continue
        rnr = [x.strip() for x in t[4].split(",") if x.strip()] if len(t) > 4 else []
        if not rnr:
            continue
        gruppen.append({
            "code": t[0], "art": t[1], "bez": t[2],
            "ean": t[5] if len(t) > 5 else "", "rnr": rnr, "datum": datum,
            # Aeltere Zeilen (vor 2026-08-21h) haben diese Spalte noch nicht ->
            # Fallback auf die bis dahin einzig moegliche Annahme "1 Stück".
            "menge_je_text": t[7] if len(t) > 7 and t[7] else "1 Stück",
        })
    return gruppen


def merge_sammelgruppen(neue_gruppen, csv_pfad):
    """Fuehrt frisch berechnete Sammeldruck-Gruppen (vorlaeufige Codes aus
    finde_sammelgruppen, reine Aufzaehlung INNERHALB dieses Laufs) mit noch
    frischen Gruppen aus einem FRUEHEREN Lauf zusammen - mit der Garantie,
    dass ein bereits vergebener Code NIE auf eine andere Bestellgruppe
    umgebogen wird.

    Hintergrund: Der Sammelcode wird auf Papier gedruckt (Packuebersicht-
    Deckblatt mit den Original-Kundennamen). Scannt jemand einen noch nicht
    abgearbeiteten Code aus einem frueheren Lauf, MUSS er weiterhin exakt
    dieselben Bestellungen ergeben wie auf dem Papier steht - sonst wuerden
    am Ende Label fuer die falschen Bestellungen gedruckt, ohne dass das beim
    Scannen auffaellt (es wird ja keine Adresse manuell gegengeprueft).
    Da archivierte Rechnungen in einem spaeteren Lauf nie wieder auftauchen,
    ueberschneiden sich alte und neue Gruppen nie in ihren Rechnungsnummern -
    das eigentliche Risiko ist rein die Wiederverwendung derselben Codenummer
    fuer eine andere Gruppe, und genau das verhindert diese Funktion.

    Rueckgabe: (fuer_pdf, fuer_csv)
      fuer_pdf: NUR die Gruppen DIESES Laufs, mit finalen kollisionsfreien
                Codes - das geht in die neu gedruckte Pickliste.
      fuer_csv: fuer_pdf PLUS alle noch frischen Alt-Gruppen unveraendert -
                das wird komplett nach sammel_zuordnung.csv geschrieben."""
    alt = _lade_bestehende_sammelgruppen(csv_pfad)
    belegt = set()
    for g in alt:
        m = re.match(r"SAM-(\d+)$", g["code"])
        if m:
            belegt.add(int(m.group(1)))

    naechste = 1
    fuer_pdf = []
    for g in neue_gruppen:
        while naechste in belegt:
            naechste += 1
        neu = dict(g)
        neu["code"] = f"SAM-{naechste}"
        # finde_sammelgruppen() liefert noch KEIN "datum" (eine Sammelgruppe
        # hat kein einzelnes Rechnungsdatum, sie fasst mehrere Bestellungen
        # zusammen) - ohne dieses Feld wuerde _bruecken_datum_frisch() die
        # Zeile jetzt (Stand 2026-09-01, leeres Datum gilt nicht mehr als
        # "immer frisch") schon beim naechsten Lauf faelschlich verwerfen,
        # obwohl der Sammelcode noch gar nicht gedruckt wurde. Heutiges Datum
        # als "seit wann bekannt" eintragen, analog zu _csv_datum().
        neu["datum"] = datetime.now().strftime("%d.%m.%Y")
        belegt.add(naechste)
        fuer_pdf.append(neu)

    return fuer_pdf, (alt + fuer_pdf)


# ----------------------------------------------------------------------------
# PARSER
# ----------------------------------------------------------------------------

# Numerische Spalten (Menge/Einzelpreis/G-Preis) werden NICHT mehr ueber feste
# x-Baender gelesen, sondern layout-unabhaengig als die (bis zu drei) Komma-
# Dezimalzahlen RECHTS der Bezeichnung, in Lesereihenfolge Menge | EP | GP.
# Damit funktionieren alle Amicron-Spaltenlagen (mit und ohne EAN-Spalte, sowie
# die verschiedenen alten Vorlagen) und gemischte Rechnungen. Eine EAN ist eine
# reine Ziffernfolge OHNE Komma und wird so nie als Preis gezaehlt. Betraege in
# der Bezeichnung (z.B. "3,2mm") liegen links von NUM_MIN_X und stoeren nicht.
NUM_MIN_X = 305        # ab hier (rechts der Einheit) stehen Menge/EP/GP


def _preiszahlen(line):
    """Zahlen einer Zeile rechts der Bezeichnung (Menge/EP/GP), links -> rechts.
    Die Menge kann eine GANZE Zahl ohne Komma sein (z.B. '2'); Einzelpreis und
    G-Preis tragen immer ein Komma. Reine EAN/GTIN-Ziffernfolgen (8 oder 12-14
    Stellen) werden ausgeschlossen. Zahlen der Bezeichnung (z.B. '3,2mm', '8mm',
    '50 Stück') liefern entweder kein num() oder liegen links von NUM_MIN_X."""
    out = []
    for w in sorted(line, key=lambda w: w["x0"]):
        t = w["text"]
        if w["x0"] < NUM_MIN_X or EAN_RE.match(t) or num(t) is None:
            continue
        out.append(w)
    return out


def parse_block(words):
    """Ein Positions-Block (Woerter zwischen zwei Trenner-Punkten) -> dict.
    Layout-unabhaengig: Menge/EP/GP = die Komma-Zahlen rechts der Bezeichnung
    (rechtsbuendig verankert: GP ganz rechts). Artikelnr (x0<ART_MAX) und
    Bezeichnung (BEZ-Band) werden ueber ALLE Blockzeilen gesammelt, weil die
    Bezeichnung mehrzeilig sein kann und die Zahlen dann in einer anderen Zeile
    stehen als die Artikelnummer."""
    lines = cluster_lines(words)
    # Kopf- und Rechnungsbetrag-Zeile gehoeren nicht zur Position (koennen aber im
    # selben Block liegen, wenn die Tabelle knapp getrennt ist) -> raus.
    lines = [ln for ln in lines
             if not any(k in line_text(ln)
                        for k in ("Artikelnr", "Bezeichnung", "Rechnungsbetrag"))]

    # Hauptzeile = Zeile mit den meisten Preis-Zahlen (Menge/EP/GP).
    main, best = None, 0
    for ln in lines:
        n = len(_preiszahlen(ln))
        if n > best:
            main, best = ln, n

    art = einh = ean = ""
    menge = ep = gp = None
    bez_parts, lagerorte = [], []
    fach = None
    gewicht = None
    kennungen = []
    je_paket = None

    # Menge / Einzelpreis / G-Preis aus der Hauptzeile (rechts verankert).
    if main:
        pz = _preiszahlen(main)
        if len(pz) >= 3:
            menge, ep, gp = num(pz[-3]["text"]), num(pz[-2]["text"]), num(pz[-1]["text"])
        elif len(pz) == 2:
            menge, gp = num(pz[-2]["text"]), num(pz[-1]["text"])
        elif len(pz) == 1:
            gp = num(pz[-1]["text"])
        # Einheit (Mittelpunkt im Einheit-Band; EAN-Ziffernfolgen ausgeschlossen)
        einh_toks = [w["text"] for w in main
                     if in_band(w, EINH_LO, EINH_HI) and not EAN_RE.match(w["text"])]
        einh = einh_toks[0] if einh_toks else ""

    # Artikelnr, Bezeichnung, EAN und Lagerort ueber ALLE Zeilen sammeln.
    # Kopfzeile ('Artikelnr Bezeichnung ...') und die 'Lagerort:'-Zeile werden
    # dabei uebersprungen. Die Artikelnr wird NUR aus der ersten Positionszeile
    # gezogen (Folgezeilen einer umbrochenen Bezeichnung koennen links ebenfalls
    # etwas haben, z.B. '10 Stück' -> darf nicht in die Artnr wandern).
    art_toks = []
    art_done = False
    i = 0
    while i < len(lines):
        ln = lines[i]
        txt = line_text(ln)
        if "Artikelnr" in txt or "Bezeichnung" in txt:      # Tabellenkopf
            i += 1
            continue
        if "Lagerort" in txt:
            # VOLLSTAENDIGEN Lagerort-Wert ziehen (nicht nur das erste Token),
            # damit mehrteilige Angaben wie 'R8 E2' oder 'A8 + A9' erhalten
            # bleiben. Wichtig fuer die Fachanzeige UND die Einsortierung: bei
            # 'R8 E2' bringt erst das 'E2' die Bestellung aufs Kleinteile-Blatt
            # (classify -> GRID_RE); 'R8' allein liefe auf 'Allgemein'.
            wert = _lagerort_wert(txt)
            if not wert and i + 1 < len(lines):
                # 'Lagerort:'-Zeile ohne Wert -> der Wert steht umbrochen in der
                # Folgezeile. Diese nur uebernehmen, wenn sie KEINE Positions-/
                # Preiszeile ist (keine Menge/EP/GP-Zahlen); die Folgezeile wird
                # dann mitverbraucht, damit sie nicht als Bezeichnung/Artnr
                # fehlinterpretiert wird.
                nxt = lines[i + 1]
                if not _preiszahlen(nxt):
                    wert = line_text(nxt).strip()
                    i += 1
            if wert:
                lagerorte.append(wert)
            i += 1
            continue
        m_fach = FACH_ARTIKEL_RE.search(txt)
        if m_fach:
            # "<N>-Fach-Artikel"-Zeile: wie Lagerort NICHT in die Bezeichnung
            # einfliessen lassen, sondern als Set-Groesse merken.
            fach = int(m_fach.group(1))
            i += 1
            continue
        m_multi = MULTIPACK_RE.search(txt)
        if m_multi:
            # "<X>,<Y>kg-Multipack"-Zeile: genau wie Lagerort/Fach-Artikel eine
            # eigene Zeile, NICHT Teil der Bezeichnung (real beobachtet an
            # Rechnung 1703025/WSG2-1,6-5: die Zeile "5kg-Multipack" stand an
            # derselben links-buendigen Position wie "Lagerort:"/"N-Fach-
            # Artikel" - x0 < ART_MAX - und wurde deshalb bisher komplett
            # verworfen, noch bevor sie ueberhaupt bei artikel_gewicht() als
            # Text ankam. Als Set-Groesse merken wie "fach".
            gewicht = float(m_multi.group(1).replace(",", "."))
            i += 1
            continue
        m_ken = KENNUNG_RE.match(txt)
        if m_ken:
            # Versand-Kennung (Pax1/Pox1/Brx1/Wapo) als eigene Zeile: merken,
            # nicht in die Bezeichnung uebernehmen (siehe KENNUNG_RE).
            kennungen.append(m_ken.group(1).lower())
            i += 1
            continue
        m_je_paket = JE_PAKET_RE.match(txt)
        if m_je_paket:
            # "<N>-je-Paket"-Zeile: wie Lagerort/Fach-Artikel/Kennung eine
            # eigene Zeile, NICHT Teil der Bezeichnung - siehe JE_PAKET_RE.
            je_paket = int(m_je_paket.group(1))
            i += 1
            continue
        links = [w for w in sorted(ln, key=lambda w: w["x0"])
                 if w["x0"] < ART_MAX and w["text"] != "."]
        if links and not art_done:                          # erste Positionszeile
            art_toks = [w["text"] for w in links]
            art_done = True
        for w in sorted(ln, key=lambda w: w["x0"]):
            t = w["text"]
            if t == "." or w["x0"] < ART_MAX:
                continue
            if in_band(w, BEZ_LO, BEZ_HI):                  # Bezeichnungs-Spalte
                bez_parts.append((round(w["top"]), w["x0"], t))
            elif EAN_RE.match(t) and w["x0"] >= EINH_LO and not ean:  # reine EAN/GTIN
                ean = t
        i += 1

    art = "".join(art_toks)
    bez = " ".join(t for _, _, t in sorted(bez_parts)).strip()

    return {
        "art": art,
        "bez": bez,
        "einh": einh,
        "ean": ean,
        "menge": menge,
        "ep": ep,
        "gp": gp,
        "lagerorte": lagerorte,
        "fach": fach,
        "gewicht": gewicht,
        "kennungen": kennungen,
        "je_paket": je_paket,
    }


def extrahiere_adresse(words):
    """Lieferadresse -> (name, strasse, hausnr, plz)."""
    send = next((w["top"] for w in words if "Gasecenter" in w["text"]), 640)
    # Untergrenze des Adressfensters. Die kompaktere Rechnungsvorlage schiebt den
    # Lieferadressblock weiter nach unten (PLZ-Zeile jetzt bei top~775); die frueheren
    # 760 schnitten Strasse+PLZ der Lieferadresse ab (-> leere post_zuordnung.csv,
    # Deutsche-Post-Zuordnung per PLZ+Hausnummer scheiterte). 800 deckt den Block ab;
    # unterhalb 776 steht in den Rechnungen kein Inhalt (Seitenhoehe 842).
    ADR_BOT = 800
    left = [w for w in words if w["x0"] < 405 and send + 5 < w["top"] < ADR_BOT]
    right = [w for w in words if w["x0"] >= 405 and send + 5 < w["top"] < ADR_BOT]
    lefts = [line_text(ln) for ln in cluster_lines(left)]
    rights = [line_text(ln) for ln in cluster_lines(right)]
    right_real = [l for l in rights if "Lieferadresse" not in l and "Rechnungsadresse" not in l]
    same = any("Rechnungsadresse" in l for l in rights)
    deliv = lefts if (same or not right_real) else right_real
    deliv = [l for l in deliv if l.strip()]
    if not deliv:
        return "", "", "", ""
    name = deliv[0]
    # Bei abweichender Lieferadresse (Bestellung 1695459, Amazon "Bäck-Kuhn"/
    # "AT-8380 Jennersdorf"): das Länderkürzel "AT-" steht oft NUR in der
    # Rechnungsadresse ("KUHN Bernhard"/"AT-8380"), die separate Lieferadresse
    # nennt nur "8380 Jennersdorf" ohne AT-Praefix und ohne "ÖSTERREICH" ->
    # Land wurde faelschlich als DE erkannt, die 4-stellige PLZ passte dann
    # nicht auf die DE-Regel (5-stellig) und blieb leer (keine Zuordnung,
    # kein Druck). Daher zusaetzlich lefts (Rechnungsadresse) auswerten.
    land = _land_aus_zeilen(deliv)
    if land != "AT" and not any(re.search(r"\b\d{5}\b", l) for l in deliv):
        # Lieferadresse liefert weder ein AT-Signal noch eine gueltige
        # 5-stellige PLZ -> evtl. fehlt nur das AT-Praefix in der separaten
        # Lieferadresse (Fall 1695459) -> zusaetzlich Rechnungsadresse pruefen.
        land = _land_aus_zeilen(deliv + lefts)
    if land == "AT":
        # Oesterreich: 4-stellige PLZ ('AT-8020', 'A-8020' oder '8020 Ort'); die
        # Zeile 'ÖSTERREICH' steht meist darunter und ist nicht die PLZ-Zeile.
        # Rueckwaerts durchsuchen (wie der DE-Zweig unten) statt vorwaerts: bei
        # Amazon-Business-Bestellungen mit angehaengter PO-Referenz in der
        # Namenszeile (z.B. "GmbHPO126-0287 VDS Getriebe") passt "0287" bereits
        # auf das Muster "4 Ziffern + Leerzeichen + Text" und wurde faelschlich
        # als PLZ genommen, bevor die echte PLZ-Zeile ueberhaupt geprueft wurde
        # (Rechnung 1700723, VDS Getriebe/Wolfern: PLZ "0287" statt "4493").
        # Das reine Ziffernmuster ist zusaetzlich auf den ZEILENANFANG verankert
        # (re.match statt re.search): eine PLZ+Ort-Zeile faengt in der Praxis
        # immer mit der PLZ an, waehrend z.B. eine Telefonzeile ("Tel. 0664
        # 1234567") mittendrin eine zufaellig passende 4er-Ziffernfolge haben
        # kann, ohne PLZ+Ort zu sein - das Präfix-Muster (AT-/A- vorneweg)
        # bleibt bewusst ungeankert, da das Kuerzel selbst schon eindeutig ist.
        plzline = next((l for l in reversed(deliv)
                        if re.search(r"\b(?:AT|A)-?\s?\d{4}\b", l)
                        or re.match(r"\s*\d{4}\b\s+\S", l)), "")
    else:
        # 2026-08-21 NACHBESSERUNG: Die Verankerung auf den Zeilenanfang (s.o.)
        # brach bei Lieferadressen mit explizitem Laenderpraefix vor der PLZ
        # (z.B. "DE-97262 Hausen b. Würzburg", Rechnung 1700861/Michael Caesar,
        # analog zu "AT-8020" beim oesterreichischen Pendant) - die Zeile
        # beginnt dort mit Buchstaben, nicht mit der Ziffer, also NICHT mehr
        # gematcht -> PLZ blieb leer, Deutsche-Post-Zuordnung schlug fehl (fast
        # alle Briefmarken einer Sitzung nicht zuordenbar). Optionales, kurzes
        # GROSSBUCHSTABEN-Praefix (+ Trenner) vor der PLZ jetzt zugelassen -
        # "GmbHPO126-0287 VDS Getriebe" (der urspruengliche Fall 1700723) bleibt
        # weiterhin ausgeschlossen, da "Gmb..." nach dem ersten Grossbuchstaben
        # "G" auf Kleinbuchstaben "m" trifft und daher nicht passt.
        plzline = next((l for l in reversed(deliv)
                        if re.match(r"\s*(?:[A-Z]{1,3}-?\s?)?\d{5}\b", l)), "")
    idx = deliv.index(plzline) if plzline in deliv else len(deliv) - 1
    # Bei genau 2 Zeilen (Name + "PLZ Ort", keine eigene Strassenzeile) ist
    # idx=1 und idx-1=0 - der alte Fallback "deliv[1] if len(deliv) > 1 else
    # ''" griff dann FAELSCHLICH, weil deliv[1] bei nur 2 Elementen identisch
    # mit der PLZ/Ort-Zeile (=plzline) selbst ist: strasse wurde faelschlich
    # mit dem PLZ/Ort-Text belegt statt leer zu bleiben. Ohne echte Strassen-
    # zeile (idx-1 zeigt sonst auf die Namenszeile oder liegt vor der Liste)
    # bleibt strasse jetzt korrekt leer.
    street = deliv[idx - 1] if idx - 1 >= 1 else ""
    return name, street, _hausnr(street), _plz(plzline, land)


def lieferadresse_zeilen(words):
    """ALLE Zeilen der Lieferadresse (bzw. der Rechnungsadresse, wenn keine
    abweichende Lieferadresse existiert) in Lesereihenfolge, z.B.
    ['Evi Schmid', 'Jaegerwirth 122', 'DE-94081 Fuerstenzell'].
    Grundlage fuer den CSV-Export des Carrier-Dashboards (Firma/c-o/Zusatzzeilen
    und Land werden dort ausgewertet). Die Zeilenauswahl ist bewusst IDENTISCH zu
    extrahiere_adresse() oben (gleiches Fenster, gleiche Links/Rechts-Regel) -
    dort wird sie fuer die Post-Zuordnung genutzt; hier bewusst als eigene
    Funktion, damit dieser bewaehrte Code unveraendert bleibt."""
    send = next((w["top"] for w in words if "Gasecenter" in w["text"]), 640)
    ADR_BOT = 800
    left = [w for w in words if w["x0"] < 405 and send + 5 < w["top"] < ADR_BOT]
    right = [w for w in words if w["x0"] >= 405 and send + 5 < w["top"] < ADR_BOT]
    lefts = [line_text(ln) for ln in cluster_lines(left)]
    rights = [line_text(ln) for ln in cluster_lines(right)]
    right_real = [l for l in rights
                  if "Lieferadresse" not in l and "Rechnungsadresse" not in l]
    same = any("Rechnungsadresse" in l for l in rights)
    deliv = lefts if (same or not right_real) else right_real
    return [l.strip() for l in deliv if l.strip()]


def parse_pdf(path):
    """Eine Rechnung -> dict mit Kopf, Positionen, Adresse, Abgleich-Status."""
    with pdfplumber.open(path) as pdf:
        page = pdf.pages[0]
        words = page.extract_words(use_text_flow=False)

    tops = sorted({round(w["top"]) for w in words})
    zeilen = [line_text([w for w in words if round(w["top"]) == t]) for t in tops]
    full = "\n".join(zeilen)

    m = re.search(r"Rechnung\s*Nr\.?\s*:?\s*(\d+)", full)
    rnr = m.group(1) if m else ""

    # Direkt unter der "Rechnung Nr."-Zeile druckt Amicron ein Fremdbeleg-Feld:
    #   - Onlineshop  -> WooCommerce-Bestellnummer (reine Ziffern, 4-6 Stellen)
    #   - Amazon      -> "nnn-nnnnnnn-nnnnnnn"
    #   - Zaehltheke  -> Kd-Nr.
    # gefolgt (bei Shop/Marktplatz) von der Kundenmail. Fuer die Pickliste heisst
    # dieses Feld weiterhin "kdnr" (Anzeige im Kopf der Packuebersicht); die
    # Kundenmail wird zusaetzlich als "email" gemerkt - beides zusammen erlaubt
    # schreibe_wc_bestellnummern_csv(), echte Onlineshop-Rechnungen sauber von
    # Amazon/eBay/Theke zu trennen.
    # Bei Kunden mit Mehrfach-Marktplatz-Historie (z.B. schon mal per eBay,
    # jetzt per Amazon bestellt) druckt Amicron MEHRERE Fremdbeleg-Zeilen
    # untereinander (Rechnung 1705582: erst der eBay-Name "*coyote86*", dann
    # die Amazon-Bestellnummer). Ohne Priorisierung haette die erste (hier: der
    # Name) gewonnen - Amazon-/WC-Bestellnummer sind aber die eindeutigeren,
    # fuer den Abgleich nuetzlicheren Kennungen und werden deshalb bevorzugt;
    # ein reiner Marktplatz-Name (eBay etc.) ist nur der Rueckfallwert.
    def _kdnr_prioritaet(cand):
        if re.fullmatch(r"\d{4,6}", cand):                        # WC-Bestellnummer
            return 0
        if re.fullmatch(r"\d{3}-\d{7}-\d{7}", cand):               # Amazon-Bestellnummer
            return 1
        return 2                                                   # sonstiges (z.B. eBay-Name)

    kdnr = ""
    email = ""
    for i, z in enumerate(zeilen):
        if re.search(r"Rechnung\s*Nr", z):
            kandidaten = []
            for j in range(i + 1, min(i + 5, len(zeilen))):
                cand = zeilen[j].strip()
                if not cand:
                    continue
                em = re.search(r"[^\s@;]+@[^\s@;]+\.[A-Za-z]{2,}", cand)
                if em:
                    if not email:
                        email = em.group(0)
                    continue
                if cand != rnr:
                    kandidaten.append(cand)
            if kandidaten:
                kdnr = min(kandidaten, key=_kdnr_prioritaet)
            break

    md = re.search(r"(\d{2}\.\d{2}\.\d{4})", full)
    datum = md.group(1) if md else ""

    # Positionen ueber Trenner-Punkte (einzelnes '.' am linken Rand) zerlegen.
    # Toleranter x-Bereich (20..40), da die Punkte je nach Vorlage bei ~28 oder
    # ~32 sitzen (z.B. 27.99 wurde vom alten Fenster 28..38 knapp verfehlt).
    def _ist_trenner(w):
        return w["text"] == "." and 20 <= w["x0"] <= 40

    # Tabellenende dynamisch bestimmen: Die Amicron-Seite wiederholt UNTER der
    # Positionstabelle einen 'Druckdatum:'-Block (Lieferschein-/Adressteil). Bis
    # kurz davor reicht die Tabelle. Fixe 388 waren zu niedrig -> bei mehreren
    # Artikeln / langen Bezeichnungen fielen Versand oder ganze Positionen weg,
    # weil deren Trennpunkte unterhalb 388 lagen.
    kopf_top = min((w["top"] for w in words if w["text"] == "Artikelnr"),
                   default=TABLE_TOP)
    wiederhol = [w["top"] for w in words
                 if w["text"].startswith("Druckdatum") and w["top"] > kopf_top + 5]
    if wiederhol:
        table_bot = min(wiederhol) - 2
    else:
        rb = [w["top"] for w in words if "Rechnungsbetrag" in w["text"]]
        table_bot = (max(rb) + 40) if rb else TABLE_BOT

    tab = [w for w in words if TABLE_TOP < w["top"] < table_bot]
    seps = sorted(w["top"] for w in tab if _ist_trenner(w))
    positionen = []
    lo = TABLE_TOP
    for s in seps:
        block = [w for w in tab if lo < w["top"] < s and not _ist_trenner(w)]
        if block:
            p = parse_block(block)
            if p["gp"] is not None or p["art"] or p["bez"]:
                positionen.append(p)
        lo = s

    mb = re.search(r"Rechnungsbetrag\s*€?\s*([\d.,]+)", full)
    total = num(mb.group(1)) if mb else None

    name, strasse, hausnr, plz = extrahiere_adresse(words)
    adresse_zeilen = lieferadresse_zeilen(words)

    # Sendungsgewicht (kg) aus der Kopfzeile "Sendungsgewicht : 0,34000" (seit
    # 2026-09 auf den Amicron-Rechnungen, Grundlage der Carrier-Zuordnung im
    # Carrier-Dashboard). None, wenn die Zeile fehlt oder nicht lesbar ist.
    mg = re.search(r"Sendungsgewicht\s*:?\s*([\d.,]+)", full)
    sendungsgewicht = num(mg.group(1)) if mg else None

    # Betragsabgleich
    zeilen_ok = True
    for p in positionen:
        if p["menge"] is not None and p["ep"] is not None and p["gp"] is not None:
            if abs(round(p["menge"] * p["ep"], 2) - p["gp"]) > CENT:
                zeilen_ok = False
    summe = round(sum(p["gp"] for p in positionen if p["gp"] is not None), 2)
    # Ein NICHT lesbarer Rechnungsbetrag (z.B. wenn eine sehr lange Bezeichnung ihn
    # im Quell-PDF ueberlagert) darf keinen falschen "Summe != 0,00"-Fehler ausloesen.
    total_lesbar = total is not None
    summe_ok = (not total_lesbar) or (abs(summe - total) <= CENT)

    # Vollstaendigkeitspruefung: jede ECHTE Position braucht Artikelnr + Menge +
    # Einzelpreis + Betrag(GP). Versandzeilen (ohne Artikel) brauchen Menge + Betrag.
    vollstaendig_ok = True
    unvollstaendig = []
    for p in positionen:
        if ist_versand(p["art"], p["bez"]):
            # Versand braucht nur den Betrag (Menge ist stets 1 und geht nicht in
            # die Summe ein). So warnt eine im Quell-PDF verschmolzene Versand-Menge
            # nicht unnoetig.
            # 2026-08-21 GETESTET UND WIEDER ENTFERNT: Ein zusaetzliches Flaggen
            # von Positionen ohne Artikelnr UND Bezeichnung wurde kurzzeitig
            # eingebaut (Sorge: ein echter Artikel koennte so unbemerkt als
            # Versand durchrutschen). In der Praxis ist "weder Artikelnr noch
            # Bezeichnung, nur ein Betrag" aber genau die NORMALE Form der
            # Versandkosten-Zeile in diesem Amicron-Layout (steht auf so gut wie
            # JEDER Rechnung, oft komplett ohne das Wort "Versandkosten" im PDF-
            # Text) - die Warnung schlug dadurch bei praktisch jeder einzelnen
            # Rechnung an und machte die Pruefung durch die Flut an Falsch-
            # meldungen wertlos. Aus den Daten allein ist der echte Versand-Fall
            # nicht von einem tatsaechlich verlorenen Artikel unterscheidbar -
            # deshalb bewusst wieder auf das reine Betrag-Erfordernis zurueck.
            fehlt = []
            if p["gp"] is None:
                fehlt.append("Betrag")
        else:
            fehlt = []
            if not p["art"]:
                fehlt.append("Artikelnr")
            if p["menge"] is None:
                fehlt.append("Menge")
            if p["ep"] is None:
                fehlt.append("Einzelpreis")
            if p["gp"] is None:
                fehlt.append("Betrag")
        if fehlt:
            vollstaendig_ok = False
            kennung = p["art"] or (p["bez"][:20] if p["bez"] else "?")
            unvollstaendig.append(f"{kennung} ({'/'.join(fehlt)})")

    return {
        "datei": os.path.basename(path),
        "rnr": rnr, "kdnr": kdnr, "email": email, "datum": datum,
        "positionen": positionen, "total": total, "summe": summe,
        "name": name, "strasse": strasse, "hausnr": hausnr, "plz": plz,
        "adresse_zeilen": adresse_zeilen, "sendungsgewicht": sendungsgewicht,
        "zeilen_ok": zeilen_ok, "summe_ok": summe_ok,
        "total_lesbar": total_lesbar,
        "vollstaendig_ok": vollstaendig_ok, "unvollstaendig": unvollstaendig,
    }


# ----------------------------------------------------------------------------
# PDF-AUSGABE
# ----------------------------------------------------------------------------

class CheckBox(Flowable):
    """Kleines leeres Abhak-Kaestchen."""
    def __init__(self, size=10):
        super().__init__()
        self.size = size
        self.width = size
        self.height = size

    def draw(self):
        self.canv.setLineWidth(0.8)
        self.canv.rect(0, 1, self.size, self.size)


class MengeFlow(Flowable):
    """Menge-Anzeige. Bei Menge > 1: rot, fett, GROESSER und eingekreist.
    Die hervorgehobene Menge wird zusaetzlich um `HERVOR_PLUS` Punkt groesser
    gesetzt, damit ein Mehrfach-Artikel wirklich sofort ins Auge faellt."""
    HERVOR_PLUS = 6   # so viele pt groesser als die normale Schrift bei Menge>1

    def __init__(self, text, hervor=False, fs=9):
        super().__init__()
        self.text = text or ""
        self.hervor = hervor
        self.fs = (fs + self.HERVOR_PLUS) if hervor else fs
        self.fs0 = self.fs

    def wrap(self, aw, ah):
        self.aw = aw
        self.fs = self.fs0    # wrap() kann mehrfach mit unterschiedlicher Breite laufen
        if self.hervor:
            # Lange Einheitenwoerter ("2 Schuber", "12 Packungen") duerfen den
            # Kreis nicht ueber die Spalte hinaus ziehen (real beobachtet an
            # Rechnung 1705307: "2 Schuber" ueberdeckte Lagerort/Checkbox).
            # Der Kreis reicht 9 pt (Einzug) + Textbreite + 6 pt (Kreisrand) +
            # ~1 pt Strichstaerke - so weit verkleinern, bis das in `aw` passt,
            # aber nie unter die Normalschrift (9 pt).
            from reportlab.pdfbase.pdfmetrics import stringWidth
            while (self.fs > 9 and
                   9 + stringWidth(self.text, "Helvetica-Bold", self.fs) + 7 > aw):
                self.fs -= 0.5
        self.height = self.fs + (12 if self.hervor else 4)
        return aw, self.height

    def draw(self):
        c = self.canv
        font = "Helvetica-Bold" if self.hervor else "Helvetica"
        c.setFont(font, self.fs)
        tw = c.stringWidth(self.text, font, self.fs)
        x = 9 if self.hervor else 1
        y = (self.height - self.fs) / 2.0 + 1
        if self.hervor:
            c.setFillColor(colors.red)
            c.drawString(x, y, self.text)
            cx = x + tw / 2.0
            cy = y + self.fs * 0.33
            hw = tw / 2.0 + 6
            hh = self.fs / 2.0 + 5
            c.setStrokeColor(colors.red)
            c.setLineWidth(1.6)
            c.ellipse(cx - hw, cy - hh, cx + hw, cy + hh)
        else:
            c.setFillColor(colors.black)
            c.drawString(x, y, self.text)


def mengentext(menge, einh):
    if menge is None:
        return ""
    if abs(menge - round(menge)) < 1e-9:
        m = str(int(round(menge)))
    else:
        m = ("%.2f" % menge).replace(".", ",")
    return f"{m} {einh}".strip()


# ---------------------------------------------------------------------------
# Gewichtsartikel (Schweissdraht/-staebe u.ae., Verkauf nach kg)
# ---------------------------------------------------------------------------
# Erkennung (Stand 2026-08-21j, Parsing-Fix 2026-09-04): EXPLIZITE Markierung
# "<X>,<Y>kg-Multipack" als EIGENE ZEILE im Positionsblock (siehe
# MULTIPACK_RE/parse_block weiter oben) - analog zur "<N>-Fach-Artikel"-
# Markierung bei Set-Artikeln. Bewusst KEINE Ableitung mehr aus der
# Artikelnummer (vorher: kg-Text musste zur letzten Zahl der Artikelnummer
# passen). Diese Inferenz war ein wiederkehrendes Fehlausloeser-Risiko bei
# irrefuehrenden Artikelnummern/-bezeichnungen (Beispiele aus der Praxis:
# "AutoPR-11" mit "...fuer 11-kg-Flaschen" im Text, "Chlorgranulat 5kg" als
# Packungsgroesse) und wurde auf ausdruecklichen Kundenwunsch durch den
# eindeutigen "-Multipack"-Marker ersetzt.
#
# BUGFIX 2026-09-04 (real beobachtet an Rechnung 1703025/WSG2-1,6-5): der
# Marker wurde urspruenglich per MULTIPACK_RE.search() INNERHALB des schon
# fertigen Bezeichnungstexts gesucht - dabei aber uebersehen, dass eine
# Markierungszeile wie "5kg-Multipack" GENAUSO wie "Lagerort:"/"N-Fach-
# Artikel" links-buendig (x0 < ART_MAX) im PDF steht und von parse_block()
# schlicht VERWORFEN wurde, bevor der Bezeichnungstext ueberhaupt
# zusammengesetzt war - artikel_gewicht() sah die Markierung also nie. Jetzt
# wird die Zeile in parse_block() wie "fach" als eigenes Feld "gewicht"
# erkannt/konsumiert; die tatsaechlich SICHTBARE Gewichtsangabe im
# Bezeichnungstext (z.B. "5,00 kg", eigenstaendig neben "5kg-Multipack" -
# analog zu "10 Stück" neben "10-Fach-Artikel") wird separat ueber
# _gewicht_token_re() gefunden und ersetzt.
def artikel_gewicht(p):
    """Stueckgewicht in kg, wenn diese Position die explizite Markierung
    '<X>,<Y>kg-Multipack' als eigene Zeile im Positionsblock hatte, sonst
    None. Keine Artikelnummer-/Bezeichnungs-Ableitung."""
    return p.get("gewicht")


def _gewicht_token_re(w):
    """Regex fuer die SICHTBARE Gewichtsangabe im Bezeichnungstext eines
    Gewichtsartikels, die durch GESAMTMENGE ersetzt wird (z.B. '5,00 kg' bei
    einem Artikel mit der Markierungszeile '5kg-Multipack' - real beobachtet
    an Rechnung 1703025/WSG2-1,6-5: Marker nennt die Zahl bar als '5', der
    Bezeichnungstext aber als '5,00 kg' mit Leerzeichen). Akzeptiert daher
    sowohl die genaue Schreibweise des Markers als auch die uebliche
    2-Nachkommastellen-Variante, jeweils mit optionalem Leerzeichen vor 'kg'.
    Passt keine Zahl im Text (z.B. weil der Katalogtext das Gewicht in einer
    dritten Schreibweise nennt), bleibt die Bezeichnung unveraendert - die
    Menge-Spalte zeigt trotzdem korrekt die tatsaechliche Gesamtmenge
    (gleiches Verhalten wie bei _fach_token_re, s.u.)."""
    varianten = {f"{w:.2f}".replace(".", ",")}
    if float(w).is_integer():
        varianten.add(str(int(w)))
    else:
        varianten.add(f"{w:g}".replace(".", ","))
    alt = "|".join(re.escape(v) for v in sorted(varianten, key=len, reverse=True))
    # (?<![\d,.]) statt nur (?<!\d): sonst wuerde z.B. bei Marker 5 in "1,5 kg"
    # nur die "5" hinter dem Komma ersetzt ("1,GESAMTMENGE").
    return re.compile(rf"(?<![\d,.])(?:{alt})\s*kg\b", re.IGNORECASE)


# ---------------------------------------------------------------------------
# Packungsartikel (Artikelnummer-Form "<Zahl>/<N>") - bewusst KEINE
# Mengen-/Text-Erkennung mehr (Stand 2026-08-21i).
# ---------------------------------------------------------------------------
# Frueher wurde aus "<Zahl>/<N>" in der Artikelnummer + "<N> Stück" im Text
# automatisch auf eine Packungsgroesse geschlossen (analog zu den
# Gewichtsartikeln). Das fuehrte zu Fehlausloesern bei irrefuehrenden
# Artikelbezeichnungen (z.B. Rechnung 1701626, 299/6: "6 Flaschenkappen" wurde
# faelschlich als 6er-Packung erkannt, obwohl es keine ist). Auf Wunsch des
# Kunden aendert eine reine Artikelnummer-/Text-Heuristik die Menge nun NICHT
# mehr - nur die explizite Markierung "<N>-Fach-Artikel" in der Bezeichnung
# (siehe FACH_ARTIKEL_RE) und die Gewichtsartikel-Erkennung (kg, s.u.) duerfen
# die Packliste-Menge/-Bezeichnung noch veraendern.


def _fach_token_re(n):
    """Regex fuer das Mengen-Token eines Set-Artikels in dessen Bezeichnung:
    '<n>x' (z.B. '2x' in '2x Flaschenkappe'), '<n> Stück' als eigene Zeile
    (z.B. '10 Stück' bei einer Gasduese, die als 10er-Set verkauft wird -
    real beobachtet an Rechnung 1701529/MB15-GD12-10), '<n> Packung(en)'
    (real beobachtet an Rechnung 1701663/BP-TILL-2), oder eine der
    Gebinde-Einheiten Kartusche(n)/Flasche(n)/Dose(n)/Eimer/Rolle(n) (auf
    Kundenwunsch ergaenzt, 2026-08-21m/n - dieselben Woerter, die auch als
    Verpackungseinheit in Bezeichnungen wie '4 Kartuschen', '6 Flaschen',
    '3 Dosen', '2 Eimer', '5 Rollen' vorkommen). Kommt keine dieser
    Schreibweisen im Katalog vor (z.B. weil die Fach-Artikel-Markierung eine
    eigenstaendige Mehrfachverkaufs-Angabe ist, die nichts mit einer Zahl in
    der Bezeichnung zu tun hat - Rechnung 1701528/52107-4: '2-Fach-Artikel',
    aber Bezeichnung nennt '4x' fuer den Karton-Inhalt), bleibt die
    Bezeichnung unveraendert; die Menge-Spalte zeigt trotzdem korrekt die
    tatsaechliche Stueckzahl."""
    # (?<![\d,.]) statt nur (?<!\d): "1,2 Stück" darf bei n=2 nicht als "2 Stück"
    # gelten (Dezimalzahl in der Bezeichnung).
    return re.compile(
        rf"(?<![\d,.]){n}\s*(?:x\b|St(?:ü|ue)?c?ke?\b|Packung(?:en)?\b|"
        rf"Kartusche(?:n)?\b|Flasche(?:n)?\b|Dose(?:n)?\b|Eimer\b|Rolle(?:n)?\b)",
        re.IGNORECASE)


def effektive_menge(p):
    """(menge, einh, bez, immer_hervorheben) - die TATSAECHLICH zu verpackende
    Menge/Einheit/Bezeichnung einer Position, nach Gewichts- oder Set-Artikel-
    Multiplikator (Mengen-Token in der Bezeichnung durch 'GESAMTMENGE'
    ersetzt). Bewusst KEINE Ableitung aus der Bezeichnung/Artikelnummer (Stand
    2026-08-21j) - nur zwei EXPLIZITE Text-Marker duerfen die Menge/
    Bezeichnung aendern: '<X>,<Y>kg-Multipack' (Gewichtsartikel, siehe
    artikel_gewicht) und '<N>-Fach-Artikel' (Set-Artikel, Feld 'fach');
    alles andere bleibt wie im PDF, auch wenn die Artikelnummer oder der Text
    irrefuehrend nach etwas anderem aussieht. Bei
    einem normalen Artikel unveraendert (menge/einh/bez wie im PDF,
    immer_hervorheben=False). Gemeinsame Grundlage fuer die Packuebersicht
    (pack_anzeige, s.u.) UND die aggregierte Sammelliste (Teil A) - beide
    sollen dieselbe tatsaechliche Stueckzahl/Einheit zeigen, nicht nur die
    Packuebersicht."""
    qty = p["menge"] if p["menge"] is not None else 1
    gw = artikel_gewicht(p)
    if gw is not None:
        bez = _gewicht_token_re(gw).sub("GESAMTMENGE", p["bez"], count=1)
        return (gw * qty, "kg", bez, True)
    fach = p.get("fach")
    if fach is not None and fach > 1:
        bez = _fach_token_re(fach).sub("GESAMTMENGE", p["bez"], count=1)
        return (fach * qty, "Stück", bez, True)
    return (p["menge"], p["einh"], p["bez"], False)


def pack_anzeige(p):
    """Anzeige (bez, mengentext, hervorheben) fuer die PACKUEBERSICHT - siehe
    effektive_menge() fuer die Multiplikator-Logik (Gewichts-/Set-Artikel). Bei einem normalen Artikel bleibt die bisherige Menge>1-
    Hervorhebung bestehen."""
    menge, einh, bez, immer_hv = effektive_menge(p)
    hv = immer_hv or (menge or 0) > 1.0000001
    return (bez, mengentext(menge, einh), hv)


def scan_bedarf(eff):
    """Noetige Scans fuer EINE Position mit effektiver Menge eff:
    genau 1 -> 1, unter 1 (z.B. 0,5 kg) -> 2, sonst aufgerundet (gedeckelt 5)."""
    if eff is None or eff <= 0:
        return 1
    if abs(eff - 1.0) < 1e-9:
        return 1
    if eff < 1:
        return 2
    e = round(eff, 6)
    ganz = int(e) + (1 if e > int(e) else 0)
    return min(5, ganz)


def order_scananzahl(r):
    """Noetige Scans fuer eine ganze Bestellung = Summe der Positions-Scans,
    gedeckelt auf 5 (mind. 1). Gewichts- UND Set-Artikel (N-Fach-Artikel)
    zaehlen nach ihrer TATSAECHLICHEN Stueckzahl (Gewicht x Menge bzw.
    Set-Groesse x Menge) statt der rohen Bestellmenge - sonst wuerde die
    Mehrfach-Scan-Sicherung z.B. bei einem als '1x' bestellten 2er-Set-Artikel
    nur 1 statt 2 Scans verlangen und ein vergessenes zweites Teil nicht
    abfangen (2026-08-21c nachgezogen, analog zur bereits bestehenden
    Gewichtsartikel-Behandlung). Alle anderen Positionen zaehlen nach der
    rohen Bestellmenge (Packungsartikel-Ableitung gibt es seit 2026-08-21i
    nicht mehr)."""
    total = 0
    for p in r["positionen"]:
        if ist_versand(p["art"], p["bez"]):
            continue
        qty = p["menge"] if p["menge"] is not None else 1
        gw = artikel_gewicht(p)
        fach = p.get("fach")
        if gw is not None:
            eff = gw * qty
        elif fach is not None and fach > 1:
            eff = fach * qty
        else:
            eff = qty
        total += scan_bedarf(eff)
    return max(1, min(5, total))


def _artikel_sortkey(s):
    """Natuerliche Sortierung: Zahlen numerisch (auf 8 Stellen aufgefuellt)."""
    return re.sub(r"\d+", lambda m: m.group().zfill(8), (s or "").lower())


def rahmen_block(flowables, breite, dashed=False):
    """Flowables in einen umrahmten 1x1-Tabellenblock packen (Rahmen optional
    gestrichelt)."""
    t = Table([[flowables]], colWidths=[breite])
    box = (("BOX", (0, 0), (-1, -1), 1.0, colors.HexColor("#555555"), None, (3, 2))
           if dashed else ("BOX", (0, 0), (-1, -1), 0.9, colors.HexColor("#37474F")))
    t.setStyle(TableStyle([
        box,
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
        ("RIGHTPADDING", (0, 0), (-1, -1), 8),
        ("TOPPADDING", (0, 0), (-1, -1), 7),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
        ("BACKGROUND", (0, 0), (-1, -1), colors.white),
    ]))
    return t


def _fuss_zeichnen(canvas, doc, text):
    """Kleiner grauer Fuss unten links auf jeder Seite (z.B. Erstellungsdatum)."""
    canvas.saveState()
    canvas.setFont("Helvetica", 7)
    canvas.setFillGray(0.5)
    canvas.drawString(14 * mm, 7 * mm, text)
    canvas.restoreState()


def _story_to_reader(story, fusstext=None):
    """Eine Flowable-Liste zu einem PdfReader rendern (fuer seitengenaues
    Zusammenfuehren). fusstext (optional) wird auf JEDER Seite unten abgedruckt."""
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=14 * mm, rightMargin=14 * mm,
                            topMargin=12 * mm, bottomMargin=12 * mm)
    if fusstext:
        cb = lambda c, d: _fuss_zeichnen(c, d, fusstext)
        doc.build(story, onFirstPage=cb, onLaterPages=cb)
    else:
        doc.build(story)
    buf.seek(0)
    return PdfReader(buf)


def baue_pdf(rechnungen, pdf_pfad, gruppen):
    styles = getSampleStyleSheet()
    st_titel = ParagraphStyle("titel", parent=styles["Title"], fontSize=16, spaceAfter=4)
    st_cell = ParagraphStyle("cell", parent=styles["Normal"], fontSize=9, leading=11)
    st_klein = ParagraphStyle("klein", parent=styles["Normal"], fontSize=8,
                              textColor=colors.grey)
    st_kopf = ParagraphStyle("kopf", parent=styles["Normal"], fontSize=10, leading=13)
    # Eigener Kopf-Stil fuer die Packuebersicht: grosszuegige Zeilenhoehe, damit
    # der vergroesserte Empfaengername (inline <font size=...>) nicht ueberlappt.
    st_pack_kopf = ParagraphStyle("packkopf", parent=styles["Normal"], fontSize=10, leading=18)
    st_warn = ParagraphStyle("warn", parent=styles["Normal"], fontSize=9,
                            textColor=colors.red)
    st_info = ParagraphStyle("info", parent=styles["Normal"], fontSize=8,
                            textColor=colors.HexColor("#7a6000"))

    SEITE = A4[0] - 28 * mm            # nutzbare Breite (= doc.width)
    anz_rechnungen = len(rechnungen)
    # Kleiner Nachlauf -> alles auf eine gemeinsame Liste statt auf mehrere
    # Lagerort-Kategorien zu verteilen (siehe KOMBINIERT_SCHWELLE oben).
    kombiniert = anz_rechnungen < KOMBINIERT_SCHWELLE
    kategorien = [KOMBINIERT_LABEL] if kombiniert else KATEGORIE_REIHENFOLGE

    def kat_von(r):
        return KOMBINIERT_LABEL if kombiniert else auftrag_kategorie(r)

    # Kopf-Tabellenstil fuer Listen
    def _liste_style():
        return TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#37474F")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONTSIZE", (0, 0), (-1, 0), 9),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F2F4F5")]),
            ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#CFD8DC")),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("ALIGN", (-1, 0), (-1, -1), "CENTER"),     # Kaestchen-Spalte zentriert
            ("LEFTPADDING", (0, 0), (-1, -1), 4),
            ("RIGHTPADDING", (0, 0), (-1, -1), 4),
            ("TOPPADDING", (0, 0), (-1, -1), 3),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ])

    # ----- Aggregation Teil A: Kategorie -> art -> {bez, einh, lager, menge} ---
    # Die Kommissionierliste eines Lagerorts listet ALLE Positionen der
    # Bestellungen, die DORT gepackt werden (Haupt-Lagerort = auftrag_kategorie),
    # nicht nur die dort gelagerten Artikel. So steht z.B. ein Wapo-Artikel, der
    # zusammen mit Topseller/Poolchemie verschickt wird, auf DEREN Liste - er muss
    # ja mitverpackt werden. Die Spalte "Lagerort" zeigt, wo der Artikel liegt.
    # Dadurch sind Kommissionierliste (Teil A) und Packuebersicht (Teil B)
    # konsistent, und jeder Artikel steht genau einmal (kein Doppel-Picken).
    # Menge/Einheit/Bezeichnung kommen ueber effektive_menge() - bis 2026-08-21c
    # wurde hier IMMER die rohe Bestellmenge summiert, auch bei Gewichts-/
    # Packungs-/Set-Artikeln (die Packuebersicht zeigte schon die tatsaechliche
    # Stueckzahl, die Sammelliste oben auf der Seite aber nur die rohe Menge -
    # bewusst nachgeruestet, damit beide Teile konsistent dieselbe Menge zeigen).
    kat_artikel = {k: {} for k in kategorien}
    for r in rechnungen:
        kat = kat_von(r)
        for p in r["positionen"]:
            if ist_versand(p["art"], p["bez"]):
                continue
            key = p["art"] or p["bez"]
            eff_menge, eff_einh, eff_bez, immer_hv = effektive_menge(p)
            eintrag = kat_artikel[kat].setdefault(
                key, {"art": p["art"], "bez": eff_bez, "einh": eff_einh,
                      "lager": set(), "menge": 0.0, "immer_hv": False})
            eintrag["menge"] += (eff_menge or 0)
            eintrag["immer_hv"] = eintrag["immer_hv"] or immer_hv
            for lo in p["lagerorte"]:
                eintrag["lager"].add(lo)

    def liste_story(artikel, titel, untertitel):
        """Eine (Kategorie-)Liste als Flowables: Kaestchen RECHTS, nach
        Artikelnummer sortiert, Menge>1 rot/fett/eingekreist."""
        story = [Paragraph(titel, st_titel)]
        if untertitel:
            story.append(Paragraph(untertitel, st_klein))
        story.append(Spacer(1, 6))
        head = ["Lagerort", "Artikelnr", "Bezeichnung", "Menge", ""]
        lw, aw, mw, cw = 66, 80, 100, 26   # mw 80 -> 100: lange Einheitenwoerter (2026-09-21c)
        bw = SEITE - lw - aw - mw - cw
        data = [head]
        for key in sorted(artikel, key=lambda k: _artikel_sortkey(artikel[k]["art"] or k)):
            a = artikel[key]
            lager = ", ".join(sorted(a["lager"])) if a["lager"] else "-"
            hv = a.get("immer_hv") or (a["menge"] or 0) > 1.0000001
            data.append([
                Paragraph(lager, st_cell),
                Paragraph(a["art"] or "-", st_cell),
                Paragraph(a["bez"], st_cell),
                MengeFlow(mengentext(a["menge"], a["einh"]), hv),
                CheckBox(),
            ])
        t = Table(data, colWidths=[lw, aw, bw, mw, cw], repeatRows=1)
        t.setStyle(_liste_style())
        story.append(t)
        return story

    def einzel_block(r):
        """Eine Bestellung als solider Rahmen (Barcode, Empfaenger, Positionen).
        Ohne Kundennummer; Kaestchen rechts; Menge>1 hervorgehoben."""
        inner = SEITE - 16
        pos = [p for p in r["positionen"] if not ist_versand(p["art"], p["bez"])]
        try:
            bc = code128.Code128(r["rnr"], barHeight=14 * mm, barWidth=0.42 * mm)
        except Exception:
            bc = Paragraph(r["rnr"], st_kopf)
        kopf_rechts = (
            f"<font size=8>Rechnung Nr. {r['rnr']} &nbsp;&middot;&nbsp; {r['datum']}</font>"
            f"<br/><font size=15><b>{r['name']}</b></font>"
        )
        kopf = Table([[bc, Paragraph(kopf_rechts, st_pack_kopf)]],
                     colWidths=[150, inner - 150])
        kopf.setStyle(TableStyle([
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("LEFTPADDING", (0, 0), (0, 0), 0),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
        ]))
        inhalt = [kopf]
        if not (r["zeilen_ok"] and r["summe_ok"] and r.get("vollstaendig_ok", True)):
            details = []
            if not r["zeilen_ok"]:
                details.append("Zeile Menge×EP≠GP")
            if not r["summe_ok"]:
                details.append(f"Summe {('%.2f' % r['summe']).replace('.', ',')} ≠ "
                                f"Rechnungsbetrag {('%.2f' % (r['total'] or 0)).replace('.', ',')}")
            if not r.get("vollstaendig_ok", True):
                details.append("Position unvollständig: "
                               + ", ".join(r.get("unvollstaendig", [])))
            inhalt.append(Paragraph("⚠ Prüfung NICHT bestanden: " + "; ".join(details), st_warn))
        elif not r.get("total_lesbar", True):
            inhalt.append(Paragraph(
                "ℹ Rechnungsbetrag im PDF nicht maschinell lesbar (vermutlich sehr "
                "lange Bezeichnung) – Positionssumme "
                + ('%.2f' % r['summe']).replace('.', ',') + " €, Positionen geprüft",
                st_info))

        head2 = ["Artikelnr", "Bezeichnung", "Menge", "Lagerort", ""]
        aw, mw, lw, cw = 74, 100, 92, 26   # mw 80 -> 100: lange Einheitenwoerter (2026-09-21c)
        bw = inner - aw - mw - lw - cw
        data = [head2]
        for p in pos:
            lager = ", ".join(p["lagerorte"]) if p["lagerorte"] else "-"
            bez_disp, menge_disp, hv = pack_anzeige(p)
            data.append([
                Paragraph(p["art"] or "-", st_cell),
                Paragraph(bez_disp, st_cell),
                MengeFlow(menge_disp, hv),
                Paragraph(lager, st_cell),
                CheckBox(),
            ])
        t = Table(data, colWidths=[aw, bw, mw, lw, cw], repeatRows=1)
        t.setStyle(_liste_style())
        inhalt.append(Spacer(1, 4))
        inhalt.append(t)
        return KeepTogether([rahmen_block(inhalt, SEITE), Spacer(1, 8)])

    def gruppe_block(g, gruppe_rechnungen):
        """Sammeldruck-Gruppe: alle betroffenen Bestellungen in einem
        GESTRICHELTEN Rahmen, der Sammel-Barcode unten INNERHALB des Rahmens."""
        n = len(gruppe_rechnungen)
        titel = Paragraph(
            f"<b>Sammeldruck {g['code']}</b> &middot; Artikel <b>{g['art']}</b> &ndash; "
            f"{g['bez']} &middot; {n} Bestellungen à {g.get('menge_je_text') or '1 Stück'}",
            st_kopf)
        try:
            bc = code128.Code128(g["code"], barHeight=18 * mm, barWidth=0.5 * mm)
        except Exception:
            bc = Paragraph(g["code"], st_kopf)
        hint = Paragraph(
            f"<para align='center'>Barcode <b>{g['code']}</b> scannen &rarr; "
            f"alle {n} Versandlabel auf einmal drucken</para>", st_klein)
        c_rnr, c_plz, c_chk = 90, 60, 26
        c_emp = SEITE - c_rnr - c_plz - c_chk
        # Reihenfolge im Rahmen: Titel -> Deckblatt (Barcode) -> Hinweis ->
        # dann erst die Auftragsliste. (Deckblatt VOR den Auftraegen.)
        data = [[titel, "", "", ""],
                [bc, "", "", ""],
                [hint, "", "", ""],
                [Paragraph("<b>Rechnung</b>", st_cell), Paragraph("<b>Empfänger</b>", st_cell),
                 Paragraph("<b>PLZ</b>", st_cell), ""]]
        kopf_rows = 4                 # Titel, Barcode, Hinweis, Tabellenkopf
        for r in gruppe_rechnungen:
            data.append([Paragraph(r["rnr"], st_cell), Paragraph(r["name"], st_cell),
                         Paragraph(r["plz"], st_cell), CheckBox()])
        last = len(data) - 1
        t = Table(data, colWidths=[c_rnr, c_emp, c_plz, c_chk], repeatRows=kopf_rows)
        t.setStyle(TableStyle([
            ("BOX", (0, 0), (-1, -1), 1.1, colors.HexColor("#444444"), None, (3, 2)),
            ("SPAN", (0, 0), (-1, 0)),                 # Titel
            ("SPAN", (0, 1), (-1, 1)),                 # Barcode (Deckblatt)
            ("SPAN", (0, 2), (-1, 2)),                 # Hinweis
            ("ALIGN", (0, 1), (-1, 2), "CENTER"),      # Barcode + Hinweis zentriert
            ("LINEBELOW", (0, 3), (-1, 3), 0.5, colors.HexColor("#999999")),
            ("ALIGN", (-1, kopf_rows), (-1, last), "CENTER"),   # Kaestchen-Spalte
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING", (0, 0), (-1, -1), 4),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ("LEFTPADDING", (0, 0), (-1, -1), 6),
            ("RIGHTPADDING", (0, 0), (-1, -1), 6),
        ]))
        # Eine Sammeldruck-Gruppe nie ueber einen Seitenumbruch trennen.
        return KeepTogether([t])

    # ============================ PDF zusammenbauen =========================
    writer = PdfWriter()

    # Erstellungszeitpunkt einmal festhalten -> auf jeder Seite als Fuss.
    erstellt = datetime.now()
    fusstext = f"Erstellt am {erstellt:%d.%m.%Y} um {erstellt:%H:%M} Uhr"

    def add_story(story, pad_even=False):
        reader = _story_to_reader(story, fusstext)
        for p in reader.pages:
            writer.add_page(p)
        if pad_even and len(reader.pages) % 2 == 1:
            writer.add_blank_page(width=float(A4[0]), height=float(A4[1]))

    # ---- Bestellungen den Lagerorten zuordnen --------------------------------
    # Jede Bestellung kommt KOMPLETT unter genau EINEN Lagerort (Haupt-Lagerort
    # = hoechste Prioritaet ueber ihre Positionen, siehe auftrag_kategorie()).
    # Sammeldruck-Gruppen (immer ein einzelner Artikel) ordnen sich ueber diesen
    # Artikel demselben Lagerort zu.
    # Die Kommissionierliste je Lagerort listet GENAU die Positionen dieser
    # Bestellungen (siehe kat_artikel oben) - also alles, was an diesem Packplatz
    # verpackt wird, inkl. Artikeln, die woanders gelagert sind (Spalte Lagerort).
    grouped_rnr = set(rnr for g in gruppen for rnr in g["rnr"])
    rnr_map = {r["rnr"]: r for r in rechnungen}
    einzel = [r for r in rechnungen if r["rnr"] not in grouped_rnr]

    einzel_nach_kat = {k: [] for k in kategorien}
    for r in einzel:
        einzel_nach_kat[kat_von(r)].append(r)

    gruppen_nach_kat = {k: [] for k in kategorien}
    for g in gruppen:
        gr = [rnr_map[x] for x in g["rnr"] if x in rnr_map]
        kat = kat_von(gr[0]) if gr else (KOMBINIERT_LABEL if kombiniert else "Allgemein")
        gruppen_nach_kat[kat].append((g, gr))

    def packuebersicht_flowables(einzel_k, gruppen_k):
        """Packuebersicht-Teil EINES Lagerorts: erst die Einzel-Bestellungen,
        dann die Sammeldruck-Gruppen. Jeder Block dank KeepTogether ungebrochen.
        Ohne zugeordnete Bestellung -> kurzer Hinweis."""
        fl = [Spacer(1, 10),
              Paragraph("Packübersicht je Rechnung", st_titel),
              Paragraph("Eine Bestellung je Rahmen &middot; Versandkosten ausgeblendet "
                        "&middot; Sammeldruck-Gruppen gestrichelt umrahmt", st_klein),
              Spacer(1, 6)]
        if not einzel_k and not gruppen_k:
            fl.append(Paragraph(
                "Keine Bestellung hat diesen Lagerort als Haupt-Lagerort &ndash; die "
                "oben gelisteten Artikel gehören zu Bestellungen, die bei deren "
                "Haupt-Lagerort gepackt werden.", st_klein))
            return fl
        for r in einzel_k:
            fl.append(einzel_block(r))
        for g, gr in gruppen_k:
            fl.append(gruppe_block(g, gr))
            fl.append(Spacer(1, 10))
        return fl

    # ---- Pro Lagerort: Kommissionierliste + DIREKT die Packuebersicht --------
    # Jeder Lagerort beginnt auf einem NEUEN Blatt (eigene Handout-Liste je
    # Packplatz). Nur Lagerorte mit Inhalt erscheinen.
    story = []
    erste = True
    for kat in kategorien:
        artikel = kat_artikel[kat]
        einzel_k = einzel_nach_kat[kat]
        gruppen_k = gruppen_nach_kat[kat]
        if not artikel and not einzel_k and not gruppen_k:
            continue
        if not erste:
            story.append(PageBreak())
        erste = False
        anz_best = len(einzel_k) + sum(len(gr) for _g, gr in gruppen_k)
        titel = "Kommissionierliste" if kombiniert else f"Kommissionierliste &ndash; {kat}"
        story += liste_story(
            artikel, titel,
            f"{len(artikel)} Artikel zu kommissionieren &nbsp;&middot;&nbsp; "
            f"{anz_best} Bestellung(en) zum Packen &nbsp;&middot;&nbsp; "
            f"sortiert nach Artikelnummer")
        story += packuebersicht_flowables(einzel_k, gruppen_k)

    if not story:
        story = [Paragraph("Keine auswertbaren Positionen.", st_titel)]
    add_story(story, pad_even=False)

    try:
        with open(pdf_pfad, "wb") as f:
            writer.write(f)
    except PermissionError:
        raise SystemExit(
            "\n*** FEHLER: Die Pickliste-PDF kann nicht geschrieben werden:\n"
            f"    {pdf_pfad}\n"
            "    Die Datei ist vermutlich noch GEOEFFNET (im PDF-Betrachter) oder\n"
            "    schreibgeschuetzt. Bitte das Pickliste-Fenster schliessen und die\n"
            "    Batch erneut starten. (Es wurde nichts kopiert/archiviert.)\n")


# ----------------------------------------------------------------------------
# post_zuordnung.csv
# ----------------------------------------------------------------------------

# ----------------------------------------------------------------------------
# post_zuordnung.csv / mengen_zuordnung.csv / ean_zuordnung.csv
# ----------------------------------------------------------------------------
# Diese drei Dateien werden bei JEDEM Lauf mit einer eventuell BESTEHENDEN
# Datei ZUSAMMENGEFUEHRT statt sie zu ueberschreiben. Grund: nach jedem Lauf
# werden die eingelesenen Rechnungs-PDFs archiviert (verschwinden aus dem
# Eingangsordner, siehe archiviere()) - ein zweiter Lauf saehe sonst nur noch
# die NEUEN Rechnungen, und die Zuordnung fuer noch nicht gescannte
# Bestellungen aus einem frueheren Lauf wuerde ersatzlos verschwinden. Das
# faellt bei der Deutschen Post am haertesten auf: dort steht KEINE
# Rechnungsnummer auf dem Label, der Abgleich laeuft ausschliesslich ueber
# post_zuordnung.csv (PLZ+Hausnummer) - fehlt der Eintrag, kann scan_druck.py
# das Label ueberhaupt nicht mehr zuordnen.
#
# Damit die Dateien dabei nicht unbegrenzt weiterwachsen, verfallen alte
# Eintraege nach BRUECKEN_CSV_MERGE_TAGE automatisch (Rechnungsdatum als
# LETZTE Spalte in jeder Datei angehaengt, damit eine aeltere scan_druck.py-
# Version, die diese Spalte nicht kennt, die Dateien weiterhin lesen kann).
BRUECKEN_CSV_MERGE_TAGE = 5


def _bruecken_datum_frisch(datum_str, max_tage=BRUECKEN_CSV_MERGE_TAGE):
    """True, wenn ein 'TT.MM.JJJJ'-Rechnungsdatum noch nicht aelter als
    max_tage ist. Unlesbares/leeres Datum gilt als NICHT frisch (Stand
    2026-09-01 - vorher als 'immer frisch' behandelt).

    FRUEHERER BUG (real beobachtet, fuehrte zum MemoryError-Absturz vom
    2026-09-01): ein leeres/unlesbares Datum wurde 'vorsichtshalber' als
    dauerhaft frisch gewertet, um einen gueltigen Eintrag nicht zu frueh zu
    verlieren. Da eine per merge() weitergereichte Zeile ihr Datum NIE
    aendert (das Feld wird nur roh kopiert), blieb ein Eintrag ohne Datum
    dadurch fuer immer 'frisch' und wuchs unbegrenzt weiter - genau die
    Bedingung, unter der der o.g. CSV-Quotierungsbug ungebremst exponentiell
    eskalieren konnte. Der eigentliche Schutzzweck (eine noch offene
    Bestellung nicht zu frueh verlieren) wird jetzt anders sichergestellt:
    schreibe_csv()/schreibe_mengen_csv()/schreibe_ean_csv()/
    schreibe_wc_bestellnummern_csv() tragen beim Schreiben IMMER ein
    gueltiges Datum ein (siehe _csv_datum()) - notfalls das heutige, wenn
    parse_pdf() keines aus der Rechnung extrahieren konnte. Sammelgruppen
    (schreibe_sammel_csv) haben kein einzelnes Rechnungsdatum - dort setzt
    merge_sammelgruppen() beim Anlegen einer neuen Gruppe ebenso das
    heutige Datum. Ein Eintrag ohne Datum kann in der Datei also nur noch
    aus einer AELTEREN Version dieses Skripts stammen und wird jetzt bewusst
    verworfen statt endlos mitgeschleppt."""
    if not datum_str:
        return False
    try:
        d = datetime.strptime(datum_str.strip(), "%d.%m.%Y")
    except ValueError:
        return False
    return (datetime.now() - d).days <= max_tage


def _csv_datum(r):
    """Rechnungsdatum fuer eine Bruecken-CSV-Zeile: das aus der Rechnung
    extrahierte Datum, sonst (parse_pdf() fand keines) ersatzweise das
    heutige Datum - NIE leer. Nur so bleibt _bruecken_datum_frisch() fuer
    diese Zeile spaeter aussagekraeftig (s. dortiger Kommentar); ein leeres
    Rechnungsdatum wuerde die Zeile sonst beim naechsten Merge sofort als
    'nicht frisch' verwerfen, obwohl die Bestellung noch offen sein kann."""
    return (r.get("datum") or "").strip() or datetime.now().strftime("%d.%m.%Y")


def _lies_bestehende_zeilen(csv_pfad, mindest_spalten):
    """Bestehende Bruecken-CSV zeilenweise als Liste von Tokenlisten einlesen
    (roh, vor der eigentlichen Verarbeitung). Bei fehlender/kaputter Datei:
    leere Liste (verhaelt sich dann wie 'erster Lauf', nichts zum
    Zusammenfuehren).

    KRITISCHER FIX (2026-09-01, real beobachtet - ean_zuordnung.csv wuchs auf
    ~940 MB und liess packliste.py mit MemoryError abstuerzen): las bisher mit
    einem simplen zeile.split(";") statt echtem CSV-Parsing. Ein Bezeichnungs-
    text, der ein Semikolon UND ein Anfuehrungszeichen enthaelt (z.B. Rechnung
    1700713: 'Regler 8,0kg/h; 0,5-4bar; Kombi x G 3/8"LH-KN; ...' - das "
    steht fuer Zoll, kommt bei Schlauch-/Gewindeartikeln haeufig vor), wird
    von csv.writer beim Schreiben korrekt in Anfuehrungszeichen gesetzt und
    das interne " dabei verdoppelt (CSV-Standard). Der naive split(";") kannte
    diese Quotierung nicht: er zerlegte das Feld an den EINGEBETTETEN
    Semikolons und uebernahm die (nicht entfernten) Anfuehrungszeichen als
    literalen Text. Wurde diese Zeile beim naechsten Lauf erneut geschrieben,
    hat csv.writer die inzwischen literalen Anfuehrungszeichen ERNEUT
    verdoppelt - bei jedem Lauf, der diese Rechnung noch als 'frisch' fand,
    verdoppelte sich die Anzahl der Anfuehrungszeichen (EXPONENTIELLES
    Wachstum: 2 -> 4 -> 8 -> ... -> nach ca. 30 Laeufen ueber 900 Millionen).
    Jetzt echtes CSV-Parsing (csv.reader) statt split(";") - Quotierung wird
    korrekt aufgeloest, das Feld bleibt bei jedem Lese-/Schreibzyklus gleich
    lang. Betraf alle Bruecken-CSVs, die diese Funktion nutzen (post_/sammel_/
    mengen_/ean_zuordnung.csv), ausgeloest ist es aber offenbar bisher nur bei
    ean_zuordnung.csv (die anderen drei sind unauffaellig gross)."""
    if not os.path.exists(csv_pfad):
        return []
    zeilen = []
    try:
        with open(csv_pfad, encoding="utf-8-sig", newline="") as f:
            reader = csv.reader(f, delimiter=";")
            next(reader, None)  # Kopfzeile
            for t in reader:
                if len(t) >= mindest_spalten and t and t[0]:
                    zeilen.append(t)
    except Exception:
        return []
    return zeilen


def schreibe_csv(rechnungen, csv_pfad):
    """post_zuordnung.csv fuer scan_druck.py - v.a. fuer das Deutsche-Post-
    Matching per PLZ+Hausnummer (auf dem Label selbst steht keine
    Rechnungsnummer). Format:
        Rechnungsnummer;Name;Strasse;Hausnummer;PLZ;Rechnungsdatum
    Wird mit einer eventuell bestehenden Datei zusammengefuehrt, siehe
    BRUECKEN_CSV_MERGE_TAGE weiter oben."""
    bestand = {}
    for t in _lies_bestehende_zeilen(csv_pfad, 5):
        datum = t[5] if len(t) > 5 else ""
        if _bruecken_datum_frisch(datum):
            bestand[t[0]] = [t[0], t[1], t[2], t[3], t[4], datum]
    for r in rechnungen:
        if not r["rnr"]:
            continue
        bestand[r["rnr"]] = [r["rnr"], r["name"], r["strasse"], r["hausnr"],
                              r["plz"], _csv_datum(r)]
    with open(csv_pfad, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f, delimiter=";", lineterminator="\n")
        w.writerow(["Rechnungsnummer", "Name", "Strasse", "Hausnummer",
                     "PLZ", "Rechnungsdatum"])
        for row in bestand.values():
            w.writerow(row)


def schreibe_sammel_csv(gruppen, csv_pfad):
    """Bruecke fuer den Sammeldruck in scan_druck.py.
    Format: Code;Artikelnr;Bezeichnung;Menge;Rechnungsnummern;EAN;Rechnungsdatum;
    MengeJeBestellung (Rechnungsnummern kommasepariert). Die EAN-Spalte ist nur
    gefuellt, wenn der Artikel tatsaechlich eine EAN traegt - scan_druck.py
    verlangt dann vor dem Sammeldruck zusaetzlich zum Sammelcode die
    Bestaetigungs-EAN des Artikels. 'Rechnungsdatum' = Datum, an dem der Code
    ERSTELLT wurde (fuer die Verfallsfrist beim Zusammenfuehren mit einem
    spaeteren Lauf, siehe merge_sammelgruppen/BRUECKEN_CSV_MERGE_TAGE).
    'MengeJeBestellung' (seit 2026-08-21h) = die tatsaechlich JE Bestellung zu
    verpackende Menge als fertiger Text (z.B. "3 Stück" bei einem
    3-Fach-Artikel, sonst "1 Stück") - kommt aus finde_sammelgruppen() ueber
    effektive_menge(). Alle Spalten ab 'EAN' bewusst ans ENDE angehaengt, damit
    eine aeltere scan_druck.py-Version die Datei weiter lesen kann (faellt dann
    nur auf die alte, hartcodierte "1 Stück"-Annahme zurueck).

    `gruppen` ist hier bereits das GEMERGTE Gesamtergebnis aus
    merge_sammelgruppen() (alte, noch frische Gruppen + neue Gruppen dieses
    Laufs) - nicht nur die Gruppen des aktuellen Laufs."""
    with open(csv_pfad, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f, delimiter=";", lineterminator="\n")
        w.writerow(["Code", "Artikelnr", "Bezeichnung", "Menge",
                     "Rechnungsnummern", "EAN", "Rechnungsdatum",
                     "MengeJeBestellung"])
        for g in gruppen:
            w.writerow([g["code"], g["art"], g["bez"], len(g["rnr"]),
                        ",".join(g["rnr"]), g.get("ean", ""), g.get("datum", ""),
                        g.get("menge_je_text") or "1 Stück"])


def schreibe_mengen_csv(rechnungen, csv_pfad):
    """Bruecke fuer die Mehrfach-Scan-Sicherung in scan_druck.py.
    Format: Rechnungsnummer;Scananzahl;Rechnungsdatum  (1..5). Die Scananzahl
    je Bestellung kommt aus order_scananzahl(): Summe der Positions-Scans
    (genau 1 -> 1, unter 1 wie 0,5 kg -> 2, sonst aufgerundete Menge),
    Gewichtsartikel zaehlen nach kg, gedeckelt auf 5. scan_druck.py liest die
    Zahl direkt als noetige Scananzahl - es muss dafuer NICHT geaendert/neu
    deployt werden. Wird mit einer eventuell bestehenden Datei
    zusammengefuehrt, siehe BRUECKEN_CSV_MERGE_TAGE weiter oben."""
    bestand = {}
    for t in _lies_bestehende_zeilen(csv_pfad, 2):
        datum = t[2] if len(t) > 2 else ""
        if _bruecken_datum_frisch(datum):
            bestand[t[0]] = [t[0], t[1], datum]
    for r in rechnungen:
        if not r["rnr"]:
            continue
        bestand[r["rnr"]] = [r["rnr"], order_scananzahl(r), _csv_datum(r)]
    with open(csv_pfad, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f, delimiter=";", lineterminator="\n")
        w.writerow(["Rechnungsnummer", "Scananzahl", "Rechnungsdatum"])
        for row in bestand.values():
            w.writerow(row)


def schreibe_ean_csv(rechnungen, csv_pfad):
    """Bruecke fuer den EAN-Verifikationsscan in scan_druck.py.
    Eine Zeile JE Position MIT EAN:
        Rechnungsnummer;EAN;Bezeichnung;Anzahl;Rechnungsdatum
    `Anzahl` = noetige Verifikations-Scans dieser Position (= scan_bedarf der
    bestellten Menge, gedeckelt auf MAX; bei einem Set-Artikel/N-Fach-Artikel
    die TATSAECHLICHE Stueckzahl, Set-Groesse x Menge - seit 2026-08-21e).
    Positionen OHNE EAN tauchen hier NICHT auf - fuer sie bleibt es bei der
    normalen Mehrfach-Scan-Sicherung ueber die Stueckzahl (mengen_zuordnung.csv).
    So wird der EAN-Scan ausschliesslich bei Artikeln verlangt, die in der
    Pickliste eine EAN tragen. Gewichtsartikel bleiben bei der rohen
    Bestellmenge (laufen ueber mengen_zuordnung.csv, nicht ueber EAN - siehe
    Versionshinweis 2026-08-21e).

    Wird mit einer eventuell bestehenden Datei zusammengefuehrt (gruppiert
    nach Rechnungsnummer): Rechnungen aus dem AKTUELLEN Lauf ersetzen ihre
    alten Zeilen komplett; Rechnungen, die NUR im Bestand stehen (frueherer
    Lauf, noch nicht gescannt), bleiben unveraendert erhalten - siehe
    BRUECKEN_CSV_MERGE_TAGE weiter oben."""
    bestand = {}  # rnr -> Liste von Zeilen (inkl. Rechnungsdatum)
    for t in _lies_bestehende_zeilen(csv_pfad, 4):
        datum = t[4] if len(t) > 4 else ""
        if _bruecken_datum_frisch(datum):
            bestand.setdefault(t[0], []).append([t[0], t[1], t[2], t[3], datum])

    for r in rechnungen:
        if not r["rnr"]:
            continue
        zeilen = []
        for p in r["positionen"]:
            ean = (p.get("ean") or "").strip()
            if not ean:
                continue
            if ist_versand(p["art"], p["bez"]):
                continue
            qty = p["menge"] if p["menge"] is not None else 1
            # Set-Artikel (N-Fach-Artikel) auch hier nach TATSAECHLICHER
            # Stueckzahl verlangen (2026-08-21e): am gedruckten Label steht ab
            # 2026-08-21c "GESAMTMENGE" statt "<N>x" - ohne diese Anpassung waere
            # beim EAN-Scan nicht mehr erkennbar, ob es sich um ein Set oder um
            # Einzelbestellungen handelt, und es wuerde trotz N Teilen im Set nur
            # 1x zur Verifikation verlangt. Gewichtsartikel bleiben hier bewusst
            # AUSSEN VOR (gewollt, siehe Chat: Gewichtsartikel laufen ueber die
            # Mehrfach-Scan-Sicherung/mengen_zuordnung.csv, nicht ueber EAN).
            fach = p.get("fach")
            eff = fach * qty if fach is not None and fach > 1 else qty
            anzahl = scan_bedarf(eff)
            zeilen.append([r["rnr"], ean, p["bez"], anzahl, _csv_datum(r)])
        # Rechnung aus dem aktuellen Lauf ERSETZT ihre alten Zeilen komplett
        # (auch mit einer leeren Liste, falls inzwischen keine EAN-Position
        # mehr vorhanden waere).
        bestand[r["rnr"]] = zeilen

    with open(csv_pfad, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f, delimiter=";", lineterminator="\n")
        w.writerow(["Rechnungsnummer", "EAN", "Bezeichnung", "Anzahl", "Rechnungsdatum"])
        for zeilen in bestand.values():
            for row in zeilen:
                w.writerow(row)


# ----------------------------------------------------------------------------
# wc_bestellnummern.csv  (Bruecke fuer wc_sendungsnummer_sync.py)
# ----------------------------------------------------------------------------
# Ordnet die Amicron-Rechnungsnummer der WooCommerce-Bestellnummer zu, damit
# ein separates Skript die Carrier-Sendungsnummern (DHL/DPD/Deutsche Post,
# jeweils per Rechnungsnummer referenziert) ueber die REST-API in den
# Onlineshop (Shiptastic) eintragen kann.
#
# Nur ONLINESHOP-Rechnungen kommen hinein. Erkennung: die Zeile unter
# "Rechnung Nr." (Feld r["kdnr"]) ist eine reine 4-6-stellige Ziffernfolge
# UND darunter steht eine Kundenmail (r["email"]), deren Domain NICHT von
# Amazon/eBay stammt. Amazon-Rechnungen tragen dort "nnn-nnnnnnn-nnnnnnn",
# Zaehltheken-Rechnungen eine Kd-Nr ohne Mail - beide fallen damit raus.
# Eine zufaellig passende Kd-Nr wird spaeter vom Sync-Skript abgefangen
# (die Nummer muss sich per API als echte Bestellung aufloesen lassen).
#
# Kumulativ wie post_zuordnung.csv (siehe _lies_bestehende_zeilen /
# BRUECKEN_CSV_MERGE_TAGE), aber mit deutlich laengerer Verfallsfrist: ein
# Carrier-Export kann eine Rechnung referenzieren, die vor Wochen gepickt und
# laengst archiviert wurde. Die Zeilen sind winzig, das faellt nicht ins
# Gewicht. Rechnungsdatum als LETZTE Spalte (Verfallslogik).
WC_BESTELLNR_CSV_MERGE_TAGE = 60

_MARKTPLATZ_MAIL = ("amazon.", "ebay.", "marketplace.")


def _ist_onlineshop_bestellung(kdnr, email):
    if not re.fullmatch(r"\d{4,6}", (kdnr or "").strip()):
        return False
    e = (email or "").strip().lower()
    if "@" not in e:
        return False
    return not any(marker in e for marker in _MARKTPLATZ_MAIL)


def schreibe_wc_bestellnummern_csv(rechnungen, csv_pfad):
    """wc_bestellnummern.csv - Rechnungsnummer -> WooCommerce-Bestellnummer.
    Format: Rechnungsnummer;WC_Bestellnummer;Email;Name;PLZ;Rechnungsdatum
    Rueckgabe: Anzahl der in DIESEM Lauf aufgenommenen Onlineshop-Rechnungen."""
    bestand = {}
    for t in _lies_bestehende_zeilen(csv_pfad, 2):
        datum = t[5] if len(t) > 5 else ""
        if _bruecken_datum_frisch(datum, WC_BESTELLNR_CSV_MERGE_TAGE):
            row = (t[:6] + ["", "", "", "", "", ""])[:6]
            bestand[t[0]] = row
    neu = 0
    for r in rechnungen:
        if not r["rnr"]:
            continue
        if not _ist_onlineshop_bestellung(r.get("kdnr"), r.get("email")):
            continue
        bestand[r["rnr"]] = [r["rnr"], (r.get("kdnr") or "").strip(),
                              (r.get("email") or "").strip(), r["name"],
                              r["plz"], _csv_datum(r)]
        neu += 1
    with open(csv_pfad, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f, delimiter=";", lineterminator="\n")
        w.writerow(["Rechnungsnummer", "WC_Bestellnummer", "Email", "Name",
                     "PLZ", "Rechnungsdatum"])
        for row in bestand.values():
            w.writerow(row)
    return neu


def archiviere(rechnungen, archiv_basis):
    """Verschiebt die AUFGENOMMENEN Rechnungs-PDFs nach
    <archiv_basis>/<JJJJ-MM-TT>/, damit sie nicht erneut auf eine Packliste
    geraten. Nur erfolgreich eingelesene Rechnungen werden verschoben -
    uebersprungene/fehlerhafte bleiben liegen. (verschoben, fehler, zielordner)."""
    ziel = os.path.join(archiv_basis, datetime.now().strftime("%Y-%m-%d"))
    os.makedirs(ziel, exist_ok=True)
    verschoben, fehler = 0, []
    for r in rechnungen:
        quelle = r.get("quelle")
        if not quelle or not os.path.exists(quelle):
            continue
        name = os.path.basename(quelle)
        zpfad = os.path.join(ziel, name)
        if os.path.exists(zpfad):           # Namenskollision vermeiden
            stamm, ext = os.path.splitext(name)
            i = 2
            while os.path.exists(os.path.join(ziel, f"{stamm}_{i}{ext}")):
                i += 1
            zpfad = os.path.join(ziel, f"{stamm}_{i}{ext}")
        try:
            shutil.move(quelle, zpfad)
            verschoben += 1
        except Exception as e:
            fehler.append(f"{name}: {e}")
    return verschoben, fehler, ziel


# ----------------------------------------------------------------------------
# MAIN
# ----------------------------------------------------------------------------

def main(argv):
    if len(argv) < 3:
        print("Aufruf: python packliste.py <rechnungs_ordner> "
              "\"<stammdaten_oder_leer>\" <ausgabe.pdf> [archiv_ordner]")
        return 2
    ordner = argv[0]
    _stammdaten = argv[1]  # optional, derzeit nicht fuer Sortierung benoetigt
    ausgabe = argv[2]
    archiv_ordner = argv[3].strip() if len(argv) >= 4 and argv[3].strip() else None

    pdfs = sorted(glob.glob(os.path.join(ordner, "*.pdf")))
    if not pdfs:
        print(f"Keine PDF-Rechnungen in {ordner} gefunden.")
        return 1

    rechnungen = []
    print(f"Lese {len(pdfs)} PDF(s) aus {ordner} ...")
    for p in pdfs:
        try:
            r = parse_pdf(p)
        except Exception as e:
            print(f"  FEHLER beim Lesen von {os.path.basename(p)}: {e}")
            continue
        if not r["positionen"]:
            print(f"  uebersprungen (keine Positionen erkannt): {os.path.basename(p)}")
            continue
        r["quelle"] = p          # Quellpfad fuer die spaetere Archivierung merken
        rechnungen.append(r)
        ok = r["zeilen_ok"] and r["summe_ok"] and r["vollstaendig_ok"]
        status = "OK" if ok else "WARNUNG"
        total_txt = f"{r['total']:.2f}" if r.get("total_lesbar", True) else "n.l."
        print(f"  {r['datei']}: Rg {r['rnr']}  {len(r['positionen'])} Pos  "
              f"Summe {r['summe']:.2f}/{total_txt}  [{status}]")
        if not r.get("total_lesbar", True):
            print(f"      Hinweis: Rechnungsbetrag nicht lesbar (lange Bezeichnung) "
                  f"-> Positionssumme {r['summe']:.2f} geprueft")
        if not r["vollstaendig_ok"]:
            print(f"      unvollstaendige Position(en): {', '.join(r['unvollstaendig'])}")

    if not rechnungen:
        print("Keine auswertbaren Rechnungen.")
        return 1

    os.makedirs(os.path.dirname(os.path.abspath(ausgabe)), exist_ok=True)
    out_dir = os.path.dirname(os.path.abspath(ausgabe))
    sammel_pfad = os.path.join(out_dir, "sammel_zuordnung.csv")

    gruppen_roh = finde_sammelgruppen(rechnungen)
    heute_str = datetime.now().strftime("%d.%m.%Y")
    for g in gruppen_roh:
        g["datum"] = heute_str
    # Codes kollisionsfrei mit einem eventuell frueheren, noch nicht
    # abgearbeiteten Lauf zusammenfuehren (siehe merge_sammelgruppen).
    gruppen, gruppen_fuer_csv = merge_sammelgruppen(gruppen_roh, sammel_pfad)
    baue_pdf(rechnungen, ausgabe, gruppen)

    csv_pfad = os.path.join(out_dir, "post_zuordnung.csv")
    schreibe_csv(rechnungen, csv_pfad)
    schreibe_sammel_csv(gruppen_fuer_csv, sammel_pfad)
    mengen_pfad = os.path.join(out_dir, "mengen_zuordnung.csv")
    schreibe_mengen_csv(rechnungen, mengen_pfad)
    ean_pfad = os.path.join(out_dir, "ean_zuordnung.csv")
    schreibe_ean_csv(rechnungen, ean_pfad)
    wc_pfad = os.path.join(out_dir, "wc_bestellnummern.csv")
    wc_neu = schreibe_wc_bestellnummern_csv(rechnungen, wc_pfad)

    warn = [r for r in rechnungen
            if not (r["zeilen_ok"] and r["summe_ok"] and r["vollstaendig_ok"])]
    print("-" * 60)
    print(f"PDF geschrieben:  {ausgabe}")
    print(f"CSV geschrieben:  {csv_pfad}")
    print(f"Mengen-CSV:       {mengen_pfad}")
    # EAN-CSV: Anzahl Zeilen (= Positionen mit EAN) zur schnellen Kontrolle melden
    _ean_pos = sum(1 for r in rechnungen for p in r["positionen"]
                   if (p.get("ean") or "").strip() and not ist_versand(p["art"], p["bez"]))
    print(f"EAN-CSV:          {ean_pfad}  ({_ean_pos} Position(en) mit EAN)")
    print(f"WC-Bestellnr-CSV: {wc_pfad}  ({wc_neu} Onlineshop-Rechnung(en) in diesem Lauf)")
    if gruppen:
        print(f"Sammeldruck-CSV:  {sammel_pfad}  ({len(gruppen)} Sammelcode(s): "
              + ", ".join(f"{g['code']}={g['art']}×{len(g['rnr'])}" for g in gruppen) + ")")
    else:
        print(f"Sammeldruck-CSV:  {sammel_pfad}  (keine Gruppe ab {SAMMEL_MIN} gleichen Einzelartikeln)")
    if warn:
        print(f"ACHTUNG: {len(warn)} Rechnung(en) mit Warnung (Betrag/Vollständigkeit): "
              + ", ".join(r["rnr"] for r in warn))
    else:
        print("Alle Rechnungen: Abgleich UND Vollständigkeit OK.")

    # Aufgenommene Rechnungen archivieren (nur wenn Archivordner uebergeben)
    if archiv_ordner:
        verschoben, fehler, ziel = archiviere(rechnungen, archiv_ordner)
        print(f"Archiviert:       {verschoben} Rechnung(en) -> {ziel}")
        if fehler:
            print(f"  NICHT verschoben ({len(fehler)}): " + "; ".join(fehler))
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
