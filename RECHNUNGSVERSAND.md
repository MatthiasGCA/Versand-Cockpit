# Rechnungsversand (Onlineshop + eBay, nie Amazon)

Ein normales Programm auf dem Faktura-PC – **keine KI**. Es läuft täglich um 15:00 (Windows-Aufgabenplanung),
liest die Rechnungs-PDFs aus `Ausgangsrechnungen` und verschickt sie gesammelt per E-Mail – ausschließlich an
Onlineshop- und eBay-Kunden.

## Wie Amazon sicher ausgeschlossen wird
Die E-Mail-Adresse und der Fremdbeleg stehen auf der **Packliste** (Rechnungsnummer = Verbindung). Es wird nur
**positiv** erkannt (Prüfung an 1.741 echten Packlisten: Amazon 1.298 / eBay 366 / Shop 77, keine Mischfälle):

| Kanal | Erkennung (alles muss stimmen) | Aktion |
|---|---|---|
| **Onlineshop** | Fremdbeleg = genau 4–6 Ziffern **und** genau eine Kunden-Adresse (kein Marktplatz) | senden |
| **eBay** | genau eine Adresse `@members.ebay.com` **und** Fremdbeleg ist ein Name | senden |
| **Amazon** | irgendwo Amazon-Bestellnummer (`123-1234567-1234567`), das Wort „Amazon“ oder eine Marketplace-Adresse – in Packliste **oder** Rechnung | **nie senden** |
| alles andere | keine Packliste, keine/mehrere Adressen, Fremdbeleg weicht ab, … | **nicht senden**, im Bericht melden |

Weitere Sicherungen: Pflicht-Startnummer `ab_rechnungsnummer` (keine alten Rechnungen), Höchstzahl Mails je Lauf,
`nie_senden.txt` (eine Rechnungsnummer pro Zeile), Sendeprotokoll `gesendet.csv` (nie doppelt; ein Versuch ohne
Bestätigung wird **nie** blind wiederholt), Gutschriften/Storno/Duplikate/Auftragsbestätigungen werden ignoriert,
Rechnungs-PDF mit internen Angaben (Lagerort/Kennung) wird nicht gesendet, bei jedem Fehler wird nicht gesendet.

## Einrichten (Faktura-PC)
1. Die Dateien `rechnungsversand.py`, `rechnungsmail_vorlage.html`, `Rechnungsversand_starten.bat` und `rechnungsversand_config.example.json` in den
   **Programmordner des Carrier-Dashboards** kopieren (dort, wo `carrier_dashboard.py` liegt) – dann erkennt das
   Dashboard den Rechnungsversand automatisch. Zusätzlich die neue `carrier_dashboard.py` austauschen.
2. `rechnungsversand_config.example.json` kopieren nach **`rechnungsversand_config.json`** und anpassen:
   Absender/SMTP, **`ab_rechnungsnummer`** (= erste Rechnung, die das Programm bearbeiten soll – die nächste noch nicht
   per Amicron versendete), Betreff/Text. **Modus zuerst auf `"bericht"` lassen.** (`lauf_uhrzeit` ist nur die Anzeige-
   Uhrzeit im Dashboard – die echte Startzeit steht in der Aufgabenplanung, Schritt 7.)
3. SMTP-Passwort **nicht** in die Datei schreiben, sondern als Windows-Umgebungsvariable
   `RECHNUNGSVERSAND_SMTP_PASSWORT` setzen (Eingabeaufforderung: `setx RECHNUNGSVERSAND_SMTP_PASSWORT "…"`).
   Postfach (DomainFactory): `onlinerechnung@gasecenter-onlineshop.de`, SMTP `sslout.de`, Port 465, `sicherheit` = `ssl`
   (in der Vorlage schon eingetragen; Verbindung am 26.10.2026 geprüft: Server erreichbar, Anmeldung per PLAIN/LOGIN).
4. Test ohne Versand: im Dashboard auf **„Rechnungsversand“ → „Prüfen (Bericht)“** klicken (oder
   `py rechnungsversand.py --trocken`). Prüfen: Amazon-Rechnungen stehen als „NICHT gesendet (Amazon)“, Shop/eBay als
   „würde senden“.
5. Mehrere Tage Modus `"entwurf"` bzw. im Fenster **„Entwürfe erstellen“**: es entstehen fertige `.eml`-Dateien
   (`Ausgabe\Entwuerfe\<Datum>`), die du in Outlook öffnen und ansehen kannst.
6. Erst danach in der Konfiguration Modus `"senden"`. **Wichtig:** Ab dann die Rechnungen für Onlineshop/eBay **nicht mehr
   zusätzlich** in Amicron per „Belege per Email versenden“ verschicken, sonst bekommen Kunden sie doppelt.
7. Aufgabenplanung (täglich 15:00), Eingabeaufforderung als Administrator (Pfad anpassen):
   `schtasks /create /tn "Rechnungsversand" /tr "C:\Carrier-Dashboard\Rechnungsversand_starten.bat" /sc daily /st 15:00`
   Die Aufgabe muss unter dem Benutzer laufen, der Zugriff auf `Ausgangsrechnungen` hat und die Umgebungsvariable besitzt.

## E-Mail-Text und Signatur
- **Kontrollkopie:** In der Anfangszeit geht jede Rechnungsmail zusätzlich als Blindkopie (BCC) an
  `Dokumente@Gasecenter.info` (`"bcc"` in der Konfiguration; der Kunde sieht das nicht). Später einfach leeren: `"bcc": ""`.
- Die Mail geht als **HTML mit Firmensignatur und Logo** raus (Vorlage `rechnungsmail_vorlage.html`, liegt neben dem
  Programm) plus einem Klartext-Teil (`text` in der Konfiguration) für Mailprogramme ohne HTML.
- Platzhalter in Betreff, Text und Vorlage: `{rnr}` (Rechnungsnummer), `{rechnungsdatum}`, `{bestellung}` – wird je nach
  Kanal zu „Ihre Bestellung Nr. 25623 in unserem Onlineshop“ bzw. „Ihre Bestellung bei eBay“.
- Text ändern: `rechnungsmail_vorlage.html` mit einem Editor öffnen (nicht Word), nur die Sätze oben anpassen, Platzhalter
  stehen lassen. Fehlt die Vorlage, wird **keine** Mail gebaut (Meldung im Bericht).
- Ansicht: `Beispiel_Rechnungsmail_Shop.html` im Browser bzw. `Beispiel_Rechnungsmail_eBay.eml` in Outlook öffnen.

## Im Carrier-Dashboard
- **Statuszeile oben rechts** („Rechnungsversand 15:00: 38 gesendet, 2 unklar“): grün = alles gut, gelb = unklare
  Rechnungen oder heute noch kein Lauf, rot = Fehler. Klick öffnet das Fenster.
- **Fenster „Rechnungsversand“:** Liste des letzten Laufs (Fehler/Unklares oben), Knöpfe **Prüfen (Bericht)**,
  **Entwürfe erstellen**, **Jetzt senden …** (nur wenn `modus` = `"senden"`; es folgt erst ein Bericht und eine Rückfrage,
  Standard Nein), **Bericht öffnen**.
- Hat der geplante Lauf Fehler oder unklare Rechnungen, meldet sich das Dashboard (falls geöffnet) einmal kurz per Dialog.
- Läuft gerade ein anderer Lauf (z. B. die geplante Aufgabe), wird ein zweiter nicht gestartet (Sperre).

## Täglich
Nach 15:00 kurz `Ausgabe\Bericht_*.csv` / `rechnungsversand.log` ansehen. Zeilen „NICHT gesendet (unklar)“ bitte prüfen
und bei Bedarf von Hand in Amicron versenden. Rechnungen, die nach 15:00 entstehen, gehen am nächsten Tag raus
(Bereitstellung `max_alter_tage`, Standard 4 Tage).
