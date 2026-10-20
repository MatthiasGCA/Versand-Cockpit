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
1. Ordner `C:\Rechnungsversand` anlegen, diese Dateien hineinkopieren: `rechnungsversand.py`,
   `Rechnungsversand_starten.bat`, `rechnungsversand_config.example.json`.
2. `rechnungsversand_config.example.json` kopieren nach **`rechnungsversand_config.json`** und anpassen:
   Absender/SMTP, **`ab_rechnungsnummer`** (= erste Rechnung, die das Programm bearbeiten soll – die nächste noch nicht
   per Amicron versendete), Betreff/Text. **Modus zuerst auf `"bericht"` lassen.**
3. SMTP-Passwort **nicht** in die Datei schreiben, sondern als Windows-Umgebungsvariable
   `RECHNUNGSVERSAND_SMTP_PASSWORT` setzen (Eingabeaufforderung: `setx RECHNUNGSVERSAND_SMTP_PASSWORT "…"`).
   Exchange Online: für das Postfach muss „Authentifizierter SMTP-Client“ erlaubt sein (ggf. App-Kennwort); Exchange vor
   Ort: Servername intern, Port 25/587, `sicherheit` = `none` oder `starttls`.
4. Test ohne Versand: `py rechnungsversand.py --trocken` → `C:\Rechnungsversand\Ausgabe\Bericht_<Datum>.csv`.
   Prüfen: Amazon-Rechnungen stehen als „NICHT gesendet (Amazon)“, Shop/eBay als „würde senden“.
5. Mehrere Tage Modus `"entwurf"`: es entstehen fertige `.eml`-Dateien (`Ausgabe\Entwuerfe\<Datum>`), die du in Outlook
   öffnen und ansehen kannst.
6. Erst danach Modus `"senden"`. **Wichtig:** Ab dann die Rechnungen für Onlineshop/eBay **nicht mehr zusätzlich**
   in Amicron per „Belege per Email versenden“ verschicken, sonst bekommen Kunden sie doppelt.
7. Aufgabenplanung (täglich 15:00), Eingabeaufforderung als Administrator:
   `schtasks /create /tn "Rechnungsversand" /tr "C:\Rechnungsversand\Rechnungsversand_starten.bat" /sc daily /st 15:00`
   Die Aufgabe muss unter dem Benutzer laufen, der Zugriff auf `Ausgangsrechnungen` hat und die Umgebungsvariable besitzt.

## Täglich
Nach 15:00 kurz `Ausgabe\Bericht_*.csv` / `rechnungsversand.log` ansehen. Zeilen „NICHT gesendet (unklar)“ bitte prüfen
und bei Bedarf von Hand in Amicron versenden. Rechnungen, die nach 15:00 entstehen, gehen am nächsten Tag raus
(Bereitstellung `max_alter_tage`, Standard 4 Tage).
