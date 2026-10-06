# Carrier-Dashboard – Inbetriebnahme auf dem Faktura-PC

Stand: 2026-10-06 · Dashboard 2026-10-06a · Regeln 2026-10-05e

## Prinzip

Das Dashboard ersetzt `Pickliste_erstellen.bat` und nutzt **dieselben Ordner**:

| Zweck | Ordner (Faktura-PC) |
|---|---|
| Pool (Amicron legt hier die Rechnungs-PDFs ab) | `C:\Packlisten` |
| Pickliste + Brücken-CSVs (lokal) | `C:\Packlisten\Pickliste` |
| Archiv | `C:\Packlisten\Archiv` |
| Carrier-CSVs (DHL/DPD/Post zum Hochladen) | `C:\Carrier_Export` |
| Brücken-CSVs für `scan_druck.py` | `\\DESKTOP-N2H75H\Netzwerk\Paketscheine` |
| WooCommerce-Sendungsnummer-Sync | `C:\Scripts\Sendungsnummern_WC` |

Weil Ausgabe- und Archivordner identisch zur `.bat` sind, laufen die **SAM-Nummern lückenlos weiter**
(`sammel_zuordnung.csv` liegt schon in `C:\Packlisten\Pickliste`) und die `.bat` bleibt als
**Rückfallebene** nutzbar. Nie beide für dieselben PDFs nacheinander laufen lassen
(wer zuerst läuft, archiviert die Rechnungen).

## 1. Voraussetzungen auf dem Faktura-PC

- Python (`py`) ist schon da (die `.bat` nutzt es). Pakete prüfen/nachinstallieren:
  `py -m pip install pdfplumber reportlab pypdf pillow`
- Ordner `C:\Carrier_Export` anlegen (falls nicht vorhanden).

## 2. Dateien kopieren

**Dashboard-Ordner `C:\Carrier-Dashboard`** (neu anlegen) – aus `Faktura-PC-Paket\Carrier-Dashboard\`:

`carrier_dashboard.py`, `carrier_regeln.py`, `carrier_export.py`, `carrier_statistik.py`,
`packliste.py`, `dpd_logo.png`, `Carrier-Dashboard.ico`, `Carrier-Dashboard_starten.vbs`,
**`carrier_dashboard_config.json`** (legt die Ordner oben fest – ohne diese Datei läuft das Dashboard
mit den Laptop-Testordnern!)

**Rückfall-Ordner `C:\Packlisten`** – aus `Faktura-PC-Paket\Packlisten\`:
neue `packliste.py` + `dpd_logo.png` (damit auch die `.bat` die neuen Regeln/Lagerorte/DPD-Logo hat).

**Scanner-/Drucker-PC (dort läuft `scan_druck.py`)** – aus `Faktura-PC-Paket\Scanner-PC\`:
neue `scan_druck.py` (Lagerort auf dem Sammeldeckblatt). Altes Programm vorher beenden.

Desktop-Verknüpfung auf `Carrier-Dashboard_starten.vbs`, ggf. „Symbol ändern“ → `Carrier-Dashboard.ico`.

## 3. Vor dem ersten Lauf (Sicherung)

- `\\DESKTOP-N2H75H\Netzwerk\Paketscheine\*.csv` einmal kopieren (Sicherung der Brücken-CSVs und
  Statistik, v. a. `carrier_statistik.csv`, `artikel_statistik.csv`, `statistik.csv`).
- Alte Pickliste/Rechnungen des Tages sind unkritisch – das Dashboard verschiebt nur, wie die `.bat`.

## 4. Erster Lauf (mit kleiner Menge, z. B. 5–10 Rechnungen)

1. Dashboard starten → Kopfzeile muss zeigen: Pool-Ordner `C:\Packlisten` und
   „Brücken-CSVs für scan_druck → \\DESKTOP-N2H75H\…\Paketscheine“. Fehlt die zweite Zeile, wurde die
   Konfigurationsdatei nicht gefunden/gelesen (beim Start erscheint dann eine Warnung).
2. Amicron-Export auslösen → PDFs erscheinen automatisch in der Tabelle (Auto-Einlesen, ca. 5 s Ruhe).
3. Zuordnung prüfen (Hinweise quittieren, Adresse/Gewicht bei Bedarf bearbeiten) → **Schritt 2**.
4. Fertig-Dialog muss melden: „Brücken-CSVs nach … kopiert (4 Dateien)“.
5. Prüfen:
   - `Paketscheine\sammel_zuordnung.csv`: neue SAM-Nummern schließen **fortlaufend** an die bisherigen an.
   - Einen Sammelbarcode und einen Einzelbarcode der neuen Pickliste mit `scan_druck.py` scannen → Label/Druck OK.
   - Carrier-CSV aus `C:\Carrier_Export` im DHL-/DPD-/Post-Portal importieren (wie bisher getestet).
   - `wc_bestellnummern_dashboard.csv` liegt in `C:\Scripts\Sendungsnummern_WC`.

Kommt im Fertig-Dialog eine **WARNUNG zu den Brücken-CSVs** (Netzfreigabe weg): Schritt 2 **nicht**
wiederholen, sondern die vier Dateien `post_/sammel_/mengen_/ean_zuordnung.csv` aus
`C:\Packlisten\Pickliste` von Hand nach `Paketscheine` kopieren.

## 5. Schutzfunktionen im Dauerbetrieb

- Pickliste wird nach Schritt 2 automatisch geöffnet und gedruckt (Konfiguration `pickliste_oeffnen` / `pickliste_drucken`, Windows-Standarddrucker; schlägt das Drucken fehl, steht im Fertig-Dialog eine Warnung).
- Gutschriften (111…) und Auftragsbestätigungen (444…) werden abgelehnt (`RNR_GESPERRT` in `packliste.py`).
- Kopien einer doppelten Rechnungsnummer im Pool sind Fehler und werden nicht verarbeitet.
- Eine Rechnungsnummer wird höchstens einmal in Carrier-CSVs geschrieben (auch über Läufe hinweg).
- Brief/Großbrief an eine Packstation wird automatisch DHL (Hinweis „auf Kleinpaket abändern“).

## 6. Später

- Automatische Archiv-Bereinigung (Aufbewahrungsdauer + Ordnerliste mit Matthias abstimmen; nie löschen:
  `carrier_statistik.csv`, `artikel_statistik.csv`, `statistik.csv`, `sammel_zuordnung.csv`).
- Wenn das Dashboard stabil läuft: `Pickliste_erstellen.bat` weglegen (nicht löschen, Rückfallebene).
