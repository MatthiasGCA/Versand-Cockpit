# Versandstatistik online – Reiter „Top-Artikel“ (vorbereitet, noch nicht veröffentlicht)

Stand: 2026-10-08. Nichts davon ist auf dem Webspace – die vorhandene Seite läuft unverändert weiter,
bis die Dateien unten bewusst ausgetauscht werden.

## Bestandteile

| Datei | Wohin | Zweck |
|---|---|---|
| `index.html` | Webspace, Ordner `versandstatistik/` (ersetzt die alte) | Seite mit neuem 4. Reiter „Top-Artikel“ |
| `vs-ingest.php` | Webspace, Wurzelverzeichnis (ersetzt die alte) | nimmt zusätzlich `artikel.json` entgegen (`?ziel=artikel`, bis 12 MB); bisheriges Verhalten für `data.json` bleibt gleich |
| `artikel_upload.py` | Lager-PC (DESKTOP-N2H75H), neben `statistik_upload.py` | verdichtet `artikel_verkaeufe.csv` (+ `artikel_historie.csv`) zu `artikel.json` und lädt sie hoch |

`vs-ingest.php.vorlage` im Repo hat KEIN Geheimwort. Die fertige Datei `vs-ingest.php` (mit Geheimwort) liegt nur
lokal/im Austauschpaket – nicht weitergeben.

## Veröffentlichen (wenn du so weit bist)

1. **Sicherung** der beiden alten Webdateien (`index.html`, `vs-ingest.php`) von Webspace herunterladen.
2. `vs-ingest.php` und `index.html` hochladen (gleiche Orte wie die alten).
3. Lager-PC: `artikel_upload.py` neben `statistik_upload.py` legen. Test ohne Upload:
   `py artikel_upload.py --dry-run --out C:\Temp\artikel.json`
4. Erster Upload: `py artikel_upload.py --force` → „Upload OK“. Das Token wird automatisch aus
   `statistik_upload.py` gelesen (oder `artikel_upload_config.json` mit `"token"`).
5. Aufgabenplanung: `py artikel_upload.py` z. B. alle 30 Minuten (lädt nur bei Änderung hoch).
6. Seite aufrufen → Reiter **Top-Artikel**.

## Was der Reiter kann

- frei wählbarer Zeitraum (Von/Bis) plus Schnellwahl (30/90 Tage, dieses/letztes Jahr, alles)
- Rangfolge nach Bestellungen, Stückzahl oder Umsatz; Top 20 / 50 / 100
- Kennzahlen: verkaufte Stück, Artikelumsatz, verschiedene Artikel – jeweils mit Vergleich zum gleichen
  Zeitraum des Vorjahres; je Artikel Vorjahres-Stückzahl und Veränderung in %
- Klick auf einen Artikel: Verlauf der verkauften Stück je Monat, Jahre übereinander (Saisonbild)

## Datenquellen und Grenzen

- Tagesgenau: alle Rechnungen, die das Carrier-Dashboard (Schritt 2) verarbeitet hat, ab 16.07.2026
  (Beginn des Archivs). Menge = wie auf der Rechnung (ein Karton „GermanFire 24 Stück“ zählt 1).
- Frühere Zeiträume (2025 bis Juni 2026): nur monatsgenau aus dem Amicron-Export
  (`artikel_historie.csv`) – diese Monate zählen im Zeitraum nur, wenn sie komplett darin liegen.
- Die Seite ist passwortgeschützt (Basic-Auth wie bisher); es werden keine Kundendaten übertragen,
  nur Artikelnummer, Bezeichnung, Mengen und Preise je Tag.
