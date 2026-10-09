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

## Was „Stück“ bedeutet (Regel von Matthias, 2026-10-09)

Die Auswertung zählt die **tatsächlich verkauften Einzelteile laut Stückliste** (Amicron-Export):
- **Mehrfachpackung** (Set aus N × einem Einzelartikel, z. B. „GermanFire 6 Stück“ = 6 × BP-GF): es zählt der
  Einzelartikel (BP-GF, 6 Stück); der Umsatz der Packung wird dem Einzelartikel zugeschlagen.
- **Zusammengesetzter Artikel** (z. B. Paella-Schlauch, Bunsenbrenner + Kartuschen): es zählt der Artikel selbst, die
  Bestandteile fallen weg.
- normale Artikel: wie auf der Rechnung. Es wird nichts mit Faktoren umgerechnet.
Quelle ist `artikel_amicron.csv` (tagesgenau, per `artikel_import_amicron.py` aus dem Amicron-Export). Rechnungen, die
dort noch fehlen, ergänzt `artikel_upload.py` aus den Dashboard-Daten (Packungen werden dann mit
`artikel_stueckfaktor.csv` auf den Einzelartikel umgerechnet).

## Datenquellen und Grenzen

- Tagesgenau (letzte ~14 Monate), ältere Tage werden zu Monaten verdichtet; Monate zählen im Zeitraum nur, wenn sie
  komplett darin liegen.
- Aktuell bleibt die Auswertung nur so frisch wie der letzte Amicron-Export; dazwischen springen die Dashboard-Daten ein.
- Die Seite ist passwortgeschützt (Basic-Auth wie bisher); es werden keine Kundendaten übertragen,
  nur Artikelnummer, Bezeichnung, Mengen und Preise je Tag.
