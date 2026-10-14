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

---

# Reiter „Live“ – Spiegel des Versandcockpits (2026-10-14)

Zeigt auf der Versandstatistik-Seite, wie weit das Lager gerade ist: je Versender gedruckt / offen, Countdown zur
Abhol-Deadline, die offenen Rechnungsnummern (mit Scan-/EAN-Fortschritt), Warnungen und oben ein Statusbalken –
**„Alles gedruckt – nächster Rechnungslauf kann gestartet werden“** wenn nichts mehr offen ist. Aktualisiert sich alle
10 Sekunden. Direktlink/Lesezeichen: `https://gasecenter-onlineshop.de/versandstatistik/#live`.

## Weg der Daten

```
scan_druck.py (Scanner-PC)  →  Paketscheine\cockpit_live.json  →  live_upload.py (Lager-PC)  →  vs-ingest.php?ziel=live  →  versandstatistik/live.json  →  Reiter „Live“
```

- `scan_druck.py` schreibt die kleine JSON-Datei bei jeder Änderung (höchstens alle 5 s) und als Lebenszeichen alle 30 s,
  in einem Hintergrund-Thread. Fehler dabei stören das Cockpit nie.
- `live_upload.py` läuft im Dauerlauf auf dem Lager-PC (dort liegt das Upload-Token aus `statistik_upload.py`; die
  Scanner-PCs brauchen es nicht) und lädt die Datei hoch, sobald sie sich ändert.
- Läuft das Cockpit nicht (oder der Upload), erscheint auf der Seite nach 2 Minuten rot „Cockpit meldet sich nicht“
  mit der Uhrzeit des letzten Stands.

## Datenschutz

Standardmäßig werden **nur Rechnungsnummern** übertragen, keine Empfängernamen. Wer die Namen der offenen Aufträge
auf der Seite sehen will: in `scan_druck.py` `LIVE_MIT_NAMEN = True` setzen (Kundendaten gehen dann auf den
Webspace – Seite ist passwortgeschützt, trotzdem bewusst entscheiden).

## Einrichten

1. **Webspace:** `vs-ingest.php` (mit Geheimwort) und `index.html` austauschen (alte vorher sichern). Neu in der PHP:
   Ziel `live` → `versandstatistik/live.json`, Server-Empfangszeit wird mit eingetragen.
2. **Scanner-PC:** neue `scan_druck.py` (Version ab `2026-10-14a`) einsetzen, Cockpit neu starten. In
   `LIVE_STATUS_DATEI` steht der Pfad (Standard: `Paketscheine\cockpit_live.json`).
3. **Lager-PC:** `live_upload.py` + `live_upload.vbs` neben `statistik_upload.py` legen (z. B. `C:\Packlisten`).
   Test: `py live_upload.py --once` → „Upload OK“. Dauerbetrieb: `live_upload.vbs` in den Autostart-Ordner legen
   (`Win+R` → `shell:startup`) und einmal doppelklicken. Log: `C:\Packlisten\live_upload.log`.
4. Seite öffnen → Reiter **Live**.
