# -*- coding: utf-8 -*-
r"""
label_erkennen.py  --  Ist eine PDF-Datei ein DHL-/DPD-/Deutsche-Post-Versandlabel?

Nutzt DIESELBEN Erkennungsmuster wie scan_druck.py (VERSENDER_PROFILE /
erkenne_profil), damit PaketImport-Mover.ps1 nur echte Versandlabel nach
\\DESKTOP-N2H75H\Netzwerk\Paketscheine kopiert - und NICHT jedes x-beliebige
PDF, das sonst noch im Downloads-Ordner landet.

Aufruf:
    py label_erkennen.py "C:\Users\Gasecenter\Downloads\irgendeine.pdf"

Exit-Code 0 + "JA <Versender>"  auf stdout  -> ist ein Versandlabel
Exit-Code 1 + "NEIN"                        -> kein erkanntes Label
Exit-Code 2 + "FEHLER: ..."                  -> Datei konnte nicht gelesen werden
                                                (z.B. kein PDF, beschaedigt,
                                                noch vom Browser gesperrt)

WICHTIG: Wenn scan_druck.py die Erkennungsmuster (VERSENDER_PROFILE) einmal
aendert, bitte HIER SPIEGELBILDLICH nachziehen, sonst laufen beide
auseinander.
"""

import re
import sys

# --- Muster 1:1 aus scan_druck.py (VERSENDER_PROFILE / erkenne_profil) -------
ERKENNUNG = [
    ("DPD", r"DPD|PARCELLetter|Parcel"),
    ("DHL", r"DHL\s*PAKET|PAKET INTERNATIONAL|Kostenstelle:"),
    ("Deutsche Post", r"NEERGOG|GOGREEN|tsoP ehcstueD|Deutsche Post"),
]


def _ist_internetmarke(text):
    """Deutsche-Post-Internetmarke: Text ist zeichenweise gespiegelt, kein
    'Deutsche Post'-Wortlaut vorhanden (nur Logo als Grafik). Erkennung ueber
    die gespiegelte IM-Frankaturzeile ('IM <TT.MM.JJ> <Betrag>')."""
    rev = "\n".join(" ".join(w[::-1] for w in ln.split())
                    for ln in (text or "").splitlines())
    return bool(re.search(r"\bIM\b", rev)) and \
           bool(re.search(r"\d{2}\.\d{2}\.\d{2}", rev))


def erkenne(text):
    for name, muster in ERKENNUNG:
        if re.search(muster, text):
            return name
    if _ist_internetmarke(text):
        return "Deutsche Post (Internetmarke)"
    return None


def main(argv):
    if len(argv) != 1:
        print("FEHLER: Aufruf mit genau einem PDF-Pfad erwartet")
        return 2
    pfad = argv[0]
    try:
        import pdfplumber
        with pdfplumber.open(pfad) as pdf:
            text = "\n".join((p.extract_text() or "") for p in pdf.pages)
    except Exception as e:
        print(f"FEHLER: {e}")
        return 2

    treffer = erkenne(text)
    if treffer:
        print(f"JA {treffer}")
        return 0
    print("NEIN")
    return 1


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
