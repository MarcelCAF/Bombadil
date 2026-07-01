"""
PDF-Sammler – täglich 21:00 Uhr via Windows Aufgabenplanung
Kopiert PDFs vom Vortag aus 5 Quell-Ordnern in den Ziel-Ordner.
Doppler (Dateiname schon vorhanden) → Unterordner "Doppler" mit Quell-Kürzel als Präfix.
"""

import os
import shutil
import datetime
import sys

# Windows-Konsole auf UTF-8 setzen damit Sonderzeichen funktionieren
if sys.stdout.encoding != "utf-8":
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

QUELLEN = [
    (r"W:\Automatisierungen\24_Auftragspipeline\data\taxierung_output", "Taxierung"),
]

ZIEL      = r"W:\HERSTELLUNG\Abfüllung\E-Sign Rezepte\unsigniert"
DOPPLER   = r"W:\HERSTELLUNG\Abfüllung\E-Sign Rezepte\Doppler"

def main():
    heute = datetime.date.today() - datetime.timedelta(days=1)  # Vortag
    os.makedirs(ZIEL,    exist_ok=True)
    os.makedirs(DOPPLER, exist_ok=True)

    kopiert    = 0
    doppler    = 0
    fehler     = 0
    kein_datum = 0

    for ordner, kuerzel in QUELLEN:
        if not os.path.isdir(ordner):
            print(f"[WARNUNG] Quell-Ordner nicht erreichbar: {ordner}")
            continue

        for dateiname in os.listdir(ordner):
            if not dateiname.lower().endswith(".pdf"):
                continue

            quelle_pfad = os.path.join(ordner, dateiname)

            try:
                stat = os.stat(quelle_pfad)
                aenderung = datetime.date.fromtimestamp(stat.st_mtime)
                erstellt  = datetime.date.fromtimestamp(stat.st_ctime)
            except OSError:
                # Ungultiger Zeitstempel (z.B. Datum im Jahr 30000) -> still uberspringen
                kein_datum += 1
                continue

            if aenderung != heute and erstellt != heute:
                continue  # nicht heute → überspringen

            ziel_pfad = os.path.join(ZIEL, dateiname)

            if not os.path.exists(ziel_pfad):
                try:
                    shutil.copy2(quelle_pfad, ziel_pfad)
                    print(f"[OK]      {kuerzel}: {dateiname}")
                    kopiert += 1
                except Exception as e:
                    print(f"[FEHLER]  {kuerzel}: {dateiname} – {e}")
                    fehler += 1
            else:
                # Datei schon im Ziel → Doppler-Ordner, Kürzel als Präfix
                doppler_name = f"{kuerzel}_{dateiname}"
                doppler_pfad = os.path.join(DOPPLER, doppler_name)
                # Falls auch im Doppler-Ordner schon vorhanden: Zeitstempel anhängen
                if os.path.exists(doppler_pfad):
                    ts = datetime.datetime.now().strftime("%H%M%S")
                    doppler_name = f"{kuerzel}_{ts}_{dateiname}"
                    doppler_pfad = os.path.join(DOPPLER, doppler_name)
                try:
                    shutil.copy2(quelle_pfad, doppler_pfad)
                    print(f"[DOPPLER] {kuerzel}: {dateiname} → Doppler/{doppler_name}")
                    doppler += 1
                except Exception as e:
                    print(f"[FEHLER]  {kuerzel}: {dateiname} – {e}")
                    fehler += 1

    print()
    print(f"=== Ergebnis Vortag ({heute}) ===")
    print(f"  Kopiert:         {kopiert}")
    print(f"  Doppler:         {doppler}")
    print(f"  Fehler (Kopie):  {fehler}")
    if kein_datum:
        print(f"  Ubersprungen (kein lesbares Datum): {kein_datum}")
    print()
    input("Fertig! Drücke Enter zum Schließen...")

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"\nUNERWARTETER FEHLER: {e}")
    input("Drücke Enter zum Schließen...")
