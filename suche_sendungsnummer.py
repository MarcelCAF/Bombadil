"""
Sucht eine DHL-Sendungsnummer in allen NAS-Archiven (Normal + Express).
Einfach starten, Nummer eingeben, Enter druecken.
"""
import os
import pandas as pd

NAS_NORMAL  = r"W:\Dokumentenaustausch\Tagesskripte\Bombadil\Archiv\DHL Normal"
NAS_EXPRESS = r"W:\Dokumentenaustausch\Tagesskripte\Bombadil\Archiv\DHL Express"

def suche_in_ordner(ordner, suchbegriff, typ):
    if not os.path.isdir(ordner):
        return []
    gefunden = []
    dateien = [f for f in os.listdir(ordner) if f.lower().endswith('.xlsx')]
    print(f"  Durchsuche {len(dateien)} Dateien ({typ}) ...")
    for f in dateien:
        try:
            xls = pd.read_excel(os.path.join(ordner, f), sheet_name=None)
            for sheet in xls.values():
                sheet.columns = [str(c).strip() for c in sheet.columns]
                bc_col   = next((c for c in sheet.columns if 'barcode' in c.lower()), None)
                scan_col = next((c for c in sheet.columns if 'scan' in c.lower() or 'date' in c.lower()), None)
                if bc_col:
                    treffer = sheet[sheet[bc_col].astype(str).str.contains(suchbegriff, na=False)]
                    if not treffer.empty:
                        for _, row in treffer.iterrows():
                            barcode = str(row[bc_col]).strip()
                            # Fuehrende 00 ergaenzen falls noetig
                            if not barcode.startswith('00') and len(barcode) <= 20:
                                barcode = '00' + barcode
                            scan_dt = str(row[scan_col]).strip() if scan_col else 'unbekannt'
                            gefunden.append({
                                'Typ':        typ,
                                'Datei':      f,
                                'Barcode':    barcode,
                                'Scan-Datum': scan_dt,
                            })
        except Exception:
            pass
    return gefunden


def main():
    print("=" * 50)
    print("  DHL Sendungsnummer-Suche")
    print("=" * 50)

    while True:
        eingabe = input("\nSendungsnummer eingeben (oder 'q' zum Beenden): ").strip()
        if eingabe.lower() == 'q':
            print("Tschuess!")
            break
        if not eingabe:
            continue

        # Fuehrende Nullen entfernen fuer robuste Suche
        suchbegriff = eingabe.lstrip('0')
        if not suchbegriff:
            print("Ungueltige Eingabe.")
            continue

        print(f"\nSuche nach: {eingabe} ...")
        alle_treffer = []
        alle_treffer += suche_in_ordner(NAS_NORMAL,  suchbegriff, "DHL Normal")
        alle_treffer += suche_in_ordner(NAS_EXPRESS, suchbegriff, "DHL Express")

        if not alle_treffer:
            print(f"\n  Nicht gefunden: {eingabe}")
        else:
            print(f"\n  {len(alle_treffer)} Treffer gefunden:")
            print(f"  {'─'*48}")
            for t in alle_treffer:
                print(f"  Barcode:    {t['Barcode']}")
                print(f"  Typ:        {t['Typ']}")
                print(f"  Scan-Datum: {t['Scan-Datum']}")
                print(f"  Datei:      {t['Datei']}")
                print(f"  {'─'*48}")

if __name__ == "__main__":
    main()
