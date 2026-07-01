"""Testet die _gimli_search-Kernlogik mit echten NAS-Daten."""
import sys, os
sys.stdout.reconfigure(encoding='utf-8')
import pandas as pd

ORCA_COL_BARCODE = ["Package Barcode", "Paket-Barcode", "Paket Barcode", "Paketbarcode", "Barcode", "barcode"]
ORCA_COL_SCAN    = ["Date of Scan", "Date Of Scan", "DateOfScan", "Scan Date", "ScanDate", "date"]

def first_existing(df, candidates):
    for c in candidates:
        if c in df.columns:
            return c
    return None

def clean_barcode(value):
    if pd.isna(value): return ""
    s = str(value).strip()
    if s.startswith('="') and s.endswith('"') and len(s) >= 4: s = s[2:-1].strip()
    if s.startswith("'"): s = s[1:].lstrip()
    s = s.strip()
    if s.endswith(".0") and s[:-2].lstrip("-").isdigit(): s = s[:-2]
    return s

# NAS Normal laden (eine Datei reicht fuer den Test)
df = pd.read_excel(r'W:\Dokumentenaustausch\Tagesskripte\Bombadil\Archiv\DHL Normal\260430.xlsx')
df.columns = [str(c).strip() for c in df.columns]

# === exakt die Gimli-Logik ===
eingabe = "00340433888321263486"
such = clean_barcode(eingabe).lstrip("0")

bc_col = first_existing(df, ORCA_COL_BARCODE)
sc_col = first_existing(df, ORCA_COL_SCAN)
print(f"Suchbegriff (ohne 0): {such}")
print(f"Barcode-Spalte: {bc_col} | Scan-Spalte: {sc_col}")

norm = df[bc_col].map(clean_barcode).str.lstrip("0")
mask = norm.str.contains(such, na=False, regex=False)

treffer = []
for idx in df.index[mask]:
    roh = clean_barcode(df.at[idx, bc_col])
    anzeige = "00" + roh if (roh.isdigit() and not roh.startswith("00")) else roh
    scan = str(df.at[idx, sc_col]) if sc_col else ""
    treffer.append((anzeige, "DHL Normal", scan))

print(f"\n{len(treffer)} Treffer:")
for a, t, s in treffer:
    print(f"  Barcode: {a} | Typ: {t} | Scan: {s}")

# Gegentest: nicht existierende Nummer
such2 = "99999999999999999"
mask2 = norm.str.contains(such2, na=False, regex=False)
print(f"\nGegentest (Fantasie-Nummer): {int(mask2.sum())} Treffer (erwartet 0)")
