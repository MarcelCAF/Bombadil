# Temporäres Analyse-Skript: Mai-Vergleich CSV vs. Bombadil-Daten (nur lesen)
import csv
import datetime
import sys

sys.path.insert(0, ".")
import pandas as pd
from Bombadil import (load_dhl_nas_archive, fetch_dhl_archiv_gdrive,
                      _merge_live_archiv, DHL_NAS_NORMAL_DIR,
                      DHL_NAS_EXPRESS_DIR, ORCA_COL_SCAN, first_existing)

# ── 1. CSV: Mai-Tageswerte nach Kategorie (CAF / GB / Gesamt) ──────────────
path = r"C:\Users\Abfuellung 15\Downloads\CAF Priority List(Report Versand).csv"
rows = list(csv.reader(open(path, encoding="latin-1"), delimiter=";"))
dates = rows[0][2:]

csv_caf, csv_gb = {}, {}
for r in rows[1:]:
    kat = (r[1] or "").strip().upper() if len(r) > 1 else ""
    # Kategorie auch aus dem Namen ableiten (manche Zeilen haben leere Spalte 2)
    if not kat:
        kat = "CAF" if "CAF" in r[0].upper() else ("GB" if "GB" in r[0].upper() else "")
    ziel = csv_caf if kat == "CAF" else (csv_gb if kat == "GB" else None)
    if ziel is None:
        continue
    for d, v in zip(dates, r[2:]):
        v = v.strip()
        if v.isdigit():
            try:
                dt = datetime.datetime.strptime(d, "%d.%m.%Y").date()
            except ValueError:
                continue
            if dt.year == 2026 and dt.month == 5:
                ziel[dt] = ziel.get(dt, 0) + int(v)

print("CSV Mai-Summen:  CAF =", sum(csv_caf.values()),
      " GB =", sum(csv_gb.values()),
      " Gesamt =", sum(csv_caf.values()) + sum(csv_gb.values()))

# ── 2. Bombadil: echte Tageswerte Mai (NAS + Drive, OHNE Phantom-Korrektur) ──
def daily_counts(nas_dir):
    nas = load_dhl_nas_archive(nas_dir)
    return nas

try:
    drive_n, drive_e = fetch_dhl_archiv_gdrive()
except Exception as e:
    print("Drive nicht erreichbar:", e)
    drive_n, drive_e = pd.DataFrame(), pd.DataFrame()

for name, nas_dir, drive_df in [("Normal", DHL_NAS_NORMAL_DIR, drive_n),
                                 ("Express", DHL_NAS_EXPRESS_DIR, drive_e)]:
    nas_df = load_dhl_nas_archive(nas_dir)
    df = _merge_live_archiv(drive_df, nas_df)
    if df.empty:
        print(name, ": keine Daten")
        continue
    c = first_existing(df, ORCA_COL_SCAN)
    d = pd.to_datetime(df[c], errors="coerce", utc=True).dt.tz_convert(None).dt.date.dropna()
    mai = d[(pd.Series(list(d)).apply(lambda x: x.year == 2026 and x.month == 5)).values] \
        if len(d) else d
    s = pd.Series(list(d))
    s = s[s.apply(lambda x: x.year == 2026 and x.month == 5)]
    vc = s.value_counts().sort_index()
    print()
    print(f"=== Bombadil {name} – Mai real (Summe {int(vc.sum())}):")
    for dt, n in vc.items():
        wd = ["Mo", "Di", "Mi", "Do", "Fr", "Sa", "So"][dt.weekday()]
        caf = csv_caf.get(dt, 0)
        gb  = csv_gb.get(dt, 0)
        print(f"  {dt} {wd}  real={int(n):5d}   CSV: CAF={caf:5d} GB={gb:5d} ges={caf+gb:5d}")
