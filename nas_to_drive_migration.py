"""
nas_to_drive_migration.py
=========================
Migriert NAS DHL-Archiv-Dateien einmalig auf Google Drive.

Was es tut:
  1. Liest alle .xlsx aus DHL Normal + DHL Express (NAS)
  2. Normalisiert Spalten (case-insensitiv) + konvertiert Excel-Serial-Datum
  3. Filtert ab Dezember 2025
  4. Dedupliziert per Barcode
  5. Gruppiert nach Monat (YYYY-MM)
  6. Prüft welche Monate auf Drive noch fehlen
  7. Lädt fehlende Monate hoch als DHL_Normal_Archiv_YYYY-MM.xlsx

Sicherheit:
  - DRY_RUN = True  → nur Vorschau, kein Upload
  - Überschreibt KEINE bestehenden Drive-Dateien
"""

import os, sys, io, pickle
from pathlib import Path
from datetime import date

import pandas as pd
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload, MediaIoBaseDownload

sys.stdout.reconfigure(encoding="utf-8")

# ── Konfiguration ─────────────────────────────────────────────────────────────

DRY_RUN = False   # ← auf False setzen um wirklich hochzuladen

CUTOFF  = date(2025, 12, 1)   # Nur Daten ab Dezember 2025

BOMBADIL_DIR     = Path(r"C:\Users\Abfuellung 15\Documents\Marcels Skripts\Bombadil")
NAS_NORMAL_DIR   = Path(r"W:\Dokumentenaustausch\Tagesskripte\Bombadil\Archiv\DHL Normal")
NAS_EXPRESS_DIR  = Path(r"W:\Dokumentenaustausch\Tagesskripte\Bombadil\Archiv\DHL Express")

GDRIVE_FOLDER_DHL_NORMAL  = "1G5_i9zUvjqVBCnE_ME-fPkhK4j5SPWpe"
GDRIVE_FOLDER_DHL_EXPRESS = "1J_oZcBmPECYfL7xj5U0oxgao--sdzS_c"

COL_BARCODE_NAMES = {"package barcode", "paket-barcode", "paket barcode", "paketbarcode", "barcode"}
COL_DATE_NAMES    = {"date of scan", "date of scan", "scandate", "scan date", "datum", "date"}


# ── Drive-Verbindung ──────────────────────────────────────────────────────────

def _get_drive_service():
    token_path = BOMBADIL_DIR / "token.json"
    with open(token_path, "rb") as f:
        creds = pickle.load(f)
    from google.auth.transport.requests import Request
    if creds.expired and creds.refresh_token:
        creds.refresh(Request())
        with open(token_path, "wb") as f:
            pickle.dump(creds, f)
    return build("drive", "v3", credentials=creds, cache_discovery=False)


# ── NAS-Dateien einlesen ──────────────────────────────────────────────────────

def _normalize_df(df: pd.DataFrame) -> pd.DataFrame | None:
    """Normalisiert Spalten auf 'Package Barcode' + 'Date of Scan', filtert ab CUTOFF."""
    df.columns = [str(c).strip() for c in df.columns]
    bc_col = date_col = None
    for col in df.columns:
        cl = col.lower()
        if cl in COL_BARCODE_NAMES:
            bc_col = col
        if cl in COL_DATE_NAMES:
            date_col = col
    if not bc_col or not date_col:
        return None

    df = df[[bc_col, date_col]].copy()
    df.columns = ["Package Barcode", "Date of Scan"]

    # Barcode bereinigen
    df["Package Barcode"] = df["Package Barcode"].astype(str).str.strip()
    df["Package Barcode"] = df["Package Barcode"].str.replace(r'^\s*=?"?\'?', "", regex=True).str.strip('"\'')
    df = df[df["Package Barcode"].str.len() > 5]

    # Datum konvertieren – Excel Serial oder echtes Datum
    col = df["Date of Scan"]
    if pd.api.types.is_numeric_dtype(col):
        df["Date of Scan"] = pd.to_datetime(col, unit="D", origin="1899-12-30", errors="coerce")
    else:
        df["Date of Scan"] = pd.to_datetime(col, errors="coerce", format="mixed")

    df = df.dropna(subset=["Date of Scan"])
    df = df[df["Date of Scan"].dt.date >= CUTOFF]
    return df if not df.empty else None


def load_nas_folder(folder: Path, label: str) -> pd.DataFrame:
    """Liest alle .xlsx aus einem NAS-Ordner, normalisiert und gibt einen DataFrame zurück."""
    print(f"\n📂  {label}: {folder}")
    frames = []
    skipped = []
    for xlsx in sorted(folder.glob("*.xlsx")):
        try:
            raw = pd.read_excel(xlsx, engine="openpyxl")
            df  = _normalize_df(raw)
            if df is None:
                skipped.append(xlsx.name)
                continue
            frames.append(df)
            print(f"   ✓  {xlsx.name:<45} {len(df):>6} Zeilen ab {CUTOFF}")
        except Exception as e:
            skipped.append(f"{xlsx.name} ({e})")

    if skipped:
        print(f"   ⏭  Übersprungen ({len(skipped)}): {', '.join(skipped[:5])}" +
              (" …" if len(skipped) > 5 else ""))

    if not frames:
        print("   ⚠  Keine verwertbaren Dateien gefunden.")
        return pd.DataFrame()

    combined = pd.concat(frames, ignore_index=True)
    before   = len(combined)
    combined = combined.drop_duplicates(subset=["Package Barcode"], keep="last")
    print(f"\n   Gesamt: {before} Zeilen → nach Dedup: {len(combined)}")
    return combined


# ── Drive: vorhandene Monate prüfen ──────────────────────────────────────────

def get_existing_months_on_drive(service, folder_id: str) -> set:
    """Gibt eine Menge von 'YYYY-MM' zurück, für die auf Drive bereits eine Datei existiert."""
    import re
    res = service.files().list(
        q=f"'{folder_id}' in parents and trashed=false",
        fields="files(name)", pageSize=200
    ).execute()
    months = set()
    for f in res.get("files", []):
        m = re.search(r"(\d{4}-\d{2})", f["name"])
        if m:
            months.add(m.group(1))
    return months


# ── Upload ────────────────────────────────────────────────────────────────────

def upload_month(service, df_month: pd.DataFrame, filename: str, folder_id: str):
    """Lädt einen Monats-DataFrame als Excel auf Drive hoch."""
    buf = io.BytesIO()
    df_month.to_excel(buf, index=False)
    buf.seek(0)
    media = MediaIoBaseUpload(buf,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    meta  = {"name": filename, "parents": [folder_id]}
    service.files().create(body=meta, media_body=media, fields="id").execute()


# ── Haupt-Logik ───────────────────────────────────────────────────────────────

def migrate(folder: Path, label: str, prefix: str, folder_id: str, service):
    df = load_nas_folder(folder, label)
    if df.empty:
        print(f"⚠  Keine Daten für {label}, überspringe.")
        return

    existing = get_existing_months_on_drive(service, folder_id)
    print(f"\n   Auf Drive bereits vorhanden: {sorted(existing) or '(keine)'}")

    # Nach Monat gruppieren
    df["_month"] = df["Date of Scan"].dt.to_period("M").astype(str)
    print(f"\n{'Monat':<12} {'Zeilen':>8}  {'Aktion'}")
    print("-" * 40)

    uploaded = skipped = 0
    for month, group in df.groupby("_month"):
        filename = f"{prefix}_Archiv_{month}.xlsx"
        if month in existing:
            print(f"  {month:<10} {len(group):>8}  ⏭  bereits auf Drive vorhanden")
            skipped += 1
            continue
        if DRY_RUN:
            print(f"  {month:<10} {len(group):>8}  🔵 DRY_RUN – würde hochladen als {filename}")
        else:
            upload_month(service, group[["Package Barcode", "Date of Scan"]], filename, folder_id)
            print(f"  {month:<10} {len(group):>8}  ✅ hochgeladen als {filename}")
        uploaded += 1

    print(f"\n  → {uploaded} Monate {'würden hochgeladen' if DRY_RUN else 'hochgeladen'}, "
          f"{skipped} bereits vorhanden.")


# ── Entry Point ───────────────────────────────────────────────────────────────

if __name__ == "__main__":
    print("=" * 60)
    print("NAS → Drive Migration")
    print(f"Cutoff: ab {CUTOFF}  |  DRY_RUN: {DRY_RUN}")
    print("=" * 60)

    if DRY_RUN:
        print("\n⚠  DRY_RUN ist aktiv – es wird NICHTS hochgeladen.")
        print("   Setze DRY_RUN = False um die Migration wirklich durchzuführen.\n")

    service = _get_drive_service()

    migrate(NAS_NORMAL_DIR,  "DHL Normal",  "DHL_Normal",  GDRIVE_FOLDER_DHL_NORMAL,  service)
    migrate(NAS_EXPRESS_DIR, "DHL Express", "DHL_Express", GDRIVE_FOLDER_DHL_EXPRESS, service)

    print("\n" + "=" * 60)
    print("Fertig!" if not DRY_RUN else "Trockenlauf abgeschlossen. Prüfe die Ausgabe oben.")
    print("=" * 60)
