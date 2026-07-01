"""
build_mai_express_quelle.py
===========================
Baut aus ALLEN verfügbaren echten Quellen eine vollständige
DHL_Express_Archiv_2026-05.xlsx für den Mai (nach dem Datencrash):

  Quellen (alle Union, dedupliziert per Barcode):
    1. Gimli: PDF-Versandlabel aus dem Datensicherungs-Ordner (Mai 2026, Express)
    2. NAS Express-Archiv (Mai)
    3. Drive: fehlplatzierte Datei + PDF-Teilarchiv (Mai)

Ziel laut Priority List (CSV): ~14.084 Express im Mai.
Das Skript zeigt wie nah die echten Daten kommen → dann entscheiden wir,
ob die Phantom-Korrektur (DHL_KORREKTUR_EXPRESS) reduziert/entfernt werden kann.

DRY_RUN = True → nur zählen, nichts hochladen.
"""

import os, re, sys, io, pickle
from datetime import datetime, date
from concurrent.futures import ProcessPoolExecutor, as_completed

import pandas as pd

sys.stdout.reconfigure(encoding="utf-8")

DRY_RUN = True

PDF_FOLDER   = r"W:\Dokumentenaustausch\Tagesskripte\Datensicherung"
NAS_EXPRESS  = r"W:\Dokumentenaustausch\Tagesskripte\Bombadil\Archiv\DHL Express"
GDRIVE_FOLDER_DHL_EXPRESS = "1J_oZcBmPECYfL7xj5U0oxgao--sdzS_c"
GDRIVE_FOLDER_ABHOLER     = "1OSnMmPf--uqt4ulDGy3ILT61pUuCPzpn"

PAT_EXPRESS = re.compile(r'\(J\)\s+(JD\d{2}\s+[\d\s]{15,})')


# ── 1. Gimli: PDF-Express-Barcodes für Mai ────────────────────────────────────

def process_pdf(path):
    """Liest ein PDF, gibt (barcode, scan_date) zurück wenn Express + Mai 2026."""
    try:
        mtime   = os.path.getmtime(path)
        scan_dt = datetime.fromtimestamp(mtime)
        if not (scan_dt.year == 2026 and scan_dt.month == 5):
            return None
        import pdfplumber
        with pdfplumber.open(path) as pdf:
            text = "".join(p.extract_text() or "" for p in pdf.pages)
        m = PAT_EXPRESS.search(text)
        if m:
            bc = "J" + m.group(1).replace(" ", "")
            return (bc, scan_dt)
    except Exception:
        pass
    return None


def load_pdf_express_mai():
    files = [os.path.join(PDF_FOLDER, f) for f in os.listdir(PDF_FOLDER)
             if f.lower().endswith(".pdf")]
    print(f"Lese {len(files)} PDFs (filtere Express + Mai) …")
    rows = []
    done = 0
    with ProcessPoolExecutor(max_workers=8) as ex:
        futures = [ex.submit(process_pdf, p) for p in files]
        for fut in as_completed(futures):
            r = fut.result()
            done += 1
            if r:
                rows.append({"Package Barcode": r[0], "Date of Scan": r[1]})
            if done % 3000 == 0:
                print(f"  {done}/{len(files)} …")
    df = pd.DataFrame(rows)
    print(f"PDF Express Mai: {len(df)} Barcodes")
    return df


# ── 2+3. NAS + Drive (wie im Cleanup) ─────────────────────────────────────────

def _norm(df):
    df.columns = [str(c).strip() for c in df.columns]
    bc = next((c for c in df.columns if c.lower() in
               ("package barcode", "paket-barcode", "barcode")), None)
    dc = next((c for c in df.columns if c.lower() in
               ("date of scan", "scan date", "datum", "date")), None)
    if not bc or not dc:
        return pd.DataFrame()
    o = df[[bc, dc]].copy()
    o.columns = ["Package Barcode", "Date of Scan"]
    o["Package Barcode"] = o["Package Barcode"].astype(str).str.strip()
    c = o["Date of Scan"]
    if pd.api.types.is_numeric_dtype(c):
        o["Date of Scan"] = pd.to_datetime(c, unit="D", origin="1899-12-30", errors="coerce")
    else:
        o["Date of Scan"] = pd.to_datetime(c, errors="coerce", utc=True).dt.tz_localize(None)
    return o.dropna(subset=["Date of Scan"])


def load_nas_express_mai():
    frames = []
    for f in os.listdir(NAS_EXPRESS):
        if not f.lower().endswith(".xlsx"):
            continue
        try:
            d = _norm(pd.read_excel(os.path.join(NAS_EXPRESS, f)))
            if not d.empty:
                frames.append(d)
        except Exception:
            pass
    if not frames:
        return pd.DataFrame()
    df = pd.concat(frames, ignore_index=True)
    return df[df["Date of Scan"].dt.to_period("M").astype(str) == "2026-05"]


def load_drive_express_mai(svc):
    from googleapiclient.http import MediaIoBaseDownload
    frames = []
    for fid in (GDRIVE_FOLDER_ABHOLER, GDRIVE_FOLDER_DHL_EXPRESS):
        res = svc.files().list(
            q=f"'{fid}' in parents and name contains 'DHL_Express' and trashed=false",
            fields="files(id,name)", pageSize=100).execute()
        for f in res.get("files", []):
            if "2026-05" not in f["name"]:
                continue
            try:
                req = svc.files().get_media(fileId=f["id"])
                buf = io.BytesIO(); dl = MediaIoBaseDownload(buf, req); done = False
                while not done:
                    _, done = dl.next_chunk()
                buf.seek(0)
                d = _norm(pd.read_excel(buf))
                if not d.empty:
                    frames.append(d[d["Date of Scan"].dt.to_period("M").astype(str) == "2026-05"])
            except Exception as e:
                print(f"   ⚠ {f['name']}: {e}")
    return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()


# ── Main ──────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    print("=" * 60)
    print(f"Mai-Express Quelldatei bauen  |  DRY_RUN: {DRY_RUN}")
    print("=" * 60)

    from nas_to_drive_migration import _get_drive_service
    svc = _get_drive_service()

    pdf   = load_pdf_express_mai()
    nas   = load_nas_express_mai()
    drive = load_drive_express_mai(svc)

    print(f"\nQuellen (vor Dedup):")
    print(f"  PDF (Gimli):  {len(pdf):>6}")
    print(f"  NAS:          {len(nas):>6}")
    print(f"  Drive:        {len(drive):>6}")

    allmai = pd.concat([pdf, nas, drive], ignore_index=True)
    allmai["Date of Scan"] = pd.to_datetime(allmai["Date of Scan"], errors="coerce", utc=True).dt.tz_localize(None)
    allmai = allmai.dropna(subset=["Date of Scan"])
    # Barcode normalisieren für sauberen Dedup
    allmai["Package Barcode"] = allmai["Package Barcode"].astype(str).str.strip().str.lstrip("0")
    vorher = len(allmai)
    allmai = allmai.drop_duplicates(subset=["Package Barcode"], keep="first")

    print(f"\n=== UNION dedupliziert: {len(allmai)} (von {vorher} vor Dedup) ===")
    print(f"    Ziel laut Priority List: ~14.084")
    print(f"    Differenz zum Ziel: {14084 - len(allmai):+}")
    tage = sorted(allmai['Date of Scan'].dt.day.unique())
    print(f"    Abgedeckte Mai-Tage: {tage}")

    if not DRY_RUN:
        from googleapiclient.http import MediaIoBaseUpload
        buf = io.BytesIO()
        allmai[["Package Barcode", "Date of Scan"]].to_excel(buf, index=False)
        buf.seek(0)
        media = MediaIoBaseUpload(buf,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        svc.files().create(
            body={"name": "DHL_Express_Archiv_2026-05.xlsx",
                  "parents": [GDRIVE_FOLDER_DHL_EXPRESS]},
            media_body=media, fields="id").execute()
        print("\n✅ DHL_Express_Archiv_2026-05.xlsx (Quelldatei) hochgeladen.")
