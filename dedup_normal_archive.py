"""
dedup_normal_archive.py
=======================
Bereinigt die DHL_Normal_Archiv_*.xlsx auf Google Drive:
Entfernt Barcode-Format-Duplikate (00340… vs 340… = dasselbe Paket).

Lädt jede Datei, dedupliziert per normalisiertem Barcode (clean_barcode +
führende Nullen entfernt), und lädt sie nur dann neu hoch, wenn sich die
Zeilenzahl ändert. Express ist nicht betroffen (keine führenden Nullen).

DRY_RUN = True → nur zeigen was sich ändern würde.
"""

import sys, io, pickle
import pandas as pd
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload

sys.stdout.reconfigure(encoding="utf-8")

DRY_RUN = False

GDRIVE_FOLDER_DHL_NORMAL = "1G5_i9zUvjqVBCnE_ME-fPkhK4j5SPWpe"


def clean_barcode(value):
    if pd.isna(value):
        return ""
    s = str(value).strip()
    if s.startswith('="') and s.endswith('"') and len(s) >= 4:
        s = s[2:-1].strip()
    if s.startswith("'"):
        s = s[1:].lstrip()
    s = s.strip()
    if s.endswith(".0") and s[:-2].lstrip("-").isdigit():
        s = s[:-2]
    return s


def _svc():
    with open("token.json", "rb") as f:
        creds = pickle.load(f)
    return build("drive", "v3", credentials=creds, cache_discovery=False)


def _download(svc, fid):
    req = svc.files().get_media(fileId=fid)
    buf = io.BytesIO(); dl = MediaIoBaseDownload(buf, req); done = False
    while not done:
        _, done = dl.next_chunk()
    buf.seek(0)
    return pd.read_excel(buf)


if __name__ == "__main__":
    print("=" * 60)
    print(f"DHL-Normal-Archive bereinigen  |  DRY_RUN: {DRY_RUN}")
    print("=" * 60)
    svc = _svc()

    res = svc.files().list(
        q=f"'{GDRIVE_FOLDER_DHL_NORMAL}' in parents and name contains 'Archiv' and trashed=false",
        fields="files(id,name)", orderBy="name", pageSize=100).execute()

    for f in sorted(res.get("files", []), key=lambda x: x["name"]):
        df = _download(svc, f["id"])
        df.columns = [str(c).strip() for c in df.columns]
        bc = next((c for c in df.columns if "barcode" in c.lower()), None)
        if not bc:
            print(f"  {f['name']}: keine Barcode-Spalte – übersprungen")
            continue

        vorher = len(df)
        key = df[bc].map(clean_barcode).str.lstrip("0")
        dedup = df.assign(_k=key)
        dedup = dedup[dedup["_k"].str.len() > 0]
        dedup = dedup.drop_duplicates(subset=["_k"], keep="last").drop(columns="_k")
        nachher = len(dedup)

        if nachher == vorher:
            print(f"  {f['name']}: {vorher} Zeilen – sauber, keine Änderung")
            continue

        print(f"  {f['name']}: {vorher} → {nachher}  (-{vorher - nachher} Duplikate)")
        if DRY_RUN:
            print(f"      🔵 DRY_RUN – würde neu hochladen")
            continue

        # alte Datei löschen, bereinigte neu hochladen (gleicher Name)
        svc.files().delete(fileId=f["id"]).execute()
        buf = io.BytesIO()
        dedup.to_excel(buf, index=False)
        buf.seek(0)
        media = MediaIoBaseUpload(buf,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        svc.files().create(
            body={"name": f["name"], "parents": [GDRIVE_FOLDER_DHL_NORMAL]},
            media_body=media, fields="id").execute()
        print(f"      ✅ bereinigt neu hochgeladen")

    print("=" * 60)
    print("Trockenlauf fertig." if DRY_RUN else "Bereinigung abgeschlossen!")
