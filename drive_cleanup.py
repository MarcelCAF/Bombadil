"""
drive_cleanup.py
================
Räumt den Google Drive nach der NAS-Migration auf:

  1. Express-Mai reparieren: vollen Mai aus NAS in den EXPRESS-Ordner laden
     (im EXPRESS-Ordner lag nur ein PDF-Teilmonat 22.–26.05.)
  2. Fehlplatzierte DHL-Archive aus dem Abholer_DB-Ordner löschen
  3. Echte Duplikate (gleicher Name im selben Ordner) löschen – ältestes behalten

Sicherheit: DRY_RUN = True → nur Vorschau, kein Löschen/Upload.
"""

import sys, io, pickle
from pathlib import Path
from collections import defaultdict

import pandas as pd
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload, MediaIoBaseDownload

sys.stdout.reconfigure(encoding="utf-8")

# Funktionen aus dem Migrations-Skript wiederverwenden
from nas_to_drive_migration import (
    _get_drive_service, load_nas_folder, NAS_EXPRESS_DIR, _normalize_df,
)

DRY_RUN = False   # ← auf False setzen um wirklich zu löschen/hochladen

GDRIVE_FOLDER_ABHOLER     = "1OSnMmPf--uqt4ulDGy3ILT61pUuCPzpn"
GDRIVE_FOLDER_DHL_EXPRESS = "1J_oZcBmPECYfL7xj5U0oxgao--sdzS_c"

ORDNER = {
    "Abholer_DB":       "1OSnMmPf--uqt4ulDGy3ILT61pUuCPzpn",
    "DHL (Normal)":     "1G5_i9zUvjqVBCnE_ME-fPkhK4j5SPWpe",
    "EXPRESS":          "1J_oZcBmPECYfL7xj5U0oxgao--sdzS_c",
    "Tagesbote Upload": "1a5Wg-fFhF11ux5d7Tl5oVqcq9fRkp5yX",
    "Tourlisten":       "1ALz0IR6JeoUVXFMAMxida4MT7lQ0ArDd",
}


def list_files(svc, folder_id):
    res = svc.files().list(
        q=f"'{folder_id}' in parents and trashed=false",
        fields="files(id,name,createdTime)", orderBy="name", pageSize=500
    ).execute()
    return res.get("files", [])


# ── 1. Express-Mai reparieren: EINE vollständige Datei aus allen Quellen ──────

def _download_df(svc, fid):
    req = svc.files().get_media(fileId=fid)
    buf = io.BytesIO(); dl = MediaIoBaseDownload(buf, req); done = False
    while not done:
        _, done = dl.next_chunk()
    buf.seek(0)
    return pd.read_excel(buf)

def repair_express_mai(svc):
    print("\n" + "=" * 60)
    print("1. Express-Mai reparieren – EINE vollständige Datei bauen")
    print("=" * 60)

    quellen = []

    # a) NAS (Anfang Mai 02.–21.)
    nas = load_nas_folder(NAS_EXPRESS_DIR, "DHL Express (NAS)")
    if not nas.empty:
        quellen.append(nas[nas["Date of Scan"].dt.to_period("M").astype(str) == "2026-05"])

    # b) Fehlplatzierte Datei im Abholer-Ordner (Ende Mai 27.–30.)
    # c) PDF-Teilarchiv im EXPRESS-Ordner (22.–26.)
    # Unterscheidung über ORDNER, nicht Name: im EXPRESS-Ordner eine evtl. schon
    # existierende finale Datei überspringen (Idempotenz bei Wiederholung).
    extra_ids = []
    for ordner_id in (GDRIVE_FOLDER_ABHOLER, GDRIVE_FOLDER_DHL_EXPRESS):
        for f in list_files(svc, ordner_id):
            n = f["name"]
            if not (n.startswith("DHL_Express") and "2026-05" in n):
                continue
            # finale Zieldatei im EXPRESS-Ordner nicht erneut einlesen
            if ordner_id == GDRIVE_FOLDER_DHL_EXPRESS and n == "DHL_Express_Archiv_2026-05.xlsx":
                continue
            try:
                d = _normalize_df(_download_df(svc, f["id"]))
                if d is not None:
                    d = d[d["Date of Scan"].dt.to_period("M").astype(str) == "2026-05"]
                    quellen.append(d)
                    extra_ids.append((n, f["id"]))
            except Exception as e:
                print(f"   ⚠ {n}: {e}")

    if not quellen:
        print("   ⚠ Keine Mai-Quellen gefunden.")
        return None

    allmai = pd.concat(quellen, ignore_index=True)
    # Timezone vereinheitlichen (NAS=naiv, Drive=UTC-aware) → alles naiv
    allmai["Date of Scan"] = pd.to_datetime(
        allmai["Date of Scan"], errors="coerce", utc=True).dt.tz_localize(None)
    allmai = allmai.dropna(subset=["Date of Scan"])
    allmai = allmai.drop_duplicates(subset=["Package Barcode"], keep="last")
    tage = sorted(allmai["Date of Scan"].dt.day.unique())
    print(f"\n   Zusammengeführt: {len(allmai)} Zeilen, "
          f"{allmai['Date of Scan'].min()} – {allmai['Date of Scan'].max()}")
    print(f"   Abgedeckte Tage: {tage}")
    print(f"   Teil-Quellen die danach gelöscht werden: {[n for n,_ in extra_ids]}")

    if DRY_RUN:
        print("   🔵 DRY_RUN – würde EINE DHL_Express_Archiv_2026-05.xlsx (EXPRESS-Ordner) hochladen")
        print("              und die Teil-Fragmente (PDF + fehlplatziert) löschen")
        return None

    buf = io.BytesIO()
    allmai[["Package Barcode", "Date of Scan"]].to_excel(buf, index=False)
    buf.seek(0)
    media = MediaIoBaseUpload(buf,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    svc.files().create(
        body={"name": "DHL_Express_Archiv_2026-05.xlsx",
              "parents": [GDRIVE_FOLDER_DHL_EXPRESS]},
        media_body=media, fields="id").execute()
    print("   ✅ Vollständige DHL_Express_Archiv_2026-05.xlsx hochgeladen.")

    # PDF-Fragment löschen (die fehlplatzierten DHL-Dateien räumt Schritt 2)
    for n, fid in extra_ids:
        if n.endswith("PDF.xlsx"):
            svc.files().delete(fileId=fid).execute()
            print(f"   ✅ Fragment gelöscht: {n}")
    return None


# ── 2. Fehlplatzierte DHL-Dateien im Abholer-Ordner ───────────────────────────

def remove_misplaced_dhl(svc):
    print("\n" + "=" * 60)
    print("2. Fehlplatzierte DHL-Archive aus Abholer_DB-Ordner entfernen")
    print("=" * 60)

    files = list_files(svc, GDRIVE_FOLDER_ABHOLER)
    misplaced = [f for f in files if f["name"].startswith("DHL_")]
    if not misplaced:
        print("   (keine fehlplatzierten DHL-Dateien gefunden)")
        return
    for f in misplaced:
        if DRY_RUN:
            print(f"   🔵 DRY_RUN – würde löschen: {f['name']}  (id={f['id']})")
        else:
            svc.files().delete(fileId=f["id"]).execute()
            print(f"   ✅ gelöscht: {f['name']}")


# ── 3. Duplikate (gleicher Name im selben Ordner) ─────────────────────────────

def remove_duplicates(svc):
    print("\n" + "=" * 60)
    print("3. Duplikate entfernen (ältestes behalten, Rest löschen)")
    print("=" * 60)

    for name, fid in ORDNER.items():
        files = list_files(svc, fid)
        by_name = defaultdict(list)
        for f in files:
            by_name[f["name"]].append(f)

        dups = {n: fs for n, fs in by_name.items() if len(fs) > 1}
        if not dups:
            continue
        print(f"\n   📂 {name}:")
        for fname, fs in dups.items():
            fs_sorted = sorted(fs, key=lambda x: x["createdTime"])  # ältestes zuerst
            keep, remove = fs_sorted[0], fs_sorted[1:]
            print(f"      {fname}: {len(fs)}× → behalte {keep['createdTime'][:19]}, "
                  f"lösche {len(remove)}")
            for f in remove:
                if DRY_RUN:
                    print(f"         🔵 DRY_RUN – würde löschen id={f['id']} ({f['createdTime'][:19]})")
                else:
                    svc.files().delete(fileId=f["id"]).execute()
                    print(f"         ✅ gelöscht id={f['id']}")


# ── Main ──────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    print("=" * 60)
    print(f"Drive-Cleanup  |  DRY_RUN: {DRY_RUN}")
    print("=" * 60)
    if DRY_RUN:
        print("\n⚠  DRY_RUN aktiv – es wird NICHTS gelöscht oder hochgeladen.\n")

    svc = _get_drive_service()
    repair_express_mai(svc)
    remove_misplaced_dhl(svc)
    remove_duplicates(svc)

    print("\n" + "=" * 60)
    print("Trockenlauf fertig – prüfe oben." if DRY_RUN else "Cleanup abgeschlossen!")
    print("=" * 60)
