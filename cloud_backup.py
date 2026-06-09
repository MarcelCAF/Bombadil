"""
cloud_backup.py
==============
Eigenständiges Backup-Skript für GitHub Actions (läuft in der Cloud, ohne PC,
ohne Bombadil). Sichert täglich aus OrcaScan auf Google Drive:

  Abholer_DB   → Abholer_DB_Backup_YYYY-MM-DD.xlsx
  DHL Normal   → DHL_Normal_Backup_YYYY-MM-DD.xlsx
  DHL Express  → DHL_Express_Backup_YYYY-MM-DD.xlsx
  Tagesbote    → Tagesbote_Backup_YYYY-MM-DD.xlsx

Auth: Google Service Account (headless, kein Browser-Login).
Config: Umgebungsvariablen (GitHub Secrets) oder lokale Dateien (.env / service_account.json).
"""

import os, sys, io, json, time, pickle, urllib.request, urllib.error
from datetime import date
from pathlib import Path

import pandas as pd
from google.oauth2.credentials import Credentials as UserCredentials
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

# ── Config aus Umgebung (GitHub Secrets) oder .env (lokal) ───────────────────
def _load_local_env():
    """Lädt .env lokal (für Testläufe). In GitHub Actions kommen die Werte
    direkt aus den Secrets/Umgebungsvariablen."""
    p = Path(__file__).resolve().parent / ".env"
    if p.exists():
        for line in p.read_text(encoding="utf-8").splitlines():
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                k, v = line.split("=", 1)
                os.environ.setdefault(k.strip(), v.strip())

_load_local_env()

ORCA_API_KEY  = os.getenv("ORCA_API_KEY")
ORCA_BASE_URL = os.getenv("ORCA_BASE_URL", "https://api.orcascan.com/v1")

SHEETS = [
    # (env-Variable Sheet-ID, Datei-Präfix, Drive-Folder-ID, drop_cols)
    (os.getenv("ORCA_ABHOLER_SHEET_ID"),    "Abholer_DB_Backup",  "1OSnMmPf--uqt4ulDGy3ILT61pUuCPzpn", {"Unterschrift", "Paketfoto"}),
    (os.getenv("ORCA_DHL_NORMAL_SHEET_ID"), "DHL_Normal_Backup",  "1G5_i9zUvjqVBCnE_ME-fPkhK4j5SPWpe", {"signature", "packagePhoto"}),
    (os.getenv("ORCA_DHL_EX_SHEET_ID"),     "DHL_Express_Backup", "1J_oZcBmPECYfL7xj5U0oxgao--sdzS_c", {"signature", "packagePhoto"}),
    (os.getenv("ORCA_TAGESBOTE_SHEET_ID"),  "Tagesbote_Backup",   "1a5Wg-fFhF11ux5d7Tl5oVqcq9fRkp5yX", set()),
]


# ── Google Drive Service (OAuth-Token, headless via refresh_token) ───────────
# WICHTIG: Service Accounts haben keinen eigenen Drive-Speicher und können
# nicht in persönliche Ordner schreiben. Daher OAuth-Token (Marcels Konto) –
# das refresh_token erlaubt headless ein neues Access-Token (kein Browser).
def _drive_service():
    """OAuth-Credentials aus GOOGLE_OAUTH_TOKEN (JSON, GitHub Secret) oder
    lokaler token.json (pickle). Wird per refresh_token erneuert."""
    tok = os.getenv("GOOGLE_OAUTH_TOKEN")
    if tok:
        info = json.loads(tok)
        creds = UserCredentials(
            token=info.get("token"),
            refresh_token=info["refresh_token"],
            client_id=info["client_id"],
            client_secret=info["client_secret"],
            token_uri=info.get("token_uri", "https://oauth2.googleapis.com/token"),
            scopes=info.get("scopes", ["https://www.googleapis.com/auth/drive.file"]),
        )
    else:
        path = Path(__file__).resolve().parent / "token.json"
        with open(path, "rb") as f:
            creds = pickle.load(f)
    creds.refresh(Request())   # neues Access-Token holen (headless, ohne Browser)
    return build("drive", "v3", credentials=creds, cache_discovery=False)


# ── OrcaScan-Sheet laden (alle Seiten, mit Retry) ────────────────────────────
def fetch_sheet(sheet_id, drop_cols=None):
    drop_cols = drop_cols or set()
    all_rows, seen, page, MAX = [], set(), 1, 50
    while page <= MAX:
        url = f"{ORCA_BASE_URL}/sheets/{sheet_id}/rows?withTitles=true&page={page}&limit=500"
        req = urllib.request.Request(url, headers={"Authorization": f"Bearer {ORCA_API_KEY}"})
        for attempt in range(4):
            try:
                with urllib.request.urlopen(req, timeout=60) as resp:
                    data = json.loads(resp.read())
                break
            except urllib.error.HTTPError as e:
                if e.code == 429 and attempt < 3:
                    time.sleep((attempt + 1) * 3); continue
                raise RuntimeError(f"OrcaScan HTTP {e.code}: {e.reason}") from e
            except Exception as e:
                if attempt < 3:
                    time.sleep((attempt + 1) * 5); continue
                raise RuntimeError(f"OrcaScan Fehler: {e}") from e
        rows = data.get("data", [])
        if not rows:
            break
        new_ids = {r.get("_id") for r in rows if r.get("_id")}
        if new_ids and new_ids.issubset(seen):
            break
        seen.update(new_ids)
        for r in rows:
            for c in drop_cols:
                r.pop(c, None)
        all_rows.extend(rows)
        if len(rows) < 500:
            break
        page += 1
        time.sleep(0.4)
    return pd.DataFrame(all_rows)


# ── Auf Drive hochladen (alte gleichnamige Datei vorher löschen) ─────────────
def upload(svc, df, folder_id, filename):
    # gleichnamige Datei(en) im Ordner entfernen (Backup ist Tages-Snapshot)
    res = svc.files().list(
        q=f"name = '{filename}' and '{folder_id}' in parents and trashed = false",
        fields="files(id)", pageSize=10,
        supportsAllDrives=True, includeItemsFromAllDrives=True).execute()
    for f in res.get("files", []):
        try: svc.files().delete(fileId=f["id"], supportsAllDrives=True).execute()
        except Exception: pass
    # tz-aware Spalten für Excel entfernen
    dfx = df.copy()
    for col in dfx.select_dtypes(include=["datetimetz"]).columns:
        dfx[col] = dfx[col].dt.tz_localize(None)
    buf = io.BytesIO()
    dfx.to_excel(buf, index=False)
    buf.seek(0)
    media = MediaIoBaseUpload(buf,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        resumable=False)
    svc.files().create(body={"name": filename, "parents": [folder_id]},
                       media_body=media, fields="id", supportsAllDrives=True).execute()


# ── Main ─────────────────────────────────────────────────────────────────────
def main():
    if not ORCA_API_KEY:
        print("FEHLER: ORCA_API_KEY fehlt."); sys.exit(1)
    today = date.today().strftime("%Y-%m-%d")
    svc = _drive_service()
    fehler = 0
    for sheet_id, prefix, folder_id, drop_cols in SHEETS:
        name = f"{prefix}_{today}.xlsx"
        if not sheet_id:
            print(f"⏭  {prefix}: Sheet-ID fehlt – übersprungen"); continue
        try:
            df = fetch_sheet(sheet_id, drop_cols)
            if df.empty:
                print(f"⚠  {name}: leer – nicht hochgeladen"); continue
            upload(svc, df, folder_id, name)
            print(f"✅  {name}: {len(df)} Zeilen hochgeladen")
        except Exception as e:
            fehler += 1
            print(f"❌  {name}: {e}")
    print(f"\nFertig. {fehler} Fehler.")
    sys.exit(1 if fehler else 0)


if __name__ == "__main__":
    main()
