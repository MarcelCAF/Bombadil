"""
Erstellt aus den PDF-extrahierten Daten eine Excel-Datei
fuer die fehlenden Express-Tage (22./23./26.05.2026)
und laedt sie auf Google Drive hoch.
"""
import pickle, sys, io
sys.stdout.reconfigure(encoding='utf-8')
import pandas as pd
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
from openpyxl import Workbook

BOMBADIL_DIR  = r"C:\Users\Abfuellung 15\Documents\Marcels Skripts\Bombadil"
EXPRESS_QUELLE = r"C:\Users\Abfuellung 15\Downloads\dhl_express_sendungen.xlsx"

# Fehlende Tage (OrcaScan = 0)
FEHLENDE_TAGE = ["2026-05-22", "2026-05-23", "2026-05-26"]

# Google Drive Ordner: EXPRESS
GDRIVE_FOLDER_EXPRESS = "1J_oZcBmPECYfL7xj5U0oxgao--sdzS_c"

# ── Drive-Verbindung ─────────────────────────────────────────────────────────
with open(f"{BOMBADIL_DIR}\\token.json", "rb") as fh:
    creds = pickle.load(fh)
service = build("drive", "v3", credentials=creds, cache_discovery=False)

# ── PDF-Daten laden und filtern ──────────────────────────────────────────────
df = pd.read_excel(EXPRESS_QUELLE)
df.columns = [str(c).strip() for c in df.columns]
df["DATE OF SCAN"] = pd.to_datetime(df["DATE OF SCAN"], errors="coerce")
df["_datum_str"]  = df["DATE OF SCAN"].dt.strftime("%Y-%m-%d")

luecke_df = df[df["_datum_str"].isin(FEHLENDE_TAGE)].copy()
luecke_df = luecke_df[["PACKAGE BARCODE", "DATE OF SCAN"]].copy()
luecke_df.columns = ["Package Barcode", "Date of Scan"]
luecke_df = luecke_df.sort_values("Date of Scan").reset_index(drop=True)

print(f"Gefundene Eintraege fuer fehlende Tage:")
for tag in FEHLENDE_TAGE:
    n = int((luecke_df["Date of Scan"].dt.strftime("%Y-%m-%d") == tag).sum())
    print(f"  {tag}: {n} Pakete")
print(f"  Gesamt: {len(luecke_df)} Eintraege")

# ── Excel erstellen ──────────────────────────────────────────────────────────
buf = io.BytesIO()
with pd.ExcelWriter(buf, engine="openpyxl") as writer:
    luecke_df.to_excel(writer, index=False, sheet_name="Sheet1")
buf.seek(0)

# ── Auf Drive hochladen ──────────────────────────────────────────────────────
dateiname = "DHL_Express_Archiv_2026-05-22_bis_26_PDF.xlsx"

file_meta = {
    "name": dateiname,
    "parents": [GDRIVE_FOLDER_EXPRESS],
}
media = MediaIoBaseUpload(
    buf,
    mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    resumable=False,
)
uploaded = service.files().create(
    body=file_meta,
    media_body=media,
    fields="id,name"
).execute()

print(f"\nErfolgreich hochgeladen:")
print(f"  Name: {uploaded['name']}")
print(f"  ID:   {uploaded['id']}")
print(f"\nBombadil liest diese Datei beim naechsten Laden automatisch ein.")
