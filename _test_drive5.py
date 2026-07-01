import pickle, sys, io
sys.stdout.reconfigure(encoding='utf-8')
import pandas as pd
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

with open(r'C:\Users\Abfuellung 15\Documents\Marcels Skripts\Bombadil\token.json', 'rb') as fh:
    creds = pickle.load(fh)

service = build('drive', 'v3', credentials=creds, cache_discovery=False)

# Vollstaendige IDs holen
result = service.files().list(
    q="trashed=false and (name='DHL_Express_Archiv_2026-05.xlsx' or name='DHL_Normal_Archiv_2026-05.xlsx')",
    fields="files(id,name)"
).execute()

for f in result.get("files", []):
    print(f"Name: {f['name']}")
    print(f"ID:   {f['id']}")
    req = service.files().get_media(fileId=f['id'])
    buf = io.BytesIO()
    dl  = MediaIoBaseDownload(buf, req)
    done = False
    while not done:
        _, done = dl.next_chunk()
    buf.seek(0)
    df = pd.read_excel(buf)
    df.columns = [str(c).strip() for c in df.columns]
    scan_col = next((c for c in df.columns if "scan" in c.lower() or "date" in c.lower()), None)
    if scan_col:
        dates = pd.to_datetime(df[scan_col], errors='coerce').dt.date.dropna()
        print(f"Zeilen: {len(df)} | von {dates.min()} bis {dates.max()}")
        for d, c in dates.value_counts().sort_index().tail(8).items():
            print(f"  {d}: {c}")
    print()
