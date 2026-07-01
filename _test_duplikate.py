import pickle, sys, io
sys.stdout.reconfigure(encoding='utf-8')
import pandas as pd
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

with open(r'C:\Users\Abfuellung 15\Documents\Marcels Skripts\Bombadil\token.json', 'rb') as fh:
    creds = pickle.load(fh)

service = build('drive', 'v3', credentials=creds, cache_discovery=False)

for typ in ["Express", "Normal"]:
    result = service.files().list(
        q=f"trashed=false and name contains 'DHL_{typ}' and mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'",
        fields="files(id,name)"
    ).execute()
    files = result.get("files", [])
    print(f"\n=== {typ} ===")
    for f in files:
        req = service.files().get_media(fileId=f['id'])
        buf = io.BytesIO()
        dl  = MediaIoBaseDownload(buf, req)
        done = False
        while not done:
            _, done = dl.next_chunk()
        buf.seek(0)
        df = pd.read_excel(buf)
        df.columns = [str(c).strip() for c in df.columns]
        bc_col = next((c for c in df.columns if "barcode" in c.lower()), None)
        if bc_col:
            gesamt = len(df)
            unique = df[bc_col].nunique()
            duplikate = gesamt - unique
            print(f"  {f['name']}: {gesamt} Zeilen, {unique} eindeutige Barcodes, {duplikate} Duplikate ({duplikate/gesamt*100:.1f}%)")
