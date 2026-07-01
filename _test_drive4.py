import pickle, sys, io
sys.stdout.reconfigure(encoding='utf-8')
import pandas as pd
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

with open(r'C:\Users\Abfuellung 15\Documents\Marcels Skripts\Bombadil\token.json', 'rb') as fh:
    creds = pickle.load(fh)

service = build('drive', 'v3', credentials=creds, cache_discovery=False)

files = {
    "Express": "148-U4kDlq2M7kn-sW2r",
    "Normal":  "10zGHE0axC0jAQnBJdEP",
}

for name, file_id in files.items():
    req  = service.files().get_media(fileId=file_id)
    buf  = io.BytesIO()
    dl   = MediaIoBaseDownload(buf, req)
    done = False
    while not done:
        _, done = dl.next_chunk()
    buf.seek(0)
    df = pd.read_excel(buf)
    df.columns = [str(c).strip() for c in df.columns]
    scan_col = next((c for c in df.columns if "scan" in c.lower() or "date" in c.lower()), None)
    if scan_col:
        dates = pd.to_datetime(df[scan_col], errors='coerce').dt.date.dropna()
        print(f"{name}: {len(df)} Zeilen | Datum von {dates.min()} bis {dates.max()}")
        # Letzten 5 Tage anzeigen
        counts = dates.value_counts().sort_index()
        print("  Letzte 5 Tage:")
        for d, c in counts.tail(5).items():
            print(f"    {d}: {c} Pakete")
    else:
        print(f"{name}: {len(df)} Zeilen, Spalten: {df.columns.tolist()}")
