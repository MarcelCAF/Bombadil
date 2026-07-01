import pickle, sys
sys.stdout.reconfigure(encoding='utf-8')
from googleapiclient.discovery import build

with open(r'C:\Users\Abfuellung 15\Documents\Marcels Skripts\Bombadil\token.json', 'rb') as fh:
    creds = pickle.load(fh)

service = build('drive', 'v3', credentials=creds, cache_discovery=False)

# Alle Excel-Dateien holen (pageToken fuer mehr als 20)
all_files = []
page_token = None
while True:
    kwargs = dict(
        fields="nextPageToken,files(id,name,parents)",
        pageSize=100,
        q="trashed=false and mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'"
    )
    if page_token:
        kwargs["pageToken"] = page_token
    result = service.files().list(**kwargs).execute()
    all_files.extend(result.get("files", []))
    page_token = result.get("nextPageToken")
    if not page_token:
        break

print(f"Gesamt Excel-Dateien im Drive: {len(all_files)}")
dhl_files = [f for f in all_files if "dhl" in f["name"].lower()]
print(f"DHL-Dateien: {len(dhl_files)}")
for f in sorted(dhl_files, key=lambda x: x["name"]):
    print(f"  {f['name']}  ID:{f['id'][:20]}  Ordner:{f.get('parents',['?'])[0][:20]}")
