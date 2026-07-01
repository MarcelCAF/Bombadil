import pickle, sys
sys.stdout.reconfigure(encoding='utf-8')
from googleapiclient.discovery import build

with open(r'C:\Users\Abfuellung 15\Documents\Marcels Skripts\Bombadil\token.json', 'rb') as fh:
    creds = pickle.load(fh)

print("Scopes:", creds.scopes)
print("Gueltig:", creds.valid)

service = build('drive', 'v3', credentials=creds, cache_discovery=False)

# Alle Dateien des Accounts anzeigen (ohne Ordner-Filter)
result = service.files().list(
    fields="files(id,name,mimeType,parents)",
    pageSize=20,
    q="trashed=false and mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'"
).execute()
files = result.get("files", [])
print(f"\nAlle Excel-Dateien im Drive: {len(files)}")
for f in files[:15]:
    parents = f.get('parents', [])
    print(f"  {f['name']}  Ordner: {parents}")
