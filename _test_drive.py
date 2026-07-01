import pickle, sys
sys.stdout.reconfigure(encoding='utf-8')
from googleapiclient.discovery import build

with open(r'C:\Users\Abfuellung 15\Documents\Marcels Skripts\Bombadil\token.json', 'rb') as fh:
    creds = pickle.load(fh)

service = build('drive', 'v3', credentials=creds, cache_discovery=False)

for name, folder_id in [
    ("Express", "1J_oZcBmPECYfL7xj5U0oxgao--sdzS_c"),
    ("Normal",  "1G5_i9zUvjqVBCnE_ME-fPkhK4j5SPWpe"),
]:
    q = f"'{folder_id}' in parents and trashed=false"
    result = service.files().list(q=q, fields="files(id,name,mimeType)", pageSize=20).execute()
    files = result.get("files", [])
    print(f"\n{name}-Ordner: {len(files)} Dateien")
    for f in files[:10]:
        print(f"  {f['name']}  ({f['mimeType'][:50]})")
