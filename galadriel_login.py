# galadriel_login.py - Einmaliger Login, erzeugt galadriel_token.json
from google_auth_oauthlib.flow import InstalledAppFlow
import pickle

# Berechtigung: nur Dateien anlegen/verwalten, die Galadriel selbst erstellt
SCOPES = ["https://www.googleapis.com/auth/drive.file"]

flow = InstalledAppFlow.from_client_secrets_file("oauth_credentials.json", SCOPES)
creds = flow.run_local_server(port=0)

with open("galadriel_token.json", "wb") as f:
    pickle.dump(creds, f)

print("Fertig! galadriel_token.json wurde erstellt.")