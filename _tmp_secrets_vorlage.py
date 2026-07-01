# Temporär: Secrets-Vorlage für Streamlit Cloud erzeugen – kann danach gelöscht werden
import json
from pathlib import Path

kandidaten = [
    Path(r"C:\Users\Abfuellung 15\Documents\Marcels Skripts\Bombadil\service_account.json"),
    Path(r"W:\Dokumentenaustausch\Tagesskripte\Bombadil\service_account.json"),
]
sa_path = next((p for p in kandidaten if p.exists()), None)
if not sa_path:
    raise SystemExit("service_account.json nicht gefunden!")
print("Nutze:", sa_path)

sa = json.loads(sa_path.read_text(encoding="utf-8"))

zeilen = ["# In Streamlit Cloud unter Settings -> Secrets einfuegen", ""]
zeilen.append('app_password = "HIER-DEIN-WUNSCH-PASSWORT"')
zeilen.append("")
zeilen.append("[gcp_service_account]")
for k, v in sa.items():
    zeilen.append(f"{k} = " + json.dumps(v, ensure_ascii=False))

out = Path(r"C:\Users\Abfuellung 15\Documents\Marcels Skripts\Bombadil-Mobile\streamlit_secrets_VORLAGE.toml")
out.write_text("\n".join(zeilen), encoding="utf-8")
print("Vorlage geschrieben:", out)
