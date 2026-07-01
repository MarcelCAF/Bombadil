import sys
sys.stdout.reconfigure(encoding='utf-8')
path = r'C:\Users\Abfuellung 15\Documents\Marcels Skripts\Bombadil\Bombadil.py'
with open(path, 'r', encoding='utf-8') as f:
    content = f.read()

# 1. DataFrame-Aufbau ersetzen
old_df = (
    '        df = _pd.DataFrame([{\n'
    '            "Paket-Barcode": r["barcode"],\n'
    '            "Name":          r["name"],\n'
    '            "Ziel-Kiosk":    r["zielkiosk"],\n'
    '            "Packstatus":    r["tb_status"],\n'
    '            "Verpackt_At":   r["verpackt_at"],\n'
    '            "Abholbereit_At": r["abholbereit_at"],\n'
    '        } for r in tour_rows])\n'
)

new_df = (
    '        def _last4(v):\n'
    '            s = str(v).strip()\n'
    '            return s[-4:] if len(s) >= 4 else s\n'
    '        df = _pd.DataFrame([{\n'
    '            "Paket-Barcode":     r["barcode"],\n'
    '            "Bestellnummer":     _last4(r["barcode"]),\n'
    '            "Datum":             r["scan_datum"],\n'
    '            "Vorname":           "",\n'
    '            "Name":              r["name"],\n'
    '            "Ziel-Kiosk":        r["zielkiosk"],\n'
    '            "Status":            r["tb_status"],\n'
    '            "Bestellwert":       "",\n'
    '            "Versicher.":        "",\n'
    '            "Email":             "",\n'
    '            "Lieferung-Adresse": "",\n'
    '            "Lieferung":         "",\n'
    '            "Notizen":           "",\n'
    '            "Kontrollstatus":    r["tb_status"],\n'
    '            "Zahlung":           "",\n'
    '            "Rezept":            "",\n'
    '        } for r in tour_rows])\n'
)

# 2. Spaltenbreiten anpassen (16 Spalten wie in konvertierung.py)
old_w = "                _col_w = [22, 28, 13, 16, 20, 20]\n"
new_w = "                _col_w = [22, 14, 16, 10, 28, 13, 16, 13, 13, 28, 28, 12, 16, 15, 12, 10]\n"

ok1 = old_df in content
ok2 = old_w in content

if ok1:
    content = content.replace(old_df, new_df, 1)
    print('DataFrame: OK')
else:
    print('DataFrame: MISS')
    idx = content.find('df = _pd.DataFrame')
    print(repr(content[idx-10:idx+80]))

if ok2:
    content = content.replace(old_w, new_w, 1)
    print('Spaltenbreiten: OK')
else:
    print('Spaltenbreiten: MISS')
    idx = content.find('_col_w')
    print(repr(content[idx-10:idx+60]))

if ok1 and ok2:
    with open(path, 'w', encoding='utf-8') as f:
        f.write(content)
    print('Datei gespeichert.')
