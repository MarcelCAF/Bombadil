import sys
sys.stdout.reconfigure(encoding='utf-8')
path = r'C:\Users\Abfuellung 15\Documents\Marcels Skripts\Bombadil\Bombadil.py'
with open(path, 'r', encoding='utf-8') as f:
    content = f.read()

errors = []

# 1. Filter-Optionen
old1 = '        _status_opts = ["Alle", "Offen", "Verpackt", "Am Standort", "Abgeholt", "Retoure"]\n'
new1 = '        _status_opts = ["Alle", "Offen", "Verpackt", "Am Standort", "Abgeholt", "Tour 1", "Tour 2"]\n'
if old1 in content:
    content = content.replace(old1, new1, 1); print('Filter-Optionen: OK')
else:
    errors.append('Filter-Optionen: MISS'); print(repr(content[content.find('_status_opts'):content.find('_status_opts')+80]))

# 2. Sortier-Optionen
old2 = (
    '        _sort_opts = [\n'
    '            "Standard",\n'
    '            "Name A→Z", "Name Z→A",\n'
    '            "Barcode A→Z",\n'
    '            "Verpackt (neu→alt)", "Verpackt (alt→neu)",\n'
    '            "Abholbereit (neu→alt)",\n'
    '            "Ziel-Kiosk A→Z",\n'
    '        ]\n'
)
new2 = (
    '        _sort_opts = [\n'
    '            "Standard",\n'
    '            "Name A→Z",\n'
    '            "Ziel-Kiosk A→Z",\n'
    '            "Älteste zuerst",\n'
    '        ]\n'
)
if old2 in content:
    content = content.replace(old2, new2, 1); print('Sortier-Optionen: OK')
else:
    errors.append('Sortier-Optionen: MISS'); print(repr(content[content.find('_sort_opts'):content.find('_sort_opts')+200]))

# 3. Filter-Logik in _refresh_ui: Retoure entfernen, Tour 1/2 hinzufügen
old3 = (
    '        elif f == "Abgeholt":\n'
    '            rows = [r for r in rows if r.get("db_status") == "abgeholt"]\n'
    '        elif f == "Retoure":\n'
    '            rows = [r for r in rows if r.get("db_status") == "retoure"]\n'
    '        elif f == "Verpackt":\n'
)
new3 = (
    '        elif f == "Abgeholt":\n'
    '            rows = [r for r in rows if r.get("db_status") == "abgeholt"]\n'
    '        elif f == "Tour 1":\n'
    '            rows = [r for r in rows if r.get("tour") == "T1"]\n'
    '        elif f == "Tour 2":\n'
    '            rows = [r for r in rows if r.get("tour") == "T2"]\n'
    '        elif f == "Verpackt":\n'
)
if old3 in content:
    content = content.replace(old3, new3, 1); print('Filter-Logik: OK')
else:
    errors.append('Filter-Logik: MISS')

# 4. Sortier-Logik: alte Optionen ersetzen, Älteste zuerst hinzufügen
old4 = (
    '        if s == "Name A→Z":\n'
    '            rows = sorted(rows, key=lambda r: str(r["name"] or "").lower() or _none_last)\n'
    '        elif s == "Name Z→A":\n'
    '            rows = sorted(rows, key=lambda r: str(r["name"] or "").lower() or _none_last,\n'
    '                          reverse=True)\n'
    '        elif s == "Barcode A→Z":\n'
    '            rows = sorted(rows, key=lambda r: str(r["barcode"] or "").lower())\n'
    '        elif s == "Verpackt (neu→alt)":\n'
    '            rows = sorted(rows, key=lambda r: r["verpackt_at"] or "", reverse=True)\n'
    '        elif s == "Verpackt (alt→neu)":\n'
    '            rows = sorted(rows, key=lambda r: r["verpackt_at"] or _none_last)\n'
    '        elif s == "Abholbereit (neu→alt)":\n'
    '            rows = sorted(rows, key=lambda r: r["abholbereit_at"] or "", reverse=True)\n'
    '        elif s == "Ziel-Kiosk A→Z":\n'
    '            rows = sorted(rows, key=lambda r: str(r["zielkiosk"] or "").lower() or _none_last)\n'
)
new4 = (
    '        if s == "Name A→Z":\n'
    '            rows = sorted(rows, key=lambda r: str(r["name"] or "").lower() or _none_last)\n'
    '        elif s == "Ziel-Kiosk A→Z":\n'
    '            rows = sorted(rows, key=lambda r: str(r["zielkiosk"] or "").lower() or _none_last)\n'
    '        elif s == "Älteste zuerst":\n'
    '            rows = sorted(rows, key=lambda r: r["verpackt_at"] or _none_last)\n'
)
if old4 in content:
    content = content.replace(old4, new4, 1); print('Sortier-Logik: OK')
else:
    errors.append('Sortier-Logik: MISS')

if not errors:
    with open(path, 'w', encoding='utf-8') as f:
        f.write(content)
    print('Datei gespeichert.')
else:
    print('FEHLER:', errors)
