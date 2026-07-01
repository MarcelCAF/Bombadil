import sys
sys.stdout.reconfigure(encoding='utf-8')
path = r'C:\Users\Abfuellung 15\Documents\Marcels Skripts\Bombadil\Bombadil.py'
with open(path, 'r', encoding='utf-8') as f:
    content = f.read()

old = (
    '        s = self._sort_var.get()\n'
    '        _none_last = "ÿ"   # leere Strings nach hinten sortieren\n'
    '        if s == "Name A→Z":\n'
    '            rows = sorted(rows, key=lambda r: str(r["name"] or "").lower() or _none_last)\n'
    '        elif s == "Ziel-Kiosk A→Z":\n'
    '            rows = sorted(rows, key=lambda r: str(r["zielkiosk"] or "").lower() or _none_last)\n'
    '        elif s == "Älteste zuerst":\n'
    '            rows = sorted(rows, key=lambda r: r["verpackt_at"] or _none_last)\n'
)
new = (
    '        _none_last = "ÿ"   # leere Strings nach hinten sortieren\n'
    '        _col_keys = {\n'
    '            0: lambda r: str(r.get("tour", "") or "").lower(),\n'
    '            1: lambda r: str(r["barcode"] or "").lower(),\n'
    '            2: lambda r: str(r["name"] or "").lower() or _none_last,\n'
    '            3: lambda r: str(r["tb_status"] or "").lower(),\n'
    '            5: lambda r: str(r["verpackt_at"] or _none_last),\n'
    '            6: lambda r: str(r["abholbereit_at"] or _none_last),\n'
    '            7: lambda r: str(r["zielkiosk"] or "").lower() or _none_last,\n'
    '        }\n'
    '        if self._sort_col is not None and self._sort_dir > 0 and self._sort_col in _col_keys:\n'
    '            rows = sorted(rows, key=_col_keys[self._sort_col], reverse=(self._sort_dir == 2))\n'
)

if old in content:
    content = content.replace(old, new, 1)
    with open(path, 'w', encoding='utf-8') as f:
        f.write(content)
    print('OK')
else:
    print('MISS')
    idx = content.find('_sort_var.get()')
    print(repr(content[idx-10:idx+400]))
