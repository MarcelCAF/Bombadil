import sys
sys.stdout.reconfigure(encoding='utf-8')
path = r'C:\Users\Abfuellung 15\Documents\Marcels Skripts\Bombadil\Bombadil.py'
with open(path, 'r', encoding='utf-8') as f:
    content = f.read()

errors = []

# 1. extra_bindings: column_header_select durch direktes Canvas-Binding ersetzen
old1 = (
    '            try:\n'
    '                self._sheet.extra_bindings([\n'
    '                    ("cell_select", self._on_cell_click),\n'
    '                    ("column_header_select", self._on_header_click),\n'
    '                ])\n'
    '            except Exception:\n'
    '                pass\n'
)
new1 = (
    '            try:\n'
    '                self._sheet.extra_bindings([("cell_select", self._on_cell_click)])\n'
    '            except Exception:\n'
    '                pass\n'
    '            try:\n'
    '                self._sheet.MT.CH.bind("<ButtonRelease-1>", self._on_header_click)\n'
    '            except Exception:\n'
    '                pass\n'
)
if old1 in content:
    content = content.replace(old1, new1, 1); print('Binding: OK')
else:
    errors.append('Binding')

# 2. _on_header_click: robuste Spalten-Erkennung via x-Position
old2 = (
    '    def _on_header_click(self, event):\n'
    '        """Spaltenheader klicken: 1x aufsteigend, 2x absteigend, 3x Standard."""\n'
    '        try:\n'
    '            col = event.column if hasattr(event, "column") else event[1]\n'
    '        except Exception:\n'
    '            return\n'
    '        _sortable = {0, 1, 2, 3, 5, 6, 7}\n'
    '        if col not in _sortable:\n'
    '            return\n'
    '        if self._sort_col == col:\n'
    '            self._sort_dir = (self._sort_dir + 1) % 3\n'
    '            if self._sort_dir == 0:\n'
    '                self._sort_col = None\n'
    '        else:\n'
    '            self._sort_col = col\n'
    '            self._sort_dir = 1\n'
    '        self._refresh_ui()\n'
)
new2 = (
    '    def _on_header_click(self, event):\n'
    '        """Spaltenheader klicken: 1x aufsteigend, 2x absteigend, 3x Standard."""\n'
    '        try:\n'
    '            col = self._sheet.MT.CH.find_col(event.x)\n'
    '        except Exception:\n'
    '            try:\n'
    '                col = self._sheet.identify_col(x=event.x)\n'
    '            except Exception:\n'
    '                return\n'
    '        if col is None:\n'
    '            return\n'
    '        _sortable = {0, 1, 2, 3, 5, 6, 7}\n'
    '        if col not in _sortable:\n'
    '            return\n'
    '        if self._sort_col == col:\n'
    '            self._sort_dir = (self._sort_dir + 1) % 3\n'
    '            if self._sort_dir == 0:\n'
    '                self._sort_col = None\n'
    '        else:\n'
    '            self._sort_col = col\n'
    '            self._sort_dir = 1\n'
    '        self._refresh_ui()\n'
)
if old2 in content:
    content = content.replace(old2, new2, 1); print('Header-Handler: OK')
else:
    errors.append('Header-Handler')

# 3. Warndreieck-Spalte entfernen
old3 = (
    '        def _is_fehler(r):\n'
    '            return (r.get("tour") in ("T1", "T2")\n'
    '                    and not r["_abholbereit_bool"]\n'
    '                    and r.get("db_status") not in ("abgeholt", "retoure"))\n'
    '\n'
    '        data = [\n'
    '            [r.get("tour", ""), r["barcode"], r["name"], r["tb_status"], r["in_db"],\n'
    '             r["verpackt_at"], r["abholbereit_at"], r["zielkiosk"],\n'
    '             "\\u26a0" if _is_fehler(r) else ""]\n'
    '            for r in rows\n'
    '        ]\n'
)
new3 = (
    '        data = [\n'
    '            [r.get("tour", ""), r["barcode"], r["name"], r["tb_status"], r["in_db"],\n'
    '             r["verpackt_at"], r["abholbereit_at"], r["zielkiosk"]]\n'
    '            for r in rows\n'
    '        ]\n'
)
if old3 in content:
    content = content.replace(old3, new3, 1); print('Warndreieck: OK')
else:
    errors.append('Warndreieck')
    idx = content.find('_is_fehler')
    print('MISS:', repr(content[idx-10:idx+200]))

if not errors:
    with open(path, 'w', encoding='utf-8') as f:
        f.write(content)
    print('Datei gespeichert.')
else:
    print('FEHLER:', errors)
