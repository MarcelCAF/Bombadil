import sys
sys.stdout.reconfigure(encoding='utf-8')
path = r'C:\Users\Abfuellung 15\Documents\Marcels Skripts\Bombadil\Bombadil.py'
with open(path, 'r', encoding='utf-8') as f:
    content = f.read()

errors = []

# 1. Header-Canvas-Binding nach dem Sheet-Setup einfuegen
old1 = (
    '            try:\n'
    '                self._sheet.extra_bindings([("cell_select", self._on_cell_click)])\n'
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
    content = content.replace(old1, new1, 1); print('Canvas-Binding: OK')
else:
    errors.append('Canvas-Binding')

# 2. _on_header_click Methode vor _on_cell_click einfuegen
old2 = '    def _on_cell_click(self, event):\n'
new2 = (
    '    def _on_header_click(self, event):\n'
    '        """Spaltenheader klicken: 1x aufsteigend, 2x absteigend, 3x Standard."""\n'
    '        try:\n'
    '            actual_x = self._sheet.MT.CH.canvasx(event.x)\n'
    '            cumx = 0\n'
    '            col = None\n'
    '            for i in range(len(self._COL_HEADERS)):\n'
    '                try:\n'
    '                    w = self._sheet.column_width(column=i)\n'
    '                except Exception:\n'
    '                    w = self._COL_WIDTHS[i] if i < len(self._COL_WIDTHS) else 80\n'
    '                cumx += w\n'
    '                if actual_x <= cumx:\n'
    '                    col = i\n'
    '                    break\n'
    '        except Exception:\n'
    '            return\n'
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
    '\n'
    '    def _on_cell_click(self, event):\n'
)
if old2 in content:
    content = content.replace(old2, new2, 1); print('Header-Methode: OK')
else:
    errors.append('Header-Methode')

# 3. Warndreieck-Logik und Spalte entfernen
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
    if idx >= 0: print(repr(content[idx-10:idx+200]))

if not errors:
    with open(path, 'w', encoding='utf-8') as f:
        f.write(content)
    print('Datei gespeichert.')
else:
    print('FEHLER:', errors)
