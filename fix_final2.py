import sys
sys.stdout.reconfigure(encoding='utf-8')
path = r'C:\Users\Abfuellung 15\Documents\Marcels Skripts\Bombadil\Bombadil.py'
with open(path, 'r', encoding='utf-8') as f:
    content = f.read()

errors = []

# 1. _on_header_click Methode vor _on_cell_click (PU heute, ab Zeile ~4165)
old1 = '\n    def _on_cell_click(self, event):\n        """Barcod'
new1 = (
    '\n'
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
    '        """Barcod'
)
# Nur den PU-heute _on_cell_click (nicht den anderen in TableTab)
idx = content.find('\n    def _on_cell_click(self, event):\n        """Barcod', 4000*4)
if idx >= 0:
    content = content[:idx] + new1[1:] + content[idx + len('\n    def _on_cell_click(self, event):\n        """Barcod'):]
    print('Header-Methode: OK')
else:
    errors.append('Header-Methode')
    print('MISS Header-Methode, idx=', content.find('def _on_cell_click'))

# 2. Warndreieck: _is_fehler entfernen und data-Liste bereinigen
idx2 = content.find('        def _is_fehler(r):\n')
if idx2 >= 0:
    # Suche den data = [...] Block danach
    data_start = content.find('        data = [\n', idx2)
    data_end = content.find('        ]\n', data_start) + len('        ]\n')
    old_block = content[idx2:data_end]
    new_block = (
        '        data = [\n'
        '            [r.get("tour", ""), r["barcode"], r["name"], r["tb_status"], r["in_db"],\n'
        '             r["verpackt_at"], r["abholbereit_at"], r["zielkiosk"]]\n'
        '            for r in rows\n'
        '        ]\n'
    )
    content = content[:idx2] + new_block + content[data_end:]
    print('Warndreieck: OK')
else:
    errors.append('Warndreieck')

with open(path, 'w', encoding='utf-8') as f:
    f.write(content)
print('Datei gespeichert.' if not errors else f'Mit Fehlern gespeichert: {errors}')
