import sys
sys.stdout.reconfigure(encoding='utf-8')
path = r'C:\Users\Abfuellung 15\Documents\Marcels Skripts\Bombadil\Bombadil.py'
with open(path, 'r', encoding='utf-8') as f:
    content = f.read()

errors = []

# 1. enable_bindings: column_select hinzufuegen
old1 = '            self._sheet.enable_bindings("single_select", "row_select", "copy", "column_width_resize")\n'
new1 = '            self._sheet.enable_bindings("single_select", "row_select", "column_select", "copy", "column_width_resize")\n'
if old1 in content:
    content = content.replace(old1, new1, 1); print('enable_bindings: OK')
else:
    errors.append('enable_bindings')

# 2. MT.CH.bind durch sauberes extra_bindings ersetzen
old2 = (
    '            try:\n'
    '                self._sheet.extra_bindings([("cell_select", self._on_cell_click)])\n'
    '            except Exception:\n'
    '                pass\n'
    '            try:\n'
    '                self._sheet.MT.CH.bind("<ButtonRelease-1>", self._on_header_click)\n'
    '            except Exception:\n'
    '                pass\n'
)
new2 = (
    '            try:\n'
    '                self._sheet.extra_bindings([\n'
    '                    ("cell_select", self._on_cell_click),\n'
    '                    ("column_select", self._on_header_click),\n'
    '                ])\n'
    '            except Exception:\n'
    '                pass\n'
)
if old2 in content:
    content = content.replace(old2, new2, 1); print('Binding: OK')
else:
    errors.append('Binding')
    idx = content.find('MT.CH.bind')
    print('MISS Binding:', repr(content[idx-100:idx+100]))

# 3. _on_header_click: saubere Implementierung via event.selected.column
old3 = (
    '    def _on_header_click(self, event):\n'
    '        """Spaltenheader klicken: 1x aufsteigend, 2x absteigend, 3x Standard."""\n'
    '        col = None\n'
    '        # Methode 1: col_positions (tksheet intern)\n'
    '        try:\n'
    '            actual_x = self._sheet.MT.CH.canvasx(event.x)\n'
    '            positions = self._sheet.MT.col_positions\n'
    '            for i in range(len(positions) - 1):\n'
    '                if positions[i] <= actual_x < positions[i + 1]:\n'
    '                    col = i\n'
    '                    break\n'
    '        except Exception:\n'
    '            pass\n'
    '        # Methode 2: canvasx + fixe Breiten\n'
    '        if col is None:\n'
    '            try:\n'
    '                actual_x = self._sheet.MT.CH.canvasx(event.x)\n'
    '                cumx = 0\n'
    '                for i, w in enumerate(self._COL_WIDTHS):\n'
    '                    cumx += w\n'
    '                    if actual_x <= cumx:\n'
    '                        col = i\n'
    '                        break\n'
    '            except Exception:\n'
    '                pass\n'
    '        # Methode 3: rohe event.x + fixe Breiten\n'
    '        if col is None:\n'
    '            cumx = 0\n'
    '            for i, w in enumerate(self._COL_WIDTHS):\n'
    '                cumx += w\n'
    '                if event.x <= cumx:\n'
    '                    col = i\n'
    '                    break\n'
)
new3 = (
    '    def _on_header_click(self, event):\n'
    '        """Spaltenheader klicken: 1x aufsteigend, 2x absteigend, 3x Standard."""\n'
    '        try:\n'
    '            col = event.selected.column\n'
    '        except Exception:\n'
    '            return\n'
    '        if col is None:\n'
    '            return\n'
)
if old3 in content:
    content = content.replace(old3, new3, 1); print('Handler: OK')
else:
    errors.append('Handler')
    idx = content.find('def _on_header_click')
    print('MISS Handler:', repr(content[idx:idx+200]))

if not errors:
    with open(path, 'w', encoding='utf-8') as f:
        f.write(content)
    print('Datei gespeichert.')
else:
    print('FEHLER:', errors)
