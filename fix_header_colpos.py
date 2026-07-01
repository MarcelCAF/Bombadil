import sys
sys.stdout.reconfigure(encoding='utf-8')
path = r'C:\Users\Abfuellung 15\Documents\Marcels Skripts\Bombadil\Bombadil.py'
with open(path, 'r', encoding='utf-8') as f:
    content = f.read()

old = (
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
)
new = (
    '    def _on_header_click(self, event):\n'
    '        """Spaltenheader klicken: 1x aufsteigend, 2x absteigend, 3x Standard."""\n'
    '        try:\n'
    '            actual_x = self._sheet.MT.CH.canvasx(event.x)\n'
    '            positions = self._sheet.MT.col_positions\n'
    '            col = None\n'
    '            for i in range(len(positions) - 1):\n'
    '                if positions[i] <= actual_x < positions[i + 1]:\n'
    '                    col = i\n'
    '                    break\n'
    '        except Exception:\n'
    '            return\n'
)

if old in content:
    content = content.replace(old, new, 1)
    with open(path, 'w', encoding='utf-8') as f:
        f.write(content)
    print('OK')
else:
    print('MISS')
    idx = content.find('def _on_header_click')
    print(repr(content[idx:idx+500]))
