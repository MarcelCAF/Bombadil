import sys
sys.stdout.reconfigure(encoding='utf-8')
path = r'C:\Users\Abfuellung 15\Documents\Marcels Skripts\Bombadil\Bombadil.py'
with open(path, 'r', encoding='utf-8') as f:
    content = f.read()

old = (
    '            try:\n'
    '                self._sheet.extra_bindings([\n'
    '                    ("cell_select", self._on_cell_click),\n'
    '                    ("column_header_select", self._on_header_click),\n'
    '                ])\n'
    '            except Exception:\n'
    '                pass\n'
)
new = (
    '            try:\n'
    '                self._sheet.extra_bindings([("cell_select", self._on_cell_click)])\n'
    '            except Exception:\n'
    '                pass\n'
    '            try:\n'
    '                self._sheet.MT.CH.bind("<ButtonRelease-1>", self._on_header_click)\n'
    '            except Exception:\n'
    '                pass\n'
)

if old in content:
    content = content.replace(old, new, 1)
    with open(path, 'w', encoding='utf-8') as f:
        f.write(content)
    print('OK')
else:
    print('MISS')
    idx = content.find('column_header_select')
    print(repr(content[idx-80:idx+100]))
