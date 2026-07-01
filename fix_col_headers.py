import sys
sys.stdout.reconfigure(encoding='utf-8')
path = r'C:\Users\Abfuellung 15\Documents\Marcels Skripts\Bombadil\Bombadil.py'
with open(path, 'r', encoding='utf-8') as f:
    content = f.read()

old1 = '    _COL_HEADERS = ["Tour", "Paket-Barcode", "Name", "Packstatus", "In DB",\n                    "Verpackt", "Abholbereit", "Ziel-Kiosk", "\\u26a0"]\n'
new1 = '    _COL_HEADERS = ["Tour", "Paket-Barcode", "Name", "Packstatus", "In DB",\n                    "Verpackt", "Abholbereit", "Ziel-Kiosk"]\n'

old2 = '    _COL_WIDTHS  = [48, 220, 240, 110, 60, 155, 155, 120, 36]\n'
new2 = '    _COL_WIDTHS  = [48, 220, 240, 110, 60, 155, 155, 120]\n'

ok1 = old1 in content
ok2 = old2 in content

if ok1:
    content = content.replace(old1, new1, 1); print('COL_HEADERS: OK')
else:
    print('COL_HEADERS: MISS')
    idx = content.find('_COL_HEADERS')
    print(repr(content[idx:idx+120]))

if ok2:
    content = content.replace(old2, new2, 1); print('COL_WIDTHS: OK')
else:
    print('COL_WIDTHS: MISS')

if ok1 or ok2:
    with open(path, 'w', encoding='utf-8') as f:
        f.write(content)
    print('Gespeichert.')
