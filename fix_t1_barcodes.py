import sys
sys.stdout.reconfigure(encoding='utf-8')
path = r'C:\Users\Abfuellung 15\Documents\Marcels Skripts\Bombadil\Bombadil.py'
with open(path, 'r', encoding='utf-8') as f:
    content = f.read()

old = '        t1_barcodes = [r["barcode"] for r in self._all_rows]\n'
new = '        t1_barcodes = [r["barcode"] for r in self._all_rows if r["tb_status"] == "Verpackt"]\n'

if old in content:
    content = content.replace(old, new, 1)
    with open(path, 'w', encoding='utf-8') as f:
        f.write(content)
    print('OK')
else:
    print('MISS')
    idx = content.find('t1_barcodes')
    print(repr(content[idx-10:idx+80]))
