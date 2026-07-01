import sys
sys.stdout.reconfigure(encoding='utf-8')
path = r'C:\Users\Abfuellung 15\Documents\Marcels Skripts\Bombadil\Bombadil.py'
with open(path, 'r', encoding='utf-8') as f:
    content = f.read()

old = (
    '        tour_nr = tour.replace("T", "")\n'
    '        filename = f"Orca_Tour{tour_nr}_{heute}.xlsx"\n'
)
new = (
    '        tour_suffix = "A" if tour == "T1" else "B"\n'
    '        filename = f"Orca_Abholer_{heute}{tour_suffix}.xlsx"\n'
)

if old in content:
    content = content.replace(old, new, 1)
    with open(path, 'w', encoding='utf-8') as f:
        f.write(content)
    print('OK')
else:
    print('MISS')
    idx = content.find('filename = f"Orca')
    print(repr(content[idx-80:idx+60]))
