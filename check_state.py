import sys
sys.stdout.reconfigure(encoding='utf-8')
path = r'C:\Users\Abfuellung 15\Documents\Marcels Skripts\Bombadil\Bombadil.py'
with open(path, 'r', encoding='utf-8') as f:
    content = f.read()

# Zeige _on_header_click Methode
idx = content.find('def _on_header_click')
print(repr(content[idx:idx+800]))

# Zeige wo Sortierung: vorkommt
idx2 = 0
while True:
    idx2 = content.find('Sortierung:', idx2)
    if idx2 < 0: break
    print('Sortierung: bei', idx2, repr(content[idx2-30:idx2+60]))
    idx2 += 1
