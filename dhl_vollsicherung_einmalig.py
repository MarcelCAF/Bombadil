"""
DHL Vollsicherung – EINMALIGES Skript
======================================
Lädt alle DHL Normal + Express Einträge aus OrcaScan und
speichert sie als Archiv-Datei auf Google Drive.

Nach erfolgreichem Durchlauf kann dieses Skript gelöscht werden.
"""

import sys
import os
sys.path.insert(0, os.path.dirname(__file__))

# Bombadil-Funktionen laden
from Bombadil import backup_dhl_to_gdrive

print("=" * 50)
print("DHL Vollsicherung – Einmaliger Export")
print("=" * 50)
print()
print("Lade DHL Normal + Express aus OrcaScan ...")
print("(Das kann bei 164k Einträgen einige Minuten dauern)")
print()

results = backup_dhl_to_gdrive()

print()
print("Ergebnis:")
for key, val in results.items():
    print(f"  {key}: {val}")

print()
print("Fertig! Dieses Skript kann jetzt gelöscht werden.")
input("Enter drücken zum Beenden ...")
