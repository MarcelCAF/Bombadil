"""
json_in_unterordner.py
=====================
Verschiebt die vielen tour_zeiten_*.json + tagesbote_cache_*.json aus dem
Bombadil-Hauptordner in Unterordner, damit der Hauptordner schlank wird.

  tour_zeiten_*.json     → tour_zeiten/
  tagesbote_cache_*.json → tagesbote_cache/

Der Bombadil-Code (ab v1.0.80) liest aus BEIDEN Orten, das Verschieben ist
also gefahrlos. Laufzeit-/Auth-JSONs (settings, token, statistik_cache,
abholer_ts_cache, oauth_credentials, service_account) bleiben unangetastet.

DRY_RUN = True → nur zeigen was verschoben würde.
ZIEL: Pfad anpassen (NAS oder lokal).
"""

import sys, shutil
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8")

DRY_RUN = True

# Welcher Bombadil-Ordner soll aufgeräumt werden?
ZIEL = Path(r"W:\Dokumentenaustausch\Tagesskripte\Bombadil")   # NAS-Installation

# (Quell-Glob, Quell-Basis relativ zu ZIEL, Ziel-Unterordner)
JOBS = [
    ("tour_zeiten_*.json",       ".",       "tour_zeiten"),
    ("tagesbote_cache_*.json",   ".",       "tagesbote_cache"),
    ("Bombadil.backup_*.py",     ".",       "_alte_versionen"),   # Hauptordner-Backups
    ("Bombadil.backup_*.py",     "Archiv",  "_alte_versionen"),   # alte Backups im Archiv
]


def main():
    print("=" * 60)
    print(f"Bombadil-Ordner aufräumen  |  DRY_RUN: {DRY_RUN}")
    print(f"Ordner: {ZIEL}")
    print("=" * 60)
    if not ZIEL.exists():
        print(f"⚠  Ordner nicht erreichbar: {ZIEL}")
        return

    for muster, basis_rel, unterordner in JOBS:
        basis   = ZIEL if basis_rel == "." else ZIEL / basis_rel
        dst_dir = ZIEL / unterordner
        if not basis.exists():
            continue
        files = [f for f in sorted(basis.glob(muster)) if f.parent == basis]
        print(f"\n📂 {basis_rel}/{muster}  →  {unterordner}/   ({len(files)} Dateien)")
        if not files:
            print("   (nichts zu verschieben)")
            continue
        if not DRY_RUN:
            dst_dir.mkdir(parents=True, exist_ok=True)
        for f in files:
            ziel = dst_dir / f.name
            if DRY_RUN:
                print(f"   🔵 würde verschieben: {f.name}")
            else:
                if ziel.exists():
                    ziel.unlink()
                shutil.move(str(f), str(ziel))
                print(f"   ✅ verschoben: {f.name}")

    print("\n" + "=" * 60)
    print("Trockenlauf fertig." if DRY_RUN else "Verschieben abgeschlossen!")
    print("=" * 60)


if __name__ == "__main__":
    main()
