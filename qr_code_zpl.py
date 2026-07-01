"""
Legolas – QR-Code Drucker
Links in die Konsole einfügen (einen pro Zeile), zweimal Enter = drucken.
Fenster bleibt offen für mehrere Runden.
"""

import re
import sys
from pathlib import Path

# ── KONFIGURATION ─────────────────────────────────────────────────────────────
DRUCKER = "ZDesigner ZD421-203dpi ZPL"   # Name des Zebra-Druckers (wie in Gandalf)
# ─────────────────────────────────────────────────────────────────────────────


def _zpl_fuer_url(url: str) -> bytes:
    """Erstellt ZPL für ein 102×38mm Etikett mit QR-Code + URL-Text."""
    # URL kürzen für Textzeile (max. 60 Zeichen)
    url_text = url if len(url) <= 60 else url[:57] + "..."
    zpl = (
        "^XA"
        "^CI28"
        "^PW816"
        "^LL304"
        # QR-Code zentriert auf dem Etikett
        f"^FO326,70^BQN,2,5^FDQA,{url}^FS"
        "^XZ"
    )
    return zpl.encode("utf-8")


def drucken(urls: list[str]) -> None:
    try:
        import win32print
    except ImportError:
        print("  FEHLER: pywin32 nicht installiert.")
        print("  Bitte ausführen: pip install pywin32")
        return

    for url in urls:
        url = url.strip()
        if not url:
            continue
        zpl = _zpl_fuer_url(url)
        try:
            h = win32print.OpenPrinter(DRUCKER)
            try:
                win32print.StartDocPrinter(h, 1, ("Legolas QR", None, "RAW"))
                win32print.StartPagePrinter(h)
                win32print.WritePrinter(h, zpl)
                win32print.EndPagePrinter(h)
                win32print.EndDocPrinter(h)
            finally:
                win32print.ClosePrinter(h)
            print(f"  OK  {url}")
        except Exception as e:
            print(f"  FEHLER bei {url}: {e}")


def main():
    print("=" * 60)
    print("  Legolas – QR-Code Drucker")
    print(f"  Drucker: {DRUCKER}")
    print("=" * 60)

    while True:
        print("\nLinks einfügen (einen pro Zeile). Zweimal Enter = drucken.")
        print("Zum Beenden: 'q' eingeben + Enter.\n")

        urls = []
        while True:
            try:
                line = input()
            except EOFError:
                break
            if line.strip().lower() == "q":
                print("Tschüss!")
                sys.exit(0)
            if line.strip() == "":
                break
            urls.append(line.strip())

        if not urls:
            print("Keine Links eingegeben.")
            continue

        print(f"\nDrucke {len(urls)} QR-Code(s)...\n")
        drucken(urls)
        print("\nFertig! Nächste Runde oder 'q' zum Beenden.")


if __name__ == "__main__":
    main()
