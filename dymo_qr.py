"""
Dymo QR-Drucker – Link eingeben → QR-Code wird auf Dymo LabelWriter 450 gedruckt.
Etikettengröße: 36x89mm
"""

from pathlib import Path

# ─── KONFIGURATION ────────────────────────────────────────────────────────────
DRUCKER_NAME = "Dymo LabelWriter 450"   # Exakt wie in Windows Druckerverwaltung
ETIKETT_BREITE_MM  = 36
ETIKETT_HOEHE_MM   = 89
DPI = 300
# ──────────────────────────────────────────────────────────────────────────────


def mm_zu_px(mm: float) -> int:
    return round(mm * DPI / 25.4)


def drucke_qr(url: str) -> None:
    import qrcode
    import win32print
    import win32ui
    from PIL import Image, ImageWin

    # Drucker-DC öffnen um echte Pixelgröße abzufragen
    hdc = win32ui.CreateDC()
    hdc.CreatePrinterDC(DRUCKER_NAME)

    drucker_breite = hdc.GetDeviceCaps(8)    # HORZRES – echte Pixel Breite
    drucker_hoehe  = hdc.GetDeviceCaps(10)   # VERTRES  – echte Pixel Höhe

    # QR-Code erzeugen: quadratisch, begrenzt durch die schmalere Seite (Breite)
    rand_px    = round(drucker_breite * 0.05)   # 5% Rand
    qr_groesse = drucker_breite - 2 * rand_px   # bleibt quadratisch

    qr = qrcode.QRCode(border=1)
    qr.add_data(url)
    qr.make(fit=True)
    qr_img = qr.make_image(fill_color="black", back_color="white").convert("RGB")
    qr_img = qr_img.resize((qr_groesse, qr_groesse), Image.LANCZOS)

    # QR-Code auf weißem Etikett zentrieren (in echter Druckerpixelgröße)
    etikett = Image.new("RGB", (drucker_breite, drucker_hoehe), "white")
    x = (drucker_breite - qr_groesse) // 2
    y = (drucker_hoehe  - qr_groesse) // 2
    etikett.paste(qr_img, (x, y))

    # 1:1 drucken – keine Skalierung mehr
    hdc.StartDoc("Dymo QR")
    hdc.StartPage()
    dib = ImageWin.Dib(etikett)
    dib.draw(hdc.GetHandleOutput(), (0, 0, drucker_breite, drucker_hoehe))
    hdc.EndPage()
    hdc.EndDoc()
    hdc.DeleteDC()


def main():
    # Abhängigkeiten prüfen
    try:
        import qrcode
    except ImportError:
        print("Fehlende Bibliothek! Bitte einmalig ausführen:")
        print("  pip install qrcode[pil]")
        input("\nEnter drücken zum Beenden...")
        return

    try:
        import win32print
        import win32ui
    except ImportError:
        print("Fehlende Bibliothek! Bitte einmalig ausführen:")
        print("  pip install pywin32")
        input("\nEnter drücken zum Beenden...")
        return

    print(f"Dymo QR-Drucker  →  {DRUCKER_NAME}")
    print("─" * 45)

    while True:
        print("\nLink eingeben (oder 'q' zum Beenden):")
        url = input("  > ").strip()

        if url.lower() == "q" or url == "":
            break

        try:
            drucke_qr(url)
            print("  ✓ Gedruckt!")
        except Exception as e:
            print(f"  ✗ Fehler: {e}")


    print("\nBis dann!")
    input("Enter drücken zum Beenden...")


if __name__ == "__main__":
    main()
