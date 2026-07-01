"""
Liest alle Versandlabel-PDFs aus dem Backup-Ordner,
extrahiert die DHL-Express-Sendungsnummer (JD01...) und
das Dateidatum, und schreibt eine Excel-Tabelle.
"""

import os
import re
import pdfplumber
from datetime import datetime
from concurrent.futures import ProcessPoolExecutor, as_completed
import pandas as pd

PDF_FOLDER = r"W:\Dokumentenaustausch\Tagesskripte\Datensicherung"
OUTPUT_FILE = r"C:\Users\Abfuellung 15\Downloads\dhl_express_sendungen.xlsx"

# Muster: (J) JD01 4600 0126 2813 8069
PATTERN = re.compile(r'\(J\)\s+(JD\d{2}\s+[\d\s]{15,})')


def process_pdf(path):
    filename = os.path.basename(path)
    mtime = os.path.getmtime(path)
    scan_dt = datetime.fromtimestamp(mtime)
    try:
        with pdfplumber.open(path) as pdf:
            text = ""
            for page in pdf.pages:
                t = page.extract_text()
                if t:
                    text += t
        m = PATTERN.search(text)
        if m:
            raw = m.group(1).replace(" ", "")  # "JD014600012628138069"
            barcode = "J" + raw               # "JJD014600012628138069"
            return barcode, scan_dt
    except Exception:
        pass
    return None, None


def main():
    pdf_files = [
        os.path.join(PDF_FOLDER, f)
        for f in os.listdir(PDF_FOLDER)
        if f.lower().endswith(".pdf")
    ]
    print(f"{len(pdf_files)} PDFs gefunden, starte Verarbeitung...")

    rows = []
    done = 0
    skipped = 0

    # Parallele Verarbeitung (8 Prozesse)
    with ProcessPoolExecutor(max_workers=8) as ex:
        futures = {ex.submit(process_pdf, p): p for p in pdf_files}
        for fut in as_completed(futures):
            barcode, dt = fut.result()
            done += 1
            if barcode:
                rows.append({"PACKAGE BARCODE": barcode, "DATE OF SCAN": dt})
            else:
                skipped += 1
            if done % 500 == 0:
                print(f"  {done}/{len(pdf_files)} verarbeitet, {len(rows)} gefunden ...")

    print(f"\nFertig: {len(rows)} Express-Sendungen gefunden, {skipped} PDFs ohne Treffer.")

    df = pd.DataFrame(rows, columns=["PACKAGE BARCODE", "DATE OF SCAN"])
    df.sort_values("DATE OF SCAN", inplace=True)
    df.to_excel(OUTPUT_FILE, index=False)
    print(f"Excel gespeichert: {OUTPUT_FILE}")


if __name__ == "__main__":
    main()
