"""
Gimli – DHL-Label-PDF-Auswertung
Liest alle PDFs aus dem Datensicherungs-Ordner und gibt
Tages-Summen für DHL Normal und DHL Express aus.
"""

import os
import re
import sys
from datetime import datetime
from concurrent.futures import ProcessPoolExecutor, as_completed

sys.stdout.reconfigure(encoding="utf-8")

PDF_FOLDER = r"W:\Dokumentenaustausch\Tagesskripte\Datensicherung"

PAT_EXPRESS = re.compile(r'\(J\)\s+(JD\d{2}\s+[\d\s]{15,})')
PAT_NORMAL  = re.compile(r'(?:Sendungsnr\.|Sendungsnummer)\s*[:\n]?\s*\(00\)\s*(\d{18,20})')


def process_pdf(path):
    mtime   = os.path.getmtime(path)
    scan_dt = datetime.fromtimestamp(mtime).date()
    try:
        import pdfplumber
        with pdfplumber.open(path) as pdf:
            text = "".join(p.extract_text() or "" for p in pdf.pages)
        m_ex = PAT_EXPRESS.search(text)
        if m_ex:
            bc = "J" + m_ex.group(1).replace(" ", "")
            return ("express", bc, scan_dt)
        m_no = PAT_NORMAL.search(text)
        if m_no:
            bc = "00" + m_no.group(1)
            return ("normal", bc, scan_dt)
    except Exception:
        pass
    return (None, None, None)


def load_pdfs():
    files = [os.path.join(PDF_FOLDER, f)
             for f in os.listdir(PDF_FOLDER) if f.lower().endswith(".pdf")]
    print(f"{len(files)} PDFs gefunden, lese aus …")

    express_rows, normal_rows = {}, {}
    done = 0
    with ProcessPoolExecutor(max_workers=8) as ex:
        futures = {ex.submit(process_pdf, p): p for p in files}
        for fut in as_completed(futures):
            typ, bc, dt = fut.result()
            done += 1
            if typ == "express":
                express_rows.setdefault(dt, set()).add(bc)
            elif typ == "normal":
                normal_rows.setdefault(dt, set()).add(bc)
            if done % 500 == 0:
                print(f"  {done}/{len(files)} PDFs verarbeitet …")

    return express_rows, normal_rows


if __name__ == "__main__":
    express_by_day, normal_by_day = load_pdfs()

    alle_tage = sorted(set(express_by_day) | set(normal_by_day))

    print("\n=== Tages-Summen (alle Tage) ===")
    print(f"{'Datum':<12}  {'Normal':>8}  {'Express':>9}")
    print("-" * 35)
    for tag in alle_tage:
        n = len(normal_by_day.get(tag, set()))
        e = len(express_by_day.get(tag, set()))
        print(f"{str(tag):<12}  {n:>8}  {e:>9}")

    # April-Summe gesondert
    april_normal  = sum(len(v) for k, v in normal_by_day.items()  if k.month == 4)
    april_express = sum(len(v) for k, v in express_by_day.items() if k.month == 4)
    print(f"\n=== April-Gesamt ===")
    print(f"  DHL Normal:  {april_normal}")
    print(f"  DHL Express: {april_express}")
    print(f"  Gesamt:      {april_normal + april_express}")
