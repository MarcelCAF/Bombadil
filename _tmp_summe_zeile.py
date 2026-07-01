# Temporär: 'Summe'-Zeile (70) der CSV für Mai auswerten + Vergleich – kann danach gelöscht werden
import csv
import datetime

path = r"C:\Users\Abfuellung 15\Downloads\CAF Priority List(Report Versand).csv"
rows = list(csv.reader(open(path, encoding="latin-1"), delimiter=";"))
dates = rows[0][2:]

# Zeile mit exakt 'Summe' als Name
summe_row = None
for r in rows:
    if r and r[0].strip() == "Summe":
        summe_row = r
        break
if summe_row is None:
    raise SystemExit("Summe-Zeile nicht gefunden!")

summe = {}
for d, v in zip(dates, summe_row[2:]):
    v = v.strip()
    if v.isdigit():
        try:
            dt = datetime.datetime.strptime(d, "%d.%m.%Y").date()
        except ValueError:
            continue
        if dt.year == 2026 and dt.month == 5:
            summe[dt.day] = int(v)

print("Mai-Summe der 'Summe'-Zeile:", sum(summe.values()))
print()

# Bombadil real (aus vorheriger Analyse)
real_n = {2: 35, 4: 508, 5: 477, 6: 633, 7: 492, 8: 384, 9: 324,
          11: 391, 12: 462, 13: 477, 15: 645, 16: 336, 18: 357,
          19: 448, 20: 561, 21: 490, 22: 603, 23: 247, 26: 438, 27: 476,
          28: 581, 29: 480, 30: 779}
real_e = {2: 341, 4: 348, 5: 594, 6: 316, 7: 417, 8: 498, 9: 265,
          11: 537, 12: 626, 13: 463, 15: 572, 16: 390, 18: 661,
          19: 683, 20: 671, 21: 260, 22: 390, 23: 226, 26: 818, 27: 743,
          28: 737, 29: 860, 30: 209}
alte_phantome = {2: 200, 4: 170, 5: 225, 6: 162, 7: 202, 8: 190, 9: 121,
                 11: 208, 12: 237, 13: 192, 15: 266, 16: 147, 18: 212,
                 19: 275, 20: 210, 21: 280, 22: 229, 23: 113, 26: 277, 27: 206}

alle_tage = sorted(set(list(real_n.keys()) + list(summe.keys())))
print(f"{'Tag':<12}{'Summe-Zeile':>12}{'JETZT':>9}{'NEU':>9}{'Diff':>8}")
print("-" * 50)
s_csv = s_jetzt = s_neu = 0
neue_phantome = {}
for d in alle_tage:
    dt = datetime.date(2026, 5, d)
    wd = ["Mo", "Di", "Mi", "Do", "Fr", "Sa", "So"][dt.weekday()]
    real = real_n.get(d, 0) + real_e.get(d, 0)
    jetzt = real + alte_phantome.get(d, 0)
    ziel = summe.get(d)
    if ziel:
        neu = max(real, ziel)            # nie echte Daten verstecken
        neue_phantome[d] = max(0, ziel - real)
    else:
        neu = real                        # CSV leer -> echte Werte behalten
    s_csv += ziel or 0
    s_jetzt += jetzt
    s_neu += neu
    ziel_s = str(ziel) if ziel else "(leer)"
    print(f"{d:02d}.05. {wd:<5}{ziel_s:>12}{jetzt:>9}{neu:>9}{neu - jetzt:>8}")
print("-" * 50)
print(f"{'SUMME':<12}{s_csv:>12}{s_jetzt:>9}{s_neu:>9}{s_neu - s_jetzt:>8}")
print()
print("Neue Phantom-Verteilung (Summe", sum(neue_phantome.values()), "):")
print("{" + ", ".join(f"{d}: {n}" for d, n in sorted(neue_phantome.items()) if n > 0) + "}")
