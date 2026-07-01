# Temporäres Analyse-Skript: Phantom-Verteilung aus CSV berechnen (kann danach gelöscht werden)
import csv
import datetime

path = r"C:\Users\Abfuellung 15\Downloads\CAF Priority List(Report Versand).csv"
rows = list(csv.reader(open(path, encoding="latin-1"), delimiter=";"))
dates = rows[0][2:]
totals = {}
for r in rows[1:]:
    for d, v in zip(dates, r[2:]):
        v = v.strip()
        if v.isdigit():
            totals[d] = totals.get(d, 0) + int(v)

mai = {}
for d, n in totals.items():
    try:
        dt = datetime.datetime.strptime(d, "%d.%m.%Y").date()
    except ValueError:
        continue
    if dt.year == 2026 and dt.month == 5 and n > 0:
        mai[dt] = n

summe = sum(mai.values())
ziel = 4122

# Größter-Rest-Verfahren: Ganzzahlen, Summe exakt = ziel
anteile = {dt: ziel * n / summe for dt, n in mai.items()}
basis = {dt: int(a) for dt, a in anteile.items()}
rest = ziel - sum(basis.values())
reste = sorted(anteile.items(), key=lambda kv: kv[1] - int(kv[1]), reverse=True)
for i in range(rest):
    basis[reste[i][0]] += 1

print("Kontrolle Summe:", sum(basis.values()))
print()
print("Python-Dict (Tag -> Anzahl):")
print("{" + ", ".join(f"{dt.day}: {basis[dt]}" for dt in sorted(basis)) + "}")
print()
for dt in sorted(basis):
    wd = ["Mo", "Di", "Mi", "Do", "Fr", "Sa", "So"][dt.weekday()]
    print(" ", dt, wd, "->", basis[dt])
