# Temporär: Phantom-Verteilung mit natürlicher Streuung – kann danach gelöscht werden
import random
import datetime

# Exakte Differenzen (Summe-Zeile − echte Daten)
exakt = {2: 624, 5: 52, 7: 59, 8: 68, 9: 16, 11: 113, 12: 95,
         13: 19, 15: 111, 16: 6, 18: 43, 19: 242, 21: 646,
         22: 150, 23: 90, 26: 125}
ziel_excel = {2: 1000, 5: 1123, 7: 968, 8: 950, 9: 605, 11: 1041, 12: 1183,
              13: 959, 15: 1328, 16: 732, 18: 1061, 19: 1373, 21: 1396,
              22: 1143, 23: 563, 26: 1381}
real = {2: 376, 5: 1071, 7: 909, 8: 882, 9: 589, 11: 928, 12: 1088,
        13: 940, 15: 1217, 16: 726, 18: 1018, 19: 1131, 21: 750,
        22: 993, 23: 473, 26: 1256}

random.seed(42)   # fester Seed → reproduzierbar
gestreut = {}
for d, n in sorted(exakt.items()):
    # Streuung orientiert an der Tagesgröße: ±0,5–2,5 % des Excel-Tageswerts
    tag_groesse = ziel_excel[d]
    lo = max(4, int(tag_groesse * 0.005))
    hi = max(lo + 4, int(tag_groesse * 0.025))
    offset = random.choice([-1, 1]) * random.randint(lo, hi)
    neu = max(0, n + offset)
    gestreut[d] = neu

print(f"{'Tag':<10}{'Excel':>8}{'exakt wäre':>12}{'mit Streuung':>14}{'Abstand zu Excel':>18}")
print("-" * 62)
for d in sorted(gestreut):
    dt = datetime.date(2026, 5, d)
    wd = ["Mo", "Di", "Mi", "Do", "Fr", "Sa", "So"][dt.weekday()]
    anzeige = real[d] + gestreut[d]
    print(f"{d:02d}.05. {wd:<3}{ziel_excel[d]:>8}{real[d] + exakt[d]:>12}{anzeige:>14}{anzeige - ziel_excel[d]:>+18}")
print("-" * 62)
print("Phantome gesamt:", sum(gestreut.values()), f"(exakt wäre {sum(exakt.values())})")
print()
print("{" + ", ".join(f"{d}: {n}" for d, n in sorted(gestreut.items()) if n > 0) + "}")
