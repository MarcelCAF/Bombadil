# Temporär: Vergleichsübersicht Mai (alt vs. neu) – kann danach gelöscht werden
import datetime

csv_ges = {2: 2000, 4: 1694, 5: 2246, 6: 1618, 7: 2062, 8: 1900, 9: 1210,
           11: 2082, 12: 2366, 13: 1918, 15: 2656, 16: 1464, 18: 2122,
           19: 2746, 20: 2094, 21: 2792, 22: 2286, 23: 1126, 26: 2762, 27: 2058}
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

print(f"{'Tag':<14}{'CSV-Ziel':>9}{'JETZT':>9}{'NEU':>9}{'Diff':>8}")
print("-" * 49)
s_csv = s_jetzt = s_neu = 0
for d in sorted(real_n):
    dt = datetime.date(2026, 5, d)
    wd = ["Mo", "Di", "Mi", "Do", "Fr", "Sa", "So"][dt.weekday()]
    real = real_n[d] + real_e[d]
    jetzt = real + alte_phantome.get(d, 0)
    csv_v = csv_ges.get(d)
    neu = csv_v if csv_v else real          # CSV-Tage = CSV; sonst real
    s_csv += csv_v or 0
    s_jetzt += jetzt
    s_neu += neu
    csv_s = str(csv_v) if csv_v else "(leer)"
    print(f"{d:02d}.05. {wd:<6}{csv_s:>9}{jetzt:>9}{neu:>9}{neu - jetzt:>8}")
print("-" * 49)
print(f"{'SUMME':<14}{s_csv:>9}{s_jetzt:>9}{s_neu:>9}{s_neu - s_jetzt:>8}")
