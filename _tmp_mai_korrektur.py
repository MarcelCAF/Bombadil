# Temporär: exakte Tages-Korrektur Mai berechnen (CSV_ges − real) – kann danach gelöscht werden
csv_ges = {2: 2000, 4: 1694, 5: 2246, 6: 1618, 7: 2062, 8: 1900, 9: 1210,
           11: 2082, 12: 2366, 13: 1918, 15: 2656, 16: 1464, 18: 2122,
           19: 2746, 20: 2094, 21: 2792, 22: 2286, 23: 1126, 26: 2762, 27: 2058}
real_n = {2: 35, 4: 508, 5: 477, 6: 633, 7: 492, 8: 384, 9: 324,
          11: 391, 12: 462, 13: 477, 15: 645, 16: 336, 18: 357,
          19: 448, 20: 561, 21: 490, 22: 603, 23: 247, 26: 438, 27: 476}
real_e = {2: 341, 4: 348, 5: 594, 6: 316, 7: 417, 8: 498, 9: 265,
          11: 537, 12: 626, 13: 463, 15: 572, 16: 390, 18: 661,
          19: 683, 20: 671, 21: 260, 22: 390, 23: 226, 26: 818, 27: 743}

korr = {}
for d in sorted(csv_ges):
    diff = csv_ges[d] - real_n[d] - real_e[d]
    korr[d] = max(0, diff)
    if diff < 0:
        print(f"ACHTUNG Tag {d}: real > CSV (diff {diff})")

print("Summe Phantome:", sum(korr.values()))
print()
print("{" + ", ".join(f"{d}: {korr[d]}" for d in sorted(korr)) + "}")
