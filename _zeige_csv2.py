import pandas as pd, sys
sys.stdout.reconfigure(encoding='utf-8')

df = pd.read_csv(
    r'C:\Users\Abfuellung 15\Downloads\CAF Priority List(Report Versand).csv',
    encoding='latin-1', sep=';', index_col=0
)

df = df.drop(columns=[c for c in df.columns if 'Unnamed' in str(c)], errors='ignore')
tages_summen = df.apply(pd.to_numeric, errors='coerce').sum(axis=0)
tages_summen.index = pd.to_datetime(tages_summen.index, format='%d.%m.%Y', errors='coerce')
tages_summen = tages_summen[tages_summen.index.notna()]
tages_summen = tages_summen[tages_summen > 0].sort_index()

print(f"Zeitraum: {tages_summen.index.min().date()} bis {tages_summen.index.max().date()}")
print(f"Tage mit Daten: {len(tages_summen)}")
print()
print("Letzte 20 Tage:")
for d, n in tages_summen.tail(20).items():
    print(f"  {d.strftime('%d.%m.%Y')}  {int(n):4} Pakete")
