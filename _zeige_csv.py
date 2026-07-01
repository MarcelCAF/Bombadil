import pandas as pd, sys
sys.stdout.reconfigure(encoding='utf-8')

df = pd.read_csv(
    r'C:\Users\Abfuellung 15\Downloads\CAF Priority List(Report Versand).csv',
    encoding='latin-1', sep=';'
)
print(f"Zeilen: {len(df)} | Spalten: {len(df.columns)}")
print(f"Erste Spalte: {df.columns[0]}")
print(f"Letzte Spalte: {df.columns[-1]}")
print()
print("Erste 10 Zeilen (erste 6 Spalten):")
print(df.iloc[:10, :6].to_string())
print()
print("Zeilennamen (erste Spalte):")
for v in df.iloc[:, 0].dropna().unique()[:30]:
    print(f"  {v}")
