import pandas as pd, sys
sys.stdout.reconfigure(encoding='utf-8')

df_ex = pd.read_excel(r'C:\Users\Abfuellung 15\Downloads\vergleich_dhl_pdfs_vs_orca.xlsx', sheet_name='DHL Express', skiprows=1)
df_no = pd.read_excel(r'C:\Users\Abfuellung 15\Downloads\vergleich_dhl_pdfs_vs_orca.xlsx', sheet_name='DHL Normal', skiprows=1)

print("EXPRESS: letzte 20 Tage")
for _, r in df_ex.tail(20).iterrows():
    d = str(r["Datum"])[:10]
    print(f"  {d}  PDF={int(r['PDF-Anzahl']):4}  Orca={int(r['OrcaScan']):4}  Diff={int(r['Differenz']):+4}")

print()
print("NORMAL: letzte 20 Tage")
for _, r in df_no.tail(20).iterrows():
    d = str(r["Datum"])[:10]
    print(f"  {d}  PDF={int(r['PDF-Anzahl']):4}  Orca={int(r['OrcaScan']):4}  Diff={int(r['Differenz']):+4}")

print()
print(f"EXPRESS Summe: PDF={int(df_ex['PDF-Anzahl'].sum())}  Orca={int(df_ex['OrcaScan'].sum())}")
print(f"NORMAL  Summe: PDF={int(df_no['PDF-Anzahl'].sum())}  Orca={int(df_no['OrcaScan'].sum())}")
