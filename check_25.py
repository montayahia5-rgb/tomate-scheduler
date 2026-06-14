# Voir si AMOR KHECHIN (PL-ELFALLEH) a une livraison RM le 25/06
# qui se cumulerait avec une livraison OR-Tools
import pandas as pd
df = pd.read_excel('Planning_Tomate_2026.xlsx', sheet_name='Planning Journalier')

# Toutes les livraisons le 25/06 pour AMOR KHECHIN
d = df[df['Date'].astype(str).str.contains('2026-06-25')]
amor25 = d[d['Agriculteur'].str.contains('AMOR KHECHIN', case=False, na=False)]
print('=== AMOR KHECHIN livraisons le 25/06 (toutes usines) ===')
for _, r in amor25.iterrows():
    print('  ' + r['Agriculteur'] + ' | ' + r['Usine'] + ' | ' + str(r['Tonnes/Jour']) + 't | Note: ' + str(r['Note']))
print()
print('Cap ELFALLEH = 150t/jour')
print('Si AMOR a 200t le 25/06 ET d autres agriculteurs livrent → cap depasse')
print()
# Total ELFALLEH le 25/06
e25 = d[d['Usine']=='ELFALLEH']
print('Total ELFALLEH 25/06:', e25['Tonnes/Jour'].sum(), 't (cap=150)')
