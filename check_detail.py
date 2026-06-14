import pandas as pd
df = pd.read_excel('Planning_Tomate_2026.xlsx', sheet_name='Planning Journalier')

# Verifier ligne 25/06 AMOR KHECHIN avec details vehicules
amor = df[(df['Agriculteur']=='AMOR KHECHIN (PL-ELFALLEH)')].copy()
amor['Date'] = pd.to_datetime(amor['Date'])
amor = amor.sort_values('Date').reset_index(drop=True)
print('=== Detail complet AMOR KHECHIN (PL-ELFALLEH) ===')
print()
for i, r in amor.iterrows():
    d = r['Date'].strftime('%a %d/%m')
    print('  ' + d + ': Tonnes=' + str(r['Tonnes/Jour']) + 't  | Voyages=' + str(r['Nb Voyages']) + '  | Vehicules=' + str(r['Vehicules']))
print()
total = amor['Tonnes/Jour'].sum()
print('TOTAL Excel:', total, 't (declare 408t)')
print('Difference:', total - 408, 't')

# Verifier si autres agriculteurs ELFALLEH le 25/06
print()
print('=== Tout ELFALLEH le 25/06 ===')
d25 = df[df['Date'].astype(str).str.contains('2026-06-25')]
e25 = d25[d25['Usine']=='ELFALLEH']
for _, r in e25.iterrows():
    print('  ' + r['Agriculteur'] + ': ' + str(r['Tonnes/Jour']) + 't | ' + str(r['Vehicules']))
