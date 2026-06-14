import pandas as pd
df = pd.read_excel('Planning_Tomate_2026.xlsx', sheet_name='Planning Journalier')

# Cibler le lot precis AMOR KHECHIN ELFALLEH
sub = df[(df['Agriculteur'].str.contains('AMOR KHECHIN', case=False, na=False)) & 
         (df['Usine'] == 'ELFALLEH')].copy()
sub['Date'] = pd.to_datetime(sub['Date'])
sub = sub.sort_values('Date').reset_index(drop=True)
print('=== AMOR KHECHIN ELFALLEH detail ===')
print('Lignes:', len(sub))
print('Total:', sub['Tonnes/Jour'].sum(), 't')
print('Date debut declaree:', sub['Date Debut'].iloc[0] if len(sub) else '?')
print('Date fin declaree:', sub['Date Fin'].iloc[0] if len(sub) else '?')
print('Accessibilite:', sub['Accessibilite'].iloc[0] if len(sub) else '?')
print('Note:', sub['Note'].iloc[0] if len(sub) else '?')
print()
print('Toutes lignes:')
for i, r in sub.iterrows():
    d = r['Date'].strftime('%a %d/%m')
    print('  ' + d + ': ' + str(r['Tonnes/Jour']) + 't  ' + str(r['Vehicules']) + '  | Note: ' + str(r['Note']))
print()

# Verifier saturation ELFALLEH ce jour la
print('=== Total ELFALLEH le 25/06 (toutes livraisons) ===')
d2 = df[df['Date'].astype(str).str.contains('2026-06-25')]
elfalleh = d2[d2['Usine']=='ELFALLEH']
print('Total tonnes:', elfalleh['Tonnes/Jour'].sum(), 't (cap 150)')
print('Nombre livraisons:', len(elfalleh))
for _, r in elfalleh.iterrows():
    print('  ' + str(r['Agriculteur']) + ': ' + str(r['Tonnes/Jour']) + 't')
