import pandas as pd
df = pd.read_excel('Planning_Tomate_2026.xlsx', sheet_name='Planning Journalier')
NOM = 'AMOR KHECHIN'   # change ici
sub = df[df['Agriculteur'].str.contains(NOM, case=False, na=False)].copy()
sub['Date'] = pd.to_datetime(sub['Date'])
for usine in sub['Usine'].unique():
    s = sub[sub['Usine']==usine].sort_values('Date').reset_index(drop=True)
    print('===', usine, '(', len(s), 'lignes,', s['Tonnes/Jour'].sum(), 't )')
    for _, r in s.iterrows():
        d = r['Date'].strftime('%a %d/%m')
        t = r['Tonnes/Jour']
        marker = ' <<< PIC' if t >= 60 else (' << gros' if t >= 40 else '')
        print('  ' + d + ': ' + str(t) + 't  ' + str(r['Vehicules']) + marker)
    print()
