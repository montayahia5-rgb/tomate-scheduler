import pandas as pd
df = pd.read_excel('Planning_Tomate_2026.xlsx', sheet_name='Planning Journalier')
habib = df[df['Agriculteur'].str.contains('HABIB BELWEAR', case=False, na=False)].copy()
habib['Date'] = pd.to_datetime(habib['Date'])
habib = habib.sort_values(['Usine','Date']).reset_index(drop=True)

for usine in habib['Usine'].unique():
    sub = habib[habib['Usine']==usine].sort_values('Date').reset_index(drop=True)
    print('===', usine, '===')
    print('Lignes:', len(sub), '| Total:', sub['Tonnes/Jour'].sum(), 't')
    print()
    for i, row in sub.iterrows():
        date_str = row['Date'].strftime('%a %d/%m')
        tons = row['Tonnes/Jour']
        veh = row['Vehicules']
        marker = ''
        if tons >= 80: marker = ' <<< GROS'
        elif tons >= 60: marker = ' << PIC'
        print('  ' + date_str + ': ' + str(tons) + 't  ' + str(veh) + marker)
    print()
