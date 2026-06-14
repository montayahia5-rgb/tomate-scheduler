import pandas as pd
df = pd.read_excel('Planning_Tomate_2026.xlsx', sheet_name='Planning Journalier')
habib = df[df['Agriculteur'].str.contains('HABIB BELWAER', case=False, na=False)]
print('Nombre de lignes Habib:', len(habib))
print()
sub = habib[['Date','Tonnes/Jour','Commercial','Usine','Vehicules']]
print(sub.to_string())
print()
total = habib['Tonnes/Jour'].sum()
print('Total:', total, 't')
print()
# Stats par phase
import datetime
habib_sorted = habib.sort_values('Date').reset_index(drop=True)
print('Phase debut (3 premiers jours):')
print(habib_sorted.head(3)[['Date','Tonnes/Jour']].to_string())
print()
print('Phase pic (jours 7-12):')
print(habib_sorted.iloc[6:12][['Date','Tonnes/Jour']].to_string())
