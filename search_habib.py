import pandas as pd
df = pd.read_excel('Planning_Tomate_2026.xlsx', sheet_name='Planning Journalier')
# Chercher tous les noms contenant BEL
matches = df[df['Agriculteur'].str.contains('BEL', case=False, na=False)]
print('Noms contenant BEL:')
for n in matches['Agriculteur'].unique():
    print(' -', n)
print()
# Et HABIB
matches2 = df[df['Agriculteur'].str.contains('HABIB', case=False, na=False)]
print('Noms contenant HABIB:')
for n in matches2['Agriculteur'].unique():
    print(' -', n)
