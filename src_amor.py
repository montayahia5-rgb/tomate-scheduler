import pandas as pd
# Lire directement depuis Supabase via le fichier source
sb_data = pd.read_excel('Planning_Tomate_2026.xlsx', sheet_name='Disponibilite Vehicules')
amor = sb_data[sb_data['Agriculteur'].str.contains('AMOR KHECHIN', case=False, na=False)]
print('=== Lots AMOR KHECHIN dans agriculteurs source ===')
for _, r in amor.iterrows():
    print('  ' + str(r['Agriculteur']) + ' | ' + str(r['Usine']) + ' | tonnage=' + str(r['Total Tonnes']) + ' | debut=' + str(r['Date Debut'])[:10] + ' | fin=' + str(r['Date Fin'])[:10])
