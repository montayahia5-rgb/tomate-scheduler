# Lire l'agriculteur AMOR KHECHIN complet pour voir s'il y a des homonymes
import pandas as pd
from supabase import create_client
sb = create_client('https://mwjefdqfzrtsfzspeppg.supabase.co','eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Im13amVmZHFmenJ0c2Z6c3BlcHBnIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODAzODg3MDgsImV4cCI6MjA5NTk2NDcwOH0.H5IX2uLneHdvwyLJrgN7OrHGrLSZSrQBAuJeuzCEz44')
data = sb.table('agriculteurs').select('*').execute().data
df = pd.DataFrame(data)
# Tous les agriculteurs nommes AMOR
amor = df[df['nom'].astype(str).str.contains('AMOR', case=False, na=False)]
print('=== Tous agriculteurs AMOR ===')
for _, r in amor.iterrows():
    print(r['nom'] + ' | comm=' + r['commercial'] + ' | usine=' + r['usine'] + ' | access=' + str(r.get('accessibilite')) + ' | ton=' + str(r.get('tonnage_total')) + ' | debut=' + str(r['date_debut'])[:10])
print()
# Lots distincts par (nom, usine) - le code differencie via name_count
print('=== Doublons exacts ? ===')
counts = df.groupby(['nom','commercial']).size().reset_index(name='count')
amor_counts = counts[counts['nom'].astype(str).str.contains('AMOR', case=False, na=False)]
print(amor_counts.to_string())
