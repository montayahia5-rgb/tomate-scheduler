import sys, datetime, math
sys.path.insert(0, '.')
import pandas as pd
from collections import defaultdict
import os

# Charger directement depuis Supabase pour avoir les VRAIES dates
from supabase import create_client
SB_URL='https://mwjefdqfzrtsfzspeppg.supabase.co'
SB_KEY='eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Im13amVmZHFmenJ0c2Z6c3BlcHBnIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODAzODg3MDgsImV4cCI6MjA5NTk2NDcwOH0.H5IX2uLneHdvwyLJrgN7OrHGrLSZSrQBAuJeuzCEz44'
sb = create_client(SB_URL, SB_KEY)
data = sb.table('agriculteurs').select('*').execute().data
df = pd.DataFrame(data)
amor = df[df['nom'].astype(str).str.contains('AMOR KHECHIN', case=False, na=False)]
amor = amor[amor['usine']=='ELFALLEH']
print('=== AMOR KHECHIN ELFALLEH dans Supabase ===')
for _, r in amor.iterrows():
    print('nom:', r['nom'])
    print('usine:', r['usine'])
    print('tonnage:', r.get('tonnage_total'))
    print('accessibilite:', r.get('accessibilite'))
    print('date_debut:', r.get('date_debut'))
    print('date_fin:', r.get('date_fin'))
    print('commercial:', r.get('commercial'))
    print()

# Maintenant simuler la fenetre comme le code le ferait
print('=== Simulation fenetre maturation pour ce lot ===')
tonnage = 408
start = datetime.date(2026, 6, 25)
end = datetime.date(2026, 7, 8)
# Extension +2 jours pour tonnage 408t (entre 300 et 700)
end = end + datetime.timedelta(days=2)
n_days = (end - start).days + 1
print('Apres extension: ' + str(start) + ' -> ' + str(end) + ' = ' + str(n_days) + 'j')

# Courbe FEDI (assumant FEDI - mais il faut savoir quel commercial)
n_debut = max(1, round(n_days * 0.30))
n_milieu = max(1, round(n_days * 0.50))
n_fin = n_days - n_debut - n_milieu
print('Phases: debut=' + str(n_debut) + 'j, milieu=' + str(n_milieu) + 'j, fin=' + str(n_fin) + 'j')
d_debut = tonnage * 0.20 / n_debut
d_milieu = tonnage * 0.60 / n_milieu
d_fin = tonnage * 0.20 / max(1, n_fin)
print('Tonnage/j: debut=' + str(round(d_debut,1)) + ', milieu=' + str(round(d_milieu,1)) + ', fin=' + str(round(d_fin,1)))
print('Bornes UB (x1.5): debut=' + str(round(d_debut*1.5,1)) + ', milieu=' + str(round(d_milieu*1.5,1)) + ', fin=' + str(round(d_fin*1.5,1)))
print()
print('Si OR-Tools place 200t le 25/06, c IMPOSSIBLE avec ces bornes')
print('=> Le pattern doit etre dans un POST-TRAITEMENT')
