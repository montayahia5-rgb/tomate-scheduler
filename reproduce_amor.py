# Reproduire le calcul exact de farmer.window pour AMOR KHECHIN ELFALLEH
import sys
sys.path.insert(0, '.')
import datetime, math
from supabase import create_client
import pandas as pd

sb = create_client('https://mwjefdqfzrtsfzspeppg.supabase.co','eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Im13amVmZHFmenJ0c2Z6c3BlcHBnIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODAzODg3MDgsImV4cCI6MjA5NTk2NDcwOH0.H5IX2uLneHdvwyLJrgN7OrHGrLSZSrQBAuJeuzCEz44')
data = sb.table('agriculteurs').select('*').execute().data
df = pd.DataFrame(data)

# Importer le module optimizer pour creer un Farmer
import importlib.util
spec = importlib.util.spec_from_file_location('opt', 'optimizer_v2.py')
print('Loading optimizer_v2.py module (may take a moment)...')

# Faire le calcul manuellement avec les EXACTES regles du code
amor_row = df[(df['nom']=='AMOR KHECHIN') & (df['usine']=='ELFALLEH')].iloc[0]
print('Donnees brutes:')
print('  nom:', amor_row['nom'])
print('  usine:', amor_row['usine'])
print('  tonnage:', amor_row['tonnage_total'])
print('  access:', amor_row.get('accessibilite'))
print('  date_debut:', amor_row['date_debut'])
print('  date_fin:', amor_row['date_fin'])
print('  commercial:', amor_row['commercial'])
print()

# Maintenant simuler EXACTEMENT le code de Farmer.__init__
tonnage = float(amor_row['tonnage_total'])
raw_start = pd.to_datetime(amor_row['date_debut']).date()
raw_end = pd.to_datetime(amor_row['date_fin']).date()

# Le code fait raw_end - 1 jour
raw_end = raw_end - datetime.timedelta(days=1)
print('Apres -1j: start=', raw_start, ' end=', raw_end)

SEASON_START = datetime.date(2026,6,15)
SEASON_END = datetime.date(2026,8,31)
def clamp(d):
    if d < SEASON_START: return SEASON_START
    if d > SEASON_END: return SEASON_END
    return d

start = clamp(raw_start)
end = clamp(raw_end)
if end <= start and tonnage > 100:
    end = clamp(start + datetime.timedelta(days=30))

# Extension +2 jours pour tonnage 300-700
ext_days = 1 if tonnage < 300 else (2 if tonnage < 700 else 3)
end = clamp(end + datetime.timedelta(days=ext_days))
n_days = max(1, (end - start).days + 1)
print('Apres ext+' + str(ext_days) + ': start=', start, ' end=', end, ' n_days=', n_days)

# Pas RM (access=PL)
# Extension faisabilite ELFALLEH cap=150, max_share=90
ucap = 150
max_share = ucap * 0.60
daily_now = tonnage / n_days
print('daily_now=', round(daily_now,2), ' max_share=', max_share)
if daily_now > max_share:
    days_needed = math.ceil(tonnage / max_share)
    extra = days_needed - n_days
    if extra > 0:
        end = clamp(end + datetime.timedelta(days=extra))
        n_days = max(1, (end - start).days + 1)
        print('Extension faisabilite +', extra, 'j: end=', end, ' n_days=', n_days)

# Courbe maturation - commercial = KHALIL → 20/60/20 sur 20%/50%/30%
commercial = amor_row['commercial']
is_khalil = commercial.strip().upper() == 'KHALIL'
print('commercial=', commercial, ' is_khalil=', is_khalil)

if is_khalil:
    pct_debut, pct_milieu = 0.20, 0.60
    time_debut, time_milieu = 0.20, 0.50
else:
    pct_debut, pct_milieu = 0.20, 0.60
    time_debut, time_milieu = 0.30, 0.50

n_debut = max(1, round(n_days * time_debut))
n_milieu = max(1, round(n_days * time_milieu))
n_fin = n_days - n_debut - n_milieu
if n_fin < 1:
    n_fin = 1
    if n_milieu > 1: n_milieu -= 1
    else: n_debut = max(1, n_debut - 1)

t_debut = tonnage * pct_debut
t_milieu = tonnage * pct_milieu
t_fin = tonnage - t_debut - t_milieu

d_debut = t_debut / n_debut
d_milieu = t_milieu / n_milieu
d_fin = t_fin / max(1, n_fin)

print()
print('=== FENETRE THEORIQUE ===')
print('n_debut=', n_debut, ' n_milieu=', n_milieu, ' n_fin=', n_fin)
print('d_debut=', round(d_debut,1), ' d_milieu=', round(d_milieu,1), ' d_fin=', round(d_fin,1))
print()

# Calculer borne UB pour chaque jour
SCALE = 10
veh_min = 15  # PL
for i in range(n_days):
    d = start + datetime.timedelta(days=i)
    if i < n_debut:
        v = round(d_debut, 1)
        ph = 'debut'
    elif i < n_debut + n_milieu:
        v = round(d_milieu, 1)
        ph = 'milieu'
    else:
        v = round(d_fin, 1)
        ph = 'fin'
    ub_scaled = max(int(v * SCALE * 1.5), int(veh_min * SCALE))
    ub_tons = ub_scaled / SCALE
    print('  ' + str(d) + ' (' + ph + '): courbe=' + str(v) + 't | UB OR-Tools=' + str(ub_tons) + 't')

print()
print('=> Borne max sur n importe quel jour =', round(max(d_debut, d_milieu, d_fin) * 1.5, 1), 't')
print('=> Excel montre 200t le 25/06')
print('=> Donc soit le code utilise PAS la borne, soit window est differente')
