# Confirmer empiriquement que la borne s applique
# en lisant le fichier source EXACTEMENT
import re

c = open('optimizer_v2.py','r',encoding='utf-8').read()

# Compter les occurrences du pattern _ub_day
n_ub_day = c.count('_ub_day')
n_ub_scaled = c.count('_ub_scaled')
n_avg_rate = c.count('_avg_rate')

print('_ub_day:', n_ub_day, '(doit etre >= 2)')
print('_ub_scaled:', n_ub_scaled, '(doit etre 0)')
print('_avg_rate:', n_avg_rate, '(doit etre 0)')

# Verifier que NewIntVar utilise _ub_day
if 'NewIntVar(0, _ub_day' in c:
    print('OK - NewIntVar utilise _ub_day')
else:
    print('PROBLEME - NewIntVar ne utilise PAS _ub_day')
    # Trouver ce qu'il utilise
    m = re.search(r'NewIntVar\(0, ([^,]+),', c)
    if m:
        print('NewIntVar utilise:', m.group(1))
