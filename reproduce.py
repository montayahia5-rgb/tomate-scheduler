# Reproduire EXACTEMENT le calcul de la fenetre AMOR KHECHIN PL-ELFALLEH
import datetime, math

print('=== Reproduction fenetre AMOR KHECHIN (PL-ELFALLEH) ===')
raw_start = datetime.date(2026, 6, 25)
raw_end_orig = datetime.date(2026, 7, 7)
# Le code fait raw_end = raw_end - 1 jour
raw_end = raw_end_orig - datetime.timedelta(days=1)
print('start=', raw_start, 'end_apres_-1j=', raw_end)
n_days = (raw_end - raw_start).days + 1
print('n_days initial =', n_days)

# Extension fenetre +X jours selon tonnage
tonnage = 408
ext_days = 1 if tonnage < 300 else (2 if tonnage < 700 else 3)
print('ext_days =', ext_days)
end = raw_end + datetime.timedelta(days=ext_days)
n_days = (end - raw_start).days + 1
print('Apres extension:', raw_start, '->', end, '=', n_days, 'j')

# Pas RM, donc pas d'extension SEMI
# Extension faisabilite usine ELFALLEH cap=150, max_share = 150*0.60 = 90
ucap = 150
max_share = ucap * 0.60
daily = tonnage / n_days
print('daily=', round(daily,1), '< max_share=', max_share, '=> pas d extension')

# Courbe KHALIL: 20/60/20 sur 20%/50%/30%
n_debut = max(1, round(n_days * 0.20))
n_milieu = max(1, round(n_days * 0.50))
n_fin = n_days - n_debut - n_milieu
print('n_debut=', n_debut, ' n_milieu=', n_milieu, ' n_fin=', n_fin)

t_debut = tonnage * 0.20
t_milieu = tonnage * 0.60
t_fin = tonnage * 0.20

d_debut = round(t_debut / n_debut, 1)
d_milieu = round(t_milieu / n_milieu, 1)
d_fin = round(t_fin / n_fin if n_fin > 0 else 0, 1)
print('d_debut=', d_debut, ' d_milieu=', d_milieu, ' d_fin=', d_fin)
print()
print('=== Fenetre theorique complete ===')
for i in range(n_days):
    d = raw_start + datetime.timedelta(days=i)
    if i < n_debut:
        v = d_debut
        ph = 'debut'
    elif i < n_debut + n_milieu:
        v = d_milieu
        ph = 'milieu'
    else:
        v = d_fin
        ph = 'fin'
    ub = max(int(v * 10 * 1.5), int(15 * 10)) / 10  # PL min=15
    print('  ' + str(d) + ' (' + ph + '): courbe=' + str(v) + 't, UB OR-Tools=' + str(ub) + 't')

print()
print('=== ANALYSE 25/06 ===')
print('Le 25/06 est le premier jour = phase debut')
print('Courbe naturelle: ' + str(d_debut) + 't, UB max=' + str(d_debut * 1.5) + 't')
print('Excel montre 200t le 25/06 = IMPOSSIBLE avec ces bornes')
print()
print('=== ANALYSE TOTAL ===')
print('Total fenetre theorique:', sum([d_debut*n_debut, d_milieu*n_milieu, d_fin*n_fin]), 't (~408)')
print('Total Excel: 440t')
print('Difference: 32t cumules sur plusieurs jours')
