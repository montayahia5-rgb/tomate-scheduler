c = open('optimizer_v2.py','r',encoding='utf-8').read()

# Supprimer le debug RAW
old1 = '''    # DEBUG: valeur RAW de x[166,5] avant le dict comprehension
    try:
        raw_val = solver.Value(x[(166, 5)])
        print('RAW x[166,5] =', raw_val, '(UB devrait etre 408)')
    except Exception as e:
        print('Erreur lecture x[166,5]:', e)
    solution'''
new1 = '    solution'
c = c.replace(old1, new1, 1)

# Supprimer le debug AMOR KHECHIN
import re
pattern = re.compile(r'# DEBUG AMOR KHECHIN ELFALLEH DETAILLE.*?=== FIN DEBUG ===.*?\n', re.DOTALL)
c = pattern.sub('', c, count=1)

# Supprimer le debug CREATION VAR
old2 = '''            # DEBUG: pour AMOR KHECHIN ELFALLEH 25/06
            if 'AMOR KHECHIN' in farmer.name.upper() and farmer.usine == 'ELFALLEH' and date.day == 25 and date.month == 6:
                print('CREATION VAR x[' + str(f_idx) + ',' + str(d_idx) + ']: day_planned=' + str(day_planned) + ' _ub_day=' + str(_ub_day) + ' min_tons=' + str(_get_min_tons(farmer)))'''
c = c.replace(old2, '', 1)

open('optimizer_v2.py','w',encoding='utf-8').write(c)
print('Debug supprime - relance pour verifier')
