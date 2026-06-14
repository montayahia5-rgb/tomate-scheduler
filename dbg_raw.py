# Ajouter un debug AVANT le calcul solution = {(fi, di): solver.Value...
# pour voir la valeur RAW
c = open('optimizer_v2.py','r',encoding='utf-8').read()
old = 'solution = {(fi, di): solver.Value(x[(fi, di)]) / SCALE'
if old in c:
    new = '''# DEBUG: valeur RAW de x[166,5] avant le dict comprehension
try:
    raw_val = solver.Value(x[(166, 5)])
    print('RAW x[166,5] =', raw_val, '(UB devrait etre 408)')
except Exception as e:
    print('Erreur lecture x[166,5]:', e)
''' + old
    c = c.replace(old, new, 1)
    open('optimizer_v2.py','w',encoding='utf-8').write(c)
    print('OK debug RAW insere')
else:
    print('Pattern non trouve')
