c = open('optimizer_v2.py','r',encoding='utf-8').read()
old = '''else:
    # DEBUG: valeur RAW de x[166,5] avant le dict comprehension
try:
    raw_val = solver.Value(x[(166, 5)])
    print('RAW x[166,5] =', raw_val, '(UB devrait etre 408)')
except Exception as e:
    print('Erreur lecture x[166,5]:', e)
solution = {(fi, di): solver.Value(x[(fi, di)]) / SCALE
                for fi in range(N_FARM) for di in range(N_DATES)}'''
new = '''else:
    # DEBUG: valeur RAW de x[166,5] avant le dict comprehension
    try:
        raw_val = solver.Value(x[(166, 5)])
        print('RAW x[166,5] =', raw_val, '(UB devrait etre 408)')
    except Exception as e:
        print('Erreur lecture x[166,5]:', e)
    solution = {(fi, di): solver.Value(x[(fi, di)]) / SCALE
                    for fi in range(N_FARM) for di in range(N_DATES)}'''
if old in c:
    c = c.replace(old, new, 1)
    open('optimizer_v2.py','w',encoding='utf-8').write(c)
    print('OK fix indentation')
else:
    print('Pattern non trouve, tentative manuelle')
    lines = c.split(chr(10))
    # Indenter lignes 1291-1297
    for i in [1290, 1291, 1292, 1293, 1294, 1295, 1296]:
        if i < len(lines) and not lines[i].startswith('    '):
            lines[i] = '    ' + lines[i]
    c_new = chr(10).join(lines)
    open('optimizer_v2.py','w',encoding='utf-8').write(c_new)
    print('OK fix manuel - re-verifier')
