c = open('optimizer_v2.py','r',encoding='utf-8').read()
# La ligne 1300 doit etre complete
# Chercher la ligne incorrecte
old_line = 'ub_d = max(int(f.window.get(d,0) * SCALE * 1.5), int(_get_min_tons(f) * SCALE)) / SCALE if d in f.window'
new_line = 'ub_d = max(int(f.window.get(d,0) * SCALE * 1.5), int(_get_min_tons(f) * SCALE)) / SCALE if d in f.window else 0'

if old_line in c and old_line + ' else 0' not in c:
    c = c.replace(old_line, new_line, 1)
    open('optimizer_v2.py','w',encoding='utf-8').write(c)
    print('OK - syntax fix applique')
elif old_line + ' else 0' in c:
    print('Deja correct')
else:
    print('Ligne non trouvee, voici ligne 1300 exacte:')
    print(repr(c.split(chr(10))[1299]))
