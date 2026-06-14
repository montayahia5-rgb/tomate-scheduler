import re
c = open('optimizer_v2.py','r',encoding='utf-8').read()

# Ajouter un print de debug juste apres OR-Tools solve, avant consolidation
# On cherche la ligne ou solution est construite
marker = 'solution = {(fi, di): solver.Value'
if marker in c:
    # Trouver la fin du dictionnaire solution
    idx = c.find(marker)
    # Trouver le } de fin de dict comprehension
    # Plus simple: ajouter le debug juste apres le marker (sur la prochaine ligne vide)
    end_block = c.find(chr(10) + chr(10), idx)
    debug = '''

# DEBUG AMOR KHECHIN ELFALLEH
print('=== DEBUG AMOR KHECHIN ELFALLEH ===')
for fi, f in enumerate(nonrm_farmers):
    if 'AMOR KHECHIN' in f.name.upper() and f.usine == 'ELFALLEH':
        print('Farmer:', f.name, '| tonnage=', f.tonnage)
        print('Window first 5:', list(f.window.items())[:5])
        total_ortools = 0
        for di, d in enumerate(all_dates):
            v = solution.get((fi,di), 0)
            if v > 0.5:
                ub_d = max(int(f.window.get(d,0) * SCALE * 1.5), int(_get_min_tons(f) * SCALE)) / SCALE if d in f.window else 0
                print('  ', d, ':', round(v,1), 't (UB etait', ub_d, ')')
                total_ortools += v
        print('TOTAL OR-Tools:', round(total_ortools,1))
        break
print('=== FIN DEBUG ===')
'''
    c = c[:end_block] + debug + c[end_block:]
    open('optimizer_v2.py','w',encoding='utf-8').write(c)
    print('OK debug insere - relancez python optimizer_v2.py')
else:
    print('Marker non trouve, cherche solution = ailleurs')
    for i, line in enumerate(c.split(chr(10))):
        if 'solution = {' in line or 'solution[' in line and 'solver.Value' in line:
            print(i+1, ':', line.strip()[:100])
