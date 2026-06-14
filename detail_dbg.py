# Ajouter un debug PLUS detaille
c = open('optimizer_v2.py','r',encoding='utf-8').read()

# Trouver le debug existant
idx = c.find('# DEBUG AMOR KHECHIN ELFALLEH')
if idx > 0:
    # Remplacer par version plus detaillee
    new_debug = '''# DEBUG AMOR KHECHIN ELFALLEH DETAILLE
print('=== DEBUG AMOR KHECHIN ELFALLEH ===')
# 1. Combien de farmers AMOR dans nonrm_farmers
amor_farmers = []
for fi, f in enumerate(nonrm_farmers):
    if 'AMOR KHECHIN' in f.name.upper():
        amor_farmers.append((fi, f))
        print('Trouve f_idx=' + str(fi) + ' name=' + f.name + ' usine=' + f.usine + ' access=' + str(f.access) + ' tonnage=' + str(f.tonnage))
print('Total farmers AMOR dans nonrm:', len(amor_farmers))
print()

# 2. Pour CHAQUE farmer AMOR, montrer ce qu OR-Tools a donne
for fi, f in amor_farmers:
    if f.usine != 'ELFALLEH':
        continue
    print('Detail farmer ELFALLEH f_idx=' + str(fi) + ' name=' + f.name)
    print('  Window keys count:', len(f.window))
    total = 0
    for di, d in enumerate(all_dates):
        v = solution.get((fi, di), 0)
        if v > 0.5:
            in_win = d in f.window
            win_val = f.window.get(d, 0)
            ub_val = max(int(win_val * SCALE * 1.5), int(_get_min_tons(f) * SCALE)) / SCALE if in_win else 0
            print('  ' + str(d) + ': v=' + str(round(v,1)) + ' window=' + str(win_val) + ' UB=' + str(ub_val) + ' in_win=' + str(in_win))
            total += v
    print('TOTAL OR-Tools pour ce farmer:', round(total,1))
    print()

# 3. Verifier f_idx total
print('N_FARM total:', len(nonrm_farmers))
print('=== FIN DEBUG ===')
'''
    # Trouver la fin du debug actuel
    end_marker = '=== FIN DEBUG ==='
    end_idx = c.find(end_marker)
    end_line_end = c.find(chr(10), end_idx) + 1
    
    # Le debut du debug
    start_idx = idx
    
    # Remplacer
    c = c[:start_idx] + new_debug + c[end_line_end:]
    open('optimizer_v2.py','w',encoding='utf-8').write(c)
    print('OK debug detaille insere')
else:
    print('Debug existant non trouve')
