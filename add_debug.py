# Instrumenter le code pour logger ce qu OR-Tools donne pour AMOR KHECHIN
# Avant la consolidation, ligne ~1620
c = open('optimizer_v2.py','r',encoding='utf-8').read()

# Inserer un debug print AVANT la consolidation
old = '''for f_idx, farmer in enumerate(nonrm_farmers):
    for d_idx, date in enumerate(all_dates):
        tons = solution[(f_idx, d_idx)]
        if tons <= 0.5: continue
        # Cle : nom agriculteur + usine + date (= un envoi unique)
        key = (farmer.name.strip().upper(), farmer.usine, date)
        consolidated[key][\"tons\"] += round(tons, 1)
        consolidated[key][\"farmer\"] = farmer  # garder la derniere reference
        consolidated[key][\"parcelles\"].append(f_idx)'''

new = '''# DEBUG: tracer AMOR KHECHIN
_debug_amor = {}
for f_idx, farmer in enumerate(nonrm_farmers):
    is_amor = 'AMOR KHECHIN' in farmer.name.upper() and 'ELFALLEH' in farmer.usine.upper()
    if is_amor:
        print(f'DEBUG AMOR farmer={farmer.name}, usine={farmer.usine}, allowed={farmer.allowed_veh}, tonnage={farmer.tonnage}')
        print(f'  window keys: {sorted(farmer.window.keys())[:5]}...')
        print(f'  window values first 5: {[farmer.window[d] for d in sorted(farmer.window.keys())[:5]]}')
    for d_idx, date in enumerate(all_dates):
        tons = solution[(f_idx, d_idx)]
        if tons <= 0.5: continue
        if is_amor:
            _debug_amor[date] = tons
        key = (farmer.name.strip().upper(), farmer.usine, date)
        consolidated[key]['tons'] += round(tons, 1)
        consolidated[key]['farmer'] = farmer
        consolidated[key]['parcelles'].append(f_idx)

if _debug_amor:
    print('DEBUG OR-Tools solution pour AMOR KHECHIN ELFALLEH:')
    for d in sorted(_debug_amor.keys()):
        print(f'  {d}: {_debug_amor[d]:.1f}t')
    print(f'  TOTAL: {sum(_debug_amor.values()):.1f}t')'''

if old.replace(chr(34), chr(39)) in c.replace(chr(34), chr(39)):
    # Le code utilise des guillemets doubles, on remplace en preservant
    import re
    # Match flexible sur les guillemets
    pattern = re.compile(r'for f_idx, farmer in enumerate\(nonrm_farmers\):\s*\n.*?consolidated\[key\]\[.parcelles.\]\.append\(f_idx\)', re.DOTALL)
    if pattern.search(c):
        c_new = pattern.sub(new.replace(chr(39), chr(34)), c, count=1)
        if c_new != c:
            open('optimizer_v2.py','w',encoding='utf-8').write(c_new)
            print('OK - debug insere, relance python optimizer_v2.py')
        else:
            print('Pas de changement')
    else:
        print('Pattern non trouve')
else:
    print('Texte exact non trouve - inspection manuelle requise')
