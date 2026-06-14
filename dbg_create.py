# Verifier que _ub_day est calcule CORRECTEMENT au moment de NewIntVar
# Ajouter un debug a la creation de variable
c = open('optimizer_v2.py','r',encoding='utf-8').read()

# Chercher le bloc de creation
old = '''for f_idx, farmer in enumerate(nonrm_farmers):
    for d_idx, date in enumerate(all_dates):
        if date in farmer.window:
            # Γ£à Borne sup├⌐rieure JOURNALI├êRE bas├⌐e sur la COURBE de maturation
            # ├ù 1.5 = tol├⌐rance mod├⌐r├⌐e (un jour "double" = 1.5├ù la courbe naturelle)
            # ce qui emp├¬che OR-Tools de mettre 100t l├á o├╣ la courbe pr├⌐voit 35t
            day_planned = farmer.window[date]
            _ub_day = max(int(day_planned * SCALE * 1.5), int(_get_min_tons(farmer) * SCALE))
            x[(f_idx, d_idx)] = model.NewIntVar(0, _ub_day, f"x_{f_idx}_{d_idx}")'''

# Si le pattern exact ne matche pas, essayer une regex
import re
pattern = re.compile(
    r'(for f_idx, farmer in enumerate\(nonrm_farmers\):\s*\n\s*for d_idx, date in enumerate\(all_dates\):\s*\n\s*if date in farmer\.window:.*?day_planned = farmer\.window\[date\].*?_ub_day = max\(.*?\)\s*\n\s*x\[\(f_idx, d_idx\)\] = model\.NewIntVar\(0, _ub_day, .*?\))',
    re.DOTALL
)
m = pattern.search(c)
if m:
    print('Pattern trouve, ajout debug...')
    original = m.group(0)
    new_code = original + '''
            # DEBUG: pour AMOR KHECHIN ELFALLEH 25/06
            if 'AMOR KHECHIN' in farmer.name.upper() and farmer.usine == 'ELFALLEH' and date.day == 25 and date.month == 6:
                print('CREATION VAR x[' + str(f_idx) + ',' + str(d_idx) + ']: day_planned=' + str(day_planned) + ' _ub_day=' + str(_ub_day) + ' min_tons=' + str(_get_min_tons(farmer)))'''
    c = c.replace(original, new_code, 1)
    open('optimizer_v2.py','w',encoding='utf-8').write(c)
    print('OK debug creation insere')
else:
    print('Pattern non trouve, voici lignes 1011-1022:')
    for i, line in enumerate(c.split(chr(10))[1010:1022], 1011):
        print(i, ':', line)
