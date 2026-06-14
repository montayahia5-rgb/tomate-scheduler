f=open('optimizer_v2.py','r',encoding='utf-8')
c=f.read()
f.close()
old='''for f_idx, farmer in enumerate(nonrm_farmers):
    _ndays_f   = max(1, len(farmer.window))
    _avg_rate  = farmer.tonnage / _ndays_f
    _ub_scaled = max(int(_avg_rate * SCALE * 3.0), 1)

    for d_idx, date in enumerate(all_dates):
        if date in farmer.window:
            x[(f_idx, d_idx)] = model.NewIntVar(0, _ub_scaled, f"x_{f_idx}_{d_idx}")
        else:
            x[(f_idx, d_idx)] = model.NewConstant(0)'''
new='''for f_idx, farmer in enumerate(nonrm_farmers):
    for d_idx, date in enumerate(all_dates):
        if date in farmer.window:
            # Borne journaliere basee sur la courbe de maturation x 1.5
            day_planned = farmer.window[date]
            _ub_day = max(int(day_planned * SCALE * 1.5), int(_get_min_tons(farmer) * SCALE))
            x[(f_idx, d_idx)] = model.NewIntVar(0, _ub_day, f"x_{f_idx}_{d_idx}")
        else:
            x[(f_idx, d_idx)] = model.NewConstant(0)'''
if old in c:
    c=c.replace(old,new,1)
    open('optimizer_v2.py','w',encoding='utf-8').write(c)
    print('OK - borne journaliere appliquee')
else:
    print('TEXTE NON TROUVE - verifier ligne 1011-1020')
    import re
    for i,line in enumerate(c.split(chr(10))[:1025],1):
        if '_avg_rate' in line or '_ub_scaled' in line:
            print(f'Ligne {i}: {line}')
