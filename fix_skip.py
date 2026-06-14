f=open('optimizer_v2.py','r',encoding='utf-8')
c=f.read()
f.close()
old='    _min_t = _get_min_tons(_farmer_ref) if _farmer_ref else 10\n    _is_semi_rm = False  # deja skipe au dessus pour les vrais RM'
new='    # SKIP correction non-RM/non-SEMI : elle empire toujours les resultats\n    # (arrondis cumules crees un surplus systematique)\n    continue'
if old in c:
    c=c.replace(old,new,1)
    open('optimizer_v2.py','w',encoding='utf-8').write(c)
    print('OK - correction non-RM desactivee')
else:
    print('TEXTE NON TROUVE - cherche manuellement')
    idx=c.find('_is_semi_rm = False')
    if idx>0: print('Ligne ~',c[:idx].count(chr(10))+1)
