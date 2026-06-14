c = open('optimizer_v2.py','r',encoding='utf-8').read()
old = '_ub_day = max(int(day_planned * SCALE * 1.5), int(_get_min_tons(farmer) * SCALE))'
new = '_ub_day = max(int(day_planned * SCALE * 2.0), int(_get_min_tons(farmer) * SCALE))'
if old in c:
    c = c.replace(old, new, 1)
    open('optimizer_v2.py','w',encoding='utf-8').write(c)
    print('OK - borne passee a x2.0')
else:
    print('Pattern non trouve')
