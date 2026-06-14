c = open('optimizer_v2.py','r',encoding='utf-8').read()
old = 'day_planned * SCALE * 2.0'
new = 'day_planned * SCALE * 1.7'
if old in c:
    c = c.replace(old, new, 1)
    open('optimizer_v2.py','w',encoding='utf-8').write(c)
    print('OK - borne passee a x1.7')
