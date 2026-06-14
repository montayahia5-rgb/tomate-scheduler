f=open('optimizer_v2.py','r',encoding='utf-8')
c=f.read()
f.close()
old='        nb_semi = max(1, int(round(tons / SEMI_CAP)))'
new='        nb_semi = max(min_semi, int(round(tons / SEMI_CAP)))'
c=c.replace(old,new,1)
f=open('optimizer_v2.py','w',encoding='utf-8')
f.write(c)
f.close()
print('OK')
