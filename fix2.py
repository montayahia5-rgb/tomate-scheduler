f=open('optimizer_v2.py','r',encoding='utf-8')
c=f.read()
f.close()
old='    remaining_diff = diff  # positif'
new='    if not _is_semi_rm:\n        continue\n    remaining_diff = diff  # positif'
c=c.replace(old,new,1)
f=open('optimizer_v2.py','w',encoding='utf-8')
f.write(c)
f.close()
print('OK')
