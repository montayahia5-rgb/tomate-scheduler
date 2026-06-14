f=open('optimizer_v2.py','r',encoding='utf-8')
c=f.read()
f.close()
if '"_is_rm":        True' not in c:
    c=c.replace('            "Note":          "RM-pre-alloc",\n        })','            "Note":          "RM-pre-alloc",\n            "_is_rm":        True,\n        })',1)
    open('optimizer_v2.py','w',encoding='utf-8').write(c)
    print('OK 1')
else:
    print('OK 1 deja present')
