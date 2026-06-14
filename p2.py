f=open('optimizer_v2.py','r',encoding='utf-8')
c=f.read()
f.close()
if '"_is_rm":        False' not in c:
    c=c.replace('        "Note":          note,\n    })\n\n# FUSION RM','        "Note":          note,\n        "_is_rm":        False,\n    })\n\n# FUSION RM',1)
    open('optimizer_v2.py','w',encoding='utf-8').write(c)
    print('OK 2')
else:
    print('OK 2 deja present')
