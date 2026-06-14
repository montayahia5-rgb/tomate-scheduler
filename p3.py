f=open('optimizer_v2.py','r',encoding='utf-8')
c=f.read()
f.close()
old='nb_jours_doubles = 0\nfor row in all_days:\n    key = (row["Commercial"], row["Date"])\n    cap_norm = CAPS_NORMAL.get(row["Commercial"], 1000)'
new='nb_jours_doubles = 0\nfor row in all_days:\n    if row.get("_is_rm", False):\n        continue\n    key = (row["Commercial"], row["Date"])\n    cap_norm = CAPS_NORMAL.get(row["Commercial"], 1000)'
c=c.replace(old,new,1)
open('optimizer_v2.py','w',encoding='utf-8').write(c)
print('OK 3')
