c = open('optimizer_v2.py','r',encoding='utf-8').read()
# Compter combien de fois DEBUG AMOR apparait
print('DEBUG AMOR present:', 'DEBUG AMOR KHECHIN' in c)
# Trouver la ligne EXACTE qui contient solver.Value
import re
for i, line in enumerate(c.split(chr(10))):
    if 'solver.Value' in line:
        print('Ligne', i+1, ':', line[:150])
