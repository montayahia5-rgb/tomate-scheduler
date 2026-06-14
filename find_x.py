import re
c = open('optimizer_v2.py','r',encoding='utf-8').read()

# Chercher s il y a plusieurs definitions de x[(f_idx, d_idx)]
n_new_intvar = c.count('x[(f_idx, d_idx)] = model.NewIntVar')
n_new_const = c.count('x[(f_idx, d_idx)] = model.NewConstant')
print('x = NewIntVar:', n_new_intvar)
print('x = NewConstant:', n_new_const)

# Chercher s'il y a une REDEFINITION de x quelque part
# Cherche x[(f_idx, d_idx)] = sans NewIntVar
import re
pattern = re.compile(r'x\[\(f_idx, d_idx\)\]\s*=\s*([^m].*)$', re.MULTILINE)
for m in pattern.finditer(c):
    line_n = c[:m.start()].count(chr(10)) + 1
    print(f'Ligne {line_n}: {m.group(0)[:80]}')

# Chercher x[(fi, di)] aussi (variations)
print()
print('Toutes les definitions x[ ... ] = ...')
pattern2 = re.compile(r'x\[\([^\]]+\)\]\s*=\s*(.*?)$', re.MULTILINE)
seen = set()
for m in pattern2.finditer(c):
    line = m.group(0)[:80]
    if line not in seen:
        seen.add(line)
        line_n = c[:m.start()].count(chr(10)) + 1
        print(f'Ligne {line_n}: {line}')
