# Verifier solver.Value retourne bien la valeur de NOTRE variable
# et que personne n a remplace x[(f_idx, d_idx)]
c = open('optimizer_v2.py','r',encoding='utf-8').read()

# Cherche tout ce qui touche x[
print('Toutes les LIGNES qui contiennent x[ (assignations OU lectures):')
for i, line in enumerate(c.split(chr(10))):
    if 'x[(' in line or 'x[' in line and ('Constant' in line or 'NewIntVar' in line):
        # Filter only lines that ASSIGN to x
        if '=' in line and ' = ' in line[:line.index('=')+3] if '=' in line else False:
            print(i+1, ':', line.strip()[:120])

print()
print('Toutes les manipulations de model:')
for i, line in enumerate(c.split(chr(10))):
    if 'model.Add' in line and 'x[' in line:
        if i < 50 or i > 2000: continue
        print(i+1, ':', line.strip()[:120])
