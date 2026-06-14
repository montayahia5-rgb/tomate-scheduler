c=open('optimizer_v2.py',encoding='utf-8').read()
lines=c.split(chr(10))
# Trouver la creation de variable NewIntVar pour x
for i,line in enumerate(lines):
    if 'NewIntVar' in line and 'x_' in line and 'f_idx' in line:
        # Afficher 8 lignes avant et 4 apres
        start=max(0,i-8); end=min(len(lines),i+4)
        print(f'=== Bloc autour ligne {i+1} ===')
        for j in range(start,end):
            print(f'{j+1:4d}: {lines[j]}')
        print()
        break
