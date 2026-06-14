lines = open('optimizer_v2.py','r',encoding='utf-8').readlines()
# Trouver la ligne avec _is_semi_rm = False et la remplacer
for i,line in enumerate(lines):
    if '_is_semi_rm = False' in line:
        print(f'Ligne {i+1}: {line.rstrip()}')
        # Remplacer cette ligne et la precedente par continue
        lines[i-1] = ''
        lines[i] = '    continue  # SKIP correction non-SEMI\n'
        print('Remplacement effectue')
        break
open('optimizer_v2.py','w',encoding='utf-8').writelines(lines)
print('Fichier sauvegarde')
