# Voir ce qu'OR-Tools donne reellement pour AMOR KHECHIN PL-ELFALLEH
# en l'instrumentant directement
c = open('optimizer_v2.py','r',encoding='utf-8').read()

# Trouver ou la consolidation se fait (autour ligne 1620)
lines = c.split(chr(10))
print('=== Consolidation block (1614-1632) ===')
for i in range(1614, 1635):
    if i < len(lines):
        print(str(i+1) + ': ' + lines[i])
