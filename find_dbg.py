c = open('optimizer_v2.py','r',encoding='utf-8').read()
lines = c.split(chr(10))
# Trouver ou est le debug
for i, line in enumerate(lines):
    if 'DEBUG AMOR' in line:
        # Afficher 5 lignes avant et 25 apres
        start = max(0, i-3)
        end = min(len(lines), i+30)
        for j in range(start, end):
            print(j+1, '|', lines[j][:120])
        break
