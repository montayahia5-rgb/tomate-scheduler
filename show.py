c = open('optimizer_v2.py','r',encoding='utf-8').read()
lines = c.split(chr(10))
# Afficher lignes 1284-1300
for i in range(1283, min(1305, len(lines))):
    print(str(i+1) + ':' + repr(lines[i][:120]))
