c = open('optimizer_v2.py','r',encoding='utf-8').read()
lines = c.split(chr(10))
# Cherche tous les usages de 'solution = {' et le contexte
for i, line in enumerate(lines):
    if 'solution = {(fi, di): solver.Value' in line or 'RAW x' in line:
        start = max(0, i-5)
        end = min(len(lines), i+3)
        for j in range(start, end):
            print(str(j+1) + ': ' + lines[j][:120])
        print('---')
