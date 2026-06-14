f=open('optimizer_v2.py','r',encoding='utf-8')
c=f.read()
f.close()

# FIX 1: NE PAS changer x2.0 (x1.1 cause INFEASIBLE - prouve dans tests)
# On garde day_planned * SCALE * 2.0 - c est le seul multiplicateur qui reste OPTIMAL

# FIX 2: ELFALLEH overflow weight (bon fix)
c=c.replace('FACTORY_OVERFLOW_WEIGHT = 500','FACTORY_OVERFLOW_WEIGHT = 2000')

# FIX 3: arrondi 5t au lieu de 10t (bon fix pour micro-livraisons)
c=c.replace('int(round(round(tons_brut, 1) / 10)) * 10','int(round(round(tons_brut, 1) / 5)) * 5')

f=open('optimizer_v2.py','w',encoding='utf-8')
f.write(c)
f.close()
print('OK - 2 corrections appliquees (overflow + arrondi 5t)')
print('ATTENTION: borne x2.0 conservee (x1.1 = INFEASIBLE confirme)')
