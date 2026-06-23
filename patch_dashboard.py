c = open('dashboard_phase10.py','r',encoding='utf-8').read()

# PATCH 1: Import comparaison_tab
old1 = '''try:
    from upload_tab import render_upload_tab, generate_template_excel
    UPLOAD_AVAILABLE = True
except ImportError:
    UPLOAD_AVAILABLE = False'''
new1 = old1 + '''

try:
    from comparaison_tab import render_comparaison_tab
    COMPARAISON_AVAILABLE = True
except ImportError:
    COMPARAISON_AVAILABLE = False'''

# PATCH 2: Ajouter tab_comp dans st.tabs
old2 = '''tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8, tab9, tab10 = st.tabs(['''
new2 = '''tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8, tab9, tab10, tab_comp = st.tabs(['''

old2b = '''        "Gamma tomate  Upload Planning",
    ])'''
new2b = '''        "Gamma tomate  Upload Planning",
        "Gamma comparaison  Comparaison Plans",
    ])'''

# PATCH 3: Contenu onglet (chercher fin du with tab10)
patch3 = '''
# ONGLET COMPARAISON PLANS
with tab_comp:
    if COMPARAISON_AVAILABLE:
        render_comparaison_tab(planning_df=planning, df_to_xlsx_styled=df_to_xlsx_styled)
    else:
        st.error("comparaison_tab.py introuvable")
        st.info("Mets comparaison_tab.py dans le meme dossier que dashboard_phase10.py")
'''

applied = []

if old1 in c and 'COMPARAISON_AVAILABLE' not in c:
    c = c.replace(old1, new1, 1)
    applied.append('PATCH 1: import OK')
else:
    applied.append('PATCH 1: deja present ou non trouve')

if old2 in c:
    c = c.replace(old2, new2, 1)
    applied.append('PATCH 2a: tab_comp OK')
else:
    applied.append('PATCH 2a: non trouve - verifier ligne st.tabs')

if '"Gamma tomate  Upload Planning"' in c:
    c = c.replace('"Gamma tomate  Upload Planning",\n    ])', '"Gamma tomate  Upload Planning",\n        "Gamma comparaison  Comparaison Plans",\n    ])', 1)
    applied.append('PATCH 2b: onglet ajoute OK')

if 'tab_comp' not in c.split('with tab10:')[-1] if 'with tab10:' in c else True:
    if 'with tab10:' in c:
        # Trouver la fin du bloc tab10 et ajouter apres
        idx = c.rfind('with tab10:')
        # Trouver la prochaine ligne qui commence par "with tab" ou fin fichier
        rest = c[idx:]
        next_tab = rest.find('\nwith tab', 10)
        if next_tab == -1:
            c = c + chr(10) + patch3
        else:
            pos = idx + next_tab
            c = c[:pos] + chr(10) + patch3 + c[pos:]
        applied.append('PATCH 3: with tab_comp ajoute OK')

open('dashboard_phase10.py','w',encoding='utf-8').write(c)

for msg in applied:
    print(msg)

# Verification finale
c2 = open('dashboard_phase10.py','r',encoding='utf-8').read()
print()
print('Verif COMPARAISON_AVAILABLE:', 'COMPARAISON_AVAILABLE' in c2)
print('Verif tab_comp dans tabs:', 'tab_comp' in c2)
print('Verif with tab_comp:', 'with tab_comp:' in c2)
print('Verif render_comparaison_tab:', 'render_comparaison_tab' in c2)
