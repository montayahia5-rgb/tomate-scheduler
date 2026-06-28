#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""patch_agroeco.py — Ajoute l'onglet 📊 Dashboard Agroéco dans dashboard_phase10.py"""
import os, sys, shutil
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
TARGET = os.path.join(SCRIPT_DIR, "dashboard_phase10.py")
BACKUP = TARGET + ".bak_preagroeco"
if not os.path.exists(TARGET):
    print("❌ dashboard_phase10.py introuvable"); sys.exit(1)
shutil.copy2(TARGET, BACKUP)
print(f"✅ Backup : {BACKUP}")
with open(TARGET, encoding="utf-8") as f:
    code = f.read()

# Patch 1 : import
OLD1 = "try:\n    from agro_performance import render_agro_tab"
NEW1 = """try:
    from agroeco_dashboard import render_agroeco_tab
    AGROECO_AVAILABLE = True
except ImportError:
    AGROECO_AVAILABLE = False

try:
    from agro_performance import render_agro_tab"""
if OLD1 in code and "agroeco_dashboard" not in code:
    code = code.replace(OLD1, NEW1, 1); print("✅ Patch 1 : import agroeco_dashboard")
else:
    print("⏭️  Patch 1 déjà appliqué")

# Patch 2 : tab directeur
OLD2 = '        "📊 Comparaison Plans",\n    ])'
NEW2 = '        "📊 Comparaison Plans",\n        "📊 Dashboard Agroéco",\n    ])'
OLD2_VAR = 'tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8, tab9, tab10, tab_comp, tab_agro = st.tabs(['
NEW2_VAR = 'tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8, tab9, tab10, tab_comp, tab_agro, tab_agroeco = st.tabs(['
if OLD2_VAR in code and "tab_agroeco" not in code:
    code = code.replace(OLD2_VAR, NEW2_VAR, 1)
    code = code.replace(
        '"📊 Comparaison Plans",\n        "🌱 Performance Agro",\n    ])',
        '"📊 Comparaison Plans",\n        "🌱 Performance Agro",\n        "📊 Dashboard Agroéco",\n    ])', 1)
    print("✅ Patch 2 : tab directeur")
else:
    print("⏭️  Patch 2 déjà appliqué")

# Patch 3 : alias usine et commercial
for old3, new3 in [
    ("    tab_agro = tab4   # usine", "    tab_agro = tab4   # usine\n    tab_agroeco = tab4"),
    ("    tab5 = tab7 = tab8 = tab4\n    tab_agro", "    tab5 = tab7 = tab8 = tab4\n    tab_agroeco = tab4\n    tab_agro"),
]:
    if old3 in code and new3 not in code:
        code = code.replace(old3, new3, 1); print(f"✅ Patch 3 : alias")

# Patch 4 : contenu tab
ANCHOR = "# ONGLET COMPARAISON PLANS\nwith tab_comp:"
AGROECO_BLOCK = """
# ── TAB DASHBOARD AGROÉCONOMIQUE ────────────────────────────
with tab_agroeco:
    if AGROECO_AVAILABLE:
        render_agroeco_tab(sb=get_supabase(),
                           CURRENT_ROLE=CURRENT_ROLE,
                           CURRENT_NAME=CURRENT_NAME)
    else:
        st.error("❌ agroeco_dashboard.py introuvable")
        st.info("Place agroeco_dashboard.py dans le même dossier que dashboard_phase10.py")
"""
if ANCHOR in code and "render_agroeco_tab" not in code:
    code = code.replace(ANCHOR, AGROECO_BLOCK + "\n" + ANCHOR, 1)
    print("✅ Patch 4 : contenu tab_agroeco")
else:
    print("⏭️  Patch 4 déjà appliqué")

with open(TARGET, "w", encoding="utf-8") as f:
    f.write(code)
print(f"\n🎉  Patches appliqués → {TARGET}")