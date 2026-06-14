"""
Diagnostic ACHREF — à exécuter sur ta machine
    python diagnostic_achref.py

Vérifie EXACTEMENT ce qui est stocké dans Supabase pour ACHREF.
"""
from supabase import create_client
import pandas as pd

URL = "https://mwjefdqfzrtsfzspeppg.supabase.co"
KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Im13amVmZHFmenJ0c2Z6c3BlcHBnIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODAzODg3MDgsImV4cCI6MjA5NTk2NDcwOH0.H5IX2uLneHdvwyLJrgN7OrHGrLSZSrQBAuJeuzCEz44"
sb = create_client(URL, KEY)

print("="*70)
print("  DIAGNOSTIC ACHREF — Supabase")
print("="*70)

# ─── 1. TABLE agriculteurs ─────────────────────────────────────────────
print("\n[1/3] Table 'agriculteurs' — ACHREF AJLANI")
data = sb.table("agriculteurs").select(
    "nom,accessibilite,tonnage_total,usine,region"
).eq("commercial","ACHREF AJLANI").execute().data

df_a = pd.DataFrame(data)
if df_a.empty:
    print("  ❌ Aucun agriculteur ACHREF dans Supabase!")
else:
    print(f"  Total : {len(df_a)} agriculteurs | {df_a['tonnage_total'].sum():.0f}t")
    print("\n  Distribution ACCESSIBILITÉ :")
    print(df_a["accessibilite"].value_counts().to_string())
    
    # ALERTE si autre chose que SEMI
    bad = df_a[~df_a["accessibilite"].isin(["SEMI","RM"])]
    if not bad.empty:
        print(f"\n  ⚠️ {len(bad)} agriculteurs avec accessibilité ≠ SEMI :")
        print(bad[["nom","accessibilite","usine"]].head(15).to_string(index=False))

# ─── 2. TABLE planning ─────────────────────────────────────────────────
print("\n" + "="*70)
print("[2/3] Table 'planning' — ACHREF AJLANI")

all_p = []
offset = 0
while True:
    batch = sb.table("planning").select(
        "agriculteur,type_vehicule,tonnes_jour,usine,date"
    ).eq("commercial","ACHREF AJLANI").range(offset, offset+999).execute().data
    if not batch: break
    all_p.extend(batch)
    if len(batch) < 1000: break
    offset += 1000

dfp = pd.DataFrame(all_p)
if dfp.empty:
    print("  ❌ Aucune ligne planning ACHREF!")
else:
    print(f"  Total : {len(dfp)} lignes")
    print("\n  Distribution Type Véhicule :")
    print(dfp["type_vehicule"].value_counts().to_string())
    
    # ALERTE si pas que SEMI
    bad_v = dfp[dfp["type_vehicule"] != "SEMI"]
    if not bad_v.empty:
        print(f"\n  ⚠️ {len(bad_v)} lignes avec véhicule ≠ SEMI :")
        print(f"  Tonnage concerné : {bad_v['tonnes_jour'].sum():.0f}t")
        print("\n  Détail par usine :")
        print(bad_v.groupby(["usine","type_vehicule"]).agg(
            nb=("agriculteur","count"),
            tonnage=("tonnes_jour","sum")
        ).to_string())

# ─── 3. TABLE transport ─────────────────────────────────────────────────
print("\n" + "="*70)
print("[3/3] Table 'transport' — ACHREF AJLANI")
data_t = sb.table("transport").select(
    "date,tracteur,petit_poilour,poilour,semi,total_tonnes"
).eq("commercial","ACHREF AJLANI").execute().data

dft = pd.DataFrame(data_t)
if dft.empty:
    print("  ❌ Aucune ligne transport ACHREF!")
else:
    print(f"  Total : {len(dft)} jours")
    print(f"  Voyages TRACTEUR     : {pd.to_numeric(dft['tracteur'], errors='coerce').sum():.0f}")
    print(f"  Voyages PETIT POILOUR: {pd.to_numeric(dft['petit_poilour'], errors='coerce').sum():.0f}")
    print(f"  Voyages POILOUR      : {pd.to_numeric(dft['poilour'], errors='coerce').sum():.0f}")
    print(f"  Voyages SEMI         : {pd.to_numeric(dft['semi'], errors='coerce').sum():.0f}")
    
    pl_total = pd.to_numeric(dft['poilour'], errors='coerce').sum()
    ppl_total = pd.to_numeric(dft['petit_poilour'], errors='coerce').sum()
    if pl_total > 0 or ppl_total > 0:
        print(f"\n  ⚠️ ACHREF a {int(pl_total)} voyages PL et {int(ppl_total)} voyages PPL — devrait être 0!")

print("\n" + "="*70)
print("  CONCLUSION :")
print("="*70)
if not df_a.empty:
    bad_count = len(df_a[~df_a["accessibilite"].isin(["SEMI","RM"])])
    if bad_count > 0:
        print(f"  ❌ Supabase contient encore {bad_count} agriculteurs ACHREF avec acc≠SEMI")
        print(f"  → ACHREF doit re-uploader son fichier via le dashboard !")
    else:
        print(f"  ✅ Table agriculteurs propre (100% SEMI)")
if not dfp.empty:
    bad_v_count = len(dfp[dfp["type_vehicule"] != "SEMI"])
    if bad_v_count > 0:
        print(f"  ❌ Table planning contient {bad_v_count} lignes avec véhicule≠SEMI")
        print(f"  → Relancer : python optimizer_v2.py && python migrate.py")
    else:
        print(f"  ✅ Table planning propre (100% SEMI)")
print("="*70)