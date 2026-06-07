# -*- coding: utf-8 -*-
"""
restore_agri.py — Restaure les données ACHRAF, KHALIL, MAKKI supprimées par --fix
==================================================================================
Le --fix a supprimé de vraies données car il a confondu :
- Agriculteur avec 2 PARCELLES différentes → 2 lignes légitimes avec mêmes nom+usine
  (ex: AMOR KHECHIN a 2 zones différentes dans KHALIL, chacune avec tonnage distinct)

Solution: supprimer et réinsérer depuis les fichiers ORGANISE corrects.
"""
import sys, io, os
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
import pandas as pd
from supabase import create_client

SUPABASE_URL = "https://mwjefdqfzrtsfzspeppg.supabase.co"
SUPABASE_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Im13amVmZHFmenJ0c2Z6c3BlcHBnIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODAzODg3MDgsImV4cCI6MjA5NTk2NDcwOH0.H5IX2uLneHdvwyLJrgN7OrHGrLSZSrQBAuJeuzCEz44"

sb = create_client(SUPABASE_URL, SUPABASE_KEY)

# Fichiers ORGANISE corrects (avec toutes les parcelles)
ORGANISE_FILES = {
    "KHALIL":       "KHALIL_ORGANISE.xlsx",
    "ACHREF AJLANI":"ACHRAF_ORGANISE.xlsx",
    "MAKKI BEN SALAH": "MAKKI_ORGANISE.xlsx",
}

REGION_NORM = {
    "NABEUL":"CAP BON 2","BEJA":"NORD","MANOUBA":"NORD",
    "GAFSA":"GAFSA / KASSRINE","KASSRINE":"GAFSA / KASSRINE",
    "CAPB1":"CAP BON 1","CAPB2":"CAP BON 2",
}

expected = {"KHALIL":15705,"ACHREF AJLANI":14140,"MAKKI BEN SALAH":21365}

print("="*60)
print("  RESTAURATION DES DONNÉES SUPPRIMÉES PAR --FIX")
print("="*60)

for commercial, filename in ORGANISE_FILES.items():
    # Chercher dans plusieurs emplacements
    paths = [filename, f"../{filename}", f"outputs/{filename}"]
    df = None
    for path in paths:
        if os.path.exists(path):
            df = pd.read_excel(path)
            print(f"\n✅ Chargé: {path}")
            break
    
    if df is None:
        print(f"\n❌ Fichier non trouvé: {filename}")
        print(f"   Cherché dans: {paths}")
        continue
    
    df['TONNAGE'] = pd.to_numeric(df['TONNAGE'], errors='coerce')
    df = df[df['TONNAGE'] > 0].copy()
    
    total = df['TONNAGE'].sum()
    exp   = expected[commercial]
    print(f"   {commercial}: {len(df)} lignes | {total:,.0f}t (attendu {exp:,}t)")
    
    if abs(total - exp) > 200:
        print(f"   ⚠️ Écart > 200t — vérifier le fichier")
    
    # Supprimer les données actuelles de ce commercial
    sb.table("agriculteurs").delete().eq("commercial", commercial).execute()
    print(f"   Données supprimées de Supabase")
    
    # Réinsérer depuis le fichier ORGANISE
    rows = []
    for _, row in df.iterrows():
        region_raw = str(row.get("REGION","") or "").strip().upper()
        region = REGION_NORM.get(region_raw, region_raw) or "CAP BON 2"
        zone   = str(row.get("ZONE","") or row.get("ZONNE","") or "").strip()
        debut  = str(row.get("DATE_DEBUT","") or "2026-07-01")
        fin    = str(row.get("DATE_FIN","") or "2026-07-31")
        
        rows.append({
            "commercial":    commercial,
            "nom":           str(row["NOM_AGRICULTEUR"]).strip(),
            "region":        region,
            "zone":          zone or None,
            "usine":         str(row["USINE"]).strip(),
            "accessibilite": str(row["ACCESSIBILITE"]).strip(),
            "tonnage_total": float(row["TONNAGE"]),
            "date_debut":    debut[:10] if len(str(debut)) >= 10 else debut,
            "date_fin":      fin[:10]   if len(str(fin))   >= 10 else fin,
        })
    
    # Insérer par batch
    for i in range(0, len(rows), 100):
        sb.table("agriculteurs").insert(rows[i:i+100]).execute()
    
    # Vérifier
    verif = sb.table("agriculteurs").select("tonnage_total").eq("commercial", commercial).execute().data
    verif_total = sum(float(r["tonnage_total"]) for r in verif if r["tonnage_total"])
    ok = "✅" if abs(verif_total - exp) < 100 else "❌"
    print(f"   Réinséré: {len(rows)} lignes | {verif_total:,.0f}t {ok}")

print("\n" + "="*60)
print("  VÉRIFICATION FINALE")
print("="*60)
all_data = sb.table("agriculteurs").select("commercial,tonnage_total").execute().data
from collections import defaultdict
by_comm = defaultdict(float)
for r in all_data:
    by_comm[r["commercial"]] += float(r["tonnage_total"] or 0)

expected_all = {
    "FEDI":31690,"MAKKI BEN SALAH":21365,"KHALIL":15705,
    "ACHREF AJLANI":14140,"JILANI OBAY":7410
}
total = 0
for comm in sorted(by_comm.keys()):
    exp = expected_all.get(comm, "?")
    ecart = by_comm[comm] - exp if isinstance(exp, int) else 0
    ok = "✅" if abs(ecart) < 100 else "❌"
    print(f"  {comm:<22}: {by_comm[comm]:>9,.0f}t (attendu {str(exp)+'t':>7}) {ok}")
    total += by_comm[comm]
print(f"  {'TOTAL':<22}: {total:>9,.0f}t")
print()
print("Prochaine étape: python optimizer_v2.py")