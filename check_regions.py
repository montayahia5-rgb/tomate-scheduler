# -*- coding: utf-8 -*-
"""
check_regions.py — Diagnostic complet des régions dans Supabase
================================================================
Identifie pourquoi l'onglet "Tonnage par Région" du dashboard est vide.

Cause typique:
- Le dashboard attend exactement: CAP BON 1, CAP BON 2, NORD, GAFSA / KASSRINE,
  KAIROUAN, SIDI BOUZID, BOUFICHA
- Mais Supabase peut contenir: "CAP BON" (sans numéro), "" (vide),
  "Cap Bon 1" (casse différente), etc.
"""
import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
from supabase import create_client
from collections import Counter

SUPABASE_URL = "https://mwjefdqfzrtsfzspeppg.supabase.co"
SUPABASE_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Im13amVmZHFmenJ0c2Z6c3BlcHBnIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODAzODg3MDgsImV4cCI6MjA5NTk2NDcwOH0.H5IX2uLneHdvwyLJrgN7OrHGrLSZSrQBAuJeuzCEz44"

sb = create_client(SUPABASE_URL, SUPABASE_KEY)

REG_ORD = ["CAP BON 1","CAP BON 2","NORD","GAFSA / KASSRINE",
           "KAIROUAN","SIDI BOUZID","BOUFICHA"]

print("="*60)
print("  DIAGNOSTIC RÉGIONS DANS SUPABASE")
print("="*60)

# ── 1. Toutes les régions présentes ────────────────────────
data = sb.table("agriculteurs").select(
    "commercial,nom,region,zone,usine,tonnage_total"
).execute().data

print(f"\n[1] TOTAL: {len(data)} lignes dans agriculteurs\n")

# Compteur régions exactes
regions_raw = Counter()
tonnage_par_region_raw = {}
for r in data:
    reg = (r.get("region") or "").strip()
    t   = float(r.get("tonnage_total") or 0)
    regions_raw[reg] += 1
    tonnage_par_region_raw[reg] = tonnage_par_region_raw.get(reg, 0) + t

print("[2] RÉGIONS EXACTES PRÉSENTES (telles que stockées):")
print(f"  {'Région (brut)':<30} {'Lignes':>7} {'Tonnage':>10} {'Statut':>15}")
print("  " + "-"*70)
total_match = 0
total_no_match = 0
for reg, count in sorted(regions_raw.items(), key=lambda x:-x[1]):
    ton = tonnage_par_region_raw.get(reg, 0)
    status = "✅ MATCH dashboard" if reg in REG_ORD else "❌ NON MATCH"
    reg_display = f"'{reg}'" if reg else "(VIDE)"
    print(f"  {reg_display:<30} {count:>7} {ton:>9,.0f}t {status:>15}")
    if reg in REG_ORD:
        total_match += ton
    else:
        total_no_match += ton

print()
print(f"  ✅ Tonnage MATCH dashboard: {total_match:>10,.0f}t")
print(f"  ❌ Tonnage NON MATCH:      {total_no_match:>10,.0f}t")

# ── 3. Suggestions de mapping ──────────────────────────────
print("\n[3] MAPPING SUGGÉRÉ POUR NORMALISER:")
SUGGESTIONS = {
    "CAP BON":          "CAP BON 1",  # par défaut → CAP BON 1
    "CAPBON 1":         "CAP BON 1",
    "CAPBON 2":         "CAP BON 2",
    "CAP BON1":         "CAP BON 1",
    "CAP BON2":         "CAP BON 2",
    "CAPBON":           "CAP BON 1",
    "NABEUL":           "CAP BON 2",
    "KORBA":            "CAP BON 1",
    "BEJA":             "NORD",
    "MANOUBA":          "NORD",
    "JANDOUBA":         "NORD",
    "TUNIS":            "NORD",
    "GAFSA":            "GAFSA / KASSRINE",
    "KASSRINE":         "GAFSA / KASSRINE",
    "KASSERINE":        "GAFSA / KASSRINE",
    "GAFSA/KASSRINE":   "GAFSA / KASSRINE",
    "SIDI BOUZID":      "SIDI BOUZID",
    "KAIROUAN":         "KAIROUAN",
    "SOUSSE":           "BOUFICHA",
    "BOUFICHA":         "BOUFICHA",
}
non_match_regions = [r for r in regions_raw if r and r not in REG_ORD]
fixable = []
unknown = []
for reg in non_match_regions:
    suggestion = SUGGESTIONS.get(reg.upper().strip(), None)
    if suggestion:
        fixable.append((reg, suggestion, regions_raw[reg], tonnage_par_region_raw.get(reg, 0)))
        print(f"  '{reg}' → '{suggestion}' ({regions_raw[reg]} lignes, {tonnage_par_region_raw.get(reg, 0):,.0f}t)")
    else:
        unknown.append((reg, regions_raw[reg], tonnage_par_region_raw.get(reg, 0)))

if unknown:
    print("\n  ⚠️ Régions NON RECONNUES (à mapper manuellement):")
    for reg, count, ton in unknown:
        print(f"    '{reg}' ({count} lignes, {ton:,.0f}t)")

# ── 4. Correction automatique ──────────────────────────────
if fixable or unknown:
    print("\n[4] APPLIQUER LA CORRECTION ?")
    print("  Lance: python check_regions.py --fix")

if "--fix" in sys.argv:
    print("\n[FIX] CORRECTION DES RÉGIONS...")
    for old_reg, new_reg, count, ton in fixable:
        sb.table("agriculteurs").update({"region": new_reg}).eq("region", old_reg).execute()
        print(f"  ✅ '{old_reg}' → '{new_reg}' ({count} lignes)")
    
    # Aussi: pour les vides, essayer de déduire depuis la zone
    print("\n  Tentative de déduction depuis 'zone' pour les régions vides...")
    ZONE_TO_REGION = {
        # CAP BON 1 (Cap Bon nord-est)
        "DAR ALLOUCH": "CAP BON 1", "KORBA": "CAP BON 1", "SOMAA": "CAP BON 1",
        "LEBNA": "CAP BON 1", "DIAR HOJJEJ": "CAP BON 1", "HTOUBA": "CAP BON 1",
        "ATHLETH": "CAP BON 1", "TEFELOUN": "CAP BON 1", "KHADHRA": "CAP BON 1",
        "SIDI KHELIFA": "CAP BON 1", "MENZEL HORR": "CAP BON 1",
        "BIR LAHFAY": "CAP BON 1", "OUED CHIBA": "CAP BON 1",
        "GOURCHIN": "CAP BON 1", "MENZEL TAMIM": "CAP BON 1",
        # CAP BON 2 (Cap Bon sud)
        "MENZEL MHIRI": "CAP BON 2", "GROMBELIA": "CAP BON 2",
        "SIDI AICH": "CAP BON 2", "SIDI OTHMAN": "CAP BON 2",
        "NABEUL": "CAP BON 2",
        # NORD
        "FARTOUNA": "NORD", "GAR DIMAOU": "NORD", "JANDOUBA": "NORD",
        "BOU SALEM": "NORD", "WED MLIZ": "NORD", "SIDI HASSOUN": "NORD",
        "AMAYMIA": "NORD", "ZAAFRIA": "NORD", "BOUJRIDA": "NORD",
        "OUED KHATEF": "NORD", "MEDJEZ BEB": "NORD",
        "BOR AMRI": "NORD", "AWAMRIYA": "NORD", "FRININ": "NORD",
        "GOMBAR": "NORD", "BENI AYECH": "NORD",
        # KAIROUAN
        "SBIKHA": "KAIROUAN", "CHRARDA": "KAIROUAN", "MAJEL BELABESS": "KAIROUAN",
        "CHEBIKA": "KAIROUAN", "HAWEREB": "KAIROUAN", "HAFOUZ": "KAIROUAN",
        "OULED ZID": "KAIROUAN", "BATTEN": "KAIROUAN", "GARAT SASSI": "KAIROUAN",
        # SIDI BOUZID
        "SIDI BOUZID": "SIDI BOUZID", "OM ADHAM": "SIDI BOUZID", "TBAG": "SIDI BOUZID",
        # GAFSA / KASSRINE
        "FERIANA": "GAFSA / KASSRINE", "GAFSA": "GAFSA / KASSRINE",
        # BOUFICHA
        "BOUFICHA": "BOUFICHA",
    }
    deduits = 0
    rows_empty = [r for r in data if not (r.get("region") or "").strip()]
    for r in rows_empty:
        zone = str(r.get("zone","") or "").strip().upper()
        for zone_key, region_val in ZONE_TO_REGION.items():
            if zone_key in zone:
                # Mettre à jour ce row spécifiquement par id
                sb.table("agriculteurs").update({"region": region_val}).eq("id", r["id"]).execute()
                deduits += 1
                break
    print(f"  ✅ {deduits} régions déduites depuis 'zone'")

# ── 5. Vérification finale ─────────────────────────────────
print("\n[5] VÉRIFICATION FINALE")
data_v = sb.table("agriculteurs").select("region,tonnage_total").execute().data
final = Counter()
final_ton = {}
for r in data_v:
    reg = (r.get("region") or "").strip()
    final[reg] += 1
    final_ton[reg] = final_ton.get(reg, 0) + float(r.get("tonnage_total") or 0)

print(f"\n  {'Région finale':<25} {'Lignes':>7} {'Tonnage':>10} {'Dashboard':>15}")
print("  " + "-"*65)
for reg in REG_ORD + [r for r in final if r not in REG_ORD and r]:
    count = final.get(reg, 0)
    ton = final_ton.get(reg, 0)
    if count > 0:
        status = "✅ Visible" if reg in REG_ORD else "❌ Invisible"
        print(f"  {reg:<25} {count:>7} {ton:>9,.0f}t {status:>15}")

vide = final.get("", 0)
if vide:
    ton_vide = final_ton.get("", 0)
    print(f"  {'(VIDE)':<25} {vide:>7} {ton_vide:>9,.0f}t {'❌ Invisible':>15}")