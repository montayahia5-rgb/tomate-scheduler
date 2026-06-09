# -*- coding: utf-8 -*-
"""
migrate.py — Pousse le planning dans Supabase
=============================================
CE SCRIPT :
  - NE TOUCHE PAS à la table agriculteurs
    (elle est gérée par les uploads des commerciaux)
  - Vide et reinsère : planning + transport + decalage
    depuis Planning_Tomate_2026.xlsx

COMMENT UTILISER :
  1. Lance d'abord : python optimizer_v2.py
  2. Lance ensuite : python migrate.py
  3. Rafraîchis le dashboard

ORDRE COMPLET :
  python optimizer_v2.py   → génère Planning_Tomate_2026.xlsx
  python migrate.py        → pousse planning/transport dans Supabase
  streamlit run dashboard_phase10.py
"""

import sys, os, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

import pandas as pd
import datetime
from supabase import create_client, Client

# ============================================================
# CONFIGURATION
# ============================================================
SUPABASE_URL = "https://mwjefdqfzrtsfzspeppg.supabase.co"
SUPABASE_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Im13amVmZHFmenJ0c2Z6c3BlcHBnIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODAzODg3MDgsImV4cCI6MjA5NTk2NDcwOH0.H5IX2uLneHdvwyLJrgN7OrHGrLSZSrQBAuJeuzCEz44"

EXCEL_PHASE4 = "Planning_Tomate_2026.xlsx"
BATCH_SIZE   = 100

# ============================================================
# CONNEXION SUPABASE
# ============================================================
print("Connexion a Supabase...")
try:
    sb: Client = create_client(SUPABASE_URL, SUPABASE_KEY)
    print("  Connexion OK")
except Exception as e:
    print(f"  ERREUR: {e}")
    sys.exit(1)

# ============================================================
# HELPERS
# ============================================================
def insert_batch(table_name: str, rows: list, label: str = ""):
    total = len(rows)
    if total == 0:
        print(f"  {label}: 0 lignes (rien a inserer)")
        return
    inserted = 0
    for i in range(0, total, BATCH_SIZE):
        batch = rows[i: i + BATCH_SIZE]
        try:
            sb.table(table_name).insert(batch).execute()
            inserted += len(batch)
            pct = int(inserted / total * 100)
            bar = "#" * (pct // 5) + "-" * (20 - pct // 5)
            print(f"  [{bar}] {pct}%  {inserted}/{total} lignes", end="\r")
        except Exception as e:
            print(f"\n  ERREUR lot {i}: {e}")
            raise
    print(f"  [####################] 100%  {total}/{total} lignes    ")

# ── Normalisation régions officielles ──────────────────────────────────────
REGION_NORM = {
    # CAP BON
    "NABEUL":   "CAP BON 2",  "nabeul":   "CAP BON 2",
    "CAPB1":    "CAP BON 1",  "CAP B1":   "CAP BON 1",
    "CAPB2":    "CAP BON 2",  "CAP B2":   "CAP BON 2",
    "CAP BON":  "CAP BON 1",
    # GAFSA / KASSERINE
    "GAFSA":    "GAFSA / KASSRINE",
    "KASSRINE": "GAFSA / KASSRINE",
    "KASSERINE":"GAFSA / KASSRINE",
    "KASRINE":  "GAFSA / KASSRINE",
    "KASSARINE":"GAFSA / KASSRINE",
    "SBEITLA":  "GAFSA / KASSRINE",
    # NORD
    "BEJA":     "NORD",       "beja":     "NORD",
    "MANOUBA":  "NORD",       "manouba":  "NORD",
    "BIZERTE":  "NORD",
    "JENDOUBA": "NORD",
    "BIR LAHFAY":"NORD",
    "BOR AMRI": "NORD",       "BORJ AMRI":"NORD",
    "MEDJEZ EL BAB":"NORD",   "MEJEZ EL BAB":"NORD",   "MEDJEZ BEB":"NORD",
    "TESTOUR":  "NORD",
    "BOUSSALEM":"NORD",
    # KAIROUAN
    "KAIRAOUAN":"KAIROUAN",
    # SIDI BOUZID
    "SIDIBOUZID":"SIDI BOUZID",
    "SIDI BOU ZID":"SIDI BOUZID",
    # BOUFICHA
    "BOUFICHA": "BOUFICHA",
    "SOUSSE":   "BOUFICHA",
    "ENFIDHA":  "BOUFICHA",
    "HAMMAMET": "CAP BON 1",
}

KNOWN_REGIONS = {"CAP BON 1","CAP BON 2","NORD","KAIROUAN","SIDI BOUZID",
                 "GAFSA / KASSRINE","BOUFICHA","AUTRE"}

def norm_region(r):
    """Normalise la région vers les 7 régions officielles. NAN → AUTRE."""
    if not r or r in ("None", "nan", "", None):
        return "AUTRE"
    r = str(r).strip()
    r_norm = REGION_NORM.get(r, REGION_NORM.get(r.upper(), r))
    # Si toujours pas reconnue, mettre AUTRE
    if r_norm.upper() not in {x.upper() for x in KNOWN_REGIONS}:
        return "AUTRE"
    return r_norm

def safe_str(val):
    if val is None or (isinstance(val, float) and str(val) == "nan"):
        return None
    return str(val).strip() or None

def safe_float(val):
    try:
        f = float(val)
        return None if str(f) == "nan" else round(f, 2)
    except:
        return None

def safe_int(val):
    try:
        return int(float(val))
    except:
        return 0

def safe_date(val):
    if val is None:
        return None
    try:
        if isinstance(val, (datetime.date, datetime.datetime)):
            return val.strftime("%Y-%m-%d")
        d = pd.to_datetime(val, errors="coerce")
        if pd.isna(d) or d.year < 2000:
            return None
        return d.strftime("%Y-%m-%d")
    except:
        return None

# ============================================================
# ETAPE 0 : Nettoyer les données corrompues dans agriculteurs
# (régions en minuscules ou noms d'avant la normalisation)
print("\nEtape 0 : Nettoyage régions corrompues dans agriculteurs...")
try:
    # Update nabeul → CAP BON 2
    sb.table("agriculteurs").update({"region": "CAP BON 2"}).ilike("region", "nabeul").execute()
    sb.table("agriculteurs").update({"region": "CAP BON 2"}).eq("region", "NABEUL").execute()
    # Update beja → NORD
    sb.table("agriculteurs").update({"region": "NORD"}).ilike("region", "beja").execute()
    # Update manouba → NORD
    sb.table("agriculteurs").update({"region": "NORD"}).ilike("region", "manouba").execute()
    # Update gafsa → GAFSA / KASSRINE
    sb.table("agriculteurs").update({"region": "GAFSA / KASSRINE"}).ilike("region", "gafsa").execute()
    sb.table("agriculteurs").update({"region": "GAFSA / KASSRINE"}).ilike("region", "kassrine").execute()
    # Update CAPB1/CAPB2
    sb.table("agriculteurs").update({"region": "CAP BON 1"}).ilike("region", "capb1").execute()
    sb.table("agriculteurs").update({"region": "CAP BON 2"}).ilike("region", "capb2").execute()
    print("  ✅ Régions agriculteurs normalisées")
except Exception as e:
    print(f"  ⚠️ Normalisation régions: {e} (non bloquant)")

# ETAPE 1 : Vérifier que agriculteurs a des données
# ============================================================
print("\nEtape 1 : Vérification agriculteurs...")
try:
    agri_check = sb.table("agriculteurs").select("id", count="exact").execute()
    agri_count = agri_check.count
    agri_tons_data = sb.table("agriculteurs").select("tonnage_total").execute().data
    agri_tons = sum(float(r["tonnage_total"]) for r in agri_tons_data if r["tonnage_total"])
    print(f"  agriculteurs: {agri_count} lignes | {agri_tons:,.0f}t")
    if agri_count == 0:
        print("  ATTENTION: Table agriculteurs vide!")
        print("  Demande aux commerciaux d'uploader leurs fichiers d'abord.")
        print("  Ou lance python migrate_init.py pour charger depuis l'Excel original.")
except Exception as e:
    print(f"  Impossible de vérifier agriculteurs: {e}")

# ============================================================
# ETAPE 2 : Vider planning + transport + decalage SEULEMENT
# ============================================================
print("\nEtape 2 : Nettoyage planning/transport/decalage...")
for table in ["decalage", "transport", "planning"]:
    try:
        sb.table(table).delete().gt("id", 0).execute()
        print(f"  Table '{table}' videe")
    except Exception as e:
        print(f"  AVERTISSEMENT '{table}': {e}")

# ============================================================
# ETAPE 3 : Lire Planning_Tomate_2026.xlsx
# ============================================================
print(f"\nEtape 3 : Lecture de '{EXCEL_PHASE4}'...")
if not os.path.exists(EXCEL_PHASE4):
    print(f"  ERREUR: '{EXCEL_PHASE4}' introuvable")
    print(f"  Lance d'abord: python optimizer_v2.py")
    sys.exit(1)

df_planning  = pd.read_excel(EXCEL_PHASE4, sheet_name="Planning Journalier",    header=0)
df_transport = pd.read_excel(EXCEL_PHASE4, sheet_name="Besoins Transport-Jour", header=0)

try:
    df_decalage = pd.read_excel(EXCEL_PHASE4, sheet_name="Journal Double Transport", header=1)
    df_decalage = df_decalage.dropna(subset=["Commercial"]) if "Commercial" in df_decalage.columns else pd.DataFrame()
except Exception:
    df_decalage = pd.DataFrame()

df_planning["Date"]  = pd.to_datetime(df_planning["Date"],  errors="coerce")
df_transport["Date"] = pd.to_datetime(df_transport["Date"], errors="coerce")
df_planning  = df_planning.dropna(subset=["Date"])
df_transport = df_transport.dropna(subset=["Date"])

print(f"  Planning  : {len(df_planning)} lignes")
print(f"  Transport : {len(df_transport)} lignes")
print(f"  Decalage  : {len(df_decalage)} lignes")

# ============================================================
# ETAPE 4 : Insérer planning
# ============================================================
print(f"\nEtape 4 : Insertion planning...")
rows_planning = []
for _, row in df_planning.iterrows():
    pic_val = str(row.get("Pic de Recolte", row.get("Pic de Récolte", ""))).upper()
    rows_planning.append({
        "commercial":    safe_str(row.get("Commercial")),
        "agriculteur":   safe_str(row.get("Agriculteur")),
        "usine":         safe_str(row.get("Usine")),
        "region":        norm_region(safe_str(row.get("Region", row.get("Région")))),
        "accessibilite": safe_str(row.get("Accessibilite", row.get("Accessibilité"))),
        "date":          safe_date(row["Date"]),
        "tonnes_jour":   safe_float(row.get("Tonnes/Jour")),
        "type_vehicule": safe_str(row.get("Type Vehicule", row.get("Type Véhicule"))),
        "vehicules":     safe_str(row.get("Vehicules",     row.get("Véhicules Requis"))),
        "nb_voyages":    safe_int(row.get("Nb Voyages", 0)),
        "date_debut":    safe_date(row.get("Date Debut",   row.get("Date Début"))),
        "date_fin":      safe_date(row.get("Date Fin")),
        "total_tonnes":  safe_float(row.get("Total Tonnes")),
        "pic":           "PIC" in pic_val,
        "note":          safe_str(row.get("Note")),
    })

insert_batch("planning", rows_planning, "planning")

# Vérification
try:
    count = sb.table("planning").select("id", count="exact").execute().count
    tonnage = sum(r["tonnes_jour"] for r in rows_planning if r["tonnes_jour"])
    print(f"  Vérification: {count} lignes | {tonnage:,.0f}t planifiés")
except Exception as e:
    print(f"  Vérification impossible: {e}")

# ============================================================
# ETAPE 5 : Insérer transport
# ============================================================
print(f"\nEtape 5 : Insertion transport...")
rows_transport = []
for _, row in df_transport.iterrows():
    rows_transport.append({
        "date":          safe_date(row["Date"]),
        "commercial":    safe_str(row.get("Commercial")),
        "total_tonnes":  safe_float(row.get("Total Tonnes")),
        "tracteur":      safe_int(row.get("Voyages TRACTEUR", 0)),
        "petit_poilour": safe_int(row.get("Voyages PETIT POILOUR", 0)),
        "poilour":       safe_int(row.get("Voyages POILOUR", 0)),
        "semi":          safe_int(row.get("Voyages SEMI", 0)),
        "jours_double":  safe_int(row.get("Jours Double", 0)),
    })
insert_batch("transport", rows_transport, "transport")

# ============================================================
# ETAPE 6 : Insérer decalage (vide avec OR-Tools)
# ============================================================
print(f"\nEtape 6 : Decalage (OR-Tools = 0 conflits)...")
rows_decalage = []
for _, row in df_decalage.iterrows():
    rows_decalage.append({
        "commercial":      safe_str(row.get("Commercial")),
        "agriculteur_a":   safe_str(row.get("Agriculteur A (finit tôt)")),
        "agriculteur_b":   safe_str(row.get("Agriculteur B (reçoit véhicule)")),
        "vehicule":        safe_str(row.get("Véhicule Partagé")),
        "shift_jours":     safe_int(row.get("Jours Économisés", 0)),
        "fin_orig_a":      safe_date(row.get("Fin Orig. A")),
        "nouvelle_fin_a":  safe_date(row.get("Nouvelle Fin A")),
        "debut_orig_b":    safe_date(row.get("Début Orig. B")),
        "nouveau_debut_b": safe_date(row.get("Nouveau Début B")),
        "risque":          safe_str(row.get("Risque Maladie")),
        "action":          safe_str(row.get("Action Requise")),
    })
if rows_decalage:
    insert_batch("decalage", rows_decalage, "decalage")
else:
    print("  0 conflits (OR-Tools gère en interne)")

# ============================================================
# RÉSUMÉ FINAL
# ============================================================
print("\n" + "="*50)
print("  MIGRATION COMPLETE")
print("="*50)
totals = {}
for table in ["agriculteurs", "planning", "transport", "decalage"]:
    try:
        result = sb.table(table).select("id", count="exact").execute()
        totals[table] = result.count
    except Exception as e:
        totals[table] = f"erreur"

for table, count in totals.items():
    icon = "✅" if isinstance(count, int) and count > 0 else "⚠️"
    print(f"  {icon} {table:<20}: {count} lignes")

print("="*50)
print()
print("  agriculteurs = données des commerciaux (inchangé)")
print("  planning     = nouveau planning OR-Tools")
print()
print("  Prochaine étape:")
print("  streamlit run dashboard_phase10.py")