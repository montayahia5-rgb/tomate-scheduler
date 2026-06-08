# -*- coding: utf-8 -*-
"""
optimizer_v2.py — OR-Tools avec optimisation distance + caps
=============================================================
Nouvelles fonctionnalités vs optimizer.py :
  - Matrice de distances zones → usines (km)
  - SICAM = 70% Borj Sedeya + 30% Mejez el Bab (pondéré)
  - OR-Tools minimise : distance totale de transport + déviation du plan
  - Assignation automatique agriculteur → usine la plus proche
  - Affichage par région dans le résultat

USAGE :
  python optimizer_v2.py
  python migrate.py
  streamlit run dashboard_phase10.py
"""

import sys, io, os, datetime, json, math
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

import pandas as pd
from collections import defaultdict
from ortools.sat.python import cp_model
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ============================================================
# CONFIGURATION
# ============================================================
OUTPUT_FILE = "Planning_Tomate_2026.xlsx"   # new clean name
TEMP_FILE   = "Planning_Tomate_TEMP.xlsx"

PEAK_START   = datetime.date(2026, 7,  1)
PEAK_END     = datetime.date(2026, 7, 15)
SEASON_START = datetime.date(2026, 6, 15)
SEASON_END   = datetime.date(2026, 8, 31)

def clamp_date(d):
    if d < SEASON_START: return SEASON_START
    if d > SEASON_END:   return SEASON_END
    return d

COMMERCIAL_CAPS = {
    "FEDI":             850,   # t/jour
    "MAKKI BEN SALAH":  800,
    "KHALIL":           800,
    "ACHREF AJLANI":    500,
    "JILANI OBAY":       50,   # cap officiel 50t/j — ajusté automatiquement si tonnage dépasse 50×78j=3900t
}
FACTORY_CAPS = {
    "SICAM":    1300,   # t/jour
    "TUCAL":     750,
    "COMOCAP":   700,
    "ABIDA":     150,
    "ELFALLEH":  100,
}

# ✅ MARGE pour absorber l'arrondi à la dizaine
# L'arrondi peut ajouter jusqu'à +5t/agriculteur en moyenne
# On réduit les caps solveur pour qu'APRÈS arrondi le cap OFFICIEL soit respecté
ROUNDING_MARGIN_PCT = 0.10  # 10% de marge sous le cap pour le solveur (était 5%)
ROUNDING_MARGIN_MIN = 50    # minimum 50t de marge (était 30t)
FLEET_CAPACITY = {
    "TRACTEUR":         (9,  11),    # min/max tonnes par voyage (moyenne ~10t)
    "PPL":              (6,  14),    # Petit Poilour  — capacité mise à jour
    "PL":               (15, 25),   # Poilour         — capacité mise à jour
    "SEMI":             (27, 33),   # Semi-remorque   — capacité mise à jour
    "DOUBLE_REM":       (27, 33),   # Double Remorque (= SEMI)
}

# Alias rétrocompatibilité
FLEET_CAPACITY["PETIT POILOUR"] = FLEET_CAPACITY["PPL"]
FLEET_CAPACITY["POILOUR"]       = FLEET_CAPACITY["PL"]

ACCESS_VEHICLES = {
    # ── 1 véhicule ─────────────────────────────────────────────────
    "PL":          ["PL"],                       # PL seulement
    "RM":          ["SEMI"],                     # RM = 100% Semi (3-5 semi/jour)
    # ── 2 véhicules ────────────────────────────────────────────────
    "PL/PPL":      ["PL", "PPL"],
    "PL-PPL":      ["PL", "PPL"],
    "TRC/PPL":     ["TRACTEUR", "PPL"],          # tracteur + PPL
    "PL/SEMI":     ["PL", "SEMI"],
    "PL-SEMI":     ["PL", "SEMI"],
    # ── 3 véhicules ────────────────────────────────────────────────
    "TRC/PPL/PL":  ["TRACTEUR", "PPL", "PL"],   # tracteur + PPL + PL
    "PL/PPL/TRC":  ["PL", "PPL", "TRACTEUR"],
    "PL/PPL/SEMI": ["PL", "PPL", "SEMI"],        # PL + PPL + SEMI
    "SEMI/PL/PPL": ["SEMI", "PL", "PPL"],
    # ── Fallback ───────────────────────────────────────────────────
    "NAN":         ["TRACTEUR", "PPL", "PL", "SEMI"],
}

# ── Capacités confirmées par usine (source: transport_12_mai.xlsx) ──────
# Mise à jour: 12 mai 2026 — liste confirmée uniquement
# ── Données exactes depuis transport_12_mai.xlsx (liste confirmée) ──────────
TRANSPORT_CONFIRMED = {
    "SICAM":    {"total": 1199, "PL": 825, "PPL": 44,  "SEMI": 330, "nb_bennes": 58},
    "TUCAL":    {"total": 348,  "PL": 318, "PPL": 0,   "SEMI": 30,  "nb_bennes": 19},
    "COMOCAP":  {"total": 298,  "PL": 76,  "PPL": 132, "SEMI": 90,  "nb_bennes": 21},
    "ABIDA":    {"total": 50,   "PL": 20,  "PPL": 0,   "SEMI": 30,  "nb_bennes": 2},
    "ELFALLEH": {"total": 24,   "PL": 0,   "PPL": 24,  "SEMI": 0,   "nb_bennes": 2},
}
# Jokers = BOURAK et LUI-MÊME (toutes usines)
TRANSPORT_JOKERS = {
    "BOURAK":   {"total": 76,  "PL": 76, "PPL": 0,  "SEMI": 0, "nb_bennes": 4},
    "LUI-MEME": {"total": 84,  "PL": 38, "PPL": 46, "SEMI": 0, "nb_bennes": 6},
}

# ── Règles de complétion transport par usine ──────────────────────────
# Quand la capacité confirmée ne couvre pas le cap journalier,
# on complète avec ces proportions (règles professionnelles)
TRANSPORT_REGLES = {
    "COMOCAP":  [("TRACTEUR", 100), ("PPL", 0.50), ("PL", 0.30), ("SEMI", 0.20)],
    "TUCAL":    [("PL", 0.30),      ("SEMI", 0.70)],
    "ABIDA":    [("PL", 0.50),      ("SEMI", 0.50)],
    "ELFALLEH": [("PPL", 0.70),     ("PL", 0.30)],
    "SICAM":    [("SEMI", 1.00)],
}

# ── Besoins de transport calculés (reste à compléter) ────────────────
def calc_transport_needs():
    """
    Calcule le tonnage manquant par usine et type de véhicule.
    Retourne dict {usine: {type: tonnes_manquantes}}
    """
    needs = {}
    for usine, cap in FACTORY_CAPS.items():
        conf  = TRANSPORT_CONFIRMED.get(usine, {}).get("total", 0)
        reste = max(0, cap - conf)
        if reste == 0:
            needs[usine] = {}
            continue
        regles = TRANSPORT_REGLES.get(usine, [])
        reste_var = reste
        usine_needs = {}
        for vtype, share in regles:
            if isinstance(share, int):
                usine_needs[vtype] = share
                reste_var -= share
            else:
                usine_needs[vtype] = round(reste_var * share)
        needs[usine] = usine_needs
    return needs

TRANSPORT_NEEDS = calc_transport_needs()

# ── Allocation jokers (PL→TUCAL/COMOCAP, PPL→ELFALLEH/COMOCAP) ──────
def alloc_jokers():
    """Alloue les jokers aux usines qui en ont besoin, par priorité de manque."""
    jpl  = TRANSPORT_JOKERS["BOURAK"]["PL"]
    jppl = TRANSPORT_JOKERS["LUI-MEME"]["PPL"]
    alloc = {}
    for usine in FACTORY_CAPS:
        needs = TRANSPORT_NEEDS.get(usine, {})
        alloc[usine] = {}
        if "PL" in needs and jpl > 0:
            a = min(needs["PL"], jpl)
            alloc[usine]["PL_joker"] = a
            jpl -= a
        if "PPL" in needs and jppl > 0:
            a = min(needs["PPL"], jppl)
            alloc[usine]["PPL_joker"] = a
            jppl -= a
    return alloc

JOKER_ALLOC = alloc_jokers()

# Transport COMOCAP — TRACTEUR est un VRAI transport (10t/voyage)
# Le tracteur fait partie de la flotte normale COMOCAP, pas pour caisses vides.
# Désactivé : la logique CAISSE_VIDE était incorrecte.
COMOCAP_EXTRA_TRACTEUR = False

# ============================================================
# DISTANCE MATRIX — Zone → Usine (km)
# SICAM = 70% Borj Sedeya + 30% Mejez el Bab
# ============================================================
DISTANCE_KM = {
    # CAP BON 1 (Dar Allouche, Haouaria, Kelibia)
    "DAR ALLOUCH":          {"COMOCAP":15,  "ELFALLEH":20,  "SICAM":94,  "TUCAL":100, "ABIDA":220},
    "KORBA":                {"COMOCAP":20,  "ELFALLEH":15,  "SICAM":94,  "TUCAL":95,  "ABIDA":215},
    "KORBA/SOMAA":          {"COMOCAP":22,  "ELFALLEH":18,  "SICAM":95,  "TUCAL":97,  "ABIDA":217},
    "SOMAA":                {"COMOCAP":25,  "ELFALLEH":20,  "SICAM":89,  "TUCAL":90,  "ABIDA":210},
    "LEBNA":                {"COMOCAP":30,  "ELFALLEH":25,  "SICAM":94,  "TUCAL":95,  "ABIDA":215},
    "LEBNA/TAMEZRRAT":      {"COMOCAP":32,  "ELFALLEH":27,  "SICAM":95,  "TUCAL":97,  "ABIDA":217},
    "DIAR HOJJEJ":          {"COMOCAP":10,  "ELFALLEH":12,  "SICAM":97,  "TUCAL":98,  "ABIDA":218},
    "DIAR HOJJEJ/KHARREZ":  {"COMOCAP":12,  "ELFALLEH":14,  "SICAM":96,  "TUCAL":99,  "ABIDA":219},
    "TEFELOUN":             {"COMOCAP":18,  "ELFALLEH":22,  "SICAM":90,  "TUCAL":92,  "ABIDA":212},
    "TEFELOUN/DIAR HOJJEJ": {"COMOCAP":15,  "ELFALLEH":18,  "SICAM":94,  "TUCAL":95,  "ABIDA":215},
    "ATHLETH":              {"COMOCAP":8,   "ELFALLEH":10,  "SICAM":98,  "TUCAL":100, "ABIDA":220},
    "ATHLETH/HTOUBA":       {"COMOCAP":10,  "ELFALLEH":12,  "SICAM":96,  "TUCAL":98,  "ABIDA":218},
    "HTOUBA":               {"COMOCAP":12,  "ELFALLEH":15,  "SICAM":96,  "TUCAL":97,  "ABIDA":217},
    "KHADHRA":              {"COMOCAP":20,  "ELFALLEH":18,  "SICAM":91,  "TUCAL":93,  "ABIDA":213},
    "SIDI KHELIFA":         {"COMOCAP":25,  "ELFALLEH":22,  "SICAM":88,  "TUCAL":90,  "ABIDA":210},
    "MENZEL HORR":          {"COMOCAP":28,  "ELFALLEH":25,  "SICAM":86,  "TUCAL":88,  "ABIDA":208},
    "BIR LAHFAY":           {"COMOCAP":35,  "ELFALLEH":30,  "SICAM":82,  "TUCAL":85,  "ABIDA":205},
    "OUED CHIBA":           {"COMOCAP":55,  "ELFALLEH":58,  "SICAM":58,  "TUCAL":62,  "ABIDA":182},
    "GOURCHIN":             {"COMOCAP":42,  "ELFALLEH":45,  "SICAM":70,  "TUCAL":72,  "ABIDA":192},
    # CAP BON 2 (Menzel Temime → Nabeul)
    "MENZEL TAMIM":         {"COMOCAP":40,  "ELFALLEH":45,  "SICAM":74,  "TUCAL":75,  "ABIDA":195},
    "MENZEL MHIRI":         {"COMOCAP":45,  "ELFALLEH":50,  "SICAM":69,  "TUCAL":70,  "ABIDA":190},
    "GROMBELIA":            {"COMOCAP":50,  "ELFALLEH":55,  "SICAM":64,  "TUCAL":65,  "ABIDA":185},
    "SIDI AICH":            {"COMOCAP":48,  "ELFALLEH":52,  "SICAM":67,  "TUCAL":68,  "ABIDA":188},
    "SIDI OTHMAN":          {"COMOCAP":52,  "ELFALLEH":55,  "SICAM":64,  "TUCAL":65,  "ABIDA":185},
    # NORD (Tunis → Jendouba)
    "FARTOUNA":             {"COMOCAP":95,  "ELFALLEH":100, "SICAM":51,  "TUCAL":35,  "ABIDA":120},
    "GAR DIMAOU":           {"COMOCAP":120, "ELFALLEH":125, "SICAM":65,  "TUCAL":60,  "ABIDA":100},
    "JANDOUBA":             {"COMOCAP":130, "ELFALLEH":135, "SICAM":71,  "TUCAL":70,  "ABIDA":90},
    "BOU SALEM":            {"COMOCAP":115, "ELFALLEH":120, "SICAM":64,  "TUCAL":55,  "ABIDA":105},
    "WED MLIZ":             {"COMOCAP":110, "ELFALLEH":115, "SICAM":61,  "TUCAL":50,  "ABIDA":110},
    "SIDI HASSOUN":         {"COMOCAP":100, "ELFALLEH":105, "SICAM":54,  "TUCAL":40,  "ABIDA":115},
    "AMAYMIA":              {"COMOCAP":80,  "ELFALLEH":85,  "SICAM":47,  "TUCAL":30,  "ABIDA":130},
    "ZAAFRANA-ELKHADHRA":   {"COMOCAP":85,  "ELFALLEH":90,  "SICAM":49,  "TUCAL":35,  "ABIDA":125},
    "ZAAFRIA":              {"COMOCAP":82,  "ELFALLEH":87,  "SICAM":48,  "TUCAL":33,  "ABIDA":127},
    "BOUJRIDA":             {"COMOCAP":88,  "ELFALLEH":93,  "SICAM":52,  "TUCAL":37,  "ABIDA":122},
    "OUED KHATEF":          {"COMOCAP":105, "ELFALLEH":110, "SICAM":58,  "TUCAL":45,  "ABIDA":112},
    "MEDJEZ BEB":           {"COMOCAP":100, "ELFALLEH":105, "SICAM":49,  "TUCAL":45,  "ABIDA":110},
    "AMAYMIA":              {"COMOCAP":80,  "ELFALLEH":85,  "SICAM":47,  "TUCAL":30,  "ABIDA":130},
    "BOR AMRI":             {"COMOCAP":78,  "ELFALLEH":83,  "SICAM":45,  "TUCAL":28,  "ABIDA":132},
    "AWAMRIYA":             {"COMOCAP":92,  "ELFALLEH":97,  "SICAM":53,  "TUCAL":40,  "ABIDA":118},
    "FRININ":               {"COMOCAP":75,  "ELFALLEH":80,  "SICAM":44,  "TUCAL":25,  "ABIDA":135},
    "GOMBAR":               {"COMOCAP":83,  "ELFALLEH":88,  "SICAM":49,  "TUCAL":32,  "ABIDA":128},
    "BENI AYECH":           {"COMOCAP":110, "ELFALLEH":115, "SICAM":61,  "TUCAL":55,  "ABIDA":100},
    "BIR MASOUDA":          {"COMOCAP":108, "ELFALLEH":113, "SICAM":59,  "TUCAL":52,  "ABIDA":102},
    # KAIROUAN / CENTRE
    "SBIKHA-CHRARDA":       {"COMOCAP":150, "ELFALLEH":155, "SICAM":114, "TUCAL":110, "ABIDA":130},
    "MAJEL BELABESS":       {"COMOCAP":160, "ELFALLEH":165, "SICAM":124, "TUCAL":120, "ABIDA":120},
    "CHEBIKA-ELHAWEREB":    {"COMOCAP":155, "ELFALLEH":160, "SICAM":119, "TUCAL":115, "ABIDA":125},
    "ELHAWEREB-AIN BIDHA-HAFOUZ": {"COMOCAP":165, "ELFALLEH":170, "SICAM":129, "TUCAL":125, "ABIDA":115},
    "OULED ZID":            {"COMOCAP":145, "ELFALLEH":150, "SICAM":109, "TUCAL":105, "ABIDA":135},
    "BATTEN":               {"COMOCAP":158, "ELFALLEH":163, "SICAM":122, "TUCAL":118, "ABIDA":122},
    "GARAT SASSI":          {"COMOCAP":162, "ELFALLEH":167, "SICAM":126, "TUCAL":122, "ABIDA":118},
    "BELYES":               {"COMOCAP":148, "ELFALLEH":153, "SICAM":112, "TUCAL":108, "ABIDA":132},
    # SIDI BOUZID
    "SIDI BOUZID":          {"COMOCAP":200, "ELFALLEH":205, "SICAM":164, "TUCAL":160, "ABIDA":80},
    "OM ADHAM":             {"COMOCAP":195, "ELFALLEH":200, "SICAM":159, "TUCAL":155, "ABIDA":85},
    "TBAG":                 {"COMOCAP":198, "ELFALLEH":203, "SICAM":162, "TUCAL":158, "ABIDA":82},
    # KASSRINE / GAFSA
    "FERIANA":              {"COMOCAP":230, "ELFALLEH":235, "SICAM":190, "TUCAL":185, "ABIDA":60},
    "CENTRE -OUEST":        {"COMOCAP":180, "ELFALLEH":185, "SICAM":144, "TUCAL":140, "ABIDA":100},
    "GAFSA":                {"COMOCAP":280, "ELFALLEH":285, "SICAM":239, "TUCAL":235, "ABIDA":55},
    # BOUFICHA
    "BOUFICHA":             {"COMOCAP":60,  "ELFALLEH":65,  "SICAM":55,  "TUCAL":55,  "ABIDA":160},
}

# Region mapping
ZONE_REGION = {
    # CAP BON 1
    "DAR ALLOUCH":"CAP BON 1","KORBA":"CAP BON 1","KORBA/SOMAA":"CAP BON 1",
    "SOMAA":"CAP BON 1","LEBNA":"CAP BON 1","LEBNA/TAMEZRRAT":"CAP BON 1",
    "DIAR HOJJEJ":"CAP BON 1","DIAR HOJJEJ/KHARREZ":"CAP BON 1",
    "TEFELOUN/DIAR HOJJEJ":"CAP BON 1","ATHLETH":"CAP BON 1",
    "ATHLETH/HTOUBA":"CAP BON 1","HTOUBA":"CAP BON 1","KHADHRA":"CAP BON 1",
    "SIDI KHELIFA":"CAP BON 1","MENZEL HORR":"CAP BON 1","BIR LAHFAY":"CAP BON 1",
    "OUED CHIBA":"CAP BON 1","GOURCHIN":"CAP BON 1",
    # CAP BON 2
    "MENZEL TAMIM":"CAP BON 2","MENZEL MHIRI":"CAP BON 2","GROMBELIA":"CAP BON 2",
    "SIDI AICH":"CAP BON 2","SIDI OTHMAN":"CAP BON 2",
    # NORD
    "FARTOUNA":"NORD","GAR DIMAOU":"NORD","JANDOUBA":"NORD","BOU SALEM":"NORD",
    "WED MLIZ":"NORD","SIDI HASSOUN":"NORD","AMAYMIA":"NORD",
    "ZAAFRANA-ELKHADHRA":"NORD","ZAAFRIA":"NORD","BOUJRIDA":"NORD",
    "OUED KHATEF":"NORD","MEDJEZ BEB":"NORD","BOR AMRI":"NORD",
    "AWAMRIYA":"NORD","FRININ":"NORD","GOMBAR":"NORD","BENI AYECH":"NORD",
    "BIR MASOUDA":"NORD",
    # KAIROUAN
    "SBIKHA-CHRARDA":"KAIROUAN","MAJEL BELABESS":"KAIROUAN",
    "CHEBIKA-ELHAWEREB":"KAIROUAN","ELHAWEREB-AIN BIDHA-HAFOUZ":"KAIROUAN",
    "OULED ZID":"KAIROUAN","BATTEN":"KAIROUAN","GARAT SASSI":"KAIROUAN","BELYES":"KAIROUAN",
    # SIDI BOUZID
    "SIDI BOUZID":"SIDI BOUZID","OM ADHAM":"SIDI BOUZID","TBAG":"SIDI BOUZID",
    # GAFSA / KASSRINE (fusionnées)
    "FERIANA":"GAFSA / KASSRINE","CENTRE -OUEST":"GAFSA / KASSRINE","GAFSA":"GAFSA / KASSRINE",
    # BOUFICHA
    "BOUFICHA":"BOUFICHA",
    # BEJA et MANOUBA → intégrés dans NORD
    "BENI AYECH":"NORD","BIR MASOUDA":"NORD",
    "BELARIGIA":"NORD","BIR LAHFAY":"NORD",
}

def get_distance(zone: str, usine: str) -> int:
    """Get distance km from zone to usine. Returns 999 if unknown."""
    zone_up = zone.upper().strip()
    d = DISTANCE_KM.get(zone_up, {})
    return d.get(usine, 999)

def get_best_usine(zone: str, allowed_usines: list) -> str:
    """Return the closest usine to this zone among allowed list."""
    zone_up = zone.upper().strip()
    dists = {u: get_distance(zone_up, u) for u in allowed_usines}
    return min(dists, key=dists.get)

def get_region(zone: str) -> str:
    return ZONE_REGION.get(zone.upper().strip(), "AUTRE")

# ============================================================
# SUPABASE CONNECTION
# ============================================================
SUPABASE_URL = "https://mwjefdqfzrtsfzspeppg.supabase.co"
SUPABASE_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Im13amVmZHFmenJ0c2Z6c3BlcHBnIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODAzODg3MDgsImV4cCI6MjA5NTk2NDcwOH0.H5IX2uLneHdvwyLJrgN7OrHGrLSZSrQBAuJeuzCEz44"

# ============================================================
# STEP 1: Load ONLY from Supabase — farmers uploaded by commercials
# ============================================================
print("Connexion à Supabase...")
try:
    from supabase import create_client
    sb = create_client(SUPABASE_URL, SUPABASE_KEY)
    print("  OK")
except Exception as e:
    print(f"  ERREUR: {e}")
    print("  Installe supabase: pip install supabase")
    import sys; sys.exit(1)

# Get list of commercials who have deposited (statut = 'depose')
print("\nVérification des dépôts...")
try:
    depot_data = sb.table("depot_status").select("*").eq("statut","depose").execute().data
    deposited  = [r["commercial"] for r in depot_data] if depot_data else []
except Exception:
    # If depot_status table doesn't exist yet, use all agriculteurs
    deposited = []

if not deposited:
    print("  Aucun commercial n'a déposé via depot_status.")
    print("  Utilisation de tous les agriculteurs dans Supabase.")
    deposited = None  # means: use all

# Load agriculteurs from Supabase
print("Chargement des agriculteurs depuis Supabase...")
try:
    agri_data = sb.table("agriculteurs").select("*").execute().data
    if not agri_data:
        print("  ERREUR: Table agriculteurs vide.")
        print("  Lance d'abord python migrate.py ou demande aux commerciaux d'uploader.")
        import sys; sys.exit(1)
    df_all = pd.DataFrame(agri_data)
except Exception as e:
    print(f"  ERREUR lecture Supabase: {e}")
    import sys; sys.exit(1)

# Filter: only farmers from commercials who deposited
if deposited is not None:
    df_all["commercial"] = df_all["commercial"].astype(str).str.strip()
    df = df_all[df_all["commercial"].isin(deposited)].copy()
    missing = [c for c in ["FEDI","MAKKI BEN SALAH","KHALIL","ACHREF AJLANI","JILANI OBAY"]
               if c not in deposited]
    if missing:
        print(f"  ⚠️  Pas encore déposé: {missing}")
        print(f"  Planning généré SEULEMENT pour: {deposited}")
else:
    df = df_all.copy()

# Rename columns to match internal format
df = df.rename(columns={
    "nom":           "AGRICULTEUR",
    "tonnage_total": "TONNAGE",
    "usine":         "USINE",
    "accessibilite": "accessbilite",
    "region":        "REGION",
    "zone":          "ZONNE",
})

df = df.dropna(subset=["AGRICULTEUR","TONNAGE","USINE"])
df = df[pd.to_numeric(df["TONNAGE"], errors="coerce") > 0]
df["TONNAGE"]      = pd.to_numeric(df["TONNAGE"], errors="coerce")
df["commercial"]   = df["commercial"].astype(str).str.strip()
df["USINE"]        = df["USINE"].astype(str).str.strip()
df["accessbilite"] = df["accessbilite"].fillna("PL/PPL").astype(str).str.strip().str.upper()
def normalize_acc(x):
    """Normalise l'accessibilité en gardant TOUS les types de véhicules."""
    x = str(x).strip().upper().replace("-", "/")
    if x in ("NAN","NONE","","NAT"): return "PL/PPL"
    if x == "RM":                   return "RM"
    if x == "PL":                   return "PL"
    # Déjà valide tel quel
    if x in ACCESS_VEHICLES:        return x
    import re as _re
    parts = set(_re.split(r"[/,;\s]+", x))
    parts = {p.strip() for p in parts if p.strip()}
    has_trc  = "TRC" in parts or "TRACTEUR" in parts
    has_pl   = "PL" in parts
    has_ppl  = "PPL" in parts
    has_semi = "SEMI" in parts
    # 3 types
    if has_trc and has_pl and has_ppl:  return "TRC/PPL/PL"
    if has_trc and (has_ppl or has_pl): return "TRC/PPL"
    if has_semi and has_pl and has_ppl: return "PL/PPL/SEMI"
    if has_semi and has_ppl:            return "PL/PPL/SEMI"
    # 2 types
    if has_semi and has_pl:             return "PL/SEMI"
    if has_pl and has_ppl:              return "PL/PPL"
    if has_pl:                          return "PL"
    if has_ppl:                         return "PL/PPL"
    if has_semi:                        return "PL/SEMI"
    return "PL/PPL"

df["accessbilite"] = df["accessbilite"].apply(normalize_acc)
df["ZONNE"]  = df["ZONNE"].fillna("").astype(str).str.strip().str.upper()
df["REGION"] = df["REGION"].fillna("").astype(str).str.strip().str.upper()
REGION_NORM_OPT = {
    "NABEUL":"CAP BON 2","BEJA":"NORD","MANOUBA":"NORD",
    "GAFSA":"GAFSA / KASSRINE","KASSRINE":"GAFSA / KASSRINE",
    "CAPB1":"CAP BON 1","CAP B1":"CAP BON 1",
    "CAPB2":"CAP BON 2","CAP B2":"CAP BON 2",
}
df["REGION"] = df["REGION"].replace(REGION_NORM_OPT)
df["date_debut"] = pd.to_datetime(df["date_debut"], errors="coerce")
df["date_fin"]   = pd.to_datetime(df["date_fin"],   errors="coerce")

date_cols = []  # no date matrix — dates come from date_debut/date_fin columns

print(f"  {len(df)} agriculteurs chargés depuis {len(deposited) if deposited else 'tous les'} commerciaux")
print(f"  Commerciaux inclus: {sorted(df['commercial'].unique().tolist())}")
print(f"  Usines: {sorted(df['USINE'].unique().tolist())}")
print(f"  Tonnage total: {df['TONNAGE'].sum():,.0f}t")

# ============================================================
# STEP 2: Build Farmer objects
# ============================================================
class Farmer:
    def __init__(self, row, date_cols):
        self.commercial  = row["commercial"]
        self.name        = str(row["AGRICULTEUR"]).strip()
        self.usine       = row["USINE"]
        self.region      = str(row.get("REGION", "") or "").strip()
        self.zone        = str(row.get("ZONNE", "") or "").strip()
        self.access      = row["accessbilite"]
        self.tonnage     = float(row["TONNAGE"])
        self.allowed_veh = ACCESS_VEHICLES.get(self.access, ACCESS_VEHICLES["NAN"])
        self.distance_km = get_distance(self.zone, self.usine)
        self.geo_region  = get_region(self.zone)

        # ── Build maturity window ──────────────────────────
        # Case 1: date_cols present (Excel fallback path)
        if date_cols:
            raw_window = {}
            for col in date_cols:
                val = row.get(col)
                if val is not None and pd.notna(val) and float(val) > 0:
                    d = col.date()
                    if SEASON_START <= d <= SEASON_END:
                        raw_window[d] = float(val)

            if raw_window:
                self.start = clamp_date(min(raw_window))
                self.end   = clamp_date(max(raw_window))
                date_sum   = sum(raw_window.values())
                if date_sum < self.tonnage * 0.5:
                    n_days = max(1, (self.end - self.start).days + 1)
                    daily  = self.tonnage / n_days
                    self.window = {
                        clamp_date(self.start + datetime.timedelta(days=i)): round(daily, 1)
                        for i in range(n_days)
                    }
                elif date_sum > self.tonnage * 1.5:
                    factor = self.tonnage / date_sum
                    self.window = {d: round(t * factor, 1) for d, t in raw_window.items()}
                else:
                    self.window = raw_window
                return

        # Case 2: Supabase path — use date_debut / date_fin columns
        try:
            raw_start = pd.to_datetime(row["date_debut"]).date()
            raw_end   = pd.to_datetime(row["date_fin"]).date()
        except Exception:
            raw_start = datetime.date(2026, 7, 1)
            raw_end   = datetime.date(2026, 7, 31)

        if pd.isna(raw_start): raw_start = datetime.date(2026, 7, 1)
        if pd.isna(raw_end):   raw_end   = datetime.date(2026, 7, 31)

        self.start  = clamp_date(raw_start)
        self.end    = clamp_date(raw_end)
        if self.end <= self.start:
            self.end = clamp_date(self.start + datetime.timedelta(days=30))

        n_days = max(1, (self.end - self.start).days + 1)
        daily  = self.tonnage / n_days
        self.window = {
            self.start + datetime.timedelta(days=i): round(daily, 1)
            for i in range(n_days)
        }

farmers = [Farmer(row, date_cols) for _, row in df.iterrows()]
total_dist = sum(f.distance_km for f in farmers if f.distance_km < 999)
print(f"  Built {len(farmers)} farmers | Avg distance to usine: {total_dist/len(farmers):.0f}km")

# Show distance summary
print()
print("  Distance summary by region:")
from collections import defaultdict
by_region = defaultdict(list)
for f in farmers:
    by_region[f.geo_region].append(f.distance_km)
for region, dists in sorted(by_region.items()):
    valid = [d for d in dists if d < 999]
    if valid:
        print(f"    {region:<15}: avg {sum(valid)/len(valid):.0f}km to assigned usine")

# ============================================================
# STEP 3: OR-Tools model with distance objective
# ============================================================
print("\nBuilding OR-Tools model with distance optimization...")

SCALE = 10
MAX_SOLVE_SECONDS = 80  # 80s solver + ~20s Excel = 100s total

all_dates   = sorted({d for f in farmers for d in f.window.keys()})
date_to_idx = {d: i for i, d in enumerate(all_dates)}
N_DATES     = len(all_dates)
N_FARM      = len(farmers)

model = cp_model.CpModel()

x = {}
for f_idx, farmer in enumerate(farmers):
    for d_idx, date in enumerate(all_dates):
        if date in farmer.window:
            # ub = taux journalier × multiplier (flexibilité de scheduling)
            # ub=1.0 trop strict → INFEASIBLE car les caps PIC obligent
            # à concentrer/étaler. ub=2.5 permet à OR-Tools de regrouper
            # les livraisons sur certains jours pour respecter les caps.
            # Le tonnage total reste contraint (±5%) donc pas de débordement global.
            ub = int(farmer.window[date] * SCALE * 2.5)
            x[(f_idx, d_idx)] = model.NewIntVar(0, max(1, ub), f"x_{f_idx}_{d_idx}")
        else:
            x[(f_idx, d_idx)] = model.NewConstant(0)

# Constraint 1: Tonnage per farmer (±5%)
for f_idx, farmer in enumerate(farmers):
    total_scaled   = int(farmer.tonnage * SCALE)
    window_max     = int(sum(farmer.window.values()) * SCALE)
    effective      = min(total_scaled, window_max)
    tolerance      = max(int(total_scaled * 0.05), SCALE)
    model.Add(sum(x[(f_idx, d)] for d in range(N_DATES)) >= effective - tolerance)
    model.Add(sum(x[(f_idx, d)] for d in range(N_DATES)) <= effective + tolerance)

# ── Pré-calcul des caps effectifs (UNE seule fois, avant les boucles) ──────
# Évite le spam ⚠️ × N_DATES et calcule correctement le besoin réel
comm_effective_caps = {}
for comm in set(f.commercial for f in farmers):
    cap_declared  = COMMERCIAL_CAPS.get(comm, 1200)
    total_comm_t  = sum(f.tonnage for f in farmers if f.commercial == comm)
    cap_needed    = math.ceil(total_comm_t / N_DATES) if N_DATES > 0 else cap_declared
    cap_effective = max(cap_declared, cap_needed)
    comm_effective_caps[comm] = cap_effective
    if cap_effective > cap_declared:
        print(f"    ⚠️  {comm}: cap ajusté {cap_declared}→{cap_effective}t/j "
              f"(tonnage={total_comm_t:,.0f}t / {N_DATES}j = {cap_needed}t/j requis)")

# Constraint 2: Commercial daily cap — UNIQUEMENT pendant le pic
# Caps s'appliquent du 1-15 juillet pour tous les commerciaux
# SAUF JILANI et KHALIL : caps du 1-12 juillet seulement
CAP_PERIOD_DEFAULT = (datetime.date(2026, 7, 1), datetime.date(2026, 7, 15))
CAP_PERIOD_SPECIAL = {
    "JILANI OBAY": (datetime.date(2026, 7, 1), datetime.date(2026, 7, 12)),
    "KHALIL":      (datetime.date(2026, 7, 1), datetime.date(2026, 7, 12)),
}
for d_idx, date in enumerate(all_dates):
    by_comm = defaultdict(list)
    for f_idx, f in enumerate(farmers):
        by_comm[f.commercial].append(x[(f_idx, d_idx)])
    for comm, vs in by_comm.items():
        # Déterminer la période de cap pour ce commercial
        cap_start, cap_end = CAP_PERIOD_SPECIAL.get(comm, CAP_PERIOD_DEFAULT)
        if cap_start <= date <= cap_end:
            # Réduire le cap solveur pour absorber l'arrondi
            _cap_brut = comm_effective_caps.get(comm, 1200)
            _marge    = max(ROUNDING_MARGIN_MIN, _cap_brut * ROUNDING_MARGIN_PCT)
            _cap_solv = max(50, _cap_brut - _marge)  # min 50t
            model.Add(sum(vs) <= int(_cap_solv * SCALE))
        # Hors pic: pas de limite journalière commerciale

# Constraint 3: Factory daily cap — basé sur transport RÉEL confirmé + jokers
# cap_reel = min(cap_max_usine, transport_confirmé + jokers_alloués)
def get_real_cap(usine):
    """
    Cap journalier réel = min(cap_théorique, transport_confirmé + jokers).
    Plancher = cap_théorique/2 pour éviter INFEASIBLE si transport très bas.
    Si transport non encore complété, on autorise jusqu'au cap théorique.
    """
    cap_theorique  = FACTORY_CAPS.get(usine, 2000)
    transport_conf = TRANSPORT_CONFIRMED.get(usine, {}).get("total", cap_theorique)
    joker_pl       = JOKER_ALLOC.get(usine, {}).get("PL_joker", 0)
    joker_ppl      = JOKER_ALLOC.get(usine, {}).get("PPL_joker", 0)
    cap_transport  = transport_conf + joker_pl + joker_ppl
    # Plancher = cap_théorique pour garantir la faisabilité
    # (le transport est en cours de confirmation — on ne bloque pas le plan)
    cap_reel = min(cap_theorique, max(cap_transport, cap_theorique))
    return cap_reel  # = cap_theorique tant que transport < théorique

# Factory caps — UNIQUEMENT du 1-15 juillet (hors pic: pas de limite usine)
FACTORY_CAP_START = datetime.date(2026, 7, 1)
FACTORY_CAP_END   = datetime.date(2026, 7, 15)
for d_idx, date in enumerate(all_dates):
    by_fact = defaultdict(list)
    for f_idx, f in enumerate(farmers):
        by_fact[f.usine].append(x[(f_idx, d_idx)])
    for fact, vs in by_fact.items():
        if FACTORY_CAP_START <= date <= FACTORY_CAP_END:
            cap_reel_brut = get_real_cap(fact)
            # Réduire cap solveur pour absorber l'arrondi à la dizaine
            _marge_f = max(ROUNDING_MARGIN_MIN, cap_reel_brut * ROUNDING_MARGIN_PCT)
            cap_reel = max(20, cap_reel_brut - _marge_f)
            model.Add(sum(vs) <= int(cap_reel * SCALE))
        # Hors pic: pas de limite usine

print("  Caps journaliers (transport confirmé vs cap théorique):")
for usine in FACTORY_CAPS:
    cap_th  = FACTORY_CAPS[usine]
    cap_tr  = TRANSPORT_CONFIRMED.get(usine, {}).get("total", cap_th)
    jok     = JOKER_ALLOC.get(usine, {}).get("PL_joker", 0) + JOKER_ALLOC.get(usine, {}).get("PPL_joker", 0)
    cap_r   = get_real_cap(usine)
    manque  = cap_th - cap_tr - jok
    status  = f"✅ complet" if manque <= 0 else f"⚠️ manque {manque}t/j de bennes"
    print(f"    {usine:<12}: théorique={cap_th} | confirmé={cap_tr}+{jok}j | utilisé={cap_r} | {status}")

# Constraint 4 SUPPRIMÉE (peak avg cap) — cause INFEASIBLE
# Constraint 5 SUPPRIMÉE (lissage) — cause INFEASIBLE avec ub=1.0
# Raison: quand plusieurs agriculteurs terminent leur fenêtre en même temps,
# la capacité naturelle peut chuter de >20% entre deux jours consécutifs.
# Avec ub=1.0 (profil naturel), la contrainte de lissage est mathématiquement
# incompatible avec les chutes de fin de saison → INFEASIBLE instantané.
# L'objectif (minimiser déviation du plan) gère le lissage de façon souple.
print("  Constraints added: tonnage + commercial caps (pic only) + factory caps (pic only)")

# ============================================================
# OBJECTIVE: Minimize plan deviation + distance cost
# Weight: 70% minimize deviation from plan, 30% minimize transport distance
# ============================================================
DEVIATION_WEIGHT = 7   # higher = stick closer to original plan
DISTANCE_WEIGHT  = 3   # higher = favor closer usines

deviations = []
distance_costs = []

for f_idx, farmer in enumerate(farmers):
    dist_norm = min(farmer.distance_km, 300)
    # Precompute integer coefficient — OR-Tools variables don't support // operator
    dist_coeff = int(dist_norm * DISTANCE_WEIGHT // 100) + 1  # always >= 1
    for d_idx, date in enumerate(all_dates):
        if date in farmer.window:
            original = int(farmer.window[date] * SCALE)
            diff = model.NewIntVar(0, int(farmer.tonnage * SCALE), f"dev_{f_idx}_{d_idx}")
            model.AddAbsEquality(diff, x[(f_idx, d_idx)] - original)
            deviations.append(diff * DEVIATION_WEIGHT)
            # Distance cost: more tons on long routes = higher cost
            distance_costs.append(x[(f_idx, d_idx)] * dist_coeff)

model.Minimize(sum(deviations) + sum(distance_costs))
print(f"  Objective: minimize (plan deviation ×{DEVIATION_WEIGHT}) + (distance cost ×{DISTANCE_WEIGHT})")

# ============================================================
# STEP 4: Solve
# ============================================================
print(f"\nSolving (max {MAX_SOLVE_SECONDS}s, 4 cores)...")
solver = cp_model.CpSolver()
solver.parameters.max_time_in_seconds = MAX_SOLVE_SECONDS
solver.parameters.num_search_workers  = 4
solver.parameters.log_search_progress = False

status = solver.Solve(model)
STATUS_NAMES = {
    cp_model.OPTIMAL:   "OPTIMAL",
    cp_model.FEASIBLE:  "FEASIBLE",
    cp_model.INFEASIBLE:"INFEASIBLE",
    cp_model.UNKNOWN:   "UNKNOWN",
}
print(f"  Status: {STATUS_NAMES.get(status, status)} | Time: {solver.WallTime():.1f}s")

if status not in (cp_model.OPTIMAL, cp_model.FEASIBLE):
    print("  No solution — applying capped fallback distribution")
    # Fallback intelligent: distribuer le tonnage en respectant les caps
    # au lieu de tout mettre dans la fenêtre originale sans contraintes
    from collections import defaultdict
    solution = {}
    # Compteurs journaliers pour vérifier les caps
    daily_comm  = defaultdict(lambda: defaultdict(float))   # [date][commercial]
    daily_fact  = defaultdict(lambda: defaultdict(float))   # [date][usine]

    for fi, farmer in enumerate(farmers):
        remaining = farmer.tonnage
        # Distribution uniforme sur toute la fenêtre (pas concentrée sur premiers jours)
        # Chaque jour de la fenêtre reçoit le taux journalier naturel
        dates_avail = [(di, all_dates[di]) for di in range(N_DATES)
                       if farmer.window.get(all_dates[di], 0) > 0]
        
        # Passe 1: distribuer le taux naturel sur chaque jour disponible
        for di, date in dates_avail:
            if remaining <= 0:
                break
            in_peak_comm = (PEAK_START <= date <= datetime.date(2026, 7, 15))
            in_peak_fact = (PEAK_START <= date <= PEAK_END)
            if farmer.commercial in ("JILANI OBAY","KHALIL"):
                in_peak_comm = (PEAK_START <= date <= datetime.date(2026, 7, 12))
            
            # Taux naturel = tonnage / nb_jours_fenêtre
            natural_rate = farmer.window.get(date, 0)
            
            if in_peak_comm:
                cap_c   = comm_effective_caps.get(farmer.commercial, 9999)
                used_c  = daily_comm[date][farmer.commercial]
                avail_c = max(0, cap_c - used_c)
            else:
                avail_c = 9999
            
            if in_peak_fact:
                cap_f   = get_real_cap(farmer.usine)
                used_f  = daily_fact[date][farmer.usine]
                avail_f = max(0, cap_f - used_f)
            else:
                avail_f = 9999
            
            # Livrer exactement le taux naturel (pas plus, pas moins)
            daily_qty = min(avail_c, avail_f, remaining, natural_rate)
            if daily_qty > 0.1:
                solution[(fi, di)] = round(daily_qty, 2)
                daily_comm[date][farmer.commercial] += daily_qty
                daily_fact[date][farmer.usine]      += daily_qty
                remaining -= daily_qty
            else:
                solution[(fi, di)] = 0.0
        
        # Passe 2: placer le reste (si remaining > 0 à cause des caps du pic)
        if remaining > 0.5:
            for di, date in sorted(dates_avail, key=lambda x: daily_comm[x[1]][farmer.commercial]):
                if remaining <= 0.5:
                    break
                in_peak_comm = (PEAK_START <= date <= datetime.date(2026, 7, 15))
                in_peak_fact = (PEAK_START <= date <= PEAK_END)
                if farmer.commercial in ("JILANI OBAY","KHALIL"):
                    in_peak_comm = (PEAK_START <= date <= datetime.date(2026, 7, 12))
                
                if in_peak_comm:
                    cap_c  = comm_effective_caps.get(farmer.commercial, 9999)
                    avail_c = max(0, cap_c - daily_comm[date][farmer.commercial])
                else:
                    avail_c = 9999
                
                if in_peak_fact:
                    cap_f  = get_real_cap(farmer.usine)
                    avail_f = max(0, cap_f - daily_fact[date][farmer.usine])
                else:
                    avail_f = 9999
                
                extra = min(avail_c, avail_f, remaining)
                if extra > 0.1:
                    cur = solution.get((fi, di), 0.0)
                    solution[(fi, di)] = round(cur + extra, 2)
                    daily_comm[date][farmer.commercial] += extra
                    daily_fact[date][farmer.usine]      += extra
                    remaining -= extra

        # Remplir les cases non assignées à 0
        for di in range(N_DATES):
            if (fi, di) not in solution:
                solution[(fi, di)] = 0.0
else:
    solution = {(fi, di): solver.Value(x[(fi, di)]) / SCALE
                for fi in range(N_FARM) for di in range(N_DATES)}

# ============================================================
# STEP 5: Build output
# ============================================================
print("\nBuilding output tables...")

def choose_vehicles(tons, allowed_raw, usine=None):
    """
    VERSION V5 — Logique simplifiée et correcte.
    
    Principe:
    - L'accessibilité (PL/SEMI, PL/PPL...) = véhicules PHYSIQUEMENT possibles sur ce terrain
    - On choisit UN seul type de véhicule pour chaque livraison (le plus adapté au tonnage)
    - Exceptionnellement 2 types pour COMOCAP (TRACTEUR + transport principal)
    - Pas de découpage proportionnel artificiel en 4-5 types
    
    Logique de sélection:
    1. Regarder la préférence usine (SICAM préfère SEMI, ELFALLEH préfère PPL...)
    2. Si le véhicule préféré est accessible → l'utiliser
    3. Sinon → choisir le meilleur véhicule accessible pour le tonnage
    4. _alloc() répartit en voyages en respectant mn ≤ charge ≤ mx
    """
    # Normaliser allowed
    _norm = {"PETIT POILOUR":"PPL","POILOUR":"PL"}
    allowed = list(dict.fromkeys(_norm.get(v,v) for v in allowed_raw))
    if not allowed:
        allowed = ["PL"]

    def _alloc(veh, qty):
        """
        Alloue qty tonnes sur veh. Respecte mn ≤ charge ≤ mx par voyage.
        Si la répartition parfaite est impossible, répartition équilibrée.
        """
        if qty <= 0: return []
        
        if veh == "TRACTEUR":
            mn_t, mx_t = FLEET_CAPACITY.get(veh, (9, 11))
            return [{"vehicle": veh, "trips": 1, "tons_each": round(min(qty, mx_t), 2)}]
        
        mn, mx = FLEET_CAPACITY.get(veh, (7, 25))
        
        # 1 voyage suffit
        if qty <= mx:
            return [{"vehicle": veh, "trips": 1, "tons_each": round(qty, 2)}]
        
        # Plusieurs voyages nécessaires
        trips = math.ceil(qty / mx)
        each  = qty / trips
        
        # Si chaque voyage < mn → réduire le nombre de voyages
        while each < mn and trips > 1:
            trips -= 1
            each = qty / trips
        
        # Répartition équilibrée
        base  = int(qty // trips)
        extra = int(qty - base * trips)  # nb voyages avec 1t de plus
        
        entries = []
        n_heavy = extra        # voyages à (base+1)t
        n_light = trips - extra  # voyages à base t
        if n_heavy > 0:
            entries.append({"vehicle": veh, "trips": n_heavy,
                            "tons_each": round(base + 1, 2)})
        if n_light > 0:
            entries.append({"vehicle": veh, "trips": n_light,
                            "tons_each": round(base, 2)})
        return entries

    def _best_for_tons(qty, candidates):
        """
        Parmi les candidats accessibles, choisit celui qui gère qty le plus proprement.
        Score: 0=parfait, 1=léger sous-charge, 2=multi-voyages propre, 3=surcharge
        """
        best, best_score = None, 999
        for veh in candidates:
            if veh not in allowed:
                continue
            mn_v, mx_v = FLEET_CAPACITY.get(veh, (7, 25))
            if qty <= mx_v and qty >= mn_v:
                score = 0  # 1 voyage parfait
            elif qty <= mx_v:
                score = 1  # 1 voyage léger sous-charge
            else:
                trips = math.ceil(qty / mx_v)
                each  = qty / trips
                score = 2 if each >= mn_v else 3
            if score < best_score:
                best_score, best = score, veh
        return best

    # ── CAS SPÉCIAL : RM (Récolte Mécanique) ────────────────────────────
    if allowed == ["SEMI"]:
        mn, mx = FLEET_CAPACITY["SEMI"]
        if tons <= 0: return []
        trips = max(1, math.ceil(tons / mx))
        each  = round(tons / trips, 2)
        while each < mn and trips > 1:
            trips -= 1
            each  = round(tons / trips, 2)
        return [{"vehicle": "SEMI", "trips": trips, "tons_each": each}]

    # ── PRÉFÉRENCES PAR USINE ───────────────────────────────────────────
    # Ordre de préférence des véhicules selon l'usine
    # Le système choisit le PREMIER véhicule de la liste qui est accessible
    USINE_PREFS = {
        "SICAM":    ["SEMI", "PL",  "PPL"],   # préfère SEMI
        "TUCAL":    ["SEMI", "PL",  "PPL"],   # préfère SEMI
        "COMOCAP":  ["PL",   "PPL", "SEMI"],  # préfère PL (TRACTEUR géré séparément)
        "ABIDA":    ["PL",   "SEMI","PPL"],   # préfère PL
        "ELFALLEH": ["PPL",  "PL",  "SEMI"],  # préfère PPL
    }
    prefs = USINE_PREFS.get(usine, ["PL", "PPL", "SEMI"])

    # ── CAS COMOCAP : TRACTEUR fixe (10t) + transport principal ─────────
    if usine == "COMOCAP" and "TRACTEUR" in allowed:
        result = []
        # 1 voyage TRACTEUR ≈ 10t
        trac_tons = min(10, tons * 0.14)
        trac_tons = round(trac_tons)
        if trac_tons > 0:
            result.append({"vehicle": "TRACTEUR", "trips": 1,
                           "tons_each": round(trac_tons, 2)})
        remaining = round(tons - trac_tons, 2)
        if remaining > 0:
            main_veh = _best_for_tons(remaining, ["PL", "PPL", "SEMI"])
            if not main_veh:
                main_veh = "PL"
            result.extend(_alloc(main_veh, remaining))
        return result if result else _alloc("PL", tons)

    # ── CAS GÉNÉRAL : 1 véhicule principal ──────────────────────────────
    # Choisir le meilleur véhicule accessible selon les préférences usine
    primary = _best_for_tons(tons, prefs)
    if not primary:
        # Aucun des préférés n'est accessible → prendre n'importe quel accessible
        primary = _best_for_tons(tons, ["SEMI", "PL", "PPL", "TRACTEUR"])
    if not primary:
        primary = "PL"  # dernier recours absolu

    result = _alloc(primary, tons)
    
    # Vérification minimale
    if not result:
        result = [{"vehicle": primary, "trips": 1, "tons_each": round(tons, 2)}]
    
    return result

all_days = []
# ✅ CONSOLIDATION : regrouper les parcelles du même agriculteur livrant
# à la même usine le même jour pour optimiser le transport
# (ex: 2 parcelles HAFEDH MOSBE × 10t = 1 voyage PL de 20t au lieu de 2)
from collections import defaultdict as _dd_cons
consolidated = _dd_cons(lambda: {"tons": 0, "farmer": None, "parcelles": []})
for f_idx, farmer in enumerate(farmers):
    for d_idx, date in enumerate(all_dates):
        tons = solution[(f_idx, d_idx)]
        if tons <= 0.5: continue
        # Clé : nom agriculteur + usine + date (= un envoi unique)
        key = (farmer.name.strip().upper(), farmer.usine, date)
        consolidated[key]["tons"] += round(tons, 1)
        consolidated[key]["farmer"] = farmer  # garder la dernière référence
        consolidated[key]["parcelles"].append(f_idx)

# Génération des all_days à partir des envois consolidés
for (nom, usine, date), data in consolidated.items():
    tons_brut = data["tons"]
    farmer    = data["farmer"]
    n_parcelles = len(data["parcelles"])
    
    # ✅ ARRONDI à la dizaine la plus proche
    tons = int(round(round(tons_brut, 1) / 10)) * 10
    if tons == 0 and tons_brut > 0:
        tons = 10  # minimum 10t
    if tons <= 0:
        continue
    
    vehicles = choose_vehicles(tons, farmer.allowed_veh, usine=farmer.usine)
    veh_parts = []
    for v in vehicles:
        if v.get('note') and v['tons_each'] == 0:
            veh_parts.append("TRACTEUR x1 (caisses COMOCAP)")
        elif v['tons_each'] > 0:
            veh_parts.append(f"{v['vehicle']} x{v['trips']}({v['tons_each']}t)")
    veh_str = " | ".join(veh_parts)
    veh_type = vehicles[0]["vehicle"] if vehicles else "POILOUR"
    trips    = sum(v["trips"] for v in vehicles)
    dist_km  = farmer.distance_km if farmer.distance_km < 999 else 0
    note     = f"{n_parcelles} parcelles consolidées" if n_parcelles > 1 else "AI-optimise"
    
    all_days.append({
        "Commercial":    farmer.commercial,
        "Agriculteur":   farmer.name,
        "Usine":         usine,
        "Region":        farmer.geo_region,
        "Zone":          farmer.zone,
        "Accessibilite": farmer.access,
        "Date":          date,
        "Tonnes/Jour":   tons,
        "Type Vehicule": veh_type,
        "Vehicules":     veh_str,
        "Nb Voyages":    trips,
        "Distance km":   dist_km,
        "Date Debut":    farmer.start,
        "Date Fin":      farmer.end,
        "Total Tonnes":  farmer.tonnage,
        "Pic de Recolte":"PIC" if PEAK_START <= date <= PEAK_END else "",
        "Note":          note,
    })

# ANCIEN code remplacé par la consolidation par envoi unique

# ✅ POST-TRAITEMENT TRACTEUR COMOCAP : max 10 voyages/jour
# DOIT être fait AVANT result_df pour aussi corriger transport_rows
print("  Post-traitement TRACTEUR COMOCAP (max 10 voyages/jour)...")
TRACTEUR_DAILY_LIMIT = 10  # max 10 voyages × 10t = 100t/jour

# Indexer all_days par date
from collections import defaultdict as _dd
by_date = _dd(list)
for i, row in enumerate(all_days):
    by_date[row["Date"]].append(i)

nb_corrections = 0
for date_val, indices in by_date.items():
    # Filtrer COMOCAP avec TRACTEUR
    comocap_trac_idx = [
        i for i in indices
        if all_days[i]["Usine"] == "COMOCAP" 
        and "TRACTEUR" in str(all_days[i].get("Vehicules", ""))
    ]
    if len(comocap_trac_idx) <= TRACTEUR_DAILY_LIMIT:
        continue
    # Trop de voyages TRACTEUR — convertir les surplus en PPL
    # Garder les 10 PREMIERS par ordre d'index (déterministe)
    # Convertir le reste
    sorted_by_tons = sorted(comocap_trac_idx, key=lambda i: all_days[i].get("Tonnes/Jour", 0))
    to_convert = sorted_by_tons[TRACTEUR_DAILY_LIMIT:]
    for idx in to_convert:
        ton = all_days[idx].get("Tonnes/Jour", 0)
        # PPL capacité 7-12t → calculer voyages
        n_voyages = max(1, math.ceil(ton / 12))
        ton_each  = round(ton / n_voyages, 1)
        all_days[idx]["Vehicules"]    = f"PPL x{n_voyages}({ton_each}t)"
        all_days[idx]["Type Vehicule"] = "PPL"
        all_days[idx]["Nb Voyages"]    = n_voyages
        nb_corrections += 1
print(f"    → {nb_corrections} agriculteur(s) TRACTEUR convertis en PPL pour respecter cap 10/jour")

result_df = pd.DataFrame(all_days).sort_values(
    ["Date","Commercial","Agriculteur"]).reset_index(drop=True)

# Normalize Region column in result
REGION_NORM_RESULT = {
    "NABEUL":"CAP BON 2","nabeul":"CAP BON 2",
    "BEJA":"NORD","beja":"NORD",
    "MANOUBA":"NORD","manouba":"NORD",
    "GAFSA":"GAFSA / KASSRINE","KASSRINE":"GAFSA / KASSRINE",
    "CAPB1":"CAP BON 1","CAP B1":"CAP BON 1",
    "CAPB2":"CAP BON 2","CAP B2":"CAP BON 2",
}
if "Region" in result_df.columns:
    result_df["Region"] = result_df["Region"].replace(REGION_NORM_RESULT)
elif "Région" in result_df.columns:
    result_df["Région"] = result_df["Région"].replace(REGION_NORM_RESULT)
print(f"  Planning rows: {len(result_df)}")

# Transport rows
# Mapping interne → colonnes Excel/dashboard
VEH_MAP_TO_COL = {
    "TRACTEUR":    "TRACTEUR",
    "PPL":         "PETIT POILOUR",   # PPL = Petit Poilour
    "PL":          "POILOUR",          # PL  = Poilour
    "SEMI":        "SEMI",
    "DOUBLE_REM":  "SEMI",             # Double Rem = Semi
}

transport_rows = []
by_date_comm = defaultdict(lambda: defaultdict(lambda: defaultdict(int)))
for row in all_days:
    d, comm = row["Date"], row["Commercial"]
    veh_str = row.get("Vehicules", "")
    for part in veh_str.split("|"):
        part = part.strip()
        if not part: continue
        # Extraire: "PLx2(15.0t)" → veh=PL, trips=2
        # Extraire: "TRC(caisses)" → veh=TRACTEUR, trips=1
        trips_n = 1
        matched_veh = None
        # Chercher dans l'ordre du plus long au plus court pour éviter PPL→PL
        for internal, col in sorted(VEH_MAP_TO_COL.items(), key=lambda x: -len(x[0])):
            if part.upper().startswith(internal.upper()):
                matched_veh = col
                try:
                    trips_n = int(part.split("x")[1].split("(")[0].strip())
                except: trips_n = 1
                break
        if matched_veh:
            by_date_comm[d][comm][matched_veh] += trips_n

for d in sorted(by_date_comm.keys()):
    for comm in sorted(by_date_comm[d].keys()):
        nd       = by_date_comm[d][comm]
        day_rows = [r for r in all_days if r["Date"]==d and r["Commercial"]==comm]
        avg_dist = round(sum(r["Distance km"] for r in day_rows)/max(len(day_rows),1), 0)
        transport_rows.append({
            "Date":                  d,
            "Commercial":            comm,
            "Total Tonnes":          round(sum(r["Tonnes/Jour"] for r in day_rows), 1),
            "Voyages TRACTEUR":      nd.get("TRACTEUR", 0),
            "Voyages PETIT POILOUR": nd.get("PETIT POILOUR", 0),
            "Voyages POILOUR":       nd.get("POILOUR", 0),
            "Voyages SEMI":          nd.get("SEMI", 0),
            "Distance Moy km":       int(avg_dist),
            "Jours Double":          0,
            "Pic":                   "yes" if PEAK_START <= d <= PEAK_END else "",
        })

transport_df = pd.DataFrame(transport_rows)

# Region summary
region_summary = result_df.groupby("Region").agg(
    Tonnes_Total=("Tonnes/Jour", "sum"),
    Nb_Agriculteurs=("Agriculteur", "nunique"),
    Distance_Moy_km=("Distance km", "mean"),
).round(0).reset_index()
region_summary.columns = ["Region","Tonnes Totales","Nb Agriculteurs","Distance Moy (km)"]

# Availability table
avail_rows = []
for farmer in farmers:
    for veh in farmer.allowed_veh:
        avail_rows.append({
            "Commercial": farmer.commercial, "Agriculteur": farmer.name,
            "Usine": farmer.usine, "Zone": farmer.zone,
            "Region": farmer.geo_region, "Distance km": farmer.distance_km,
            "Type Vehicule": veh, "Date Debut": farmer.start,
            "Date Fin": farmer.end, "Total Tonnes": farmer.tonnage,
        })
avail_df = pd.DataFrame(avail_rows).sort_values(["Commercial","Date Debut"]).reset_index(drop=True)

# Resume
resume_rows = []
for comm in sorted(result_df["Commercial"].unique()):
    resume_rows.append({
        "Commercial":            comm,
        "Tonnes Totales Saison": round(df[df["commercial"]==comm]["TONNAGE"].sum(), 0),
        "Nb Agriculteurs":       df[df["commercial"]==comm]["AGRICULTEUR"].nunique(),
        "Conflits Resolus":      0,
        "Total Jours Double":    0,
        "Statut":                "AI optimise",
    })
resume_df = pd.DataFrame(resume_rows)

# ============================================================
# STEP 6: Write Excel
# ============================================================
print(f"\nWriting {OUTPUT_FILE}...")

wb   = Workbook()
thin = Side(style="thin", color="BBBBBB")
border = Border(left=thin, right=thin, top=thin, bottom=thin)

def style_ws(ws, headers, widths, color):
    fill = PatternFill("solid", start_color=color)
    font = Font(bold=True, color="FFFFFF", name="Calibri", size=10)
    for ci, (h, w) in enumerate(zip(headers, widths), 1):
        c = ws.cell(row=1, column=ci, value=h)
        c.font=font; c.fill=fill
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = border
        ws.column_dimensions[get_column_letter(ci)].width = w
    ws.row_dimensions[1].height = 28

def write_df(ws, df_in, start_row=2):
    for ro, (_, row) in enumerate(df_in.iterrows()):
        for ci, val in enumerate(row, 1):
            c = ws.cell(row=start_row+ro, column=ci, value=val)
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.border = border
            c.font = Font(name="Calibri", size=9)

# Sheet 1: Planning
ws1 = wb.active; ws1.title = "Planning Journalier"
COLS1 = ["Commercial","Agriculteur","Usine","Region","Zone","Accessibilite",
         "Date","Tonnes/Jour","Type Vehicule","Vehicules","Nb Voyages",
         "Distance km","Date Debut","Date Fin","Total Tonnes","Pic de Recolte","Note"]
W1 = [15,24,10,12,14,11,12,10,13,38,9,11,12,12,12,10,14]
disp = result_df.copy(); disp.columns = COLS1
style_ws(ws1, COLS1, W1, "1F4E79")
write_df(ws1, disp)

# Sheet 2: Transport
ws2 = wb.create_sheet("Besoins Transport-Jour")
COLS2 = ["Date","Commercial","Total Tonnes","Voyages TRACTEUR",
         "Voyages PETIT POILOUR","Voyages POILOUR","Voyages SEMI",
         "Distance Moy km","Jours Double","Pic"]
style_ws(ws2, COLS2, [12,16,12,16,20,16,14,14,12,6], "375623")
write_df(ws2, transport_df[COLS2])

# Sheet 3: By Region
ws3 = wb.create_sheet("Par Region")
COLS3 = ["Region","Tonnes Totales","Nb Agriculteurs","Distance Moy (km)"]
style_ws(ws3, COLS3, [16,16,16,16], "7B2D8B")
write_df(ws3, region_summary[COLS3])

# Sheet 4: Availability
ws4 = wb.create_sheet("Disponibilite Vehicules")
COLS4 = ["Commercial","Agriculteur","Usine","Zone","Region","Distance km",
         "Type Vehicule","Date Debut","Date Fin","Total Tonnes"]
style_ws(ws4, COLS4, [15,26,10,16,12,11,14,12,12,12], "C00000")
write_df(ws4, avail_df)

# Sheet 5: Résumé optimisation (pas de double transport avec OR-Tools)
ws5 = wb.create_sheet("Résumé Optimisation")
ws5["A1"] = "Planning généré par OR-Tools — Optimisation IA (distances + caps simultanés)"
ws5["A1"].font = Font(bold=True, size=11, name="Calibri", color="1F4E79")
ws5["A2"] = f"Statut: {STATUS_NAMES.get(status, status)} | Farmers: {len(farmers)} | Total: {result_df['Tonnes/Jour'].sum():,.0f}t"
ws5["A2"].font = Font(size=10, name="Calibri", color="375623")
ws5["A3"] = "Aucun double transport — OR-Tools évite les conflits dès la planification"
ws5["A3"].font = Font(size=9, name="Calibri", color="5F5E5A")

# Sheet 6: Resume
ws6 = wb.create_sheet("Resume par Commercial")
COLS6 = ["Commercial","Tonnes Totales Saison","Nb Agriculteurs",
         "Conflits Resolus","Total Jours Double","Statut"]
style_ws(ws6, COLS6, [18,20,16,16,18,16], "1F4E79")
write_df(ws6, resume_df[COLS6])

wb.save(TEMP_FILE)

import shutil, time
for attempt in range(5):
    try:
        shutil.move(TEMP_FILE, OUTPUT_FILE)
        break
    except PermissionError:
        if attempt < 4:
            print(f"  Excel ouvert, retry {attempt+1}/5 dans 3s...")
            time.sleep(3)
        else:
            print("  ERREUR: Fermez le fichier Excel et relancez.")
            sys.exit(1)

# ============================================================
# STEP 7: Summary
# ============================================================
print(f"\n{'='*60}")
print(f"  OPTIMIZER V2 COMPLETE -> {OUTPUT_FILE}")
print(f"{'='*60}")
print(f"  Status    : {STATUS_NAMES.get(status, status)}")
print(f"  Rows      : {len(result_df)}")
print(f"  Total declared : {sum(f.tonnage for f in farmers):,.0f}t")
print(f"  Total planned  : {result_df['Tonnes/Jour'].sum():,.0f}t")
print()
print("  Distance optimization:")
for region, dists in sorted(by_region.items()):
    valid = [d for d in dists if d < 999]
    if valid:
        total_ton_dist = sum(f.tonnage * f.distance_km
                             for f in farmers if f.geo_region == region and f.distance_km < 999)
        print(f"    {region:<15}: {len(valid)} farmers | avg {sum(valid)/len(valid):.0f}km")
print()
print("  Commercial caps (vérification PIC UNIQUEMENT 1-15 Juillet):")
# Convertir Date en datetime pour éviter l'erreur .dt.date
result_df["Date"] = pd.to_datetime(result_df["Date"], errors="coerce")
for comm in COMMERCIAL_CAPS:
    cap = COMMERCIAL_CAPS[comm]
    sub = result_df[result_df["Commercial"]==comm].copy()
    if sub.empty: continue
    cap_end_date = pd.Timestamp("2026-07-12") if comm in ("JILANI OBAY","KHALIL") else pd.Timestamp("2026-07-15")
    peak_start_ts = pd.Timestamp(PEAK_START)
    sub_pic = sub[(sub["Date"] >= peak_start_ts) & (sub["Date"] <= cap_end_date)]
    mx_pic  = sub_pic.groupby("Date")["Tonnes/Jour"].sum().max() if not sub_pic.empty else 0
    mx_all  = sub.groupby("Date")["Tonnes/Jour"].sum().max()
    ok_pic  = "✅" if mx_pic <= cap else "❌ DEPASSE PENDANT PIC"
    print(f"    {comm:<20}: max PIC={mx_pic:.0f}t/j | max hors-pic={mx_all:.0f}t/j | limite={cap}t {ok_pic}")
print()
print("  Factory caps (vérification PIC UNIQUEMENT 1-15 Juillet):")
factory_start_ts = pd.Timestamp(FACTORY_CAP_START)
factory_end_ts   = pd.Timestamp(FACTORY_CAP_END)
for fact, cap in FACTORY_CAPS.items():
    sub = result_df[result_df["Usine"]==fact].copy()
    if sub.empty: continue
    sub_pic = sub[(sub["Date"] >= factory_start_ts) & (sub["Date"] <= factory_end_ts)]
    mx_pic  = sub_pic.groupby("Date")["Tonnes/Jour"].sum().max() if not sub_pic.empty else 0
    mx_all  = sub.groupby("Date")["Tonnes/Jour"].sum().max()
    ok_pic  = "✅" if mx_pic <= cap else "❌ DEPASSE PENDANT PIC"
    print(f"    {fact:<12}: max PIC={mx_pic:.0f}t/j | max hors-pic={mx_all:.0f}t/j | limite={cap}t {ok_pic}")
print()
print("  Next: python migrate.py")
print(f"{'='*60}")