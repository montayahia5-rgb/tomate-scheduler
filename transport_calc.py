# -*- coding: utf-8 -*-
"""
transport_calc.py — Calcul des besoins de transport a partir du PLAN RECTIFIE
===============================================================================
Ce module reprend FIDELEMENT (sans reduction) la logique de choix de vehicules
et de flotte de optimizer_v2.py (FLEET_CAPACITY, ACCESS_VEHICLES,
TRANSPORT_CONFIRMED, TRANSPORT_JOKERS, TRANSPORT_REGLES, normalize_acc,
choose_vehicles version V5, chargement de la flotte reelle depuis
transport_etat_final.xlsx).

Difference cle avec optimizer_v2.py : ici, on n'execute PAS de solve OR-Tools.
On applique choose_vehicles() directement sur le TONNAGE RECTIFIE (issu de
build_effective_planning dans comparaison_tab.py) pour chaque
(Commercial, Agriculteur, Usine, Region, Accessibilite, Date, Tonnes/Jour),
puis on agrege le resultat par Usine / Commercial / Region / Type Vehicule,
par jour ET par semaine, et on compare au besoin contre la flotte reelle
disponible pour determiner Disponible vs Manquant (a louer).
"""
import math
import re
import pandas as pd
from collections import defaultdict

# ═══ FLOTTE THEORIQUE (min,max tonnes par voyage) — IDENTIQUE optimizer_v2 ═══
FLEET_CAPACITY = {
    "TRACTEUR":   (9,  11),    # min/max tonnes par voyage (moyenne ~10t)
    "PPL":        (6,  14),    # Petit Poilour
    "PL":         (15, 25),    # Poilour
    "SEMI":       (27, 33),    # Semi-remorque
    "DOUBLE_REM": (27, 33),    # Double Remorque
}
FLEET_CAPACITY["PETIT POILOUR"] = FLEET_CAPACITY["PPL"]
FLEET_CAPACITY["POILOUR"]       = FLEET_CAPACITY["PL"]

# ═══ ACCESSIBILITE -> VEHICULES AUTORISES — IDENTIQUE optimizer_v2 ═══
ACCESS_VEHICLES = {
    "PL":          ["PL"],
    "PPL":         ["PPL"],
    "SEMI":        ["SEMI"],
    "RM":          ["SEMI"],
    "PL/PPL":      ["PL", "PPL"],
    "PL-PPL":      ["PL", "PPL"],
    "TRC/PPL":     ["TRACTEUR", "PPL"],
    "PL/SEMI":     ["PL", "SEMI"],
    "PL-SEMI":     ["PL", "SEMI"],
    "TRC/PPL/PL":  ["TRACTEUR", "PPL", "PL"],
    "PL/PPL/TRC":  ["PL", "PPL", "TRACTEUR"],
    "PL/PPL/SEMI": ["PL", "PPL", "SEMI"],
    "SEMI/PL/PPL": ["SEMI", "PL", "PPL"],
    "NAN":         ["TRACTEUR", "PPL", "PL", "SEMI"],
}

# ═══ CAPACITES CONFIRMEES PAR USINE (theorique, en tonnes/jour) ═══
# Source: transport_etat_final.xlsx (10/06/2026) — IDENTIQUE optimizer_v2
TRANSPORT_CONFIRMED = {
    "SICAM":    {"total": 1381, "PL": 927, "PPL": 64,  "SEMI": 390, "nb_bennes": 67},
    "TUCAL":    {"total": 363,  "PL": 303, "PPL": 0,   "SEMI": 60,  "nb_bennes": 19},
    "COMOCAP":  {"total": 328,  "PL": 91,  "PPL": 147, "SEMI": 90,  "nb_bennes": 23},
    "ABIDA":    {"total": 80,   "PL": 20,  "PPL": 0,   "SEMI": 60,  "nb_bennes": 3},
    "ELFALLEH": {"total": 24,   "PL": 0,   "PPL": 24,  "SEMI": 0,   "nb_bennes": 2},
}
TRANSPORT_JOKERS = {
    "BOURAK":   {"total": 114, "PL": 114, "PPL": 0,  "SEMI": 0, "nb_bennes": 6},
    "LUIMEME":  {"total": 101, "PL": 55,  "PPL": 46, "SEMI": 0, "nb_bennes": 7},
}
TRANSPORT_REGLES = {
    "COMOCAP":  [("TRACTEUR", 100), ("PPL", 0.50), ("PL", 0.30), ("SEMI", 0.20)],
    "TUCAL":    [("PL", 0.30),      ("SEMI", 0.70)],
    "ABIDA":    [("PL", 0.50),      ("SEMI", 0.50)],
    "ELFALLEH": [("PPL", 0.70),     ("PL", 0.30)],
    "SICAM":    [("SEMI", 1.00)],
}
FACTORY_CAPS = {"SICAM": 1500, "TUCAL": 800, "COMOCAP": 800, "ABIDA": 200, "ELFALLEH": 150}


def calc_transport_needs():
    """Tonnage manquant par usine/type au-dela du transport confirme — IDENTIQUE optimizer_v2."""
    needs = {}
    for usine, cap in FACTORY_CAPS.items():
        conf = TRANSPORT_CONFIRMED.get(usine, {}).get("total", 0)
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


def alloc_jokers():
    """Alloue les jokers BOURAK/LUIMEME aux usines en manque — IDENTIQUE optimizer_v2."""
    jpl = TRANSPORT_JOKERS["BOURAK"]["PL"]
    jppl = TRANSPORT_JOKERS["LUIMEME"]["PPL"]
    jpl_luimeme = TRANSPORT_JOKERS["LUIMEME"]["PL"]
    alloc = {}
    for usine in FACTORY_CAPS:
        needs = TRANSPORT_NEEDS.get(usine, {})
        alloc[usine] = {}
        if "PL" in needs and jpl > 0:
            a = min(needs["PL"], jpl)
            alloc[usine]["PL_joker"] = a
            jpl -= a
        if "PL" in needs and jpl_luimeme > 0:
            a = min(needs.get("PL", 0) - alloc[usine].get("PL_joker", 0), jpl_luimeme)
            if a > 0:
                alloc[usine]["PL_joker"] = alloc[usine].get("PL_joker", 0) + a
                jpl_luimeme -= a
        if "PPL" in needs and jppl > 0:
            a = min(needs["PPL"], jppl)
            alloc[usine]["PPL_joker"] = a
            jppl -= a
    return alloc


JOKER_ALLOC = alloc_jokers()


def normalize_acc(x):
    """Normalise un code d'accessibilite — IDENTIQUE optimizer_v2 (respecte exactement
    ce que le commercial a ecrit, ne convertit jamais un type simple en un autre)."""
    x = str(x).strip().upper().replace("-", "/")
    if x in ("NAN", "NONE", "", "NAT"):
        return "PL/PPL"
    if x == "RM":
        return "RM"
    if x in ACCESS_VEHICLES:
        return x
    if x == "PL":
        return "PL"
    if x == "PPL":
        return "PPL"
    parts = set(re.split(r"[/,;\s]+", x))
    parts = {p.strip() for p in parts if p.strip()}
    has_trc = "TRC" in parts or "TRACTEUR" in parts
    has_pl = "PL" in parts
    has_ppl = "PPL" in parts
    has_semi = "SEMI" in parts
    if has_trc and has_pl and has_ppl: return "TRC/PPL/PL"
    if has_trc and has_pl:             return "TRC/PPL/PL"
    if has_trc and has_ppl:            return "TRC/PPL"
    if has_semi and has_pl and has_ppl:return "PL/PPL/SEMI"
    if has_semi and has_ppl:           return "PL/PPL/SEMI"
    if has_semi and has_pl:            return "PL/SEMI"
    if has_semi:                       return "PL/SEMI"
    if has_pl and has_ppl:             return "PL/PPL"
    if has_pl:                         return "PL"
    if has_ppl:                        return "PPL"
    return "PL/PPL"


# ═══ CHARGEMENT FLOTTE REELLE depuis transport_etat_final.xlsx ═══
# Adapte de _load_real_fleet() de optimizer_v2.py : au lieu de lire un chemin
# disque fixe, accepte un fichier uploade (Streamlit UploadedFile) ou un chemin.
def parse_real_fleet_file(file_obj):
    """
    file_obj : fichier uploade Streamlit (ou chemin str) du fichier
    transport_etat_final.xlsx (ou transport_disponible.xlsx).
    Retourne {usine: {type: [capacites triees desc]}} — meme format que
    REAL_FLEET dans optimizer_v2.py.
    """
    try:
        xl = pd.ExcelFile(file_obj)
        sheet = "liste confirmé" if "liste confirmé" in xl.sheet_names else xl.sheet_names[0]
        df = pd.read_excel(file_obj, sheet_name=sheet)
        cols_upper = [str(c).strip().lower() for c in df.columns]
        # Normaliser les accents pour la détection des colonnes
        def _norm(s): return (s.replace("é","e").replace("è","e").replace("ê","e")
                               .replace("à","a").replace("â","a").replace("î","i")
                               .replace("ô","o").replace("û","u").replace("ç","c"))
        cols_norm = [_norm(c) for c in cols_upper]
        if "usine" in cols_norm and "tonnage" in cols_norm:
            df.columns = [str(c).strip() for c in df.columns]
            usine_col = next(c for c in df.columns if _norm(c.lower()) == "usine")
            ton_col   = next(c for c in df.columns if _norm(c.lower()) == "tonnage")
            type_col  = next(c for c in df.columns if "type" in _norm(c.lower()) and "vehicule" in _norm(c.lower()))
            conf_col  = next((c for c in df.columns if _norm(c.lower()) in ("confirmation", "actif")), None)
            cont_col  = next((c for c in df.columns if _norm(c.lower()) == "contrat"), None)
            df["_usine"]   = df[usine_col].astype(str).str.strip().str.upper()
            df["_tonnage"] = pd.to_numeric(df[ton_col], errors="coerce")
            df["_type"]    = df[type_col].astype(str).str.strip().str.upper()
            df["_actif"]   = df[conf_col].astype(str).str.strip().str.lower() if conf_col else pd.Series("ok", index=df.index)
            df["_exclu"]   = df[cont_col].astype(str).str.strip().str.lower().str.contains("attente", na=False) if cont_col else False
        else:
            df["_usine"]   = df.iloc[:, 1].astype(str).str.strip().str.upper()
            df["_tonnage"] = pd.to_numeric(df.iloc[:, 0], errors="coerce")
            df["_type"]    = df.iloc[:, 4].astype(str).str.strip().str.upper()
            df["_actif"]   = df.iloc[:, 11].astype(str).str.strip().str.lower() if len(df.columns) > 11 else pd.Series("ok", index=df.index)
            df["_exclu"]   = False

        USINE_N = {
            "SICAM": "SICAM", "COMOCAP": "COMOCAP", "COMOCAB": "COMOCAP", "TUCAL": "TUCAL", "ABIDA": "ABIDA",
            "EL FALLEH": "ELFALLEH", "ELFALLEH": "ELFALLEH", "FALLEH": "ELFALLEH", "FELLA": "ELFALLEH",
            "LUI-MEME": "LUIMEME", "LUI-MÊME": "LUIMEME", "LUIMEME": "LUIMEME", "BOURAK": "BOURAK", "TOTAL": "SKIP",
        }
        df["_usine"] = df["_usine"].map(lambda x: USINE_N.get(x, x))

        def norm_vtype(t):
            t = str(t).strip().upper()
            if "SEMI" in t or "2*6" in t or "DOUBLE" in t or "REMORQUE" in t: return "SEMI"
            if "PPL" in t or "PELÉE" in t or "PELEE" in t: return "PPL"
            if t.startswith("PL"): return "PL"
            if "TRACTEUR" in t: return "TRACTEUR"
            return t
        df["_type"] = df["_type"].apply(norm_vtype)

        df_ok = df[(df["_actif"] == "ok") & (~df["_exclu"]) & df["_tonnage"].notna() & (df["_tonnage"] > 0) & (df["_usine"] != "SKIP")]
        fleet = defaultdict(lambda: defaultdict(list))
        for _, row in df_ok.iterrows():
            fleet[row["_usine"]][row["_type"]].append(float(row["_tonnage"]))
        result = {u: {vt: sorted(caps, reverse=True) for vt, caps in vtypes.items()} for u, vtypes in fleet.items()}
        return result, None
    except Exception as e:
        return {}, str(e)


# ═══ CHOOSE_VEHICLES — VERSION V5, port fidele de optimizer_v2.py ═══
def choose_vehicles(tons, allowed_raw, usine=None, region=None, semi_coeff=1.0, rm_day_rank=0, real_fleet=None):
    """
    Port fidele de la version V5 de optimizer_v2.py.
    real_fleet remplace la globale REAL_FLEET du module d'origine (injectee
    en parametre ici car ce module ne fait pas de solve, juste du reporting
    sur le plan deja rectifie).
    """
    real_fleet = real_fleet or {}
    _norm = {"PETIT POILOUR": "PPL", "POILOUR": "PL"}
    allowed = list(dict.fromkeys(_norm.get(v, v) for v in allowed_raw))
    if not allowed:
        allowed = ["PL"]

    def _alloc_real(veh, qty, usine_name=None):
        if qty <= 0:
            return []
        if veh == "TRACTEUR":
            return [{"vehicle": "TRACTEUR", "trips": 1, "tons_each": round(min(10.0, qty), 1), "real_cap": 10.0}]
        mn_veh, mx_veh = FLEET_CAPACITY.get(veh, (7, 25))
        _semi_only = (allowed == ["SEMI"])
        if veh == "SEMI" and _semi_only:
            SEMI_CAP = 30.0
            SEMI_MIN = 27.0
            if qty < SEMI_MIN:
                return []
            nb = max(1, int(round(qty / SEMI_CAP)))
            return [{"vehicle": "SEMI", "trips": nb, "tons_each": SEMI_CAP, "real_cap": SEMI_CAP}]
        real_caps = []
        if usine_name and usine_name in real_fleet:
            real_caps = list(real_fleet[usine_name].get(veh, []))
            if not real_caps and usine_name == "TUCAL":
                real_caps = list(real_fleet.get("BOURAK", {}).get(veh, []))
        if not real_caps:
            return _alloc_theory(veh, qty)
        MAX_SOLDE_FIXE = 4.0
        result = []
        remaining = round(qty, 1)
        for idx, cap in enumerate(real_caps):
            if remaining <= 0:
                break
            used = len(result)
            caps_restantes = real_caps[used:]
            min_dispo = min(caps_restantes) if caps_restantes else cap
            if result and (remaining <= MAX_SOLDE_FIXE or remaining < min_dispo * 0.5):
                break
            actual_load = min(cap, remaining)
            result.append({"vehicle": veh, "trips": 1, "tons_each": round(actual_load, 1), "real_cap": round(cap, 1), "solde": 0.0})
            remaining = round(remaining - actual_load, 1)
        if remaining > 0.05 and result:
            nb = len(result)
            solde_base = math.floor(remaining / nb * 10) / 10
            solde_last = round(remaining - solde_base * (nb - 1), 1)
            for i in range(nb):
                result[i]["solde"] = solde_last if i == nb - 1 else solde_base
        return result if result else _alloc_theory(veh, qty)

    def _alloc_theory(veh, qty):
        if qty <= 0:
            return []
        mn, mx = FLEET_CAPACITY.get(veh, (7, 25))
        if qty <= mx:
            return [{"vehicle": veh, "trips": 1, "tons_each": round(qty, 2)}]
        trips = math.ceil(qty / mx)
        each = qty / trips
        while each < mn and trips > 1:
            trips -= 1
            each = qty / trips
        base = int(qty // trips)
        extra = int(qty - base * trips)
        entries = []
        if extra:
            entries.append({"vehicle": veh, "trips": extra, "tons_each": round(base + 1, 2)})
        if trips - extra:
            entries.append({"vehicle": veh, "trips": trips - extra, "tons_each": round(base, 2)})
        return entries

    def _alloc(veh, qty):
        return _alloc_real(veh, qty, usine_name=usine)

    def _best_for_tons(qty, candidates):
        best, best_score, best_ratio, best_trips = None, 999, 0, 999
        for veh in candidates:
            mn_v, mx_v = FLEET_CAPACITY.get(veh, (7, 25))
            ratio = min(qty / mn_v, 1.0)
            if mn_v <= qty <= mx_v:
                score = 0; n_trips = 1
            elif qty > mx_v:
                trips = math.ceil(qty / mx_v)
                each = qty / trips
                if each >= mn_v:
                    score = 1; n_trips = trips
                else:
                    t2 = trips
                    while each < mn_v and t2 > 1:
                        t2 -= 1; each = qty / t2
                    score = 2 if each >= mn_v else 4; n_trips = t2
            elif qty >= mn_v * 0.8:
                score = 3; n_trips = 1
            else:
                score = 5; n_trips = 1
            if (score < best_score or (score == best_score and ratio > best_ratio) or
                    (score == best_score and ratio == best_ratio and n_trips < best_trips)):
                best_score, best, best_ratio, best_trips = score, veh, ratio, n_trips
        return best

    if allowed == ["SEMI"]:
        if tons <= 0 or tons < 27.0:
            return []
        nb = max(1, int(round(tons / 30.0)))
        return [{"vehicle": "SEMI", "trips": nb, "tons_each": 30.0, "real_cap": 30.0, "solde": 0.0}]

    _is_long_distance = (region and str(region).strip().upper() in
                          {"GAFSA / KASSRINE", "GAFSA/KASSRINE", "SIDI BOUZID", "KAIROUAN"})
    if _is_long_distance:
        USINE_PREFS = {"SICAM": ["SEMI", "PL", "PPL"], "TUCAL": ["SEMI", "PL", "PPL"],
                        "COMOCAP": ["SEMI", "PL", "PPL"], "ABIDA": ["PL", "SEMI", "PPL"],
                        "ELFALLEH": ["PPL", "PL", "SEMI"]}
    else:
        USINE_PREFS = {"SICAM": ["PL", "PPL", "SEMI"], "TUCAL": ["PL", "PPL", "SEMI"],
                        "COMOCAP": ["PL", "PPL", "SEMI"], "ABIDA": ["PL", "SEMI", "PPL"],
                        "ELFALLEH": ["PPL", "PL", "SEMI"]}
    prefs = USINE_PREFS.get(usine, ["PL", "PPL", "SEMI"])
    prefs_allowed = [v for v in prefs if v in allowed] or allowed

    _r = str(region or "").strip().upper()
    _region_effective = _r if _r not in ("", "AUTRE") else ""
    use_tracteur = (_region_effective in {"CAP BON 1"}) and ("TRACTEUR" in allowed or "TRC" in str(allowed_raw).upper())

    if usine == "COMOCAP":
        result = []
        trac_mn, trac_mx = FLEET_CAPACITY.get("TRACTEUR", (9, 11))
        if use_tracteur:
            trac_tons = min(trac_mx, max(trac_mn, round(tons * 0.14)))
            result.append({"vehicle": "TRACTEUR", "trips": 1, "tons_each": round(trac_tons, 2)})
            remaining = round(tons - trac_tons, 2)
        else:
            remaining = tons
        if remaining > 0:
            candidates = [v for v in ["PL", "PPL", "SEMI"] if v in allowed] or allowed
            main_veh = _best_for_tons(remaining, candidates) or candidates[0]
            result.extend(_alloc(main_veh, remaining))
        return result if result else _alloc(allowed[0] if allowed else "PL", tons)

    primary = _best_for_tons(tons, prefs_allowed) or _best_for_tons(tons, allowed) or (allowed[0] if allowed else "PL")
    if primary and primary in FLEET_CAPACITY:
        mn_p, mx_p = FLEET_CAPACITY[primary]
        if tons < mn_p * 0.8:
            smaller = {"SEMI": ["PL", "PPL"], "PL": ["PPL"], "PPL": []}
            for alt in smaller.get(primary, []):
                if alt in allowed:
                    alt_mn, alt_mx = FLEET_CAPACITY[alt]
                    if tons >= alt_mn * 0.5:
                        primary = alt
                        break

    # ── FIX: Préférer PPL quand tonnage ≤ PPL_max et PPL est autorisé ──
    # Si PL/PPL autorisés et tonnage ≤ 14t → PPL est plus adapté que PL
    # (PL min = 15t, charger 10t dans un PL = sous-utilisation)
    if "PPL" in allowed and primary == "PL":
        ppl_mn, ppl_mx = FLEET_CAPACITY.get("PPL", (6, 14))
        pl_mn, _ = FLEET_CAPACITY.get("PL", (15, 25))
        if tons <= ppl_mx and tons >= ppl_mn * 0.8:
            primary = "PPL"
        elif tons < pl_mn and tons >= ppl_mn * 0.5:
            primary = "PPL"

    result = _alloc(primary, tons)
    if not result:
        result = [{"vehicle": primary, "trips": 1, "tons_each": round(tons, 2)}]
    if "SEMI" not in allowed:
        cleaned = []
        for v in result:
            if v.get("vehicle") == "SEMI":
                best_alt = "PL" if "PL" in allowed else ("PPL" if "PPL" in allowed else None)
                if best_alt:
                    cleaned.extend(_alloc(best_alt, v["trips"] * v.get("tons_each", 0)))
            else:
                cleaned.append(v)
        if cleaned:
            result = cleaned
    return result


# ═══ AGREGATION POUR LE RAPPORT TRANSPORT ═══
def build_transport_detail(p_corrige, real_fleet=None):
    """
    p_corrige : DataFrame du planning DEJA CORRIGE (sortie de
    build_effective_planning) avec colonnes Commercial, Agriculteur, Usine,
    Région, Accessibilité, Date, Tonnes/Jour.

    Retourne un DataFrame detail (une ligne par voyage/vehicule alloue) avec
    colonnes : Date, Commercial, Usine, Région, Type Véhicule, Voyages, Tonnes.
    Calcule via choose_vehicles() sur le TONNAGE RECTIFIE (jamais sur le
    tonnage OR-Tools d'origine).
    """
    real_fleet = real_fleet or {}
    if p_corrige is None or p_corrige.empty:
        return pd.DataFrame(columns=["Date", "Commercial", "Agriculteur", "Usine", "Région", "Type Véhicule", "Voyages", "Tonnes"])

    # Detecter si Accessibilite est une VRAIE colonne du DataFrame
    _has_accessibilite = "Accessibilité" in p_corrige.columns
    rows = []
    for _, r in p_corrige.iterrows():
        tons = float(r.get("Tonnes/Jour", 0) or 0)
        if tons <= 0:
            continue
        usine = str(r.get("Usine", "") or "").strip()
        region = str(r.get("Région", "") or "").strip()

        if _has_accessibilite:
            # Accessibilite disponible -> recalculer les vehicules (meilleure precision)
            acc_raw = str(r.get("Accessibilité", "") or "").strip()
            acc = normalize_acc(acc_raw)
            allowed = ACCESS_VEHICLES.get(acc, ACCESS_VEHICLES["NAN"])
            vehicles = choose_vehicles(tons, allowed, usine=usine, region=region, real_fleet=real_fleet)
        else:
            # Pas d'Accessibilite (p_display du dashboard) -> se fier au
            # Type Vehicule deja calcule par build_effective_planning
            existing_veh = str(r.get("Type Véhicule", "") or "").strip().upper()
            if existing_veh and existing_veh in FLEET_CAPACITY:
                allowed = [existing_veh]
                vehicles = choose_vehicles(tons, allowed, usine=usine, region=region, real_fleet=real_fleet)
            else:
                allowed = ACCESS_VEHICLES["NAN"]
                vehicles = choose_vehicles(tons, allowed, usine=usine, region=region, real_fleet=real_fleet)
        if not vehicles:
            vehicles = [{"vehicle": allowed[0] if allowed else "PL", "trips": 1, "tons_each": tons}]
        for v in vehicles:
            rows.append({
                "Date": r.get("Date"), "Commercial": r.get("Commercial", ""), "Agriculteur": r.get("Agriculteur", ""),
                "Usine": usine, "Région": region, "Type Véhicule": v.get("vehicle", "PL"),
                "Voyages": int(v.get("trips", 1)),
                "Tonnes": round(v.get("trips", 1) * v.get("tons_each", 0) + v.get("solde", 0.0), 1),
            })
    return pd.DataFrame(rows)


def summarize_transport_usine(detail_df, real_fleet=None, period="day"):
    """
    Agrege par (Periode, Usine, Type Vehicule) UNIQUEMENT — c'est le seul
    niveau ou "Disponible vs Manquant" a un sens, car la flotte d'une usine
    est PARTAGEE entre tous les commerciaux qui y livrent ce jour-la (on ne
    duplique pas le compte de camions disponibles par commercial/region).
    """
    real_fleet = real_fleet or {}
    if detail_df is None or detail_df.empty:
        return pd.DataFrame()
    df = detail_df.copy()
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    df = df.dropna(subset=["Date"])
    df["Periode"] = df["Date"].dt.to_period("W-SUN").apply(lambda p: p.start_time) if period == "week" else df["Date"].dt.normalize()

    g = df.groupby(["Periode", "Usine", "Type Véhicule"], dropna=False).agg(
        Tonnes_requises=("Tonnes", "sum"), Voyages_requis=("Voyages", "sum")
    ).reset_index()
    g["Camions disponibles"] = g.apply(lambda r: len(real_fleet.get(r["Usine"], {}).get(r["Type Véhicule"], [])), axis=1)
    g["Tonnage disponible"] = g.apply(lambda r: sum(real_fleet.get(r["Usine"], {}).get(r["Type Véhicule"], [])), axis=1)
    g["Camions manquants (a louer)"] = (g["Voyages_requis"] - g["Camions disponibles"]).clip(lower=0).astype(int)
    g["Tonnage manquant (a louer)"] = (g["Tonnes_requises"] - g["Tonnage disponible"]).clip(lower=0).round(1)
    g = g.rename(columns={"Tonnes_requises": "Tonnes requises", "Voyages_requis": "Voyages requis"})
    return g.sort_values(["Periode", "Usine", "Type Véhicule"]).reset_index(drop=True)


def summarize_transport_detail(detail_df, period="day"):
    """
    Agrege par (Periode, Usine, Commercial, Région, Type Vehicule) — le
    besoin (tonnes/voyages) par commercial et par region, SANS recalculer
    Disponible/Manquant ici (cf. summarize_transport_usine pour ce niveau,
    la flotte n'etant pas sous-allouee par commercial/region)."""
    if detail_df is None or detail_df.empty:
        return pd.DataFrame()
    df = detail_df.copy()
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    df = df.dropna(subset=["Date"])
    df["Periode"] = df["Date"].dt.to_period("W-SUN").apply(lambda p: p.start_time) if period == "week" else df["Date"].dt.normalize()
    g = df.groupby(["Periode", "Usine", "Commercial", "Région", "Type Véhicule"], dropna=False).agg(
        Tonnes_requises=("Tonnes", "sum"), Voyages_requis=("Voyages", "sum")
    ).reset_index()
    g = g.rename(columns={"Tonnes_requises": "Tonnes requises", "Voyages_requis": "Voyages requis"})
    return g.sort_values(["Periode", "Usine", "Commercial", "Région", "Type Véhicule"]).reset_index(drop=True)