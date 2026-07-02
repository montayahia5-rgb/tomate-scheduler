# -*- coding: utf-8 -*-
"""
agroeco_dashboard.py v2 — Dashboard Agroéconomique Tomate 2026
===============================================================
CORRECTIONS v2 :
  ✅ Tous les fichiers entrée = "centre" + "client" obligatoires
  ✅ Royal : date_debut_repiquage → date_debut_livraison
  ✅ Caisses vides : condition = date_debut_RECOLTE (fichier rectifié Supabase)
     1ère affectation = date_debut_recolte < 10 juillet  → avec caisses
     2ème affectation = date_debut_recolte ≥ 10 juillet  → sans caisses
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import io
from datetime import date

# ══════════════════════════════════════════════════════════════
# CONSTANTES
# ══════════════════════════════════════════════════════════════
DATE_CAISSE_LIMITE = date(2026, 7, 10)

# Paramètres caisses vides PAR USINE (modifiables dans l'UI)
# Format : {usine: {"nb_ha": nb_caisses/ha, "prix": DT/caisse, "type": description}}
CAISSES_USINE_DEFAULTS = {
    "SICAM":    {"nb_ha": 80,  "prix": 3.0,  "type": "Caisse plastique 25kg",  "cap_kg": 25},
    "TUCAL":    {"nb_ha": 80,  "prix": 3.0,  "type": "Caisse plastique 25kg",  "cap_kg": 25},
    "COMOCAP":  {"nb_ha": 60,  "prix": 2.5,  "type": "Bac tracteur (forfait)", "cap_kg": 30},
    "ABIDA":    {"nb_ha": 60,  "prix": 2.5,  "type": "Caisse plastique 25kg",  "cap_kg": 25},
    "ELFALLEH": {"nb_ha": 50,  "prix": 2.0,  "type": "Caisse métal 20kg",      "cap_kg": 20},
}
MO_TONNE_DEFAULT   = 50.0
DENSITE_STD = {
    "CAP BON 1":25000,"CAP BON 2":25000,"NORD":22000,
    "KAIROUAN":20000,"BOUFICHA":22000,
    "GAFSA / KASSRINE":18000,"SIDI BOUZID":18000,
}
FAM_NORM_MAP = {
    "engrais":"Engrais","engrais ":"Engrais",
    "fertilissant":"Fertilisant","fertilisant":"Fertilisant",
    "fongicide":"Fongicide","insecticide":"Insecticide",
    "irrigations ":"Irrigation","irrigations":"Irrigation",
    "irrigations turk":"Irrigation","irrigation":"Irrigation","irrigation ":"Irrigation",
    "herbicide":"Herbicide","divers":"Divers","divers ":"Divers",
    "materiel":"Matériel","traitement":"Traitement",
}

# ══════════════════════════════════════════════════════════════
# HELPERS
# ══════════════════════════════════════════════════════════════
def _norm(c):
    return (str(c).strip().lower()
            .replace("é","e").replace("è","e").replace("ê","e")
            .replace("â","a").replace("ô","o").replace("î","i")
            .replace("û","u").replace(" ","_").replace("/","_")
            .replace("(","").replace(")","").replace("°","")
            .replace("'",""))

def _find_header(raw, keywords, max_rows=8):
    for i in range(min(max_rows, len(raw))):
        vals = [_norm(str(v)) for v in raw.iloc[i].values if pd.notna(v)]
        if any(any(kw in v for v in vals) for kw in keywords):
            return i
    return 0

def _read_auto(file_obj, keywords):
    raw = pd.read_excel(file_obj, sheet_name=0, header=None)
    hr  = _find_header(raw, keywords)
    file_obj.seek(0)
    df  = pd.read_excel(file_obj, sheet_name=0, header=hr)
    df.columns = [_norm(c) for c in df.columns]
    return df

def _find_col(df, cands):
    for c in cands:
        if c in df.columns: return c
    return None

def _check_required(df, source_name, required=["centre","client"]):
    """Vérifie que les colonnes obligatoires sont présentes."""
    missing = [c for c in required if c not in df.columns]
    if missing:
        return False, f"⚠️ {source_name} — colonnes manquantes : {missing}"
    return True, ""

def _metric(label, value, color="#f0f6fc", delta=None, delta_label=""):
    dh = ""
    if delta is not None:
        dc = "#3dd68c" if delta >= 0 else "#ef5350"
        sign = "+" if delta >= 0 else ""
        dh = f"<div style='font-size:.75rem;color:{dc}'>{sign}{delta:,.1f} {delta_label}</div>"
    return f"""<div style='background:#161b22;border:1px solid #30363d;border-radius:10px;
padding:12px 16px;border-top:3px solid {color}'>
<div style='font-size:.7rem;color:#8b949e;text-transform:uppercase;letter-spacing:.05em'>{label}</div>
<div style='font-size:1.4rem;font-weight:700;color:#f0f6fc'>{value}</div>{dh}</div>"""

# ══════════════════════════════════════════════════════════════
# SUPABASE — Date début récolte par agriculteur
# ══════════════════════════════════════════════════════════════
def load_date_debut_recolte(sb):
    """
    Charge depuis Supabase (plan_rectifie_detail) la date MIN de livraison
    par agriculteur = date début récolte.
    C'est cette date qui détermine l'affectation des caisses vides :
      < 10 juillet  → 1ère affectation (avec caisses vides)
      ≥ 10 juillet  → 2ème affectation (sans caisses vides)
    """
    if sb is None:
        return pd.DataFrame()
    try:
        data = []
        offset = 0
        while True:
            batch = sb.table("plan_rectifie_detail").select(
                "agriculteur,date"
            ).range(offset, offset+999).execute().data
            if not batch: break
            data.extend(batch)
            if len(batch) < 1000: break
            offset += 1000
        if not data:
            return pd.DataFrame()
        df = pd.DataFrame(data)
        df["date"] = pd.to_datetime(df["date"], errors="coerce")
        df = df.dropna(subset=["date"])
        # Min date = date début récolte par agriculteur
        result = df.groupby("agriculteur")["date"].min().reset_index()
        result.columns = ["agriculteur","date_debut_recolte"]
        result["agriculteur"] = result["agriculteur"].astype(str).str.strip().str.upper()
        # Déduire l'affectation
        result["affectation_caisse"] = result["date_debut_recolte"].apply(
            lambda d: "1ère (avec caisses)" if pd.notna(d) and d.date() < DATE_CAISSE_LIMITE
                      else "2ème (sans caisses)")
        return result
    except Exception as e:
        return pd.DataFrame()


def load_prevision_juin(sb):
    """Tonnage prévu Juin depuis plan_rectifie_detail (somme par agriculteur)."""
    if sb is None:
        return pd.DataFrame()
    try:
        data = []
        offset = 0
        while True:
            batch = sb.table("plan_rectifie_detail").select(
                "agriculteur,tonnes"
            ).range(offset, offset+999).execute().data
            if not batch: break
            data.extend(batch)
            if len(batch) < 1000: break
            offset += 1000
        if not data:
            return pd.DataFrame()
        df = pd.DataFrame(data)
        df["tonnes"] = pd.to_numeric(df["tonnes"], errors="coerce").fillna(0)
        grp = df.groupby("agriculteur")["tonnes"].sum().reset_index()
        grp.columns = ["agriculteur","prevision_juin"]
        grp["agriculteur"] = grp["agriculteur"].astype(str).str.strip().str.upper()
        return grp
    except Exception:
        return pd.DataFrame()


# ══════════════════════════════════════════════════════════════
# PARSERS — "centre" et "client" obligatoires dans tous les fichiers
# ══════════════════════════════════════════════════════════════

def parse_bourak(file_obj):
    """
    BOURAK — Financement & Transport
    Colonnes OBLIGATOIRES : client (agriculteur) · centre
    Colonnes attendues    : commercial · ingenieur · region ·
                            hectares · avance · report
    """
    df = _read_auto(file_obj,
         ["client","responsable","ingenieur","centre","avance","report"])

    MAP = {
        "client":     ["client","agriculteur","nom"],
        "commercial": ["responsable","commercial","resp"],
        "ingenieur":  ["ingenieur","ing","ingenieur_agronome"],
        "centre":     ["centre","centre_collecte"],
        "region":     ["region","zone"],
        "hectares":   ["hectares","ha","surface","nb_hectares"],
        "avance":     ["avance","avances","total_avance","montant_avance"],
        "report":     ["report","reste","solde_precedent","non_paye"],
    }
    rename = {}
    for tgt, cands in MAP.items():
        for c in cands:
            if c in df.columns and tgt not in rename.values():
                rename[c] = tgt; break
    df = df.rename(columns=rename)

    # Vérification colonnes obligatoires
    if "client" not in df.columns:
        return None, "BOURAK — colonne client manquante"
    if "centre" not in df.columns:
        df["centre"] = ""

    for c in ["hectares","avance","report"]:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)
    df["client"] = df["client"].astype(str).str.strip()
    # Filtrer : vides, TOTAL, et séparateurs "── COMM ──" générés par les fichiers test
    df = df[~df["client"].str.upper().isin(["","NAN","TOTAL","SOUS-TOTAL"])]
    df = df[~df["client"].str.startswith("──")]
    df = df[~df["client"].str.startswith("--")]
    df = df[df["client"].str.len() > 2]
    # S'assurer que commercial est bien présent (depuis "responsable" ou "ingenieur")
    if "commercial" not in df.columns:
        for _try_col in ["responsable","ingenieur","Responsable","Commercial"]:
            if _try_col in df.columns:
                df["commercial"] = df[_try_col].astype(str).str.strip()
                break
        else:
            df["commercial"] = ""
    return df, ""


def parse_royal(file_obj):
    """
    ROYAL — Plants livrés (Pépinière)
    Colonnes OBLIGATOIRES : client · centre
    Colonnes attendues    : zone · variete · qte_livree · valeur ·
                            date_debut_livraison · date_fin_livraison ·
                            type_plateau · nb_plateaux

    NOTE : date_debut_livraison = date de la 1ère livraison de plants
           (≠ date début récolte — celle-ci vient du fichier rectifié)
    """
    df = _read_auto(file_obj,
         ["client","centre","variete","quantite","date","livraison"])

    MAP = {
        "client":               ["client","agriculteur","nom"],
        "centre":               ["centre","centre_collecte"],
        "zone":                 ["zone","destination","direction","localisation"],
        "variete":              ["variete","article","type_plant"],
        "qte_livree":           ["quantite_livree","qte_livree","qte","plants_livres","nb_plants"],
        "valeur_plants":        ["valeur","montant","total","prix_total","valeur_plants"],
        "date_debut_livraison": ["date_debut_livraison","date_premiere_livraison",
                                 "date_debut","debut_livraison","debut"],
        "date_fin_livraison":   ["date_fin_livraison","date_fin","fin_livraison","fin"],
        "type_plateau":         ["type_plateau","unite","plateau"],
        "nb_plateaux":          ["nb_plateaux","plateaux","qte_plateaux"],
    }
    rename = {}
    for tgt, cands in MAP.items():
        for c in cands:
            if c in df.columns and tgt not in rename.values():
                rename[c] = tgt; break
    df = df.rename(columns=rename)

    if "client" not in df.columns:
        return None, "ROYAL — colonne client manquante"
    if "centre" not in df.columns:
        df["centre"] = ""

    for c in ["qte_livree","valeur_plants","nb_plateaux"]:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)
    for c in ["date_debut_livraison","date_fin_livraison"]:
        if c in df.columns:
            df[c] = pd.to_datetime(df[c], errors="coerce")

    df["client"] = df["client"].astype(str).str.strip()
    df = df[~df["client"].str.upper().isin(["","NAN","TOTAL"])]
    return df, ""


def parse_sotusfa(file_obj):
    """
    SOTUSFA — Engrais & Pesticides
    Colonnes OBLIGATOIRES : client · centre
    Colonnes attendues    : famille · article · qte · valeur
    """
    df = _read_auto(file_obj,
         ["client","agriculteur","centre","famille","article","valeur"])

    MAP = {
        "client":    ["client","agriculteur","nom"],
        "centre":    ["centre","centre_collecte"],
        "famille":   ["famille"],
        "article":   ["article","produit"],
        "qte":       ["qte","quantite"],
        "valeur":    ["total_ttc","total","valeur","montant"],
        "prix_u":    ["prix_un_ttc","prix_ttc","prix"],
    }
    rename = {}
    for tgt, cands in MAP.items():
        for c in cands:
            if c in df.columns and tgt not in rename.values():
                rename[c] = tgt; break
    df = df.rename(columns=rename)

    if "client" not in df.columns:
        return None, None, "SOTUSFA — colonne client/agriculteur manquante"
    if "centre" not in df.columns:
        df["centre"] = ""

    for c in ["qte","valeur","prix_u"]:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)
    df["client"] = df["client"].astype(str).str.strip()
    df = df[~df["client"].str.upper().isin(["","NAN","TOTAL"])]

    # Normaliser famille
    if "famille" in df.columns:
        df["famille_norm"] = df["famille"].astype(str).str.strip().str.lower()\
                             .map(FAM_NORM_MAP).fillna("Autre")

    # Pivot par (client + centre) + famille → une ligne par client
    if "famille_norm" in df.columns and "valeur" in df.columns:
        grp_cols = ["client"]
        if "centre" in df.columns:
            grp_cols.append("centre")
        pivot = df.groupby(grp_cols + ["famille_norm"])["valeur"].sum().unstack(
            fill_value=0).reset_index()
        # Normaliser noms colonnes (familles)
        new_cols = []
        for i, c in enumerate(pivot.columns):
            if i < len(grp_cols):
                new_cols.append(c)
            else:
                new_cols.append(_norm(str(c)))
        pivot.columns = new_cols
        pivot["total_intrants"] = pivot.select_dtypes("number").sum(axis=1)
    else:
        pivot = pd.DataFrame()

    return df, pivot, ""


def parse_quantite(file_obj):
    """
    TABLEAU QUANTITÉ — Plan livré / actif / extra
    Colonnes OBLIGATOIRES : client · centre
    Colonnes attendues    : qte_livree · qte_actif · qte_extra ·
                            tonnage_livre · prix_vente
    """
    df = _read_auto(file_obj,
         ["client","centre","livree","actif","extra","quantite","tonnage"])

    MAP = {
        "client":       ["client","agriculteur","nom"],
        "centre":       ["centre","centre_collecte"],
        "qte_livree":   ["quantite_livree","qte_livree","livree","plants_livres"],
        "qte_actif":    ["quantite_actif","qte_actif","actif","plants_actifs"],
        "qte_extra":    ["quantite_extra","qte_extra","extra","pertes"],
        "tonnage_livre":["tonnage_livre","tonnage","recolte","livraison_t"],
        "prix_vente":   ["prix_vente","prix","prix_unitaire_vente"],
    }
    rename = {}
    for tgt, cands in MAP.items():
        for c in cands:
            if c in df.columns and tgt not in rename.values():
                rename[c] = tgt; break
    df = df.rename(columns=rename)

    if "client" not in df.columns:
        return None, "TABLEAU QUANTITÉ — colonne client manquante"
    if "centre" not in df.columns:
        df["centre"] = ""

    for c in ["qte_livree","qte_actif","qte_extra","tonnage_livre","prix_vente"]:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)
    if "qte_extra" not in df.columns and "qte_livree" in df.columns and "qte_actif" in df.columns:
        df["qte_extra"] = df["qte_livree"] - df["qte_actif"]
    df["client"] = df["client"].astype(str).str.strip()
    df = df[~df["client"].str.upper().isin(["","NAN","TOTAL"])]
    return df, ""


def parse_prevision(file_obj, col_name):
    """
    Prévision Déc ou Mai.
    Accepte le format réel (responsable région / AGRICULTEUR / TONNAGE)
    centre est OPTIONNEL dans ce fichier.
    """
    try:
        df = _read_auto(file_obj,
             ["agriculteur","client","tonnage","prevision","responsable"])

        MAP = {
            "client":     ["client","agriculteur","nom"],
            "commercial": ["responsable_region","responsable region",
                           "commercial","responsable","resp"],
            "centre":     ["centre","centre_collecte"],
            "region":     ["region"],
            col_name:     ["tonnage","prevision","tonnes","quantite",
                           "tonnage_total","total_tonnage"],
        }
        rename = {}
        for tgt, cands in MAP.items():
            for c in cands:
                if c in df.columns and tgt not in rename.values():
                    rename[c] = tgt; break
        df = df.rename(columns=rename)

        # client obligatoire
        if "client" not in df.columns:
            return None, f"PRÉVISION ({col_name}) — colonne client/AGRICULTEUR manquante"

        df[col_name] = pd.to_numeric(df.get(col_name, 0), errors="coerce").fillna(0)
        df["client"] = df["client"].astype(str).str.strip()

        # Filtrer lignes TOTAL, vides, sous-totaux
        df = df[~df["client"].str.upper().str.strip().isin(
            ["","NAN","TOTAL","TOTAL FEDI","TOTAL MEKKI","TOTAL KHALIL",
             "TOTAL MAKKI","TOTAL ACHREF","TOTAL JILANI","TOTAL MAKKI BEN SALAH",
             "TOTAL ACHREF AJLANI","TOTAL JILANI OBAY","SOUS-TOTAL"])]
        df = df[df[col_name] > 0]

        # centre optionnel — créer vide si absent
        if "centre" not in df.columns:
            df["centre"] = ""

        keep = ["client","centre",col_name]
        for extra in ["commercial","region"]:
            if extra in df.columns:
                keep.append(extra)
        df = df[keep]
        # ← CLEF : sommer par client pour éviter double comptage
        # (un agriculteur qui livre à 2 usines = 2 lignes dans le fichier → 1 après sum)
        num_cols = [col_name]
        df = df.groupby("client", as_index=False)[num_cols].sum()
        df["centre"] = ""  # centre non disponible après groupby
        return df, ""
    except Exception as e:
        return None, str(e)


# ══════════════════════════════════════════════════════════════
# FUSION ET CALCULS
# ══════════════════════════════════════════════════════════════

def merge_and_calculate(df_bourak, df_royal, df_sotusfa_raw,
                        df_sotusfa_pivot, df_quantite,
                        df_prev_dec, df_prev_mai, df_prev_juin,
                        df_dates_recolte, params):
    """
    Fusionne les 4 sources sur (client + centre) et calcule tous les indicateurs.

    Clé de jointure : client + centre (présents dans tous les fichiers)

    df_dates_recolte : DataFrame(agriculteur, date_debut_recolte, affectation_caisse)
                       issu de Supabase plan_rectifie_detail
                       ⇒ détermine 1ère/2ème affectation caisses vides
    """
    # ── Base = BOURAK ──────────────────────────────────────
    if df_bourak is not None and not df_bourak.empty:
        base = df_bourak.copy()
    elif df_quantite is not None and not df_quantite.empty:
        base = df_quantite[["client","centre"]].copy()
    else:
        return pd.DataFrame()

    def _upper(df, cols):
        for c in cols:
            if c in df.columns:
                df[c] = df[c].astype(str).str.strip().str.upper()
        return df

    base = _upper(base, ["client","centre"])
    KEY = ["client","centre"]
    # Propager commercial depuis Bourak si disponible
    if "commercial" in base.columns:
        base["commercial"] = base["commercial"].astype(str).str.strip()
    else:
        base["commercial"] = ""

    # ── Merge ROYAL ────────────────────────────────────────
    if df_royal is not None and not df_royal.empty:
        r = _upper(df_royal.copy(), KEY)
        r_grp = r.groupby(KEY).agg(
            qte_royal       = ("qte_livree","sum"),
            valeur_plants   = ("valeur_plants","sum"),
            variete         = ("variete", lambda x: x.mode()[0] if len(x) else ""),
            zone            = ("zone", lambda x: x.mode()[0] if len(x) else ""),
            date_debut_liv  = ("date_debut_livraison","min"),
            date_fin_liv    = ("date_fin_livraison","max"),
        ).reset_index()
        # Consigne plateau
        prix_c = params.get("prix_consigne", {})
        if "type_plateau" in r.columns and "nb_plateaux" in r.columns:
            r["consigne_pl"] = r.apply(
                lambda row: row["nb_plateaux"] * prix_c.get(
                    str(row.get("type_plateau","")), 0), axis=1)
            r_cons = r.groupby(KEY)["consigne_pl"].sum().reset_index()
            r_grp  = r_grp.merge(r_cons, on=KEY, how="left")
            r_grp  = r_grp.rename(columns={"consigne_pl":"consigne_plateau"})
        else:
            r_grp["consigne_plateau"] = 0
        base = base.merge(r_grp, on=KEY, how="outer")

    # ── Merge dates récolte (Supabase) ─────────────────────
    # ⚠️  La condition caisses vides = date_debut_RECOLTE (pas livraison)
    if df_dates_recolte is not None and not df_dates_recolte.empty:
        dr = df_dates_recolte.copy()
        dr["client"] = dr["agriculteur"].astype(str).str.strip().str.upper()
        base = base.merge(
            dr[["client","date_debut_recolte","affectation_caisse"]],
            on="client", how="left")
        base["affectation_caisse"] = base["affectation_caisse"].fillna(
            "2ème (sans caisses)")
    else:
        base["affectation_caisse"] = "2ème (sans caisses)"
        base["date_debut_recolte"] = pd.NaT

    # ── Merge SOTUSFA ──────────────────────────────────────
    if df_sotusfa_pivot is not None and not df_sotusfa_pivot.empty:
        s = _upper(df_sotusfa_pivot.copy(), KEY)
        # Merge flexible : KEY disponibles dans les 2 DataFrames
        sot_key = [c for c in KEY if c in s.columns and c in base.columns]
        if sot_key:
            base = base.merge(s, on=sot_key, how="left")

    # ── Merge QUANTITÉ ─────────────────────────────────────
    if df_quantite is not None and not df_quantite.empty:
        q = _upper(df_quantite.copy(), KEY)
        base = base.merge(q, on=KEY, how="left")

    # ── Merge PRÉVISIONS (concordance + fuzzy matching) ────
    for df_p, col in [(df_prev_dec,"prevision_dec"),
                      (df_prev_mai,"prevision_mai"),
                      (df_prev_juin,"prevision_juin")]:
        if df_p is not None and not df_p.empty and col in df_p.columns:
            p = df_p.copy()
            # Appliquer concordance sur les noms du fichier prévision
            if "client" in p.columns:
                p["client"] = p["client"].apply(
                    lambda x: _get_concordance_key(x) or x)
                p[col] = pd.to_numeric(p[col], errors="coerce").fillna(0)
                p = p.groupby("client")[col].sum().reset_index()
            # Merge avec fuzzy
            if "client" in p.columns and "client" in base.columns:
                base, n_m, n_t = _fuzzy_match_clients(base, p, col)
            else:
                base[col] = np.nan

    # ══ CALCULS ═══════════════════════════════════════════
    df = base.copy()
    def g(col, d=0):
        return pd.to_numeric(df.get(col, d), errors="coerce").fillna(d)

    # Charges
    df["charge_plants"]   = g("valeur_plants")
    df["charge_intrants"] = g("total_intrants")
    df["avance_bourak"]   = g("avance")
    df["charge_totale"]   = df["charge_plants"] + df["charge_intrants"] + df["avance_bourak"]

    # ── Enrichir commercial si absent (depuis quantite ou bourak) ──
    if "commercial" not in df.columns or df["commercial"].fillna("").eq("").all():
        # Essayer depuis df_quantite
        if df_quantite is not None and "commercial" in df_quantite.columns and "client" in df_quantite.columns:
            _q_comm = _upper(df_quantite[["client","commercial"]].dropna().drop_duplicates("client").copy(), ["client"])
            df = df.merge(_q_comm.rename(columns={"commercial":"_comm_q"}), on="client", how="left")
            if "_comm_q" in df.columns:
                df["commercial"] = df.get("commercial", pd.Series([""]*len(df))).fillna("").replace("", pd.NA)
                df["commercial"] = df["commercial"].fillna(df["_comm_q"])
                df = df.drop(columns=["_comm_q"])
        # Fallback : depuis Bourak
        if df_bourak is not None and "commercial" in df_bourak.columns and "client" in df_bourak.columns:
            _b_comm = _upper(df_bourak[["client","commercial"]].dropna().drop_duplicates("client").copy(), ["client"])
            _b_comm.columns = ["client","_comm_b"]
            df = df.merge(_b_comm, on="client", how="left")
            if "_comm_b" in df.columns:
                df["commercial"] = df["commercial"].fillna("").replace("", pd.NA).fillna(df["_comm_b"])
                df = df.drop(columns=["_comm_b"])
    if "commercial" not in df.columns:
        df["commercial"] = ""
    df["commercial"] = df["commercial"].fillna("").astype(str).str.strip()

    # Consigne caisse — PAR USINE (1ère affectation uniquement)
    caisses_par_usine = params.get("caisses_par_usine", {})
    # Fallback global si pas de config par usine
    _px_global  = params.get("prix_caisse", 3.0)
    _nb_global  = params.get("nb_caisses_ha", 80.0)

    def _calc_caisse(row):
        if not str(row.get("affectation_caisse","")).startswith("1ère"):
            return 0.0
        ha = float(row.get("hectares", 0) or 0)
        # Déterminer l'usine de l'agriculteur
        usine = str(row.get("usine", row.get("usine_livraison", ""))).upper().strip()
        # Chercher dans les usines connues
        cfg = None
        for u_key in caisses_par_usine:
            if u_key.upper() in usine or usine in u_key.upper():
                cfg = caisses_par_usine[u_key]
                break
        if cfg:
            return round(ha * cfg["nb_ha"] * cfg["prix"], 2)
        # Fallback global
        return round(ha * _nb_global * _px_global, 2)

    df["consigne_caisse"] = df.apply(_calc_caisse, axis=1)

    # Détail par usine pour affichage
    def _detail_caisse(row):
        if not str(row.get("affectation_caisse","")).startswith("1ère"):
            return "2ème — 0 DT"
        ha = float(row.get("hectares", 0) or 0)
        usine = str(row.get("usine", "")).upper().strip()
        for u_key, cfg in caisses_par_usine.items():
            if u_key.upper() in usine or usine in u_key.upper():
                nb = cfg["nb_ha"]; px = cfg["prix"]
                total = round(ha * nb * px, 0)
                return f"1ère — {int(ha*nb)} caisses × {px} DT = {total:,.0f} DT"
        nb = _nb_global; px = _px_global
        return f"1ère — {int(ha*nb)} caisses × {px} DT = {round(ha*nb*px):,.0f} DT"

    df["detail_caisse"] = df.apply(_detail_caisse, axis=1)

    df["consigne_plateau"] = g("consigne_plateau")
    mo = params.get("mo_tonne", MO_TONNE_DEFAULT)
    df["tonnage_livre"]    = g("tonnage_livre")
    df["mo_recolte"]       = df["tonnage_livre"] * mo

    # Plants
    df["qte_livree"]  = g("qte_livree")
    df["qte_actif"]   = g("qte_actif")
    df["qte_extra"]   = g("qte_extra")
    df["hectares"]    = g("hectares")

    df["taux_prise"]  = np.where(df["qte_livree"]>0,
                         (df["qte_actif"]/df["qte_livree"]*100).round(1), 0)
    df["densite_ha"]  = np.where(df["hectares"]>0,
                         (df["qte_actif"]/df["hectares"]).round(0), 0)

    # Prix vente
    df["prix_vente"]  = g("prix_vente")
    df["prix_vente"]  = df["prix_vente"].where(df["prix_vente"]>0,
                         params.get("prix_vente_global", 0))

    # Tonnage recouvrement
    charges_totales = (df["charge_totale"] + df["consigne_plateau"]
                     + df["consigne_caisse"] + df["mo_recolte"])
    df["charges_totales"]       = charges_totales
    df["tonnage_recouvrement"]  = np.where(df["prix_vente"]>0,
                                   (charges_totales / df["prix_vente"]).round(2), 0)

    # Indicateurs /ha
    df["recouvrement_ha"]   = np.where(df["hectares"]>0,
                               (df["tonnage_recouvrement"]/df["hectares"]).round(2),0)
    df["tonnage_ha_realise"]= np.where(df["hectares"]>0,
                               (df["tonnage_livre"]/df["hectares"]).round(2),0)
    df["cout_ha"]           = np.where(df["hectares"]>0,
                               (df["charge_totale"]/df["hectares"]).round(0),0)
    df["cout_plant_actif"]  = np.where(df["qte_actif"]>0,
                               (df["charge_totale"]/df["qte_actif"]).round(4),0)

    # Solde et valeur
    df["valeur_livree"] = (df["tonnage_livre"] * df["prix_vente"]).round(0)
    df["ecart_tonnage"] = (df["tonnage_livre"] - df["tonnage_recouvrement"]).round(2)
    df["solde_final"]   = (df["valeur_livree"] - charges_totales).round(0)
    df["report"]        = g("report")

    # Alertes
    def _alerte(row):
        ecart      = row.get("ecart_tonnage", 0) or 0
        taux       = row.get("taux_prise", 100) or 100
        report_v   = row.get("report", 0) or 0
        charge     = row.get("charge_totale", 1) or 1
        prev_mai   = row.get("prevision_mai", 0) or 0
        recouvr    = row.get("tonnage_recouvrement", 0) or 0

        if ecart < -5:
            return "🔴 DÉFICIT RECOUVREMENT"
        if taux < 85:
            return "🔴 PRISE FAIBLE"
        if report_v > charge * 0.5:
            return "🔴 RISQUE FINANCIER"
        if recouvr > 0 and prev_mai > 0:
            ratio = prev_mai / recouvr
            if ratio < 0.90: return "🔴 PRÉVISION INSUFFISANTE"
            if ratio < 1.00: return "🟡 ATTENTION"
        if ecart >= 0:
            return "🟢 OK"
        return "🟡 ATTENTION"

    df["alerte"] = df.apply(_alerte, axis=1)

    # Renommer client → agriculteur pour affichage
    df = df.rename(columns={"client":"agriculteur"})
    return df


# ══════════════════════════════════════════════════════════════
# EXPORT EXCEL
# ══════════════════════════════════════════════════════════════

def export_excel(df, df_sotusfa_raw=None):
    from openpyxl import Workbook
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.formatting.rule import ColorScaleRule, DataBarRule

    def hf(h): return PatternFill("solid",start_color=h,end_color=h)
    def bf(bold=True,white=False,color="000000",size=10):
        return Font(bold=bold,name="Calibri",size=size,
                    color="FFFFFF" if white else color)
    t = Side(style="thin",color="CCCCCC")
    BD  = Border(left=t,right=t,top=t,bottom=t)
    CTR = Alignment(horizontal="center",vertical="center",wrap_text=True)
    LFT = Alignment(horizontal="left",vertical="center")

    wb = Workbook()

    # ── Feuille 1 : Dashboard ──────────────────────────────
    ws = wb.active
    ws.title = "📊 Dashboard"
    ws.sheet_view.showGridLines = False

    BLOCS = [
        ("IDENTIFICATION",    ["agriculteur","commercial","ingenieur","centre","region"],
         "1F3864"),
        ("PLANT",             ["variete","hectares","qte_livree","qte_actif","qte_extra",
                                "taux_prise","densite_ha"],          "1A5C2A"),
        ("AFFECTATION",       ["affectation_caisse","date_debut_recolte"], "8B3A00"),
        ("CHARGES (DT)",      ["charge_plants","charge_intrants","avance_bourak",
                                "charge_totale","consigne_plateau","consigne_caisse",
                                "mo_recolte","charges_totales"],     "375623"),
        ("PRÉVISIONS (T)",    ["prevision_dec","prevision_mai",
                                "prevision_juin","tonnage_livre"],   "4A235A"),
        ("RECOUVREMENT",      ["prix_vente","tonnage_recouvrement",
                                "recouvrement_ha","ecart_tonnage"],  "C0392B"),
        ("RÉSULTAT",          ["tonnage_ha_realise","cout_ha","cout_plant_actif",
                                "valeur_livree","solde_final","report","alerte"],
         "0B4F6C"),
    ]

    # Titre
    all_cols = [c for bloc in BLOCS for c in bloc[1]]
    avail_cols = [c for c in all_cols if c in df.columns]
    ws.merge_cells(start_row=1,start_column=1,end_row=1,end_column=len(avail_cols))
    ws["A1"] = "📊 DASHBOARD AGROÉCONOMIQUE TOMATE 2026 — v2"
    ws["A1"].font = bf(True,white=True,size=13)
    ws["A1"].fill = hf("0D1B2A"); ws["A1"].alignment = CTR
    ws.row_dimensions[1].height = 32

    # Ligne 2 : blocs
    ci = 1
    BLOC_STARTS = {}
    for bloc_name, bloc_cols, bloc_color in BLOCS:
        bloc_avail = [c for c in bloc_cols if c in df.columns]
        if not bloc_avail: continue
        BLOC_STARTS[bloc_name] = (ci, ci+len(bloc_avail)-1, bloc_color)
        ws.merge_cells(start_row=2,start_column=ci,
                       end_row=2,end_column=ci+len(bloc_avail)-1)
        c = ws.cell(2,ci,value=bloc_name)
        c.font=bf(True,white=True,size=8); c.fill=hf(bloc_color)
        c.alignment=CTR; c.border=BD
        ci += len(bloc_avail)
    ws.row_dimensions[2].height = 18

    # Ligne 3 : en-têtes
    HDRS_LABELS = {
        "agriculteur":"Agriculteur","commercial":"Commercial",
        "ingenieur":"Ingénieur","centre":"Centre","region":"Région",
        "variete":"Variété","hectares":"Ha","qte_livree":"Plants Livrés",
        "qte_actif":"Plants Actifs","qte_extra":"Extra (pertes)",
        "taux_prise":"Taux prise %","densite_ha":"Densité/ha",
        "affectation_caisse":"Affectation","date_debut_recolte":"Déb. Récolte",
        "charge_plants":"Plants (DT)","charge_intrants":"Intrants (DT)",
        "avance_bourak":"Avance Bourak (DT)","charge_totale":"Charge Totale (DT)",
        "consigne_plateau":"Consigne Plateau","consigne_caisse":"Consigne Caisse",
        "mo_recolte":"MO Récolte (DT)","charges_totales":"TOTAL Charges",
        "prevision_dec":"Prév. Déc (T)","prevision_mai":"Prév. Mai (T)",
        "prevision_juin":"Prév. Juin (T)","tonnage_livre":"Livré (T)",
        "prix_vente":"Prix Vente","tonnage_recouvrement":"RECOUVREMENT (T)",
        "recouvrement_ha":"Recouv./ha","ecart_tonnage":"Écart (T)",
        "tonnage_ha_realise":"T/ha réalisé","cout_ha":"Coût/ha",
        "cout_plant_actif":"Coût/plant","valeur_livree":"Valeur Livrée",
        "solde_final":"Solde Final","report":"Report (DT)","alerte":"Alerte",
    }
    COL_WIDTHS = {
        "agriculteur":26,"commercial":16,"ingenieur":18,"centre":16,"region":16,
        "variete":14,"hectares":8,"qte_livree":13,"qte_actif":13,"qte_extra":12,
        "taux_prise":11,"densite_ha":11,"affectation_caisse":18,
        "date_debut_recolte":14,"charge_plants":13,"charge_intrants":14,
        "avance_bourak":15,"charge_totale":15,"consigne_plateau":15,
        "consigne_caisse":15,"mo_recolte":13,"charges_totales":14,
        "prevision_dec":13,"prevision_mai":13,"prevision_juin":13,
        "tonnage_livre":12,"prix_vente":11,"tonnage_recouvrement":17,
        "recouvrement_ha":13,"ecart_tonnage":12,"tonnage_ha_realise":13,
        "cout_ha":12,"cout_plant_actif":13,"valeur_livree":14,
        "solde_final":13,"report":12,"alerte":18,
    }
    ci = 1
    col_color_map = {}
    for bloc_name, bloc_cols, bloc_color in BLOCS:
        for col in bloc_cols:
            if col not in df.columns: continue
            col_color_map[col] = bloc_color
            c = ws.cell(3,ci,value=HDRS_LABELS.get(col,col))
            c.font=bf(True,white=True,size=9); c.fill=hf(bloc_color)
            c.alignment=CTR; c.border=BD
            ws.column_dimensions[get_column_letter(ci)].width = COL_WIDTHS.get(col,12)
            ci += 1
    ws.row_dimensions[3].height = 34

    # Données
    ALERT_BG = {
        "🔴 DÉFICIT RECOUVREMENT":"5c1a1a",
        "🔴 PRISE FAIBLE":         "5c1a1a",
        "🔴 RISQUE FINANCIER":     "5c1a1a",
        "🔴 PRÉVISION INSUFFISANTE":"5c1a1a",
        "🟡 ATTENTION":            "3d3000",
        "🟢 OK":                   "0a2a0a",
    }
    NUM_FMT = {
        "qte_livree":"#,##0","qte_actif":"#,##0","qte_extra":"#,##0",
        "charge_plants":"#,##0","charge_intrants":"#,##0",
        "avance_bourak":"#,##0","charge_totale":"#,##0",
        "consigne_plateau":"#,##0","consigne_caisse":"#,##0",
        "mo_recolte":"#,##0","charges_totales":"#,##0",
        "valeur_livree":"#,##0","solde_final":"+#,##0;-#,##0;0",
        "report":"#,##0","cout_ha":"#,##0",
        "taux_prise":"0.0","densite_ha":"#,##0",
        "tonnage_livre":"0.0","tonnage_recouvrement":"0.0",
        "ecart_tonnage":"+0.0;-0.0;0","recouvrement_ha":"0.0",
        "tonnage_ha_realise":"0.0","cout_plant_actif":"0.0000",
        "prix_vente":"0.0",
    }

    for ri, (_, row) in enumerate(df.iterrows()):
        r = ri + 4
        alt = ri % 2 == 0
        alerte_val = str(row.get("alerte","🟢 OK"))
        ci2 = 1
        for bloc_name, bloc_cols, bloc_color in BLOCS:
            for col in bloc_cols:
                if col not in df.columns: continue
                val = row.get(col,"")
                if isinstance(val, float) and np.isnan(val): val = ""
                if isinstance(val, pd.Timestamp): val = val.strftime("%d/%m/%Y")
                c = ws.cell(r,ci2,value=val)
                c.border = BD
                c.alignment = LFT if col in ("agriculteur","ingenieur","centre") else CTR
                c.font = bf(False,size=9)
                if col == "alerte":
                    c.fill = hf(ALERT_BG.get(alerte_val,"0a2a0a"))
                    c.font = bf(True,white=True,size=9)
                elif col == "ecart_tonnage":
                    v = row.get(col,0) or 0
                    c.fill = hf("1a3a1a" if v>=0 else "3a1a1a")
                    c.font = bf(True,size=9,
                                color="4CAF50" if v>=0 else "ef5350")
                elif col == "solde_final":
                    v = row.get(col,0) or 0
                    c.fill = hf("E8F5E9" if v>=0 else "FFEBEE")
                    c.font = bf(False,size=9,
                                color="1E8449" if v>=0 else "C0392B")
                else:
                    c.fill = hf("F0F5FF" if alt else "FFFFFF")
                if col in NUM_FMT:
                    c.number_format = NUM_FMT[col]
                ci2 += 1
        ws.row_dimensions[r].height = 18

    last = len(df)+4
    # Color scale écart tonnage
    if "ecart_tonnage" in df.columns:
        idx = avail_cols.index("ecart_tonnage")+1
        col_l = get_column_letter(idx)
        ws.conditional_formatting.add(
            f"{col_l}4:{col_l}{last}",
            ColorScaleRule(start_type="min",start_color="FFCDD2",
                           mid_type="num",mid_value=0,mid_color="FFF9C4",
                           end_type="max",end_color="C8E6C9"))
    ws.freeze_panes = "A4"

    # ── Feuille 2 : Par Ingénieur ──────────────────────────
    if "ingenieur" in df.columns:
        ws2 = wb.create_sheet("👤 Par Ingénieur")
        ws2.sheet_view.showGridLines = False
        g2 = df.groupby(["ingenieur","centre"]).agg(
            Agriculteurs   = ("agriculteur","count"),
            Hectares       = ("hectares","sum"),
            Charge_DT      = ("charge_totale","sum"),
            Plants_livres  = ("qte_livree","sum"),
            Plants_actifs  = ("qte_actif","sum"),
            Taux_prise     = ("taux_prise","mean"),
            Recouvrement_T = ("tonnage_recouvrement","sum"),
            Livre_T        = ("tonnage_livre","sum"),
            Ecart_T        = ("ecart_tonnage","sum"),
            Alertes_rouge  = ("alerte",lambda x:(x.str.contains("🔴")).sum()),
        ).reset_index()
        g2 = g2.round(1)
        ws2.merge_cells(start_row=1,start_column=1,end_row=1,end_column=len(g2.columns))
        ws2["A1"]="👤 Synthèse par Ingénieur × Centre"
        ws2["A1"].font=bf(True,white=True,size=12)
        ws2["A1"].fill=hf("0B4F6C"); ws2["A1"].alignment=CTR
        ws2.row_dimensions[1].height=26
        for ci3,col in enumerate(g2.columns,1):
            c=ws2.cell(2,ci3,value=col)
            c.font=bf(True,white=True,size=10); c.fill=hf("0B4F6C")
            c.alignment=CTR; c.border=BD
            ws2.column_dimensions[get_column_letter(ci3)].width=max(14,len(str(col))+4)
        ws2.row_dimensions[2].height=24
        for ri,(_, row) in enumerate(g2.iterrows()):
            r2=ri+3
            for ci3,val in enumerate(row,1):
                if isinstance(val,float) and np.isnan(val): val=0
                c=ws2.cell(r2,ci3,value=val)
                c.border=BD; c.alignment=CTR
                c.fill=hf("F0F5FF" if ri%2==0 else "FFFFFF")
                c.font=bf(False,size=9)
                if g2.columns[ci3-1]=="Alertes_rouge" and val>0:
                    c.fill=hf("FFEBEE"); c.font=bf(True,size=9,color="C0392B")
        ws2.freeze_panes="A3"

    # ── Feuille 3 : Caisses vides detail ──────────────────
    ws3 = wb.create_sheet("📦 Caisses Vides")
    ws3.sheet_view.showGridLines = False
    caisse_cols = [c for c in ["agriculteur","centre","region","usine","affectation_caisse",
                                "detail_caisse","date_debut_recolte","hectares",
                                "consigne_caisse","consigne_plateau"] if c in df.columns]
    if caisse_cols:
        ws3.merge_cells(start_row=1,start_column=1,end_row=1,end_column=len(caisse_cols))
        ws3["A1"]="📦 Détail Caisses Vides — 1ère vs 2ème Affectation"
        ws3["A1"].font=bf(True,white=True,size=12)
        ws3["A1"].fill=hf("8B3A00"); ws3["A1"].alignment=CTR
        ws3.row_dimensions[1].height=26
        for ci3,col in enumerate(caisse_cols,1):
            c=ws3.cell(2,ci3,value=HDRS_LABELS.get(col,col))
            c.font=bf(True,white=True,size=10); c.fill=hf("8B3A00")
            c.alignment=CTR; c.border=BD
            ws3.column_dimensions[get_column_letter(ci3)].width=18
        ws3.row_dimensions[2].height=26
        for ri,(_, row) in enumerate(df[caisse_cols].iterrows()):
            r3=ri+3
            is_1ere = "1ère" in str(row.get("affectation_caisse",""))
            for ci3,val in enumerate(row,1):
                if isinstance(val,pd.Timestamp): val=val.strftime("%d/%m/%Y")
                if isinstance(val,float) and np.isnan(val): val=""
                c=ws3.cell(r3,ci3,value=val); c.border=BD; c.alignment=CTR
                c.fill=hf("FFF3E0" if is_1ere else "E8F5E9")
                c.font=bf(bold=is_1ere,size=9,
                           color="8B3A00" if is_1ere else "1A5C2A")
        ws3.freeze_panes="A3"

    # ── Feuille 4 : Prévisions ─────────────────────────────
    ws4 = wb.create_sheet("📈 Prévisions")
    ws4.sheet_view.showGridLines = False
    prev_c = [c for c in ["agriculteur","centre","prevision_dec","prevision_mai",
                           "prevision_juin","tonnage_livre","tonnage_recouvrement",
                           "recouvrement_ha","ecart_tonnage"] if c in df.columns]
    if len(prev_c) > 2:
        ws4.merge_cells(start_row=1,start_column=1,end_row=1,end_column=len(prev_c))
        ws4["A1"]="📈 Évolution Prévisions Déc → Mai → Juin → Réalisé vs Recouvrement"
        ws4["A1"].font=bf(True,white=True,size=12)
        ws4["A1"].fill=hf("4A235A"); ws4["A1"].alignment=CTR
        ws4.row_dimensions[1].height=26
        for ci3,col in enumerate(prev_c,1):
            c=ws4.cell(2,ci3,value=HDRS_LABELS.get(col,col))
            c.font=bf(True,white=True,size=10); c.fill=hf("4A235A")
            c.alignment=CTR; c.border=BD
            ws4.column_dimensions[get_column_letter(ci3)].width=15
        ws4.row_dimensions[2].height=24
        for ri,(_, row) in enumerate(df[prev_c].iterrows()):
            r4=ri+3
            for ci3,val in enumerate(row,1):
                if isinstance(val,float) and np.isnan(val): val=""
                c=ws4.cell(r4,ci3,value=val); c.border=BD; c.alignment=CTR
                c.fill=hf("F0F0FF" if ri%2==0 else "FFFFFF")
                c.font=bf(False,size=9)
                col_n = prev_c[ci3-1]
                if col_n == "ecart_tonnage":
                    try:
                        v=float(val)
                        c.fill=hf("E8F5E9" if v>=0 else "FFEBEE")
                        c.font=bf(True,size=9,
                                  color="1E8449" if v>=0 else "C0392B")
                        c.number_format="+0.0;-0.0;0"
                    except: pass
        ws4.freeze_panes="A3"

    buf = io.BytesIO()
    wb.save(buf); buf.seek(0)
    return buf.read()


# ══════════════════════════════════════════════════════════════
# RENDER PRINCIPAL
# ══════════════════════════════════════════════════════════════




# ══════════════════════════════════════════════════════════════
# TABLE DE CONCORDANCE MANUELLE
# Mappe les noms de la RÉFÉRENCE vers les noms exacts dans SOTUSFA
# Résout les 3 types de problèmes :
#   1. Fautes de frappe : ATTIYA → ATTIAA
#   2. Entités splittées : BILEL GHA 1/2/3 → SOCIETE BILEL GHA
#   3. Alias : STE BACCARA → SOCIETE BACCARA ET FILS
# ══════════════════════════════════════════════════════════════
CONCORDANCE_NOMS = {
    # ── Fautes d'orthographe (CAUSE 1) ──────────────────────
    "SAMIR ATTIYA":               "SAMIR ATTIAA",
    "SAMIR ATTIA":                "SAMIR ATTIAA",
    "NEGI ZAAFOURI":              "NEJI ZAAFOURI",
    "ALI KOTLI":                  "ALI EL KOTLI",
    "ABDELKADER KALBOUSI":        "ABDELKADER KALBOUSSI",
    "MAKRAM HAFFAR":              "MAKREM HAFFAR",
    "MAKREM HAFFAR":              "MAKREM HAFFAR",
    "LASSED NEILI":               "LASSAAD NEILI",
    "SALEH BEN HAMOUDA":          "SALAH BEN HAMOUDA",
    "HSSINE BRINI":               "HSSIN BEN MED BEN ABDELKADER BRINI",
    "HSSINE BRINI":               "HSSIN BEN MED BEN ABDELKADER BRINI",
    "SAMI KAAB":                  "SAMI BEN HEDI KAAB",
    "ABDELFATEH BEN SLIMEN":      "ABDELFATEH BEN SLIMENE",
    "HAMED BEN YOUNIS":           "HAMED BEN YOUNES",
    "TAREK BEN ABDALAH":          "TAREK BEN ABDALLAH",
    "TAREK EL BAHRI":             "TAREK ELBAHRI",
    "TAREK ELBAHRI":              "TAREK ELBAHRI",
    "SEBTI JABALI":               "Sebti jaballi",
    "SOUHAIL BOUZANA":            "SOUHAIL BOUZENA",
    "MOHAMED THAMER BEN ALAYA":   "MOHAMED THAMEUR BEN ALAYA",
    "SASSI MANSOUR":              "SASSI BEN MANSOUR",
    "HASSEN BEN ALIA":            "HASSEN BEN ALAYA",
    "AZAIZ BEN ISSA":             "AZAIZ BEN ISSA",
    "HAMMADI BENZRIBIA":          "HAMMADI BEN ZRIBIA",
    "HAFEDH MOSBEH":              "HAFEDH MESBEH",
    "MOHAMED BEN HEDI MEHEMDI":   "MOHAMED  HEDI MEHEMDI",
    "MAHER BELHAJ FRAJ":          "MAHER BEL HAJ FRAJ",
    "ZOUHAIR BEN ECHIK":          "ZOUHAIR BEN  ECHIK",
    "NOOMEN ECHAGRAOUI":          "NOOMEN ECHAGRAOUI",
    "MOHAMED ALI MBAREK":         "MOHAMED ALI  BEN MBAREK",
    "BADIA SAAFI":                "BADIA SAAFI",
    "IBRAHIM BEN BOUBAKER":       "IBRAHIM BEN BOUBAKER",
    "IMED TRABILSI":              "IMED TRABILSI",
    "Mohamed BEDIA NEJI":         "MOHAMED BADIA NEJI",
    "CHOKRI SAAFI":               "CHOKRI SAAFI",
    "HICHEM SAAFI":               "HICHEM SAAFI",

    # ── Entités splittées dans référence (CAUSE 2) ──────────
    # BILEL GHA (4 lignes référence → 1 entrée Sotusfa)
    "BILEL GHA 1":                "SOCIETE BILEL GHA SERVICE AGRICOLE",
    "BILEL GHA 2":                "SOCIETE BILEL GHA SERVICE AGRICOLE",
    "BILEL GHA 3":                "SOCIETE BILEL GHA SERVICE AGRICOLE",
    "BILEL GHA 4":                "SOCIETE BILEL GHA SERVICE AGRICOLE",
    "BILEL KEHIL":                "SOCIETE BILEL GHA SERVICE AGRICOLE",
    # GARMALAH (plusieurs lignes → 1 Sotusfa)
    "KARIM GARMALAH 1":           "ABDELKARIM GARMALLAH",
    "KARIM GARMALAH 2":           "ABDELKARIM GARMALLAH",
    "AMAR GARMALAH":              "ABDELKARIM GARMALLAH",
    "MED ALI GARMALAH":           "ABDELKARIM GARMALLAH",
    "HSAN GARMALAH":              "ABDELKARIM GARMALLAH",
    "JAMEL GARMALAH":             "ABDELKARIM GARMALLAH",
    "KARIM AMAR":                 "ABDELKARIM GARMALLAH",
    "KARIM GARMALAH":             "ABDELKARIM GARMALLAH",
    "ABDELKARIM SAAD":            "ABDELKARIM GARMALLAH",
    "HAFEDH MOSBEH":              "HAFEDH MESBEH",
    # BACCARA
    "STE BACCARA":                "SOCIETE BACCARA ET FILS",
    "SOCIETE BACCARA":            "SOCIETE BACCARA ET FILS",
    "SAMEH BACCOUCH":             "SAMAH BACCOUCH",
    "NAJIB BACCOUCH":             "NEJIB BAKOUCHE",
    # MANSOURI (plusieurs membres famille)
    "NOUREDIN MANSOURI":          "MOURAD MANSOURI",
    "ELIFA MANSOURI":             "MOURAD MANSOURI",
    "TAHER MANSOURI":             "MOURAD MANSOURI",
    "ABELSAMII MANSOURI":         "MOURAD MANSOURI",
    "CHOKRI MANSOURI":            "MOURAD MANSOURI",
    "LAMINE MANSOURI":            "MOURAD MANSOURI",
    "AHMED MANSOURI":             "MOURAD MANSOURI",

    # ── Alias et abréviations (CAUSE 2 bis) ─────────────────
    "HICHEM TRABELSI":            "ABDELHEDI TRABELSSI",
    "KAMEL TRABELSSI":            "ABDELHEDI TRABELSSI",
    "MOUHAMED TRABELSI":          "ABDELHEDI TRABELSSI",
    "ABDELKARIM TRABELSSI":       "ABDELHEDI TRABELSSI",
    "SABER KHARBESH":             "Saber Kherbech",
    "ALAEDINE KILENI":            "ALAEDDINE BEN KILANI",
    "ALI LTIFI":                  "ALI EL KOTLI",
    "HANI BEN KILANI":            "HANI BELKILANI",
    "HANI BEN KILANI":            "HANI BEN KILANI",
    "ADEL JAZI":                  "ADEL ALJAZI",
    "ABEDRAZEK BEY":              "ABDERRAZEG BEY",
    "ABDESLEM BEN SOUISSI":       "ABDESLEM BEN SOUISSI",
    "ABDELKADER MANNA":           "ABDELKADER ELMANAA",
    "AHMED ATTIA":                "AHMED ECHIKH",
    "AHMED IDRISSI":              "AHMED ELIDRISSI",
    "SALEM LEGRERI":              "SALEM LEGRERI",
    "FEDI AMAYMIA":               "FEDI LEGRERI",
    "HOUSSEM BRAYEK":             "HOUSSEM BRAYEK",
    "MED TAHER":                  "MOHAMED TAHRI",
    "MOUHAMED AOUINET":           "MOUHAMED ALI GHZALA",
    "MOUHAMED MESSII":            "MOUHAMED ALI GHZALA",
    "ROMDHAN SAAFI":              "ROMDHAN SAAFI",
    "BOUBAKER FILALI":            "BOUBAKER FILELI",
    "KAIS DHAOUI":                "KAIS EDHAOUI",
    "EZZEDINE GUESMI":            "EZZEDDIN ELGUESMI",
    "HEDI SLAMA":                 "HEDI SLEMA",
    "MOURAD HEMMEDI":             "MOURAD BEN SAID HAMMADI",
    "SAMI FERGENI":               "SAMI BEN AMOR FERJENI",
    "NEGI ZAAFOURI":              "NEJI ZAAFOURI",
    "NABIL EL MRIDH":             "NABIL EL MRIDH",
    "NAWFEL EL KORBI":            "NAWFEL EL  KORBI",
    "OSAMA KAAB":                 "OSAMA KAAB",
    # Lotfi
    "LOTFY HAJIJ":                "LOTFI ABDALLAH",
    # Slim Marzougui
    "SLIM MARZOUGUI":             "SLIM ELMARZOUGUI",
    # Imed
    "IMED AMDOUNI":               "IMEDDINE AMDOUNI",
    # Riadh
    "RIADH KOUKI":                "RIADH OBAY",
    "ISSAM KOUKI":                "ISSAM KOUKI",
}

def _get_concordance_key(nom_ref):
    """Trouve le nom Sotusfa correspondant via la table de concordance."""
    nom_up = str(nom_ref).strip().upper()
    # Recherche exacte dans la table
    for ref_key, sot_val in CONCORDANCE_NOMS.items():
        if ref_key.upper() == nom_up:
            return sot_val
    # Recherche normalisée (sans accents, espaces)
    import unicodedata, re as _re
    def _n(s):
        s = str(s).upper()
        s = ''.join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn')
        return _re.sub(r'\s+', ' ', _re.sub(r'[^A-Z0-9 ]', ' ', s)).strip()
    nom_clean = _n(nom_up)
    for ref_key, sot_val in CONCORDANCE_NOMS.items():
        if _n(ref_key) == nom_clean:
            return sot_val
    return None

def _fuzzy_match_clients(df_base, df_prev, col_prev):
    """Merge fuzzy SANS double comptage. Un prev_key → un seul base_key."""
    import unicodedata as _uni, re as _re2

    def _clean(s):
        s = str(s).upper().strip()
        s = _re2.sub(r"\bSOCIETE\b", "STE", s)
        s = "".join(c for c in _uni.normalize("NFD", s) if _uni.category(c) != "Mn")
        s = _re2.sub(r"\([^)]*\)", " ", s)
        s = _re2.sub(r"[^A-Z0-9 ]", " ", s)
        return _re2.sub(r"\s+", " ", s).strip()

    def _score(a, b):
        wa, wb = set(a.split()), set(b.split())
        if not wa or not wb: return 0.0
        inter = len(wa & wb); union = len(wa | wb)
        sj = inter / union if union else 0
        sh, lo = (wa, wb) if len(wa) <= len(wb) else (wb, wa)
        sc = sum(1 for w in sh if any(lw.startswith(w[:4]) for lw in lo)
                 and len(w) > 2) / max(len(sh), 1) * 0.85
        return max(sj, sc)

    THRESHOLD = 0.38
    df_base = df_base.copy()
    df_prev = df_prev.copy()
    df_base["_km"] = df_base["client"].apply(_clean)
    df_prev["_km"] = df_prev["client"].apply(_clean)

    # Agréger prévisions par _km (sécurité contre doublons)
    prev_agg = df_prev.groupby("_km", as_index=False)[col_prev].sum()
    prev_dict = dict(zip(prev_agg["_km"], prev_agg[col_prev]))

    # Merge exact
    result = df_base.merge(prev_agg, on="_km", how="left")

    # Fuzzy pour non-matchés (bijectif : chaque prev_key → 1 base_key max)
    unmatched_mask = result[col_prev].isna()
    if unmatched_mask.any():
        used_prev = set(result.loc[~unmatched_mask, "_km"].values)
        avail = {k: v for k, v in prev_dict.items() if k not in used_prev}
        assigned = {}
        for bk in result.loc[unmatched_mask, "_km"].unique():
            best_sc = 0; best_v = None
            for pk, pv in avail.items():
                sc = _score(bk, pk)
                if sc > best_sc and sc >= THRESHOLD:
                    best_sc = sc; best_v = pv
            if best_v is not None:
                assigned[bk] = best_v
        for bk, val in assigned.items():
            result.loc[result["_km"] == bk, col_prev] = val

    result = result.drop(columns=["_km"])
    return result, result[col_prev].notna().sum(), len(result)


def _export_excel_table(df, sheet_title="Data",
                        header_text="Export", color_hex="1F3864"):
    """Excel formaté attractif pour n'importe quel DataFrame."""
    from openpyxl import Workbook
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    import io as _io, numpy as _np
    wb = Workbook()
    ws = wb.active
    ws.title = sheet_title[:31]
    ws.sheet_view.showGridLines = False
    T = Side(style="thin", color="CCCCCC")
    BD = Border(left=T, right=T, top=T, bottom=T)
    CTR = Alignment(horizontal="center", vertical="center", wrap_text=True)
    LFT = Alignment(horizontal="left", vertical="center")
    def hf(h): return PatternFill("solid", start_color=h, end_color=h)
    def bf(bold=True, white=False, size=10, color=None):
        """bf local : supporte white=True (blanc) ET color hex explicite."""
        if white:
            final_color = "FFFFFF"
        elif color:
            final_color = str(color).lstrip("#")
        else:
            final_color = "000000"
        return Font(bold=bold, name="Calibri", size=size, color=final_color)
    nc = max(len(df.columns), 1)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=nc)
    ws["A1"] = header_text
    ws["A1"].font = bf(True, True, 12)
    ws["A1"].fill = hf(color_hex)
    ws["A1"].alignment = CTR
    ws.row_dimensions[1].height = 30
    for ci, col in enumerate(df.columns, 1):
        c = ws.cell(2, ci, value=str(col))
        c.font = bf(True, True, 10)
        c.fill = hf(color_hex)
        c.alignment = CTR
        c.border = BD
        ws.column_dimensions[get_column_letter(ci)].width = max(14, len(str(col)) + 4)
    ws.row_dimensions[2].height = 28
    ALERTE_COLORS = {
        "🔴":("FFCDD2","C0392B"), "🟡":("FFF9C4","D4AC0D"),
        "🟢":("E8F5E9","1E8449"),
    }
    # Détecter colonnes numériques
    _num_cols = {col: i+1 for i, col in enumerate(df.columns)
                 if str(df[col].dtype).startswith(("int","float"))}

    for ri, (_, row) in enumerate(df.iterrows()):
        r = ri + 3
        # Couleur ligne selon alerte si présente
        alerte_val = str(row.get("alerte","")) if "alerte" in df.columns else ""
        if "🔴" in alerte_val: row_bg = "FFEBEE"
        elif "🟡" in alerte_val: row_bg = "FFF9E6"
        elif "🟢" in alerte_val: row_bg = "E8F5E9"
        else: row_bg = "F0F5FF" if ri % 2 == 0 else "FFFFFF"

        for ci, val in enumerate(row, 1):
            col_name = df.columns[ci-1]
            if isinstance(val, float) and _np.isnan(val):
                val = ""
            c = ws.cell(r, ci, value=val)
            c.border = BD
            c.alignment = LFT if ci == 1 else CTR
            c.font = bf(ci == 1, size=9)

            # Couleur spéciale selon colonne
            if col_name == "alerte" and val:
                for emoji,(bg2,fg2) in ALERTE_COLORS.items():
                    if emoji in str(val):
                        c.fill = hf(bg2)
                        c.font = bf(True, size=9, color=fg2)
                        break
                else:
                    c.fill = hf(row_bg)
            elif col_name in ("ecart_tonnage","solde_final") and val != "" and val is not None:
                try:
                    fv = float(val)
                    c.fill = hf("E8F5E9") if fv >= 0 else hf("FFEBEE")
                    c.font = bf(True, size=9, color="1E8449" if fv >= 0 else "C0392B")
                    c.number_format = "+#,##0;-#,##0;0"
                except (TypeError, ValueError):
                    c.fill = hf(row_bg)
            elif col_name == "taux_prise" and isinstance(val,(int,float)) and val==val:
                v = float(val)
                c.fill = hf("E8F5E9" if v>=90 else ("FFF9E6" if v>=85 else "FFEBEE"))
                c.number_format = "0.0"
            elif col_name == "affectation_caisse":
                c.fill = hf("FBE9E7") if "1ère" in str(val) else hf("E8F5E9")
            else:
                c.fill = hf(row_bg)

            if val != "" and val is not None and col_name not in ("ecart_tonnage","solde_final","taux_prise"):
                try:
                    fv2 = float(val)
                    if fv2 == fv2:  # pas NaN
                        c.number_format = "#,##0" if abs(fv2) >= 100 else "0.0"
                except (TypeError, ValueError):
                    pass
    num_ci = [i + 1 for i, col in enumerate(df.columns)
              if str(df[col].dtype).startswith(("int", "float"))]
    if num_ci:
        tr = len(df) + 3
        for ci in range(1, nc + 1):
            c = ws.cell(tr, ci)
            c.fill = hf(color_hex)
            c.font = bf(True, True)
            c.border = BD
            c.alignment = CTR
        ws.cell(tr, 1).value = "TOTAL"
        ws.cell(tr, 1).alignment = LFT
        for ci in num_ci:
            col_l = get_column_letter(ci)
            ws.cell(tr, ci).value = f"=SUM({col_l}3:{col_l}{tr-1})"
            ws.cell(tr, ci).number_format = "#,##0"
        ws.row_dimensions[tr].height = 22
    ws.freeze_panes = "A3"
    buf = _io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()

def render_agroeco_tab(sb=None, CURRENT_ROLE="directeur", CURRENT_NAME=""):

    st.markdown("""
<div style='background:#0a1a0a;border:1px solid #1E8449;border-radius:12px;
padding:16px 20px;margin-bottom:18px'>
  <div style='font-size:1.05rem;font-weight:700;color:#f0f6fc;margin-bottom:6px'>
    📊 Dashboard Agroéconomique — Tomate 2026
  </div>
  <div style='font-size:.82rem;color:#8b949e;line-height:1.8'>
    Clé de jointure : <b style='color:#FFD700'>client + centre</b> (obligatoires dans tous les fichiers) &nbsp;·&nbsp;
    Caisses vides : <b style='color:#FF9800'>date début RÉCOLTE</b> (fichier rectifié Supabase) &nbsp;·&nbsp;
    <b style='color:#ef5350'>Tonnage recouvrement</b> = (Charges + Consignes + MO 50DT/T) ÷ Prix vente
  </div>
</div>""", unsafe_allow_html=True)

    # Session state
    KEYS = ["abo_bourak","abo_royal","abo_sotusfa_raw","abo_sotusfa_pivot",
            "abo_quantite","abo_prev_dec","abo_prev_mai","abo_prev_juin",
            "abo_dates_recolte","abo_merged","abo_params","abo_errors"]
    for k in KEYS:
        if k not in st.session_state:
            st.session_state[k] = None

    t0,t1,t2,t3,t4,t5,t6,t7 = st.tabs([
        "⚙️ Paramètres & Import",
        "📋 Par Agriculteur",
        "👤 Par Ingénieur / Centre",
        "🗺️ Par Région",
        "🍅 Par Variété",
        "💊 Par Famille Intrant",
        "📈 Prévisions vs Réalisé",
        "🏆 Analyse Efficacité Pro",
    ])

    # ══ TAB 0 — PARAMÈTRES ET IMPORT ══════════════════════
    with t0:
        # ── Contrôle d'accès : upload réservé au directeur ───
        _is_admin = (CURRENT_ROLE in ("directeur", "admin"))

        if not _is_admin:
            # Utilisateur non-admin : affiche statut session partagée
            _shared_data = load_session_from_supabase(sb, "SHARED_2026") if sb else None
            if st.session_state.get("abo_merged") is not None:
                st.success("✅ Données chargées automatiquement depuis la session partagée.")
                st.info("🔒 L'upload des fichiers est réservé à l'administrateur. "
                        "Les données se mettent à jour automatiquement.")
                if st.button("🔄 Rafraîchir les données"):
                    if sb:
                        _sd = load_session_from_supabase(sb, "SHARED_2026")
                        if _sd and _sd.get("merged") is not None:
                            st.session_state["abo_merged"]        = _sd["merged"]
                            st.session_state["abo_bourak"]        = _sd.get("bourak")
                            st.session_state["abo_royal"]         = _sd.get("royal")
                            st.session_state["abo_sotusfa_raw"]   = _sd.get("sotusfa_raw")
                            st.session_state["abo_sotusfa_pivot"] = _sd.get("sotusfa_pivot")
                            st.session_state["abo_quantite"]      = _sd.get("quantite")
                            st.session_state["abo_prev_mai"]      = _sd.get("prev_mai")
                            st.cache_data.clear()
                            st.rerun()
                        else:
                            st.warning("Aucune session partagée trouvée.")
            else:
                st.warning("⏳ En attente des données — l'administrateur doit importer et sauvegarder les fichiers.")
            st.stop()  # Ne pas afficher le reste du tab pour les non-admins

        # ── Paramètres (admin seulement) ─────────────────────
        st.markdown("### ⚙️ Paramètres de calcul")
        pc1,pc2,pc3 = st.columns(3)
        with pc1:
            st.markdown("**💰 Prix vente global (DT/tonne)**")
            st.caption("Utilisé si absent du tableau quantité")
            prix_global = st.number_input("Prix vente DT/T",0.0,1000.0,240.0,10.0,key="px_g")
        with pc2:
            st.markdown("**🔲 Consigne plateaux (DT/plateau)**")
            p228pvc  = st.number_input("Pltx 228 PVC", 0.0,50.0,2.5,0.1,key="p1")
            p228poly = st.number_input("Pltx 228 POLY",0.0,50.0,2.0,0.1,key="p2")
            p160pvc  = st.number_input("Pltx 160 PVC", 0.0,50.0,2.0,0.1,key="p3")
            p160poly = st.number_input("Pltx 160 POLY",0.0,50.0,1.8,0.1,key="p4")
        with pc3:
            st.markdown("**📦 Caisses vides — MO récolte**")
            mo_tonne = st.number_input("MO récolte (DT/T)",0.0,200.0,50.0,5.0,key="mo")
            st.caption("Condition caisses : date début RÉCOLTE < 10 juil. → 1ère affectation")

        # ── Caisses vides PAR USINE ───────────────────────────
        st.markdown("---")
        st.markdown("#### 📦 Caisses vides — Paramètres par usine")
        st.caption("1ère affectation (début récolte < 10 juillet) = caisses facturées | 2ème = 0 DT")

        caisses_par_usine = {}
        _saved_caisses = (st.session_state.get("abo_params") or {}).get("caisses_par_usine", {})
        usine_cols = st.columns(5)
        usine_names = ["SICAM","TUCAL","COMOCAP","ABIDA","ELFALLEH"]
        usine_colors = {"SICAM":"#F5A623","TUCAL":"#8B5CF6","COMOCAP":"#3B82F6",
                        "ABIDA":"#FF6B9D","ELFALLEH":"#00E5A0"}

        for ci2, usine in enumerate(usine_names):
            dft = CAISSES_USINE_DEFAULTS.get(usine, {"nb_ha":80,"prix":3.0,"type":"Caisse 25kg","cap_kg":25})
            saved_u = _saved_caisses.get(usine, dft)
            uc = usine_colors.get(usine,"#888")
            with usine_cols[ci2]:
                st.markdown(f"<div style='background:#1a2332;border-radius:8px;padding:8px;"
                            f"border-top:3px solid {uc};margin-bottom:4px'>"
                            f"<b style='color:{uc};font-size:12px'>{usine}</b><br>"
                            f"<span style='font-size:10px;color:#aaa'>{dft['type']}</span>"
                            f"</div>", unsafe_allow_html=True)
                nb_ha = st.number_input(f"Nb caisses/ha",
                    min_value=0.0, max_value=300.0,
                    value=float(saved_u.get("nb_ha", dft["nb_ha"])),
                    step=5.0, key=f"nb_c_{usine}")
                prix_c = st.number_input(f"Prix/caisse (DT)",
                    min_value=0.0, max_value=20.0,
                    value=float(saved_u.get("prix", dft["prix"])),
                    step=0.25, key=f"px_c_{usine}")
                cout_ha = round(nb_ha * prix_c, 2)
                st.caption(f"→ **{cout_ha:.1f} DT/ha** (1ère affectation)")
                caisses_par_usine[usine] = {"nb_ha": nb_ha, "prix": prix_c,
                                            "type": dft["type"], "cap_kg": dft["cap_kg"]}

        params = {
            "prix_vente_global": prix_global,
            "prix_consigne": {
                "Pltx 228 PVC":p228pvc,"Pltx 228 POLY":p228poly,
                "Pltx 160 PVC":p160pvc,"Pltx 160 POLY":p160poly,
            },
            "caisses_par_usine": caisses_par_usine,
            # Rétrocompat : valeurs globales = moyenne pondérée SICAM (usine principale)
            "prix_caisse":   caisses_par_usine.get("SICAM",{}).get("prix", 3.0),
            "nb_caisses_ha": caisses_par_usine.get("SICAM",{}).get("nb_ha", 80.0),
            "mo_tonne":      mo_tonne,
        }
        st.session_state["abo_params"] = params

        st.divider()
        st.markdown("### 📥 Import fichiers")
        st.markdown("""<div style='background:#161b22;border:1px solid #FFD700;
border-radius:8px;padding:10px 14px;margin-bottom:14px;font-size:.85rem'>
⭐ <b style='color:#FFD700'>Colonnes OBLIGATOIRES dans tous les fichiers :</b>
&nbsp;<code>centre</code> &nbsp;+&nbsp; <code>client</code>
</div>""", unsafe_allow_html=True)

        fi1,fi2 = st.columns(2)
        fi3,fi4 = st.columns(2)

        def _upload_block(col, icon, name, color, desc, key, parse_fn, extra_args=()):
            with col:
                st.markdown(f"""<div style='background:#111;border:1px solid #{color};
border-radius:8px;padding:10px 14px;margin-bottom:8px'>
<b style='color:#{color}'>{icon} {name}</b><br>
<span style='font-size:.78rem;color:#aaa'>{desc}</span></div>""",
                    unsafe_allow_html=True)
                f = st.file_uploader(name,type=["xlsx","xls"],
                                     key=key,label_visibility="collapsed")
                if f:
                    try:
                        result = parse_fn(f, *extra_args)
                        if isinstance(result, tuple):
                            if len(result) == 2:
                                df_res, msg = result
                            else:
                                df_res, pivot, msg = result
                        else:
                            df_res, msg = result, ""
                        if msg:
                            st.error(msg)
                            return None
                        return f, df_res, pivot if len(result)==3 else None
                    except Exception as e:
                        st.error(f"Erreur {name}: {e}")
                return None

        # BOURAK
        with fi1:
            st.markdown(f"""<div style='background:#111;border:1px solid #FF9800;
border-radius:8px;padding:10px 14px;margin-bottom:8px'>
<b style='color:#FF9800'>🚛 BOURAK</b> — Financement<br>
<span style='font-size:.78rem;color:#aaa'>Obligatoire : <b>client · centre</b><br>
Attendu : responsable · ingenieur · region · hectares · avance · report</span></div>""",
                unsafe_allow_html=True)
            f_b = st.file_uploader("Bourak",type=["xlsx","xls"],
                                    key="up_b",label_visibility="collapsed")
            if f_b:
                res = parse_bourak(f_b)
                if isinstance(res, tuple): df_b, msg = res
                else: df_b, msg = res, ""
                if msg: st.error(msg)
                else:
                    st.session_state["abo_bourak"] = df_b
                    tot_av = df_b["avance"].sum() if "avance" in df_b.columns else 0
                    st.success(f"✅ {len(df_b)} lignes · {tot_av:,.0f} DT avances")

        # ROYAL
        with fi2:
            st.markdown(f"""<div style='background:#111;border:1px solid #9C27B0;
border-radius:8px;padding:10px 14px;margin-bottom:8px'>
<b style='color:#9C27B0'>🌱 ROYAL</b> — Plants<br>
<span style='font-size:.78rem;color:#aaa'>Obligatoire : <b>client · centre</b><br>
Attendu : zone · variete · qte_livree · valeur · <b>date_debut_livraison</b> · date_fin</span></div>""",
                unsafe_allow_html=True)
            f_r = st.file_uploader("Royal",type=["xlsx","xls"],
                                    key="up_r",label_visibility="collapsed")
            if f_r:
                df_r, msg = parse_royal(f_r)
                if msg: st.error(msg)
                else:
                    st.session_state["abo_royal"] = df_r
                    st.success(f"✅ {len(df_r)} lignes")

        # SOTUSFA
        with fi3:
            st.markdown(f"""<div style='background:#111;border:1px solid #4CAF50;
border-radius:8px;padding:10px 14px;margin-bottom:8px'>
<b style='color:#4CAF50'>🌿 SOTUSFA</b> — Intrants<br>
<span style='font-size:.78rem;color:#aaa'>Obligatoire : <b>client · centre</b><br>
Attendu : famille · article · qte · valeur (DAP / fumure / fumier / pest.)</span></div>""",
                unsafe_allow_html=True)
            f_s = st.file_uploader("Sotusfa",type=["xlsx","xls"],
                                    key="up_s",label_visibility="collapsed")
            if f_s:
                df_s_raw, df_s_piv, msg = parse_sotusfa(f_s)
                if msg: st.error(msg)
                else:
                    st.session_state["abo_sotusfa_raw"]   = df_s_raw
                    st.session_state["abo_sotusfa_pivot"] = df_s_piv
                    tot = df_s_raw["valeur"].sum() if "valeur" in df_s_raw.columns else 0
                    st.success(f"✅ {len(df_s_raw)} lignes · {tot:,.0f} DT")

        # QUANTITÉ
        with fi4:
            st.markdown(f"""<div style='background:#111;border:1px solid #2196F3;
border-radius:8px;padding:10px 14px;margin-bottom:8px'>
<b style='color:#2196F3'>📊 QUANTITÉ</b> — Actif/Extra<br>
<span style='font-size:.78rem;color:#aaa'>Obligatoire : <b>client · centre</b><br>
Attendu : qte_livree · qte_actif · qte_extra · tonnage_livre · prix_vente</span></div>""",
                unsafe_allow_html=True)
            f_q = st.file_uploader("Quantité",type=["xlsx","xls"],
                                    key="up_q",label_visibility="collapsed")
            if f_q:
                df_q, msg = parse_quantite(f_q)
                if msg: st.error(msg)
                else:
                    st.session_state["abo_quantite"] = df_q
                    st.success(f"✅ {len(df_q)} agriculteurs")

        # ── Prévisions ──────────────────────────────────────
        st.divider()
        st.markdown("### 📅 Prévisions tonnage")
        pv1,pv2,pv3 = st.columns(3)
        with pv1:
            f_dec = st.file_uploader("📋 Prévision Décembre",
                                      type=["xlsx","xls"],key="up_dec")
            if f_dec:
                df_d, msg = parse_prevision(f_dec, "prevision_dec")
                if msg:
                    st.warning(f"⚠️ Déc (non bloquant): {msg}")
                    # Essayer quand même avec ce qu'on a
                    if df_d is not None and not df_d.empty:
                        st.session_state["abo_prev_dec"] = df_d
                        st.success(f"✅ Déc chargé malgré avertissement: {df_d['prevision_dec'].sum():,.0f} T")
                elif df_d is not None and not df_d.empty:
                    st.session_state["abo_prev_dec"] = df_d
                    st.success(f"✅ Déc: {df_d['prevision_dec'].sum():,.0f} T")
        with pv2:
            f_mai = st.file_uploader("📋 Prévision Mai",
                                      type=["xlsx","xls"],key="up_mai")
            if f_mai:
                df_m, msg = parse_prevision(f_mai, "prevision_mai")
                if msg:
                    st.warning(f"⚠️ Mai (non bloquant): {msg}")
                    if df_m is not None and not df_m.empty:
                        st.session_state["abo_prev_mai"] = df_m
                        st.success(f"✅ Mai chargé malgré avertissement: {df_m['prevision_mai'].sum():,.0f} T")
                elif df_m is not None and not df_m.empty:
                    st.session_state["abo_prev_mai"] = df_m
                    st.success(f"✅ Mai: {df_m['prevision_mai'].sum():,.0f} T")
        with pv3:
            st.markdown("**☁️ Juin — Supabase (fichier rectifié)**")
            c1,c2 = st.columns(2)
            with c1:
                if st.button("🔄 Charger Juin", use_container_width=True):
                    df_j = load_prevision_juin(sb)
                    if not df_j.empty:
                        st.session_state["abo_prev_juin"] = df_j
                        st.success(f"✅ {df_j['prevision_juin'].sum():,.0f} T")
                    else:
                        st.warning("Aucune donnée Juin")
            with c2:
                if st.button("🔄 Dates récolte", use_container_width=True):
                    df_dr = load_date_debut_recolte(sb)
                    if not df_dr.empty:
                        st.session_state["abo_dates_recolte"] = df_dr
                        n1 = (df_dr["affectation_caisse"].str.startswith("1ère")).sum()
                        st.success(f"✅ {n1} agriculteurs 1ère affectation")
                    else:
                        st.warning("Aucune date récolte")

        # Statut caisses vides
        dr = st.session_state.get("abo_dates_recolte")
        if dr is not None and not dr.empty:
            n1 = (dr["affectation_caisse"].str.startswith("1ère")).sum()
            n2 = len(dr) - n1
            st.info(f"📦 Caisses vides — **1ère affectation** (< 10 juil.) : {n1} agriculteurs · "
                    f"**2ème** (≥ 10 juil.) : {n2} agriculteurs")
        else:
            st.warning("⚠️ Dates de récolte non chargées → tous les agriculteurs "
                       "seront mis en 2ème affectation (sans caisses vides). "
                       "Cliquez '🔄 Dates récolte' ci-dessus.")

        # ── Fusionner ────────────────────────────────────────
        st.divider()
        if st.button("🔗 Fusionner et calculer",
                     type="primary",use_container_width=True):
            with st.spinner("Calcul en cours…"):
                df_merged = merge_and_calculate(
                    st.session_state.get("abo_bourak"),
                    st.session_state.get("abo_royal"),
                    st.session_state.get("abo_sotusfa_raw"),
                    st.session_state.get("abo_sotusfa_pivot"),
                    st.session_state.get("abo_quantite"),
                    st.session_state.get("abo_prev_dec"),
                    st.session_state.get("abo_prev_mai"),
                    st.session_state.get("abo_prev_juin"),
                    st.session_state.get("abo_dates_recolte"),
                    st.session_state["abo_params"],
                )
            if df_merged.empty:
                st.error("❌ Aucune donnée fusionnée — vérifiez les fichiers.")
            else:
                st.session_state["abo_merged"] = df_merged
                n_r = (df_merged["alerte"].str.contains("🔴")).sum()
                n_y = (df_merged["alerte"].str.contains("🟡")).sum()
                n_g = (df_merged["alerte"].str.contains("🟢")).sum()
                st.success(
                    f"✅ {len(df_merged)} agriculteurs · "
                    f"🔴 {n_r} critiques · 🟡 {n_y} attention · 🟢 {n_g} OK")
                xl = export_excel(df_merged, st.session_state.get("abo_sotusfa_raw"))
                st.download_button(
                    "📥 Télécharger Excel complet (4 feuilles)",
                    data=xl,
                    file_name="dashboard_agroeco_2026.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",use_container_width=True)

    # ── Données fusionnées ─────────────────────────────────
    df = st.session_state.get("abo_merged")

    def _no_data():
        st.info("📥 Importez les fichiers dans l'onglet **⚙️ Paramètres & Import** "
                "puis cliquez **Fusionner**.")

    # ══ TAB 1 — PAR AGRICULTEUR ════════════════════════════
    with t1:
        if df is None or df.empty:
            _no_data()
        else:
            kc = st.columns(6)
            kc[0].markdown(_metric("Agriculteurs",len(df)), unsafe_allow_html=True)
            kc[1].markdown(_metric("Charge à Recouvrir",f"{df['charge_totale'].sum():,.0f} DT",color="#FF9800"),unsafe_allow_html=True)
            kc[2].markdown(_metric("Recouvrement",f"{df['tonnage_recouvrement'].sum():,.1f} T",color="#ef5350"),unsafe_allow_html=True)

            # Tonnage réalisé : depuis Quantité ou prévision
            _ton_reel = df["tonnage_livre"].fillna(0).sum() if "tonnage_livre" in df.columns else 0
            _has_quantite = st.session_state.get("abo_quantite") is not None
            _ton_label = f"{_ton_reel:,.1f} T" if _has_quantite else "Non importé"
            _ton_color = "#4CAF50" if _has_quantite and _ton_reel > 0 else "#888"
            kc[3].markdown(_metric("Livré réel", _ton_label, color=_ton_color),
                           unsafe_allow_html=True)

            # Prévision Mai : total brut du fichier (pas juste les matchés)
            _prev_mai_brut = 0
            _df_prev_mai = st.session_state.get("abo_prev_mai")
            if _df_prev_mai is not None and "prevision_mai" in _df_prev_mai.columns:
                _prev_mai_brut = _df_prev_mai["prevision_mai"].sum()
            _prev_mai_merged = df["prevision_mai"].fillna(0).sum() if "prevision_mai" in df.columns else 0

            if _prev_mai_brut > 0:
                _pct_match = round(_prev_mai_merged / _prev_mai_brut * 100) if _prev_mai_brut > 0 else 0
                kc[4].markdown(_metric("Prévision Mai",
                    f"{_prev_mai_brut:,.0f} T",
                    color="#42A5F5",
                    delta=_prev_mai_merged - _prev_mai_brut,
                    delta_label=f"T matchés ({_pct_match}%)"),
                    unsafe_allow_html=True)
            else:
                n_crit=(df["alerte"].str.contains("🔴")).sum()
                kc[4].markdown(_metric("⚠️ Critiques",n_crit,color="#ef5350"),unsafe_allow_html=True)

            kc[5].markdown(_metric("Solde global",f"{df['solde_final'].sum():+,.0f} DT",
                color="#4CAF50" if df["solde_final"].sum()>=0 else "#ef5350"),unsafe_allow_html=True)

            # Avertissement si données fictives
            if not _has_quantite:
                st.warning("⚠️ **Tonnage réalisé = 0** — Le fichier **Tableau Quantité** n'est pas encore importé. "
                           "Les calculs de recouvrement et solde sont basés sur les prévisions uniquement.")
            if _prev_mai_brut > 0 and _prev_mai_merged < _prev_mai_brut * 0.5:
                st.info(f"ℹ️ **Prévision Mai** : {_prev_mai_brut:,.0f} T dans le fichier, "
                        f"mais seulement **{_prev_mai_merged:,.0f} T matchés** ({round(_prev_mai_merged/_prev_mai_brut*100)}%) "
                        f"car certains noms d'agriculteurs diffèrent entre les fichiers. "
                        f"Le total affiché dans les tableaux = valeurs matchées uniquement.")

            fc1,fc2,fc3,fc4 = st.columns(4)
            alerte_f = fc1.selectbox("Alerte",["Toutes","🔴","🟡","🟢"],key="t1a")
            comm_f   = fc2.selectbox("Commercial",
                ["Tous"]+sorted(df["commercial"].dropna().unique().tolist())
                if "commercial" in df.columns else ["Tous"],key="t1c")
            ing_f    = fc3.selectbox("Ingénieur",
                ["Tous"]+sorted(df["ingenieur"].dropna().unique().tolist())
                if "ingenieur" in df.columns else ["Tous"],key="t1i")
            ctr_f    = fc4.selectbox("Centre",
                ["Tous"]+sorted(df["centre"].dropna().unique().tolist())
                if "centre" in df.columns else ["Tous"],key="t1ct")

            df_f = df.copy()
            if alerte_f != "Toutes": df_f = df_f[df_f["alerte"].str.contains(alerte_f)]
            if comm_f   != "Tous" and "commercial" in df_f.columns: df_f = df_f[df_f["commercial"]==comm_f]
            if ing_f    != "Tous" and "ingenieur"  in df_f.columns: df_f = df_f[df_f["ingenieur"]== ing_f]
            if ctr_f    != "Tous" and "centre"     in df_f.columns: df_f = df_f[df_f["centre"]   == ctr_f]

            # Graphique écart tonnage
            if "ecart_tonnage" in df_f.columns and "agriculteur" in df_f.columns:
                df_c = df_f.dropna(subset=["ecart_tonnage"]).sort_values("ecart_tonnage")
                if not df_c.empty:
                    fig = go.Figure()
                    fig.add_trace(go.Bar(
                        y=df_c["agriculteur"], x=df_c["ecart_tonnage"],
                        orientation="h",
                        marker_color=["#ef5350" if v<0 else "#4CAF50"
                                      for v in df_c["ecart_tonnage"]],
                        text=df_c["ecart_tonnage"].apply(lambda v:f"{v:+.1f} T"),
                        textposition="outside"))
                    fig.add_vline(x=0,line_color="#888",line_width=1.5,
                                  line_dash="dash")
                    fig.update_layout(
                        title="Écart = Tonnage livré − Tonnage recouvrement",
                        template="plotly_dark",paper_bgcolor="#161b22",
                        plot_bgcolor="#0d1117",
                        height=max(350,len(df_c)*26+80),
                        margin=dict(l=220,r=80,t=40,b=30))
                    st.plotly_chart(fig, use_container_width=True)

            # ── Section Caisses Vides par usine ─────────────────
            if "affectation_caisse" in df_f.columns:
                st.markdown("#### 📦 Détail caisses vides par usine")
                # Récap par usine des consignes caisses
                caisse_cfg = st.session_state.get("abo_params",{}).get("caisses_par_usine",{})
                if caisse_cfg:
                    cv_cols = st.columns(len(caisse_cfg))
                    usine_colors_disp = {"SICAM":"#F5A623","TUCAL":"#8B5CF6",
                                         "COMOCAP":"#3B82F6","ABIDA":"#FF6B9D","ELFALLEH":"#00E5A0"}
                    for ci_u, (usine_u, cfg_u) in enumerate(caisse_cfg.items()):
                        # Filtrer agriculteurs de cette usine en 1ère affectation
                        _mask_1 = (df_f.get("affectation_caisse","").str.startswith("1ère")
                                   if hasattr(df_f.get("affectation_caisse",""),"str") else pd.Series([False]*len(df_f)))
                        _usine_mask = df_f.get("usine", pd.Series([""] * len(df_f))).astype(str).str.upper().str.contains(usine_u.upper(), na=False)
                        _agri_1ere = df_f[_mask_1 & _usine_mask] if "usine" in df_f.columns else df_f[_mask_1]
                        n_1ere = len(_agri_1ere)
                        total_cv = _agri_1ere["consigne_caisse"].sum() if "consigne_caisse" in _agri_1ere.columns else 0
                        cout_ha = round(cfg_u["nb_ha"] * cfg_u["prix"], 1)
                        uc2 = usine_colors_disp.get(usine_u,"#888")
                        with cv_cols[ci_u]:
                            st.markdown(f"""<div style='background:#1a2332;border-radius:10px;
padding:10px;text-align:center;border-top:3px solid {uc2}'>
<div style='font-size:13px;font-weight:bold;color:{uc2}'>{usine_u}</div>
<div style='font-size:11px;color:#aaa'>{cfg_u['type']}</div>
<div style='font-size:11px;color:#ccc;margin:4px 0'>
{cfg_u['nb_ha']:.0f} caisses/ha × {cfg_u['prix']:.2f} DT = <b style='color:#fff'>{cout_ha} DT/ha</b>
</div>
<div style='font-size:11px;color:#FFD700'>{n_1ere} agri. 1ère affect.</div>
<div style='font-size:13px;font-weight:bold;color:{"#FF7043" if total_cv>0 else "#4CAF50"}'>
{total_cv:,.0f} DT total</div>
</div>""", unsafe_allow_html=True)

                # Total global caisses
                if "consigne_caisse" in df_f.columns:
                    tot_cv = df_f["consigne_caisse"].sum()
                    n_1e = df_f["affectation_caisse"].str.startswith("1ère").sum() if "affectation_caisse" in df_f.columns else 0
                    n_2e = len(df_f) - n_1e
                    cc1,cc2,cc3 = st.columns(3)
                    cc1.markdown(_metric("Total consigne caisses",
                        f"{tot_cv:,.0f} DT", color="#FF7043"), unsafe_allow_html=True)
                    cc2.markdown(_metric("1ère affectation (< 10 juil.)",
                        f"{n_1e} agriculteurs", color="#FF9800"), unsafe_allow_html=True)
                    cc3.markdown(_metric("2ème affectation (≥ 10 juil.)",
                        f"{n_2e} agriculteurs — 0 DT", color="#4CAF50"), unsafe_allow_html=True)
                st.markdown("---")

            VIEW = [c for c in ["agriculteur","commercial","ingenieur","centre","variete",
                                  "hectares","affectation_caisse","detail_caisse","taux_prise",
                                  "report","charge_a_recouvrir","consigne_caisse","consigne_plateau",
                                  "mo_recolte","tonnage_recouvrement","tonnage_livre",
                                  "ecart_tonnage","valeur_livree","solde_final",
                                  "alerte"] if c in df_f.columns]
            st.dataframe(df_f[VIEW].round(1),
                use_container_width=True,hide_index=True,height=400,
                column_config={
                    "taux_prise":st.column_config.ProgressColumn(
                        "Taux prise %",min_value=0,max_value=100,format="%.1f%%"),
                    "ecart_tonnage":st.column_config.NumberColumn("Écart (T)",format="%+.1f"),
                    "solde_final":st.column_config.NumberColumn("Solde (DT)",format="%+,.0f"),
                })
            # Exports multiples tab1
            _dl1, _dl2 = st.columns(2)
            with _dl1:
                xl2 = export_excel(df_f, st.session_state.get("abo_sotusfa_raw"))
                st.download_button("📥 Excel complet (4 feuilles)",data=xl2,
                    file_name="agroeco_vue.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True, type="primary")
            with _dl2:
                _view_df = df_f[[c for c in VIEW if c in df_f.columns]].round(1)
                st.download_button("📥 Excel — Vue actuelle",
                    data=_export_excel_table(
                        _view_df, "Par Agriculteur",
                        "Tableau Agroéconomique par Agriculteur — Campagne 2026",
                        "1F3864"),
                    file_name="agroeco_par_agriculteur.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True)

    # ══ TAB 2 — PAR INGÉNIEUR / CENTRE ════════════════════
    with t2:
        if df is None or df.empty: _no_data()
        else:
            grp_col = st.radio("Regrouper par", ["Ingénieur","Centre","Ingénieur × Centre"],
                                horizontal=True, key="t2_grp")
            if grp_col == "Ingénieur":
                gc = ["ingenieur"] if "ingenieur" in df.columns else ["centre"]
            elif grp_col == "Centre":
                gc = ["centre"] if "centre" in df.columns else ["ingenieur"]
            else:
                gc = [c for c in ["ingenieur","centre"] if c in df.columns]

            if gc:
                g2 = df.groupby(gc).agg(
                    Agriculteurs       = ("agriculteur","count"),
                    Hectares           = ("hectares","sum"),
                    Plants_actifs      = ("qte_actif","sum"),
                    Taux_prise_moy     = ("taux_prise","mean"),
                    Charge_totale      = ("charge_totale","sum"),
                    Recouvrement_T     = ("tonnage_recouvrement","sum"),
                    Livre_T            = ("tonnage_livre","sum"),
                    Ecart_T            = ("ecart_tonnage","sum"),
                    Solde_DT           = ("solde_final","sum"),
                    Alertes_rouges     = ("alerte",lambda x:(x.str.contains("🔴")).sum()),
                ).reset_index().round(1)

                fig2 = go.Figure()
                x_col = gc[-1] if gc else "centre"
                fig2.add_trace(go.Bar(name="Recouvrement (T)",
                    x=g2[x_col],y=g2["Recouvrement_T"],marker_color="#ef5350"))
                fig2.add_trace(go.Bar(name="Livré réel (T)",
                    x=g2[x_col],y=g2["Livre_T"],marker_color="#4CAF50"))
                fig2.update_layout(barmode="group",template="plotly_dark",
                    paper_bgcolor="#161b22",plot_bgcolor="#0d1117",
                    height=340,title=f"Recouvrement vs Livré par {grp_col}")
                st.plotly_chart(fig2,use_container_width=True)
                st.dataframe(g2,use_container_width=True,hide_index=True)
                st.download_button(
                    "📥 Excel — Par Ingénieur/Centre",
                    data=_export_excel_table(g2,
                        "Par Ingenieur Centre",
                        f"Synthèse par {grp_col} — Campagne 2026",
                        "0B4F6C"),
                    file_name="agroeco_par_ingenieur.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True)

    # ══ TAB 3 — PAR RÉGION ════════════════════════════════
    with t3:
        if df is None or df.empty: _no_data()
        elif "region" not in df.columns:
            st.warning("Colonne 'region' absente — vérifiez le fichier Bourak.")
        else:
            rg = df.groupby("region").agg(
                Agriculteurs    = ("agriculteur","count"),
                Hectares        = ("hectares","sum"),
                Cout_ha_moy     = ("cout_ha","mean"),
                Rendement_moy   = ("tonnage_ha_realise","mean"),
                Taux_prise_moy  = ("taux_prise","mean"),
                Recouvrement_ha = ("recouvrement_ha","mean"),
                Tonnage_total   = ("tonnage_livre","sum"),
                Alertes_rouges  = ("alerte",lambda x:(x.str.contains("🔴")).sum()),
            ).reset_index().round(1)

            fig3 = px.bar(rg,x="region",y="Rendement_moy",
                color="Rendement_moy",
                color_continuous_scale=["#ef5350","#FF9800","#4CAF50"],
                template="plotly_dark",text_auto=".1f",
                title="Rendement moyen (T/ha) par région")
            fig3.update_layout(paper_bgcolor="#161b22",plot_bgcolor="#0d1117",height=340)
            st.plotly_chart(fig3,use_container_width=True)
            st.dataframe(rg,use_container_width=True,hide_index=True)
            st.download_button(
                "📥 Excel — Par Région",
                data=_export_excel_table(rg,
                    "Par Region",
                    "Performance par Région — Campagne 2026",
                    "1A5C2A"),
                file_name="agroeco_par_region.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True)

    # ══ TAB 4 — PAR VARIÉTÉ ═══════════════════════════════
    with t4:
        if df is None or df.empty: _no_data()
        elif "variete" not in df.columns:
            st.warning("Colonne 'variete' absente — vérifiez le fichier Royal.")
        else:
            vg = df.dropna(subset=["variete"]).groupby("variete").agg(
                Agriculteurs  = ("agriculteur","count"),
                Hectares      = ("hectares","sum"),
                Densite_moy   = ("densite_ha","mean"),
                Rendement_moy = ("tonnage_ha_realise","mean"),
                Taux_prise    = ("taux_prise","mean"),
                Cout_ha_moy   = ("cout_ha","mean"),
                Tonnage_total = ("tonnage_livre","sum"),
            ).reset_index().sort_values("Rendement_moy",ascending=False).round(1)

            fig4 = px.bar(vg,x="variete",y="Rendement_moy",
                color="Rendement_moy",
                color_continuous_scale=["#ef5350","#FF9800","#4CAF50"],
                template="plotly_dark",text_auto=".1f",
                title="Rendement moyen (T/ha) par variété")
            fig4.update_layout(paper_bgcolor="#161b22",plot_bgcolor="#0d1117",height=340)
            st.plotly_chart(fig4,use_container_width=True)
            st.dataframe(vg,use_container_width=True,hide_index=True)
            st.download_button(
                "📥 Excel — Par Variété",
                data=_export_excel_table(vg,
                    "Par Variete",
                    "Performance par Variété — Campagne 2026",
                    "375623"),
                file_name="agroeco_par_variete.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True)

    # ══ TAB 5 — PAR FAMILLE INTRANT ═══════════════════════
    with t5:
        ds = st.session_state.get("abo_sotusfa_raw")
        if ds is None or ds.empty:
            _no_data()
            st.caption("Importez le fichier Sotusfa dans l'onglet ⚙️.")
        else:
            # Recalculer famille_norm si absent
            if "famille_norm" not in ds.columns and "famille" in ds.columns:
                ds = ds.copy()
                ds["famille_norm"] = ds["famille"].astype(str).str.strip().str.lower()\
                                     .map(FAM_NORM_MAP).fillna("Autre")
            if "famille_norm" not in ds.columns:
                st.warning("Colonne famille absente.")
            else:
                _cl = "client" if "client" in ds.columns else (
                      "agriculteur" if "agriculteur" in ds.columns else None)
                _agg = {"Valeur_DT":("valeur","sum")}
                if _cl: _agg["Nb_agri"] = (_cl,"nunique")
                fg = ds.groupby("famille_norm").agg(**_agg).reset_index()
                fg["Valeur_DT"] = fg["Valeur_DT"].round(0)
                fg["Part_pct"]  = (fg["Valeur_DT"]/fg["Valeur_DT"].sum()*100).round(1)
            fg = fg.sort_values("Valeur_DT",ascending=False)

            c5a,c5b = st.columns(2)
            with c5a:
                fig5 = px.pie(fg,names="famille_norm",values="Valeur_DT",hole=0.4,
                    template="plotly_dark",title="Dépenses intrants par famille (DT)",
                    color_discrete_sequence=px.colors.qualitative.Set2)
                fig5.update_layout(paper_bgcolor="#161b22",height=370)
                st.plotly_chart(fig5, use_container_width=True)
            with c5b:
                fig5b = px.bar(fg, x="famille_norm", y="Valeur_DT",
                    color="Valeur_DT",
                    color_continuous_scale=["#1A5C2A","#4CAF50","#A5D6A7"],
                    template="plotly_dark", text_auto=",.0f",
                    title="Dépenses par famille (DT)")
                fig5b.update_traces(textposition="outside", textfont_size=11)
                fig5b.update_layout(paper_bgcolor="#161b22",
                    plot_bgcolor="#0d1117", height=370,
                    xaxis_tickangle=-30)
                st.plotly_chart(fig5b, use_container_width=True)

            # Tableau + téléchargement
            fg_disp = fg.rename(columns={"famille_norm":"Famille","Part_pct":"Part %"})
            st.dataframe(fg_disp, use_container_width=True, hide_index=True,
                column_config={
                    "Valeur_DT": st.column_config.NumberColumn("Valeur (DT)", format="%,.0f"),
                    "Part %":    st.column_config.ProgressColumn(
                        "Part %", min_value=0, max_value=100, format="%.1f%%"),
                })
            st.download_button(
                "📥 Télécharger Famille Intrant (Excel)",
                data=_export_excel_table(fg_disp,
                    "Famille Intrant",
                    "Dépenses Intrants par Famille — Campagne 2026",
                    "1A5C2A"),
                file_name="famille_intrant_2026.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True)

            # ══ ANALYSE AGRONOMIQUE ══════════════════════════════════
            st.markdown("---")
            st.markdown("#### 🌱 Analyse Agronomique — Intrants vs Production")
            st.caption("DAP · Engrais · Fertilissants · Fongicides · Insecticides par agriculteur")

            # ── Données disponibles ──────────────────────────────
            _ds_agro = st.session_state.get("abo_sotusfa_raw")
            _df_main = df  # merged data

            if _ds_agro is None or _ds_agro.empty:
                st.warning("⚠️ Importez le fichier **Sotusfa** dans ⚙️ pour voir cette analyse.")
            else:
                import plotly.graph_objects as _go2
                import numpy as _np_agro

                _ds_agro = _ds_agro.copy()

                # Colonnes
                _cl_agro  = next((c for c in ["client","agriculteur"] if c in _ds_agro.columns), None)
                _fam_agro = next((c for c in ["famille","famille_norm"] if c in _ds_agro.columns), None)
                _val_agro = next((c for c in ["valeur","total_ttc","Total TTC"] if c in _ds_agro.columns), None)
                _qte_agro = next((c for c in ["qte","Qte"] if c in _ds_agro.columns), None)
                _art_agro = next((c for c in ["article","Article"] if c in _ds_agro.columns), None)

                if not _cl_agro or not _fam_agro:
                    st.info("Structure Sotusfa incompatible — colonnes client/famille manquantes.")
                else:
                    # Mapping familles réelles Sotusfa → catégories agro
                    _FAM_MAP = {
                        "engrais":      "🧪 DAP & Engrais",
                        "Engrais":      "🧪 DAP & Engrais",
                        "fertilissant": "🌿 Fertilissants",
                        "Fertilissant": "🌿 Fertilissants",
                        "fongicide":    "🛡️ Fongicides",
                        "Fongicide":    "🛡️ Fongicides",
                        "insecticide":  "🦟 Insecticides",
                        "Insecticide":  "🦟 Insecticides",
                        "IRRIGATIONS":  "💧 Irrigation",
                        "IRRIGATIONS TURK": "💧 Irrigation",
                        "HERBICIDE":    "🌾 Herbicides",
                        "Divers":       "📦 Divers",
                    }
                    _ds_agro["_cat"] = _ds_agro[_fam_agro].astype(str).str.strip().map(_FAM_MAP)
                    _ds_agro = _ds_agro[_ds_agro["_cat"].notna()]

                    _val_col = _val_agro if _val_agro else (_qte_agro if _qte_agro else None)
                    if _val_col:
                        _ds_agro[_val_col] = pd.to_numeric(_ds_agro[_val_col], errors="coerce").fillna(0)

                    # Pivot : 1 ligne par client, colonnes = catégories
                    if _val_col:
                        _pivot = _ds_agro.groupby([_cl_agro, "_cat"])[_val_col].sum()                                         .unstack("_cat", fill_value=0).reset_index()
                        _pivot.columns.name = None
                        _pivot = _pivot.rename(columns={_cl_agro: "client"})
                    else:
                        _pivot = _ds_agro.groupby([_cl_agro, "_cat"])["_cat"].count()                                         .unstack("_cat", fill_value=0).reset_index()                                         .rename(columns={_cl_agro: "client"})

                    # Tonnage : depuis df merged ou depuis prévision Mai
                    _ton_col = None
                    if _df_main is not None and not _df_main.empty:
                        _agri_col = next((c for c in ["agriculteur","client"] if c in _df_main.columns), None)
                        for tc in ["tonnage_livre","prevision_juin","prevision_mai","prevision_dec"]:
                            if tc in _df_main.columns and _df_main[tc].fillna(0).sum() > 0:
                                _ton_col = tc; break
                        if _agri_col and _ton_col:
                            # Construire la liste de colonnes en vérifiant leur existence
                            _keep_cols = [_agri_col, _ton_col]
                            for _extra in ["commercial", "region"]:
                                if _extra in _df_main.columns:
                                    _keep_cols.append(_extra)
                            _ton_df = _df_main[_keep_cols].copy()
                            _ton_df = _ton_df.rename(columns={_agri_col: "client"})
                            _ton_df[_ton_col] = pd.to_numeric(_ton_df[_ton_col], errors="coerce").fillna(0)
                            _ton_df = _ton_df[_ton_df[_ton_col] > 0]
                            _ma = _pivot.merge(_ton_df, on="client", how="inner")
                        else:
                            _ma = _pivot.copy()
                            _ton_col = None
                    else:
                        _ma = _pivot.copy()
                        _ton_col = None

                    _cats = [c for c in _ma.columns if c not in
                             ["client","commercial","region",_ton_col or "x"]]

                    # ── KPIs intrants ────────────────────────────────
                    st.markdown("**📊 Total intrants Sotusfa par catégorie**")
                    _kpi_cols = st.columns(min(len(_cats), 5))
                    _COLORS_AGR = {
                        "🧪 DAP & Engrais": "#42A5F5",
                        "🌿 Fertilissants": "#66BB6A",
                        "🛡️ Fongicides":    "#AB47BC",
                        "🦟 Insecticides":  "#FF7043",
                        "💧 Irrigation":    "#26C6DA",
                        "🌾 Herbicides":    "#FFA726",
                        "📦 Divers":        "#78909C",
                    }
                    for ci4, cat in enumerate(_cats):
                        if ci4 < 5:
                            tot_cat = _ma[cat].sum() if cat in _ma.columns else 0
                            _kpi_cols[ci4].metric(cat, f"{tot_cat:,.0f} DT")

                    # ── GRAPHIQUE 1 : Scatter corrélation ────────────
                    if _ton_col and _cats:
                        st.markdown(f"**📈 Corrélation Intrants → {_ton_col.replace('_',' ').title()}**")
                        _scatter_cats = [c for c in _cats if c in ["🧪 DAP & Engrais",
                                         "🌿 Fertilissants","🛡️ Fongicides","🦟 Insecticides"]]
                        if not _scatter_cats:
                            _scatter_cats = _cats[:4]

                        _sc_cols = st.columns(min(len(_scatter_cats), 3))
                        for ci5, cat in enumerate(_scatter_cats[:3]):
                            _df_sc = _ma[_ma[cat] > 0][["client", cat, _ton_col] +
                                        (["region"] if "region" in _ma.columns else [])].copy()
                            if len(_df_sc) < 3:
                                continue
                            _x = _df_sc[cat].values
                            _y = _df_sc[_ton_col].values
                            try:
                                _a, _b = _np_agro.polyfit(_x, _y, 1)
                                _r = float(_np_agro.corrcoef(_x, _y)[0,1])
                            except Exception:
                                _a = _b = _r = 0.0

                            _fig_sc = px.scatter(
                                _df_sc, x=cat, y=_ton_col,
                                hover_name="client",
                                color="region" if "region" in _df_sc.columns else None,
                                labels={cat: f"{cat} (DT)",
                                        _ton_col: "Tonnage (T)"},
                                title=f"{cat}<br><sup>r = {_r:.2f} | {len(_df_sc)} agriculteurs</sup>",
                                template="plotly_dark",
                                color_discrete_sequence=px.colors.qualitative.Set2,
                            )
                            # Ligne de tendance
                            if abs(_r) > 0.05 and _a != 0:
                                _xl = _np_agro.linspace(_x.min(), _x.max(), 50)
                                _fig_sc.add_scatter(
                                    x=_xl, y=_a*_xl+_b,
                                    mode="lines", name="Tendance",
                                    line=dict(color=_COLORS_AGR.get(cat,"#fff"),
                                              width=2.5, dash="dot"),
                                    showlegend=False,
                                )
                            _fig_sc.update_layout(
                                paper_bgcolor="#161b22",
                                plot_bgcolor="#0d1117",
                                height=340,
                                font=dict(color="#f0f6fc", size=10),
                                showlegend=("region" in _df_sc.columns),
                                legend=dict(font=dict(size=8)),
                                margin=dict(t=60,b=30,l=50,r=20),
                            )
                            with _sc_cols[ci5 % 3]:
                                st.plotly_chart(_fig_sc, use_container_width=True)

                    # ── GRAPHIQUE 2 : Top 20 barres horizontales ─────
                    st.markdown("**🏆 Top 20 Agriculteurs — Production vs Intrants**")
                    _sort_col = _ton_col if _ton_col and _ton_col in _ma.columns else                                 (_cats[0] if _cats else None)
                    if _sort_col:
                        _top20 = _ma.nlargest(20, _sort_col).sort_values(_sort_col, ascending=True)
                        _fig_top = _go2.Figure()

                        # Barre principale = tonnage
                        if _ton_col and _ton_col in _top20.columns:
                            _fig_top.add_trace(_go2.Bar(
                                y=_top20["client"], x=_top20[_ton_col],
                                name="Tonnage (T)", orientation="h",
                                marker_color="#FF9800", marker_opacity=0.9,
                                text=_top20[_ton_col].apply(lambda v: f"{v:,.0f}T"),
                                textposition="outside", textfont=dict(size=9),
                            ))

                        # Barres intrants (en % max pour superposition)
                        for cat in (_scatter_cats if _ton_col else _cats[:4]):
                            if cat not in _top20.columns: continue
                            _max_v = _top20[cat].max()
                            _max_t = _top20[_ton_col].max() if _ton_col in _top20.columns else 1
                            _norm = _top20[cat] / _max_v * _max_t * 0.4 if _max_v > 0 else 0
                            _fig_top.add_trace(_go2.Bar(
                                y=_top20["client"], x=_norm,
                                name=f"{cat} (normalisé)",
                                orientation="h",
                                marker_color=_COLORS_AGR.get(cat,"#888"),
                                marker_opacity=0.55,
                                visible="legendonly",
                                customdata=_top20[[cat]],
                                hovertemplate="%{y}<br>" + cat + ": %{customdata[0]:,.0f} DT<extra></extra>",
                            ))

                        _fig_top.update_layout(
                            barmode="overlay",
                            template="plotly_dark",
                            paper_bgcolor="#161b22",
                            plot_bgcolor="#0d1117",
                            height=max(400, len(_top20)*22),
                            title="Top 20 — Activer les intrants dans la légende pour comparer",
                            yaxis=dict(tickfont=dict(size=9)),
                            legend=dict(orientation="h", yanchor="bottom",
                                        y=1.01, font=dict(size=9)),
                            font=dict(color="#f0f6fc"),
                            margin=dict(l=200, r=80, t=80, b=30),
                        )
                        st.plotly_chart(_fig_top, use_container_width=True)

                    # ── TABLEAU ──────────────────────────────────────
                    st.markdown("**📋 Tableau complet**")
                    _tbl_cols = ["client"] + _cats +                                 ([_ton_col] if _ton_col else []) +                                 [c for c in ["commercial","region"] if c in _ma.columns]
                    _tbl_show = _ma[[c for c in _tbl_cols if c in _ma.columns]]                                .sort_values(_sort_col if _sort_col else _tbl_cols[1],
                                             ascending=False).round(0)
                    st.dataframe(_tbl_show, use_container_width=True,
                                 hide_index=True, height=350)
                    st.download_button(
                        "📥 Excel — Analyse Intrants vs Production",
                        data=_export_excel_table(
                            _tbl_show, "Analyse Intrants",
                            "Intrants vs Production — Campagne 2026", "1A5C2A"),
                        file_name="analyse_intrants_production_2026.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True)

    # ══ TAB 6 — PRÉVISIONS VS RÉALISÉ ═════════════════════    # ══ TAB 6 — PRÉVISIONS VS RÉALISÉ ═════════════════════
    with t6:
        if df is None or df.empty:
            _no_data()
        else:
            prev_exist = [c for c in ["prevision_dec","prevision_mai",
                                       "prevision_juin","tonnage_livre"]
                          if c in df.columns]
            # Afficher même avec une seule source
            if not any(c in df.columns for c in
                       ["prevision_dec","prevision_mai","prevision_juin"]):
                st.info("Importez un fichier de prévision (Déc ou Mai) "
                        "dans l'onglet ⚙️.")
            else:
                # Totaux par période
                labels_map = {
                    "prevision_dec":  "Prévision Déc",
                    "prevision_mai":  "Prévision Mai",
                    "prevision_juin": "Prévision Juin",
                    "tonnage_livre":  "Réalisé",
                }
                tots = {}
                for col, lbl in labels_map.items():
                    if col in df.columns:
                        val = df[col].fillna(0).sum()
                        if val > 0:
                            tots[lbl] = val

                if tots:
                    bar_colors = {
                        "Prévision Déc": "#78909C",
                        "Prévision Mai": "#42A5F5",
                        "Prévision Juin":"#26A69A",
                        "Réalisé":       "#FF9800",
                    }
                    LINE_STYLES = {
                        "Prévision Déc": dict(color="#78909C", width=2, dash="dot"),
                        "Prévision Mai": dict(color="#42A5F5", width=2, dash="dash"),
                        "Prévision Juin":dict(color="#26A69A", width=2, dash="dashdot"),
                        "Réalisé":       dict(color="#FF9800", width=3),
                    }
                    MARKER_SYMS = {
                        "Prévision Déc": "circle",
                        "Prévision Mai": "square",
                        "Prévision Juin":"triangle-up",
                        "Réalisé":       "diamond",
                    }

                    # ── GRAPHIQUE 1 : Courbes superposées par commercial ──
                    st.markdown("##### 📈 Courbes superposées par commercial — Déc vs Mai vs Réalisé")
                    st.caption("Chaque courbe = une prévision | Écarts verticaux = décalages entre versions")

                    # Construire données par commercial
                    _comms_all = sorted(df["commercial"].dropna().unique()) if "commercial" in df.columns else []
                    fig6_lines = go.Figure()

                    _col_map = {
                        "Prévision Déc":  "prevision_dec",
                        "Prévision Mai":  "prevision_mai",
                        "Prévision Juin": "prevision_juin",
                        "Réalisé":        "tonnage_livre",
                    }
                    _has_any = False
                    for lbl, col in _col_map.items():
                        if col in df.columns and df[col].fillna(0).sum() > 0:
                            _has_any = True
                            if "commercial" in df.columns:
                                _y = [df[df["commercial"]==c][col].fillna(0).sum() for c in _comms_all]
                                _x = list(_comms_all)
                            else:
                                _y = [df[col].fillna(0).sum()]
                                _x = ["TOTAL"]

                            # Barre + courbe superposée
                            fig6_lines.add_trace(go.Bar(
                                name=lbl, x=_x, y=_y,
                                marker_color=bar_colors.get(lbl,"#888"),
                                marker_opacity=0.65,
                                text=[f"{v:,.0f}T" for v in _y],
                                textposition="outside",
                                textfont=dict(size=9, color=bar_colors.get(lbl,"#fff")),
                            ))
                            fig6_lines.add_trace(go.Scatter(
                                name=f"Courbe {lbl}", x=_x, y=_y,
                                mode="lines+markers",
                                line=LINE_STYLES.get(lbl, dict(color="#fff", width=2)),
                                marker=dict(size=10, symbol=MARKER_SYMS.get(lbl,"circle"),
                                            color=bar_colors.get(lbl,"#888"),
                                            line=dict(width=2, color="#fff")),
                                showlegend=True,
                            ))

                    # Ligne recouvrement
                    if "tonnage_recouvrement" in df.columns:
                        recouv = df["tonnage_recouvrement"].fillna(0).sum()
                        if recouv > 0:
                            fig6_lines.add_hline(y=recouv,
                                line_dash="dot", line_color="#ef5350", line_width=2.5,
                                annotation_text=f"⚠️ Seuil recouvrement : {recouv:,.0f} T",
                                annotation_font_color="#ef5350", annotation_font_size=11,
                                annotation_position="top right")

                    fig6_lines.update_layout(
                        barmode="group", template="plotly_dark",
                        paper_bgcolor="#161b22", plot_bgcolor="#0d1117",
                        height=520,
                        title="<b>Prévisions vs Réalisé par Commercial</b>"
                              "<br><sup>Barres = volumes | Courbes = tendances | Écart vertical = décalage entre versions</sup>",
                        yaxis_title="Tonnes (T)",
                        yaxis=dict(gridcolor="#21262d"),
                        legend=dict(orientation="h", yanchor="bottom", y=1.02,
                                    traceorder="normal"),
                        font=dict(color="#f0f6fc"),
                        bargap=0.15, bargroupgap=0.04,
                    )
                    if _has_any:
                        st.plotly_chart(fig6_lines, use_container_width=True)

                    # ── GRAPHIQUE 2 : Radar statistique ──────────────
                    if len(_comms_all) >= 3:
                        st.markdown("##### 🕷️ Radar — Comparaison globale par commercial")
                        fig_radar = go.Figure()
                        _theta = list(_comms_all) + [_comms_all[0]]
                        for lbl, col in _col_map.items():
                            if col in df.columns and df[col].fillna(0).sum() > 0 and "commercial" in df.columns:
                                _r = [df[df["commercial"]==c][col].fillna(0).sum() for c in _comms_all]
                                _r += [_r[0]]
                                fig_radar.add_trace(go.Scatterpolar(
                                    r=_r, theta=_theta, fill="toself", name=lbl,
                                    line=dict(color=bar_colors.get(lbl,"#888"), width=2),
                                    fillcolor=bar_colors.get(lbl,"#888"),
                                    opacity=0.30,
                                ))
                        fig_radar.update_layout(
                            template="plotly_dark", paper_bgcolor="#161b22",
                            polar=dict(bgcolor="#0d1117",
                                       radialaxis=dict(gridcolor="#21262d"),
                                       angularaxis=dict(gridcolor="#21262d")),
                            height=420,
                            title="Radar — Répartition par commercial (toutes prévisions)",
                            legend=dict(orientation="h", yanchor="bottom", y=-0.2),
                            font=dict(color="#f0f6fc"),
                        )
                        st.plotly_chart(fig_radar, use_container_width=True)

                # KPIs comparaison
                if tots:
                    cols_k = st.columns(len(tots))
                    for i, (lbl, val) in enumerate(tots.items()):
                        ref = tots.get("Prévision Déc", val)
                        delta = val - ref if lbl != "Prévision Déc" and ref > 0 else None
                        cols_k[i].markdown(
                            _metric(lbl, f"{val:,.0f} T",
                                color=bar_colors.get(lbl, "#888"),
                                delta=delta),
                            unsafe_allow_html=True)

                # Tableau par agriculteur
                pv_cols = [c for c in [
                    "agriculteur","centre","commercial",
                    "prevision_dec","prevision_mai",
                    "prevision_juin","tonnage_livre",
                    "tonnage_recouvrement","ecart_tonnage","alerte"
                ] if c in df.columns]

                df_pv = df[pv_cols].copy()
                df_pv.columns = [c.replace("_"," ").replace("prevision","Prév.").title()
                                  for c in df_pv.columns]

                st.markdown("#### 📋 Tableau prévisions vs réalisé par agriculteur")
                st.dataframe(df_pv.round(1),
                    use_container_width=True,
                    hide_index=True,
                    height=400,
                    column_config={
                        c: st.column_config.NumberColumn(c, format="%,.1f")
                        for c in df_pv.select_dtypes("number").columns
                    })

                st.download_button(
                    "📥 Télécharger Prévisions vs Réalisé (Excel)",
                    data=_export_excel_table(
                        df_pv,
                        "Previsions",
                        "Prévisions vs Réalisé — Campagne 2026",
                        "4A235A"),
                    file_name="previsions_vs_realise_2026.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True)

    # ══════════════════════════════════════════════════════════
    # TAB 7 — ANALYSE EFFICACITÉ PROFESSIONNELLE
    # ══════════════════════════════════════════════════════════
    with t7:
        if df is None or df.empty:
            _no_data()
            st.caption("Fusionnez les données dans ⚙️ pour voir cette analyse.")
        else:
            import numpy as _npro
            import plotly.graph_objects as _gop

            st.markdown("""<div style='background:linear-gradient(90deg,#0a1628,#1a2332);
border-radius:12px;padding:16px 24px;margin-bottom:20px;border-left:4px solid #FFD700'>
<div style='font-size:1.1rem;font-weight:800;color:#f0f6fc'>
🏆 Analyse Efficacité Professionnelle — Campagne 2026</div>
<div style='font-size:.82rem;color:#8b949e;margin-top:6px'>
Score d'efficacité · Benchmark commerciaux · Matrice ROI · Recommandations automatiques
</div></div>""", unsafe_allow_html=True)

            _df7 = df.copy()
            for _col in ["hectares","tonnage_livre","taux_prise","charge_totale",
                         "valeur_livree","total_intrants","solde_final"]:
                if _col not in _df7.columns:
                    _df7[_col] = 0
                _df7[_col] = pd.to_numeric(_df7[_col], errors="coerce").fillna(0)
            for _sc2 in ["region","variete","commercial"]:
                if _sc2 not in _df7.columns:
                    _df7[_sc2] = ""

            # Métriques dérivées
            _df7["rendement_ha"] = _df7.apply(
                lambda r: round(r["tonnage_livre"]/r["hectares"],1) if r["hectares"]>0 else 0, axis=1)
            _df7["cout_intrant_tonne"] = _df7.apply(
                lambda r: round(r["total_intrants"]/r["tonnage_livre"],1) if r["tonnage_livre"]>0 else 0, axis=1)
            _df7["cout_intrant_ha"] = _df7.apply(
                lambda r: round(r["total_intrants"]/r["hectares"],1) if r["hectares"]>0 else 0, axis=1)
            _df7["roi_pct"] = _df7.apply(
                lambda r: round((r["valeur_livree"]-r["charge_totale"])/r["charge_totale"]*100,1)
                          if r["charge_totale"]>0 else 0, axis=1)

            # Score Efficacité (0-100)
            def _norm7(s, inv=False):
                mn,mx=s.min(),s.max()
                if mx==mn: return pd.Series([0.5]*len(s),index=s.index)
                n=(s-mn)/(mx-mn)
                return 1-n if inv else n

            _df7["_s_rend"] = 0.0
            for _rg in _df7["region"].dropna().unique():
                _mk = _df7["region"]==_rg
                if _mk.sum()>1:
                    _df7.loc[_mk,"_s_rend"] = _norm7(_df7.loc[_mk,"rendement_ha"])*0.40
                else:
                    _df7.loc[_mk,"_s_rend"] = 0.20

            _sot_gt0 = _df7["cout_intrant_tonne"].where(_df7["cout_intrant_tonne"]>0, 0)
            _df7["_s_int"] = _norm7(_sot_gt0, inv=True)*0.30
            _df7["_s_prise"] = _norm7(_df7["taux_prise"])*0.20
            _df7["_s_roi"] = (_df7["roi_pct"]>0).astype(float)*0.10
            _df7["score_efficacite"] = ((_df7["_s_rend"]+_df7["_s_int"]+_df7["_s_prise"]+_df7["_s_roi"])*100).round(1)

            def _cat7(s):
                if s>=75: return "🏆 Excellent"
                elif s>=55: return "✅ Bon"
                elif s>=35: return "⚠️ Moyen"
                else: return "🔴 À améliorer"
            _df7["categorie"] = _df7["score_efficacite"].apply(_cat7)

            # ── KPIs ─────────────────────────────────────────
            st.markdown("### 📊 Indicateurs Clés")
            _kp = st.columns(5)
            _kp[0].markdown(_metric("Score moyen",f"{_df7['score_efficacite'].mean():.1f}/100",color="#FFD700"),unsafe_allow_html=True)
            _rend_pos = _df7["rendement_ha"][_df7["rendement_ha"]>0]
            _kp[1].markdown(_metric("Rendement moyen",f"{_rend_pos.mean():.1f} t/ha" if len(_rend_pos)>0 else "N/A",color="#4CAF50"),unsafe_allow_html=True)
            _cout_pos = _df7["cout_intrant_tonne"][_df7["cout_intrant_tonne"]>0]
            _kp[2].markdown(_metric("Coût intrant/tonne",f"{_cout_pos.mean():.0f} DT/T" if len(_cout_pos)>0 else "N/A",color="#FF9800"),unsafe_allow_html=True)
            _kp[3].markdown(_metric("ROI moyen",f"{_df7['roi_pct'].mean():+.1f}%",
                color="#4CAF50" if _df7["roi_pct"].mean()>=0 else "#ef5350"),unsafe_allow_html=True)
            _kp[4].markdown(_metric("Excellents (≥75)",f"{(_df7['score_efficacite']>=75).sum()} agri",color="#FFD700"),unsafe_allow_html=True)

            # ── Matrice Efficacité ────────────────────────────
            st.markdown("---")
            st.markdown("### 🔷 Matrice Efficacité — Coût Intrant vs Rendement")
            st.caption("4 quadrants : Efficace · Surinvesti · Potentiel · Inefficace")
            _dfm = _df7[(_df7["cout_intrant_ha"]>0)&(_df7["rendement_ha"]>0)].copy()
            if not _dfm.empty:
                _med_c = _dfm["cout_intrant_ha"].median()
                _med_r = _dfm["rendement_ha"].median()
                _ac = next((c for c in ["agriculteur","client"] if c in _dfm.columns),None)
                _CAT_COL = {"🏆 Excellent":"#FFD700","✅ Bon":"#4CAF50","⚠️ Moyen":"#FF9800","🔴 À améliorer":"#ef5350"}
                _figm = _gop.Figure()
                for _cat7v, _cc in _CAT_COL.items():
                    _sub = _dfm[_dfm["categorie"]==_cat7v]
                    if _sub.empty: continue
                    _figm.add_trace(_gop.Scatter(
                        x=_sub["cout_intrant_ha"], y=_sub["rendement_ha"],
                        mode="markers", name=_cat7v,
                        marker=dict(color=_cc,size=10,opacity=0.85,line=dict(width=1,color="#fff")),
                        text=_sub[_ac] if _ac else None,
                        hovertemplate="<b>%{text}</b><br>Coût: %{x:,.0f} DT/ha<br>Rend: %{y:.1f} t/ha<extra></extra>",
                    ))
                _figm.add_hline(y=_med_r,line_dash="dash",line_color="#555",line_width=1.5,
                    annotation_text=f"Médiane {_med_r:.1f}t/ha",annotation_font_color="#999")
                _figm.add_vline(x=_med_c,line_dash="dash",line_color="#555",line_width=1.5,
                    annotation_text=f"Médiane {_med_c:.0f}DT/ha",annotation_font_color="#999")
                _xmax = _dfm["cout_intrant_ha"].quantile(0.9)
                _ymax = _dfm["rendement_ha"].quantile(0.9)
                for _ql,_xa,_ya,_qc in [
                    ("⭐ EFFICACE",_med_c*0.3,_ymax*0.9,"#4CAF50"),
                    ("💸 SURINVESTI",_xmax*0.75,_ymax*0.9,"#FF9800"),
                    ("🔍 POTENTIEL",_med_c*0.3,_med_r*0.3,"#42A5F5"),
                    ("❌ INEFFICACE",_xmax*0.75,_med_r*0.3,"#ef5350")]:
                    _figm.add_annotation(x=_xa,y=_ya,text=_ql,showarrow=False,
                        font=dict(size=10,color=_qc),bgcolor="rgba(0,0,0,0.5)",
                        bordercolor=_qc,borderwidth=1,borderpad=4)
                _figm.update_layout(template="plotly_dark",paper_bgcolor="#161b22",
                    plot_bgcolor="#0d1117",height=460,
                    xaxis_title="Coût intrants / ha (DT)",yaxis_title="Rendement (t/ha)",
                    legend=dict(orientation="h",y=1.02,font=dict(size=10)),font=dict(color="#f0f6fc"))
                st.plotly_chart(_figm,use_container_width=True)

            # ── Benchmark Commerciaux ─────────────────────────
            st.markdown("---")
            st.markdown("### 👔 Benchmark Commerciaux — Radar 5 axes")
            if _df7["commercial"].ne("").any():
                _cb = _df7.groupby("commercial").agg(
                    Rend_moy=("rendement_ha",lambda x: round(x[x>0].mean(),1) if (x>0).any() else 0),
                    Taux_prise=("taux_prise",lambda x: round(x[x>0].mean(),1) if (x>0).any() else 0),
                    Cout_T=("cout_intrant_tonne",lambda x: round(x[x>0].mean(),0) if (x>0).any() else 0),
                    ROI_pos=("roi_pct",lambda x: round((x>0).mean()*100,0)),
                    Score=("score_efficacite","mean"),
                    Nb=("score_efficacite","count"),
                    Tonnage=("tonnage_livre","sum"),
                ).reset_index().round(1)

                def _n100(s,inv=False):
                    mn,mx=s.min(),s.max()
                    if mx==mn: return pd.Series([50.0]*len(s),index=s.index)
                    n=(s-mn)/(mx-mn)*100
                    return 100-n if inv else n

                _rb = _cb.copy()
                _rb["nRend"]=_n100(_rb["Rend_moy"])
                _rb["nPrise"]=_n100(_rb["Taux_prise"])
                _rb["nCout"]=_n100(_rb["Cout_T"],inv=True)
                _rb["nROI"]=_n100(_rb["ROI_pos"])
                _rb["nScore"]=_n100(_rb["Score"])

                _CCOL={"KHALIL":"#F5A623","KHALIL MAIRECH":"#F5A623","MAKKI BEN SALAH":"#00E5A0",
                       "FEDI":"#3B82F6","JILANI OBAY":"#FF6B9D","ACHREF AJLANI":"#8B5CF6"}
                _theta7=["Rendement/ha","Taux prise","Coût maîtrisé","ROI agri","Score global"]

                _figr = _gop.Figure()
                for _,_rw in _rb.iterrows():
                    _cm = str(_rw["commercial"])
                    _rv = [_rw["nRend"],_rw["nPrise"],_rw["nCout"],_rw["nROI"],_rw["nScore"],_rw["nRend"]]
                    _figr.add_trace(_gop.Scatterpolar(
                        r=_rv,theta=_theta7+[_theta7[0]],fill="toself",name=_cm,
                        line=dict(color=_CCOL.get(_cm,"#888"),width=2),
                        fillcolor=_CCOL.get(_cm,"#888"),opacity=0.2))
                _figr.update_layout(template="plotly_dark",paper_bgcolor="#161b22",
                    polar=dict(bgcolor="#0d1117",
                               radialaxis=dict(gridcolor="#21262d",range=[0,100]),
                               angularaxis=dict(gridcolor="#21262d")),
                    height=420,title="Radar Efficacité — 5 Commerciaux",
                    legend=dict(orientation="h",y=-0.15,font=dict(size=10)),font=dict(color="#f0f6fc"))

                _cb_sorted = _cb.sort_values("Score",ascending=True)
                _figb7 = _gop.Figure()
                _figb7.add_trace(_gop.Bar(
                    y=_cb_sorted["commercial"],x=_cb_sorted["Score"],orientation="h",
                    marker_color=[_CCOL.get(c,"#888") for c in _cb_sorted["commercial"]],
                    text=_cb_sorted["Score"].apply(lambda v: f"{v:.1f}/100"),
                    textposition="outside",textfont=dict(size=11)))
                _figb7.update_layout(template="plotly_dark",paper_bgcolor="#161b22",
                    plot_bgcolor="#0d1117",height=260,title="Classement Score",
                    xaxis=dict(range=[0,100]),font=dict(color="#f0f6fc"),
                    margin=dict(l=160,r=80,t=50,b=20))

                _cr1,_cr2 = st.columns([3,2])
                with _cr1: st.plotly_chart(_figr,use_container_width=True)
                with _cr2:
                    st.plotly_chart(_figb7,use_container_width=True)
                    st.dataframe(_cb[["commercial","Score","Rend_moy","Taux_prise",
                                      "ROI_pos","Nb","Tonnage"]]                        .rename(columns={"commercial":"Commercial","Score":"Score/100",
                            "Rend_moy":"Rend(t/ha)","Taux_prise":"Taux prise",
                            "ROI_pos":"% ROI positif","Nb":"Nb agri","Tonnage":"Tonnage(T)"})                        .sort_values("Score/100",ascending=False),
                        hide_index=True,use_container_width=True)

            # ── Analyse par Variété ───────────────────────────
            if _df7["variete"].ne("").any() and (_df7["rendement_ha"]>0).any():
                st.markdown("---")
                st.markdown("### 🍅 Efficacité par Variété")
                _vg = _df7[_df7["rendement_ha"]>0].groupby("variete").agg(
                    Rend_moy=("rendement_ha","mean"),Score_moy=("score_efficacite","mean"),
                    Cout_moy=("cout_intrant_tonne","mean"),Nb=("rendement_ha","count"),
                ).reset_index().sort_values("Rend_moy",ascending=False).round(1)
                _vg["Recommandation"] = ["⭐ MEILLEURE" if i==0
                    else ("✅ Bonne" if r["Rend_moy"]>=_vg["Rend_moy"].median() else "💡 À optimiser")
                    for i,(_,r) in enumerate(_vg.iterrows())]
                _vg_c1,_vg_c2 = st.columns(2)
                with _vg_c1:
                    _figv=_gop.Figure()
                    _figv.add_trace(_gop.Bar(x=_vg["variete"],y=_vg["Rend_moy"],
                        marker_color=["#FFD700" if i==0 else "#42A5F5" for i in range(len(_vg))],
                        text=_vg["Rend_moy"].apply(lambda v: f"{v:.1f} t/ha"),textposition="outside"))
                    _figv.update_layout(template="plotly_dark",paper_bgcolor="#161b22",
                        plot_bgcolor="#0d1117",height=320,title="Rendement moyen par variété",
                        yaxis_title="t/ha",font=dict(color="#f0f6fc"))
                    st.plotly_chart(_figv,use_container_width=True)
                with _vg_c2:
                    st.dataframe(_vg.rename(columns={"variete":"Variété","Rend_moy":"Rend(t/ha)",
                        "Score_moy":"Score/100","Cout_moy":"Coût/T(DT)","Nb":"Nb agri"}),
                        hide_index=True,use_container_width=True,height=320)

            # ── Tableau complet avec score ────────────────────
            st.markdown("---")
            st.markdown("### 📋 Tableau Score Efficacité Complet")
            _ac7 = next((c for c in ["agriculteur","client"] if c in _df7.columns),None)
            _v7 = [c for c in [_ac7,"commercial","region","variete","score_efficacite",
                "categorie","rendement_ha","taux_prise","cout_intrant_tonne","roi_pct","solde_final"]
                if c and c in _df7.columns]
            _df7d = _df7[_v7].sort_values("score_efficacite",ascending=False).round(1)
            st.dataframe(_df7d,hide_index=True,use_container_width=True,height=380,
                column_config={
                    "score_efficacite":st.column_config.ProgressColumn("Score/100",min_value=0,max_value=100,format="%.1f"),
                    "roi_pct":st.column_config.NumberColumn("ROI%",format="%+.1f%%"),
                    "rendement_ha":st.column_config.NumberColumn("Rend(t/ha)",format="%.1f"),
                    "cout_intrant_tonne":st.column_config.NumberColumn("Coût/T(DT)",format="%.0f"),
                    "solde_final":st.column_config.NumberColumn("Solde(DT)",format="%+,.0f"),
                })

            # ── Top & Bottom ──────────────────────────────────
            _tb1,_tb2 = st.columns(2)
            with _tb1:
                st.markdown("**⭐ Top 10**")
                for idx,(_,r) in enumerate(_df7d.head(10).iterrows()):
                    _med = ["🥇","🥈","🥉","4️⃣","5️⃣","6️⃣","7️⃣","8️⃣","9️⃣","🔟"][idx]
                    _nm7 = r[_ac7] if _ac7 and _ac7 in r else str(r.name)
                    st.caption(f"{_med} **{_nm7}** — Score {r['score_efficacite']:.0f}/100 | Rend {r['rendement_ha']:.1f}t/ha | ROI {r['roi_pct']:+.0f}%")
            with _tb2:
                st.markdown("**🔴 À améliorer (10 derniers)**")
                for _,r in _df7d.tail(10).sort_values("score_efficacite").iterrows():
                    _nm7 = r[_ac7] if _ac7 and _ac7 in r else str(r.name)
                    _cause = ("faible rendement" if r["rendement_ha"]<15
                              else ("coût élevé" if r["cout_intrant_tonne"]>80 else "taux prise bas"))
                    st.caption(f"⚠️ **{_nm7}** — {r['score_efficacite']:.0f}/100 | Cause : {_cause}")

            # ── Recommandations ───────────────────────────────
            st.markdown("---")
            st.markdown("### 💡 Recommandations Automatiques")
            _rc1,_rc2 = st.columns(2)
            with _rc1:
                if _df7["variete"].ne("").any() and (_df7["rendement_ha"]>0).any():
                    _bv = _df7[_df7["rendement_ha"]>0].groupby("variete")["rendement_ha"].mean().idxmax()
                    st.info(f"🌱 **Variété recommandée : {_bv}** — meilleur rendement moyen. Priorité pour la prochaine campagne.")
                if _df7["commercial"].ne("").any():
                    _bc = _df7.groupby("commercial")["score_efficacite"].mean().idxmax()
                    st.success(f"👔 **Meilleur commercial : {_bc}** — partager ses méthodes de suivi avec les autres équipes.")
            with _rc2:
                _n_sous = (_df7["score_efficacite"]<35).sum()
                _n_sur  = (_df7["cout_intrant_ha"]>_df7["cout_intrant_ha"].quantile(0.75)).sum()
                if _n_sous > 0:
                    st.warning(f"⚠️ **{_n_sous} agriculteurs** ont un score < 35/100 — nécessitent un accompagnement terrain urgent.")
                if _n_sur > 0:
                    st.error(f"💸 **{_n_sur} agriculteurs** surinvestissent en intrants (quartile supérieur) sans rendement proportionnel — rationaliser les doses DAP/fongicides.")

            st.download_button("📥 Excel — Analyse Efficacité Complète",
                data=_export_excel_table(_df7d.rename(columns={
                    "score_efficacite":"Score/100","categorie":"Catégorie",
                    "rendement_ha":"Rend(t/ha)","cout_intrant_tonne":"Coût/T(DT)",
                    "roi_pct":"ROI%","solde_final":"Solde(DT)"}),
                    "Analyse Efficacite","Score Efficacité & ROI — 2026","FFD700"),
                file_name="analyse_efficacite_2026.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True)