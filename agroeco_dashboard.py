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
    df = df[~df["client"].str.upper().isin(["","NAN","TOTAL","SOUS-TOTAL"])]
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
        return df[keep], ""
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

    # ── Merge PRÉVISIONS ───────────────────────────────────
    for df_p, col in [(df_prev_dec,"prevision_dec"),
                      (df_prev_mai,"prevision_mai"),
                      (df_prev_juin,"prevision_juin")]:
        if df_p is not None and not df_p.empty and col in df_p.columns:
            p = df_p.copy()
            p = _upper(p, [c for c in ["client","centre"] if c in p.columns])
            # Merge sur client seul si centre vide/absent dans le fichier prévision
            merge_key = []
            for k in ["client","centre"]:
                if k in p.columns and k in base.columns:
                    vals = p[k].astype(str).str.strip()
                    if vals.replace("","NaN").ne("NaN").any():
                        merge_key.append(k)
            if not merge_key and "client" in p.columns and "client" in base.columns:
                merge_key = ["client"]
            if merge_key:
                p_clean = p[merge_key + [col]].drop_duplicates(subset=merge_key)
                base = base.merge(p_clean, on=merge_key, how="left")

    # ══ CALCULS ═══════════════════════════════════════════
    df = base.copy()
    def g(col, d=0):
        return pd.to_numeric(df.get(col, d), errors="coerce").fillna(d)

    # Charges
    df["charge_plants"]   = g("valeur_plants")
    df["charge_intrants"] = g("total_intrants")
    df["avance_bourak"]   = g("avance")
    df["charge_totale"]   = df["charge_plants"] + df["charge_intrants"] + df["avance_bourak"]

    # Consigne caisse — condition : affectation_caisse = "1ère"
    prix_caisse   = params.get("prix_caisse", 0)
    nb_caisses_ha = params.get("nb_caisses_ha", 0)
    df["consigne_caisse"] = df.apply(
        lambda row: (float(row.get("hectares",0) or 0) * nb_caisses_ha * prix_caisse
                     if str(row.get("affectation_caisse","")).startswith("1ère")
                     else 0.0), axis=1)

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
    caisse_cols = [c for c in ["agriculteur","centre","region","affectation_caisse",
                                "date_debut_recolte","hectares",
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

    t0,t1,t2,t3,t4,t5,t6 = st.tabs([
        "⚙️ Paramètres & Import",
        "📋 Par Agriculteur",
        "👤 Par Ingénieur / Centre",
        "🗺️ Par Région",
        "🍅 Par Variété",
        "💊 Par Famille Intrant",
        "📈 Prévisions vs Réalisé",
    ])

    # ══ TAB 0 — PARAMÈTRES ET IMPORT ══════════════════════
    with t0:
        # ── Paramètres ──────────────────────────────────────
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
            st.markdown("**📦 Caisses vides**")
            st.info("Condition : **date début RÉCOLTE** (Supabase) < 10 juillet\n\n"
                    "→ 1ère affectation : avec caisses\n\n"
                    "→ 2ème affectation : 0 DT automatique")
            prix_caisse   = st.number_input("Prix/caisse (DT)",0.0,50.0,3.0,0.5,key="pc")
            nb_caisses_ha = st.number_input("Nb caisses/ha",  0.0,500.0,80.0,10.0,key="nc")
            mo_tonne = st.number_input("MO récolte (DT/T)",0.0,200.0,50.0,5.0,key="mo")

        params = {
            "prix_vente_global": prix_global,
            "prix_consigne": {
                "Pltx 228 PVC":p228pvc,"Pltx 228 POLY":p228poly,
                "Pltx 160 PVC":p160pvc,"Pltx 160 POLY":p160poly,
            },
            "prix_caisse":   prix_caisse,
            "nb_caisses_ha": nb_caisses_ha,
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
            kc[1].markdown(_metric("Charge totale",f"{df['charge_totale'].sum():,.0f} DT",color="#FF9800"),unsafe_allow_html=True)
            kc[2].markdown(_metric("Recouvrement",f"{df['tonnage_recouvrement'].sum():,.1f} T",color="#ef5350"),unsafe_allow_html=True)
            kc[3].markdown(_metric("Livré réel",f"{df['tonnage_livre'].sum():,.1f} T",color="#4CAF50"),unsafe_allow_html=True)
            n_crit=(df["alerte"].str.contains("🔴")).sum()
            kc[4].markdown(_metric("⚠️ Critiques",n_crit,color="#ef5350"),unsafe_allow_html=True)
            kc[5].markdown(_metric("Solde global",f"{df['solde_final'].sum():+,.0f} DT",
                color="#4CAF50" if df["solde_final"].sum()>=0 else "#ef5350"),unsafe_allow_html=True)

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

            VIEW = [c for c in ["agriculteur","commercial","ingenieur","centre","variete",
                                  "hectares","affectation_caisse","taux_prise",
                                  "charge_totale","consigne_caisse","consigne_plateau",
                                  "mo_recolte","tonnage_recouvrement","tonnage_livre",
                                  "ecart_tonnage","valeur_livree","solde_final",
                                  "report","alerte"] if c in df_f.columns]
            st.dataframe(df_f[VIEW].round(1),
                use_container_width=True,hide_index=True,height=400,
                column_config={
                    "taux_prise":st.column_config.ProgressColumn(
                        "Taux prise %",min_value=0,max_value=100,format="%.1f%%"),
                    "ecart_tonnage":st.column_config.NumberColumn("Écart (T)",format="%+.1f"),
                    "solde_final":st.column_config.NumberColumn("Solde (DT)",format="%+,.0f"),
                })
            xl2 = export_excel(df_f, st.session_state.get("abo_sotusfa_raw"))
            st.download_button("📥 Exporter cette vue",data=xl2,
                file_name="agroeco_vue.xlsx",
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

    # ══ TAB 6 — PRÉVISIONS VS RÉALISÉ ═════════════════════
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
                    # Graphique évolution
                    fig6 = go.Figure()
                    bar_colors = {
                        "Prévision Déc": "#78909C",
                        "Prévision Mai": "#42A5F5",
                        "Prévision Juin":"#1E8449",
                        "Réalisé":       "#FF9800",
                    }
                    for lbl, val in tots.items():
                        fig6.add_trace(go.Bar(
                            name=lbl, x=[lbl], y=[val],
                            marker_color=bar_colors.get(lbl,"#888"),
                            text=[f"{val:,.0f} T"],
                            textposition="outside",
                            textfont_size=13))

                    # Ligne recouvrement
                    if "tonnage_recouvrement" in df.columns:
                        recouv = df["tonnage_recouvrement"].fillna(0).sum()
                        if recouv > 0:
                            fig6.add_hline(y=recouv,
                                line_dash="dash", line_color="#ef5350",
                                line_width=2.5,
                                annotation_text=f"⚠️ Recouvrement : {recouv:,.0f} T",
                                annotation_font_color="#ef5350",
                                annotation_font_size=12)

                    fig6.update_layout(
                        template="plotly_dark",
                        paper_bgcolor="#161b22",
                        plot_bgcolor="#0d1117",
                        height=420, showlegend=False,
                        title="Évolution prévisions vs Recouvrement minimal",
                        yaxis_title="Tonnes")
                    st.plotly_chart(fig6, use_container_width=True)

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