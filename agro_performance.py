# ============================================================
# agro_performance.py — Module Performance Agronomique
# Onglet "🌱 Performance Agronomique" du dashboard tomate 2026
#
# À placer dans le même dossier que dashboard_phase10.py
# Import dans dashboard_phase10.py :
#   from agro_performance import render_agro_tab
#
# Supabase : créer la table agri_performance (SQL ci-dessous)
# ============================================================

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import io

# ══════════════════════════════════════════════════════════════
# CONSTANTES AGRONOMIQUES
# ══════════════════════════════════════════════════════════════

# Densité de plantation standard tomate Tunisie (plants/ha)
DENSITE_PLANTS = {
    "CAP BON 1":          25000,
    "CAP BON 2":          25000,
    "NORD":               22000,
    "KAIROUAN":           20000,
    "GAFSA / KASSRINE":   18000,
    "SIDI BOUZID":        18000,
    "BOUFICHA":           22000,
}
DENSITE_DEFAULT = 22000

# Rendement moyen national tomate industrielle Tunisie (t/ha)
RENDEMENT_MOYEN_NATIONAL = 35.0

# Rendement moyen par région (t/ha) — références terrain 2025
RENDEMENT_MOYEN_REGION = {
    "CAP BON 1":          45.0,
    "CAP BON 2":          42.0,
    "NORD":               35.0,
    "KAIROUAN":           30.0,
    "GAFSA / KASSRINE":   28.0,
    "SIDI BOUZID":        28.0,
    "BOUFICHA":           38.0,
}

# Doses de référence (kg/ha) — bonnes pratiques Tunisie
REF_DOSES = {
    "DAP":      {"optimal": 150, "max": 250, "unit": "kg/ha"},
    "FUMURE":   {"optimal": 200, "max": 400, "unit": "kg/ha"},
    "FUMIER":   {"optimal": 20000, "max": 40000, "unit": "kg/ha"},
    "PESTICIDE":{"optimal": 3, "max": 8, "unit": "L/ha"},
}

# Seuils de performance (indice rendement vs. moyenne région)
SCORE_LABELS = {
    (1.15, 999):  ("⭐ Excellent",    "#1E8449"),
    (1.05, 1.15): ("✅ Bon",          "#2E86C1"),
    (0.90, 1.05): ("🟡 Moyen",        "#D4AC0D"),
    (0.75, 0.90): ("🟠 Faible",       "#CA6F1E"),
    (-999, 0.75): ("🔴 Sous-perf.",   "#C0392B"),
}

def get_score_label(rendement, region):
    """Retourne le label et couleur de performance d'un agriculteur."""
    ref = RENDEMENT_MOYEN_REGION.get(region, RENDEMENT_MOYEN_NATIONAL)
    ratio = rendement / ref if ref > 0 else 0
    for (lo, hi), (label, color) in SCORE_LABELS.items():
        if lo <= ratio < hi:
            return label, color, ratio
    return "🔴 Sous-perf.", "#C0392B", ratio


# ══════════════════════════════════════════════════════════════
# FORMULES DE CALCUL AGRONOMIQUE
# ══════════════════════════════════════════════════════════════

def calculer_rendement(tonnage_t, hectares):
    """Rendement = Tonnage (t) / Hectares (ha)"""
    if hectares and hectares > 0:
        return round(tonnage_t / hectares, 2)
    return None

def calculer_rendement_par_plant(tonnage_kg, hectares, region=None):
    """Rendement par plant = Tonnage (kg) / Nb plants total"""
    densite = DENSITE_PLANTS.get(region, DENSITE_DEFAULT)
    total_plants = densite * hectares if hectares and hectares > 0 else None
    if total_plants and total_plants > 0:
        return round(tonnage_kg / total_plants, 3)  # kg/plant
    return None

def calculer_dose_ha(quantite_totale, hectares, unite="kg"):
    """Dose par hectare = Quantité totale / Hectares"""
    if hectares and hectares > 0 and quantite_totale is not None:
        return round(quantite_totale / hectares, 2)
    return None

def calculer_efficacite_intrants(tonnage_t, total_engrais_kg):
    """
    Efficacité = kg de tomate produits / kg d'engrais consommés
    (plus le ratio est élevé, plus l'agriculteur est efficace)
    """
    if total_engrais_kg and total_engrais_kg > 0:
        return round((tonnage_t * 1000) / total_engrais_kg, 3)
    return None

def calculer_indice_intrant(dose_reelle, dose_optimale):
    """
    Indice d'utilisation = dose_réelle / dose_optimale
    < 0.8  = sous-fertilisation
    0.8-1.2 = optimal
    > 1.2  = sur-fertilisation (gaspillage + risque)
    """
    if dose_optimale and dose_optimale > 0 and dose_reelle is not None:
        return round(dose_reelle / dose_optimale, 3)
    return None

def score_commercial(df_agri):
    """
    Score agronomique d'un commercial = moyenne pondérée par tonnage
    des ratios rendement/moyenne_région de ses agriculteurs.
    """
    if df_agri.empty or "rendement_t_ha" not in df_agri.columns:
        return 0
    total_tonnage = df_agri["tonnage_t"].sum()
    if total_tonnage == 0:
        return 0
    df_agri = df_agri.copy()
    df_agri["ref"] = df_agri["region"].map(
        lambda r: RENDEMENT_MOYEN_REGION.get(r, RENDEMENT_MOYEN_NATIONAL)
    )
    df_agri["ratio"] = df_agri["rendement_t_ha"] / df_agri["ref"].replace(0, np.nan)
    df_agri["poids"] = df_agri["tonnage_t"] / total_tonnage
    score = (df_agri["ratio"] * df_agri["poids"]).sum()
    return round(score, 3)


# ══════════════════════════════════════════════════════════════
# SQL SUPABASE — CREATE TABLE
# ══════════════════════════════════════════════════════════════
SQL_CREATE_TABLE = """
-- ============================================================
-- Table Supabase : agri_performance
-- Exécuter une seule fois dans l'éditeur SQL de Supabase
-- ============================================================
CREATE TABLE IF NOT EXISTS agri_performance (
    id               BIGSERIAL PRIMARY KEY,
    commercial       TEXT NOT NULL,
    ingenieur        TEXT,
    agriculteur      TEXT NOT NULL,
    region           TEXT,
    variete          TEXT,
    hectares         NUMERIC,
    tonnage_t        NUMERIC,          -- tonnage récolté (tonnes)
    dap_kg           NUMERIC,          -- engrais DAP (kg total)
    fumure_kg        NUMERIC,          -- fumure organique (kg total)
    fumier_kg        NUMERIC,          -- fumier (kg total)
    pesticide_l      NUMERIC,          -- pesticides (litres total)
    fongicide_l      NUMERIC,          -- fongicides (litres)
    insecticide_l    NUMERIC,          -- insecticides (litres)
    irrigation_m3    NUMERIC,          -- irrigation (m³)
    notes            TEXT,
    saison           TEXT DEFAULT '2026',
    created_at       TIMESTAMPTZ DEFAULT NOW(),
    updated_at       TIMESTAMPTZ DEFAULT NOW()
);

CREATE INDEX IF NOT EXISTS idx_ap_commercial ON agri_performance(commercial);
CREATE INDEX IF NOT EXISTS idx_ap_region     ON agri_performance(region);
CREATE INDEX IF NOT EXISTS idx_ap_saison     ON agri_performance(saison);

-- Vue calculée (colonnes dérivées automatiques)
CREATE OR REPLACE VIEW agri_performance_calcule AS
SELECT
    *,
    CASE WHEN hectares > 0 THEN ROUND(tonnage_t / hectares, 2) END AS rendement_t_ha,
    CASE WHEN hectares > 0 AND dap_kg IS NOT NULL
         THEN ROUND(dap_kg / hectares, 1) END AS dap_kg_ha,
    CASE WHEN hectares > 0 AND fumure_kg IS NOT NULL
         THEN ROUND(fumure_kg / hectares, 1) END AS fumure_kg_ha,
    CASE WHEN hectares > 0 AND pesticide_l IS NOT NULL
         THEN ROUND(pesticide_l / hectares, 2) END AS pesticide_l_ha,
    CASE WHEN (COALESCE(dap_kg,0) + COALESCE(fumure_kg,0)) > 0
         THEN ROUND((tonnage_t * 1000) / (COALESCE(dap_kg,0) + COALESCE(fumure_kg,0)), 3)
         END AS efficacite_intrants
FROM agri_performance;
"""


# ══════════════════════════════════════════════════════════════
# TEMPLATE EXCEL D'IMPORT
# ══════════════════════════════════════════════════════════════
def generate_import_template():
    """Génère le fichier Excel template pour la saisie des données."""
    from openpyxl import Workbook
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.data_validation import DataValidation

    wb = Workbook()
    ws = wb.active
    ws.title = "Données Agriculteurs"

    HDR_FILL = PatternFill("solid", start_color="1F3864", end_color="1F3864")
    HDR_FONT = Font(bold=True, color="FFFFFF", name="Calibri", size=11)
    REQ_FILL = PatternFill("solid", start_color="E8F4FD", end_color="E8F4FD")
    OPT_FILL = PatternFill("solid", start_color="FFF9E6", end_color="FFF9E6")
    THIN     = Border(
        left=Side(style="thin", color="CCCCCC"),
        right=Side(style="thin", color="CCCCCC"),
        top=Side(style="thin", color="CCCCCC"),
        bottom=Side(style="thin", color="CCCCCC"),
    )
    CTR = Alignment(horizontal="center", vertical="center")

    COLS = [
        # (header, largeur, obligatoire)
        ("Commercial",       18, True),
        ("Ingénieur",        18, False),
        ("Agriculteur",      28, True),
        ("Région",           18, True),
        ("Variété",          16, False),
        ("Hectares (ha)",    14, True),
        ("Tonnage récolté (t)", 18, True),
        ("DAP (kg total)",   14, False),
        ("Fumure (kg total)",14, False),
        ("Fumier (kg total)",14, False),
        ("Pesticides (L total)", 16, False),
        ("Fongicides (L total)", 16, False),
        ("Insecticides (L total)", 18, False),
        ("Irrigation (m³)", 14, False),
        ("Notes",           30, False),
    ]

    # En-tête
    for ci, (header, width, required) in enumerate(COLS, 1):
        c = ws.cell(1, ci)
        c.value    = header
        c.fill     = HDR_FILL
        c.font     = HDR_FONT
        c.alignment= CTR
        c.border   = THIN
        ws.column_dimensions[get_column_letter(ci)].width = width

    ws.row_dimensions[1].height = 30

    # Ligne exemple
    ex = ["FEDI", "Ing. BEN ALI", "AMOR KHECHIN", "CAP BON 1", "Heinz",
          5.0, 180.0, 750, 1000, 20000, 15.0, 4.0, 2.0, 1500, ""]
    for ci, val in enumerate(ex, 1):
        c = ws.cell(2, ci)
        c.value    = val
        required   = COLS[ci-1][2]
        c.fill     = REQ_FILL if required else OPT_FILL
        c.border   = THIN
        c.alignment= CTR

    # Validation région
    dv_reg = DataValidation(
        type="list",
        formula1='"CAP BON 1,CAP BON 2,NORD,KAIROUAN,GAFSA / KASSRINE,SIDI BOUZID,BOUFICHA"',
        allow_blank=True,
    )
    ws.add_data_validation(dv_reg)
    dv_reg.add(f"D2:D500")

    # Feuille légende
    ws2 = wb.create_sheet("Légende et Formules")
    ws2.column_dimensions["A"].width = 30
    ws2.column_dimensions["B"].width = 50
    legend = [
        ("CHAMP",           "DESCRIPTION ET FORMULE"),
        ("Hectares",        "Surface cultivée en hectares (ha)"),
        ("Tonnage récolté", "Poids total récolté en TONNES (pas kg)"),
        ("DAP",             "Diammonium phosphate — engrais starter (kg total saison)"),
        ("Fumure",          "Fumure organique / engrais de fond (kg total)"),
        ("Fumier",          "Fumier animal (kg total — 20 à 40 tonnes/ha normal)"),
        ("Pesticides",      "Total fongicides + insecticides + herbicides (litres)"),
        ("",""),
        ("FORMULE : Rendement",       "= Tonnage (t) ÷ Hectares → t/ha (cible: >35 t/ha)"),
        ("FORMULE : Dose DAP/ha",     "= DAP (kg) ÷ Hectares → kg/ha (optimal: 150 kg/ha)"),
        ("FORMULE : Dose fumier/ha",  "= Fumier (kg) ÷ Hectares → t/ha (optimal: 20 t/ha)"),
        ("FORMULE : Efficacité",      "= (Tonnage×1000) ÷ (DAP+Fumure) → kg tomate/kg engrais"),
        ("FORMULE : Rendement/plant", "= Tonnage(kg) ÷ (Densité × Ha) → kg par plant"),
        ("",""),
        ("Densité plants CAP BON",    "25 000 plants/ha"),
        ("Densité plants KAIROUAN",   "20 000 plants/ha"),
        ("Densité plants GAFSA",      "18 000 plants/ha"),
        ("Rendement moy. CAP BON 1",  "45 t/ha (référence)"),
        ("Rendement moy. KAIROUAN",   "30 t/ha (référence)"),
        ("Rendement moy. GAFSA",      "28 t/ha (référence)"),
    ]
    for ri, (a, b) in enumerate(legend, 1):
        ws2.cell(ri, 1).value = a
        ws2.cell(ri, 2).value = b
        if ri == 1:
            for ci in [1,2]:
                ws2.cell(ri, ci).fill = HDR_FILL
                ws2.cell(ri, ci).font = HDR_FONT

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


# ══════════════════════════════════════════════════════════════
# PARSING DU FICHIER D'IMPORT
# ══════════════════════════════════════════════════════════════
def parse_performance_file(file_obj):
    """Parse le fichier Excel d'import et retourne un DataFrame enrichi."""
    try:
        df = pd.read_excel(file_obj, sheet_name=0, header=0)
    except Exception as e:
        return None, str(e)

    # Normaliser les colonnes
    RENAME = {
        "commercial":           "commercial",
        "ingénieur":            "ingenieur",
        "agriculteur":          "agriculteur",
        "région":               "region",
        "variété":              "variete",
        "hectares (ha)":        "hectares",
        "tonnage récolté (t)":  "tonnage_t",
        "dap (kg total)":       "dap_kg",
        "fumure (kg total)":    "fumure_kg",
        "fumier (kg total)":    "fumier_kg",
        "pesticides (l total)": "pesticide_l",
        "fongicides (l total)": "fongicide_l",
        "insecticides (l total)":"insecticide_l",
        "irrigation (m³)":      "irrigation_m3",
        "notes":                "notes",
    }
    df.columns = [str(c).strip().lower() for c in df.columns]
    df = df.rename(columns={k: v for k, v in RENAME.items() if k in df.columns})

    # Colonnes numériques
    NUM_COLS = ["hectares","tonnage_t","dap_kg","fumure_kg","fumier_kg",
                "pesticide_l","fongicide_l","insecticide_l","irrigation_m3"]
    for col in NUM_COLS:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    # Filtrer lignes invalides
    if "agriculteur" in df.columns and "tonnage_t" in df.columns:
        df = df[df["agriculteur"].notna() & (df["agriculteur"] != "")]
        df = df[df["tonnage_t"].notna() & (df["tonnage_t"] > 0)]

    if df.empty:
        return None, "Aucune ligne valide trouvée dans le fichier."

    # Calculs automatiques
    df["rendement_t_ha"] = df.apply(
        lambda r: calculer_rendement(r.get("tonnage_t", 0), r.get("hectares")), axis=1
    )
    df["rendement_kg_plant"] = df.apply(
        lambda r: calculer_rendement_par_plant(
            r.get("tonnage_t", 0) * 1000,
            r.get("hectares"),
            r.get("region")
        ), axis=1
    )
    df["dap_kg_ha"] = df.apply(
        lambda r: calculer_dose_ha(r.get("dap_kg"), r.get("hectares")), axis=1
    )
    df["fumure_kg_ha"] = df.apply(
        lambda r: calculer_dose_ha(r.get("fumure_kg"), r.get("hectares")), axis=1
    )
    df["fumier_t_ha"] = df.apply(
        lambda r: calculer_dose_ha(r.get("fumier_kg"), r.get("hectares"), "kg")
            and round(calculer_dose_ha(r.get("fumier_kg"), r.get("hectares")) / 1000, 2)
            if r.get("fumier_kg") else None, axis=1
    )
    df["pesticide_l_ha"] = df.apply(
        lambda r: calculer_dose_ha(r.get("pesticide_l"), r.get("hectares"), "L"), axis=1
    )
    total_engrais = df.get("dap_kg", pd.Series(dtype=float)).fillna(0) + \
                    df.get("fumure_kg", pd.Series(dtype=float)).fillna(0)
    df["efficacite"] = df.apply(
        lambda r: calculer_efficacite_intrants(
            r.get("tonnage_t", 0),
            r.get("dap_kg", 0) + r.get("fumure_kg", 0)
        ), axis=1
    )
    df["indice_dap"] = df.apply(
        lambda r: calculer_indice_intrant(r.get("dap_kg_ha"), REF_DOSES["DAP"]["optimal"]),
        axis=1
    )

    # Score et label
    results = df.apply(
        lambda r: get_score_label(r.get("rendement_t_ha") or 0, r.get("region", "")),
        axis=1
    )
    df["score_label"] = results.apply(lambda x: x[0])
    df["score_color"] = results.apply(lambda x: x[1])
    df["score_ratio"] = results.apply(lambda x: x[2])

    return df, None


# ══════════════════════════════════════════════════════════════
# SAUVEGARDE / CHARGEMENT SUPABASE
# ══════════════════════════════════════════════════════════════
def save_to_supabase(sb, df, saison="2026"):
    """Sauvegarde le DataFrame dans la table agri_performance."""
    if sb is None or df is None or df.empty:
        return False
    try:
        # Supprimer les anciennes données de la même saison/commercial
        if "commercial" in df.columns:
            comms = df["commercial"].dropna().unique().tolist()
            for comm in comms:
                sb.table("agri_performance").delete()\
                  .eq("commercial", comm).eq("saison", saison).execute()

        # Insérer
        COLS_SB = ["commercial","ingenieur","agriculteur","region","variete",
                   "hectares","tonnage_t","dap_kg","fumure_kg","fumier_kg",
                   "pesticide_l","fongicide_l","insecticide_l","irrigation_m3",
                   "notes","saison"]
        rows = []
        for _, row in df.iterrows():
            record = {"saison": saison}
            for col in COLS_SB:
                val = row.get(col)
                if pd.isna(val) if val is not None else True:
                    record[col] = None
                else:
                    record[col] = val
            rows.append(record)
        sb.table("agri_performance").insert(rows).execute()
        return True
    except Exception as e:
        st.error(f"Erreur Supabase save: {e}")
        return False


def load_from_supabase(sb, saison="2026"):
    """Charge les données depuis agri_performance et recalcule."""
    if sb is None:
        return pd.DataFrame()
    try:
        data = sb.table("agri_performance").select("*")\
                 .eq("saison", saison).execute().data
        if not data:
            return pd.DataFrame()
        df = pd.DataFrame(data)
        # Recalculer colonnes dérivées
        df["rendement_t_ha"] = df.apply(
            lambda r: calculer_rendement(r.get("tonnage_t") or 0, r.get("hectares")), axis=1
        )
        df["efficacite"] = df.apply(
            lambda r: calculer_efficacite_intrants(
                r.get("tonnage_t") or 0,
                (r.get("dap_kg") or 0) + (r.get("fumure_kg") or 0)
            ), axis=1
        )
        results = df.apply(
            lambda r: get_score_label(r.get("rendement_t_ha") or 0, r.get("region", "")),
            axis=1
        )
        df["score_label"] = results.apply(lambda x: x[0])
        df["score_color"] = results.apply(lambda x: x[1])
        df["score_ratio"] = results.apply(lambda x: x[2])
        return df
    except Exception as e:
        return pd.DataFrame()


# ══════════════════════════════════════════════════════════════
# RENDU PRINCIPAL : render_agro_tab()
# ══════════════════════════════════════════════════════════════
def render_agro_tab(sb=None, planning_df=None, CURRENT_ROLE="directeur", CURRENT_NAME=""):
    """Onglet Performance Agronomique — à appeler dans dashboard_phase10.py."""

    # ── En-tête ──────────────────────────────────────────────
    st.markdown("""
<div style='background:#0d2b0d;border:1px solid #1E8449;border-radius:12px;
padding:16px 20px;margin-bottom:20px'>
  <div style='font-size:1.1rem;font-weight:700;color:#f0f6fc;margin-bottom:6px'>
    🌱 Performance Agronomique — Tomate 2026
  </div>
  <div style='font-size:.82rem;color:#8b949e'>
    Analyse rendement (t/ha) • Efficacité des intrants • Classement commerciaux •
    Recommandations par région et variété
  </div>
</div>""", unsafe_allow_html=True)

    # ── Session state ──────────────────────────────────────
    if "agro_df" not in st.session_state:
        # Essayer de charger depuis Supabase au démarrage
        df_sb = load_from_supabase(sb)
        st.session_state["agro_df"] = df_sb if not df_sb.empty else pd.DataFrame()

    # ── Tabs internes ──────────────────────────────────────
    tab_saisie, tab_perf, tab_rank, tab_best, tab_sql = st.tabs([
        "📥 Saisie des données",
        "📊 Performance agriculteurs",
        "🏆 Classement commerciaux",
        "🔬 Meilleures pratiques",
        "⚙️ SQL / Export",
    ])

    # ══════════════════════════════════════════════════════
    # TAB 1 : SAISIE
    # ══════════════════════════════════════════════════════
    with tab_saisie:
        st.markdown("### Import des données agronomiques")
        st.caption(
            "Téléchargez le template, remplissez-le avec vos ingénieurs, "
            "puis importez-le ici. Les calculs sont automatiques."
        )

        col_dl, col_up = st.columns([1, 2])

        with col_dl:
            st.download_button(
                "📥 Télécharger le template Excel",
                data=generate_import_template(),
                file_name="agro_performance_template_2026.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True,
            )
            st.caption("Remplissez 1 ligne par agriculteur")

            # Saisie manuelle rapide
            st.markdown("---")
            st.markdown("**Ou saisie manuelle rapide :**")
            with st.form("agro_form"):
                f_comm    = st.selectbox("Commercial", ["FEDI","MAKKI BEN SALAH","KHALIL","ACHREF AJLANI","JILANI OBAY"])
                f_ing     = st.text_input("Ingénieur agronome")
                f_agri    = st.text_input("Nom agriculteur *")
                f_reg     = st.selectbox("Région", list(RENDEMENT_MOYEN_REGION.keys()))
                f_var     = st.text_input("Variété (ex: Heinz, Rio Grande)")
                c1, c2    = st.columns(2)
                f_ha      = c1.number_input("Hectares *", min_value=0.1, step=0.5)
                f_ton     = c2.number_input("Tonnage récolté (t) *", min_value=0.0, step=1.0)
                c3, c4    = st.columns(2)
                f_dap     = c3.number_input("DAP (kg total)", min_value=0.0, step=50.0)
                f_fumure  = c4.number_input("Fumure (kg total)", min_value=0.0, step=100.0)
                c5, c6    = st.columns(2)
                f_fumier  = c5.number_input("Fumier (kg total)", min_value=0.0, step=1000.0)
                f_pest    = c6.number_input("Pesticides (L total)", min_value=0.0, step=1.0)
                f_notes   = st.text_area("Notes", height=60)

                submitted = st.form_submit_button("➕ Ajouter cet agriculteur", type="primary")
                if submitted and f_agri and f_ha > 0 and f_ton > 0:
                    new_row = {
                        "commercial": f_comm, "ingenieur": f_ing,
                        "agriculteur": f_agri, "region": f_reg,
                        "variete": f_var, "hectares": f_ha,
                        "tonnage_t": f_ton, "dap_kg": f_dap or None,
                        "fumure_kg": f_fumure or None, "fumier_kg": f_fumier or None,
                        "pesticide_l": f_pest or None, "notes": f_notes,
                    }
                    df_new = parse_performance_file.__wrapped__(pd.DataFrame([new_row])) \
                             if hasattr(parse_performance_file, "__wrapped__") \
                             else pd.DataFrame([new_row])
                    # Recalcul simple
                    nr = new_row.copy()
                    nr["rendement_t_ha"]   = calculer_rendement(f_ton, f_ha)
                    nr["efficacite"]       = calculer_efficacite_intrants(f_ton, (f_dap or 0) + (f_fumure or 0))
                    nr["dap_kg_ha"]        = calculer_dose_ha(f_dap, f_ha)
                    lbl, col, ratio        = get_score_label(nr["rendement_t_ha"] or 0, f_reg)
                    nr["score_label"]      = lbl
                    nr["score_color"]      = col
                    nr["score_ratio"]      = ratio
                    df_cur = st.session_state.get("agro_df", pd.DataFrame())
                    st.session_state["agro_df"] = pd.concat(
                        [df_cur, pd.DataFrame([nr])], ignore_index=True
                    )
                    st.success(f"✅ {f_agri} ajouté — rendement: {nr['rendement_t_ha']} t/ha {lbl}")

        with col_up:
            uploaded = st.file_uploader(
                "Importer un fichier Excel complété",
                type=["xlsx","xls"],
                key="agro_upload",
            )
            if uploaded:
                df_parsed, err = parse_performance_file(uploaded)
                if err:
                    st.error(f"Erreur: {err}")
                elif df_parsed is not None and not df_parsed.empty:
                    st.session_state["agro_df"] = df_parsed
                    ok = save_to_supabase(sb, df_parsed)
                    st.success(
                        f"✅ {len(df_parsed)} agriculteurs importés"
                        f" {'+ sauvegardés Supabase' if ok else '(session uniquement)'}"
                    )

            df_cur = st.session_state.get("agro_df", pd.DataFrame())
            if not df_cur.empty:
                st.markdown(f"**{len(df_cur)} agriculteurs chargés :**")
                st.dataframe(
                    df_cur[["commercial","agriculteur","region","hectares","tonnage_t",
                             "rendement_t_ha","score_label"]].rename(columns={
                        "commercial":"Commercial","agriculteur":"Agriculteur",
                        "region":"Région","hectares":"Ha","tonnage_t":"Tonnage (t)",
                        "rendement_t_ha":"Rendement (t/ha)","score_label":"Score",
                    }),
                    use_container_width=True, hide_index=True, height=320,
                )
                if st.button("🗑️ Effacer toutes les données session", use_container_width=True):
                    st.session_state["agro_df"] = pd.DataFrame()
                    st.rerun()

    # ══════════════════════════════════════════════════════
    # TAB 2 : PERFORMANCE AGRICULTEURS
    # ══════════════════════════════════════════════════════
    with tab_perf:
        df = st.session_state.get("agro_df", pd.DataFrame())
        if df.empty:
            st.info("📥 Importez des données dans l'onglet 'Saisie' pour voir les analyses.")
        else:
            # Filtres
            fc1, fc2, fc3 = st.columns(3)
            sel_comm = fc1.selectbox("Commercial", ["Tous"] + sorted(df["commercial"].dropna().unique().tolist()), key="ap_comm")
            sel_reg  = fc2.selectbox("Région", ["Toutes"] + sorted(df["region"].dropna().unique().tolist()), key="ap_reg")
            sel_var  = fc3.selectbox("Variété", ["Toutes"] + sorted(df.get("variete", pd.Series()).dropna().unique().tolist()), key="ap_var") if "variete" in df.columns else "Toutes"

            df_f = df.copy()
            if sel_comm != "Tous":     df_f = df_f[df_f["commercial"] == sel_comm]
            if sel_reg  != "Toutes":   df_f = df_f[df_f["region"]     == sel_reg]
            if sel_var  != "Toutes" and "variete" in df_f.columns:
                df_f = df_f[df_f["variete"] == sel_var]

            if df_f.empty:
                st.warning("Aucune donnée pour cette sélection.")
            else:
                # ── KPIs globaux ─────────────────────────────
                k1, k2, k3, k4, k5 = st.columns(5)
                k1.metric("Agriculteurs",   len(df_f))
                k2.metric("Total ha",       f"{df_f['hectares'].sum():.1f} ha")
                k3.metric("Total tonnage",  f"{df_f['tonnage_t'].sum():.0f} t")
                avg_rend = df_f["rendement_t_ha"].mean()
                ref_moy  = RENDEMENT_MOYEN_NATIONAL
                k4.metric("Rendement moy.", f"{avg_rend:.1f} t/ha",
                          delta=f"{avg_rend - ref_moy:+.1f} vs nat.",
                          delta_color="normal" if avg_rend >= ref_moy else "inverse")
                k5.metric("Efficacité moy.", f"{df_f['efficacite'].mean():.2f}" if "efficacite" in df_f else "—")

                # ── Graphique rendement par agriculteur ──────
                st.markdown("#### 🌾 Rendement (t/ha) par agriculteur")
                df_sort = df_f.dropna(subset=["rendement_t_ha"]).sort_values("rendement_t_ha", ascending=True)
                if not df_sort.empty:
                    ref_reg = df_sort["region"].map(lambda r: RENDEMENT_MOYEN_REGION.get(r, RENDEMENT_MOYEN_NATIONAL))

                    fig_r = go.Figure()
                    fig_r.add_trace(go.Bar(
                        y=df_sort["agriculteur"],
                        x=df_sort["rendement_t_ha"],
                        orientation="h",
                        marker_color=df_sort["score_color"],
                        text=df_sort.apply(
                            lambda r: f"{r['rendement_t_ha']:.1f} t/ha  {r['score_label']}",
                            axis=1
                        ),
                        textposition="outside",
                        textfont=dict(size=11),
                    ))
                    # Ligne référence nationale
                    fig_r.add_vline(
                        x=RENDEMENT_MOYEN_NATIONAL,
                        line_dash="dash", line_color="#e8543a", line_width=1.5,
                        annotation_text=f"Moy. nat. {RENDEMENT_MOYEN_NATIONAL}t/ha",
                        annotation_position="top",
                        annotation_font_color="#e8543a",
                    )
                    fig_r.update_layout(
                        template="plotly_dark", paper_bgcolor="#161b22", plot_bgcolor="#0d1117",
                        height=max(350, len(df_sort) * 30 + 100),
                        xaxis_title="Rendement (t/ha)",
                        margin=dict(l=200, r=120, t=40, b=40),
                        showlegend=False,
                    )
                    st.plotly_chart(fig_r, use_container_width=True)

                # ── Scatter intrants vs rendement ──────────
                st.markdown("#### 💊 Efficacité intrants : DAP vs Rendement")
                st.caption("Objectif : faible dose, haut rendement → bulle en haut à gauche")
                df_eff = df_f.dropna(subset=["rendement_t_ha","dap_kg_ha"])
                if not df_eff.empty:
                    fig_s = px.scatter(
                        df_eff,
                        x="dap_kg_ha", y="rendement_t_ha",
                        color="score_label",
                        color_discrete_map={k: v for (_, _), (k, v) in SCORE_LABELS.items()},
                        size="tonnage_t" if "tonnage_t" in df_eff.columns else None,
                        hover_data=["agriculteur","commercial","region","variete"]
                                   if "variete" in df_eff.columns
                                   else ["agriculteur","commercial","region"],
                        labels={"dap_kg_ha":"DAP (kg/ha)","rendement_t_ha":"Rendement (t/ha)"},
                        title="DAP consommé vs. Rendement obtenu",
                        template="plotly_dark",
                    )
                    # Quadrants
                    ref_dap = REF_DOSES["DAP"]["optimal"]
                    fig_s.add_vline(x=ref_dap, line_dash="dot", line_color="#8b949e", line_width=1)
                    fig_s.add_hline(y=RENDEMENT_MOYEN_NATIONAL, line_dash="dot", line_color="#8b949e", line_width=1)
                    fig_s.update_layout(paper_bgcolor="#161b22", plot_bgcolor="#0d1117", height=420)
                    st.plotly_chart(fig_s, use_container_width=True)

                # ── Tableau détaillé ─────────────────────────
                st.markdown("#### 📋 Tableau complet")
                disp_cols = ["commercial","ingenieur","agriculteur","region","variete",
                             "hectares","tonnage_t","rendement_t_ha","rendement_kg_plant",
                             "dap_kg_ha","fumure_kg_ha","pesticide_l_ha",
                             "efficacite","indice_dap","score_label"]
                disp_cols = [c for c in disp_cols if c in df_f.columns]
                df_disp = df_f[disp_cols].rename(columns={
                    "commercial":"Commercial","ingenieur":"Ingénieur",
                    "agriculteur":"Agriculteur","region":"Région","variete":"Variété",
                    "hectares":"Ha","tonnage_t":"T (t)","rendement_t_ha":"Rend. t/ha",
                    "rendement_kg_plant":"kg/plant","dap_kg_ha":"DAP kg/ha",
                    "fumure_kg_ha":"Fumure kg/ha","pesticide_l_ha":"Pest. L/ha",
                    "efficacite":"Efficacité","indice_dap":"Indice DAP",
                    "score_label":"Score",
                })
                st.dataframe(df_disp.sort_values("Rend. t/ha", ascending=False),
                             use_container_width=True, hide_index=True, height=400)

    # ══════════════════════════════════════════════════════
    # TAB 3 : CLASSEMENT COMMERCIAUX
    # ══════════════════════════════════════════════════════
    with tab_rank:
        df = st.session_state.get("agro_df", pd.DataFrame())
        if df.empty:
            st.info("Données non disponibles.")
        else:
            st.markdown("### 🏆 Classement agronomique des commerciaux")
            st.caption(
                "Score = moyenne pondérée (par tonnage) des ratios Rendement/Référence_Région "
                "de chaque agriculteur. Score > 1.0 = au-dessus de la moyenne régionale."
            )

            rows = []
            for comm in sorted(df["commercial"].dropna().unique()):
                df_comm = df[df["commercial"] == comm]
                sc = score_commercial(df_comm)
                n_agri  = len(df_comm)
                avg_r   = df_comm["rendement_t_ha"].mean()
                tot_t   = df_comm["tonnage_t"].sum()
                tot_ha  = df_comm["hectares"].sum()
                pct_exc = len(df_comm[df_comm["score_ratio"] >= 1.15]) / n_agri * 100 if n_agri else 0
                pct_sub = len(df_comm[df_comm["score_ratio"] < 0.75]) / n_agri * 100 if n_agri else 0
                ing     = df_comm["ingenieur"].dropna().iloc[0] if "ingenieur" in df_comm and not df_comm["ingenieur"].dropna().empty else "—"
                rows.append({
                    "Commercial": comm,
                    "Ingénieur": ing,
                    "Score agro": sc,
                    "Agri.": n_agri,
                    "Rend. moy. (t/ha)": round(avg_r, 1) if not pd.isna(avg_r) else 0,
                    "Total (t)": int(tot_t),
                    "Total (ha)": round(tot_ha, 1),
                    "% Excellents": round(pct_exc, 1),
                    "% Sous-perf.": round(pct_sub, 1),
                })
            df_rank = pd.DataFrame(rows).sort_values("Score agro", ascending=False).reset_index(drop=True)
            df_rank.index += 1  # classement 1-based

            # Médailles
            medals = ["🥇", "🥈", "🥉"] + [""] * 10
            df_rank.insert(0, "🏅", medals[:len(df_rank)])

            st.dataframe(
                df_rank,
                use_container_width=True, hide_index=False,
                column_config={
                    "Score agro": st.column_config.ProgressColumn(
                        "Score agro", min_value=0, max_value=1.5, format="%.3f"
                    ),
                    "% Excellents": st.column_config.ProgressColumn(
                        "% Excellents", min_value=0, max_value=100, format="%.0f%%"
                    ),
                    "% Sous-perf.": st.column_config.NumberColumn(
                        "% Sous-perf.", format="%.0f%%"
                    ),
                }
            )

            # Graphique radar
            st.markdown("---")
            st.markdown("#### 🕸️ Profil agronomique par commercial")
            METRICS = ["Rend. moy. (t/ha)", "Score agro", "% Excellents"]
            fig_radar = go.Figure()
            COLORS_COMM = {
                "FEDI":"#1A5276","MAKKI BEN SALAH":"#1F7A1F",
                "KHALIL":"#7D3C98","ACHREF AJLANI":"#C0392B","JILANI OBAY":"#D4AC0D",
            }
            for _, row in df_rank.iterrows():
                fig_radar.add_trace(go.Scatterpolar(
                    r=[
                        min(row["Rend. moy. (t/ha)"] / 50, 1.0),
                        min(row["Score agro"], 1.5) / 1.5,
                        row["% Excellents"] / 100,
                    ],
                    theta=["Rendement", "Score agro", "% Excellents"],
                    fill="toself",
                    name=row["Commercial"],
                    line_color=COLORS_COMM.get(row["Commercial"], "#888"),
                ))
            fig_radar.update_layout(
                polar=dict(bgcolor="#161b22"),
                template="plotly_dark", paper_bgcolor="#161b22",
                height=420,
            )
            st.plotly_chart(fig_radar, use_container_width=True)

    # ══════════════════════════════════════════════════════
    # TAB 4 : MEILLEURES PRATIQUES
    # ══════════════════════════════════════════════════════
    with tab_best:
        df = st.session_state.get("agro_df", pd.DataFrame())
        if df.empty:
            st.info("Données non disponibles.")
        else:
            st.markdown("### 🔬 Meilleures pratiques par région et variété")
            st.caption(
                "Analyse des agriculteurs ⭐ Excellents (rendement > 115% de la moyenne régionale) "
                "pour identifier les doses d'engrais et méthodes les plus efficaces."
            )

            # Sélection région
            sel_reg2 = st.selectbox(
                "Choisir une région",
                sorted(df["region"].dropna().unique()),
                key="bp_reg"
            )
            df_reg = df[df["region"] == sel_reg2]
            ref_reg = RENDEMENT_MOYEN_REGION.get(sel_reg2, RENDEMENT_MOYEN_NATIONAL)

            st.markdown(f"**Référence région {sel_reg2} :** {ref_reg} t/ha")

            if not df_reg.empty:
                # Seuil excellent
                df_exc = df_reg[df_reg["score_ratio"] >= 1.05].copy()
                df_avg = df_reg.copy()

                def _fmt_mean(series):
                    v = series.dropna().mean()
                    return f"{v:.1f}" if not pd.isna(v) else "—"

                metrics_bp = {
                    "Nb agriculteurs":         [len(df_exc), len(df_avg)],
                    "Rendement moy. (t/ha)":   [_fmt_mean(df_exc["rendement_t_ha"]),   _fmt_mean(df_avg["rendement_t_ha"])],
                    "DAP (kg/ha)":             [_fmt_mean(df_exc.get("dap_kg_ha",    pd.Series())), _fmt_mean(df_avg.get("dap_kg_ha",    pd.Series()))],
                    "Fumure (kg/ha)":          [_fmt_mean(df_exc.get("fumure_kg_ha",  pd.Series())), _fmt_mean(df_avg.get("fumure_kg_ha",  pd.Series()))],
                    "Pesticides (L/ha)":       [_fmt_mean(df_exc.get("pesticide_l_ha",pd.Series())), _fmt_mean(df_avg.get("pesticide_l_ha",pd.Series()))],
                    "Efficacité (kg/kg)":      [_fmt_mean(df_exc.get("efficacite",    pd.Series())), _fmt_mean(df_avg.get("efficacite",    pd.Series()))],
                }
                df_bp = pd.DataFrame(metrics_bp, index=["✅ Excellents (>105%)", "📊 Moyenne région"]).T
                df_bp.index.name = "Indicateur"
                st.dataframe(df_bp, use_container_width=True)

                # Recommandation automatique
                st.markdown("#### 💡 Recommandations pour cette région")
                recs = []
                avg_dap = df_exc["dap_kg_ha"].dropna().mean() if "dap_kg_ha" in df_exc else None
                avg_fum = df_exc["fumure_kg_ha"].dropna().mean() if "fumure_kg_ha" in df_exc else None
                avg_pest = df_exc["pesticide_l_ha"].dropna().mean() if "pesticide_l_ha" in df_exc else None

                if avg_dap:
                    recs.append(f"**DAP optimal :** {avg_dap:.0f} kg/ha "
                                f"(référence nationale : {REF_DOSES['DAP']['optimal']} kg/ha)")
                if avg_fum:
                    recs.append(f"**Fumure organique :** {avg_fum:.0f} kg/ha")
                if avg_pest:
                    recs.append(f"**Pesticides :** {avg_pest:.2f} L/ha "
                                f"— {'🟢 sous la limite' if avg_pest < REF_DOSES['PESTICIDE']['optimal'] else '🟡 surveiller'}")

                recs.append(f"**Rendement cible {sel_reg2} :** {ref_reg} t/ha (régionale) "
                            f"→ viser {ref_reg * 1.15:.0f} t/ha pour niveau Excellent")

                for r in recs:
                    st.markdown(f"• {r}")

            # Comparaison variétés
            st.markdown("---")
            st.markdown("#### 🍅 Comparaison par variété")
            if "variete" in df.columns:
                df_var_grp = df.groupby("variete").agg(
                    n=("agriculteur", "count"),
                    rendement_moy=("rendement_t_ha", "mean"),
                    efficacite_moy=("efficacite", "mean"),
                    dap_moy=("dap_kg_ha", "mean"),
                ).reset_index().dropna(subset=["rendement_moy"])
                df_var_grp.columns = ["Variété","Nb agri.","Rend. moy. (t/ha)","Efficacité moy.","DAP moy. (kg/ha)"]
                df_var_grp = df_var_grp.sort_values("Rend. moy. (t/ha)", ascending=False)
                st.dataframe(df_var_grp, use_container_width=True, hide_index=True)
            else:
                st.info("Ajoutez la colonne 'variete' dans votre fichier pour cette analyse.")

    # ══════════════════════════════════════════════════════
    # TAB 5 : SQL / EXPORT
    # ══════════════════════════════════════════════════════
    with tab_sql:
        st.markdown("### ⚙️ SQL Supabase — Créer la table")
        st.code(SQL_CREATE_TABLE, language="sql")
        st.caption("Exécutez ce SQL une seule fois dans l'éditeur SQL de Supabase.")

        # Export Excel
        df = st.session_state.get("agro_df", pd.DataFrame())
        if not df.empty:
            st.markdown("---")
            st.markdown("### 📤 Exporter les données calculées")
            buf = io.BytesIO()
            with pd.ExcelWriter(buf, engine="openpyxl") as writer:
                df.to_excel(writer, sheet_name="Performance", index=False)
            buf.seek(0)
            st.download_button(
                "📥 Exporter Performance Agronomique (Excel)",
                data=buf.getvalue(),
                file_name="agro_performance_2026.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
            )