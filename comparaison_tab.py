# -*- coding: utf-8 -*-
"""
comparaison_tab.py — Onglet Comparaison OR-Tools vs Plan Rectifié
=================================================================
À importer dans dashboard_phase10.py :

    from comparaison_tab import render_comparaison_tab

Usage dans le dashboard (ajouter un tab) :
    tab_comp = st.tabs([..., "📊 Comparaison Plans"])
    with tab_comp:
        render_comparaison_tab(planning_df=planning, df_to_xlsx_styled=df_to_xlsx_styled)
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
import io


# ── Palette couleurs ─────────────────────────────────────────
COMM_COLORS = {
    "ACHREF AJLANI":   "#8b5cf6",
    "FEDI":            "#3b82f6",
    "JILANI OBAY":     "#e8543a",
    "KHALIL":          "#f5a623",
    "MAKKI BEN SALAH": "#00e5a0",
}
PLAN_ORTOOLS_COLOR = "#3b82f6"   # bleu = OR-Tools
PLAN_MANUEL_COLOR  = "#f5a623"   # orange = plan manuel rectifié


def _parse_rectification_file(uploaded_file) -> pd.DataFrame:
    """
    Parse un fichier de rectification (format Planning Horizontal).
    Format attendu :
      Ligne 0 = en-têtes : Commercial/Responsable | Région | Agriculteur | Tonnage | Accessibilité | Caisses Vides | 2026-06-20 | ...
      Lignes suivantes = données agriculteurs

    Retourne un DataFrame long avec colonnes :
      commercial, agriculteur, tonnage_total, accessibilite, region, date, tonnes_jour
    """
    try:
        df_raw = pd.read_excel(uploaded_file, sheet_name=0, header=0)
    except Exception as e:
        st.error(f"Erreur lecture fichier: {e}")
        return pd.DataFrame()

    # Détecter les colonnes de dates (à partir de la colonne 6)
    date_cols_raw = df_raw.columns[6:]
    parsed_dates = []
    for c in date_cols_raw:
        try:
            d = pd.to_datetime(str(c).split(' ')[0], errors='coerce')
            if pd.notna(d) and d.year >= 2026:
                parsed_dates.append((c, d.date()))
        except Exception:
            pass

    if not parsed_dates:
        st.warning("Aucune colonne de date trouvée dans ce fichier.")
        return pd.DataFrame()

    rows = []
    for _, row in df_raw.iterrows():
        # Colonnes fixes
        comm    = str(row.iloc[0]).strip()
        region  = str(row.iloc[1]).strip()
        agri    = str(row.iloc[2]).strip()
        tonnage = pd.to_numeric(row.iloc[3], errors='coerce')
        acc     = str(row.iloc[4]).strip()

        # Ignorer lignes vides ou totaux
        if pd.isna(tonnage) or tonnage <= 0:
            continue
        if agri.upper() in ('NAN', '', 'AGRICULTEUR', 'TOTAL', 'SOUS-TOTAL'):
            continue
        if comm.upper() in ('NAN', '', 'COMMERCIAL', 'RESPONSABLE RÉGIONAL'):
            continue

        # Extraire les tonnages journaliers
        for orig_col, date in parsed_dates:
            val = pd.to_numeric(row[orig_col], errors='coerce')
            if pd.notna(val) and val > 0:
                rows.append({
                    'commercial':    comm,
                    'region':        region,
                    'agriculteur':   agri,
                    'tonnage_total': float(tonnage),
                    'accessibilite': acc,
                    'date':          pd.Timestamp(date),
                    'tonnes_jour':   float(val),
                    'source':        'Manuel rectifié',
                })

    return pd.DataFrame(rows)


def _extract_ortools_for_comparison(planning_df: pd.DataFrame, commercial: str) -> pd.DataFrame:
    """
    Extrait les données OR-Tools pour un commercial donné depuis le planning Supabase.
    Retourne un DataFrame long similaire au fichier rectifié.
    """
    if planning_df is None or planning_df.empty:
        return pd.DataFrame()

    # Chercher la colonne Commercial
    comm_col = next((c for c in planning_df.columns if c.lower() == 'commercial'), None)
    if not comm_col:
        return pd.DataFrame()

    sub = planning_df[planning_df[comm_col] == commercial].copy()
    if sub.empty:
        return pd.DataFrame()

    # Normaliser les colonnes
    col_map = {
        'agriculteur': next((c for c in sub.columns if c.lower() == 'agriculteur'), None),
        'date':        next((c for c in sub.columns if c.lower() == 'date'), None),
        'tonnes_jour': next((c for c in sub.columns if 'tonne' in c.lower() and 'jour' in c.lower()), None),
        'tonnage':     next((c for c in sub.columns if c.lower() == 'total tonnes' or 'total' in c.lower() and 'tonne' in c.lower()), None),
        'access':      next((c for c in sub.columns if 'access' in c.lower()), None),
        'region':      next((c for c in sub.columns if 'region' in c.lower() or 'région' in c.lower()), None),
    }

    if not col_map['date'] or not col_map['tonnes_jour']:
        return pd.DataFrame()

    result = pd.DataFrame({
        'commercial':    commercial,
        'agriculteur':   sub[col_map['agriculteur']] if col_map['agriculteur'] else '',
        'tonnage_total': pd.to_numeric(sub[col_map['tonnage']], errors='coerce') if col_map['tonnage'] else 0,
        'accessibilite': sub[col_map['access']].fillna('') if col_map['access'] else '',
        'region':        sub[col_map['region']].fillna('') if col_map['region'] else '',
        'date':          pd.to_datetime(sub[col_map['date']], errors='coerce'),
        'tonnes_jour':   pd.to_numeric(sub[col_map['tonnes_jour']], errors='coerce').fillna(0),
        'source':        'OR-Tools',
    })

    return result.dropna(subset=['date']).reset_index(drop=True)


def _compute_comparison_stats(df_ortools: pd.DataFrame, df_manuel: pd.DataFrame) -> dict:
    """
    Calcule les statistiques de comparaison entre OR-Tools et plan manuel.
    """
    stats = {}

    for label, df in [('OR-Tools', df_ortools), ('Manuel', df_manuel)]:
        if df.empty:
            stats[label] = {}
            continue

        daily = df.groupby('date')['tonnes_jour'].sum()
        agri_profiles = df.groupby('agriculteur').apply(
            lambda g: {
                'n_days':   len(g),
                'max':      g['tonnes_jour'].max(),
                'avg':      g['tonnes_jour'].mean(),
                'total':    g['tonnes_jour'].sum(),
            }
        )

        stats[label] = {
            'total_tonnes':       round(df['tonnes_jour'].sum(), 0),
            'max_jour':           round(daily.max(), 0),
            'avg_jour':           round(daily.mean(), 0),
            'n_jours_actifs':     len(daily[daily > 0]),
            'n_agriculteurs':     df['agriculteur'].nunique(),
            'avg_days_per_agri':  round(agri_profiles.apply(lambda x: x['n_days']).mean(), 1),
            'avg_max_per_agri':   round(agri_profiles.apply(lambda x: x['max']).mean(), 1),
            'avg_avg_per_agri':   round(agri_profiles.apply(lambda x: x['avg']).mean(), 1),
        }

    return stats


def render_comparaison_tab(planning_df: pd.DataFrame, df_to_xlsx_styled=None):
    """
    Render l'onglet de comparaison OR-Tools vs Plan Rectifié.
    
    Args:
        planning_df: DataFrame du planning OR-Tools (depuis Supabase)
        df_to_xlsx_styled: fonction d'export Excel (optionnelle)
    """
    st.subheader("📊 Comparaison OR-Tools vs Plan Rectifié Manuel")
    st.caption(
        "Uploadez les fichiers Excel rectifiés par les commerciaux pour comparer "
        "avec le planning calculé par OR-Tools. Identifie les écarts et explique "
        "les causes racines."
    )

    # ── Session state pour stocker les plans uploadés ────────────
    if 'plans_rectifies' not in st.session_state:
        st.session_state['plans_rectifies'] = {}

    # ── Zone d'upload ────────────────────────────────────────────
    st.markdown("### 1️⃣ Uploader les plans rectifiés")

    COMM_LABELS = {
        "FEDI":            "FEDI",
        "MAKKI BEN SALAH": "MAKKI",
        "KHALIL":          "KHALIL",
        "ACHREF AJLANI":   "ACHREF",
        "JILANI OBAY":     "JILANI",
    }

    col_up1, col_up2 = st.columns([2, 1])
    with col_up1:
        uploaded_files = st.file_uploader(
            "Déposer 1 ou plusieurs fichiers de rectification (Excel)",
            type=["xlsx", "xls"],
            accept_multiple_files=True,
            key="upload_rectif",
            help="Format attendu: Planning Horizontal par commercial (colonnes = dates)"
        )

    with col_up2:
        st.markdown("**Fichiers attendus :**")
        for comm in COMM_LABELS:
            stored = comm in st.session_state['plans_rectifies']
            icon = "✅" if stored else "⬜"
            st.markdown(f"{icon} {comm}")

    # Parser les fichiers uploadés
    if uploaded_files:
        for f in uploaded_files:
            df_parsed = _parse_rectification_file(f)
            if not df_parsed.empty:
                # Détecter le commercial depuis les données
                comm_detected = df_parsed['commercial'].iloc[0]
                # Normaliser: "FEDI" → "FEDI", "JILANI OBAY" → "JILANI OBAY", etc.
                for comm_key in COMM_LABELS:
                    if comm_key.upper() in comm_detected.upper() or comm_detected.upper() in comm_key.upper():
                        st.session_state['plans_rectifies'][comm_key] = df_parsed
                        st.success(f"✅ {comm_key} chargé — {len(df_parsed)} lignes, "
                                   f"{df_parsed['tonnes_jour'].sum():.0f}t planifiées")
                        break
                else:
                    # Si commercial non reconnu, utiliser le nom du fichier
                    fname = f.name.upper()
                    for comm_key in COMM_LABELS:
                        short = COMM_LABELS[comm_key].upper()
                        if short in fname or comm_key.split()[0].upper() in fname:
                            st.session_state['plans_rectifies'][comm_key] = df_parsed
                            st.success(f"✅ {comm_key} chargé (détecté depuis nom fichier)")
                            break
                    else:
                        st.warning(f"⚠️ Commercial non reconnu dans '{f.name}'. "
                                   f"Valeur trouvée: '{comm_detected}'")

    # Bouton reset
    if st.session_state['plans_rectifies']:
        if st.button("🗑️ Effacer tous les plans uploadés", key="clear_rectif"):
            st.session_state['plans_rectifies'] = {}
            st.rerun()

    st.markdown("---")

    # ── Vérifier qu'on a des données ────────────────────────────
    loaded_comms = list(st.session_state['plans_rectifies'].keys())
    if not loaded_comms:
        st.info("👆 Uploadez au moins un fichier de rectification pour afficher la comparaison.")
        
        # Afficher quand même le diagnostic si on a le planning OR-Tools
        if planning_df is not None and not planning_df.empty:
            st.markdown("### 🔍 Diagnostic OR-Tools (sans fichier rectifié)")
            _render_ortools_diagnostic(planning_df)
        return

    # ── Sélection du commercial à analyser ──────────────────────
    st.markdown("### 2️⃣ Analyser un commercial")

    selected_comm = st.selectbox(
        "Choisir le commercial à comparer",
        loaded_comms,
        format_func=lambda x: f"{x} ({'✅ rectifié' if x in loaded_comms else '⬜ manquant'})"
    )

    df_manuel   = st.session_state['plans_rectifies'].get(selected_comm, pd.DataFrame())
    df_ortools  = _extract_ortools_for_comparison(planning_df, selected_comm)

    if df_manuel.empty:
        st.warning(f"Fichier rectifié de {selected_comm} non chargé.")
        return

    if df_ortools.empty:
        st.warning(
            f"Aucune donnée OR-Tools pour {selected_comm} dans le planning Supabase. "
            f"Régénérez le planning d'abord."
        )

    # ── KPIs comparatifs ────────────────────────────────────────
    st.markdown(f"### 3️⃣ Comparaison — {selected_comm}")

    stats = _compute_comparison_stats(df_ortools, df_manuel)

    def _kpi_delta(ort_val, man_val, label, fmt="{:.0f}", inverse=False):
        """Affiche 2 métriques côte à côte avec delta."""
        col_o, col_m = st.columns(2)
        with col_o:
            st.metric(
                f"🤖 OR-Tools — {label}",
                fmt.format(ort_val) if ort_val else "—",
                help="Valeur calculée automatiquement par OR-Tools"
            )
        with col_m:
            delta = None
            if ort_val and man_val:
                delta = fmt.format(man_val - ort_val)
            st.metric(
                f"✍️ Manuel — {label}",
                fmt.format(man_val) if man_val else "—",
                delta=delta,
                delta_color="inverse" if inverse else "normal",
                help="Valeur rectifiée manuellement par le commercial"
            )

    s_ort = stats.get('OR-Tools', {})
    s_man = stats.get('Manuel', {})

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.metric("🤖 Total OR-Tools",  f"{s_ort.get('total_tonnes', 0):,.0f}t")
    with c2:
        st.metric("✍️ Total Manuel",    f"{s_man.get('total_tonnes', 0):,.0f}t",
                  delta=f"{s_man.get('total_tonnes',0)-s_ort.get('total_tonnes',0):+,.0f}t")
    with c3:
        st.metric("🤖 Max/jour OR-Tools", f"{s_ort.get('max_jour', 0):.0f}t/j")
    with c4:
        st.metric("✍️ Max/jour Manuel",   f"{s_man.get('max_jour', 0):.0f}t/j",
                  delta=f"{s_man.get('max_jour',0)-s_ort.get('max_jour',0):+.0f}t",
                  delta_color="inverse")

    c5, c6, c7, c8 = st.columns(4)
    with c5:
        st.metric("🤖 Jours actifs OR-T",   str(s_ort.get('n_jours_actifs', 0)))
    with c6:
        st.metric("✍️ Jours actifs Manuel", str(s_man.get('n_jours_actifs', 0)),
                  delta=str(s_man.get('n_jours_actifs',0)-s_ort.get('n_jours_actifs',0)))
    with c7:
        st.metric("🤖 Jours/agri OR-T",  f"{s_ort.get('avg_days_per_agri', 0):.1f}j")
    with c8:
        st.metric("✍️ Jours/agri Manuel", f"{s_man.get('avg_days_per_agri', 0):.1f}j",
                  delta=f"{s_man.get('avg_days_per_agri',0)-s_ort.get('avg_days_per_agri',0):+.1f}j")

    st.markdown("---")

    # ── Graphique 1 : Courbes journalières superposées ──────────
    st.markdown("#### 📉 Courbes journalières totales — OR-Tools vs Manuel")

    fig_daily = go.Figure()

    if not df_ortools.empty:
        daily_ort = df_ortools.groupby('date')['tonnes_jour'].sum().reset_index()
        daily_ort = daily_ort.sort_values('date')
        fig_daily.add_trace(go.Scatter(
            x=daily_ort['date'], y=daily_ort['tonnes_jour'],
            name="🤖 OR-Tools",
            line=dict(color=PLAN_ORTOOLS_COLOR, width=2.5),
            fill='tozeroy', fillcolor='rgba(59,130,246,0.1)',
            mode='lines',
        ))

    daily_man = df_manuel.groupby('date')['tonnes_jour'].sum().reset_index()
    daily_man = daily_man.sort_values('date')
    fig_daily.add_trace(go.Scatter(
        x=daily_man['date'], y=daily_man['tonnes_jour'],
        name="✍️ Plan Manuel",
        line=dict(color=PLAN_MANUEL_COLOR, width=2.5, dash='dot'),
        fill='tozeroy', fillcolor='rgba(245,166,35,0.1)',
        mode='lines',
    ))

    # Zone PIC
    fig_daily.add_vrect(
        x0="2026-07-01", x1="2026-07-15",
        fillcolor="gold", opacity=0.07, line_width=0,
        annotation_text="⚡ PIC", annotation_position="top left",
        annotation_font_color="#f5a623"
    )

    fig_daily.update_layout(
        template="plotly_dark", paper_bgcolor="#161b22", plot_bgcolor="#0d1117",
        height=400, hovermode="x unified",
        legend=dict(orientation="h", yanchor="bottom", y=1.02),
        yaxis_title="Tonnes/jour",
        title=f"Tonnes journalières totales — {selected_comm}",
    )
    st.plotly_chart(fig_daily, use_container_width=True)

    # ── Graphique 2 : Distribution du max par agriculteur ───────
    st.markdown("#### 📊 Distribution max journalier par agriculteur")
    col_g1, col_g2 = st.columns(2)

    with col_g1:
        if not df_ortools.empty:
            max_ort = df_ortools.groupby('agriculteur')['tonnes_jour'].max().reset_index()
            max_ort.columns = ['Agriculteur', 'Max/jour']
            fig_hist_ort = px.histogram(
                max_ort, x='Max/jour', nbins=20,
                title="🤖 OR-Tools — distribution max/jour",
                color_discrete_sequence=[PLAN_ORTOOLS_COLOR],
                template="plotly_dark",
            )
            fig_hist_ort.update_layout(
                paper_bgcolor="#161b22", plot_bgcolor="#0d1117", height=300,
                xaxis_title="Max tonnes/jour par agriculteur",
                yaxis_title="Nb agriculteurs",
            )
            st.plotly_chart(fig_hist_ort, use_container_width=True)

    with col_g2:
        max_man = df_manuel.groupby('agriculteur')['tonnes_jour'].max().reset_index()
        max_man.columns = ['Agriculteur', 'Max/jour']
        fig_hist_man = px.histogram(
            max_man, x='Max/jour', nbins=20,
            title="✍️ Manuel — distribution max/jour",
            color_discrete_sequence=[PLAN_MANUEL_COLOR],
            template="plotly_dark",
        )
        fig_hist_man.update_layout(
            paper_bgcolor="#161b22", plot_bgcolor="#0d1117", height=300,
            xaxis_title="Max tonnes/jour par agriculteur",
            yaxis_title="Nb agriculteurs",
        )
        st.plotly_chart(fig_hist_man, use_container_width=True)

    # ── Graphique 3 : Comparaison par agriculteur (top 10) ──────
    st.markdown("#### 🌾 Comparaison par agriculteur — Top 15 (tonnage total)")

    if not df_ortools.empty:
        agri_ort = df_ortools.groupby('agriculteur').agg(
            total_ort=('tonnes_jour', 'sum'),
            max_ort=('tonnes_jour', 'max'),
            n_days_ort=('date', 'nunique'),
        ).reset_index()
    else:
        agri_ort = pd.DataFrame(columns=['agriculteur','total_ort','max_ort','n_days_ort'])

    agri_man = df_manuel.groupby('agriculteur').agg(
        total_man=('tonnes_jour', 'sum'),
        max_man=('tonnes_jour', 'max'),
        n_days_man=('date', 'nunique'),
    ).reset_index()

    # Fusion
    agri_comp = agri_man.merge(agri_ort, on='agriculteur', how='outer').fillna(0)
    agri_comp['ecart_total'] = agri_comp['total_man'] - agri_comp['total_ort']
    agri_comp['ecart_max']   = agri_comp['max_man']   - agri_comp['max_ort']

    # Top 15 par tonnage manuel
    top15 = agri_comp.nlargest(15, 'total_man').sort_values('total_man')

    fig_agri = go.Figure()
    fig_agri.add_trace(go.Bar(
        name="🤖 OR-Tools", y=top15['agriculteur'], x=top15['total_ort'],
        orientation='h', marker_color=PLAN_ORTOOLS_COLOR,
        text=top15['total_ort'].apply(lambda x: f"{x:.0f}t" if x > 0 else ""),
        textposition='inside', textfont=dict(size=9),
    ))
    fig_agri.add_trace(go.Bar(
        name="✍️ Manuel", y=top15['agriculteur'], x=top15['total_man'],
        orientation='h', marker_color=PLAN_MANUEL_COLOR,
        text=top15['total_man'].apply(lambda x: f"{x:.0f}t"),
        textposition='outside', textfont=dict(size=9, color="#FFD700"),
    ))
    fig_agri.update_layout(
        barmode='overlay', template="plotly_dark",
        paper_bgcolor="#161b22", plot_bgcolor="#0d1117",
        height=max(400, len(top15)*32+80),
        title=f"Tonnage total — top 15 agriculteurs de {selected_comm}",
        xaxis_title="Tonnes saison",
        legend=dict(orientation="h", y=-0.1),
        margin=dict(l=250, r=80),
    )
    st.plotly_chart(fig_agri, use_container_width=True)

    # ── Graphique 4 : Écart max journalier par agriculteur ──────
    st.markdown("#### ⚠️ Agriculteurs avec le plus grand écart de max journalier")

    ecart_df = agri_comp.copy()
    ecart_df['|ecart_max|'] = ecart_df['ecart_max'].abs()
    top_ecart = ecart_df.nlargest(15, '|ecart_max|').sort_values('ecart_max')

    colors_bar = [PLAN_MANUEL_COLOR if v > 0 else PLAN_ORTOOLS_COLOR
                  for v in top_ecart['ecart_max']]
    fig_ecart = go.Figure(go.Bar(
        y=top_ecart['agriculteur'], x=top_ecart['ecart_max'],
        orientation='h',
        marker_color=colors_bar,
        text=[f"ORT:{row.max_ort:.0f}t → MAN:{row.max_man:.0f}t"
              for _, row in top_ecart.iterrows()],
        textposition='outside',
        textfont=dict(size=9),
    ))
    fig_ecart.add_vline(x=0, line_color="#888", line_width=1)
    fig_ecart.update_layout(
        template="plotly_dark", paper_bgcolor="#161b22", plot_bgcolor="#0d1117",
        height=max(380, len(top_ecart)*30+80),
        title="Écart de max journalier (Manuel − OR-Tools) par agriculteur",
        xaxis_title="Écart tonnes/jour (+ = manuel plus lissé)",
        margin=dict(l=250),
        annotations=[dict(
            x=0.02, y=1.08, xref='paper', yref='paper',
            text="<b>Orange</b> = OR-Tools trop fort | <b>Bleu</b> = OR-Tools trop faible",
            showarrow=False, font=dict(size=10, color="#aaa")
        )]
    )
    st.plotly_chart(fig_ecart, use_container_width=True)

    # ── Tableau détaillé ────────────────────────────────────────
    st.markdown("#### 📋 Tableau comparatif par agriculteur")

    agri_comp_display = agri_comp.copy()
    agri_comp_display['total_ort']   = agri_comp_display['total_ort'].round(0).astype(int)
    agri_comp_display['total_man']   = agri_comp_display['total_man'].round(0).astype(int)
    agri_comp_display['max_ort']     = agri_comp_display['max_ort'].round(0).astype(int)
    agri_comp_display['max_man']     = agri_comp_display['max_man'].round(0).astype(int)
    agri_comp_display['n_days_ort']  = agri_comp_display['n_days_ort'].round(0).astype(int)
    agri_comp_display['n_days_man']  = agri_comp_display['n_days_man'].round(0).astype(int)
    agri_comp_display['ecart_total'] = agri_comp_display['ecart_total'].round(0).astype(int)
    agri_comp_display['ecart_max']   = agri_comp_display['ecart_max'].round(0).astype(int)
    agri_comp_display['% ecart']     = (
        (agri_comp_display['ecart_total'] / agri_comp_display['total_man'].replace(0,1) * 100)
        .round(1).astype(str) + '%'
    )

    agri_comp_display = agri_comp_display.rename(columns={
        'agriculteur': 'Agriculteur',
        'total_ort':   'Total OR-T (t)',
        'total_man':   'Total Manuel (t)',
        'max_ort':     'Max/j OR-T (t)',
        'max_man':     'Max/j Manuel (t)',
        'n_days_ort':  'Jours OR-T',
        'n_days_man':  'Jours Manuel',
        'ecart_total': 'Δ Total (t)',
        'ecart_max':   'Δ Max/j (t)',
    })

    agri_comp_display = agri_comp_display.sort_values('Total Manuel (t)', ascending=False)
    st.dataframe(
        agri_comp_display[[
            'Agriculteur','Total OR-T (t)','Total Manuel (t)','% ecart',
            'Max/j OR-T (t)','Max/j Manuel (t)','Δ Max/j (t)',
            'Jours OR-T','Jours Manuel'
        ]],
        use_container_width=True, height=350, hide_index=True,
        column_config={
            'Δ Max/j (t)': st.column_config.NumberColumn(
                'Δ Max/j (t)',
                help="Négatif = OR-Tools trop fort (excès à corriger)"
            ),
        }
    )

    # ── Export ──────────────────────────────────────────────────
    if df_to_xlsx_styled:
        col_e1, col_e2 = st.columns(2)
        with col_e1:
            st.download_button(
                f"📊 Exporter comparaison {selected_comm} (Excel)",
                data=df_to_xlsx_styled(agri_comp_display),
                file_name=f"comparaison_{selected_comm.split()[0]}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        with col_e2:
            csv_data = agri_comp_display.to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')
            st.download_button(
                f"⬇️ CSV brut {selected_comm}",
                data=csv_data,
                file_name=f"comparaison_{selected_comm.split()[0]}.csv",
                mime="text/csv",
                use_container_width=True,
            )

    st.markdown("---")

    # ── Section Diagnostic ──────────────────────────────────────
    st.markdown("### 4️⃣ Diagnostic — Pourquoi OR-Tools diffère du plan manuel ?")
    _render_ortools_diagnostic(planning_df, selected_comm=selected_comm,
                               df_manuel=df_manuel, df_ortools=df_ortools,
                               agri_comp=agri_comp)

    # ── Vue globale multi-commerciaux ───────────────────────────
    if len(loaded_comms) > 1:
        st.markdown("---")
        st.markdown("### 5️⃣ Vue globale — Tous les commerciaux chargés")
        _render_global_comparison(planning_df, loaded_comms)


def _render_ortools_diagnostic(planning_df, selected_comm=None,
                                df_manuel=None, df_ortools=None, agri_comp=None):
    """
    Affiche le diagnostic des problèmes OR-Tools avec recommandations.
    """

    # ── Identification des problèmes ────────────────────────────
    problems_found = []

    if agri_comp is not None and not agri_comp.empty:
        # Pb 1: Max/jour trop élevé dans OR-Tools
        too_high = agri_comp[agri_comp['ecart_max'] < -20]
        if not too_high.empty:
            problems_found.append({
                'icon': '🔴',
                'titre': f"OR-Tools trop concentré — {len(too_high)} agriculteurs",
                'detail': (
                    f"OR-Tools met jusqu'à **{agri_comp['max_ort'].max():.0f}t/j** "
                    f"là où le plan manuel prévoit **{agri_comp['max_man'].max():.0f}t/j**.\n\n"
                    f"Cas les plus graves : {', '.join(too_high.nsmallest(3, 'ecart_max')['agriculteur'].tolist())}"
                ),
                'cause': (
                    "**Cause**: La borne `_ub_day = day_planned × 1.5` est trop permissive. "
                    "OR-Tools concentre le tonnage sur peu de jours au lieu de le lisser."
                ),
                'fix': (
                    "**Fix dans `optimizer_v2.py`** ligne ~880 :\n"
                    "```python\n"
                    "# AVANT (trop permissif):\n"
                    "_ub_day = max(int(day_planned * SCALE * 1.5), int(_get_min_tons(farmer) * SCALE))\n\n"
                    "# APRÈS (plus strict):\n"
                    "_ub_day = max(int(day_planned * SCALE * 1.1), int(_get_min_tons(farmer) * SCALE))\n"
                    "```"
                )
            })

        # Pb 2: Moins de jours actifs dans OR-Tools
        fewer_days = agri_comp[agri_comp['n_days_man'] > agri_comp['n_days_ort'] * 1.5]
        if not fewer_days.empty:
            avg_ort = agri_comp['n_days_ort'].mean()
            avg_man = agri_comp['n_days_man'].mean()
            problems_found.append({
                'icon': '🟡',
                'titre': f"OR-Tools trop peu de jours actifs — {len(fewer_days)} cas",
                'detail': (
                    f"OR-Tools active en moyenne **{avg_ort:.0f} jours** par agriculteur "
                    f"vs **{avg_man:.0f} jours** dans le plan manuel.\n\n"
                    f"Le plan manuel lisse sur plus de jours avec des micro-livraisons (10-30t/j)."
                ),
                'cause': (
                    "**Cause**: L'arrondi à la dizaine + borne haute trop large "
                    "pousse OR-Tools à concentrer sur peu de jours. "
                    "Les micro-livraisons (10-15t/j) sont arrondies à 10t mais "
                    "l'algorithme préfère des blocs de 50-100t."
                ),
                'fix': (
                    "**Fix dans `optimizer_v2.py`** section arrondi :\n"
                    "```python\n"
                    "# AVANT (arrondi à 10t):\n"
                    "tons = int(round(round(tons_brut, 1) / 10)) * 10\n\n"
                    "# APRÈS (arrondi à 5t pour granularité fine):\n"
                    "tons = int(round(round(tons_brut, 1) / 5)) * 5\n"
                    "_min_t_agri = max(10, _get_min_tons(farmer))  # minimum 10t\n"
                    "```"
                )
            })

        # Pb 3: Écart tonnage total > 5%
        big_ecart = agri_comp[abs(agri_comp['ecart_total']) > agri_comp['total_man'] * 0.05]
        if not big_ecart.empty:
            problems_found.append({
                'icon': '🟠',
                'titre': f"Écart tonnage total > 5% — {len(big_ecart)} agriculteurs",
                'detail': (
                    f"Le tonnage planifié par OR-Tools diffère de plus de 5% "
                    f"du plan manuel pour {len(big_ecart)} agriculteurs.\n\n"
                    f"Exemples: {', '.join(big_ecart.nlargest(3, '|ecart_max|')['agriculteur'].tolist() if '|ecart_max|' in big_ecart.columns else big_ecart.head(3)['agriculteur'].tolist())}"
                ),
                'cause': (
                    "**Cause**: La tolérance ±2% de la contrainte tonnage dans OR-Tools "
                    "plus l'arrondi à la dizaine peut cumuler jusqu'à 8-10% d'écart. "
                    "La correction post-traitement est désactivée."
                ),
                'fix': (
                    "**Pas de fix nécessaire** si l'écart reste < 5% globalement. "
                    "Le plan manuel a lui-même des écarts dus aux arrondis manuels."
                )
            })

    # ── Affichage des problèmes ──────────────────────────────────
    if not problems_found:
        st.success("✅ Aucun problème majeur détecté — OR-Tools est aligné avec le plan manuel.")
    else:
        for pb in problems_found:
            with st.expander(f"{pb['icon']} {pb['titre']}", expanded=True):
                st.markdown(pb['detail'])
                st.markdown("---")
                st.markdown(pb['cause'])
                st.code(pb['fix'].replace("```python\n", "").replace("\n```", "").strip()
                        if "```" in pb['fix'] else pb['fix'])

    # ── Recommandations globales ─────────────────────────────────
    st.markdown("---")
    st.markdown("#### 🔧 Corrections recommandées dans `optimizer_v2.py`")

    with st.expander("Voir toutes les corrections (copier-coller dans PowerShell)", expanded=False):
        st.markdown("""
**Correction 1 — Borne journalière plus stricte** (ligne `_ub_day`) :
```python
# Chercher cette ligne (~880):
_ub_day = max(int(day_planned * SCALE * 1.5), int(_get_min_tons(farmer) * SCALE))
# Remplacer par:
_ub_day = max(int(day_planned * SCALE * 1.1), int(_get_min_tons(farmer) * SCALE))
```

**Correction 2 — FACTORY_OVERFLOW_WEIGHT ELFALLEH** (ligne `FACTORY_OVERFLOW_WEIGHT`) :
```python
# Chercher:
FACTORY_OVERFLOW_WEIGHT = 500
# Remplacer par:
FACTORY_OVERFLOW_WEIGHT = 2000
```

**Correction 3 — Arrondi à 5t au lieu de 10t** (section génération all_days) :
```python
# Chercher:
tons = int(round(round(tons_brut, 1) / 10)) * 10
# Remplacer par:
tons = int(round(round(tons_brut, 1) / 5)) * 5
if tons == 0 and tons_brut > 0:
    tons = max(10, _min_t_agri)
elif 0 < tons < _min_t_agri:
    tons = _min_t_agri
```
""")

        # Script PowerShell complet
        ps_script = r"""@"
import re
f=open('optimizer_v2.py','r',encoding='utf-8')
c=f.read()
f.close()

# Fix 1: borne journaliere
c=c.replace(
    'day_planned * SCALE * 1.5',
    'day_planned * SCALE * 1.1'
)

# Fix 2: ELFALLEH overflow
c=c.replace(
    'FACTORY_OVERFLOW_WEIGHT = 500',
    'FACTORY_OVERFLOW_WEIGHT = 2000'
)

# Fix 3: arrondi 5t
c=c.replace(
    'tons = int(round(round(tons_brut, 1) / 10)) * 10',
    'tons = int(round(round(tons_brut, 1) / 5)) * 5'
)

f=open('optimizer_v2.py','w',encoding='utf-8')
f.write(c)
f.close()
print('OK - 3 corrections appliquees')
"@ | Out-File -Encoding utf8 fix_comparaison.py
python fix_comparaison.py"""

        st.code(ps_script, language="powershell")
        st.caption("⚠️ Après ces corrections, relancer `python optimizer_v2.py` puis `python migrate.py`")


def _render_global_comparison(planning_df, loaded_comms):
    """
    Vue globale de la comparaison pour tous les commerciaux chargés.
    """
    rows_global = []
    for comm in loaded_comms:
        df_man = st.session_state['plans_rectifies'].get(comm, pd.DataFrame())
        df_ort = _extract_ortools_for_comparison(planning_df, comm)

        if df_man.empty:
            continue

        tot_man = df_man['tonnes_jour'].sum()
        max_man = df_man.groupby('date')['tonnes_jour'].sum().max()
        n_days_man = df_man['date'].nunique()

        tot_ort = df_ort['tonnes_jour'].sum() if not df_ort.empty else 0
        max_ort = df_ort.groupby('date')['tonnes_jour'].sum().max() if not df_ort.empty else 0
        n_days_ort = df_ort['date'].nunique() if not df_ort.empty else 0

        rows_global.append({
            'Commercial':       comm,
            'Total Manuel (t)': round(tot_man, 0),
            'Total OR-T (t)':   round(tot_ort, 0),
            'Δ Total (t)':      round(tot_man - tot_ort, 0),
            'Max/j Manuel':     round(max_man, 0),
            'Max/j OR-T':       round(max_ort, 0),
            'Δ Max/j':          round(max_man - max_ort, 0),
            'Jours Manuel':     n_days_man,
            'Jours OR-T':       n_days_ort,
        })

    if not rows_global:
        return

    df_global = pd.DataFrame(rows_global)

    # Graphique comparatif global
    fig_g = go.Figure()
    fig_g.add_trace(go.Bar(
        name="🤖 OR-Tools", x=df_global['Commercial'], y=df_global['Total OR-T (t)'],
        marker_color=PLAN_ORTOOLS_COLOR,
        text=df_global['Total OR-T (t)'].apply(lambda x: f"{x:.0f}t"),
        textposition='inside',
    ))
    fig_g.add_trace(go.Bar(
        name="✍️ Manuel", x=df_global['Commercial'], y=df_global['Total Manuel (t)'],
        marker_color=PLAN_MANUEL_COLOR,
        text=df_global['Total Manuel (t)'].apply(lambda x: f"{x:.0f}t"),
        textposition='outside',
    ))
    fig_g.update_layout(
        barmode='group', template="plotly_dark",
        paper_bgcolor="#161b22", plot_bgcolor="#0d1117",
        height=360, title="Tonnage total saison — OR-Tools vs Manuel (tous commerciaux)",
        legend=dict(orientation="h", y=-0.15),
        yaxis_title="Tonnes",
    )
    st.plotly_chart(fig_g, use_container_width=True)
    st.dataframe(df_global, use_container_width=True, hide_index=True)