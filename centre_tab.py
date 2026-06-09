import streamlit as st
import pandas as pd
import plotly.express as px

CENTRES = {
    "baccara":    "BACCARA",
    "kerkouane":  "KERKOUANE",
    "centre428":  "428",
}

def render_centre_dashboard(sb, role_filter, centre_name):
    """Dashboard dédié pour les centres de collecte."""
    centre_key = CENTRES.get(role_filter.lower(), centre_name.upper())
    
    st.title(f"🏭 Centre {centre_key}")
    st.caption(f"Vue des agriculteurs rattachés à ce centre")
    
    # ── Charger les agriculteurs de ce centre ──────────────────────
    try:
        res = sb.table("agriculteurs").select(
            "commercial,nom,tonnage_total,usine,region,zone,nbr_hectares,centre,accessibilite,date_debut,date_fin"
        ).eq("centre", centre_key).execute()
        agri_data = res.data or []
    except Exception as e:
        st.error(f"Erreur chargement agriculteurs: {e}")
        agri_data = []

    if not agri_data:
        st.warning(f"⚠️ Aucun agriculteur rattaché au centre {centre_key} pour le moment.")
        st.info("Les commerciaux doivent indiquer le centre dans leur fichier upload "
                "(colonne **CENTRE**).")
        
        # Afficher les agriculteurs sans centre pour info
        try:
            res2 = sb.table("agriculteurs").select(
                "commercial,nom,tonnage_total,usine,region,centre"
            ).is_("centre", "null").execute()
            no_centre = res2.data or []
            if no_centre:
                st.markdown(f"---")
                st.markdown(f"**{len(no_centre)} agriculteurs sans centre assigné :**")
                df_nc = pd.DataFrame(no_centre)
                st.dataframe(df_nc[["commercial","nom","tonnage_total","usine","region"]],
                             hide_index=True, use_container_width=True)
        except Exception:
            pass
        return

    df = pd.DataFrame(agri_data)
    df["tonnage_total"] = pd.to_numeric(df["tonnage_total"], errors="coerce").fillna(0)
    df["nbr_hectares"]  = pd.to_numeric(df.get("nbr_hectares", 0), errors="coerce").fillna(0)
    
    # ── KPIs ───────────────────────────────────────────────────────
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("👨‍🌾 Agriculteurs", len(df))
    col2.metric("🍅 Tonnage total", f"{df['tonnage_total'].sum():,.0f}t")
    col3.metric("🌿 Hectares", f"{df['nbr_hectares'].sum():,.1f}ha")
    avg_tpha = df["tonnage_total"].sum() / df["nbr_hectares"].sum()                if df["nbr_hectares"].sum() > 0 else 0
    col4.metric("📊 t/ha moyen", f"{avg_tpha:.1f}")
    
    st.markdown("---")
    
    # ── Tableau par commercial ──────────────────────────────────────
    st.subheader("Par commercial")
    comm_grp = df.groupby("commercial").agg(
        Agriculteurs=("nom","count"),
        Tonnage=("tonnage_total","sum"),
        Hectares=("nbr_hectares","sum")
    ).reset_index()
    comm_grp["t/ha"] = (comm_grp["Tonnage"] / comm_grp["Hectares"]).round(1)
    comm_grp["Tonnage"] = comm_grp["Tonnage"].round(0).astype(int)
    st.dataframe(comm_grp, hide_index=True, use_container_width=True)
    
    # ── Répartition par usine ───────────────────────────────────────
    col_a, col_b = st.columns(2)
    with col_a:
        st.subheader("Par usine")
        usine_grp = df.groupby("usine")["tonnage_total"].sum().reset_index()
        usine_grp.columns = ["Usine","Tonnage"]
        fig = px.pie(usine_grp, names="Usine", values="Tonnage",
                     color_discrete_sequence=px.colors.qualitative.Set2,
                     template="plotly_dark")
        fig.update_layout(paper_bgcolor="#161b22", height=280, margin=dict(t=20,b=20))
        st.plotly_chart(fig, use_container_width=True)
    
    with col_b:
        st.subheader("Tonnage par commercial")
        fig2 = px.bar(comm_grp, x="commercial", y="Tonnage",
                      color="commercial", template="plotly_dark",
                      color_discrete_sequence=px.colors.qualitative.Pastel)
        fig2.update_layout(paper_bgcolor="#161b22", height=280,
                           showlegend=False, margin=dict(t=20,b=20))
        st.plotly_chart(fig2, use_container_width=True)
    
    # ── Liste complète des agriculteurs ────────────────────────────
    st.markdown("---")
    st.subheader(f"Tous les agriculteurs du centre {centre_key}")
    
    # Filtres
    comm_list = ["Tous"] + sorted(df["commercial"].unique().tolist())
    usine_list = ["Toutes"] + sorted(df["usine"].unique().tolist())
    fc1, fc2 = st.columns(2)
    sel_comm  = fc1.selectbox("Filtrer par commercial", comm_list)
    sel_usine = fc2.selectbox("Filtrer par usine", usine_list)
    
    df_filt = df.copy()
    if sel_comm  != "Tous":  df_filt = df_filt[df_filt["commercial"]==sel_comm]
    if sel_usine != "Toutes": df_filt = df_filt[df_filt["usine"]==sel_usine]
    
    # Afficher
    show_cols = ["nom","commercial","usine","region","zone","tonnage_total","nbr_hectares"]
    show_cols = [c for c in show_cols if c in df_filt.columns]
    df_show = df_filt[show_cols].copy()
    df_show.columns = [c.replace("_"," ").title() for c in df_show.columns]
    if "Tonnage Total" in df_show.columns:
        df_show["Tonnage Total"] = df_show["Tonnage Total"].round(0)
    
    st.dataframe(df_show.sort_values("Tonnage Total" if "Tonnage Total" in df_show.columns 
                                     else df_show.columns[0], ascending=False),
                 hide_index=True, use_container_width=True, height=400)
    
    # Export
    st.download_button(
        f"⬇️ Exporter agriculteurs centre {centre_key}",
        data=df_filt.to_csv(index=False).encode("utf-8-sig"),
        file_name=f"agriculteurs_centre_{centre_key}.csv",
        mime="text/csv"
    )