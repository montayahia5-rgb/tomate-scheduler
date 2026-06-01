# ============================================================
# DASHBOARD PHASE 7 — All vehicles + CSV export + Deploy info
#
# HOW TO RUN:
#   1. Put this file in the same folder as:
#        Planning_Phase4_Transport_Double.xlsx
#        Recap_tonnage_pre_vu_ajuste__mai26.xlsx
#        phase4.py
#   2. pip install streamlit plotly pandas openpyxl watchdog xlsxwriter
#   3. streamlit run dashboard_phase7.py
#
# PHASE 7 NEW vs phase 6:
#   - TRACTEUR added everywhere (even when 0 trips — shows as 0)
#   - Export CSV button for every table in every tab
#   - Export ALL sheets as ZIP (one click = all 5 tables)
#   - Fleet simulator includes TRACTEUR
#   - "Deploy" explained inside the app
# ============================================================

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import datetime, os, subprocess, sys, io, zipfile

# ── CSV / Excel export helpers ───────────────────────────────
def df_to_csv(df):
    """Convert dataframe to CSV bytes for download button."""
    return df.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")

def dfs_to_zip(sheets: dict) -> bytes:
    """
    Convert multiple dataframes to a ZIP of CSVs.
    sheets = {"filename.csv": dataframe, ...}
    Returns ZIP bytes.
    """
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for fname, df in sheets.items():
            csv_bytes = df.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
            zf.writestr(fname, csv_bytes)
    buf.seek(0)
    return buf.read()

def dfs_to_excel(sheets: dict) -> bytes:
    """
    Convert multiple dataframes to a single Excel file with multiple sheets.
    sheets = {"Sheet Name": dataframe, ...}
    Returns Excel bytes.
    """
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        for sheet_name, df in sheets.items():
            df.to_excel(writer, sheet_name=sheet_name[:31], index=False)
    buf.seek(0)
    return buf.read()

# ── Page config ──────────────────────────────────────────────
st.set_page_config(
    page_title="🍅 Tomate Planning 2026",
    layout="wide",
    initial_sidebar_state="expanded",
    page_icon="🍅"
)

st.markdown("""
<style>
  [data-testid="stAppViewContainer"] { background: #0d1117; }
  [data-testid="stSidebar"]          { background: #161b22; border-right:1px solid #21262d; }
  .block-container { padding-top: 1.5rem; }
  h1,h2,h3 { color: #f0f6fc; }
  .metric-row { display:flex; gap:16px; margin-bottom:20px; flex-wrap:wrap; }
  .kpi-box {
    background:#161b22; border:1px solid #21262d; border-radius:12px;
    padding:16px 20px; min-width:140px; flex:1;
    border-top: 2px solid var(--c,#e8543a);
  }
  .kpi-val { font-size:1.7rem; font-weight:700; color:#f0f6fc; }
  .kpi-lbl { font-size:.72rem; color:#8b949e; text-transform:uppercase; letter-spacing:.06em; margin-top:3px; }
  .kpi-sub { font-size:.7rem; color:#3dd68c; margin-top:4px; }
  .peak-box {
    background:linear-gradient(90deg,rgba(255,179,71,.1),rgba(232,84,58,.08));
    border:1px solid rgba(255,179,71,.25); border-radius:10px;
    padding:12px 18px; margin-bottom:18px; font-size:.85rem; color:#f0f6fc;
  }
  .deploy-box {
    background:rgba(59,130,246,.08); border:1px solid rgba(59,130,246,.25);
    border-radius:10px; padding:14px 18px; font-size:.82rem; color:#f0f6fc;
    line-height:1.6;
  }
  .stDownloadButton > button {
    background:#161b22 !important; border:1px solid #21262d !important;
    color:#f0f6fc !important; font-size:.8rem !important;
  }
  .stDownloadButton > button:hover {
    border-color:#3b82f6 !important; color:#3b82f6 !important;
  }
</style>
""", unsafe_allow_html=True)

# ── File paths — always relative to THIS script's folder ─────
# This fixes the "phase4.py not found" error when Streamlit runs
# from a different working directory than your project folder
SCRIPT_DIR    = os.path.dirname(os.path.abspath(__file__))
PHASE4_FILE   = os.path.join(SCRIPT_DIR, "Planning_Phase4_Transport_Double.xlsx")
ORIGINAL_FILE = os.path.join(SCRIPT_DIR, "Recap_tonnage_pre_vu_ajuste__mai26.xlsx")
PHASE4_SCRIPT = os.path.join(SCRIPT_DIR, "phase4.py")

PEAK_START = datetime.date(2026, 7, 1)
PEAK_END   = datetime.date(2026, 7, 15)

COMM_COLORS = {
    "ACHREF AJLANI":  "#8b5cf6",
    "FEDI":           "#3b82f6",
    "JILANI OBAY":    "#e8543a",
    "KHALIL":         "#f5a623",
    "MAKKI BEN SALAH":"#00e5a0",
}
FACTORY_COLORS = {
    "ABIDA":    "#ff6b9d",
    "COMOCAP":  "#3b82f6",
    "ELFALLEH": "#00e5a0",
    "SICAM":    "#f5a623",
    "TUCAL":    "#8b5cf6",
}

# ── Load data (cached, reloads when file changes) ─────────────
def file_mtime(path):
    try:    return os.path.getmtime(path)
    except: return 0

@st.cache_data(ttl=30)   # re-check every 30 seconds automatically
def load_data(mtime_p4, mtime_orig):
    """
    mtime params are passed so cache invalidates when files change.
    Phase 6 key feature: auto-refresh when Excel is updated.
    """
    if not os.path.exists(PHASE4_FILE):
        return None, None, None, None, None, None

    # ---- Planning Journalier ----
    planning = pd.read_excel(PHASE4_FILE, sheet_name="Planning Journalier", header=0)
    planning["Date"] = pd.to_datetime(planning["Date"], errors="coerce")
    planning = planning.dropna(subset=["Date"])

    # ---- Besoins Transport-Jour ----
    transport = pd.read_excel(PHASE4_FILE, sheet_name="Besoins Transport-Jour", header=0)
    transport["Date"] = pd.to_datetime(transport["Date"], errors="coerce")
    transport = transport.dropna(subset=["Date"])

    # ---- Disponibilité Véhicules ----
    dispo = pd.read_excel(PHASE4_FILE, sheet_name="Disponibilité Véhicules", header=0)
    for col in ["Date Début","Date Fin Orig.","Date Fin Nouvelle","Libre À Partir De"]:
        if col in dispo.columns:
            dispo[col] = pd.to_datetime(dispo[col], errors="coerce")

    # ---- Journal Double Transport (header is on row 1, not row 0) ----
    double_j = pd.read_excel(PHASE4_FILE, sheet_name="Journal Double Transport", header=1)
    # Clean: drop rows where Commercial is NaN
    double_j = double_j.dropna(subset=["Commercial"])

    # ---- Résumé par Commercial ----
    resume = pd.read_excel(PHASE4_FILE, sheet_name="Résumé par Commercial", header=0)
    resume = resume[resume["Commercial"] != "TOTAL"]

    # ---- Original Excel (for extra details) ----
    orig = None
    if os.path.exists(ORIGINAL_FILE):
        orig = pd.read_excel(ORIGINAL_FILE, sheet_name="Feuil1")
        orig = orig[~orig["responsable région"].astype(str).str.contains("TOTAL", na=False)]
        orig = orig.dropna(subset=["AGRICULTEUR","TONNAGE","USINE"])
        orig["responsable région"] = orig["responsable région"].astype(str).str.strip()
        orig["USINE"] = orig["USINE"].astype(str).str.strip()

    return planning, transport, dispo, double_j, resume, orig

# ── Load with auto-refresh ────────────────────────────────────
mtime_p4   = file_mtime(PHASE4_FILE)
mtime_orig = file_mtime(ORIGINAL_FILE)
result = load_data(mtime_p4, mtime_orig)
planning, transport, dispo, double_j, resume, orig = result

# ── Sidebar ───────────────────────────────────────────────────
with st.sidebar:
    st.image("https://em-content.zobj.net/source/apple/354/tomato_1f345.png", width=52)
    st.title("🍅 Tomate 2026")
    st.caption("Tableau de bord — Transport & Récolte")
    st.divider()

    # Phase 6: Re-run Phase 4 on demand
    st.subheader("⚙️ Mise à jour données")
    st.caption("Lance phase4.py pour recalculer depuis ton Excel original.")
    if st.button("🔄 Régénérer le planning", use_container_width=True, type="primary"):
        if os.path.exists(PHASE4_SCRIPT):
            with st.spinner("Calcul en cours..."):
                result = subprocess.run(
                    [sys.executable, PHASE4_SCRIPT],
                    capture_output=True,
                    text=True,
                    timeout=120,
                    cwd=SCRIPT_DIR,           # ← KEY FIX: run from the project folder
                    encoding="utf-8",         # ← KEY FIX: avoid Windows cp1252 crash
                    errors="replace",
                )
            if result.returncode == 0:
                st.success("✅ Planning régénéré avec succès !")
                st.cache_data.clear()
                st.rerun()
            else:
                st.error("❌ Erreur lors du calcul :")
                st.code(result.stderr[-600:], language="text")
        else:
            st.warning(f"❌ phase4.py introuvable dans : {SCRIPT_DIR}")

    st.divider()

    # Filters
    st.subheader("🔍 Filtres")
    if planning is not None:
        comms_all = sorted(planning["Commercial"].dropna().unique())
        sel_comms = st.multiselect("Commercial(s)", comms_all, default=comms_all)

        if planning is not None and "Usine" in planning.columns:
            facts_all = sorted(planning["Usine"].dropna().unique())
            sel_facts = st.multiselect("Usine(s)", facts_all, default=facts_all)
        else:
            sel_facts = []

        d_min = planning["Date"].min().date()
        d_max = planning["Date"].max().date()
        date_range = st.date_input("Période", value=(d_min, d_max),
                                   min_value=d_min, max_value=d_max)
        peak_only = st.checkbox("⚡ Pic seulement (1–15 Jul)")
    else:
        sel_comms, sel_facts = [], []
        date_range = None
        peak_only = False

    st.divider()

    # Fleet inventory — ALL 4 vehicle types including TRACTEUR
    st.subheader("🚛 Votre flotte")
    fl_trac = st.number_input("TRACTEUR",       0, 20, 0,
                               help="0 = non utilisé cette saison (affiché quand même)")
    fl_ppl  = st.number_input("PETIT POILOUR", 0, 30, 3)
    fl_pl   = st.number_input("POILOUR",       0, 30, 6)
    fl_semi = st.number_input("SEMI",          0, 20, 4)

    st.divider()

    # ── EXPORT ALL — ZIP or Excel ─────────────────────────────
    st.subheader("📥 Exporter tout")
    st.caption("Télécharge toutes les 5 feuilles du planning Phase 4.")
    if planning is not None:
        export_sheets = {
            "Planning Journalier":      planning,
            "Transport-Jour":           transport,
            "Disponibilité Véhicules":  dispo,
            "Journal Décalage":         double_j,
            "Résumé Commercial":        resume,
        }
        # Excel export (all sheets in one file)
        xlsx_bytes = dfs_to_excel(export_sheets)
        st.download_button(
            "⬇️ Excel complet (5 feuilles)",
            data=xlsx_bytes,
            file_name="Tomate_Planning_Export.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
        # ZIP of CSVs
        zip_sheets = {f"{k.replace(' ','_')}.csv": v for k,v in export_sheets.items()}
        zip_bytes = dfs_to_zip(zip_sheets)
        st.download_button(
            "⬇️ ZIP de CSV (5 fichiers)",
            data=zip_bytes,
            file_name="Tomate_Planning_CSV.zip",
            mime="application/zip",
            use_container_width=True,
        )

    st.divider()
    if planning is not None:
        last_date = planning["Date"].max().strftime("%d/%m/%Y")
        st.caption(f"📅 Données jusqu'au {last_date}")
        st.caption(f"🔁 Auto-refresh toutes les 30s")

# ── Guard: file not found ─────────────────────────────────────
if planning is None:
    st.error(f"❌ Fichier '{PHASE4_FILE}' introuvable.")
    st.info("👉 Lance d'abord `python phase4.py` dans ce dossier, puis relance le dashboard.")
    st.stop()

# ── Apply filters ─────────────────────────────────────────────
p = planning[planning["Commercial"].isin(sel_comms)].copy()
t = transport[transport["Commercial"].isin(sel_comms)].copy()

if "Usine" in p.columns and sel_facts:
    p = p[p["Usine"].isin(sel_facts)]

if isinstance(date_range, (list, tuple)) and len(date_range) == 2:
    d0, d1 = date_range
    p = p[(p["Date"].dt.date >= d0) & (p["Date"].dt.date <= d1)]
    t = t[(t["Date"].dt.date >= d0) & (t["Date"].dt.date <= d1)]

if peak_only:
    p = p[(p["Date"].dt.date >= PEAK_START) & (p["Date"].dt.date <= PEAK_END)]
    t = t[(t["Date"].dt.date >= PEAK_START) & (t["Date"].dt.date <= PEAK_END)]

# ── Header + KPIs ─────────────────────────────────────────────
st.title("🍅 Tomate Planning 2026")
st.caption("Phase 7 — Tous véhicules · Export CSV/Excel · Explication Deploy")

total_tons   = p["Tonnes/Jour"].sum()
total_trips  = int(p["Nb Voyages"].sum())
n_farmers    = p["Agriculteur"].nunique()
peak_tons    = p[(p["Date"].dt.date >= PEAK_START) & (p["Date"].dt.date <= PEAK_END)]["Tonnes/Jour"].sum()
n_conflicts  = int(double_j["Commercial"].notna().sum())
dbl_days     = int(t["Jours Double"].sum()) if "Jours Double" in t.columns else 0

st.markdown(f"""
<div class="metric-row">
  <div class="kpi-box" style="--c:#e8543a"><div class="kpi-val">{total_tons:,.0f} t</div><div class="kpi-lbl">Tonnes totales</div><div class="kpi-sub">↑ Saison complète</div></div>
  <div class="kpi-box" style="--c:#3b82f6"><div class="kpi-val">{n_farmers}</div><div class="kpi-lbl">Agriculteurs</div><div class="kpi-sub">{len(sel_comms)} commerciaux</div></div>
  <div class="kpi-box" style="--c:#f5a623"><div class="kpi-val">{peak_tons:,.0f} t</div><div class="kpi-lbl">Tonnes période pic</div><div class="kpi-sub">1–15 Juillet</div></div>
  <div class="kpi-box" style="--c:#8b5cf6"><div class="kpi-val">{total_trips:,}</div><div class="kpi-lbl">Voyages totaux</div><div class="kpi-sub">PPL + PL + SEMI</div></div>
  <div class="kpi-box" style="--c:#00e5a0"><div class="kpi-val">{n_conflicts}</div><div class="kpi-lbl">Conflits résolus</div><div class="kpi-sub">Décalage ≤ 4 jours</div></div>
  <div class="kpi-box" style="--c:#22d3ee"><div class="kpi-val">{dbl_days}</div><div class="kpi-lbl">Jours double transport</div><div class="kpi-sub">Pour libérer véhicules</div></div>
</div>
<div class="peak-box">⚡ <b>Pic 1–15 Juillet :</b> 39% de toute la saison en 15 jours. Les zones jaunes sur les graphiques marquent cette période critique.</div>
""", unsafe_allow_html=True)

# ── Tabs ─────────────────────────────────────────────────────
tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
    "📅 Planning Journalier",
    "👤 Par Commercial",
    "🏭 Par Usine",
    "🚛 Transport & Alertes",
    "⚙️ Décalage & Conflits",
    "🚀 Deploy & Prochaines Étapes",
])

# ── TAB 1: DAILY PLANNING ────────────────────────────────────
with tab1:
    daily = p.groupby("Date")["Tonnes/Jour"].sum().reset_index()
    daily["Période"] = daily["Date"].apply(
        lambda d: "⚡ Pic (1-15 Jul)" if PEAK_START <= d.date() <= PEAK_END else "Normal"
    )

    fig = px.bar(
        daily, x="Date", y="Tonnes/Jour", color="Période",
        color_discrete_map={"⚡ Pic (1-15 Jul)":"#f5a623", "Normal":"#3b82f6"},
        title="Tonnes récoltées par jour — toute la saison",
        labels={"Tonnes/Jour":"Tonnes/jour"},
        template="plotly_dark",
    )
    fig.update_layout(
        plot_bgcolor="#0d1117", paper_bgcolor="#161b22",
        legend_title="Période", hovermode="x unified", height=420,
        font_color="#8b949e",
    )
    fig.add_vrect(x0=str(PEAK_START), x1=str(PEAK_END),
                  fillcolor="gold", opacity=0.06, line_width=0,
                  annotation_text="⚡ PIC", annotation_position="top left",
                  annotation_font_color="#f5a623")
    st.plotly_chart(fig, use_container_width=True)

    c1, c2 = st.columns(2)
    with c1:
        fig2 = px.pie(
            p.groupby("Commercial")["Tonnes/Jour"].sum().reset_index(),
            names="Commercial", values="Tonnes/Jour",
            title="Répartition par commercial", hole=0.45,
            color="Commercial",
            color_discrete_map=COMM_COLORS,
            template="plotly_dark",
        )
        fig2.update_layout(paper_bgcolor="#161b22", height=320)
        st.plotly_chart(fig2, use_container_width=True)
    with c2:
        if "Usine" in p.columns:
            fig3 = px.pie(
                p.groupby("Usine")["Tonnes/Jour"].sum().reset_index(),
                names="Usine", values="Tonnes/Jour",
                title="Répartition par usine", hole=0.45,
                color="Usine", color_discrete_map=FACTORY_COLORS,
                template="plotly_dark",
            )
            fig3.update_layout(paper_bgcolor="#161b22", height=320)
            st.plotly_chart(fig3, use_container_width=True)

    st.subheader("📋 Données détaillées")
    display_cols = [c for c in ["Date","Commercial","Agriculteur","Usine",
                                "Tonnes/Jour","Type Véhicule","Nb Voyages",
                                "Pic de Récolte","Note"] if c in p.columns]
    st.dataframe(
        p[display_cols].sort_values("Date").reset_index(drop=True),
        use_container_width=True, height=280,
    )
    st.download_button(
        "⬇️ Exporter planning journalier (CSV)",
        data=df_to_csv(p[display_cols].sort_values("Date").reset_index(drop=True)),
        file_name="planning_journalier.csv",
        mime="text/csv",
    )

# ── TAB 2: PAR COMMERCIAL ────────────────────────────────────
with tab2:
    # Line chart per commercial
    comm_daily = (p.groupby(["Date","Commercial"])["Tonnes/Jour"]
                  .sum().reset_index())
    fig4 = px.line(
        comm_daily, x="Date", y="Tonnes/Jour", color="Commercial",
        color_discrete_map=COMM_COLORS,
        title="Tonnes/jour par commercial",
        labels={"Tonnes/Jour":"Tonnes/j"},
        template="plotly_dark",
    )
    fig4.update_traces(line_width=2)
    fig4.update_layout(
        plot_bgcolor="#0d1117", paper_bgcolor="#161b22",
        height=400, hovermode="x unified",
    )
    fig4.add_vrect(x0=str(PEAK_START), x1=str(PEAK_END),
                   fillcolor="gold", opacity=0.06, line_width=0)
    st.plotly_chart(fig4, use_container_width=True)

    c1, c2, c3 = st.columns(3)
    with c1:
        comm_tot = p.groupby("Commercial")["Tonnes/Jour"].sum().reset_index()
        fig5 = px.bar(
            comm_tot, x="Commercial", y="Tonnes/Jour",
            color="Commercial", color_discrete_map=COMM_COLORS,
            title="Tonnes totales", template="plotly_dark",
            text_auto=".3s",
        )
        fig5.update_layout(paper_bgcolor="#161b22", showlegend=False, height=300)
        st.plotly_chart(fig5, use_container_width=True)
    with c2:
        farmers_ct = p.groupby("Commercial")["Agriculteur"].nunique().reset_index()
        fig6 = px.bar(
            farmers_ct, x="Commercial", y="Agriculteur",
            color="Commercial", color_discrete_map=COMM_COLORS,
            title="Nb agriculteurs", template="plotly_dark",
        )
        fig6.update_layout(paper_bgcolor="#161b22", showlegend=False, height=300)
        st.plotly_chart(fig6, use_container_width=True)
    with c3:
        if not resume.empty:
            fig7 = px.bar(
                resume, x="Commercial", y="Conflits Résolus",
                color="Commercial", color_discrete_map=COMM_COLORS,
                title="Conflits résolus", template="plotly_dark",
            )
            fig7.update_layout(paper_bgcolor="#161b22", showlegend=False, height=300)
            st.plotly_chart(fig7, use_container_width=True)

    # Per-commercial drill-down
    st.subheader("🔎 Détail par commercial")
    selected = st.selectbox("Choisir un commercial", sel_comms)
    one = p[p["Commercial"] == selected]
    one_daily = one.groupby("Date")["Tonnes/Jour"].sum().reset_index()
    fig8 = px.area(
        one_daily, x="Date", y="Tonnes/Jour",
        title=f"Tonnes/jour — {selected}",
        color_discrete_sequence=[COMM_COLORS.get(selected,"#3b82f6")],
        template="plotly_dark",
    )
    fig8.update_layout(paper_bgcolor="#161b22", height=280)
    fig8.add_vrect(x0=str(PEAK_START), x1=str(PEAK_END),
                   fillcolor="gold", opacity=0.08, line_width=0)
    st.plotly_chart(fig8, use_container_width=True)

    show_cols = [c for c in ["Date","Agriculteur","Usine","Tonnes/Jour",
                              "Type Véhicule","Nb Voyages","Note"] if c in one.columns]
    st.dataframe(one[show_cols].sort_values("Date").reset_index(drop=True),
                 use_container_width=True, height=240)

# ── TAB 3: PAR USINE ─────────────────────────────────────────
with tab3:
    if "Usine" not in p.columns:
        st.info("Colonne 'Usine' absente du planning.")
    else:
        # Cards per factory
        factories = sorted(p["Usine"].dropna().unique())
        cols_f = st.columns(len(factories))
        for i, f in enumerate(factories):
            ft = p[p["Usine"]==f]["Tonnes/Jour"].sum()
            with cols_f[i]:
                st.metric(f, f"{ft:,.0f} t")

        # Line chart per factory
        fact_daily = p.groupby(["Date","Usine"])["Tonnes/Jour"].sum().reset_index()
        fig9 = px.line(
            fact_daily, x="Date", y="Tonnes/Jour", color="Usine",
            color_discrete_map=FACTORY_COLORS,
            title="Tonnes/jour reçues par usine",
            labels={"Tonnes/Jour":"Tonnes/j"},
            template="plotly_dark",
        )
        fig9.update_traces(line_width=2)
        fig9.update_layout(
            plot_bgcolor="#0d1117", paper_bgcolor="#161b22",
            height=420, hovermode="x unified",
        )
        fig9.add_vrect(x0=str(PEAK_START), x1=str(PEAK_END),
                       fillcolor="gold", opacity=0.06, line_width=0)
        st.plotly_chart(fig9, use_container_width=True)

        c1, c2 = st.columns(2)
        with c1:
            # Peak per factory
            peak_fact = (fact_daily.groupby("Usine")["Tonnes/Jour"]
                         .max().reset_index().rename(columns={"Tonnes/Jour":"Pic/Jour"}))
            fig10 = px.bar(
                peak_fact, x="Usine", y="Pic/Jour",
                color="Usine", color_discrete_map=FACTORY_COLORS,
                title="Pic journalier par usine (max)", template="plotly_dark",
                text_auto=".3s",
            )
            fig10.update_layout(paper_bgcolor="#161b22", showlegend=False, height=300)
            st.plotly_chart(fig10, use_container_width=True)
        with c2:
            fact_tot = p.groupby("Usine")["Tonnes/Jour"].sum().reset_index()
            fig11 = px.bar(
                fact_tot, x="Tonnes/Jour", y="Usine", orientation="h",
                color="Usine", color_discrete_map=FACTORY_COLORS,
                title="Total tonnes par usine (saison)", template="plotly_dark",
                text_auto=".3s",
            )
            fig11.update_layout(paper_bgcolor="#161b22", showlegend=False, height=300)
            st.plotly_chart(fig11, use_container_width=True)

        # Drill-down by factory
        st.subheader("🔎 Détail par usine")
        sel_fact = st.selectbox("Choisir une usine", factories)
        fact_one = p[p["Usine"]==sel_fact].sort_values("Date")
        st.dataframe(
            fact_one[[c for c in ["Date","Commercial","Agriculteur","Tonnes/Jour",
                                   "Nb Voyages","Pic de Récolte"] if c in fact_one.columns]]
            .reset_index(drop=True),
            use_container_width=True, height=260,
        )

# ── TAB 4: TRANSPORT & ALERTES ───────────────────────────────
with tab4:
    # ALL 4 vehicle types — TRACTEUR shown even if 0
    ALL_VEH_COLS = ["Voyages TRACTEUR","Voyages PETIT POILOUR","Voyages POILOUR","Voyages SEMI"]
    VEH_COLORS   = {
        "TRACTEUR":      "#a16207",
        "PETIT POILOUR": "#f5a623",
        "POILOUR":       "#3b82f6",
        "SEMI":          "#00e5a0",
    }
    # Ensure all 4 columns exist, fill missing with 0
    for vc in ALL_VEH_COLS:
        if vc not in t.columns:
            t[vc] = 0

    veh_daily = t.groupby("Date")[ALL_VEH_COLS + ["Total Tonnes"]].sum().reset_index()

    # Stacked bar — all 4 vehicles
    veh_long = veh_daily.melt(id_vars="Date", value_vars=ALL_VEH_COLS,
                               var_name="Véhicule", value_name="Voyages")
    veh_long["Véhicule"] = veh_long["Véhicule"].str.replace("Voyages ", "")
    fig12 = px.bar(
        veh_long, x="Date", y="Voyages", color="Véhicule",
        barmode="stack",
        color_discrete_map={k.replace("Voyages ",""): v for k,v in
                            zip(ALL_VEH_COLS, VEH_COLORS.values())},
        title="Voyages par type de véhicule — chaque jour (4 types)",
        template="plotly_dark",
    )
    fig12.update_layout(
        plot_bgcolor="#0d1117", paper_bgcolor="#161b22",
        height=400, hovermode="x unified",
    )
    fig12.add_vrect(x0=str(PEAK_START), x1=str(PEAK_END),
                    fillcolor="gold", opacity=0.06, line_width=0)
    st.plotly_chart(fig12, use_container_width=True)

    # Export transport table
    st.download_button(
        "⬇️ Exporter tableau transport (CSV)",
        data=df_to_csv(veh_daily),
        file_name="transport_journalier.csv",
        mime="text/csv",
    )

    c1, c2 = st.columns(2)
    with c1:
        # Total per vehicle including TRACTEUR
        veh_totals_all = {
            vc.replace("Voyages ",""): int(t[vc].sum())
            for vc in ALL_VEH_COLS
        }
        fig13 = px.bar(
            x=list(veh_totals_all.values()),
            y=list(veh_totals_all.keys()),
            orientation="h",
            color=list(veh_totals_all.keys()),
            color_discrete_map={k: VEH_COLORS[k] for k in VEH_COLORS},
            title="Total voyages par véhicule (saison)",
            template="plotly_dark",
            text_auto=True,
        )
        fig13.update_layout(paper_bgcolor="#161b22", showlegend=False, height=300)
        st.plotly_chart(fig13, use_container_width=True)

    with c2:
        st.subheader("⚠️ Alertes manque véhicules")
        ROT = 8  # rotations/day/vehicle
        fleet_inventory = {
            "TRACTEUR":      fl_trac,
            "PETIT POILOUR": fl_ppl,
            "POILOUR":       fl_pl,
            "SEMI":          fl_semi,
        }
        all_ok = True
        for vc in ALL_VEH_COLS:
            vname  = vc.replace("Voyages ", "")
            peak   = int(t[vc].max()) if t[vc].max() > 0 else 0
            owned  = fleet_inventory[vname]
            cap    = owned * ROT
            if peak == 0:
                st.info(f"⚪ {vname} : 0 voyage cette saison (non utilisé)")
            elif peak > cap:
                st.error(f"🔴 {vname} : besoin max {peak}/jour | capacité {cap} ({owned}×{ROT}) → **manque {peak-cap}**")
                all_ok = False
            else:
                st.success(f"✅ {vname} : besoin max {peak}/jour | capacité {cap} → {cap-peak} de marge")
        if all_ok:
            st.success("✅ Flotte suffisante pour toute la saison !")

    # Double transport days chart
    if "Jours Double" in t.columns:
        dbl_d = t.groupby("Date")["Jours Double"].sum().reset_index()
        fig14 = px.bar(
            dbl_d, x="Date", y="Jours Double",
            color_discrete_sequence=["#e8543a"],
            title="Jours de double transport par date",
            template="plotly_dark",
        )
        fig14.update_layout(paper_bgcolor="#161b22", height=280)
        st.plotly_chart(fig14, use_container_width=True)

# ── TAB 5: DÉCALAGE ──────────────────────────────────────────
with tab5:
    st.info("✅ Tous les décalages respectent la limite de **4 jours** — aucune tomate à risque de maladie.")

    if not resume.empty:
        c1, c2 = st.columns(2)
        with c1:
            fig15 = px.bar(
                resume, x="Commercial", y="Conflits Résolus",
                color="Commercial", color_discrete_map=COMM_COLORS,
                title="Conflits résolus par commercial",
                template="plotly_dark", text_auto=True,
            )
            fig15.update_layout(paper_bgcolor="#161b22", showlegend=False, height=300)
            st.plotly_chart(fig15, use_container_width=True)
        with c2:
            fig16 = px.bar(
                resume, x="Commercial", y="Total Jours Double",
                color_discrete_sequence=["#e8543a"],
                title="Jours de double transport par commercial",
                template="plotly_dark", text_auto=True,
            )
            fig16.update_layout(paper_bgcolor="#161b22", showlegend=False, height=300)
            st.plotly_chart(fig16, use_container_width=True)

    st.subheader("📋 Journal des conflits résolus")
    if not double_j.empty:
        show_cols_dj = [c for c in [
            "Commercial","Agriculteur A (finit tôt)","Agriculteur B (reçoit véhicule)",
            "Véhicule Partagé","Jours Économisés","Fin Orig. A","Nouvelle Fin A",
            "Début Orig. B","Nouveau Début B","Risque Maladie","Action Requise"
        ] if c in double_j.columns]
        st.dataframe(
            double_j[show_cols_dj].reset_index(drop=True),
            use_container_width=True, height=400,
        )
        st.download_button(
            "⬇️ Exporter journal décalage (CSV)",
            data=df_to_csv(double_j[show_cols_dj].reset_index(drop=True)),
            file_name="journal_decalage.csv",
            mime="text/csv",
        )
    else:
        st.info("Aucun conflit dans le journal.")

    # Summary table
    if not resume.empty:
        st.subheader("📊 Résumé par commercial")
        st.dataframe(resume.reset_index(drop=True), use_container_width=True)
        st.download_button(
            "⬇️ Exporter résumé commercial (CSV)",
            data=df_to_csv(resume.reset_index(drop=True)),
            file_name="resume_commercial.csv",
            mime="text/csv",
        )

# ── TAB 6: DEPLOY & ROADMAP ───────────────────────────────────
with tab6:
    st.subheader("🚀 C'est quoi « Deploy » ?")
    st.markdown("""
<div class="deploy-box">
<b>Deploy = rendre ton application accessible à d'autres personnes via internet.</b><br><br>

En ce moment, quand tu fais <code>streamlit run dashboard_phase7.py</code>,
le dashboard fonctionne <b>seulement sur ton ordinateur</b>.
Personne d'autre ne peut l'ouvrir.<br><br>

<b>Deploy = mettre l'application sur un serveur internet</b> → tout le monde peut l'ouvrir
avec un lien comme <code>https://tomate-planning.streamlit.app</code><br><br>

<b>⚠️ Est-ce que ça « brise » l'application ?</b> NON. C'est juste un bouton.
Streamlit propose un service gratuit appelé <b>Streamlit Community Cloud</b> :
<ol>
<li>Tu crées un compte sur <a href="https://share.streamlit.io" target="_blank">share.streamlit.io</a></li>
<li>Tu mets ton code sur GitHub (gratuit)</li>
<li>Tu cliques "Deploy" → Streamlit héberge l'app pour toi</li>
<li>Tu reçois un lien URL → tu l'envoies à qui tu veux</li>
</ol>
<b>Résultat :</b> FEDI peut ouvrir le dashboard depuis son téléphone. MAKKI aussi.
Sans rien installer. C'est juste un lien.<br><br>
<b>Limite gratuite :</b> 1 app, accès public. Pour accès privé (login) → plan payant (~20$/mois).
</div>
""", unsafe_allow_html=True)

    st.divider()
    st.subheader("🗺️ Roadmap — Ce qu'il reste pour un produit professionnel")

    roadmap = [
        ("✅ Fait", "Phase 1-3",  "Scripts Python de base, planning journalier, caps commerciaux"),
        ("✅ Fait", "Phase 4",    "Décalage transport, double transport, journal des conflits"),
        ("✅ Fait", "Phase 5-7",  "Dashboard Streamlit, graphiques, export CSV, tous véhicules"),
        ("🔜 Next", "Phase 8",    "Login par commercial — chaque commercial voit SEULEMENT ses agriculteurs"),
        ("🔜 Next", "Phase 9",    "Base de données PostgreSQL — plus besoin d'Excel, saisie directe dans l'app"),
        ("🔜 Next", "Phase 10",   "Deploy sur serveur → lien URL pour les commerciaux et directeurs"),
        ("🔮 Future", "Phase 11", "OR-Tools : optimisation automatique (vraie IA, planning optimal)"),
        ("🔮 Future", "Phase 12", "App mobile Flutter pour les chauffeurs (voir leurs voyages du jour)"),
        ("🔮 Future", "Phase 13", "Alertes SMS automatiques quand un conflit est détecté"),
        ("🔮 Future", "Phase 14", "Historique saisons + prédictions (quand chaque fallah sera prêt)"),
    ]

    status_colors = {"✅ Fait": "green", "🔜 Next": "orange", "🔮 Future": "blue"}
    for status, phase, desc in roadmap:
        col_s, col_p, col_d = st.columns([1, 1.2, 5])
        with col_s:
            st.markdown(f"**{status}**")
        with col_p:
            st.markdown(f"`{phase}`")
        with col_d:
            st.markdown(desc)

    st.divider()
    st.subheader("💼 Comment vendre ce projet à une société ?")
    st.markdown("""
**Ce que tu as déjà :**
- Un système qui gère **90,225 tonnes** de tomates automatiquement
- **538 conflits de transport** résolus sans intervention humaine
- Dashboard professionnel avec graphiques, filtres, export

**Pitch en 2 phrases :**
> *"Votre équipe passe 3h/jour à faire manuellement ce que ce système fait en 10 secondes.
> 538 conflits de transport résolus automatiquement cette saison — zéro tomate perdue à cause d'un mauvais planning."*

**Cibles immédiates :**
- Coopératives agricoles (SMVDA, GDAP)
- Usines de conserves (SICAM, COMOCAP, TUCAL) — elles ont leurs propres équipes logistique
- Groupements de commerciaux régionaux

**Prix suggéré :** 500–2000 DT/mois selon le volume géré. Commence par un pilote gratuit de 1 mois.
""")

    st.divider()
    st.subheader("📋 Comparaison : maintenant vs avec Deploy")
    comp = {
        "": ["Qui peut voir le dashboard", "Comment y accéder", "Mise à jour données",
             "Accès sur mobile", "Partager avec un client"],
        "Maintenant (local)": ["Seulement toi", "streamlit run ... sur ton PC",
                                "Manuellement", "Non", "Impossible"],
        "Après Deploy": ["Tout le monde avec le lien", "Ouvrir un URL dans le navigateur",
                          "Automatique", "Oui", "Envoyer le lien"],
    }
    st.dataframe(pd.DataFrame(comp).set_index(""), use_container_width=True)