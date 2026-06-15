# ============================================================
# ONGLET COMPARAISON — Plans rectifiés vs OR-Tools
# Fichier : comparaison_tab.py
# A placer dans le meme dossier que dashboard_phase10.py
# ============================================================

import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
import io

COMM_COLORS = {
    "FEDI":           "#3b82f6",
    "MAKKI BEN SALAH":"#00e5a0",
    "KHALIL":         "#f5a623",
    "ACHREF AJLANI":  "#8b5cf6",
    "JILANI OBAY":    "#e8543a",
}

MANUAL_STATS = {
    "FEDI":           {"total":32736,"max_j":1216,"avg_j":595,"n_jours":55,"n_agri":73,"max_agri":200,"avg_days":13.2,"cap":1300},
    "MAKKI BEN SALAH":{"total":24310,"max_j":1315,"avg_j":552,"n_jours":44,"n_agri":82,"max_agri":80, "avg_days":11.0,"cap":1200},
    "KHALIL":         {"total":25290,"max_j":1500,"avg_j":550,"n_jours":46,"n_agri":27,"max_agri":120,"avg_days":16.4,"cap":1100},
    "ACHREF AJLANI":  {"total":17486,"max_j":570, "avg_j":287,"n_jours":61,"n_agri":68,"max_agri":120,"avg_days":5.3, "cap":700},
    "JILANI OBAY":    {"total":7000, "max_j":450, "avg_j":280,"n_jours":25,"n_agri":11,"max_agri":110,"avg_days":12.9,"cap":150},
}

USINE_CAPS = {"SICAM":1300,"TUCAL":700,"COMOCAP":700,"ELFALLEH":150,"ABIDA":200}

USINE_DAILY_PDF = {
    "SICAM":[0,0,140,170,220,220,260,320,320,415,515,625,560,910,1090,1170,1285,
             1425,1410,1540,1490,1440,1440,1455,1650,1600,1595,1550,1385,1345,1125,
             1000,935,835,840,575,456,660,750,730,660,760,535,615,580,610,520,470,
             530,500,380,320,290,240,200,140,170,140,210,180,150,150,150,120,120,120,120,120],
    "TUCAL":[0,0,0,0,15,15,15,85,85,105,230,235,285,465,560,590,590,625,610,700,
             670,690,645,660,620,655,585,665,645,595,490,500,405,405,395,345,365,475,
             490,410,430,380,380,350,350,330,315,310,340,350,300,270,200,120,90,90,
             0,0,0,0,0,0,0,0,0,0,0],
    "COMOCAP":[0,0,0,0,0,15,15,15,65,145,195,305,305,345,375,425,435,480,515,615,
               760,760,755,785,775,785,826,875,805,760,755,625,610,650,595,540,500,
               490,490,455,455,390,340,280,200,150,100,100,100,80,80,80,80,70,40,
               0,0,0,0,0,0,0,0,0,0,0],
    "ELFALLEH":[0,0,0,0,0,0,0,0,0,20,50,50,40,40,30,40,60,60,60,60,60,60,60,60,
                40,40,40,40,40,50,40,50,75,60,60,60,60,80,80,80,80,80,80,70,70,60,
                20,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0],
    "ABIDA":[0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,50,100,150,200,200,210,
             220,230,240,250,260,270,280,290,300,300,310,310,315,315,320,310,300,280,
             260,230,200,170,140,110,80,60,40,30,20,0,0,0,0,0,0,0,0,0,0],
}


def _detect_comm_from_filename(filename):
    fn = filename.upper()
    mapping = {
        "FEDI":"FEDI","MEKKI":"MAKKI BEN SALAH","MAKKI":"MAKKI BEN SALAH",
        "KHALIL":"KHALIL","ACHRAF":"ACHREF AJLANI","ACHREF":"ACHREF AJLANI",
        "JILENI":"JILANI OBAY","JILANI":"JILANI OBAY",
    }
    for key, comm in mapping.items():
        if key in fn:
            return comm
    return None


def _parse_rectification(uploaded_file):
    try:
        df = pd.read_excel(uploaded_file, header=0)
    except Exception as e:
        return None, None, str(e)

    header_row = 0
    for i in range(min(5, len(df))):
        row_vals = [str(v).strip().lower() for v in df.iloc[i].values]
        if any("agriculteur" in v for v in row_vals):
            header_row = i
            break
    if header_row > 0:
        uploaded_file.seek(0)
        df = pd.read_excel(uploaded_file, header=header_row)

    cols = df.columns.tolist()
    col_comm    = next((c for c in cols if "responsable" in str(c).lower() or "commercial" in str(c).lower()), cols[0] if cols else None)
    col_agri    = next((c for c in cols if "agriculteur" in str(c).lower()), None)
    col_tonnage = next((c for c in cols if "tonnage" in str(c).lower()), None)

    parsed_dates = []
    for c in cols[5:]:
        if isinstance(c, pd.Timestamp):
            parsed_dates.append((c, c)); continue
        cs = str(c).strip()
        if "/" in cs and len(cs) <= 5:
            try:
                p = cs.split("/")
                d = pd.Timestamp(f"2026-{int(p[1]):02d}-{int(p[0]):02d}")
                parsed_dates.append((c, d)); continue
            except Exception:
                pass
        try:
            d = pd.to_datetime(str(c).split(" ")[0], errors="coerce")
            if pd.notna(d) and d.year == 2026:
                parsed_dates.append((c, d))
        except Exception:
            pass

    if not parsed_dates or col_agri is None:
        return None, None, "Format non reconnu"

    rows = []
    comm_detected = None
    for _, row in df.iterrows():
        comm = str(row.get(col_comm, "") or "").strip()
        agri = str(row.get(col_agri, "") or "").strip()
        if not agri or agri.upper() in ("NAN","","AGRICULTEUR","TOTAL"): continue
        if "sous-total" in agri.lower() or "total" in agri.lower(): continue
        if comm and comm.upper() not in ("NAN",""):
            for known in MANUAL_STATS:
                if known.upper() in comm.upper() or comm.upper() in known.upper():
                    comm_detected = known; break
        ton = pd.to_numeric(row.get(col_tonnage, 0), errors="coerce") if col_tonnage else 0
        for orig_col, date in parsed_dates:
            val = pd.to_numeric(row.get(orig_col, 0), errors="coerce")
            if pd.notna(val) and val > 0:
                rows.append({"agriculteur": agri, "date": date, "tonnes": float(val),
                             "tonnage_total": float(ton) if pd.notna(ton) else 0})

    if not rows:
        return comm_detected, None, "Aucune donnee trouvee"

    df_rows = pd.DataFrame(rows)
    daily = df_rows.groupby("date")["tonnes"].sum().reset_index().sort_values("date")
    daily_dict = {str(r["date"].date()): float(r["tonnes"]) for _, r in daily.iterrows()}
    return comm_detected, daily_dict, df_rows


def _ortools_profile(comm, planning_df):
    if planning_df is not None and not planning_df.empty and "Commercial" in planning_df.columns:
        sub = planning_df[planning_df["Commercial"] == comm].copy()
        if not sub.empty and "Date" in sub.columns:
            sub["Date"] = pd.to_datetime(sub["Date"], errors="coerce")
            daily = sub.groupby("Date")["Tonnes/Jour"].sum().reset_index()
            return {str(r["Date"].date()): float(r["Tonnes/Jour"]) for _, r in daily.iterrows()}
    return {}


def _build_reference_curve(comm):
    stat = MANUAL_STATS[comm]
    n = stat["n_jours"]
    mx = stat["max_j"]
    starts = {
        "FEDI":"2026-06-28","MAKKI BEN SALAH":"2026-06-22",
        "KHALIL":"2026-06-20","ACHREF AJLANI":"2026-06-20","JILANI OBAY":"2026-07-23",
    }
    start = pd.Timestamp(starts.get(comm, "2026-06-20"))
    p1 = int(n * 0.30); p2 = int(n * 0.50); p3 = n - p1 - p2
    result = {}
    for i in range(n):
        if i < p1:
            v = mx * 0.2 + (mx * 0.6) * (i / max(p1, 1))
        elif i < p1 + p2:
            v = mx * 0.8 + (mx * 0.2) * ((i - p1) / max(p2, 1))
        else:
            v = mx * (1 - ((i - p1 - p2) / max(p3, 1)) * 0.8)
        d = start + pd.Timedelta(days=i)
        result[str(d.date())] = round(max(0, v))
    return result


def _kpi_row(comm, stat_m, ort_dict, man_dict):
    man_total = sum(man_dict.values()) if man_dict else stat_m["total"]
    man_max   = max(man_dict.values()) if man_dict else stat_m["max_j"]
    ort_total = sum(ort_dict.values()) if ort_dict else 0
    ort_max   = max(ort_dict.values()) if ort_dict else 0
    k1, k2, k3, k4 = st.columns(4)
    delta_t = int(man_total - ort_total)
    delta_m = int(man_max - ort_max)
    k1.metric("Total Manuel", f"{int(man_total):,}t".replace(",", " "),
              delta=f"{delta_t:+,}t vs OR-T".replace(",", " "), delta_color="off")
    k2.metric("Max/j Manuel", f"{int(man_max)}t",
              delta=f"{delta_m:+d}t vs OR-T",
              delta_color="inverse" if man_max > stat_m["cap"] else "off")
    k3.metric("Jours actifs", f"{len(man_dict)}j" if man_dict else f"{stat_m['n_jours']}j")
    k4.metric("Cap PIC", f"{stat_m['cap']}t/j",
              delta="OK" if man_max <= stat_m["cap"] else "depasse",
              delta_color="normal" if man_max <= stat_m["cap"] else "inverse")


def _build_comparison_chart(comm, man_dict, ort_dict, color, cap):
    all_dates_str = sorted(set(list(man_dict.keys()) + list(ort_dict.keys())))
    all_dates = [pd.Timestamp(d) for d in all_dates_str]
    man_vals = [man_dict.get(d, 0) for d in all_dates_str]
    ort_vals = [ort_dict.get(d, 0) for d in all_dates_str]

    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=all_dates, y=man_vals, name="Plan manuel rectifie",
        line=dict(color=color, width=2.5),
        fill="tozeroy", fillcolor=color + "14", mode="lines",
    ))
    if any(v > 0 for v in ort_vals):
        fig.add_trace(go.Scatter(
            x=all_dates, y=ort_vals, name="OR-Tools (actuel)",
            line=dict(color="#ffffff", width=1.5, dash="dot"), mode="lines",
        ))
    fig.add_hline(y=cap, line_dash="dash", line_color="#e8543a", line_width=1,
                  annotation_text=f"Cap PIC: {cap}t/j",
                  annotation_position="top right", annotation_font_color="#e8543a")
    fig.add_vrect(x0=pd.Timestamp("2026-07-01"), x1=pd.Timestamp("2026-07-15"),
                  fillcolor="rgba(245,166,35,0.06)", line_width=0,
                  annotation_text="PIC", annotation_position="top left",
                  annotation_font_color="#f5a623")
    fig.update_layout(
        title=f"{comm} — Profil journalier : Plan rectifie vs OR-Tools",
        template="plotly_dark", plot_bgcolor="#0d1117", paper_bgcolor="#161b22",
        height=380, hovermode="closest",
        legend=dict(orientation="h", yanchor="bottom", y=1.02),
        xaxis=dict(title="Date", gridcolor="#21262d"),
        yaxis=dict(title="Tonnes/jour", gridcolor="#21262d"),
    )
    return fig


def render_comparaison_tab(planning_df=None, df_to_xlsx_styled=None):
    """Point d'entree principal de l'onglet comparaison."""

    st.markdown("""
    <div style='background:#1a2332;border:1px solid #21262d;border-radius:12px;
    padding:16px 20px;margin-bottom:20px'>
      <div style='font-size:1.1rem;font-weight:700;color:#f0f6fc;margin-bottom:6px'>
        Comparaison Plans — OR-Tools vs Plans Rectifies Manuellement
      </div>
      <div style='font-size:.82rem;color:#8b949e'>
        Deposez les fichiers Excel rectifies par les commerciaux pour visualiser les ecarts
        avec le plan OR-Tools. Formats: <code>Rectification_Fedi_13_JUIN.xlsx</code>, etc.
      </div>
    </div>
    """, unsafe_allow_html=True)

    if "comp_uploaded" not in st.session_state:
        st.session_state["comp_uploaded"] = {}
    if "comp_raw" not in st.session_state:
        st.session_state["comp_raw"] = {}

    # ── ZONE UPLOAD ─────────────────────────────────────────
    st.subheader("Deposer les fichiers rectifies")
    col_up, col_status = st.columns([3, 2])

    with col_up:
        uploaded_files = st.file_uploader(
            "Fichiers Excel rectifies",
            type=["xlsx","xls"],
            accept_multiple_files=True,
            label_visibility="collapsed",
            help="Glisser-deposer les 5 fichiers Excel des commerciaux",
        )

    if uploaded_files:
        for f in uploaded_files:
            comm_from_name = _detect_comm_from_filename(f.name)
            comm_from_data, daily_dict, raw_or_err = _parse_rectification(f)
            comm = comm_from_name or comm_from_data
            if comm and daily_dict:
                st.session_state["comp_uploaded"][comm] = daily_dict
                if isinstance(raw_or_err, pd.DataFrame):
                    st.session_state["comp_raw"][comm] = raw_or_err
                st.success(f"Charge : {comm} ({len(daily_dict)} jours, {sum(daily_dict.values()):,.0f}t)".replace(",", " "))
            elif comm and not daily_dict:
                st.warning(f"{f.name} : {raw_or_err}")
            else:
                st.warning(f"{f.name} : commercial non detecte dans le nom du fichier.")

    with col_status:
        st.markdown("**Statut des fichiers :**")
        for c in MANUAL_STATS:
            if c in st.session_state["comp_uploaded"]:
                n = len(st.session_state["comp_uploaded"][c])
                t = sum(st.session_state["comp_uploaded"][c].values())
                st.success(f"{c.split()[0]} — {n}j | {int(t):,}t".replace(",", " "))
            else:
                st.info(f"{c.split()[0]} — reference interne")
        if st.session_state["comp_uploaded"]:
            if st.button("Effacer uploads", use_container_width=True):
                st.session_state["comp_uploaded"] = {}
                st.session_state["comp_raw"] = {}
                st.rerun()

    if not st.session_state["comp_uploaded"]:
        st.info("Aucun fichier uploade — courbes basees sur les donnees de reference du 13 juin 2026.")

    st.divider()

    # ── ONGLETS INTERNES ────────────────────────────────────
    c1, c2, c3, c4 = st.tabs([
        "Courbes par commercial",
        "Par usine",
        "Statistiques globales",
        "Corrections optimizer",
    ])

    # ─ C1 : COURBES PAR COMMERCIAL ──────────────────────────
    with c1:
        sel_comm = st.selectbox("Choisir un commercial", list(MANUAL_STATS.keys()), key="comp_sel_comm")
        color = COMM_COLORS[sel_comm]
        stat_m = MANUAL_STATS[sel_comm]

        if sel_comm in st.session_state["comp_uploaded"]:
            man_dict = st.session_state["comp_uploaded"][sel_comm]
            src_lbl = "fichier uploade"
        else:
            man_dict = _build_reference_curve(sel_comm)
            src_lbl = "donnees de reference (13/06/2026)"

        ort_dict = _ortools_profile(sel_comm, planning_df)

        _kpi_row(sel_comm, stat_m, ort_dict, man_dict)
        st.caption(f"Source plan rectifie : {src_lbl}")

        fig = _build_comparison_chart(sel_comm, man_dict, ort_dict, color, stat_m["cap"])
        st.plotly_chart(fig, use_container_width=True)

        # Vue agriculteur si fichier uploade
        if sel_comm in st.session_state["comp_raw"]:
            st.markdown("---")
            st.markdown(f"**Profil par agriculteur — {sel_comm}**")
            df_raw = st.session_state["comp_raw"][sel_comm]
            agri_stats = df_raw.groupby("agriculteur").agg(
                total=("tonnes","sum"), max_jour=("tonnes","max"), jours=("date","nunique")
            ).reset_index().sort_values("total", ascending=False)
            agri_stats.columns = ["Agriculteur","Total (t)","Max/jour (t)","Jours actifs"]
            agri_stats[["Total (t)","Max/jour (t)"]] = agri_stats[["Total (t)","Max/jour (t)"]].round(0).astype(int)

            n_agri = len(agri_stats)
            fig_agri = px.bar(
                agri_stats, x="Total (t)", y="Agriculteur", orientation="h",
                height=max(350, n_agri * 28 + 80),
                color="Total (t)", color_continuous_scale="Viridis",
                title=f"Tonnage par agriculteur — {sel_comm} ({n_agri} agri)",
                template="plotly_dark", text="Total (t)",
            )
            fig_agri.update_traces(textposition="outside", texttemplate="%{x:.0f}t",
                                   textfont=dict(size=10))
            fig_agri.update_layout(paper_bgcolor="#161b22", showlegend=False,
                                   margin=dict(l=230, r=60, t=50, b=40))
            st.plotly_chart(fig_agri, use_container_width=True)

    # ─ C2 : PAR USINE ────────────────────────────────────────
    with c2:
        sel_usine = st.selectbox("Choisir une usine",
                                 ["SICAM","TUCAL","COMOCAP","ELFALLEH","ABIDA"],
                                 key="comp_usine_sel")
        cap = USINE_CAPS[sel_usine]
        data_u = USINE_DAILY_PDF.get(sel_usine, [])
        dates_u = [pd.Timestamp("2026-06-20") + pd.Timedelta(days=i) for i in range(len(data_u))]

        total_u = sum(data_u); max_u = max(data_u) if data_u else 0
        jours_u = sum(1 for v in data_u if v > 0)
        pic_data = [data_u[i] for i in range(11, 26) if i < len(data_u)]
        max_pic  = max(pic_data) if pic_data else 0

        k1, k2, k3, k4, k5 = st.columns(5)
        k1.metric("Total saison", f"{total_u:,}t".replace(",", " "))
        k2.metric("Max journalier", f"{max_u}t",
                  delta="depasse" if max_u > cap else "OK",
                  delta_color="inverse" if max_u > cap else "normal")
        k3.metric("Cap officiel", f"{cap}t/j")
        k4.metric("Max PIC 1-15/07", f"{max_pic}t")
        k5.metric("Jours actifs", f"{jours_u}j")

        fig_u = go.Figure()
        fig_u.add_trace(go.Bar(x=dates_u, y=data_u, name="Plan rectifie",
                               marker_color="#3b82f6", marker_line_width=0))
        fig_u.add_hline(y=cap, line_dash="dash", line_color="#e8543a", line_width=1.5,
                        annotation_text=f"Cap: {cap}t/j",
                        annotation_position="top right", annotation_font_color="#e8543a")
        fig_u.add_vrect(x0=pd.Timestamp("2026-07-01"), x1=pd.Timestamp("2026-07-15"),
                        fillcolor="rgba(245,166,35,0.08)", line_width=0,
                        annotation_text="PIC", annotation_position="top left",
                        annotation_font_color="#f5a623")

        if planning_df is not None and not planning_df.empty and "Usine" in planning_df.columns:
            sub_u = planning_df[planning_df["Usine"] == sel_usine].copy()
            if not sub_u.empty:
                sub_u["Date"] = pd.to_datetime(sub_u["Date"], errors="coerce")
                ort_u = sub_u.groupby("Date")["Tonnes/Jour"].sum().reset_index()
                fig_u.add_trace(go.Scatter(
                    x=ort_u["Date"], y=ort_u["Tonnes/Jour"], name="OR-Tools",
                    line=dict(color="#f5a623", width=2, dash="dash"), mode="lines",
                ))

        fig_u.update_layout(
            title=f"{sel_usine} — Reception journaliere (plan rectifie vs OR-Tools)",
            template="plotly_dark", plot_bgcolor="#0d1117", paper_bgcolor="#161b22",
            height=380, hovermode="closest",
            legend=dict(orientation="h", yanchor="bottom", y=1.02),
        )
        st.plotly_chart(fig_u, use_container_width=True)

        _rules = {
            "SICAM":   ["Montee progressive juin (140-625t)","Pic 1-15 juillet (910-1650t)",
                        "Declin rapide aout","Gros agris: 80-200t/j","Petits agris: 15-30t/j"],
            "TUCAL":   ["Demarrage 24 juin","Plateau 590-700t/j juillet",
                        "STE 428: 120t/j x21j","STE BACCARA: 80t/j x33j","MAKKI: 15-35t/j"],
            "COMOCAP": ["Pic tardif 27/07 (875t)","TRC/PPL/PL: 10-20t/j majoritaires",
                        "YASIN MNASRI SEMI: 60-120t/j","STE 428: 130t/j x15j","Declin net aout"],
            "ELFALLEH":["Cap 150t/j: JAMAIS depasse","Max plan: 80t/j (aout)",
                        "HABIB BELWEAR: 10-60t montee","Tous agris: 30-60t/j max","Fenetres: 20-30j"],
            "ABIDA":   ["Demarrage tardif (juillet)","Montee 50-320t/j",
                        "SAMIR ATTIYA: 50t/j constant","ACHREF DIVERS: 120t/j","Declin aout"],
        }
        for rule in _rules.get(sel_usine, []):
            st.markdown(f"• {rule}")

    # ─ C3 : STATISTIQUES GLOBALES ────────────────────────────
    with c3:
        st.markdown("### Comparaison globale — tous commerciaux")

        rows_cmp = []
        man_totals_list = []
        ort_totals_list = []
        comms_list = list(MANUAL_STATS.keys())

        for comm in comms_list:
            stat = MANUAL_STATS[comm]
            ort_d = _ortools_profile(comm, planning_df)
            ort_total = sum(ort_d.values()) if ort_d else 0
            ort_max   = max(ort_d.values()) if ort_d else 0

            if comm in st.session_state["comp_uploaded"]:
                man_d = st.session_state["comp_uploaded"][comm]
                man_total = sum(man_d.values()); man_max = max(man_d.values()) if man_d else 0
                man_jours = len(man_d); src = "uploade"
            else:
                man_total = stat["total"]; man_max = stat["max_j"]
                man_jours = stat["n_jours"]; src = "reference"

            man_totals_list.append(man_total)
            ort_totals_list.append(ort_total)
            rows_cmp.append({
                "Commercial": comm, "Source": src,
                "Total Manuel (t)": int(man_total), "Total OR-Tools (t)": int(ort_total),
                "Ecart total (t)": f"{int(man_total-ort_total):+,}".replace(",", " "),
                "Max/j Manuel": f"{int(man_max)}t", "Max/j OR-Tools": f"{int(ort_max)}t",
                "Jours Manuel": man_jours, "Cap PIC": f"{stat['cap']}t/j",
            })

        df_cmp = pd.DataFrame(rows_cmp)
        st.dataframe(df_cmp, use_container_width=True, hide_index=True)

        fig_cmp = go.Figure()
        fig_cmp.add_trace(go.Bar(name="Plan Manuel rectifie", x=comms_list, y=man_totals_list,
                                 marker_color="#e8543a",
                                 text=[f"{int(v):,}t".replace(",", " ") for v in man_totals_list],
                                 textposition="outside"))
        fig_cmp.add_trace(go.Bar(name="OR-Tools", x=comms_list, y=ort_totals_list,
                                 marker_color="#3b82f6",
                                 text=[f"{int(v):,}t".replace(",", " ") for v in ort_totals_list],
                                 textposition="outside"))
        fig_cmp.update_layout(
            barmode="group", title="Tonnage total saison — Plan rectifie vs OR-Tools",
            template="plotly_dark", paper_bgcolor="#161b22", plot_bgcolor="#0d1117",
            height=400, legend=dict(orientation="h", yanchor="bottom", y=1.02),
            yaxis_title="Tonnes",
        )
        st.plotly_chart(fig_cmp, use_container_width=True)

        # 4 causes racines
        st.markdown("---")
        st.markdown("### 4 causes racines des ecarts OR-Tools vs Manuel")
        causes = [
            ("1 — Borne journaliere trop large","#e8543a",
             "OR-Tools met 870t/j pour un seul agriculteur FEDI (reel: 200t max). "
             "ub_day x 1.5 trop permissif -> livraisons anormalement grosses."),
            ("2 — Arrondi 10t supprime micro-livraisons","#f5a623",
             "Le plan manuel utilise 10t, 15t, 20t pour COMOCAP/TUCAL. "
             "Arrondi a 10t minimum supprime les vraies petites livraisons."),
            ("3 — FACTORY_OVERFLOW_WEIGHT trop faible","#8b5cf6",
             "Poids 500 insuffisant: ELFALLEH 130t OR-Tools vs 60t/j reel. "
             "Petites usines pas assez protegees contre depassements."),
            ("4 — Pas de cap max par agriculteur","#00e5a0",
             "Un agri avec 200t total ne doit pas livrer 60t/j. "
             "Regle terrain: <200t=30t/j max, <500t=60t/j max, >500t=200t/j max."),
        ]
        for title, color, desc in causes:
            st.markdown(f"""
            <div style='border-left:4px solid {color};padding:10px 16px;
            margin-bottom:8px;background:#161b22;border-radius:0 8px 8px 0'>
              <div style='font-size:13px;font-weight:600;color:{color}'>{title}</div>
              <div style='font-size:12px;color:#8b949e;margin-top:4px'>{desc}</div>
            </div>""", unsafe_allow_html=True)

        if df_to_xlsx_styled:
            st.download_button("Exporter tableau comparatif (Excel)",
                data=df_to_xlsx_styled(df_cmp, sheet_name="Comparaison"),
                file_name="comparaison_ortools_vs_manuel.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    # ─ C4 : CORRECTIONS OPTIMIZER ───────────────────────────
    with c4:
        st.markdown("### Corrections a appliquer dans optimizer_v2.py")
        st.caption("Copie chaque correction dans le fichier, puis relance optimizer + migrate.")

        corrections = [
            {"num":"1","titre":"Borne journaliere trop permissive (CRITIQUE)",
             "couleur":"#e8543a","impact":"Reduit max/agri de 870t a ~200t pour FEDI/MAKKI",
             "avant":"day_planned * SCALE * 1.5",
             "apres":"day_planned * SCALE * 1.1",
             "ctx":"Dans la boucle creation variables OR-Tools (_ub_day):"},
            {"num":"2","titre":"Arrondi 10t -> 5t (micro-livraisons)",
             "couleur":"#f5a623","impact":"Autorise 10t, 15t, 20t pour COMOCAP/TUCAL",
             "avant":"int(round(round(tons_brut, 1) / 10)) * 10",
             "apres":"int(round(round(tons_brut, 1) / 5)) * 5",
             "ctx":"Dans la fonction de calcul du tonnage journalier:"},
            {"num":"3","titre":"FACTORY_OVERFLOW_WEIGHT insuffisant",
             "couleur":"#8b5cf6","impact":"ELFALLEH 60t/j max (au lieu de 130t OR-Tools)",
             "avant":"FACTORY_OVERFLOW_WEIGHT = 500",
             "apres":"FACTORY_OVERFLOW_WEIGHT = 2000",
             "ctx":"Dans les constantes globales (debut du fichier):"},
            {"num":"4","titre":"Cap max par agriculteur selon tonnage (NOUVEAU)",
             "couleur":"#00e5a0","impact":"Petits agris plafonnes 30t/j, moyens 60t/j, gros 200t/j",
             "avant":"_ub_day = max(int(day_planned * SCALE * 1.5), int(_get_min_tons(farmer) * SCALE))",
             "apres":"""# Cap max selon tonnage total
if farmer.tonnage < 200:
    _max_daily = min(30, farmer.tonnage)
elif farmer.tonnage < 500:
    _max_daily = min(60, farmer.tonnage / 8)
else:
    _max_daily = min(200, farmer.tonnage / 5)
_cap_agri = int(_max_daily * SCALE)
_ub_day = min(
    max(int(day_planned * SCALE * 1.1), int(_get_min_tons(farmer) * SCALE)),
    _cap_agri
)""",
             "ctx":"Dans la boucle creation variables OR-Tools:"},
        ]

        for corr in corrections:
            with st.expander(f"Correction {corr['num']} — {corr['titre']}",
                             expanded=(corr["num"] in ("1","3"))):
                st.markdown(f"""
                <div style='color:{corr["couleur"]};font-size:12px;font-weight:600;
                margin-bottom:8px'>Impact : {corr['impact']}</div>
                <div style='font-size:12px;color:#8b949e;margin-bottom:6px'>{corr['ctx']}</div>
                """, unsafe_allow_html=True)
                c_l, c_r = st.columns(2)
                with c_l:
                    st.markdown("**AVANT**")
                    st.code(corr["avant"], language="python")
                with c_r:
                    st.markdown("**APRES**")
                    st.code(corr["apres"], language="python")

        st.markdown("---")
        st.markdown("### Script PowerShell (corrections 1-3 automatiques)")
        ps_script = """\
@"
f=open('optimizer_v2.py','r',encoding='utf-8')
c=f.read()
f.close()
c=c.replace('day_planned * SCALE * 1.5','day_planned * SCALE * 1.1')
c=c.replace('FACTORY_OVERFLOW_WEIGHT = 500','FACTORY_OVERFLOW_WEIGHT = 2000')
c=c.replace('int(round(round(tons_brut, 1) / 10)) * 10','int(round(round(tons_brut, 1) / 5)) * 5')
f=open('optimizer_v2.py','w',encoding='utf-8')
f.write(c)
f.close()
print('OK - 3 corrections appliquees')
"@ | Out-File -Encoding utf8 fix_optimizer.py
python fix_optimizer.py"""
        st.code(ps_script, language="powershell")

        st.markdown("**Puis relancer :**")
        st.code("""\
cd C:\\Users\\amset\\Desktop\\Tomate_scheduler
python optimizer_v2.py
python migrate.py
streamlit run dashboard_phase10.py""", language="bash")

        st.markdown("### Resultats attendus apres corrections")
        result_rows = [
            {"Indicateur":"Max/jour FEDI","Avant":"870t","Apres":"~200t","Cible":"200t"},
            {"Indicateur":"Max/jour MAKKI","Avant":"130t","Apres":"~80t","Cible":"80t"},
            {"Indicateur":"Max ELFALLEH","Avant":"130t","Apres":"~80t","Cible":"80t"},
            {"Indicateur":"Max ABIDA","Avant":"320t","Apres":"~200t","Cible":"200t"},
            {"Indicateur":"Livraisons 10-20t","Avant":"0%","Apres":"~30%","Cible":"30%"},
            {"Indicateur":"Ecart total","Avant":"+0.2%","Apres":"+-0.5%","Cible":"<2%"},
        ]
        st.dataframe(pd.DataFrame(result_rows), use_container_width=True, hide_index=True)