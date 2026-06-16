# ============================================================
# ONGLET COMPARAISON v4 — Plans rectifiés vs OR-Tools
# SOURCE PRINCIPALE = Plans Rectifiés (reference_interne)
# OR-Tools = comparaison en pointillé uniquement
# ELFALLEH = 5190t / ABIDA = 8010t (données réelles)
# ============================================================
import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
import io
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import ColorScaleRule

SEASON  = pd.date_range("2026-06-20", "2026-08-25", freq="D")
PIC_S   = pd.Timestamp("2026-07-01").date()
PIC_E   = pd.Timestamp("2026-07-15").date()
DAYS_FR = ["Lun","Mar","Mer","Jeu","Ven","Sam","Dim"]

COMM_COLORS = {
    "FEDI":"#3b82f6","MAKKI BEN SALAH":"#00e5a0","KHALIL":"#f5a623",
    "ACHREF AJLANI":"#8b5cf6","JILANI OBAY":"#e8543a",
}
COMM_HEX = {
    "FEDI":"1F5FA6","MAKKI BEN SALAH":"1A6B3C","KHALIL":"B45309",
    "ACHREF AJLANI":"5B21B6","JILANI OBAY":"C0392B",
}
USINE_HEX = {
    "SICAM":"1F3864","TUCAL":"4A235A","COMOCAP":"0B4F6C","ELFALLEH":"196F3D","ABIDA":"922B21",
}

# ── SOURCE PRINCIPALE : Profils journaliers rectifiés (reference_interne 13/06/2026) ──
# Ces données proviennent directement des fichiers uploadés par les commerciaux
# RECTIFIÉ = vérité terrain | OR-Tools = calcul algorithme (comparaison)

RECTIF_COMM_DAILY = {
    "FEDI": [0,0,0,0,0,40,40,80,180,230,220,340,420,470,540,580,590,700,720,920,1045,1010,1050,1130,1200,1170,1216,1150,1070,1120,1135,980,1010,970,970,950,890,800,900,820,770,700,650,610,500,410,370,340,300,270,200,160,160,150,150,110,80,60,60,30,0,0,0,0,0,0,0],
    "MAKKI BEN SALAH": [0,0,35,35,55,160,160,225,235,260,285,555,575,585,670,840,895,1030,1115,1095,1015,1085,1215,1315,1180,1230,1150,1185,970,865,695,655,565,425,285,280,280,230,200,165,160,115,100,80,35,20,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0],
    "KHALIL": [140,170,200,200,220,220,190,240,435,545,510,650,745,865,935,1020,995,1060,940,870,770,765,655,570,560,505,390,330,300,175,160,180,180,125,125,125,110,70,70,20,20,20,20,20,20,20,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0],
    "ACHREF AJLANI": [0,0,0,0,0,0,30,30,60,120,180,360,450,450,390,390,360,360,330,270,480,510,570,450,480,570,430,480,390,390,420,340,210,150,126,300,390,390,300,300,270,270,300,330,300,210,240,240,210,150,150,120,90,90,150,270,330,450,390,330,210,240,210,120,120,120,120],
    "JILANI OBAY": [0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,60,60,60,130,150,150,260,310,340,310,420,280,400,410,410,370,380,450,430,330,310,310,270,220,180,0,0,0,0,0,0,0,0,0,0,0,0,0],
}

# Profils usines RÉELS (sources: PDFs et Excel reception_prevue)
RECTIF_USINE_DAILY = {
    # Source: Re_ception_Pre_v_TF26_SICAM.pdf — total 44871t
    "SICAM":[0,0,140,170,220,220,260,320,320,415,515,625,560,910,1090,1170,1285,
             1425,1410,1540,1490,1440,1440,1455,1650,1600,1595,1550,1385,1345,1125,
             1000,935,835,840,575,456,660,750,730,660,760,535,615,580,610,520,470,
             530,500,380,320,290,240,200,140,170,140,210,180,150,150,150,120,120,120,120,120],
    # Source: Re_ception_Pre_v__TF_Tucal26.pdf — total 20525t
    "TUCAL":[0,0,0,0,15,15,15,85,85,105,230,235,285,465,560,590,590,625,610,700,
             670,690,645,660,620,655,585,665,645,595,490,500,405,405,395,345,365,475,
             490,410,430,380,380,350,350,330,315,310,340,350,300,270,200,120,90,90,
             0,0,0,0,0,0,0,0,0,0,0],
    # Source: Re_ception_Pre_v_TF_26_Comocap.pdf — total 20391t
    "COMOCAP":[0,0,0,0,0,15,15,15,65,145,195,305,305,345,375,425,435,480,515,615,
               760,760,755,785,775,785,826,875,805,760,755,625,610,650,595,540,500,
               490,490,455,455,390,340,280,200,150,100,100,100,80,80,80,80,70,40,
               0,0,0,0,0,0,0,0,0,0,0],
    # Source: Re_ception_Pre_v_TF_Fallah_26.xlsx — total 5190t / 38 jours / max 240t/j
    "ELFALLEH":[0,0,0,0,0,0,0,0,40,70,70,105,105,115,125,165,215,215,230,230,240,
                225,235,215,175,155,155,140,130,120,130,155,140,140,130,110,130,130,
                130,110,110,80,70,70,60,20,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0],
    # Source: Re_ception_Pre_v_TF_ABIDA26.xlsx — total 8010t / max 320t/j
    "ABIDA":[0,0,0,0,0,0,0,0,30,80,80,120,140,140,140,170,170,200,200,200,240,270,
             250,230,290,320,190,180,190,220,220,210,110,110,110,110,110,70,70,20,20,
             80,140,170,170,110,90,60,60,60,60,60,60,60,120,210,270,300,270,210,60,
             90,90,0,0,0,0],
}

# Vérification totaux
def _daily_to_dict(vals):
    """Convertit liste de valeurs saisonnières en dict {str_date: valeur}"""
    return {str((pd.Timestamp("2026-06-20")+pd.Timedelta(days=i)).date()): float(v)
            for i,v in enumerate(vals) if v > 0}

RECTIF_COMM_DICT = {k: _daily_to_dict(v) for k, v in RECTIF_COMM_DAILY.items()}
RECTIF_USINE_DICT = {k: _daily_to_dict(v) for k, v in RECTIF_USINE_DAILY.items()}

# Totaux réels (source: plans rectifiés)
RECTIF_STATS_COMM = {
    "FEDI":           {"total": int(sum(RECTIF_COMM_DAILY["FEDI"])),           "max_j": max(RECTIF_COMM_DAILY["FEDI"]),           "n_j": sum(1 for v in RECTIF_COMM_DAILY["FEDI"] if v>0),           "cap": 1300},
    "MAKKI BEN SALAH":{"total": int(sum(RECTIF_COMM_DAILY["MAKKI BEN SALAH"])),  "max_j": max(RECTIF_COMM_DAILY["MAKKI BEN SALAH"]),  "n_j": sum(1 for v in RECTIF_COMM_DAILY["MAKKI BEN SALAH"] if v>0),  "cap": 1200},
    "KHALIL":         {"total": int(sum(RECTIF_COMM_DAILY["KHALIL"])),            "max_j": max(RECTIF_COMM_DAILY["KHALIL"]),            "n_j": sum(1 for v in RECTIF_COMM_DAILY["KHALIL"] if v>0),            "cap": 1100},
    "ACHREF AJLANI":  {"total": int(sum(RECTIF_COMM_DAILY["ACHREF AJLANI"])),     "max_j": max(RECTIF_COMM_DAILY["ACHREF AJLANI"]),     "n_j": sum(1 for v in RECTIF_COMM_DAILY["ACHREF AJLANI"] if v>0),     "cap": 700},
    "JILANI OBAY":    {"total": int(sum(RECTIF_COMM_DAILY["JILANI OBAY"])),       "max_j": max(RECTIF_COMM_DAILY["JILANI OBAY"]),       "n_j": sum(1 for v in RECTIF_COMM_DAILY["JILANI OBAY"] if v>0),       "cap": 150},
}
RECTIF_STATS_USINE = {
    "SICAM":   {"total":44871,"max_j":1650,"n_j":67,"cap_officiel":1300},
    "TUCAL":   {"total":20525,"max_j":700, "n_j":52,"cap_officiel":700},
    "COMOCAP": {"total":20391,"max_j":875, "n_j":50,"cap_officiel":700},
    "ELFALLEH":{"total":5190, "max_j":240, "n_j":38,"cap_officiel":150},
    "ABIDA":   {"total":8010, "max_j":320, "n_j":55,"cap_officiel":200},
}

USINE_CAPS      = {"SICAM":1300,"TUCAL":700,"COMOCAP":700,"ELFALLEH":150,"ABIDA":200}
USINE_CAPS_PLAN = {"SICAM":1300,"TUCAL":700,"COMOCAP":700,"ELFALLEH":240,"ABIDA":320}

# OR-Tools = profils calculés (comparaison)
ORT_COMM_BASE = {
    "FEDI":          {str((pd.Timestamp("2026-06-28")+pd.Timedelta(days=i)).date()):v
                     for i,v in enumerate([80,120,180,250,350,420,540,650,780,850,950,1050,
                     1100,1150,1175,1100,1050,900,750,600,500,450,400,350,300,250,200,180,
                     160,150,140,130,120,100,80,60,40,20,10]) if v>0},
    "MAKKI BEN SALAH":{str((pd.Timestamp("2026-06-22")+pd.Timedelta(days=i)).date()):v
                     for i,v in enumerate([50,100,150,200,280,380,480,580,650,750,850,950,
                     1050,1095,1000,900,780,680,580,480,380,300,250,200,160,130,100,80,60,
                     40,30,20,10]) if v>0},
    "KHALIL":        {str((pd.Timestamp("2026-06-20")+pd.Timedelta(days=i)).date()):v
                     for i,v in enumerate([200,280,350,430,520,630,720,820,920,1000,1020,980,
                     900,800,700,600,500,420,350,280,220,180,150,120,90,70,50,30,20,10]) if v>0},
    "ACHREF AJLANI": {str((pd.Timestamp("2026-06-22")+pd.Timedelta(days=i)).date()):v
                     for i,v in enumerate([30,60,120,180,240,300,360,420,480,540,600,650,700,
                     750,790,810,810,790,750,700,640,570,490,400,320,250,190,140,100,70,50,
                     30,20,10,5]) if v>0},
    "JILANI OBAY":   {str((pd.Timestamp("2026-07-23")+pd.Timedelta(days=i)).date()):v
                     for i,v in enumerate([80,130,180,240,310,380,440,440,410,370,320,270,220,
                     180,140,100,70,50,30,15]) if v>0},
}
ORT_USINE_BASE = {
    "SICAM":   {str((pd.Timestamp("2026-06-20")+pd.Timedelta(days=i)).date()):v
               for i,v in enumerate([100,120,180,250,250,420,440,500,560,653,743,898,895,1000,
               1105,1128,1213,1285,1340,1490,1530,1540,1510,1495,1440,1350,1280,1210,1140,
               1015,970,960,965,930,930,1155,1180,1220,1270,1280,1180,1090,1045,1040,940,860,
               850,835,795,578,600,540,520,380,260,210,150,90,30]) if v>0},
    "TUCAL":   {str((pd.Timestamp("2026-06-20")+pd.Timedelta(days=i)).date()):v
               for i,v in enumerate([20,60,80,80,110,195,195,245,315,320,325,330,365,390,420,
               415,395,380,350,370,340,300,290,320,300,230,210,205,185,255,240,240,140,160,235,
               315,345,380,410,410,385,385,375,370,330,370,350,340,270,260,270,255,250,230,190,
               190,90,100,90]) if v>0},
    "COMOCAP": {str((pd.Timestamp("2026-06-20")+pd.Timedelta(days=i)).date()):v
               for i,v in enumerate([0,0,0,0,0,0,0,0,20,55,55,60,135,145,160,165,165,165,170,
               400,520,500,505,585,600,515,558,545,580,542,520,421,380,463,343,293,275,245,290,
               306,293,250,247,197,160,118,45,40,15,15,8]) if v>0},
    "ELFALLEH": {},  # OR-Tools n'a pas planifié ELFALLEH
    "ABIDA":    {},  # OR-Tools n'a pas planifié ABIDA
}

# ── Helpers données ───────────────────────────────────────────
def _get_rectif_comm(comm, uploaded_sess, sb_comm):
    """Retourne le profil rectifié pour un commercial (priorité: upload > Supabase > reference_interne)"""
    if comm in uploaded_sess:
        d = {str(k):v for k,v in uploaded_sess[comm].items()}
        return d, "fichier upload"
    if comm in sb_comm:
        d = {str(k):v for k,v in sb_comm[comm].items()}
        return d, "Supabase"
    # Fallback: reference_interne (données hardcodées depuis fichiers 13/06/2026)
    return RECTIF_COMM_DICT[comm], "reference interne (13/06/2026)"

def _get_rectif_usine(usine, sb_usine):
    """Retourne le profil rectifié pour une usine"""
    if usine in sb_usine:
        return {str(k):v for k,v in sb_usine[usine].items()}, "Supabase"
    return RECTIF_USINE_DICT.get(usine, {}), "reference interne"

def _get_ort_profile(entity, planning_df, entity_type="commercial"):
    """Extrait le profil OR-Tools depuis Supabase ou utilise le profil de base"""
    if planning_df is not None and not planning_df.empty:
        col = "Commercial" if entity_type=="commercial" else "Usine"
        if col in planning_df.columns:
            sub = planning_df[planning_df[col]==entity].copy()
            if not sub.empty and "Date" in sub.columns:
                sub["Date"] = pd.to_datetime(sub["Date"],errors="coerce")
                daily = sub.groupby("Date")["Tonnes/Jour"].sum().reset_index()
                r = {str(r["Date"].date()):float(r["Tonnes/Jour"]) for _,r in daily.iterrows()}
                if r: return r
    if entity_type=="commercial":
        return ORT_COMM_BASE.get(entity, {})
    if entity in ORT_USINE_BASE and ORT_USINE_BASE[entity]:
        return ORT_USINE_BASE[entity]
    return RECTIF_USINE_DICT.get(entity, {})  # Fallback sur rectifié pour OR-T

def _ns(d):
    """Normalise les clés en str"""
    return {str(k):v for k,v in d.items()} if d else {}

# ── Helpers openpyxl ─────────────────────────────────────────
def _hf(h): return PatternFill("solid", start_color=h, end_color=h)
def _ft(bold=False, color="1F1F1F", size=10, white=False):
    return Font(bold=bold, name="Calibri", size=size, color="FFFFFF" if white else color)
_CTR=Alignment(horizontal="center",vertical="center")
_LFT=Alignment(horizontal="left",vertical="center")
_THIN=Side(style="thin",color="CCCCCC")
_BORD=Border(left=_THIN,right=_THIN,top=_THIN,bottom=_THIN)
F_RECT=_hf("E2EFDA");F_ORT=_hf("DEEBF7");F_POS=_hf("C6EFCE");F_NEG=_hf("FFC7CE")
F_NEU=_hf("F2F2F2");F_PIC=_hf("FFF2CC");F_ALT1=_hf("FFFFFF");F_ALT2=_hf("F8F9FA")

# ═══ SUPABASE PERSISTENCE ════════════════════════════════════
def _save_rectifie(sb, entity_name, daily_dict, entity_type="commercial"):
    if sb is None: return False
    try:
        sb.table("plan_rectifie").delete().eq("entity_type",entity_type).eq("entity_name",entity_name).execute()
        rows=[{"entity_type":entity_type,"entity_name":entity_name,"date":str(d),"tonnes":float(t)}
              for d,t in daily_dict.items() if float(t)>0]
        if rows:
            for i in range(0,len(rows),500):
                sb.table("plan_rectifie").insert(rows[i:i+500]).execute()
        return True
    except Exception as e:
        st.warning(f"Supabase save: {e}"); return False

def _load_all_rectifie(sb):
    if sb is None: return {},{}
    try:
        data=sb.table("plan_rectifie").select("entity_type,entity_name,date,tonnes").execute().data
        if not data: return {},{}
        cp={};up={}
        for row in data:
            et=row.get("entity_type","commercial"); en=row["entity_name"]
            d=str(row["date"]); t=float(row["tonnes"])
            if et=="commercial": cp.setdefault(en,{})[d]=t
            else: up.setdefault(en,{})[d]=t
        return cp,up
    except Exception: return {},{}

# ── Parser upload ─────────────────────────────────────────────
def _detect_comm_from_filename(filename):
    fn=filename.upper()
    for key,comm in {"FEDI":"FEDI","MEKKI":"MAKKI BEN SALAH","MAKKI":"MAKKI BEN SALAH",
                     "KHALIL":"KHALIL","ACHRAF":"ACHREF AJLANI","ACHREF":"ACHREF AJLANI",
                     "JILENI":"JILANI OBAY","JILANI":"JILANI OBAY"}.items():
        if key in fn: return comm
    return None

def _parse_rectification(uploaded_file):
    try: df=pd.read_excel(uploaded_file,header=0)
    except Exception as e: return None,None,str(e)
    cols=df.columns.tolist()
    date_cols=[(c,pd.Timestamp(c)) for c in cols if isinstance(c,pd.Timestamp) or (hasattr(c,'year') and getattr(c,'year',0)==2026)]
    if not date_cols:
        # Essai header=1
        try:
            uploaded_file.seek(0); df=pd.read_excel(uploaded_file,header=1)
            cols=df.columns.tolist()
            for c in cols[5:]:
                cs=str(c).strip()
                if '/' in cs and len(cs)<=5:
                    try: p=cs.split('/'); date_cols.append((c,pd.Timestamp(f"2026-{int(p[1]):02d}-{int(p[0]):02d}")))
                    except: pass
                else:
                    try:
                        d=pd.to_datetime(str(c).split(' ')[0],errors='coerce')
                        if pd.notna(d) and d.year==2026: date_cols.append((c,d))
                    except: pass
        except: pass
    col_agri=next((c for c in cols if 'agriculteur' in str(c).lower()),None)
    col_ton =next((c for c in cols if 'tonnage' in str(c).lower()),None)
    if not date_cols or not col_agri: return None,None,"Format non reconnu"
    rows=[];comm_detected=None
    col_comm=next((c for c in cols if 'responsable' in str(c).lower() or 'commercial' in str(c).lower()),None)
    for _,row in df.iterrows():
        agri=str(row.get(col_agri,'') or '').strip()
        if not agri or agri.upper() in ('NAN','','AGRICULTEUR','TOTAL'): continue
        if 'total' in agri.lower() or 'sous' in agri.lower(): continue
        if col_comm:
            comm=str(row[col_comm] or '').strip()
            if comm and comm.upper() not in ('NAN',''):
                for known in RECTIF_STATS_COMM:
                    if known.upper() in comm.upper() or comm.upper() in known.upper():
                        comm_detected=known; break
        ton=float(pd.to_numeric(row.get(col_ton,0),errors='coerce') or 0)
        for oc,d in date_cols:
            v=pd.to_numeric(row[oc],errors='coerce')
            if pd.notna(v) and v>0:
                rows.append({"agriculteur":agri,"date":d,"tonnes":float(v),"ton_total":ton})
    if not rows: return comm_detected,None,"Aucune donnée trouvée"
    df_rows=pd.DataFrame(rows)
    daily=df_rows.groupby("date")["tonnes"].sum().reset_index().sort_values("date")
    daily_dict={str(r["date"].date()):float(r["tonnes"]) for _,r in daily.iterrows()}
    return comm_detected,daily_dict,df_rows

# ── KPIs et chart ─────────────────────────────────────────────
def _kpi_row(name, man_dict, ort_dict, cap, label_man="Plan Rectifié", label_ort="OR-Tools"):
    ms=_ns(man_dict); os_=_ns(ort_dict)
    mt=sum(ms.values()) if ms else 0; mm=max(ms.values()) if ms else 0
    ot=sum(os_.values()) if os_ else 0; om=max(os_.values()) if os_ else 0
    k1,k2,k3,k4=st.columns(4)
    k1.metric(f"Total {label_man}",f"{int(mt)}t",delta=f"{int(mt-ot):+d}t vs OR-T",delta_color="off")
    k2.metric(f"Max/j {label_man}",f"{int(mm)}t",
              delta=f"OR-T max: {int(om)}t",
              delta_color="inverse" if mm>cap else "off")
    k3.metric("Jours actifs",f"{sum(1 for v in ms.values() if v>0)}j" if ms else "0j")
    k4.metric("Cap PIC",f"{cap}t/j",
              delta="OK" if mm<=cap else "DEPASSE",
              delta_color="normal" if mm<=cap else "inverse")

def _build_chart(name, man_dict, ort_dict, color, cap):
    ms=_ns(man_dict); os_=_ns(ort_dict)
    all_d=sorted(set(list(ms.keys())+list(os_.keys())))
    dates=[pd.Timestamp(d) for d in all_d]
    mv=[ms.get(d,0) for d in all_d]; ov=[os_.get(d,0) for d in all_d]
    fig=go.Figure()
    fig.add_trace(go.Scatter(x=dates,y=mv,name="Plan Rectifié (SOURCE)",
        line=dict(color=color,width=2.5),fill="tozeroy",
        fillcolor="rgba(59,130,246,0.1)",mode="lines"))
    if any(v>0 for v in ov):
        fig.add_trace(go.Scatter(x=dates,y=ov,name="OR-Tools (comparaison)",
            line=dict(color="#555555",width=1.5,dash="dot"),mode="lines"))
    fig.add_hline(y=cap,line_dash="dash",line_color="#e8543a",line_width=1,
                  annotation_text=f"Cap: {cap}t/j",annotation_position="top right",
                  annotation_font_color="#e8543a")
    fig.add_vrect(x0=pd.Timestamp("2026-07-01"),x1=pd.Timestamp("2026-07-15"),
                  fillcolor="rgba(245,166,35,0.07)",line_width=0,
                  annotation_text="PIC",annotation_position="top left",annotation_font_color="#f5a623")
    fig.update_layout(
        title=f"{name} — Plan Rectifié (SOURCE) vs OR-Tools (comparaison)",
        template="plotly_dark",plot_bgcolor="#0d1117",paper_bgcolor="#161b22",
        height=360,hovermode="closest",
        legend=dict(orientation="h",yanchor="bottom",y=1.02),
        xaxis=dict(title="Date",gridcolor="#21262d"),
        yaxis=dict(title="Tonnes/jour",gridcolor="#21262d"))
    return fig

# ═══ EXCEL EXPORT ════════════════════════════════════════════
def _build_entity_sheet(ws,name,man_dict,ort_dict,cap,hex_color):
    def td(d):
        out={}
        for k,v in d.items():
            try: out[pd.Timestamp(k).date()]=float(v)
            except: pass
        return out
    man_d=td(_ns(man_dict)); ort_d=td(_ns(ort_dict))
    HDR=_hf(hex_color); N=8
    ws.merge_cells(start_row=1,start_column=1,end_row=1,end_column=N)
    ws.cell(1,1).value=f"Comparaison  —  {name}  —  Saison 2026  |  SOURCE = Plan Rectifié"
    ws.cell(1,1).font=_ft(bold=True,size=12,white=True); ws.cell(1,1).fill=HDR; ws.cell(1,1).alignment=_CTR
    ws.row_dimensions[1].height=30
    ws.merge_cells(start_row=2,start_column=1,end_row=2,end_column=N)
    ws.cell(2,1).value="Vert=Rectifié>OR-T | Rouge=Rectifié<OR-T | Jaune=PIC | Cap officiel="+str(cap)+"t/j"
    ws.cell(2,1).font=_ft(size=9,color="595959"); ws.cell(2,1).fill=_hf("F0F4F8"); ws.cell(2,1).alignment=_LFT
    for ci,h in enumerate(["Date","Jour","Rectifié (t)","OR-Tools (t)","Ecart (t)","Ecart %","Statut",f"Cap {cap}t/j"],1):
        ws.cell(3,ci).value=h; ws.cell(3,ci).font=_ft(bold=True,size=9,white=True)
        ws.cell(3,ci).fill=HDR; ws.cell(3,ci).alignment=_CTR; ws.cell(3,ci).border=_BORD
    data_rows=[]
    for d in SEASON:
        dk=d.date(); mv=man_d.get(dk,0); ov=ort_d.get(dk,0)
        if mv==0 and ov==0: continue
        data_rows.append((d,dk,mv,ov))
    for ri,(d,dk,mv,ov) in enumerate(data_rows):
        r=ri+4; is_pic=(PIC_S<=dk<=PIC_E); is_alt=ri%2==0
        base=F_PIC if is_pic else (F_ALT1 if is_alt else F_ALT2)
        ecart=mv-ov; pct=round(ecart/ov*100,1) if ov>0 else None
        ws.cell(r,1).value=d.strftime("%d/%m/%Y"); ws.cell(r,1).fill=base; ws.cell(r,1).border=_BORD; ws.cell(r,1).alignment=_CTR
        ws.cell(r,1).font=_ft(bold=is_pic,color="7D4F00" if is_pic else "1F1F1F")
        ws.cell(r,2).value=DAYS_FR[d.weekday()]; ws.cell(r,2).fill=base; ws.cell(r,2).border=_BORD; ws.cell(r,2).alignment=_CTR; ws.cell(r,2).font=_ft(size=9,color="595959")
        ws.cell(r,3).value=int(mv) if mv>0 else ""; ws.cell(r,3).fill=F_RECT if mv>0 else base; ws.cell(r,3).border=_BORD; ws.cell(r,3).alignment=_CTR; ws.cell(r,3).number_format="#,##0"
        ws.cell(r,4).value=int(ov) if ov>0 else ""; ws.cell(r,4).fill=F_ORT if ov>0 else base; ws.cell(r,4).border=_BORD; ws.cell(r,4).alignment=_CTR; ws.cell(r,4).number_format="#,##0"
        if mv>0 or ov>0:
            ws.cell(r,5).value=int(ecart); ws.cell(r,5).fill=F_POS if ecart>50 else (F_NEG if ecart<-50 else F_NEU); ws.cell(r,5).number_format="+#,##0;-#,##0;0"
        ws.cell(r,5).border=_BORD; ws.cell(r,5).alignment=_CTR
        if pct is not None:
            ws.cell(r,6).value=pct/100; ws.cell(r,6).fill=F_POS if pct>5 else (F_NEG if pct<-5 else F_NEU); ws.cell(r,6).number_format="+0.0%;-0.0%;0%"
        else: ws.cell(r,6).value="—" if mv>0 else ""; ws.cell(r,6).fill=base
        ws.cell(r,6).border=_BORD; ws.cell(r,6).alignment=_CTR
        if mv>0 and ov==0:   txt,fc,tc="Rect only","E8F5E9","2E7D32"
        elif mv==0 and ov>0: txt,fc,tc="OR-T only","FFF3E0","E65100"
        elif abs(ecart)<=50: txt,fc,tc="OK","E8F5E9","2E7D32"
        elif ecart>0:        txt,fc,tc="Rect+","E3F2FD","0D47A1"
        else:                txt,fc,tc="OR-T+","FCE4D6","993200"
        ws.cell(r,7).value=txt; ws.cell(r,7).fill=_hf(fc); ws.cell(r,7).font=_ft(size=9,color=tc); ws.cell(r,7).border=_BORD; ws.cell(r,7).alignment=_CTR
        if is_pic:
            peak=max(mv,ov)
            ws.cell(r,8).value=f"{'DEPASSE' if peak>cap else 'OK'} {int(peak)}t"
            ws.cell(r,8).fill=F_NEG if peak>cap else _hf("C6EFCE")
        else: ws.cell(r,8).value="—"; ws.cell(r,8).fill=base
        ws.cell(r,8).border=_BORD; ws.cell(r,8).alignment=_CTR
        ws.row_dimensions[r].height=17
    tr=len(data_rows)+4
    man_t=sum(r[2] for r in data_rows); ort_t=sum(r[3] for r in data_rows)
    ec_t=man_t-ort_t; pct_t=round(ec_t/ort_t*100,1) if ort_t>0 else 0
    ws.merge_cells(start_row=tr,start_column=1,end_row=tr,end_column=2)
    for ci in range(1,N+1):
        ws.cell(tr,ci).fill=HDR; ws.cell(tr,ci).border=_BORD; ws.cell(tr,ci).alignment=_CTR
        ws.cell(tr,ci).font=_ft(bold=True,size=10,white=True)
    ws.cell(tr,1).value="TOTAL SAISON (RECTIFIÉ)"
    ws.cell(tr,3).value=int(man_t); ws.cell(tr,3).number_format="#,##0"
    ws.cell(tr,4).value=int(ort_t); ws.cell(tr,4).number_format="#,##0"
    ws.cell(tr,5).value=int(ec_t); ws.cell(tr,5).number_format="+#,##0;-#,##0;0"
    ws.cell(tr,6).value=pct_t/100; ws.cell(tr,6).number_format="+0.0%;-0.0%;0%"
    ws.cell(tr,7).value=f"R:{sum(1 for r in data_rows if r[2]>0)}j|OT:{sum(1 for r in data_rows if r[3]>0)}j"
    ws.cell(tr,8).value=f"Max R:{int(max((r[2] for r in data_rows),default=0))}t OT:{int(max((r[3] for r in data_rows),default=0))}t"
    ws.row_dimensions[tr].height=24
    for ci,w in enumerate([12,6,13,13,12,10,12,16],1):
        ws.column_dimensions[get_column_letter(ci)].width=w
    ws.freeze_panes="A4"

def _build_synthese_sheet(ws,man_comm,ort_comm,man_usine,ort_usine):
    HDR=_hf("1F3864"); N=11
    ws.merge_cells(start_row=1,start_column=1,end_row=1,end_column=N)
    ws.cell(1,1).value="Synthèse — Plan Rectifié (SOURCE) vs OR-Tools — Saison 2026"
    ws.cell(1,1).font=_ft(bold=True,size=14,white=True); ws.cell(1,1).fill=HDR; ws.cell(1,1).alignment=_CTR
    ws.row_dimensions[1].height=30
    for ci,h in enumerate(["Entité","Type","Total Rectifié(t)","Total OR-T(t)","Ecart(t)",
                            "Ecart%","Max Rect(t)","Max OR-T(t)","Jours Rect","Jours OR-T","Cap PIC"],1):
        ws.cell(3,ci).value=h; ws.cell(3,ci).font=_ft(bold=True,size=9,white=True)
        ws.cell(3,ci).fill=HDR; ws.cell(3,ci).alignment=_CTR; ws.cell(3,ci).border=_BORD
    row=4
    comm_order=["FEDI","MAKKI BEN SALAH","KHALIL","ACHREF AJLANI","JILANI OBAY"]
    usine_order=["SICAM","TUCAL","COMOCAP","ELFALLEH","ABIDA"]
    ws.merge_cells(start_row=row,start_column=1,end_row=row,end_column=N)
    ws.cell(row,1).value="COMMERCIAUX — Source: Plans Rectifiés (reference_interne 13/06/2026)"
    ws.cell(row,1).font=_ft(bold=True,size=10,white=True); ws.cell(row,1).fill=_hf("2F4F7F"); ws.cell(row,1).alignment=_LFT
    ws.row_dimensions[row].height=16; row+=1
    for ci,comm in enumerate(comm_order):
        man=_ns(man_comm.get(comm,{})); ort=_ns(ort_comm.get(comm,{}))
        alld=sorted(set(list(man.keys())+list(ort.keys())))
        mt=sum(man.get(d,0) for d in alld); ot=sum(ort.get(d,0) for d in alld)
        ec=mt-ot; pct=round(ec/ot*100,1) if ot>0 else 0
        mm=max((man.get(d,0) for d in alld),default=0); om=max((ort.get(d,0) for d in alld),default=0)
        mj=sum(1 for d in alld if man.get(d,0)>0); oj=sum(1 for d in alld if ort.get(d,0)>0)
        hx=COMM_HEX.get(comm,"1F3864"); cap=RECTIF_STATS_COMM.get(comm,{}).get("cap",0)
        for j,v in enumerate([comm,"Commercial",int(mt),int(ot),int(ec),f"{pct:+.1f}%",int(mm),int(om),mj,oj,cap],1):
            ws.cell(row,j).value=v; ws.cell(row,j).border=_BORD; ws.cell(row,j).alignment=_LFT if j==1 else _CTR
            ws.cell(row,j).font=_ft(bold=(j==1),color="FFFFFF" if j==1 else "1F1F1F")
            ws.cell(row,j).fill=_hf(hx) if j==1 else (F_ALT1 if ci%2==0 else F_ALT2)
            if j==5: ws.cell(row,j).fill=F_POS if ec>=0 else F_NEG
        ws.row_dimensions[row].height=17; row+=1
    ws.merge_cells(start_row=row,start_column=1,end_row=row,end_column=N)
    ws.cell(row,1).value="USINES — Source: Réceptions prévues (PDFs/Excel 13/06/2026)"
    ws.cell(row,1).font=_ft(bold=True,size=10,white=True); ws.cell(row,1).fill=_hf("0B4F6C"); ws.cell(row,1).alignment=_LFT
    ws.row_dimensions[row].height=16; row+=1
    for ui,usine in enumerate(usine_order):
        man=_ns(man_usine.get(usine,{})); ort=_ns(ort_usine.get(usine,{}))
        alld=sorted(set(list(man.keys())+list(ort.keys())))
        mt=sum(man.get(d,0) for d in alld); ot=sum(ort.get(d,0) for d in alld)
        ec=mt-ot; pct=round(ec/ot*100,1) if ot>0 else 0
        mm=max((man.get(d,0) for d in alld),default=0); om=max((ort.get(d,0) for d in alld),default=0)
        mj=sum(1 for d in alld if man.get(d,0)>0); oj=sum(1 for d in alld if ort.get(d,0)>0)
        hx=USINE_HEX.get(usine,"1F3864"); cap=USINE_CAPS.get(usine,0)
        for j,v in enumerate([usine,"Usine",int(mt),int(ot),int(ec),f"{pct:+.1f}%",int(mm),int(om),mj,oj,cap],1):
            ws.cell(row,j).value=v; ws.cell(row,j).border=_BORD; ws.cell(row,j).alignment=_LFT if j==1 else _CTR
            ws.cell(row,j).font=_ft(bold=(j==1),color="FFFFFF" if j==1 else "1F1F1F")
            ws.cell(row,j).fill=_hf(hx) if j==1 else (F_ALT1 if ui%2==0 else F_ALT2)
            if j==5: ws.cell(row,j).fill=F_POS if ec>=0 else F_NEG
        ws.row_dimensions[row].height=17; row+=1

def _build_legende_sheet(ws):
    HDR=_hf("1F3864")
    ws.merge_cells("A1:D1"); ws.cell(1,1).value="Légende — Codes couleur"
    ws.cell(1,1).font=_ft(bold=True,size=13,white=True); ws.cell(1,1).fill=HDR; ws.cell(1,1).alignment=_CTR
    for ci,h in enumerate(["Couleur","Code Hex","Signification","Usage"],1):
        ws.cell(3,ci).value=h; ws.cell(3,ci).font=_ft(bold=True,size=10,white=True)
        ws.cell(3,ci).fill=HDR; ws.cell(3,ci).alignment=_CTR; ws.cell(3,ci).border=_BORD
    items=[("E2EFDA","Vert clair","Plan Rectifié — SOURCE PRINCIPALE","Colonne Rectifié"),
           ("DEEBF7","Bleu clair","OR-Tools — comparaison algorithme","Colonne OR-Tools"),
           ("C6EFCE","Vert moyen","Rectifié > OR-Tools de +50t","Ecart positif"),
           ("FFC7CE","Rose","Rectifié < OR-Tools de +50t","Ecart négatif"),
           ("FFF2CC","Jaune","Période PIC 1-15 juillet","Caps actifs")]
    for ri,(hex_c,nom,desc,usage) in enumerate(items,4):
        ws.cell(ri,1).fill=_hf(hex_c); ws.cell(ri,1).border=_BORD; ws.cell(ri,1).value=nom; ws.cell(ri,1).font=_ft(size=10)
        ws.cell(ri,2).value=f"#{hex_c}"; ws.cell(ri,2).fill=_hf(hex_c); ws.cell(ri,2).border=_BORD; ws.cell(ri,2).alignment=_CTR
        ws.cell(ri,3).value=desc; ws.cell(ri,3).border=_BORD
        ws.cell(ri,4).value=usage; ws.cell(ri,4).border=_BORD
    for ci,w in enumerate([18,12,50,25],1):
        ws.column_dimensions[get_column_letter(ci)].width=w

def _build_planning_export_sheet(ws, planning_df, rectif_comm, rectif_usine):
    """
    Onglet 'Planning Journalier Rectifié' avec structure:
    Date | Commercial | Agriculteur | Usine | Tonnes/Jour | Type Véhicule |
    Véhicules Requis | Disponibles | Manquants | Nb Voyages | Pic | 25/06/2026 | 26/06/2026 | ...
    Basé sur les plans RECTIFIÉS comme tonnage principal
    """
    HDR=_hf("1F3864")
    SEASON_DATES = list(pd.date_range("2026-06-20","2026-08-25",freq="D"))
    PIC_S_ = pd.Timestamp("2026-07-01").date()
    PIC_E_ = pd.Timestamp("2026-07-15").date()

    FIXED_HDRS = ["Date","Commercial","Agriculteur","Usine","Tonnes/Jour",
                  "Type Véhicule","Véhicules Requis","Disponibles","Manquants (à louer)","Nb Voyages","Pic de Récolte"]
    DATE_HDRS  = [d.strftime("%d/%m/%Y") for d in SEASON_DATES]
    ALL_HDRS   = FIXED_HDRS + DATE_HDRS
    N = len(ALL_HDRS)

    # Titre
    ws.merge_cells(start_row=1,start_column=1,end_row=1,end_column=min(N,30))
    ws.cell(1,1).value="Planning Journalier Rectifié — SOURCE = Plans Rectifiés (reference_interne)"
    ws.cell(1,1).font=_ft(bold=True,size=12,white=True); ws.cell(1,1).fill=HDR; ws.cell(1,1).alignment=_CTR
    ws.row_dimensions[1].height=28

    # En-têtes
    for ci,h in enumerate(ALL_HDRS,1):
        c=ws.cell(2,ci); c.value=h; c.border=_BORD; c.alignment=_CTR
        if ci<=len(FIXED_HDRS):
            c.fill=HDR; c.font=_ft(bold=True,size=9,white=True)
        else:
            d_obj=SEASON_DATES[ci-len(FIXED_HDRS)-1].date()
            is_pic_col=PIC_S_<=d_obj<=PIC_E_
            c.fill=_hf("7D6608") if is_pic_col else _hf("1F4E79")
            c.font=_ft(bold=True,size=8,white=True)
    ws.row_dimensions[2].height=36

    COMM_FILLS_HEX={
        "FEDI":"DEEBF7","MAKKI BEN SALAH":"E2EFDA","KHALIL":"FFF2CC",
        "ACHREF AJLANI":"EDEDED","JILANI OBAY":"FCE4D6",
    }
    F_GRN_L=_hf("E2EFDA"); F_PIC_L=_hf("FFF2CC")
    FLEET={"SICAM":{"PL":48,"PPL":6,"SEMI":13},"TUCAL":{"PL":17,"PPL":0,"SEMI":2},
           "COMOCAP":{"PL":6,"PPL":14,"SEMI":3},"ABIDA":{"PL":1,"PPL":0,"SEMI":2},"ELFALLEH":{"PL":0,"PPL":2,"SEMI":0}}

    row=3
    comm_order=["FEDI","MAKKI BEN SALAH","KHALIL","ACHREF AJLANI","JILANI OBAY"]

    for comm in comm_order:
        rect_daily=_ns(rectif_comm.get(comm, RECTIF_COMM_DICT.get(comm,{})))
        comm_fill=_hf(COMM_FILLS_HEX.get(comm,"F0F0F0"))
        comm_bg=_hf("1F3864") if comm=="FEDI" else _hf(COMM_HEX.get(comm,"1F3864"))

        # Si planning_df disponible: lignes par agriculteur
        if planning_df is not None and not planning_df.empty and "Commercial" in planning_df.columns:
            sub=planning_df[planning_df["Commercial"]==comm].copy()
            if not sub.empty and "Agriculteur" in sub.columns:
                grps=sub.groupby(["Agriculteur","Usine"] if "Usine" in sub.columns else ["Agriculteur"])
                for key,grp in grps:
                    agri=key[0] if isinstance(key,tuple) else key
                    usine=key[1] if isinstance(key,tuple) and len(key)>1 else ""
                    # Tonnage depuis OR-Tools (structure agriculteur)
                    ort_daily_agri={}
                    if "Date" in grp.columns and "Tonnes/Jour" in grp.columns:
                        grp2=grp.copy(); grp2["Date"]=pd.to_datetime(grp2["Date"],errors='coerce')
                        for _,r in grp2.iterrows():
                            if pd.notna(r["Date"]):
                                ort_daily_agri[str(r["Date"].date())]=float(r.get("Tonnes/Jour",0) or 0)
                    r0=grp.iloc[0]
                    tv=str(r0.get("Type Véhicule","") or "")
                    vr=str(r0.get("Véhicules Requis","") or "")
                    nv=str(r0.get("Nb Voyages","") or "")
                    total_ort=sum(ort_daily_agri.values())
                    is_pic_agri=any(PIC_S_<=pd.Timestamp(k).date()<=PIC_E_ for k,v in ort_daily_agri.items() if v>0)
                    fleet_usine=FLEET.get(usine.upper(),{})
                    dispo=fleet_usine.get(tv.upper(),0)
                    vr_n=int(pd.to_numeric(vr,errors='coerce') or 0)
                    manque=max(0,vr_n-dispo)
                    is_alt=(row%2==0)
                    base_fill=_hf("F5F8FF") if is_alt else _hf("FFFFFF")
                    fixed_vals=["",comm,agri,usine,round(total_ort,0) if total_ort>0 else "",
                                tv,vr,dispo if vr_n>0 else "",manque if vr_n>0 else "",nv,
                                "⚡ PIC" if is_pic_agri else ""]
                    for ci,val in enumerate(fixed_vals,1):
                        cell=ws.cell(row,ci); cell.value=val; cell.border=_BORD; cell.alignment=_CTR
                        if ci==2: cell.fill=comm_fill; cell.font=_ft(bold=True,size=9,color="1F1F1F")
                        elif ci==9 and isinstance(val,int) and val>0: cell.fill=_hf("FFC7CE"); cell.font=_ft(size=9,bold=True,color="9C0006")
                        else: cell.fill=base_fill; cell.font=_ft(size=9)
                    for di,d in enumerate(SEASON_DATES):
                        dk=str(d.date()); val=ort_daily_agri.get(dk,"")
                        if val==0: val=""
                        ci=len(FIXED_HDRS)+di+1
                        cell=ws.cell(row,ci); cell.value=int(val) if val!="" else ""; cell.border=_BORD; cell.alignment=_CTR
                        is_pic_d=PIC_S_<=d.date()<=PIC_E_
                        cell.fill=F_PIC_L if (is_pic_d and val!="") else (F_GRN_L if val!="" else base_fill)
                        cell.font=_ft(size=8)
                    ws.row_dimensions[row].height=15; row+=1
                continue  # passer au commercial suivant

        # Fallback: ligne agrégée depuis rectifié
        total_rect=sum(float(v) for v in rect_daily.values())
        is_alt=(row%2==0); base_fill=_hf("F5F8FF") if is_alt else _hf("FFFFFF")
        fixed_vals=["",comm,"— Agrégé commercial —","Toutes usines",round(total_rect,0),"","","","","",
                    "⚡ PIC" if any(PIC_S_<=pd.Timestamp(k).date()<=PIC_E_ for k,v in rect_daily.items() if float(v)>0) else ""]
        for ci,val in enumerate(fixed_vals,1):
            cell=ws.cell(row,ci); cell.value=val; cell.border=_BORD; cell.alignment=_CTR
            cell.fill=comm_fill if ci==2 else base_fill; cell.font=_ft(size=9,bold=(ci==2))
        for di,d in enumerate(SEASON_DATES):
            dk=str(d.date()); val=float(rect_daily.get(dk,0) or 0)
            ci=len(FIXED_HDRS)+di+1
            cell=ws.cell(row,ci); cell.value=int(val) if val>0 else ""; cell.border=_BORD; cell.alignment=_CTR
            is_pic_d=PIC_S_<=d.date()<=PIC_E_
            cell.fill=F_PIC_L if (is_pic_d and val>0) else (F_GRN_L if val>0 else base_fill)
            cell.font=_ft(size=8)
        ws.row_dimensions[row].height=15; row+=1

    # Largeurs
    for ci,w in enumerate([12,16,28,12,11,12,11,11,13,10,11],1):
        ws.column_dimensions[get_column_letter(ci)].width=w
    for di in range(len(SEASON_DATES)):
        ws.column_dimensions[get_column_letter(len(FIXED_HDRS)+di+1)].width=8
    ws.freeze_panes="L3"

def generate_comparison_excel(man_comm,ort_comm,man_usine,ort_usine,planning_df=None):
    wb=Workbook(); wb.remove(wb.active)
    comm_order=["FEDI","MAKKI BEN SALAH","KHALIL","ACHREF AJLANI","JILANI OBAY"]
    usine_order=["SICAM","TUCAL","COMOCAP","ELFALLEH","ABIDA"]

    # Onglet Synthèse
    ws=wb.create_sheet("Synthese Globale")
    _build_synthese_sheet(ws,man_comm,ort_comm,man_usine,ort_usine)

    # Onglet Planning Journalier Rectifié
    ws=wb.create_sheet("Planning Journalier Rectifie")
    _build_planning_export_sheet(ws,planning_df,man_comm,man_usine)

    # Onglets par commercial
    for comm in comm_order:
        ws=wb.create_sheet(f"C - {comm[:14]}")
        _build_entity_sheet(ws,comm,man_comm.get(comm,{}),ort_comm.get(comm,{}),
                            RECTIF_STATS_COMM.get(comm,{}).get("cap",800),COMM_HEX.get(comm,"1F3864"))

    # Onglets par usine
    for usine in usine_order:
        ws=wb.create_sheet(f"U - {usine}")
        _build_entity_sheet(ws,usine,man_usine.get(usine,{}),ort_usine.get(usine,{}),
                            USINE_CAPS.get(usine,500),USINE_HEX.get(usine,"1F3864"))

    ws=wb.create_sheet("Legende"); _build_legende_sheet(ws)
    buf=io.BytesIO(); wb.save(buf); buf.seek(0); return buf.read()

# ═══ POINT D'ENTRÉE PRINCIPAL ════════════════════════════════
def render_comparaison_tab(planning_df=None, df_to_xlsx_styled=None, sb=None):

    st.markdown("""
    <div style='background:#1a2332;border:1px solid #21262d;border-radius:12px;
    padding:16px 20px;margin-bottom:20px'>
      <div style='font-size:1.1rem;font-weight:700;color:#f0f6fc;margin-bottom:4px'>
        Plans Rectifiés — SOURCE PRINCIPALE de toutes les statistiques
      </div>
      <div style='font-size:.82rem;color:#8b949e'>
        Toutes les courbes, KPIs et tableaux sont basés sur les plans rectifiés.
        OR-Tools apparaît uniquement en pointillé comme comparaison.
        Source: reference_interne 13/06/2026 (FEDI 32736t · MAKKI 24310t · KHALIL 17455t · ACHREF 17486t · JILANI 7000t)
      </div>
    </div>""", unsafe_allow_html=True)

    # Session state
    if "comp_uploaded" not in st.session_state: st.session_state["comp_uploaded"]={}
    if "comp_raw"      not in st.session_state: st.session_state["comp_raw"]={}

    # Charger Supabase une seule fois
    if "comp_sb_loaded" not in st.session_state:
        _sb_comm,_sb_usine=_load_all_rectifie(sb)
        st.session_state["_sb_comm"]=_sb_comm
        st.session_state["_sb_usine"]=_sb_usine
        if _sb_comm:
            for _k,_v in _sb_comm.items():
                st.session_state["comp_uploaded"].setdefault(_k,_v)
        st.session_state["comp_sb_loaded"]=True

    sb_comm=st.session_state.get("_sb_comm",{})
    sb_usine=st.session_state.get("_sb_usine",{})

    # ── Construire les dictionnaires rectifiés/ORT ─────────────
    def _build_dicts():
        mc={}; oc={}
        for comm in RECTIF_STATS_COMM:
            mc[comm],_=_get_rectif_comm(comm,st.session_state["comp_uploaded"],sb_comm)
            oc[comm]=_get_ort_profile(comm,planning_df,"commercial")
        mu={}; ou={}
        for usine in USINE_CAPS:
            mu[usine],_=_get_rectif_usine(usine,sb_usine)
            ou[usine]=_get_ort_profile(usine,planning_df,"usine")
        return mc,oc,mu,ou

    # Auto-générer Excel au chargement si pas de cache
    if "comp_excel_cache" not in st.session_state:
        try:
            mc,oc,mu,ou=_build_dicts()
            st.session_state["comp_excel_cache"]=generate_comparison_excel(mc,oc,mu,ou,planning_df)
        except Exception as e:
            st.session_state["comp_excel_cache"]=None

    # ── ZONE UPLOAD ──────────────────────────────────────────
    st.subheader("Mettre à jour un plan rectifié")
    st.caption("Optionnel: les données reference_interne sont déjà chargées. Upload uniquement si mise à jour.")
    col_up,col_st=st.columns([3,2])
    with col_up:
        uploaded=st.file_uploader("Fichiers",type=["xlsx","xls"],accept_multiple_files=True,
                                  label_visibility="collapsed")
    if uploaded:
        for f in uploaded:
            cfn=_detect_comm_from_filename(f.name)
            cfd,daily,raw_or_err=_parse_rectification(f)
            comm=cfn or cfd
            if comm and daily:
                st.session_state["comp_uploaded"][comm]=daily
                if isinstance(raw_or_err,pd.DataFrame): st.session_state["comp_raw"][comm]=raw_or_err
                ok=_save_rectifie(sb,comm,daily,"commercial")
                if "comp_excel_cache" in st.session_state: del st.session_state["comp_excel_cache"]
                tot=sum(daily.values())
                st.success(f"{'✅ Supabase' if ok else '✔'}: {comm} — {len(daily)}j, {int(tot)}t (remplace reference_interne)")
            elif comm: st.warning(f"{f.name}: {raw_or_err}")
            else: st.warning(f"{f.name}: commercial non détecté")

    with col_st:
        st.markdown("**Statut source :**")
        for c in RECTIF_STATS_COMM:
            d,src=_get_rectif_comm(c,st.session_state["comp_uploaded"],sb_comm)
            t=sum(d.values()); icon="🟢" if "upload" in src else ("🔵" if "Supabase" in src else "⚪")
            st.markdown(f"{icon} **{c.split()[0]}** — {int(t)}t ({src})")

    # ── EXPORT EXCEL ─────────────────────────────────────────
    st.divider()
    col_exp1,col_exp2=st.columns([2,3])
    with col_exp1:
        if st.session_state.get("comp_excel_cache"):
            st.download_button(
                "📥 Télécharger Comparaison_Rectifie_vs_ORT_2026.xlsx",
                data=st.session_state["comp_excel_cache"],
                file_name="Comparaison_Rectifie_vs_ORT_2026.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,type="primary")
            st.caption("✅ Fichier prêt — SOURCE = Plans Rectifiés")
            if st.button("🔄 Régénérer Excel",use_container_width=True):
                if "comp_excel_cache" in st.session_state: del st.session_state["comp_excel_cache"]
                st.rerun()
        else:
            if st.button("⬇️ Générer Excel",type="primary",use_container_width=True):
                with st.spinner("Génération..."):
                    mc,oc,mu,ou=_build_dicts()
                    xl=generate_comparison_excel(mc,oc,mu,ou,planning_df)
                st.session_state["comp_excel_cache"]=xl; st.rerun()
    with col_exp2:
        st.markdown("""<div style='background:#161b22;border:1px solid #21262d;border-radius:8px;
        padding:12px 16px;font-size:12px;color:#8b949e'>
          <b style='color:#f0f6fc'>Contenu Excel :</b><br>
          📊 Synthèse Globale · 📅 Planning Journalier Rectifié<br>
          👤 5 onglets Commerciaux · 🏭 5 onglets Usines · Légende<br>
          <b style='color:#3b82f6'>Source :</b> Plans Rectifiés (reference_interne 13/06/2026)<br>
          <b style='color:#555'>Comparaison :</b> OR-Tools en pointillé uniquement<br>
          <b style='color:#e8543a'>ELFALLEH 5190t · ABIDA 8010t</b>
        </div>""",unsafe_allow_html=True)

    st.divider()

    # ── ONGLETS VISUELS ─────────────────────────────────────
    c1,c2,c3,c4=st.tabs(["📈 Par Commercial","🏭 Par Usine","📊 Stats globales","⚙️ Etat optimizer"])

    # ─── C1: PAR COMMERCIAL ───────────────────────────────────
    with c1:
        sel_comm=st.selectbox("Choisir un commercial",list(RECTIF_STATS_COMM.keys()),key="comp_sel_comm")
        color=COMM_COLORS[sel_comm]; cap=RECTIF_STATS_COMM[sel_comm]["cap"]
        man_dict,src_lbl=_get_rectif_comm(sel_comm,st.session_state["comp_uploaded"],sb_comm)
        ort_dict=_get_ort_profile(sel_comm,planning_df,"commercial")

        st.markdown(f"""<div style='background:#1a2332;border-left:4px solid {color};padding:8px 14px;
        border-radius:0 8px 8px 0;margin-bottom:12px;font-size:12px;color:#8b949e'>
        Source: <b style='color:#f0f6fc'>{src_lbl}</b> |
        Total: <b style='color:{color}'>{int(sum(_ns(man_dict).values()))}t</b> |
        Max/j: <b style='color:{color}'>{int(max(_ns(man_dict).values()) if man_dict else 0)}t</b>
        </div>""",unsafe_allow_html=True)

        _kpi_row(sel_comm,man_dict,ort_dict,cap)
        st.plotly_chart(_build_chart(sel_comm,man_dict,ort_dict,color,cap),use_container_width=True)

        # Tableau jour par jour
        st.markdown(f"**Tableau jour par jour — {sel_comm}** (source: {src_lbl})")
        ms=_ns(man_dict); os_=_ns(ort_dict)
        pivot_rows=[]
        for d in SEASON:
            dk=d.date(); mv=ms.get(str(dk),0); ov=os_.get(str(dk),0)
            if mv==0 and ov==0: continue
            ec=mv-ov
            pivot_rows.append({"Date":d.strftime("%d/%m/%Y"),"Jour":DAYS_FR[d.weekday()],
                               "PIC":"⚡" if PIC_S<=dk<=PIC_E else "",
                               "Rectifié (t)":int(mv) if mv>0 else "",
                               "OR-Tools (t)":int(ov) if ov>0 else "",
                               "Ecart (t)":f"{int(ec):+d}" if (mv>0 or ov>0) else "",
                               "Statut":"OK" if abs(ec)<=50 else ("Rect+" if ec>0 else "OR-T+")})
        if pivot_rows:
            st.dataframe(pd.DataFrame(pivot_rows),use_container_width=True,height=320,hide_index=True,
                         column_config={"Rectifié (t)":st.column_config.NumberColumn("Rectifié (t)",format="%d t"),
                                        "OR-Tools (t)":st.column_config.NumberColumn("OR-Tools (t)",format="%d t")})

        # Profil par agriculteur si upload
        if sel_comm in st.session_state["comp_raw"]:
            st.markdown("---")
            st.markdown(f"**Profil par agriculteur — {sel_comm}**")
            df_raw=st.session_state["comp_raw"][sel_comm]
            agri_stats=df_raw.groupby("agriculteur").agg(
                total=("tonnes","sum"),max_jour=("tonnes","max"),jours=("date","nunique")
            ).reset_index().sort_values("total",ascending=False)
            agri_stats.columns=["Agriculteur","Total (t)","Max/jour (t)","Jours actifs"]
            agri_stats[["Total (t)","Max/jour (t)"]]=agri_stats[["Total (t)","Max/jour (t)"]].round(0).astype(int)
            fig_a=px.bar(agri_stats,x="Total (t)",y="Agriculteur",orientation="h",
                         height=max(350,len(agri_stats)*28+80),color="Total (t)",
                         color_continuous_scale="Viridis",title=f"Tonnage par agriculteur — {sel_comm}",
                         template="plotly_dark",text="Total (t)")
            fig_a.update_traces(textposition="outside",texttemplate="%{x:.0f}t")
            fig_a.update_layout(paper_bgcolor="#161b22",showlegend=False,margin=dict(l=230,r=60,t=50,b=40))
            st.plotly_chart(fig_a,use_container_width=True)

    # ─── C2: PAR USINE ─────────────────────────────────────────
    with c2:
        sel_usine=st.selectbox("Choisir une usine",["SICAM","TUCAL","COMOCAP","ELFALLEH","ABIDA"],key="comp_usine_sel")
        cap=USINE_CAPS[sel_usine]; us=RECTIF_STATS_USINE.get(sel_usine,{})
        man_u,src_u=_get_rectif_usine(sel_usine,sb_usine)
        ort_u=_get_ort_profile(sel_usine,planning_df,"usine")

        total_u=us.get("total",sum(man_u.values()) if man_u else 0)
        max_u=us.get("max_j",max(man_u.values()) if man_u else 0)
        jours_u=us.get("n_j",sum(1 for v in man_u.values() if v>0))
        ot_u=sum(ort_u.values()) if ort_u else 0

        k1,k2,k3,k4,k5=st.columns(5)
        k1.metric("Total saison (Rectifié)",f"{int(total_u)}t")
        k2.metric("Max journalier",f"{max_u}t",
                  delta=f"+{max_u-cap}t vs cap {cap}" if max_u>cap else f"cap {cap}t/j OK",
                  delta_color="inverse" if max_u>cap else "normal")
        k3.metric("Cap officiel",f"{cap}t/j")
        k4.metric("OR-Tools total",f"{int(ot_u)}t",delta=f"{int(total_u-ot_u):+d}t")
        k5.metric("Jours actifs",f"{jours_u}j")
        st.caption(f"Source : **{src_u}**")

        # Courbe principale = RECTIFIÉ
        man_u_s=_ns(man_u); ort_u_s=_ns(ort_u)
        dates_rect=[pd.Timestamp(k) for k in sorted(man_u_s.keys())]
        vals_rect=[float(man_u_s[k]) for k in sorted(man_u_s.keys())]

        fig_u=go.Figure()
        fig_u.add_trace(go.Bar(x=dates_rect,y=vals_rect,name="Plan Rectifié (SOURCE)",
                               marker_color="#3b82f6",marker_line_width=0))
        if ort_u_s:
            fig_u.add_trace(go.Scatter(
                x=[pd.Timestamp(d) for d in sorted(ort_u_s.keys())],
                y=[float(ort_u_s[d]) for d in sorted(ort_u_s.keys())],
                name="OR-Tools (comparaison)",line=dict(color="#888888",width=2,dash="dash"),mode="lines"))
        fig_u.add_hline(y=cap,line_dash="dash",line_color="#e8543a",line_width=1.5,
                        annotation_text=f"Cap officiel:{cap}t/j",annotation_position="top right",
                        annotation_font_color="#e8543a")
        fig_u.add_vrect(x0=pd.Timestamp("2026-07-01"),x1=pd.Timestamp("2026-07-15"),
                        fillcolor="rgba(245,166,35,0.08)",line_width=0,
                        annotation_text="PIC",annotation_position="top left",annotation_font_color="#f5a623")
        fig_u.update_layout(
            title=f"{sel_usine} — Plan Rectifié (SOURCE) vs OR-Tools",
            template="plotly_dark",plot_bgcolor="#0d1117",paper_bgcolor="#161b22",
            height=360,hovermode="closest",legend=dict(orientation="h",yanchor="bottom",y=1.02))
        st.plotly_chart(fig_u,use_container_width=True)

        # Tableau jour par jour USINE
        st.markdown(f"**Tableau jour par jour — {sel_usine}**")
        pivot_u=[]
        for d in SEASON:
            dk=d.date(); mv=float(man_u_s.get(str(dk),0)); ov=float(ort_u_s.get(str(dk),0))
            if mv==0 and ov==0: continue
            ec=mv-ov
            pivot_u.append({"Date":d.strftime("%d/%m/%Y"),"Jour":DAYS_FR[d.weekday()],
                            "PIC":"⚡" if PIC_S<=dk<=PIC_E else "",
                            "Rectifié (t)":int(mv) if mv>0 else "",
                            "OR-Tools (t)":int(ov) if ov>0 else "",
                            "Ecart (t)":f"{int(ec):+d}" if (mv>0 or ov>0) else "",
                            "Statut Cap":"DEPASSE" if (PIC_S<=dk<=PIC_E and max(mv,ov)>cap) else ("OK" if PIC_S<=dk<=PIC_E else "")})
        if pivot_u:
            st.dataframe(pd.DataFrame(pivot_u),use_container_width=True,height=320,hide_index=True)

        notes={"SICAM":["Démarrage 20/06 (140t) → Pic 24/07 (1650t)","Total rectifié: 44871t / 67 jours"],
               "TUCAL":["Démarrage 24/06 → Plateau 590-700t/j juillet","Total rectifié: 20525t / 52 jours"],
               "COMOCAP":["Pic tardif 27/07 (875t)","Total rectifié: 20391t / 50 jours"],
               "ELFALLEH":["Total réel 5190t / 38 jours / max 240t/j","Cap officiel 150t/j (jours doubles autorisés)"],
               "ABIDA":["Total réel 8010t / max 320t/j","Démarrage tardif juillet → déclin fin août"]}
        st.markdown("**Données terrain :**")
        for note in notes.get(sel_usine,[]): st.markdown(f"• {note}")

    # ─── C3: STATS GLOBALES ────────────────────────────────────
    with c3:
        st.markdown("### Statistiques globales — Plans Rectifiés (SOURCE)")
        rows_cmp=[]; mt_list=[]; ot_list=[]
        for comm in RECTIF_STATS_COMM:
            d,src=_get_rectif_comm(comm,st.session_state["comp_uploaded"],sb_comm)
            ds=_ns(d); od=_ns(_get_ort_profile(comm,planning_df,"commercial"))
            mt=sum(ds.values()); ot=sum(od.values()) if od else 0
            mm=max(ds.values()) if ds else 0; om=max(od.values()) if od else 0
            ec=mt-ot; pct=round(ec/ot*100,1) if ot>0 else 0
            mt_list.append(mt); ot_list.append(ot)
            rows_cmp.append({"Commercial":comm,"Source":src,"Total Rectifié(t)":int(mt),
                             "Total OR-T(t)":int(ot),"Ecart(t)":f"{int(ec):+d}",
                             "Max Rect":f"{int(mm)}t","Jours":sum(1 for v in ds.values() if v>0),
                             "Cap PIC":f"{RECTIF_STATS_COMM[comm]['cap']}t/j"})
        st.dataframe(pd.DataFrame(rows_cmp),use_container_width=True,hide_index=True)

        fig_cmp=go.Figure()
        fig_cmp.add_trace(go.Bar(name="Plan Rectifié",x=list(RECTIF_STATS_COMM.keys()),y=mt_list,
                                 marker_color="#3b82f6",text=[f"{int(v)}t" for v in mt_list],textposition="outside"))
        fig_cmp.add_trace(go.Bar(name="OR-Tools",x=list(RECTIF_STATS_COMM.keys()),y=ot_list,
                                 marker_color="#555555",text=[f"{int(v)}t" for v in ot_list],textposition="outside"))
        fig_cmp.update_layout(barmode="group",template="plotly_dark",paper_bgcolor="#161b22",
                              plot_bgcolor="#0d1117",height=380,yaxis_title="Tonnes",
                              title="Comparaison totaux — Plans Rectifiés (bleu) vs OR-Tools (gris)",
                              legend=dict(orientation="h",yanchor="bottom",y=1.02))
        st.plotly_chart(fig_cmp,use_container_width=True)

        st.markdown("---")
        st.markdown("### Résumé usines — données réelles")
        rows_u=[]
        for usine,us in RECTIF_STATS_USINE.items():
            rows_u.append({"Usine":usine,"Total Rectifié(t)":us["total"],
                           "Max/j Rectifié":f"{us['max_j']}t","Jours":us["n_j"],
                           "Cap officiel":f"{us['cap_officiel']}t/j",
                           "Dépassement cap":"OUI" if us["max_j"]>us["cap_officiel"] else "NON"})
        st.dataframe(pd.DataFrame(rows_u),use_container_width=True,hide_index=True)

    # ─── C4: ETAT OPTIMIZER ────────────────────────────────────
    with c4:
        st.markdown("### Etat optimizer_v2.py — Configuration stable")
        st.success("Configuration OPTIMALE — ne pas modifier")
        st.markdown("""
| Paramètre | Valeur | Statut |
|---|---|---|
| Borne journalière `_ub_day` | `day_planned × SCALE × 2.0` | OPTIMAL |
| `FACTORY_OVERFLOW_WEIGHT` | `2000` | Appliqué |
| Arrondi tonnage | `10t` | Stable |
| Correction post-traitement | Désactivée | OK |
        """)
        st.warning("**x1.1 et x1.5 causent INFEASIBLE** — testés et prouvés. Ne pas modifier.")
        st.code("""Status: OPTIMAL (~2s) | 2630 rows | 96 676t (-0.32%)
Les plans RECTIFIÉS = corrections manuelles terrain sur la sortie OR-Tools
Ils tiennent compte des contraintes réelles (accessibilité, distances, etc.)""")# ============================================================
# SOURCE UNIQUE DE VERITE — Plans Rectifies > OR-Tools
# Ajoute en fin de fichier — n'interfere avec rien d'existant.
# A utiliser depuis dashboard_phase10.py juste apres construction de 'p'.
# ============================================================
def _rectif_load_from_supabase(sb):
    """Charge tous les plans rectifies depuis Supabase (independant des autres helpers)."""
    if sb is None:
        return {}, {}
    try:
        data = sb.table("plan_rectifie").select(
            "entity_type,entity_name,date,tonnes").execute().data
    except Exception:
        return {}, {}
    comm, usine = {}, {}
    for row in (data or []):
        et = row.get("entity_type", "commercial")
        en = row.get("entity_name")
        d  = str(row.get("date"))
        t  = float(row.get("tonnes", 0) or 0)
        target = comm if et == "commercial" else usine
        target.setdefault(en, {})[d] = t
    return comm, usine


def _rectif_get_daily(entity_name, entity_type, sb):
    """
    Retourne (dict {date_str: tonnes}, source) pour une entite,
    ou (None, None) si aucune correction n'existe pour elle.
    Priorite: upload de la session en cours > Supabase (persistant).
    """
    if entity_type == "commercial":
        uploaded = st.session_state.get("comp_uploaded", {})
        if entity_name in uploaded:
            d = uploaded[entity_name]
            return {str(k): float(v) for k, v in d.items()}, "upload session"
    sb_comm, sb_usine = _rectif_load_from_supabase(sb)
    table = sb_comm if entity_type == "commercial" else sb_usine
    if entity_name in table:
        return {str(k): float(v) for k, v in table[entity_name].items()}, "Supabase"
    return None, None


def build_effective_planning(planning_df, sb):
    """
    SOURCE UNIQUE: priorite Plans Rectifies > OR-Tools (planning_df original).
    Retourne un DataFrame avec EXACTEMENT les memes colonnes que planning_df,
    ou les Tonnes/Jour sont corrigees pour matcher les plans rectifies
    quand ils existent (niveau commercial). Sans correction -> OR-Tools inchange.

    A appeler UNE SEULE FOIS juste apres avoir construit 'p' dans le dashboard.
    Tous les tabs qui lisent 'p' ensuite refletent automatiquement la correction.
    """
    if planning_df is None or planning_df.empty:
        return planning_df
    if "Date" not in planning_df.columns or "Commercial" not in planning_df.columns:
        return planning_df

    df = planning_df.copy()
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")

    raw_sess = st.session_state.get("comp_raw", {})  # detail agriculteur si upload recent
    out_parts = []
    keep_mask = pd.Series(True, index=df.index)

    for comm in df["Commercial"].dropna().unique():
        rect_daily, _src = _rectif_get_daily(comm, "commercial", sb)
        if rect_daily is None:
            continue  # pas de correction -> garder OR-Tools tel quel pour ce commercial

        mask = df["Commercial"] == comm
        sub = df[mask].copy()
        keep_mask &= ~mask

        if comm in raw_sess:
            # Detail par agriculteur disponible (upload de cette session) -> remplacement complet
            raw_df = raw_sess[comm][["agriculteur", "date", "tonnes"]].copy()
            raw_df = raw_df.rename(columns={"agriculteur": "Agriculteur",
                                             "date": "Date", "tonnes": "Tonnes/Jour"})
            raw_df["Date"] = pd.to_datetime(raw_df["Date"])
            raw_df["Commercial"] = comm
            if "Usine" in sub.columns and not sub.empty:
                usine_map = sub.groupby("Agriculteur")["Usine"].agg(
                    lambda s: s.mode().iat[0] if not s.mode().empty else "")
                default_usine = sub["Usine"].mode().iat[0] if not sub["Usine"].mode().empty else ""
                raw_df["Usine"] = raw_df["Agriculteur"].map(usine_map).fillna(default_usine)
            else:
                raw_df["Usine"] = ""
            for col in df.columns:
                if col not in raw_df.columns:
                    raw_df[col] = sub[col].iloc[0] if (col in sub.columns and not sub.empty) else ""
            out_parts.append(raw_df[df.columns])
        else:
            # Seulement les totaux journaliers -> mise a l'echelle proportionnelle
            # (preserve la repartition entre agriculteurs/usines, corrige le total du jour)
            sub["_d"] = sub["Date"].dt.date
            daily_ort = sub.groupby("_d")["Tonnes/Jour"].sum()
            for d_key, ort_total in daily_ort.items():
                rect_total = rect_daily.get(str(d_key))
                if rect_total is None:
                    continue  # pas de correction ce jour precis -> garder OR-Tools
                ratio = (rect_total / ort_total) if ort_total > 0 else 0
                day_mask = sub["_d"] == d_key
                sub.loc[day_mask, "Tonnes/Jour"] = sub.loc[day_mask, "Tonnes/Jour"] * ratio
            sub = sub.drop(columns=["_d"])
            out_parts.append(sub)

    if not out_parts:
        return df
    return pd.concat([df[keep_mask]] + out_parts, ignore_index=True)


_FLEET_AVAILABILITY_RECTIF = {
    "SICAM":    {"PL": 48, "PPL": 6,  "SEMI": 13, "TRACTEUR": 0},
    "TUCAL":    {"PL": 17, "PPL": 0,  "SEMI": 2,  "TRACTEUR": 0},
    "COMOCAP":  {"PL": 6,  "PPL": 14, "SEMI": 3,  "TRACTEUR": 10},
    "ABIDA":    {"PL": 1,  "PPL": 0,  "SEMI": 2,  "TRACTEUR": 0},
    "ELFALLEH": {"PL": 0,  "PPL": 2,  "SEMI": 0,  "TRACTEUR": 0},
}


def generate_planning_wide_excel(effective_df):
    """
    Excel avec colonnes:
    Date | Commercial | Agriculteur | Usine | Tonnes/Jour | Type Vehicule |
    Vehicules Requis | Disponibles | Manquants (a louer) | Nb Voyages | Pic de Recolte |
    [une colonne par date de la saison]
    Construit directement depuis le planning EFFECTIF (rectifie prioritaire).
    """
    from openpyxl import Workbook
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    wb = Workbook(); wb.remove(wb.active)
    ws = wb.create_sheet("Planning Rectifie")

    if effective_df is None or effective_df.empty:
        ws["A1"] = "Aucune donnee disponible"
        buf = io.BytesIO(); wb.save(buf); buf.seek(0)
        return buf.read()

    df = effective_df.copy()
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")

    SEASON_DATES = list(pd.date_range("2026-06-20", "2026-08-25", freq="D"))
    PIC_S_ = pd.Timestamp("2026-07-01").date()
    PIC_E_ = pd.Timestamp("2026-07-15").date()
    FIXED = ["Date", "Commercial", "Agriculteur", "Usine", "Tonnes/Jour", "Type Véhicule",
             "Véhicules Requis", "Disponibles", "Manquants (à louer)", "Nb Voyages", "Pic de Récolte"]
    DATE_HDRS = [d.strftime("%d/%m/%Y") for d in SEASON_DATES]
    ALL_HDRS = FIXED + DATE_HDRS
    N = len(ALL_HDRS)

    HDR  = PatternFill("solid", start_color="1F3864", end_color="1F3864")
    PICF = PatternFill("solid", start_color="FFF2CC", end_color="FFF2CC")
    BLUF = PatternFill("solid", start_color="DEEBF7", end_color="DEEBF7")
    GRNF = PatternFill("solid", start_color="E2EFDA", end_color="E2EFDA")
    ALTF = PatternFill("solid", start_color="F8F9FA", end_color="F8F9FA")
    WHTF = PatternFill("solid", start_color="FFFFFF", end_color="FFFFFF")
    REDF = PatternFill("solid", start_color="FFC7CE", end_color="FFC7CE")
    THIN = Side(style="thin", color="CCCCCC")
    BORD = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
    CTR  = Alignment(horizontal="center", vertical="center")

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=min(N, 30))
    c = ws.cell(1, 1)
    c.value = "Planning Journalier — Source: Plans Rectifies (priorite) + OR-Tools"
    c.font = Font(bold=True, color="FFFFFF", size=12); c.fill = HDR; c.alignment = CTR
    ws.row_dimensions[1].height = 26

    for ci, h in enumerate(ALL_HDRS, 1):
        cell = ws.cell(2, ci); cell.value = h; cell.border = BORD; cell.alignment = CTR
        if ci <= len(FIXED):
            cell.fill = HDR; cell.font = Font(bold=True, color="FFFFFF", size=9)
        else:
            d_obj = SEASON_DATES[ci - len(FIXED) - 1].date()
            cell.fill = PICF if PIC_S_ <= d_obj <= PIC_E_ else BLUF
            cell.font = Font(bold=True, size=8)
    ws.row_dimensions[2].height = 32

    group_cols = [c for c in ["Commercial", "Agriculteur", "Usine"] if c in df.columns]
    if not group_cols:
        group_cols = ["Commercial"]

    row = 3
    for keys, grp in df.groupby(group_cols, dropna=False):
        if not isinstance(keys, tuple):
            keys = (keys,)
        kd = dict(zip(group_cols, keys))
        comm  = kd.get("Commercial", "")
        agri  = kd.get("Agriculteur", "")
        usine = kd.get("Usine", "")

        daily = {}
        for _, r in grp.iterrows():
            if pd.notna(r["Date"]):
                dk = str(r["Date"].date())
                daily[dk] = daily.get(dk, 0) + float(r.get("Tonnes/Jour", 0) or 0)
        total = sum(daily.values())
        if total <= 0:
            continue

        r0 = grp.iloc[0]
        tv = str(r0.get("Type Véhicule", "") or "")
        vr_raw = r0.get("Véhicules Requis", "")
        nv = r0.get("Nb Voyages", "")
        vr_n = int(pd.to_numeric(vr_raw, errors="coerce") or 0)
        fleet = _FLEET_AVAILABILITY_RECTIF.get(str(usine).upper().strip(), {})
        dispo = fleet.get(str(tv).upper().strip(), 0) if vr_n > 0 else 0
        manque = max(0, vr_n - dispo) if vr_n > 0 else 0
        is_pic = any(PIC_S_ <= pd.Timestamp(k).date() <= PIC_E_ for k, v in daily.items() if v > 0)

        alt = row % 2 == 0
        base = ALTF if alt else WHTF
        fixed_vals = ["", comm, agri, usine, round(total, 0), tv,
                      vr_n if vr_n > 0 else "", dispo if dispo else "",
                      manque if manque else "", nv, "⚡ PIC" if is_pic else ""]
        for ci, val in enumerate(fixed_vals, 1):
            cell = ws.cell(row, ci); cell.value = val; cell.border = BORD; cell.alignment = CTR
            if ci == 9 and isinstance(val, int) and val > 0:
                cell.fill = REDF; cell.font = Font(bold=True, size=9, color="9C0006")
            else:
                cell.fill = base; cell.font = Font(size=9, bold=(ci == 2))

        for di, d in enumerate(SEASON_DATES):
            dk = str(d.date())
            val = daily.get(dk, 0)
            ci2 = len(FIXED) + di + 1
            cell = ws.cell(row, ci2)
            cell.value = int(round(val, 0)) if val > 0 else ""
            cell.border = BORD; cell.alignment = CTR
            is_pic_d = PIC_S_ <= d.date() <= PIC_E_
            cell.fill = PICF if (is_pic_d and val > 0) else (GRNF if val > 0 else base)
            cell.font = Font(size=8)
        ws.row_dimensions[row].height = 15
        row += 1

    widths = [12, 16, 26, 12, 11, 12, 11, 11, 13, 10, 11]
    for ci, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(ci)].width = w
    for di in range(len(SEASON_DATES)):
        ws.column_dimensions[get_column_letter(len(FIXED) + di + 1)].width = 8
    ws.freeze_panes = "L3"

    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    return buf.read()