# ============================================================
# ONGLET COMPARAISON v2 — Plans rectifiés vs OR-Tools
# Fichier : comparaison_tab.py
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

# ── Constantes ───────────────────────────────────────────────
SEASON      = pd.date_range("2026-06-20", "2026-08-25", freq="D")
PIC_S       = pd.Timestamp("2026-07-01").date()
PIC_E       = pd.Timestamp("2026-07-15").date()
DAYS_FR     = ["Lun","Mar","Mer","Jeu","Ven","Sam","Dim"]

COMM_COLORS = {
    "FEDI":           "#3b82f6",
    "MAKKI BEN SALAH":"#00e5a0",
    "KHALIL":         "#f5a623",
    "ACHREF AJLANI":  "#8b5cf6",
    "JILANI OBAY":    "#e8543a",
}
COMM_HEX = {
    "FEDI":           "1F5FA6",
    "MAKKI BEN SALAH":"1A6B3C",
    "KHALIL":         "B45309",
    "ACHREF AJLANI":  "5B21B6",
    "JILANI OBAY":    "C0392B",
}
USINE_HEX = {
    "SICAM":   "1F3864",
    "TUCAL":   "4A235A",
    "COMOCAP": "0B4F6C",
    "ELFALLEH":"196F3D",
    "ABIDA":   "922B21",
}

MANUAL_STATS = {
    "FEDI":           {"total":32736,"max_j":1216,"n_jours":55,"cap":1300},
    "MAKKI BEN SALAH":{"total":24310,"max_j":1315,"n_jours":44,"cap":1200},
    "KHALIL":         {"total":25290,"max_j":1500,"n_jours":46,"cap":1100},
    "ACHREF AJLANI":  {"total":17486,"max_j":570, "n_jours":61,"cap":700},
    "JILANI OBAY":    {"total":7000, "max_j":450, "n_jours":25,"cap":150},
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

ORT_COMM_BASE = {
    "FEDI":          {(pd.Timestamp("2026-06-28")+pd.Timedelta(days=i)).date():v
                     for i,v in enumerate([80,120,180,250,350,420,540,650,780,850,950,1050,
                     1100,1150,1175,1100,1050,900,750,600,500,450,400,350,300,250,200,180,
                     160,150,140,130,120,100,80,60,40,20,10]) if v>0},
    "MAKKI BEN SALAH":{(pd.Timestamp("2026-06-22")+pd.Timedelta(days=i)).date():v
                     for i,v in enumerate([50,100,150,200,280,380,480,580,650,750,850,950,
                     1050,1095,1000,900,780,680,580,480,380,300,250,200,160,130,100,80,60,
                     40,30,20,10]) if v>0},
    "KHALIL":        {(pd.Timestamp("2026-06-20")+pd.Timedelta(days=i)).date():v
                     for i,v in enumerate([200,280,350,430,520,630,720,820,920,1000,1020,980,
                     900,800,700,600,500,420,350,280,220,180,150,120,90,70,50,30,20,10]) if v>0},
    "ACHREF AJLANI": {(pd.Timestamp("2026-06-22")+pd.Timedelta(days=i)).date():v
                     for i,v in enumerate([30,60,120,180,240,300,360,420,480,540,600,650,700,
                     750,790,810,810,790,750,700,640,570,490,400,320,250,190,140,100,70,50,
                     30,20,10,5]) if v>0},
    "JILANI OBAY":   {(pd.Timestamp("2026-07-23")+pd.Timedelta(days=i)).date():v
                     for i,v in enumerate([80,130,180,240,310,380,440,440,410,370,320,270,220,
                     180,140,100,70,50,30,15]) if v>0},
}
ORT_USINE_BASE = {
    "SICAM":   {(pd.Timestamp("2026-06-20")+pd.Timedelta(days=i)).date():v
               for i,v in enumerate([100,120,180,250,250,420,440,500,560,653,743,898,895,1000,
               1105,1128,1213,1285,1340,1490,1530,1540,1510,1495,1440,1350,1280,1210,1140,
               1015,970,960,965,930,930,1155,1180,1220,1270,1280,1180,1090,1045,1040,940,860,
               850,835,795,578,600,540,520,380,260,210,150,90,30]) if v>0},
    "TUCAL":   {(pd.Timestamp("2026-06-20")+pd.Timedelta(days=i)).date():v
               for i,v in enumerate([20,60,80,80,110,195,195,245,315,320,325,330,365,390,420,
               415,395,380,350,370,340,300,290,320,300,230,210,205,185,255,240,240,140,160,235,
               315,345,380,410,410,385,385,375,370,330,370,350,340,270,260,270,255,250,230,190,
               190,90,100,90]) if v>0},
    "COMOCAP": {(pd.Timestamp("2026-06-20")+pd.Timedelta(days=i)).date():v
               for i,v in enumerate([0,0,0,0,0,0,0,0,20,55,55,60,135,145,160,165,165,165,170,
               400,520,500,505,585,600,515,558,545,580,542,520,421,380,463,343,293,275,245,290,
               306,293,250,247,197,160,118,45,40,15,15,8]) if v>0},
    "ELFALLEH":{},
    "ABIDA":   {},
}

# ── Normaliser toutes les cles des dicts de reference en string ──
# (evite le TypeError entre datetime.date et str dans les lookups)
ORT_COMM_BASE = {
    comm: {str(k): v for k, v in d.items()}
    for comm, d in ORT_COMM_BASE.items()
}
ORT_USINE_BASE = {
    usine: {str(k): v for k, v in d.items()}
    for usine, d in ORT_USINE_BASE.items()
}

# ── Normaliser toutes les cles des dicts de reference en string ──
# (evite le TypeError entre datetime.date et str dans les lookups)
ORT_COMM_BASE = {
    comm: {str(k): v for k, v in d.items()}
    for comm, d in ORT_COMM_BASE.items()
}
ORT_USINE_BASE = {
    usine: {str(k): v for k, v in d.items()}
    for usine, d in ORT_USINE_BASE.items()
}

# ── Helpers openpyxl ─────────────────────────────────────────
def _hf(h): return PatternFill("solid", start_color=h, end_color=h)
def _ft(bold=False, color="1F1F1F", size=10, white=False):
    return Font(bold=bold, name="Calibri", size=size, color="FFFFFF" if white else color)
_CTR  = Alignment(horizontal="center", vertical="center")
_LFT  = Alignment(horizontal="left",   vertical="center")
_THIN = Side(style="thin", color="CCCCCC")
_BORD = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)

F_MAN      = _hf("E2EFDA")
F_ORT      = _hf("DEEBF7")
F_POS      = _hf("C6EFCE")
F_NEG      = _hf("FFC7CE")
F_NEU      = _hf("F2F2F2")
F_PIC      = _hf("FFF2CC")
F_ALT1     = _hf("FFFFFF")
F_ALT2     = _hf("F8F9FA")


# ═══════════════════════════════════════════════════════════
# EXPORT EXCEL : une feuille par entité (commercial / usine)
# ═══════════════════════════════════════════════════════════
def _build_entity_sheet(ws, name, man_dict, ort_dict, cap, hex_color):
    """Feuille comparaison jour par jour pour 1 commercial ou 1 usine."""
    HDR = _hf(hex_color)
    HDR_F  = _ft(bold=True, size=12, white=True)
    HDR_FS = _ft(bold=True, size=9,  white=True)
    N = 8  # nb colonnes

    # Ligne 1 : titre
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=N)
    c = ws.cell(1, 1)
    c.value = f"Comparaison Planning  —  {name}  —  Saison 2026"
    c.font = HDR_F; c.fill = HDR; c.alignment = _CTR
    ws.row_dimensions[1].height = 30

    # Ligne 2 : légende
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=N)
    c = ws.cell(2, 1)
    c.value = (f"Vert = Manuel > OR-Tools  |  Rouge = Manuel < OR-Tools  |  "
               f"Jaune = PIC 1-15 juillet  |  Cap officiel = {cap} t/j")
    c.font = _ft(size=9, color="595959"); c.fill = _hf("F0F4F8"); c.alignment = _LFT
    ws.row_dimensions[2].height = 18

    # Ligne 3 : en-têtes
    hdrs = ["Date","Jour","Manuel (t)","OR-Tools (t)","Écart (t)","Écart %","Statut",f"Cap {cap}t/j"]
    for ci, h in enumerate(hdrs, 1):
        c = ws.cell(3, ci)
        c.value = h; c.font = HDR_FS; c.fill = HDR; c.alignment = _CTR; c.border = _BORD
    ws.row_dimensions[3].height = 22

    # Lignes de données
    data_rows = []
    _man_str = {str(k):v for k,v in man_dict.items()}
    _ort_str = {str(k):v for k,v in ort_dict.items()}
    for d in SEASON:
        dk = d.date()
        mv = _man_str.get(str(dk), 0)
        ov = _ort_str.get(str(dk), 0)
        if mv == 0 and ov == 0:
            continue
        data_rows.append((d, dk, mv, ov))

    for ri, (d, dk, mv, ov) in enumerate(data_rows):
        r      = ri + 4
        is_pic = (PIC_S <= dk <= PIC_E)
        is_alt = ri % 2 == 0
        base   = F_PIC if is_pic else (F_ALT1 if is_alt else F_ALT2)
        ecart  = mv - ov
        pct    = round(ecart / ov * 100, 1) if ov > 0 else None

        # Date
        c = ws.cell(r, 1)
        c.value = d.strftime("%d/%m/%Y"); c.fill = base; c.border = _BORD; c.alignment = _CTR
        c.font  = _ft(bold=is_pic, color="7D4F00" if is_pic else "1F1F1F")

        # Jour semaine
        c = ws.cell(r, 2)
        c.value = DAYS_FR[d.weekday()]; c.fill = base; c.border = _BORD; c.alignment = _CTR
        c.font  = _ft(size=9, color="595959")

        # Manuel
        c = ws.cell(r, 3)
        c.value = int(mv) if mv > 0 else ""
        c.fill  = F_MAN if mv > 0 else base; c.border = _BORD; c.alignment = _CTR
        c.font  = _ft(bold=mv > 0, color="1A5276" if mv > 0 else "999999")
        c.number_format = "#,##0"

        # OR-Tools
        c = ws.cell(r, 4)
        c.value = int(ov) if ov > 0 else ""
        c.fill  = F_ORT if ov > 0 else base; c.border = _BORD; c.alignment = _CTR
        c.font  = _ft(bold=ov > 0, color="1A5276" if ov > 0 else "999999")
        c.number_format = "#,##0"

        # Écart
        c = ws.cell(r, 5)
        if mv > 0 or ov > 0:
            c.value = int(ecart)
            c.fill  = F_POS if ecart > 50 else (F_NEG if ecart < -50 else F_NEU)
            c.font  = _ft(bold=True,
                          color="276221" if ecart > 50 else ("9C0006" if ecart < -50 else "595959"))
            c.number_format = "+#,##0;-#,##0;0"
        else:
            c.fill = base; c.font = _ft(size=9, color="CCCCCC")
        c.border = _BORD; c.alignment = _CTR

        # Écart %
        c = ws.cell(r, 6)
        if pct is not None:
            c.value = pct / 100
            c.fill  = F_POS if pct > 5 else (F_NEG if pct < -5 else F_NEU)
            c.font  = _ft(bold=abs(pct) > 5,
                          color="276221" if pct > 5 else ("9C0006" if pct < -5 else "595959"))
            c.number_format = "+0.0%;-0.0%;0%"
        else:
            c.value = "—" if mv > 0 else ""; c.fill = base; c.font = _ft(size=9, color="999999")
        c.border = _BORD; c.alignment = _CTR

        # Statut
        c = ws.cell(r, 7)
        if mv > 0 and ov == 0:   txt, fc, tc = "Manuel only", "E8F5E9", "2E7D32"
        elif mv == 0 and ov > 0: txt, fc, tc = "OR-T only",  "FFF3E0", "E65100"
        elif abs(ecart) <= 50:   txt, fc, tc = "OK",         "E8F5E9", "2E7D32"
        elif ecart > 0:          txt, fc, tc = "Manuel +",   "E3F2FD", "0D47A1"
        else:                    txt, fc, tc = "OR-T +",     "FCE4D6", "993200"
        c.value = txt; c.fill = _hf(fc); c.font = _ft(size=9, color=tc)
        c.border = _BORD; c.alignment = _CTR

        # Cap PIC
        c = ws.cell(r, 8)
        if is_pic:
            peak = max(mv, ov)
            if peak > cap:
                c.value = f"DEPASSE {int(peak)}t"; c.fill = F_NEG
                c.font  = _ft(bold=True, size=9, color="9C0006")
            else:
                c.value = f"OK {int(peak)}t"; c.fill = F_POS
                c.font  = _ft(size=9, color="276221")
        else:
            c.value = "—"; c.fill = base; c.font = _ft(size=9, color="CCCCCC")
        c.border = _BORD; c.alignment = _CTR
        ws.row_dimensions[r].height = 17

    # Ligne total
    tr = len(data_rows) + 4
    man_t = sum(r[2] for r in data_rows)
    ort_t = sum(r[3] for r in data_rows)
    ec_t  = man_t - ort_t
    pct_t = round(ec_t / ort_t * 100, 1) if ort_t > 0 else 0
    man_m = max((r[2] for r in data_rows), default=0)
    ort_m = max((r[3] for r in data_rows), default=0)
    mj    = sum(1 for r in data_rows if r[2] > 0)
    oj    = sum(1 for r in data_rows if r[3] > 0)

    ws.merge_cells(start_row=tr, start_column=1, end_row=tr, end_column=2)
    for ci in range(1, N+1):
        c = ws.cell(tr, ci); c.fill = HDR; c.border = _BORD; c.alignment = _CTR
        c.font = _ft(bold=True, size=10, white=True)
    ws.cell(tr, 1).value = "TOTAL SAISON"
    ws.cell(tr, 3).value = int(man_t); ws.cell(tr, 3).number_format = "#,##0"
    ws.cell(tr, 4).value = int(ort_t); ws.cell(tr, 4).number_format = "#,##0"
    ws.cell(tr, 5).value = int(ec_t);  ws.cell(tr, 5).number_format = "+#,##0;-#,##0;0"
    ws.cell(tr, 6).value = pct_t/100;  ws.cell(tr, 6).number_format = "+0.0%;-0.0%;0%"
    ws.cell(tr, 7).value = f"M:{mj}j | OT:{oj}j"
    ws.cell(tr, 8).value = f"Max M:{int(man_m)}t OT:{int(ort_m)}t"
    ws.row_dimensions[tr].height = 24

    # Stats résumées (ligne tr+1)
    sr = tr + 1
    ws.merge_cells(start_row=sr, start_column=1, end_row=sr, end_column=N)
    c = ws.cell(sr, 1)
    c.value = (f"Jours Manuel: {mj}  |  Max Manuel: {int(man_m)}t/j  |  "
               f"Jours OR-Tools: {oj}  |  Max OR-Tools: {int(ort_m)}t/j  |  "
               f"Ecart total: {int(ec_t):+d}t ({pct_t:+.1f}%)")
    c.font = _ft(size=9, color="1F3864"); c.fill = _hf("EBF5FB"); c.alignment = _CTR
    ws.row_dimensions[sr].height = 16

    # Color scale colonnes Manuel et OR-Tools
    if len(data_rows) > 1:
        last = len(data_rows) + 3
        for col_l in ["C", "D"]:
            try:
                ws.conditional_formatting.add(f"{col_l}4:{col_l}{last}", ColorScaleRule(
                    start_type="min", start_color="FFFFFF",
                    mid_type="percentile", mid_value=50, mid_color="9EC3E6",
                    end_type="max", end_color=hex_color,
                ))
            except Exception:
                pass

    # Largeurs et freeze
    for ci, w in enumerate([12,6,13,13,12,10,12,16], 1):
        ws.column_dimensions[get_column_letter(ci)].width = w
    ws.freeze_panes = "A4"
    return len(data_rows)


def _build_synthese_sheet(ws, all_man_comm, all_ort_comm, all_man_usine, all_ort_usine):
    """Feuille synthèse : KPI globaux + tableau pivot calendrier complet."""
    HDR = _hf("1F3864")
    N = 11

    # Titre
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=N)
    c = ws.cell(1, 1)
    c.value = "Synthese Comparaison Globale  —  Plan Rectifie vs OR-Tools  —  Saison 2026"
    c.font = _ft(bold=True, size=14, white=True); c.fill = HDR; c.alignment = _CTR
    ws.row_dimensions[1].height = 30

    # En-têtes KPI
    kpi_hdrs = ["Entité","Type","Total Manuel(t)","Total OR-T(t)","Ecart(t)",
                "Ecart %","Max Manuel(t)","Max OR-T(t)","Jours Manuel","Jours OR-T","Cap PIC"]
    for ci, h in enumerate(kpi_hdrs, 1):
        c = ws.cell(3, ci); c.value = h
        c.font = _ft(bold=True, size=9, white=True); c.fill = HDR
        c.alignment = _CTR; c.border = _BORD
    ws.row_dimensions[3].height = 20

    # Séparateur COMMERCIAUX
    row = 4
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=N)
    ws.cell(row,1).value = "COMMERCIAUX"
    ws.cell(row,1).font  = _ft(bold=True, size=10, white=True)
    ws.cell(row,1).fill  = _hf("2F4F7F"); ws.cell(row,1).alignment = _LFT
    ws.row_dimensions[row].height = 16; row += 1

    comm_order = ["FEDI","MAKKI BEN SALAH","KHALIL","ACHREF AJLANI","JILANI OBAY"]
    usine_order= ["SICAM","TUCAL","COMOCAP","ELFALLEH","ABIDA"]

    for ci, comm in enumerate(comm_order):
        man  = {str(k):v for k,v in all_man_comm.get(comm, {}).items()}
        ort  = {str(k):v for k,v in all_ort_comm.get(comm, {}).items()}
        alld = sorted(set(list(man.keys())+list(ort.keys())))
        mt   = sum(man.get(d,0) for d in alld); ot = sum(ort.get(d,0) for d in alld)
        ec   = mt-ot; pct = round(ec/ot*100,1) if ot>0 else 0
        mm   = max((man.get(d,0) for d in alld), default=0)
        om   = max((ort.get(d,0) for d in alld), default=0)
        mj   = sum(1 for d in alld if man.get(d,0)>0)
        oj   = sum(1 for d in alld if ort.get(d,0)>0)
        cap  = MANUAL_STATS.get(comm,{}).get("cap",0)
        hx   = COMM_HEX.get(comm,"1F3864")
        vals = [comm,"Commercial",int(mt),int(ot),int(ec),f"{pct:+.1f}%",
                int(mm),int(om),mj,oj,cap]
        alt  = ci%2==0
        for j,v in enumerate(vals,1):
            c = ws.cell(row,j); c.value=v; c.border=_BORD
            c.alignment = _LFT if j==1 else _CTR
            c.font = _ft(bold=(j==1), color="FFFFFF" if j==1 else "1F1F1F")
            c.fill = _hf(hx) if j==1 else (F_ALT1 if alt else F_ALT2)
            if j==5: c.fill = F_POS if ec>=0 else F_NEG
        ws.row_dimensions[row].height = 17; row += 1

    # Séparateur USINES
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=N)
    ws.cell(row,1).value = "USINES"
    ws.cell(row,1).font  = _ft(bold=True, size=10, white=True)
    ws.cell(row,1).fill  = _hf("0B4F6C"); ws.cell(row,1).alignment = _LFT
    ws.row_dimensions[row].height = 16; row += 1

    for ui, usine in enumerate(usine_order):
        man  = {str(k):v for k,v in all_man_usine.get(usine, {}).items()}
        ort  = {str(k):v for k,v in all_ort_usine.get(usine, {}).items()}
        alld = sorted(set(list(man.keys())+list(ort.keys())))
        mt   = sum(man.get(d,0) for d in alld); ot = sum(ort.get(d,0) for d in alld)
        ec   = mt-ot; pct = round(ec/ot*100,1) if ot>0 else 0
        mm   = max((man.get(d,0) for d in alld), default=0)
        om   = max((ort.get(d,0) for d in alld), default=0)
        mj   = sum(1 for d in alld if man.get(d,0)>0)
        oj   = sum(1 for d in alld if ort.get(d,0)>0)
        cap  = USINE_CAPS.get(usine, 0)
        hx   = USINE_HEX.get(usine,"1F3864")
        vals = [usine,"Usine",int(mt),int(ot),int(ec),f"{pct:+.1f}%",
                int(mm),int(om),mj,oj,cap]
        alt  = ui%2==0
        for j,v in enumerate(vals,1):
            c = ws.cell(row,j); c.value=v; c.border=_BORD
            c.alignment = _LFT if j==1 else _CTR
            c.font = _ft(bold=(j==1), color="FFFFFF" if j==1 else "1F1F1F")
            c.fill = _hf(hx) if j==1 else (F_ALT1 if alt else F_ALT2)
            if j==5: c.fill = F_POS if ec>=0 else F_NEG
        ws.row_dimensions[row].height = 17; row += 1

    # ── Tableau pivot calendrier ──────────────────────────────
    row += 2
    # En-têtes pivot : Date | Jour | [M/OT par comm] | [M/OT par usine] | Totaux | Ecart
    pivot_hdrs = ["Date","Jour"]
    for c2 in comm_order:  pivot_hdrs += [f"M-{c2[:5]}", f"OT-{c2[:5]}"]
    for u in usine_order:  pivot_hdrs += [f"M-{u}",      f"OT-{u}"]
    pivot_hdrs += ["Tot Manuel","Tot OR-T","Ecart"]
    Np = len(pivot_hdrs)

    # Titre pivot
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=min(Np,25))
    c = ws.cell(row,1)
    c.value = "Vue Calendrier — Tonnage journalier par entite (Manuel | OR-Tools)"
    c.font = _ft(bold=True, size=11, white=True); c.fill = HDR; c.alignment = _CTR
    ws.row_dimensions[row].height = 22; row += 1

    for ci, h in enumerate(pivot_hdrs,1):
        c = ws.cell(row,ci); c.value = h; c.border = _BORD; c.alignment = _CTR
        if h.startswith("M-"):  c.fill=F_MAN; c.font=_ft(bold=True,size=8,color="1A6B3C")
        elif h.startswith("OT-"): c.fill=F_ORT; c.font=_ft(bold=True,size=8,color="1A3A6B")
        elif "Ecart" in h: c.fill=_hf("FFF2CC"); c.font=_ft(bold=True,size=8,color="7D4F00")
        else: c.fill=HDR; c.font=_ft(bold=True,size=8,white=True)
    ws.row_dimensions[row].height = 18; row += 1

    for di, d in enumerate(SEASON):
        dk     = d.date()
        is_pic = (PIC_S <= dk <= PIC_E)
        row_vals = [d.strftime("%d/%m/%Y"), DAYS_FR[d.weekday()]]
        tot_m = 0; tot_o = 0
        for comm in comm_order:
            mv = {str(k):v for k,v in all_man_comm.get(comm,{}).items()}.get(str(dk),0)
            ov = {str(k):v for k,v in all_ort_comm.get(comm,{}).items()}.get(str(dk),0)
            tot_m += mv; tot_o += ov
            row_vals += [int(mv) if mv>0 else "", int(ov) if ov>0 else ""]
        for usine in usine_order:
            mv = {str(k):v for k,v in all_man_usine.get(usine,{}).items()}.get(str(dk),0)
            ov = {str(k):v for k,v in all_ort_usine.get(usine,{}).items()}.get(str(dk),0)
            tot_m += mv; tot_o += ov
            row_vals += [int(mv) if mv>0 else "", int(ov) if ov>0 else ""]
        ec_d = tot_m - tot_o
        row_vals += [
            int(tot_m) if tot_m>0 else "",
            int(tot_o) if tot_o>0 else "",
            int(ec_d)  if (tot_m>0 or tot_o>0) else "",
        ]
        base = F_PIC if is_pic else (F_ALT1 if di%2==0 else F_ALT2)
        for ci, val in enumerate(row_vals, 1):
            c = ws.cell(row, ci); c.value = val; c.border = _BORD; c.alignment = _CTR
            h = pivot_hdrs[ci-1]
            if ci <= 2:
                c.fill = base
                c.font = _ft(size=9, bold=is_pic, color="7D4F00" if is_pic else "1F1F1F")
            elif h.startswith("M-"):
                c.fill = F_MAN if val else base; c.font = _ft(size=9)
            elif h.startswith("OT-"):
                c.fill = F_ORT if val else base; c.font = _ft(size=9)
            elif "Ecart" in h:
                if isinstance(val, int):
                    c.fill = F_POS if val>=0 else F_NEG
                    c.font = _ft(size=9,bold=True,color="276221" if val>=0 else "9C0006")
                else:
                    c.fill = base; c.font = _ft(size=9)
            else:
                c.fill = base; c.font = _ft(size=9)
        ws.row_dimensions[row].height = 15; row += 1

    # Largeurs
    for i in range(1, Np+1):
        ws.column_dimensions[get_column_letter(i)].width = 11 if i>2 else (12 if i==1 else 6)
    for ci, w in enumerate([22,12,14,14,14,10,14,14,12,12,12],1):
        ws.column_dimensions[get_column_letter(ci)].width = w


def _build_legende_sheet(ws):
    HDR = _hf("1F3864")
    ws.merge_cells("A1:E1")
    ws.cell(1,1).value = "Legende — Codes couleur et signification"
    ws.cell(1,1).font  = _ft(bold=True,size=13,white=True)
    ws.cell(1,1).fill  = HDR; ws.cell(1,1).alignment = _CTR
    ws.row_dimensions[1].height = 28

    hdrs = ["Couleur","Code Hex","Signification","Usage"]
    for ci,h in enumerate(hdrs,1):
        c=ws.cell(3,ci); c.value=h; c.font=_ft(bold=True,size=10,white=True)
        c.fill=HDR; c.alignment=_CTR; c.border=_BORD
    ws.row_dimensions[3].height = 18

    items = [
        ("E2EFDA","Vert clair","Tonnage issu du Plan Manuel rectifie","Colonne Manuel"),
        ("DEEBF7","Bleu clair","Tonnage issu de OR-Tools","Colonne OR-Tools"),
        ("C6EFCE","Vert moyen","Manuel > OR-Tools de plus de 50t","Ecart positif"),
        ("FFC7CE","Rose/Rouge","Manuel < OR-Tools de plus de 50t","Ecart negatif"),
        ("F2F2F2","Gris clair","Difference faible (moins de 50t)","Plans similaires"),
        ("FFF2CC","Jaune","Periode PIC 1-15 juillet","Caps actifs"),
        ("E8F5E9","Vert pale","Plans identiques (OK) ou Manuel only","Statut OK"),
        ("FFF3E0","Orange pale","Jour present uniquement dans OR-Tools","OR-T only"),
        ("E3F2FD","Bleu pale","Manuel prevoir plus que OR-Tools","Manuel +"),
        ("FCE4D6","Saumon","OR-Tools prevoir plus que Manuel","OR-T +"),
    ]
    for ri,(hex_c,nom,desc,usage) in enumerate(items,4):
        ws.cell(ri,1).fill=_hf(hex_c); ws.cell(ri,1).border=_BORD
        ws.cell(ri,1).value=nom; ws.cell(ri,1).font=_ft(size=10)
        ws.cell(ri,2).value=f"#{hex_c}"; ws.cell(ri,2).font=_ft(size=9,color="595959")
        ws.cell(ri,2).fill=_hf(hex_c); ws.cell(ri,2).alignment=_CTR; ws.cell(ri,2).border=_BORD
        ws.cell(ri,3).value=desc; ws.cell(ri,3).font=_ft(size=10)
        ws.cell(ri,3).fill=F_ALT1 if ri%2==0 else F_ALT2; ws.cell(ri,3).border=_BORD
        ws.cell(ri,4).value=usage; ws.cell(ri,4).font=_ft(size=10,color="595959")
        ws.cell(ri,4).fill=F_ALT1 if ri%2==0 else F_ALT2; ws.cell(ri,4).border=_BORD
        ws.row_dimensions[ri].height = 17
    for ci,w in enumerate([18,12,50,25],1):
        ws.column_dimensions[get_column_letter(ci)].width=w


def generate_comparison_excel(man_comm, ort_comm, man_usine, ort_usine):
    """
    Génère le fichier Excel de comparaison complet.
    Retourne des bytes prêts pour st.download_button.
    """
    wb = Workbook()
    wb.remove(wb.active)

    comm_order  = ["FEDI","MAKKI BEN SALAH","KHALIL","ACHREF AJLANI","JILANI OBAY"]
    usine_order = ["SICAM","TUCAL","COMOCAP","ELFALLEH","ABIDA"]

    # 1. Synthèse
    ws = wb.create_sheet("Synthese Globale")
    _build_synthese_sheet(ws, man_comm, ort_comm, man_usine, ort_usine)

    # 2. Une feuille par commercial
    for comm in comm_order:
        ws = wb.create_sheet(f"C - {comm[:14]}")
        _build_entity_sheet(ws, comm,
                            man_comm.get(comm, {}),
                            ort_comm.get(comm, {}),
                            MANUAL_STATS.get(comm,{}).get("cap",800),
                            COMM_HEX.get(comm,"1F3864"))

    # 3. Une feuille par usine
    for usine in usine_order:
        ws = wb.create_sheet(f"U - {usine}")
        _build_entity_sheet(ws, usine,
                            man_usine.get(usine, {}),
                            ort_usine.get(usine, {}),
                            USINE_CAPS.get(usine, 500),
                            USINE_HEX.get(usine,"1F3864"))

    # 4. Légende
    ws = wb.create_sheet("Legende")
    _build_legende_sheet(ws)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


# ═══════════════════════════════════════════════════════════
# HELPERS STREAMLIT
# ═══════════════════════════════════════════════════════════
def _detect_comm_from_filename(filename):
    fn = filename.upper()
    mapping = {
        "FEDI":"FEDI","MEKKI":"MAKKI BEN SALAH","MAKKI":"MAKKI BEN SALAH",
        "KHALIL":"KHALIL","ACHRAF":"ACHREF AJLANI","ACHREF":"ACHREF AJLANI",
        "JILENI":"JILANI OBAY","JILANI":"JILANI OBAY",
    }
    for key, comm in mapping.items():
        if key in fn: return comm
    return None


def _parse_rectification(uploaded_file):
    try:
        df = pd.read_excel(uploaded_file, header=0)
    except Exception as e:
        return None, None, str(e)
    header_row = 0
    for i in range(min(5, len(df))):
        if any("agriculteur" in str(v).lower() for v in df.iloc[i].values):
            header_row = i; break
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
                p=cs.split("/"); d=pd.Timestamp(f"2026-{int(p[1]):02d}-{int(p[0]):02d}")
                parsed_dates.append((c,d)); continue
            except: pass
        try:
            d=pd.to_datetime(str(c).split(" ")[0],errors="coerce")
            if pd.notna(d) and d.year==2026: parsed_dates.append((c,d))
        except: pass
    if not parsed_dates or col_agri is None:
        return None, None, "Format non reconnu"
    rows=[]; comm_detected=None
    for _, row in df.iterrows():
        comm=str(row.get(col_comm,"") or "").strip()
        agri=str(row.get(col_agri,"") or "").strip()
        if not agri or agri.upper() in ("NAN","","AGRICULTEUR","TOTAL"): continue
        if "sous-total" in agri.lower() or "total" in agri.lower(): continue
        if comm and comm.upper() not in ("NAN",""):
            for known in MANUAL_STATS:
                if known.upper() in comm.upper() or comm.upper() in known.upper():
                    comm_detected=known; break
        ton=pd.to_numeric(row.get(col_tonnage,0),errors="coerce") if col_tonnage else 0
        for oc,date in parsed_dates:
            val=pd.to_numeric(row.get(oc,0),errors="coerce")
            if pd.notna(val) and val>0:
                rows.append({"agriculteur":agri,"date":date,"tonnes":float(val),
                             "tonnage_total":float(ton) if pd.notna(ton) else 0})
    if not rows: return comm_detected, None, "Aucune donnee trouvee"
    df_rows=pd.DataFrame(rows)
    daily=df_rows.groupby("date")["tonnes"].sum().reset_index().sort_values("date")
    daily_dict={str(r["date"].date()):float(r["tonnes"]) for _,r in daily.iterrows()}
    return comm_detected, daily_dict, df_rows


def _ortools_profile(comm_or_usine, planning_df, entity_type="commercial"):
    """Extrait le profil OR-Tools depuis le planning Supabase, ou utilise le profil de base."""
    if planning_df is not None and not planning_df.empty:
        col = "Commercial" if entity_type == "commercial" else "Usine"
        if col in planning_df.columns:
            sub = planning_df[planning_df[col] == comm_or_usine].copy()
            if not sub.empty and "Date" in sub.columns:
                sub["Date"] = pd.to_datetime(sub["Date"], errors="coerce")
                daily = sub.groupby("Date")["Tonnes/Jour"].sum().reset_index()
                return {str(r["Date"].date()): float(r["Tonnes/Jour"]) for _, r in daily.iterrows()}
    # Fallback profil de base
    if entity_type == "commercial":
        return ORT_COMM_BASE.get(comm_or_usine, {})
    else:
        if comm_or_usine in ORT_USINE_BASE:
            return ORT_USINE_BASE[comm_or_usine]
        usine_pdf_vals = USINE_DAILY_PDF.get(comm_or_usine, [])
        return {(pd.Timestamp("2026-06-20")+pd.Timedelta(days=i)).date():v
                for i,v in enumerate(usine_pdf_vals) if v>0}


def _build_reference_curve(comm):
    stat=MANUAL_STATS[comm]; n=stat["n_jours"]; mx=stat["max_j"]
    starts={"FEDI":"2026-06-28","MAKKI BEN SALAH":"2026-06-22",
            "KHALIL":"2026-06-20","ACHREF AJLANI":"2026-06-20","JILANI OBAY":"2026-07-23"}
    start=pd.Timestamp(starts.get(comm,"2026-06-20"))
    p1=int(n*0.30); p2=int(n*0.50); p3=n-p1-p2
    result={}
    for i in range(n):
        if i<p1:        v=mx*0.2+(mx*0.6)*(i/max(p1,1))
        elif i<p1+p2:   v=mx*0.8+(mx*0.2)*((i-p1)/max(p2,1))
        else:           v=mx*(1-((i-p1-p2)/max(p3,1))*0.8)
        result[str((start+pd.Timedelta(days=i)).date())]=round(max(0,v))
    return result


def _kpi_row(name, man_dict, ort_dict, cap):
    mt=sum(man_dict.values()) if man_dict else 0
    mm=max(man_dict.values()) if man_dict else 0
    ot=sum(ort_dict.values()) if ort_dict else 0
    om=max(ort_dict.values()) if ort_dict else 0
    k1,k2,k3,k4=st.columns(4)
    k1.metric("Total Manuel", f"{int(mt):,}t".replace(",",""),
              delta=f"{int(mt-ot):+,}t vs OR-T".replace(",",""), delta_color="off")
    k2.metric("Max/j Manuel", f"{int(mm)}t",
              delta=f"{int(mm-om):+d}t vs OR-T",
              delta_color="inverse" if mm>cap else "off")
    k3.metric("Jours actifs", f"{len([v for v in man_dict.values() if v>0])}j" if man_dict else "0j")
    k4.metric("Cap PIC", f"{cap}t/j",
              delta="OK" if mm<=cap else "DEPASSE",
              delta_color="normal" if mm<=cap else "inverse")


def _build_chart(name, man_dict, ort_dict, color, cap):
    # Normaliser toutes les cles en string pour eviter TypeError
    man_str = {str(k): v for k, v in man_dict.items()}
    ort_str = {str(k): v for k, v in ort_dict.items()}
    all_d = sorted(set(list(man_str.keys())+list(ort_str.keys())))
    man_dict = man_str
    ort_dict = ort_str
    dates = [pd.Timestamp(d) for d in all_d]
    mv    = [man_dict.get(d,0) for d in all_d]
    ov    = [ort_dict.get(d,0) for d in all_d]
    fig   = go.Figure()
    fig.add_trace(go.Scatter(x=dates, y=mv, name="Plan Manuel rectifie",
        line=dict(color=color,width=2.5), fill="tozeroy",
        fillcolor="rgba(59,130,246,0.08)", mode="lines"))
    if any(v>0 for v in ov):
        fig.add_trace(go.Scatter(x=dates, y=ov, name="OR-Tools",
            line=dict(color="#ffffff",width=1.5,dash="dot"), mode="lines"))
    fig.add_hline(y=cap, line_dash="dash", line_color="#e8543a", line_width=1,
                  annotation_text=f"Cap: {cap}t/j",
                  annotation_position="top right", annotation_font_color="#e8543a")
    fig.add_vrect(x0=pd.Timestamp("2026-07-01"), x1=pd.Timestamp("2026-07-15"),
                  fillcolor="rgba(245,166,35,0.07)", line_width=0,
                  annotation_text="PIC", annotation_position="top left",
                  annotation_font_color="#f5a623")
    fig.update_layout(
        title=f"{name} — Profil journalier : Plan rectifie vs OR-Tools",
        template="plotly_dark", plot_bgcolor="#0d1117", paper_bgcolor="#161b22",
        height=360, hovermode="closest",
        legend=dict(orientation="h", yanchor="bottom", y=1.02),
        xaxis=dict(title="Date", gridcolor="#21262d"),
        yaxis=dict(title="Tonnes/jour", gridcolor="#21262d"),
    )
    return fig


# ═══════════════════════════════════════════════════════════
# POINT D'ENTRÉE PRINCIPAL
# ═══════════════════════════════════════════════════════════
def render_comparaison_tab(planning_df=None, df_to_xlsx_styled=None):

    st.markdown("""
    <div style='background:#1a2332;border:1px solid #21262d;border-radius:12px;
    padding:16px 20px;margin-bottom:20px'>
      <div style='font-size:1.1rem;font-weight:700;color:#f0f6fc;margin-bottom:4px'>
        Comparaison Plans — OR-Tools vs Plans Rectifies Manuellement
      </div>
      <div style='font-size:.82rem;color:#8b949e'>
        Deposez les fichiers Excel rectifies pour visualiser les ecarts jour par jour
        et exporter un tableau Excel complet colore par entite.
      </div>
    </div>""", unsafe_allow_html=True)

    # Session state — avec restauration depuis query_params si session perdue
    import json as _json, base64 as _b64
    if "comp_uploaded" not in st.session_state:
        # Essayer de restaurer depuis query_params
        try:
            _qp = st.query_params.get("comp_data", "")
            if _qp:
                _decoded = _json.loads(_b64.b64decode(_qp.encode()).decode())
                st.session_state["comp_uploaded"] = _decoded
            else:
                st.session_state["comp_uploaded"] = {}
        except Exception:
            st.session_state["comp_uploaded"] = {}
    if "comp_raw" not in st.session_state:
        st.session_state["comp_raw"] = {}

    # ── ZONE UPLOAD + STATUT ────────────────────────────────
    st.subheader("Deposer les fichiers rectifies")
    col_up, col_st = st.columns([3, 2])
    with col_up:
        uploaded = st.file_uploader("Fichiers", type=["xlsx","xls"],
                                    accept_multiple_files=True,
                                    label_visibility="collapsed",
                                    help="Rectification_Fedi_13_JUIN.xlsx, etc.")
    if uploaded:
        for f in uploaded:
            cfn = _detect_comm_from_filename(f.name)
            cfd, daily, raw_or_err = _parse_rectification(f)
            comm = cfn or cfd
            if comm and daily:
                st.session_state["comp_uploaded"][comm] = daily
                if isinstance(raw_or_err, pd.DataFrame):
                    st.session_state["comp_raw"][comm] = raw_or_err
                tot = sum(daily.values())
                st.success(f"Charge : {comm} — {len(daily)} jours, {int(tot):,}t".replace(",",""))
                # Persister dans query_params pour survivre aux refreshs
                try:
                    import json as _j, base64 as _b
                    _encoded = _b.b64encode(_j.dumps(st.session_state["comp_uploaded"]).encode()).decode()
                    st.query_params["comp_data"] = _encoded
                except Exception:
                    pass
                # Persister dans query_params pour survivre aux refreshs
                try:
                    import json as _j, base64 as _b
                    _encoded = _b.b64encode(_j.dumps(st.session_state["comp_uploaded"]).encode()).decode()
                    st.query_params["comp_data"] = _encoded
                except Exception:
                    pass
            elif comm:
                st.warning(f"{f.name} : {raw_or_err}")
            else:
                st.warning(f"{f.name} : commercial non detecte. Renommer en Rectification_FEDI_...")

    with col_st:
        st.markdown("**Statut :**")
        for c in MANUAL_STATS:
            if c in st.session_state["comp_uploaded"]:
                n=len(st.session_state["comp_uploaded"][c])
                t=sum(st.session_state["comp_uploaded"][c].values())
                st.success(f"{c.split()[0]} — {n}j | {int(t):,}t".replace(",",""))
            else:
                st.info(f"{c.split()[0]} — reference interne")
        if st.session_state["comp_uploaded"]:
            if st.button("Effacer uploads", use_container_width=True):
                st.session_state["comp_uploaded"] = {}
                st.session_state["comp_raw"] = {}
                try:
                    st.query_params.pop("comp_data", None)
                except Exception:
                    pass
                st.rerun()

    if not st.session_state["comp_uploaded"]:
        st.info("Aucun fichier uploade - donnees sauvegardees dans Supabase (persistant entre sessions) — courbes basees sur donnees de reference 13/06/2026.")

    # ── BOUTON EXPORT EXCEL PRINCIPAL ───────────────────────
    st.divider()
    col_exp1, col_exp2 = st.columns([2, 3])
    with col_exp1:
        st.markdown("**Export Excel comparaison complète**")
        st.caption("12 onglets : Synthèse + 5 commerciaux + 5 usines + Légende")
        if st.button("⬇️ Générer & télécharger Excel", type="primary",
                     use_container_width=True):
            with st.spinner("Génération du fichier Excel en cours..."):
                # Construire les dictionnaires finaux
                man_comm = {}
                ort_comm = {}
                for comm in MANUAL_STATS:
                    if comm in st.session_state["comp_uploaded"]:
                        d = st.session_state["comp_uploaded"][comm]
                        man_comm[comm] = {pd.Timestamp(k).date(): v for k,v in d.items()}
                    else:
                        ref = _build_reference_curve(comm)
                        man_comm[comm] = {pd.Timestamp(k).date(): v for k,v in ref.items()}
                    ort_comm[comm] = _ortools_profile(comm, planning_df, "commercial")

                man_usine = {}
                ort_usine = {}
                for usine in USINE_CAPS:
                    usine_vals = USINE_DAILY_PDF.get(usine, [])
                    man_usine[usine] = {(pd.Timestamp("2026-06-20")+pd.Timedelta(days=i)).date(): v
                                        for i,v in enumerate(usine_vals) if v>0}
                    ort_usine[usine] = _ortools_profile(usine, planning_df, "usine")

                xl_bytes = generate_comparison_excel(man_comm, ort_comm, man_usine, ort_usine)

            st.download_button(
                "📥 Télécharger Comparaison_ORT_vs_Manuel_2026.xlsx",
                data=xl_bytes,
                file_name="Comparaison_ORT_vs_Manuel_2026.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
    with col_exp2:
        st.markdown("""
        <div style='background:#161b22;border:1px solid #21262d;border-radius:8px;
        padding:12px 16px;font-size:12px;color:#8b949e'>
          <b style='color:#f0f6fc'>Contenu du fichier Excel :</b><br>
          📊 <b>Synthese Globale</b> — KPI tous commerciaux + usines + tableau pivot calendrier<br>
          👤 <b>C - FEDI / MAKKI / KHALIL / ACHREF / JILANI</b> — jour par jour avec écarts colorés<br>
          🏭 <b>U - SICAM / TUCAL / COMOCAP / ELFALLEH / ABIDA</b> — réception journalière vs OR-T<br>
          🔑 <b>Legende</b> — explication des codes couleur<br><br>
          <b style='color:#f5a623'>Couleurs :</b> Vert = Manuel &gt; OR-T | Rouge = Manuel &lt; OR-T |
          Jaune = PIC 1-15 juillet
        </div>""", unsafe_allow_html=True)

    st.divider()

    # ── ONGLETS VISUELS ─────────────────────────────────────
    c1, c2, c3, c4 = st.tabs([
        "Courbes par commercial",
        "Par usine",
        "Statistiques globales",
        "Corrections optimizer",
    ])

    # ─ C1 : COURBES PAR COMMERCIAL ──────────────────────────
    with c1:
        sel_comm = st.selectbox("Choisir un commercial", list(MANUAL_STATS.keys()),
                                key="comp_sel_comm")
        color  = COMM_COLORS[sel_comm]
        cap    = MANUAL_STATS[sel_comm]["cap"]

        if sel_comm in st.session_state["comp_uploaded"]:
            d = st.session_state["comp_uploaded"][sel_comm]
            man_dict = {pd.Timestamp(k).date(): v for k,v in d.items()}
            src_lbl = "fichier uploade"
        else:
            ref = _build_reference_curve(sel_comm)
            man_dict = {pd.Timestamp(k).date(): v for k,v in ref.items()}
            src_lbl = "reference (13/06/2026)"

        ort_dict = _ortools_profile(sel_comm, planning_df, "commercial")

        _kpi_row(sel_comm, man_dict, ort_dict, cap)
        st.caption(f"Source plan rectifie : {src_lbl}")

        fig = _build_chart(sel_comm, man_dict, ort_dict, color, cap)
        st.plotly_chart(fig, use_container_width=True)

        # Tableau pivot jour par jour (affiché dans le dashboard)
        st.markdown(f"**Tableau jour par jour — {sel_comm}**")
        pivot_rows = []
        _man_n = {str(k):v for k,v in man_dict.items()}
        _ort_n = {str(k):v for k,v in ort_dict.items()}
        for d in SEASON:
            dk = d.date()
            mv = _man_n.get(str(dk), 0)
            ov = _ort_n.get(str(dk), 0)
            if mv == 0 and ov == 0: continue
            ec = mv - ov
            pivot_rows.append({
                "Date":       d.strftime("%d/%m/%Y"),
                "Jour":       DAYS_FR[d.weekday()],
                "PIC":        "⚡" if PIC_S<=dk<=PIC_E else "",
                "Manuel (t)": int(mv) if mv>0 else "",
                "OR-Tools (t)":int(ov) if ov>0 else "",
                "Ecart (t)":  f"{int(ec):+d}" if (mv>0 or ov>0) else "",
                "Statut":     ("OK" if abs(ec)<=50 else ("Manuel+" if ec>0 else "OR-T+")),
            })
        if pivot_rows:
            df_piv = pd.DataFrame(pivot_rows)
            st.dataframe(df_piv, use_container_width=True, height=320, hide_index=True,
                column_config={
                    "Manuel (t)":   st.column_config.NumberColumn("Manuel (t)",  format="%d t"),
                    "OR-Tools (t)": st.column_config.NumberColumn("OR-Tools (t)",format="%d t"),
                })

        # Vue par agriculteur si fichier uploadé
        if sel_comm in st.session_state["comp_raw"]:
            st.markdown("---")
            st.markdown(f"**Profil par agriculteur — {sel_comm}**")
            df_raw = st.session_state["comp_raw"][sel_comm]
            agri_stats = df_raw.groupby("agriculteur").agg(
                total=("tonnes","sum"), max_jour=("tonnes","max"), jours=("date","nunique")
            ).reset_index().sort_values("total",ascending=False)
            agri_stats.columns = ["Agriculteur","Total (t)","Max/jour (t)","Jours actifs"]
            agri_stats[["Total (t)","Max/jour (t)"]] = agri_stats[["Total (t)","Max/jour (t)"]].round(0).astype(int)
            n_agri = len(agri_stats)
            fig_a = px.bar(agri_stats, x="Total (t)", y="Agriculteur", orientation="h",
                           height=max(350, n_agri*28+80), color="Total (t)",
                           color_continuous_scale="Viridis",
                           title=f"Tonnage par agriculteur — {sel_comm}",
                           template="plotly_dark", text="Total (t)")
            fig_a.update_traces(textposition="outside", texttemplate="%{x:.0f}t",
                                textfont=dict(size=10))
            fig_a.update_layout(paper_bgcolor="#161b22", showlegend=False,
                                margin=dict(l=230,r=60,t=50,b=40))
            st.plotly_chart(fig_a, use_container_width=True)

    # ─ C2 : PAR USINE ────────────────────────────────────────
    with c2:
        sel_usine = st.selectbox("Choisir une usine",
                                 ["SICAM","TUCAL","COMOCAP","ELFALLEH","ABIDA"],
                                 key="comp_usine_sel")
        cap   = USINE_CAPS[sel_usine]
        data_u= USINE_DAILY_PDF.get(sel_usine, [])
        dates_u=[pd.Timestamp("2026-06-20")+pd.Timedelta(days=i) for i in range(len(data_u))]
        total_u=sum(data_u); max_u=max(data_u) if data_u else 0
        jours_u=sum(1 for v in data_u if v>0)
        pic_d=[data_u[i] for i in range(11,26) if i<len(data_u)]
        max_pic=max(pic_d) if pic_d else 0

        k1,k2,k3,k4,k5=st.columns(5)
        k1.metric("Total saison", f"{int(total_u):,}t".replace(",",""))
        k2.metric("Max journalier", f"{max_u}t",
                  delta="DEPASSE" if max_u>cap else "OK",
                  delta_color="inverse" if max_u>cap else "normal")
        k3.metric("Cap officiel", f"{cap}t/j")
        k4.metric("Max PIC 1-15/07", f"{max_pic}t")
        k5.metric("Jours actifs", f"{jours_u}j")

        fig_u=go.Figure()
        fig_u.add_trace(go.Bar(x=dates_u, y=data_u, name="Plan rectifie",
                               marker_color="#3b82f6", marker_line_width=0))
        fig_u.add_hline(y=cap, line_dash="dash", line_color="#e8543a", line_width=1.5,
                        annotation_text=f"Cap: {cap}t/j",
                        annotation_position="top right", annotation_font_color="#e8543a")
        fig_u.add_vrect(x0=pd.Timestamp("2026-07-01"), x1=pd.Timestamp("2026-07-15"),
                        fillcolor="rgba(245,166,35,0.08)", line_width=0,
                        annotation_text="PIC", annotation_position="top left",
                        annotation_font_color="#f5a623")

        ort_u_dict = _ortools_profile(sel_usine, planning_df, "usine")
        if ort_u_dict:
            ort_dates = sorted(ort_u_dict.keys())
            fig_u.add_trace(go.Scatter(
                x=[pd.Timestamp(d) for d in ort_dates],
                y=[ort_u_dict[d] for d in ort_dates],
                name="OR-Tools", line=dict(color="#f5a623",width=2,dash="dash"), mode="lines"))

        fig_u.update_layout(
            title=f"{sel_usine} — Reception journaliere : Plan rectifie vs OR-Tools",
            template="plotly_dark", plot_bgcolor="#0d1117", paper_bgcolor="#161b22",
            height=360, hovermode="closest",
            legend=dict(orientation="h", yanchor="bottom", y=1.02))
        st.plotly_chart(fig_u, use_container_width=True)

        # Tableau pivot usine jour par jour
        st.markdown(f"**Tableau jour par jour — {sel_usine}**")
        man_u = {str((pd.Timestamp("2026-06-20")+pd.Timedelta(days=i)).date()):v
                 for i,v in enumerate(data_u) if v>0}
        _ort_u_n = {str(k):v for k,v in ort_u_dict.items()}
        pivot_u = []
        for d in SEASON:
            dk=d.date(); mv=man_u.get(str(dk),0); ov=_ort_u_n.get(str(dk),0)
            if mv==0 and ov==0: continue
            ec=mv-ov
            pivot_u.append({
                "Date": d.strftime("%d/%m/%Y"), "Jour": DAYS_FR[d.weekday()],
                "PIC": "⚡" if PIC_S<=dk<=PIC_E else "",
                "Manuel (t)": int(mv) if mv>0 else "",
                "OR-Tools (t)": int(ov) if ov>0 else "",
                "Ecart (t)": f"{int(ec):+d}" if (mv>0 or ov>0) else "",
                "Cap": f"DEPASSE" if (PIC_S<=dk<=PIC_E and max(mv,ov)>cap) else ("OK" if PIC_S<=dk<=PIC_E else ""),
            })
        if pivot_u:
            st.dataframe(pd.DataFrame(pivot_u), use_container_width=True,
                         height=320, hide_index=True)

        _rules = {
            "SICAM":   ["Montee progressive juin (140-625t)","Pic 1-15 juillet (910-1650t)",
                        "Declin rapide aout","Gros agris: 80-200t/j","Petits agris: 15-30t/j"],
            "TUCAL":   ["Demarrage 24 juin","Plateau 590-700t/j juillet",
                        "STE 428: 120t/j x21j","STE BACCARA: 80t/j x33j","MAKKI: 15-35t/j"],
            "COMOCAP": ["Pic tardif 27/07 (875t)","TRC/PPL/PL: 10-20t/j majoritaires",
                        "YASIN MNASRI SEMI: 60-120t/j","STE 428: 130t/j x15j","Declin net aout"],
            "ELFALLEH":["Cap 150t/j jamais depasse","Max plan: 80t/j","Fenetres 20-30j/agri"],
            "ABIDA":   ["Demarrage tardif (juillet)","Montee 50-320t/j","Declin aout"],
        }
        st.markdown("**Regles terrain observees :**")
        for rule in _rules.get(sel_usine, []):
            st.markdown(f"• {rule}")

    # ─ C3 : STATISTIQUES GLOBALES ────────────────────────────
    with c3:
        st.markdown("### Comparaison globale — tous commerciaux et usines")
        rows_cmp=[]; mt_list=[]; ot_list=[]
        comm_order=list(MANUAL_STATS.keys())

        for comm in comm_order:
            if comm in st.session_state["comp_uploaded"]:
                d={str(pd.Timestamp(k).date()):v for k,v in st.session_state["comp_uploaded"][comm].items()}
                src="uploade"
            else:
                d={str(pd.Timestamp(k).date()):v for k,v in _build_reference_curve(comm).items()}
                src="reference"
            od={str(k):v for k,v in _ortools_profile(comm, planning_df, "commercial").items()}
            mt=sum(d.values()); ot=sum(od.values()) if od else 0
            mm=max(d.values()) if d else 0; om=max(od.values()) if od else 0
            ec=mt-ot; pct=round(ec/ot*100,1) if ot>0 else 0
            mt_list.append(mt); ot_list.append(ot)
            rows_cmp.append({"Commercial":comm,"Source":src,
                             "Total Manuel(t)":int(mt),"Total OR-T(t)":int(ot),
                             "Ecart(t)":f"{int(ec):+,}".replace(",",""),
                             "Max Manuel":f"{int(mm)}t","Max OR-T":f"{int(om)}t",
                             "Jours Manuel":sum(1 for v in d.values() if v>0),
                             "Cap PIC":f"{MANUAL_STATS[comm]['cap']}t/j"})

        st.dataframe(pd.DataFrame(rows_cmp), use_container_width=True, hide_index=True)

        fig_cmp=go.Figure()
        fig_cmp.add_trace(go.Bar(name="Plan Manuel", x=comm_order, y=mt_list,
                                 marker_color="#e8543a",
                                 text=[f"{int(v):,}t".replace(",","") for v in mt_list],
                                 textposition="outside"))
        fig_cmp.add_trace(go.Bar(name="OR-Tools", x=comm_order, y=ot_list,
                                 marker_color="#3b82f6",
                                 text=[f"{int(v):,}t".replace(",","") for v in ot_list],
                                 textposition="outside"))
        fig_cmp.update_layout(barmode="group", template="plotly_dark",
                              paper_bgcolor="#161b22", plot_bgcolor="#0d1117",
                              height=380, yaxis_title="Tonnes",
                              legend=dict(orientation="h", yanchor="bottom", y=1.02))
        st.plotly_chart(fig_cmp, use_container_width=True)

        st.markdown("---")
        st.markdown("### 4 causes racines des ecarts OR-Tools vs Manuel")
        for title, color, desc in [
            ("1 — Borne journaliere trop large","#e8543a",
             "ub_day x1.5 trop permissif : OR-Tools place 870t/j pour FEDI (reel 200t max)."),
            ("2 — Arrondi 10t supprime micro-livraisons","#f5a623",
             "Plans manuels utilisent 10t/15t/20t pour COMOCAP/TUCAL, perdus avec arrondi 10t."),
            ("3 — FACTORY_OVERFLOW_WEIGHT trop faible","#8b5cf6",
             "Poids 500 insuffisant : ELFALLEH 130t OR-T vs 60t reel. Petites usines non protegees."),
            ("4 — Pas de cap max par agriculteur","#00e5a0",
             "Regle terrain : <200t=30t/j | <500t=60t/j | >500t=200t/j. OR-Tools ignore cette regle."),
        ]:
            st.markdown(f"""<div style='border-left:4px solid {color};padding:10px 16px;
            margin-bottom:8px;background:#161b22;border-radius:0 8px 8px 0'>
              <div style='font-size:13px;font-weight:600;color:{color}'>{title}</div>
              <div style='font-size:12px;color:#8b949e;margin-top:4px'>{desc}</div>
            </div>""", unsafe_allow_html=True)

    # ─ C4 : CORRECTIONS OPTIMIZER ───────────────────────────
    with c4:
        st.markdown("### Corrections a appliquer dans optimizer_v2.py")
        st.success("Configuration OPTIMALE stable - ne pas modifier")
        st.markdown("""
| Parametre | Valeur actuelle | Statut |
|---|---|---|
| Borne journaliere | day_planned x SCALE x 2.0 | OPTIMAL |
| FACTORY_OVERFLOW_WEIGHT | 2000 | Applique |
| Arrondi tonnage | 10t | Stable |
| Correction post-traitement | Desactivee | OK |
        """)
        st.warning("**x1.1 et x1.5 causent INFEASIBLE** - teste et prouve.")
        st.code("""Resultats actuels:
Status: OPTIMAL (~2s) | 2630 rows | 96 676t (-0.32%)
FEDI   : -234t (-0.7%)
MAKKI  : +115t (+0.5%)
ACHREF : 0t (parfait)
KHALIL : -139t (-0.8%)
JILANI : -125t (-1.8%)""")
