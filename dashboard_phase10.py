# ============================================================
# DASHBOARD PHASE 10 — Tomate Planning 2026
#
# FONCTIONNALITES :
#   - Login par rôle (directeur / commercial / usine)
#   - Données depuis Supabase + fallback Excel local
#   - OR-Tools optimizer_v2.py (distances + caps)
#   - Upload Excel par commercial (upload_tab.py)
#   - Gestion agriculteurs directement dans le dashboard
#   - Historique 2025 vs plan 2026
#
# FICHIERS REQUIS dans le même dossier :
#   optimizer_v2.py  migrate.py  upload_tab.py
#   Planning_Tomate_2026.xlsx
#   Recap_tonnage_pre_vu_ajuste__mai26.xlsx
#
# LANCEMENT :
#   streamlit run dashboard_phase10.py
# ============================================================

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import datetime, os, subprocess, sys, io, zipfile, hashlib

# Import upload system (upload_tab.py must be in the same folder)
try:
    from upload_tab import render_upload_tab, generate_template_excel
    UPLOAD_AVAILABLE = True
except ImportError:
    UPLOAD_AVAILABLE = False

# set_page_config MUST be the very first Streamlit call
st.set_page_config(
    page_title="🍅 Tomate Planning 2026",
    layout="wide",
    initial_sidebar_state="expanded",
    page_icon="🍅"
)

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
# ============================================================
# LOGIN SYSTEM
# ============================================================

# Rôles disponibles :
#   directeur  → voit tout, peut régénérer le planning
#   commercial → voit seulement ses agriculteurs et ses données
#   usine      → voit seulement les livraisons à son usine

def hash_password(password: str) -> str:
    """Hash simple SHA-256 pour ne pas stocker les mots de passe en clair."""
    return hashlib.sha256(password.encode()).hexdigest()

def load_users() -> dict:
    """
    Charge les utilisateurs depuis st.secrets (Streamlit Cloud)
    ou depuis un dict local si secrets non disponibles (développement local).

    Format dans .streamlit/secrets.toml :
    [users.directeur]
    password = "hash_sha256_ici"
    role     = "directeur"
    name     = "Directeur Général"
    filter   = ""

    [users.fedi]
    password = "hash_sha256_ici"
    role     = "commercial"
    name     = "FEDI"
    filter   = "FEDI"

    [users.comocap]
    password = "hash_sha256_ici"
    role     = "usine"
    name     = "COMOCAP"
    filter   = "COMOCAP"
    """
    try:
        # Production : lire depuis secrets.toml
        users = {}
        for username, info in st.secrets.get("users", {}).items():
            users[username] = {
                "password": info["password"],
                "role":     info["role"],
                "name":     info["name"],
                "filter":   info.get("filter", ""),
            }
        if users:
            return users
    except Exception:
        pass

    # Fallback local pour développement (mots de passe en clair hashés)
    # CHANGE CES MOTS DE PASSE avant de mettre en production !
    return {
        "directeur": {
            "password": hash_password("admin2026"),
            "role":     "directeur",
            "name":     "Directeur Général",
            "filter":   "",
        },
        "fedi": {
            "password": hash_password("fedi2026"),
            "role":     "commercial",
            "name":     "FEDI",
            "filter":   "FEDI",
        },
        "makki": {
            "password": hash_password("makki2026"),
            "role":     "commercial",
            "name":     "MAKKI BEN SALAH",
            "filter":   "MAKKI BEN SALAH",
        },
        "khalil": {
            "password": hash_password("khalil2026"),
            "role":     "commercial",
            "name":     "KHALIL",
            "filter":   "KHALIL",
        },
        "achref": {
            "password": hash_password("achref2026"),
            "role":     "commercial",
            "name":     "ACHREF AJLANI",
            "filter":   "ACHREF AJLANI",
        },
        "jilani": {
            "password": hash_password("jilani2026"),
            "role":     "commercial",
            "name":     "JILANI OBAY",
            "filter":   "JILANI OBAY",
        },
        "comocap": {
            "password": hash_password("comocap2026"),
            "role":     "usine",
            "name":     "COMOCAP",
            "filter":   "COMOCAP",
        },
        "sicam": {
            "password": hash_password("sicam2026"),
            "role":     "usine",
            "name":     "SICAM",
            "filter":   "SICAM",
        },
        "tucal": {
            "password": hash_password("tucal2026"),
            "role":     "usine",
            "name":     "TUCAL",
            "filter":   "TUCAL",
        },
    }

def show_login_page():
    """Affiche la page de login — appelée si non connecté."""
    st.markdown("""
    <style>
    [data-testid="stAppViewContainer"]{background:#080b12}
    [data-testid="stSidebar"]{display:none}
    .login-wrap{max-width:420px;margin:80px auto 0;padding:0 16px}
    .login-logo{text-align:center;margin-bottom:32px}
    .login-logo h1{font-size:2rem;color:#f0f6fc;margin-top:12px}
    .login-logo p{color:#8b949e;font-size:.85rem}
    .login-card{background:#161b22;border:1px solid #21262d;border-radius:16px;padding:32px}
    .login-card h2{color:#f0f6fc;font-size:1.1rem;margin-bottom:24px;font-weight:600}
    </style>
    <div class="login-wrap">
      <div class="login-logo">
        <div style="font-size:3rem">🍅</div>
        <h1>Tomate Planning</h1>
        <p>Système de planification transport & récolte 2026</p>
      </div>
    </div>
    """, unsafe_allow_html=True)

    col_l, col_c, col_r = st.columns([1, 2, 1])
    with col_c:
        with st.container(border=True):
            st.markdown("#### 🔐 Connexion")
            username = st.text_input("Identifiant",
                                     placeholder="Entrez votre identifiant",
                                     label_visibility="visible")
            password = st.text_input("Mot de passe", type="password",
                                     placeholder="Entrez votre mot de passe")
            login_btn = st.button("Se connecter", use_container_width=True,
                                  type="primary")

            if login_btn:
                users = load_users()
                u = username.strip().lower()
                if u in users and users[u]["password"] == hash_password(password):
                    st.session_state["logged_in"]  = True
                    st.session_state["username"]   = u
                    st.session_state["role"]       = users[u]["role"]
                    st.session_state["name"]       = users[u]["name"]
                    st.session_state["filter"]     = users[u]["filter"]
                    # Sauvegarder dans l'URL pour survivre aux refreshs
                    try:
                        st.query_params["u"] = u
                        st.query_params["t"] = _make_token(u)
                    except Exception:
                        pass
                    st.rerun()
                else:
                    st.error("❌ Identifiant ou mot de passe incorrect.")

# ── Timeout de session : déconnexion après 30 min d'inactivité ─
SESSION_TIMEOUT_MIN = 30

def _check_session_timeout():
    """Déconnecte automatiquement après SESSION_TIMEOUT_MIN minutes."""
    import time
    now = time.time()
    last = st.session_state.get("last_activity", now)
    if now - last > SESSION_TIMEOUT_MIN * 60:
        for key in list(st.session_state.keys()):
            del st.session_state[key]
        try:
            st.query_params.clear()
        except Exception:
            pass
        st.warning("⏱️ Session expirée après inactivité. Reconnectez-vous.")
        st.stop()
    st.session_state["last_activity"] = now

# ── Helpers tokens URL pour persister la session entre refreshs ─
def _make_token(username):
    """Génère un token simple à partir du username + secret."""
    import hashlib
    secret = "tomate2026_secret_seed"
    return hashlib.sha256(f"{username}|{secret}".encode()).hexdigest()[:32]

def _verify_token(username, token):
    """Vérifie qu'un token correspond au username."""
    return token == _make_token(username)

# ── Check login state — vérifier d'abord les query params (refresh) ─
if "logged_in" not in st.session_state:
    st.session_state["logged_in"] = False

# Si pas connecté mais URL contient un token valide → reconnecter auto
if not st.session_state["logged_in"]:
    try:
        qp = st.query_params
        url_user  = qp.get("u", "")
        url_token = qp.get("t", "")
        if url_user and url_token:
            users = load_users()
            u = url_user.strip().lower()
            if u in users and _verify_token(u, url_token):
                # Restaurer la session depuis l'URL
                st.session_state["logged_in"] = True
                st.session_state["username"]  = u
                st.session_state["role"]      = users[u]["role"]
                st.session_state["name"]      = users[u]["name"]
                st.session_state["filter"]    = users[u]["filter"]
    except Exception:
        pass

if not st.session_state["logged_in"]:
    show_login_page()
    st.stop()   # ← Stop here if not logged in. Nothing below executes.

# ── At this point: user is logged in ─────────────────────────
_check_session_timeout()  # ← Vérifier timeout d'inactivité
CURRENT_USER   = st.session_state["username"]
CURRENT_ROLE   = st.session_state["role"]
CURRENT_NAME   = st.session_state["name"]
CURRENT_FILTER = st.session_state["filter"]

# ============================================================
# PAGE CONFIG — must be first Streamlit call
# ============================================================

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
  .stDownloadButton > button {
    background:#161b22 !important; border:1px solid #21262d !important;
    color:#f0f6fc !important; font-size:.8rem !important;
  }
  .stDownloadButton > button:hover {
    border-color:#3b82f6 !important; color:#3b82f6 !important;
  }
</style>
""", unsafe_allow_html=True)

# ── Supabase connection ──────────────────────────────────────
from supabase import create_client, Client

SCRIPT_DIR    = os.path.dirname(os.path.abspath(__file__))
PHASE4_SCRIPT = os.path.join(SCRIPT_DIR, "optimizer_v2.py")

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

@st.cache_resource
def get_supabase() -> Client:
    """Create Supabase client once and reuse it."""
    try:
        url = st.secrets["supabase"]["url"]
        key = st.secrets["supabase"]["key"]
    except Exception:
        # Fallback: read from environment or hardcoded (local dev only)
        url = os.environ.get("SUPABASE_URL", "")
        key = os.environ.get("SUPABASE_KEY", "")
    if not url or not key:
        st.error("Supabase credentials not found. Add them to .streamlit/secrets.toml")
        st.stop()
    return create_client(url, key)

EXCEL_PHASE4 = os.path.join(SCRIPT_DIR, "Planning_Tomate_2026.xlsx")

def load_from_excel():
    """Load planning data directly from local Excel file."""
    if not os.path.exists(EXCEL_PHASE4):
        return None, None, None, None, None

    planning = pd.read_excel(EXCEL_PHASE4, sheet_name="Planning Journalier", header=0)
    planning["Date"]             = pd.to_datetime(planning["Date"],       errors="coerce")
    planning["Commercial"]       = planning.get("Commercial", pd.Series())
    planning["Agriculteur"]      = planning.get("Agriculteur", pd.Series())
    planning["Usine"]            = planning.get("Usine", pd.Series())
    planning["Région"]           = planning.get("Region", planning.get("Région", ""))
    planning["Accessibilité"]    = planning.get("Accessibilite", planning.get("Accessibilité", ""))
    planning["Tonnes/Jour"]      = pd.to_numeric(planning.get("Tonnes/Jour", 0), errors="coerce")
    planning["Type Véhicule"]    = planning.get("Type Vehicule", planning.get("Type Véhicule", ""))
    planning["Véhicules Requis"] = planning.get("Vehicules", planning.get("Véhicules Requis", ""))
    planning["Nb Voyages"]       = pd.to_numeric(planning.get("Nb Voyages", 0), errors="coerce")
    planning["Date Début"]       = pd.to_datetime(planning.get("Date Debut", planning.get("Date Début", None)), errors="coerce")
    planning["Date Fin"]         = pd.to_datetime(planning.get("Date Fin", None), errors="coerce")
    planning["Total Tonnes"]     = pd.to_numeric(planning.get("Total Tonnes", 0), errors="coerce")
    pic_col = "Pic de Recolte" if "Pic de Recolte" in planning.columns else "Pic de Récolte"
    planning["Pic de Récolte"]   = planning.get(pic_col, "").apply(lambda x: "🟡 PIC" if "PIC" in str(x).upper() else "")
    planning["Note"]             = planning.get("Note", "")
    planning = planning.dropna(subset=["Date"])

    transport = pd.read_excel(EXCEL_PHASE4, sheet_name="Besoins Transport-Jour", header=0)
    transport["Date"]      = pd.to_datetime(transport["Date"], errors="coerce")
    transport["Commercial"]= transport.get("Commercial", "")
    transport["Total Tonnes"] = pd.to_numeric(transport.get("Total Tonnes", 0), errors="coerce")
    transport["Voyages TRACTEUR"]      = pd.to_numeric(transport.get("Voyages TRACTEUR", 0),      errors="coerce").fillna(0)
    transport["Voyages PETIT POILOUR"] = pd.to_numeric(transport.get("Voyages PETIT POILOUR", 0), errors="coerce").fillna(0)
    transport["Voyages POILOUR"]       = pd.to_numeric(transport.get("Voyages POILOUR", 0),       errors="coerce").fillna(0)
    transport["Voyages SEMI"]          = pd.to_numeric(transport.get("Voyages SEMI", 0),          errors="coerce").fillna(0)
    transport["Jours Double"]          = pd.to_numeric(transport.get("Jours Double", 0),          errors="coerce").fillna(0)
    transport = transport.dropna(subset=["Date"])

    double_j = pd.DataFrame(columns=[
        "Commercial","Agriculteur A (finit tôt)","Agriculteur B (reçoit véhicule)",
        "Véhicule Partagé","Jours Économisés","Fin Orig. A","Nouvelle Fin A",
        "Début Orig. B","Nouveau Début B","Risque Maladie","Action Requise"
    ])

    return planning, transport, pd.DataFrame(), double_j, None

@st.cache_data(ttl=60)
def load_data(_sb_version: int = 0):
    """
    Load data — tries Supabase first, falls back to local Excel
    if Supabase has incomplete data (stops before August).
    """
    sb = get_supabase()
    DATA_SOURCE = "supabase"

    def fetch(table, order_col=None, limit=10000):
        """Récupère TOUTES les lignes avec pagination (Supabase limite à 1000/req)."""
        all_rows = []
        page_size = 1000
        offset = 0
        while True:
            q = sb.table(table).select("*")
            if order_col:
                q = q.order(order_col)
            batch = q.range(offset, offset + page_size - 1).execute().data
            if not batch:
                break
            all_rows.extend(batch)
            if len(batch) < page_size or len(all_rows) >= limit:
                break
            offset += page_size
        return pd.DataFrame(all_rows)

    # ── Planning ──
    planning = fetch("planning", order_col="date")
    if not planning.empty:
        planning["Date"]       = pd.to_datetime(planning["date"],       errors="coerce")
        planning["Commercial"] = planning["commercial"]
        # ✅ Ne garder que les commerciaux qui ont des données dans agriculteurs
        # Évite d'afficher de vieilles données d'anciens commerciaux
        try:
            _active_comms = sb.table("agriculteurs").select("commercial").execute().data
            _active_comms = list({r["commercial"] for r in _active_comms if r.get("commercial")})
            if _active_comms:
                planning = planning[planning["commercial"].isin(_active_comms)]
        except Exception:
            pass  # si erreur → garder tout
        planning["Agriculteur"]= planning["agriculteur"]
        planning["Usine"]      = planning["usine"]
        planning["Région"]     = planning["region"]
        planning["Accessibilité"] = planning["accessibilite"]
        planning["Tonnes/Jour"]= pd.to_numeric(planning["tonnes_jour"], errors="coerce")
        planning["Type Véhicule"] = planning["type_vehicule"]
        planning["Véhicules Requis"] = planning["vehicules"]
        planning["Nb Voyages"] = pd.to_numeric(planning["nb_voyages"], errors="coerce")
        planning["Date Début"] = pd.to_datetime(planning["date_debut"], errors="coerce")
        planning["Date Fin"]   = pd.to_datetime(planning["date_fin"],   errors="coerce")
        planning["Total Tonnes"] = pd.to_numeric(planning["total_tonnes"], errors="coerce")
        planning["Pic de Récolte"] = planning["pic"].apply(lambda x: "🟡 PIC" if x else "")
        planning["Note"]       = planning["note"]
        planning = planning.dropna(subset=["Date"])

        # ── Check if Supabase data is complete (must reach August) ──
        if planning["Date"].max().date() < datetime.date(2026, 7, 31):
            DATA_SOURCE = "excel_fallback"
            planning = pd.DataFrame()  # force fallback

    # ── Fallback to local Excel if Supabase is empty or incomplete ──
    if planning.empty and os.path.exists(EXCEL_PHASE4):
        DATA_SOURCE = "excel"
        result = load_from_excel()
        if result[0] is not None:
            planning, transport, dispo, double_j, _ = result
            # build resume from Supabase agriculteurs (not hardcoded)
            try:
                _agri = sb.table("agriculteurs").select("commercial,nom,tonnage_total").execute().data
                _adf  = pd.DataFrame(_agri) if _agri else pd.DataFrame()
                if not _adf.empty:
                    _adf["tonnage_total"] = pd.to_numeric(_adf["tonnage_total"], errors="coerce")
                    _adf = _adf[_adf["tonnage_total"] > 0]
                    # Filtrer lignes TOTAL et noms invalides
                    _nom_e = _adf["nom"].astype(str).str.strip().str.upper()
                    _adf = _adf[~_nom_e.str.startswith("TOTAL")]
                    _adf = _adf[~_nom_e.str.startswith("SOUS-TOTAL")]
                    _adf = _adf[_nom_e.str.len() > 2]
                    if not _adf.empty:
                        _grp  = _adf.groupby("commercial")
                        resume = pd.DataFrame({
                            "Commercial":            _grp["tonnage_total"].sum().round(0).index,
                            "Tonnes Totales Saison": _grp["tonnage_total"].sum().round(0).values,
                            "Nb Agriculteurs":       _grp["nom"].nunique().values,
                            "Conflits Résolus":      0,
                            "Total Jours Double":    0,
                        })
                    else:
                        resume = pd.DataFrame()
                else:
                    resume = pd.DataFrame()
            except Exception:
                resume = pd.DataFrame()
            # Store source for display
            st.session_state["data_source"] = "📁 Local Excel"
            return planning, transport, pd.DataFrame(), double_j, resume

    st.session_state["data_source"] = "🗄️ Supabase" if DATA_SOURCE == "supabase" else "📁 Excel (Supabase incomplet)"

    # ── Transport ──
    transport = fetch("transport", order_col="date")
    if not transport.empty:
        transport["Date"]      = pd.to_datetime(transport["date"], errors="coerce")
        transport["Commercial"]= transport["commercial"]
        transport["Total Tonnes"] = pd.to_numeric(transport["total_tonnes"], errors="coerce")
        transport["Voyages TRACTEUR"]      = pd.to_numeric(transport["tracteur"],      errors="coerce").fillna(0)
        transport["Voyages PETIT POILOUR"] = pd.to_numeric(transport["petit_poilour"], errors="coerce").fillna(0)
        transport["Voyages POILOUR"]       = pd.to_numeric(transport["poilour"],       errors="coerce").fillna(0)
        transport["Voyages SEMI"]          = pd.to_numeric(transport["semi"],          errors="coerce").fillna(0)
        transport["Jours Double"]          = pd.to_numeric(transport["jours_double"],  errors="coerce").fillna(0)
        transport = transport.dropna(subset=["Date"])

    # ── Decalage (journal double transport) ──
    double_j = fetch("decalage")
    if not double_j.empty and "commercial" in double_j.columns:
        double_j["Commercial"] = double_j["commercial"]
        double_j["Agriculteur A (finit tôt)"]       = double_j.get("agriculteur_a", "")
        double_j["Agriculteur B (reçoit véhicule)"] = double_j.get("agriculteur_b", "")
        double_j["Véhicule Partagé"]  = double_j.get("vehicule", "")
        double_j["Jours Économisés"]  = double_j.get("shift_jours", 0)
        double_j["Fin Orig. A"]       = pd.to_datetime(double_j.get("fin_orig_a"),       errors="coerce")
        double_j["Nouvelle Fin A"]    = pd.to_datetime(double_j.get("nouvelle_fin_a"),   errors="coerce")
        double_j["Début Orig. B"]     = pd.to_datetime(double_j.get("debut_orig_b"),     errors="coerce")
        double_j["Nouveau Début B"]   = pd.to_datetime(double_j.get("nouveau_debut_b"),  errors="coerce")
        double_j["Risque Maladie"]    = double_j.get("risque", "")
        double_j["Action Requise"]    = double_j.get("action", "")
        double_j = double_j.dropna(subset=["Commercial"])
    else:
        # Empty table (optimizer handles conflicts internally — no decalage rows)
        double_j = pd.DataFrame(columns=[
            "Commercial","Agriculteur A (finit tôt)","Agriculteur B (reçoit véhicule)",
            "Véhicule Partagé","Jours Économisés","Fin Orig. A","Nouvelle Fin A",
            "Début Orig. B","Nouveau Début B","Risque Maladie","Action Requise"
        ])

    # ── Resume par commercial — from agriculteurs table (real uploaded data) ──
    resume = pd.DataFrame()
    try:
        agri_raw = fetch("agriculteurs")
        if not agri_raw.empty and "commercial" in agri_raw.columns:
            agri_raw["tonnage_total"] = pd.to_numeric(agri_raw["tonnage_total"], errors="coerce")
            agri_raw = agri_raw[agri_raw["tonnage_total"] > 0]
            # Filtrer les lignes TOTAL et noms invalides
            _nom_r = agri_raw["nom"].astype(str).str.strip().str.upper()
            agri_raw = agri_raw[~_nom_r.str.startswith("TOTAL")]
            agri_raw = agri_raw[~_nom_r.str.startswith("SOUS-TOTAL")]
            agri_raw = agri_raw[_nom_r.str.len() > 2]
            if not agri_raw.empty:
                agri_grp = agri_raw.groupby("commercial")
                resume   = agri_grp["tonnage_total"].sum().reset_index()
                resume.columns = ["Commercial", "Tonnes Totales Saison"]
                resume["Tonnes Totales Saison"] = resume["Tonnes Totales Saison"].round(0)
                resume["Nb Agriculteurs"] = agri_grp["nom"].nunique().values
                resume["Conflits Résolus"]   = 0
                resume["Total Jours Double"] = 0
    except Exception:
        pass

    # dispo not in Supabase yet — return empty
    dispo = pd.DataFrame()

    return planning, transport, dispo, double_j, resume

# ── Load data from Supabase ───────────────────────────────────
# sb_refresh_counter increments when user clicks "Régénérer"
if "sb_refresh" not in st.session_state:
    st.session_state["sb_refresh"] = 0

result = load_data(_sb_version=st.session_state["sb_refresh"])
planning, transport, dispo, double_j, resume = result
orig = None  # not needed for dashboard display

# ── GLOBAL CONSTANTS — read from Supabase agriculteurs ──────
@st.cache_data(ttl=30)
def load_global_stats(_version: int = 0):
    """
    Charge les stats globales depuis Supabase.
    ✅ Écrase les anciennes données avec les nouvelles uploads
    ✅ Normalise les régions incohérentes (nabeul→CAP BON 2, beja→NORD…)
    ✅ Filtre les données corrompues (tonnage=0, dates invalides)
    """
    try:
        sb = get_supabase()

        # Commerciaux qui ont uploadé via le dashboard
        try:
            depot_data = sb.table("depot_status").select(
                "commercial,statut,depose_le").execute().data
            deposited  = {r["commercial"]: r.get("depose_le","")
                          for r in depot_data if r.get("statut") == "depose"}
        except Exception:
            deposited = {}

        # Toutes les données agriculteurs
        data = sb.table("agriculteurs").select(
            "commercial,nom,tonnage_total,usine,region,zone,date_debut,nbr_hectares,centre").execute().data
        df = pd.DataFrame(data) if data else pd.DataFrame()
        if df.empty:
            raise ValueError("empty")

        df["tonnage_total"] = pd.to_numeric(df["tonnage_total"], errors="coerce")
        df = df[df["tonnage_total"] > 0]   # filtre tonnage nul
        # Filtrer les lignes de TOTAL qui ont pu être insérées par erreur
        _nom = df["nom"].astype(str).str.strip().str.upper()
        df = df[~_nom.str.startswith("TOTAL")]
        df = df[~_nom.str.startswith("SOUS-TOTAL")]
        df = df[_nom.str.len() > 2]

        # Normalisation régions (évite les doublons nabeul/NABEUL/beja/MANOUBA)
        REGION_NORM = {
            "nabeul": "CAP BON 2", "NABEUL": "CAP BON 2",
            "beja": "NORD",        "BEJA": "NORD",
            "manouba": "NORD",     "MANOUBA": "NORD",
            "gafsa": "GAFSA / KASSRINE", "GAFSA": "GAFSA / KASSRINE",
            "kassrine": "GAFSA / KASSRINE", "KASSRINE": "GAFSA / KASSRINE",
            "capb1": "CAP BON 1",  "CAPB1": "CAP BON 1",
            "capb2": "CAP BON 2",  "CAPB2": "CAP BON 2",
        }
        df["region"] = df["region"].fillna("").astype(str).str.strip()
        df["region"] = df["region"].replace(REGION_NORM)

        # Filtre dates invalides — permissif pour ne pas perdre des données valides
        # (On garde tout, les dates invalides ne suppriement pas les tonnages)
        df["date_debut"] = pd.to_datetime(df["date_debut"], errors="coerce")
        # Supprimer SEULEMENT les lignes avec des dates vraiment absurdes (avant 2000)
        _bad_dates = df["date_debut"].notna() & (df["date_debut"].dt.year < 2000)
        if _bad_dates.sum() < len(df):   # ne pas tout supprimer
            df = df[~_bad_dates]

        # Priorité upload récent : si un commercial a uploadé,
        # ses données Supabase sont les plus récentes — garder toutes
        # (le delete() lors de l'upload a déjà écrasé l'ancien)
        df_use = df

        return {
            "total_tons":          round(df_use["tonnage_total"].sum(), 0),
            "n_farmers":           int(df_use["nom"].nunique()),
            "commercial_tons":     df_use.groupby("commercial")["tonnage_total"].sum().round(0).to_dict(),
            "commercial_farmers":  df_use.groupby("commercial")["nom"].nunique().to_dict(),
            "usine_tons":          df_use.groupby("usine")["tonnage_total"].sum().round(0).to_dict(),
            "region_tons":         df_use.groupby("region")["tonnage_total"].sum().round(0).to_dict(),
            "deposited":           list(deposited.keys()),
            "total_rows":          len(df_use),
            "df_agri":             df_use,   # ← DataFrame complet partagé globalement
            "data_quality": {
                "regions_ok":   df_use["region"].isin([
                    "CAP BON 1","CAP BON 2","NORD","GAFSA / KASSRINE",
                    "KAIROUAN","SIDI BOUZID","BOUFICHA"
                ]).sum(),
                "total_rows":   len(df_use),
            }
        }
    except Exception as e:
        return {
            "total_tons": 0.0, "n_farmers": 0,
            "commercial_tons": {}, "commercial_farmers": {},
            "usine_tons": {}, "region_tons": {},
            "deposited": [], "total_rows": 0,
            "df_agri": pd.DataFrame(),   # DataFrame vide en cas d'erreur
            "data_quality": {"regions_ok": 0, "total_rows": 0},
        }

_stats = load_global_stats(_version=st.session_state["sb_refresh"])
GLOBAL_TOTAL_TONS         = _stats["total_tons"]
GLOBAL_N_FARMERS          = _stats["n_farmers"]
AGRI_DF                   = _stats.get("df_agri", pd.DataFrame())  # DataFrame global agriculteurs
GLOBAL_COMMERCIAL_TONS    = _stats["commercial_tons"]
GLOBAL_COMMERCIAL_FARMERS = _stats["commercial_farmers"]
GLOBAL_USINE_TONS         = _stats["usine_tons"]

GLOBAL_PEAK_TONS   = round(planning[
    (planning["Date"].dt.date >= PEAK_START) &
    (planning["Date"].dt.date <= PEAK_END)
]["Tonnes/Jour"].sum(), 0) if not planning.empty else 0
GLOBAL_N_CONFLICTS = len(double_j) if not double_j.empty else 0

# ── Apply role-based data filtering ──────────────────────────
if planning is not None and CURRENT_FILTER:
    if CURRENT_ROLE == "commercial":
        if not planning.empty and "Commercial" in planning.columns:
            planning = planning[planning["Commercial"] == CURRENT_FILTER]
        if not transport.empty and "Commercial" in transport.columns:
            transport = transport[transport["Commercial"] == CURRENT_FILTER]
        if dispo is not None and not dispo.empty and "Commercial" in dispo.columns:
            dispo = dispo[dispo["Commercial"] == CURRENT_FILTER]
        if double_j is not None and not double_j.empty and "Commercial" in double_j.columns:
            double_j = double_j[double_j["Commercial"] == CURRENT_FILTER]
        if resume is not None and not resume.empty and "Commercial" in resume.columns:
            resume = resume[resume["Commercial"] == CURRENT_FILTER]
    elif CURRENT_ROLE == "usine":
        if not planning.empty and "Usine" in planning.columns:
            planning = planning[planning["Usine"] == CURRENT_FILTER]
        # FIX: ne pas vider transport pour usine — garder tout le transport
        # (le transport est par commercial, pas filtrable par usine sans perte info)
        # avant: transport.iloc[0:0] rendait les courbes transport VIDES
        double_j = double_j.iloc[0:0]

# ── Sidebar ───────────────────────────────────────────────────
with st.sidebar:
    # ── User info + logout ────────────────────────────────────
    role_icons = {"directeur": "👑", "commercial": "👤", "usine": "🏭"}
    role_labels = {"directeur": "Directeur", "commercial": "Commercial", "usine": "Usine"}
    st.markdown(f"""
    <div style='background:#1c2333;border:1px solid #21262d;border-radius:10px;
    padding:14px 16px;margin-bottom:16px'>
      <div style='font-size:1.1rem;font-weight:700;color:#f0f6fc'>
        {role_icons.get(CURRENT_ROLE,"👤")} {CURRENT_NAME}
      </div>
      <div style='font-size:.72rem;color:#8b949e;margin-top:3px;text-transform:uppercase;letter-spacing:.07em'>
        {role_labels.get(CURRENT_ROLE, CURRENT_ROLE)}
      </div>
    </div>
    """, unsafe_allow_html=True)
    if st.button("🚪 Déconnexion", use_container_width=True):
        for key in list(st.session_state.keys()):
            del st.session_state[key]
        # Effacer le token de l'URL pour ne pas rester connecté au refresh
        try:
            st.query_params.clear()
        except Exception:
            pass
        st.rerun()
    st.divider()


    # Régénérer planning — directeur seulement
    if CURRENT_ROLE == "directeur":
        st.subheader("⚙️ Mise à jour données")
        st.caption("Recalcule le planning et met à jour Supabase.")
        if st.button("🔄 Régénérer le planning", use_container_width=True, type="primary"):
            # ✅ SÉCURITÉ: vérifier qu'aucune régénération n'est en cours
            _regen_locked = False
            try:
                _sb_lock = get_supabase()
                _lock = _sb_lock.table("app_locks").select("*").eq(
                    "lock_name", "planning_regen").execute().data
                import datetime as _dt
                if _lock and _lock[0].get("expires_at"):
                    _exp = _dt.datetime.fromisoformat(
                        _lock[0]["expires_at"].replace("Z","+00:00"))
                    if _exp > _dt.datetime.now(_dt.timezone.utc):
                        _regen_locked = True
                        _by = _lock[0].get("locked_by","?")
                        st.warning(f"⚠️ Régénération déjà en cours par {_by}. Réessaye dans 2 min.")
            except Exception:
                pass  # Si table n'existe pas, continuer

            if not _regen_locked and os.path.exists(PHASE4_SCRIPT):
                # Poser le verrou
                try:
                    import datetime as _dt2
                    _exp_time = (_dt2.datetime.now(_dt2.timezone.utc) + 
                                 _dt2.timedelta(minutes=10)).isoformat()
                    get_supabase().table("app_locks").upsert({
                        "lock_name":  "planning_regen",
                        "locked_by":  CURRENT_USER,
                        "locked_at":  _dt2.datetime.now(_dt2.timezone.utc).isoformat(),
                        "expires_at": _exp_time,
                    }).execute()
                except Exception:
                    pass
                
                migrate_script = os.path.join(SCRIPT_DIR, "migrate.py")
                with st.spinner("Etape 1/2 : Calcul du planning..."):
                    r1 = subprocess.run(
                        [sys.executable, PHASE4_SCRIPT],
                        capture_output=True, text=True, timeout=600,
                        cwd=SCRIPT_DIR, encoding="utf-8", errors="replace",
                    )
                if r1.returncode != 0:
                    st.error("Erreur optimizer_v2.py :")
                    st.code(r1.stderr[-400:], language="text")
                elif os.path.exists(migrate_script):
                    with st.spinner("Etape 2/2 : Mise à jour Supabase..."):
                        r2 = subprocess.run(
                            [sys.executable, migrate_script],
                            capture_output=True, text=True, timeout=300,
                            cwd=SCRIPT_DIR, encoding="utf-8", errors="replace",
                        )
                    if r2.returncode == 0:
                        # Libérer le verrou
                        try:
                            get_supabase().table("app_locks").delete().eq(
                                "lock_name", "planning_regen").execute()
                            # Audit log
                            get_supabase().table("audit_log").insert({
                                "user_name": CURRENT_USER,
                                "user_role": CURRENT_ROLE,
                                "action":    "planning_regenerated",
                                "details":   f"Planning régénéré avec succès",
                            }).execute()
                        except Exception:
                            pass
                        st.success("✅ Planning recalculé et Supabase mis à jour !")
                        st.session_state["sb_refresh"] += 1
                        st.cache_data.clear()
                        st.rerun()
                    else:
                        st.error("Erreur migrate.py :")
                        st.code(r2.stderr[-400:], language="text")
                else:
                    st.success("✅ optimizer_v2.py terminé (migrate.py non trouvé)")
                    st.cache_data.clear()
                    st.rerun()
            else:
                st.warning(f"❌ optimizer_v2.py introuvable dans : {SCRIPT_DIR}")
        st.divider()

    # Filters — directeur sees all, commercial/usine see only their data
    st.subheader("🔍 Filtres")
    if planning is not None and not planning.empty:
        if CURRENT_ROLE == "directeur":
            comms_all = sorted(planning["Commercial"].dropna().unique())
            sel_comms = st.multiselect("Commercial(s)", comms_all, default=comms_all)
            if "Usine" in planning.columns:
                facts_all = sorted(planning["Usine"].dropna().unique())
                sel_facts = st.multiselect("Usine(s)", facts_all, default=facts_all)
            else:
                sel_facts = []
        else:
            # Commercial or usine: filter already applied above, no choice
            sel_comms = list(planning["Commercial"].dropna().unique()) if "Commercial" in planning.columns else []
            sel_facts = list(planning["Usine"].dropna().unique()) if "Usine" in planning.columns else []
            if CURRENT_ROLE == "commercial":
                st.info(f"👤 Vue filtrée : **{CURRENT_NAME}**")
            elif CURRENT_ROLE == "usine":
                st.info(f"🏭 Vue filtrée : **{CURRENT_NAME}**")

        # Use full season range — not limited by what's in Supabase
        d_min = datetime.date(2026, 6, 15)
        d_max = datetime.date(2026, 8, 31)
        # Default value: actual data range from planning
        d_data_min = planning["Date"].min().date() if not planning.empty else d_min
        d_data_max = planning["Date"].max().date() if not planning.empty else d_max
        date_range = st.date_input("Période", value=(d_data_min, d_data_max),
                                   min_value=d_min, max_value=d_max)
        peak_only = st.checkbox("⚡ Pic seulement (1–15 Jul)")
    else:
        sel_comms, sel_facts = [], []
        date_range = None
        peak_only = False

    st.divider()

    # Fleet inventory — directeur only
    if CURRENT_ROLE == "directeur":
        st.subheader("🚛 Votre flotte")
        st.caption("ℹ️ Ces valeurs servent uniquement aux **alertes de capacité** dans l'onglet Transport — elles ne recalculent pas le planning. Pour modifier le planning, il faut relancer `optimizer_v2.py`.")
        fl_trac = st.number_input("TRACTEUR",       0, 20, 0)
        fl_ppl  = st.number_input("PETIT POILOUR", 0, 30, 3)
        fl_pl   = st.number_input("POILOUR",       0, 30, 6)
        fl_semi = st.number_input("SEMI",          0, 20, 4)
        st.divider()
    else:
        fl_trac, fl_ppl, fl_pl, fl_semi = 0, 3, 6, 4  # defaults, not shown

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
    if planning is not None and not planning.empty:
        last_date = planning["Date"].max().strftime("%d/%m/%Y")
        st.caption(f"📅 Données jusqu'au {last_date}")
        source = st.session_state.get("data_source", "🗄️ Supabase")
        st.caption(f"Source : {source}")

# ── Guard: Supabase empty ─────────────────────────────────────
if planning is None or planning.empty:
    st.markdown("---")
    if CURRENT_ROLE == "directeur":
        st.warning("⚠️ Aucune donnée de planning dans Supabase pour le moment.")
        st.markdown("### 🚀 Pour démarrer — suivez ces 3 étapes :")

        col1, col2, col3 = st.columns(3)
        with col1:
            st.markdown("""
            **Étape 1 — Commerciaux uploadent**
            - Chaque commercial se connecte
            - Onglet **📤 Upload Planning**
            - Télécharge le modèle Excel
            - Le remplit et l'uploade
            """)
        with col2:
            st.markdown("""
            **Étape 2 — Générer le planning**
            - Dans le terminal, lance :
            ```
            python optimizer_v2.py
            python migrate.py
            ```
            - Ou clique le bouton ci-dessous ↓
            """)
        with col3:
            st.markdown("""
            **Étape 3 — Rafraîchir**
            - Le dashboard se met à jour
            - Toutes les données s'affichent
            - Graphiques et statistiques OK
            """)

        st.markdown("---")

        # Allow directeur to trigger generation even when empty
        col_a, col_b = st.columns(2)
        with col_a:
            if st.button("🔄 Générer le planning maintenant",
                         type="primary", use_container_width=True):
                optimizer_script = os.path.join(SCRIPT_DIR, "optimizer_v2.py")
                migrate_script   = os.path.join(SCRIPT_DIR, "migrate.py")
                if os.path.exists(optimizer_script):
                    with st.spinner("Étape 1/2 : Calcul OR-Tools en cours (peut prendre 2 min)..."):
                        r1 = subprocess.run(
                            [sys.executable, optimizer_script],
                            capture_output=True, text=True, timeout=300,
                            cwd=SCRIPT_DIR, encoding="utf-8", errors="replace",
                        )
                    if r1.returncode != 0:
                        st.error("Erreur optimizer_v2.py :")
                        st.code(r1.stderr[-600:], language="text")
                    elif os.path.exists(migrate_script):
                        with st.spinner("Étape 2/2 : Mise à jour Supabase..."):
                            r2 = subprocess.run(
                                [sys.executable, migrate_script],
                                capture_output=True, text=True, timeout=300,
                                cwd=SCRIPT_DIR, encoding="utf-8", errors="replace",
                            )
                        if r2.returncode == 0:
                            st.success("✅ Planning généré et Supabase mis à jour !")
                            st.session_state["sb_refresh"] += 1
                            st.cache_data.clear()
                            st.rerun()
                        else:
                            st.error("Erreur migrate.py :")
                            st.code(r2.stderr[-400:], language="text")
                else:
                    st.error(f"optimizer_v2.py introuvable dans : {SCRIPT_DIR}")
        with col_b:
            if st.button("🔄 Rafraîchir les données",
                         use_container_width=True):
                st.session_state["sb_refresh"] += 1
                st.cache_data.clear()
                st.rerun()

        # Still show upload tab
        st.markdown("---")
        st.info("💡 Les commerciaux peuvent déposer leurs fichiers via l'onglet **📤 Upload Planning** même si le planning n'est pas encore généré.")

    elif CURRENT_ROLE == "commercial":
        st.info(f"👋 Bonjour **{CURRENT_NAME}** — Le planning n'est pas encore disponible.")
        st.markdown("Déposez votre fichier Excel via l'onglet **📤 Upload Planning** pour commencer.")

    elif CURRENT_ROLE == "usine":
        st.info(f"👋 Bonjour **{CURRENT_NAME}** — Le planning n'est pas encore disponible.")
        st.markdown("Contactez le directeur pour qu'il génère le planning.")

    # Show upload tab even when no planning data
    if CURRENT_ROLE in ("directeur", "commercial"):
        st.markdown("---")
        try:
            with st.expander("📤 Accéder à l'upload de planning", expanded=True):
                from upload_tab import render_upload_tab
                render_upload_tab(
                    sb=get_supabase(),
                    CURRENT_ROLE=CURRENT_ROLE,
                    CURRENT_NAME=CURRENT_NAME,
                    CURRENT_FILTER=CURRENT_FILTER,
                    GLOBAL_COMMERCIAL_FARMERS=GLOBAL_COMMERCIAL_FARMERS,
                    GLOBAL_COMMERCIAL_TONS=GLOBAL_COMMERCIAL_TONS,
                    df_to_csv=df_to_csv,
                )
        except Exception:
            pass
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
st.caption("Phase 10 — Connecté à Supabase · Login par rôle · OR-Tools Optimizer")

# KPIs — GLOBAL_N_FARMERS uses unique names from agriculteurs table
total_tons   = GLOBAL_TOTAL_TONS

# Statut OR-Tools: lire depuis planning (si FEASIBLE = solution trouvée, sinon vide)
_n_planning_days = len(planning["Date"].dt.date.unique()) if not planning.empty else 0
if _n_planning_days >= 60:
    ortools_status = "FEASIBLE"
    ortools_sub    = f"{_n_planning_days} jours planifiés ✅"
elif _n_planning_days > 0:
    ortools_status = "PARTIEL"
    ortools_sub    = f"⚠️ {_n_planning_days}/67 jours — relancer optimizer"
else:
    ortools_status = "—"
    ortools_sub    = "Planning non généré"

# Tonnage planifié (somme planning) vs déclaré
planned_tons = round(planning["Tonnes/Jour"].sum(), 0) if not planning.empty else 0
ecart_pct    = round((planned_tons - total_tons) / total_tons * 100, 1) if total_tons > 0 else 0
total_trips  = int(p["Nb Voyages"].sum()) if not p.empty and "Nb Voyages" in p.columns else 0
# Unique farmer names — not rows (one farmer can deliver to 2 usines = 2 rows)
# Toujours utiliser le nombre réel d'agriculteurs déclarés (table agriculteurs)
# pas le nombre dans le planning (qui peut être incomplet ou filtré)
n_farmers = GLOBAL_N_FARMERS
peak_tons    = GLOBAL_PEAK_TONS
n_conflicts  = GLOBAL_N_CONFLICTS
dbl_days     = int(t["Jours Double"].sum()) if not t.empty and "Jours Double" in t.columns else 0

st.markdown(f"""
<div class="metric-row">
  <div class="kpi-box" style="--c:#e8543a"><div class="kpi-val">{int(total_tons):,} t</div><div class="kpi-lbl">Déclaré (saison)</div><div class="kpi-sub">{n_farmers} agriculteurs</div></div>
  <div class="kpi-box" style="--c:#22c55e"><div class="kpi-val">{int(planned_tons):,} t</div><div class="kpi-lbl">Planifié OR-Tools</div><div class="kpi-sub">Écart {ecart_pct:+.1f}% vs déclaré</div></div>
  <div class="kpi-box" style="--c:#f5a623"><div class="kpi-val">{int(peak_tons):,} t</div><div class="kpi-lbl">Tonnes période pic</div><div class="kpi-sub">1–15 Juillet</div></div>
  <div class="kpi-box" style="--c:#8b5cf6"><div class="kpi-val">{total_trips:,}</div><div class="kpi-lbl">Voyages (vue actuelle)</div><div class="kpi-sub">PPL + PL + SEMI</div></div>
  <div class="kpi-box" style="--c:#00e5a0"><div class="kpi-val">{ortools_status}</div><div class="kpi-lbl">Statut OR-Tools</div><div class="kpi-sub">{ortools_sub}</div></div>
  <div class="kpi-box" style="--c:#3b82f6"><div class="kpi-val">{len(sel_comms)}</div><div class="kpi-lbl">Commerciaux</div><div class="kpi-sub">{_n_planning_days} jours planifiés</div></div>
</div>
<div class="peak-box">⚡ <b>Pic 1–15 Juillet :</b> Les caps usines et commerciaux s'appliquent UNIQUEMENT pendant cette période. Hors pic: distribution libre selon les fenêtres de maturité.</div>
""", unsafe_allow_html=True)

# ── Warning planning incomplet (UNIQUEMENT pour directeur) ──
# Pour usine/commercial, le planning est NATURELLEMENT filtré :
#   - COMOCAP voit ~52j (ses jours de réception)
#   - JILANI voit ~27j (ses fermiers commencent Jul 23)
#   → Ce n'est PAS un planning incomplet, c'est la vue filtrée !
if CURRENT_ROLE == "directeur":
    if _n_planning_days > 0 and _n_planning_days < 60:
        st.warning(
            f"⚠️ **Planning incomplet** : seulement **{_n_planning_days} jours** dans Supabase "
            f"au lieu de 67. Les graphiques et calculs sont partiels. "
            f"**Solution :** Sidebar → 🔄 Régénérer le planning"
        )
    elif _n_planning_days == 0:
        st.error("❌ Aucun planning dans Supabase. Lance : `python optimizer_v2.py` puis `python migrate.py`")
else:
    # Pour usine/commercial : afficher seulement si VRAIMENT vide (= 0 jour)
    if _n_planning_days == 0:
        st.warning(f"⚠️ Pas de planning trouvé pour {CURRENT_NAME}. "
                   f"Demande au directeur de relancer optimizer + migrate.")

# ── Tabs — visibility depends on role ────────────────────────
# Usine    : voit seulement Par Usine + Transport (pas de commercial, décalage, historique, admin)
# Commercial: voit Planning + Par Commercial + Transport + Gestion de ses agriculteurs
# Directeur : voit tout

if CURRENT_ROLE == "centre":
    # ── Session CENTRE (BACCARA, KERKOUANE, 428) ──────────────
    # Rend un dashboard dédié et stoppe l'exécution du dashboard global
    try:
        from centre_tab import render_centre_dashboard
        render_centre_dashboard(get_supabase(), CURRENT_FILTER, CURRENT_NAME)
    except ImportError:
        st.error("Module centre_tab.py introuvable. Vérifier le déploiement.")
    except Exception as e:
        st.error(f"Erreur centre dashboard: {e}")
    st.stop()

if CURRENT_ROLE == "usine":
    tab3, tab4 = st.tabs([
        "🏭 Par Usine",
        "🚛 Transport & Alertes",
    ])
    tab1 = tab2 = tab3
    tab5 = tab6 = tab9 = tab10 = tab4
    tab7 = tab8 = tab4

elif CURRENT_ROLE == "commercial":
    tab1, tab2, tab4, tab9, tab10 = st.tabs([
        "📅 Planning Journalier",
        "👤 Par Commercial",
        "🚛 Transport & Alertes",
        "🌾 Mes Agriculteurs",
        "📤 Upload Planning",
    ])
    tab3 = tab1
    tab5 = tab6 = tab7 = tab8 = tab4

else:
    # Directeur : all tabs
    tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8, tab9, tab10 = st.tabs([
        "📅 Planning Journalier",
        "👤 Par Commercial",
        "🏭 Par Usine",
        "🚛 Transport & Alertes",
        "⚙️ Décalage & Conflits",
        "📈 Historique 2025 vs 2026",
        "🗺️ Tonnage par Région",
        "📊 Prévisions Déc→Mai→Juin",
        "🌾 Gestion Agriculteurs",
        "📤 Upload Planning",
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
        # Use DECLARED tonnage from agriculteurs (correct) not planning sum (filtered/incomplete)
        if GLOBAL_COMMERCIAL_TONS:
            comm_df = pd.DataFrame(list(GLOBAL_COMMERCIAL_TONS.items()),
                                    columns=["Commercial","Tonnes/Jour"])
        else:
            comm_df = p.groupby("Commercial")["Tonnes/Jour"].sum().reset_index()
        fig2 = px.pie(
            comm_df,
            names="Commercial", values="Tonnes/Jour",
            title="Répartition par commercial (tonnage déclaré)", hole=0.45,
            color="Commercial",
            color_discrete_map=COMM_COLORS,
            template="plotly_dark",
        )
        fig2.update_layout(paper_bgcolor="#161b22", height=320)
        st.plotly_chart(fig2, use_container_width=True)
    with c2:
        if "Usine" in p.columns:
            # Use declared tonnage from agriculteurs (correct) not planning sum (wrong)
            if GLOBAL_USINE_TONS:
                usine_df = pd.DataFrame(list(GLOBAL_USINE_TONS.items()),
                                         columns=["Usine","Tonnes/Jour"])
            else:
                usine_df = p.groupby("Usine")["Tonnes/Jour"].sum().reset_index()
            fig3 = px.pie(
                usine_df,
                names="Usine", values="Tonnes/Jour",
                title="Répartition par usine", hole=0.45,
                color="Usine", color_discrete_map=FACTORY_COLORS,
                template="plotly_dark",
            )
            fig3.update_layout(paper_bgcolor="#161b22", height=320)
            st.plotly_chart(fig3, use_container_width=True)

    st.subheader("📋 Données détaillées")
    
    # Barre de recherche par nom d'agriculteur
    col_search1, col_search2 = st.columns([3, 1])
    with col_search1:
        search_farmer = st.text_input(
            "🔍 Rechercher un agriculteur",
            placeholder="Tapez le nom (ex: ZOUHAIR ou SOUHAIL)",
            key="search_planning_farmer"
        )
    with col_search2:
        st.write("")  # spacer
        st.caption(f"Total: {len(p)} lignes")
    
    display_cols = [c for c in ["Date","Commercial","Agriculteur","Usine",
                                "Tonnes/Jour","Type Véhicule","Véhicules Requis","Nb Voyages",
                                "Pic de Récolte","Note"] if c in p.columns]
    p_display = p[display_cols].sort_values("Date").reset_index(drop=True)
    
    # Appliquer le filtre de recherche
    if search_farmer.strip():
        mask = p_display["Agriculteur"].astype(str).str.upper().str.contains(
            search_farmer.upper().strip(), na=False
        )
        p_display = p_display[mask]
        if len(p_display) == 0:
            st.warning(f"Aucun agriculteur trouvé pour '{search_farmer}'")
        else:
            total_tons_found = p_display["Tonnes/Jour"].sum()
            n_unique = p_display["Agriculteur"].nunique()
            st.success(f"✅ {n_unique} agriculteur(s) trouvé(s) — {len(p_display)} lignes — {total_tons_found:,.0f}t")
    
    st.dataframe(
        p_display,
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
        # Use DECLARED tonnage from agriculteurs table (source of truth)
        # NOT planning sum which is filtered/partial
        if GLOBAL_COMMERCIAL_TONS:
            comm_tot = pd.DataFrame(list(GLOBAL_COMMERCIAL_TONS.items()),
                                     columns=["Commercial","Tonnes/Jour"])
        else:
            comm_tot = p.groupby("Commercial")["Tonnes/Jour"].sum().reset_index()
        fig5 = px.bar(
            comm_tot, x="Commercial", y="Tonnes/Jour",
            color="Commercial", color_discrete_map=COMM_COLORS,
            title="Tonnes totales (déclarées par chaque commercial)",
            template="plotly_dark",
            text_auto=".3s",
        )
        fig5.update_layout(paper_bgcolor="#161b22", showlegend=False, height=300)
        st.plotly_chart(fig5, use_container_width=True)
    with c2:
        # Use DECLARED unique farmer count from agriculteurs table
        if GLOBAL_COMMERCIAL_FARMERS:
            farmers_ct = pd.DataFrame(list(GLOBAL_COMMERCIAL_FARMERS.items()),
                                       columns=["Commercial","Agriculteur"])
        else:
            farmers_ct = p.groupby("Commercial")["Agriculteur"].nunique().reset_index()
        fig6 = px.bar(
            farmers_ct, x="Commercial", y="Agriculteur",
            color="Commercial", color_discrete_map=COMM_COLORS,
            title="Nb agriculteurs (déclarés)", template="plotly_dark",
            text_auto=True,
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
                              "Type Véhicule","Véhicules Requis","Nb Voyages","Note"] if c in one.columns]
    st.dataframe(one[show_cols].sort_values("Date").reset_index(drop=True),
                 use_container_width=True, height=240)
    
    # ── 📊 NOUVEAU: Courbes Agriculteurs × Tonnage pour ce commercial ─
    st.markdown("---")
    st.subheader(f"📈 Agriculteurs de {selected} — Répartition des tonnages")
    
    if "Agriculteur" in one.columns and "Tonnes/Jour" in one.columns:
        # 1. Tonnage DÉCLARÉ depuis AGRI_DF (chargé une seule fois au démarrage)
        # AGRI_DF = table agriculteurs complète en mémoire — fiable, rapide, pas de pagination
        if not AGRI_DF.empty and "commercial" in AGRI_DF.columns:
            _comm_mask = AGRI_DF["commercial"] == selected
            _sel_agri  = AGRI_DF[_comm_mask].copy()
            _sel_agri["tonnage_total"] = pd.to_numeric(
                _sel_agri["tonnage_total"], errors="coerce").fillna(0)
            # Grouper par NOM (toutes parcelles et usines confondues)
            agri_totals = _sel_agri.groupby("nom")["tonnage_total"].sum().reset_index()
            agri_totals.columns = ["Agriculteur", "Tonnage Total (t)"]
            agri_totals = agri_totals.sort_values("Tonnage Total (t)", ascending=True)
            agri_totals["Tonnage Total (t)"] = agri_totals["Tonnage Total (t)"].round(0).astype(int)
            # ✅ Ajouter NBR_HECTARES et t/ha
            if "nbr_hectares" in _sel_agri.columns:
                ha_map = _sel_agri.groupby("nom")["nbr_hectares"].first()
                _ha = pd.to_numeric(agri_totals["Agriculteur"].map(ha_map), errors="coerce")
                agri_totals["Hectares"] = _ha.round(2)
                # t/ha calculé seulement si hectares > 0
                agri_totals["t/ha"] = (agri_totals["Tonnage Total (t)"] / _ha.replace(0, pd.NA)).round(1)
        else:
            # Fallback uniquement si AGRI_DF vide (connexion échouée au démarrage)
            agri_totals = one.groupby("Agriculteur")["Tonnes/Jour"].sum().reset_index()
            agri_totals.columns = ["Agriculteur", "Tonnage Total (t)"]
            agri_totals = agri_totals.sort_values("Tonnage Total (t)", ascending=True)
            agri_totals["Tonnage Total (t)"] = agri_totals["Tonnage Total (t)"].round(0).astype(int)
        
        n_agri = len(agri_totals)
        bar_height = max(280, n_agri * 22)
        
        col_a, col_b = st.columns([3, 1])
        with col_a:
            fig_bar = px.bar(
                agri_totals, x="Tonnage Total (t)", y="Agriculteur",
                orientation="h",
                title=f"Tonnage total par agriculteur — {selected} ({n_agri} agriculteurs)",
                color="Tonnage Total (t)", color_continuous_scale="Viridis",
                template="plotly_dark", height=bar_height,
                text="Tonnage Total (t)",
            )
            fig_bar.update_traces(textposition="outside")
            fig_bar.update_layout(
                paper_bgcolor="#161b22",
                xaxis_title="Tonnes (t)",
                yaxis_title="",
                showlegend=False,
                margin=dict(l=180),
            )
            st.plotly_chart(fig_bar, use_container_width=True)
        
        with col_b:
            if agri_totals.empty or agri_totals['Tonnage Total (t)'].sum() == 0:
                st.info("Aucune donnée")
            else:
                st.metric("Nb agriculteurs", n_agri)
                st.metric("Tonnage total",   f"{int(agri_totals['Tonnage Total (t)'].sum()):,}t")
                st.metric("Moyenne / agri",  f"{int(agri_totals['Tonnage Total (t)'].mean()):,}t")
                st.metric("Plus gros",       f"{int(agri_totals['Tonnage Total (t)'].max()):,}t")
                st.caption(f"Top: **{agri_totals.iloc[-1]['Agriculteur']}**")
        
        # 2. Évolution temporelle TOP 8 agriculteurs (courbes superposées)
        st.markdown(f"#### 📉 Courbes journalières — TOP 8 agriculteurs de {selected}")
        top_agri = agri_totals.tail(8)["Agriculteur"].tolist()
        one_top = one[one["Agriculteur"].isin(top_agri)].copy()
        
        if not one_top.empty:
            fig_lines = px.line(
                one_top.sort_values("Date"), x="Date", y="Tonnes/Jour",
                color="Agriculteur", template="plotly_dark", height=380,
                title=f"Évolution journalière des tonnes — TOP 8 agriculteurs de {selected}",
                markers=True,
            )
            fig_lines.update_layout(
                paper_bgcolor="#161b22",
                hovermode="x unified",
                legend=dict(orientation="h", yanchor="top", y=-0.15),
            )
            fig_lines.add_vrect(x0=str(PEAK_START), x1=str(PEAK_END),
                              fillcolor="gold", opacity=0.08, line_width=0,
                              annotation_text="PIC", annotation_position="top left")
            st.plotly_chart(fig_lines, use_container_width=True)
        
        # 3. Répartition par usine (donut)
        if "Usine" in one.columns:
            st.markdown(f"#### 🏭 Répartition par usine — {selected}")
            usine_dist = one.groupby("Usine")["Tonnes/Jour"].sum().reset_index()
            usine_dist["Tonnes/Jour"] = usine_dist["Tonnes/Jour"].round(0).astype(int)
            usine_dist = usine_dist.sort_values("Tonnes/Jour", ascending=False)
            
            col_pie, col_table = st.columns([1, 1])
            with col_pie:
                fig_pie = px.pie(
                    usine_dist, names="Usine", values="Tonnes/Jour",
                    hole=0.45, template="plotly_dark",
                    title=f"Tonnage par usine — {selected}",
                )
                fig_pie.update_traces(textinfo="label+percent",
                                       textposition="inside")
                fig_pie.update_layout(paper_bgcolor="#161b22", height=380)
                st.plotly_chart(fig_pie, use_container_width=True)
            
            with col_table:
                usine_dist["Part"] = (usine_dist["Tonnes/Jour"] / 
                                      usine_dist["Tonnes/Jour"].sum() * 100).round(1)
                usine_dist["Part"] = usine_dist["Part"].astype(str) + "%"
                usine_dist["Tonnes/Jour"] = usine_dist["Tonnes/Jour"].apply(lambda x: f"{x:,}t")
                st.markdown("**Détail par usine**")
                st.dataframe(usine_dist.rename(columns={"Tonnes/Jour": "Tonnage"}),
                            use_container_width=True, hide_index=True)

    # ══════════════════════════════════════════════════════════════
    # 📅 TABLEAU JOURNALIER PAR AGRICULTEUR — vue calendrier pivot
    # ══════════════════════════════════════════════════════════════
    st.markdown("---")
    st.subheader(f"📅 Planning journalier détaillé — {selected}")
    
    # ✅ Afficher CLAIREMENT les 2 totaux : déclaré (agriculteurs) vs planifié (OR-Tools)
    _declared_tot = 0
    if not AGRI_DF.empty and "commercial" in AGRI_DF.columns:
        _sel_d = AGRI_DF[AGRI_DF["commercial"] == selected]
        _declared_tot = pd.to_numeric(_sel_d["tonnage_total"], errors="coerce").fillna(0).sum()
    _planned_tot = one["Tonnes/Jour"].sum() if not one.empty else 0
    _ecart = _planned_tot - _declared_tot
    _pct   = (_ecart/_declared_tot*100) if _declared_tot > 0 else 0
    
    cd1, cd2, cd3 = st.columns(3)
    cd1.metric("📋 Tonnage DÉCLARÉ (saison)", f"{_declared_tot:,.0f}t",
               help="Tonnage total que le commercial a déclaré dans son fichier d'upload")
    cd2.metric("📊 Tonnage PLANIFIÉ (OR-Tools)", f"{_planned_tot:,.0f}t",
               delta=f"{_pct:+.1f}% vs déclaré",
               help="Tonnage qu'OR-Tools a placé dans le calendrier (~95% du déclaré, tolérance ±5%)")
    cd3.metric("Écart placement", f"{_ecart:+,.0f}t",
               help="Écart entre déclaré et planifié — normal jusqu'à -5% (tolérance OR-Tools)")
    
    st.caption("📋 Déclaré = ce que le commercial a annoncé | 📊 Planifié = ce qu'OR-Tools place dans le calendrier "
               "(la différence est normale jusqu'à 5% — tolérance solveur)")
    st.markdown(" ")
    st.caption("⬇️ Le tableau ci-dessous montre le tonnage **planifié** journalier par agriculteur "
               "(0 = pas de livraison ce jour)")

    if not one.empty and "Agriculteur" in one.columns and "Tonnes/Jour" in one.columns:

        # Construire le pivot: lignes=Agriculteur, colonnes=Date
        one_pivot_src = one[["Date","Agriculteur","Tonnes/Jour","Usine"]].copy()
        one_pivot_src["Date_str"] = one_pivot_src["Date"].dt.strftime("%d/%m")

        # Aggregate (un agriculteur peut livrer à plusieurs usines le même jour)
        piv = one_pivot_src.groupby(["Agriculteur","Date_str"])["Tonnes/Jour"].sum().reset_index()
        piv["Tonnes/Jour"] = piv["Tonnes/Jour"].round(0).astype(int)

        # Pivot table
        piv_table = piv.pivot(index="Agriculteur", columns="Date_str", values="Tonnes/Jour").fillna(0).astype(int)

        # Trier les colonnes par ordre chronologique
        try:
            all_dates_sorted = sorted(
                one_pivot_src["Date"].dt.normalize().unique()
            )
            date_str_ordered = [d.strftime("%d/%m") for d in all_dates_sorted
                                 if d.strftime("%d/%m") in piv_table.columns]
            # Garder uniquement les colonnes existantes dans l'ordre
            piv_table = piv_table[[c for c in date_str_ordered if c in piv_table.columns]]
        except Exception:
            pass

        # Ajouter colonne TOTAL
        piv_table.insert(0, "TOTAL (t)", piv_table.sum(axis=1))
        piv_table = piv_table.sort_values("TOTAL (t)", ascending=False)

        # Ajouter ligne TOTAL
        totals_row = piv_table.sum(axis=0)
        totals_row.name = "── TOTAL JOUR ──"
        piv_table = pd.concat([piv_table, totals_row.to_frame().T])

        # Options d'affichage
        col_opt1, col_opt2 = st.columns([2, 1])
        with col_opt1:
            n_days_total = len([c for c in piv_table.columns if c != "TOTAL (t)"])
            n_farmers_piv = len(piv_table) - 1  # -1 pour la ligne TOTAL
            st.caption(f"📊 {n_farmers_piv} agriculteurs × {n_days_total} jours | "
                       f"Période: {one['Date'].min().strftime('%d/%m/%Y')} → {one['Date'].max().strftime('%d/%m/%Y')}")
        with col_opt2:
            # Filtre période rapide
            show_pic_only = st.checkbox("⚡ PIC uniquement (1-15 Jul)", key=f"piv_pic_{selected}")

        if show_pic_only:
            # Garder seulement les colonnes PIC
            import re as _re
            def _is_pic_col(col):
                if col == "TOTAL (t)": return True
                try:
                    d, m = col.split("/")
                    return int(m) == 7 and 1 <= int(d) <= 15
                except Exception:
                    return False
            piv_table = piv_table[[c for c in piv_table.columns if _is_pic_col(c)]]

        # Affichage avec couleur selon valeur (0=gris, >0=dégradé vert)
        def _color_cell(val):
            if isinstance(val, str):
                return "background-color: #1a1a2e; color: #888; font-weight: bold"
            if val == 0:
                return "background-color: #0d1117; color: #444"
            intensity = min(1.0, val / 300)  # normalise à 300t max
            r = int(14  + (0   - 14)  * intensity)
            g = int(149 + (229 - 149) * intensity)
            b = int(160 + (160 - 160) * intensity)
            return f"background-color: rgb({r},{g},{b}); color: {'white' if intensity > 0.4 else '#0d1117'}; font-weight: {'bold' if val > 50 else 'normal'}"

        try:
            styled = piv_table.style.applymap(_color_cell)
            st.dataframe(styled, use_container_width=True,
                         height=min(900, max(300, (n_farmers_piv + 2) * 35 + 40)))
        except Exception:
            st.dataframe(piv_table, use_container_width=True,
                         height=min(900, max(300, (n_farmers_piv + 2) * 35 + 40)))

        # Export planning journalier
        st.download_button(
            f"⬇️ Exporter planning journalier {selected} (CSV)",
            data=piv_table.to_csv(index=True),
            file_name=f"planning_{selected.replace(' ','_')}_journalier.csv",
            mime="text/csv",
        )
        
        # ── 🚛 NOUVEAU: Export TRANSPORT du commercial ─────────────────
        st.markdown("---")
        st.subheader(f"🚛 Transport & Voyages — {selected}")
        
        if "Véhicules Requis" in one.columns or "Type Véhicule" in one.columns:
            # Construire le tableau transport: agriculteur × date × véhicules × voyages
            transp_cols = [c for c in ["Date","Agriculteur","Usine","Tonnes/Jour",
                                       "Type Véhicule","Véhicules Requis","Nb Voyages"] 
                          if c in one.columns]
            df_transport = one[transp_cols].sort_values(["Date","Agriculteur"]).reset_index(drop=True)
            df_transport["Date"] = pd.to_datetime(df_transport["Date"]).dt.strftime("%d/%m/%Y")
            
            # Affichage tableau
            st.dataframe(df_transport, use_container_width=True, height=400, hide_index=True)
            
            # Stats résumées
            col_t1, col_t2, col_t3, col_t4 = st.columns(4)
            total_voyages = int(df_transport["Nb Voyages"].sum()) if "Nb Voyages" in df_transport.columns else 0
            total_jours   = df_transport["Date"].nunique()
            avg_per_day   = total_voyages / total_jours if total_jours > 0 else 0
            
            col_t1.metric("Total voyages saison", f"{total_voyages:,}")
            col_t2.metric("Jours actifs", total_jours)
            col_t3.metric("Voyages/jour (moy)", f"{avg_per_day:.1f}")
            
            # Comptage par type véhicule
            if "Type Véhicule" in df_transport.columns:
                veh_count = df_transport["Type Véhicule"].value_counts().to_dict()
                veh_str = " | ".join(f"{v}: {c}" for v, c in veh_count.items())
                col_t4.metric("Répartition véhicules", veh_str[:50])
            
            # Bouton export Excel (formaté)
            import io
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                df_transport.to_excel(writer, sheet_name="Transport_jour", index=False)
                # Résumé par jour
                daily_summary = one.groupby(one["Date"].dt.date if "Date" in one.columns else "Date").agg({
                    "Tonnes/Jour": "sum",
                    "Nb Voyages": "sum" if "Nb Voyages" in one.columns else "count",
                    "Agriculteur": "nunique",
                }).reset_index()
                daily_summary.columns = ["Date", "Total tonnes", "Total voyages", "Nb agriculteurs"]
                daily_summary.to_excel(writer, sheet_name="Resume_par_jour", index=False)
            buffer.seek(0)
            
            st.download_button(
                f"⬇️ Exporter Transport & Voyages {selected} (Excel)",
                data=buffer,
                file_name=f"transport_{selected.replace(' ','_')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
            )
            
            # Aussi CSV simple
            st.download_button(
                f"⬇️ Exporter Transport {selected} (CSV)",
                data=df_transport.to_csv(index=False).encode("utf-8-sig"),
                file_name=f"transport_{selected.replace(' ','_')}.csv",
                mime="text/csv",
            )
        else:
            st.info("Colonnes véhicules non disponibles dans le planning.")
    else:
        st.info("Aucune donnée disponible pour ce commercial.")

# ── TAB 3: PAR USINE ─────────────────────────────────────────
with tab3:
    if "Usine" not in p.columns:
        st.info("Colonne 'Usine' absente du planning.")
    else:
        # For usine role: show only data for their factory
        # Header message for usine users
        if CURRENT_ROLE == "usine":
            st.info(f"🏭 Vue **{CURRENT_NAME}** — Tonnage et voyages vous concernant.")

        factories = sorted(p["Usine"].dropna().unique())

        # Cards — total per factory (usine sees only theirs)
        if CURRENT_ROLE != "usine":
            cols_f = st.columns(len(factories))
            for i, f in enumerate(factories):
                # Use declared tonnage from agriculteurs table (more accurate)
                ft = GLOBAL_USINE_TONS.get(f, p[p["Usine"]==f]["Tonnes/Jour"].sum())
                with cols_f[i]:
                    st.metric(f, f"{ft:,.0f} t")

        # Line chart per factory
        fact_daily = p.groupby(["Date","Usine"])["Tonnes/Jour"].sum().reset_index()
        fig9 = px.line(
            fact_daily, x="Date", y="Tonnes/Jour", color="Usine",
            color_discrete_map=FACTORY_COLORS,
            title="Tonnes/jour reçues" + (f" — {CURRENT_NAME}" if CURRENT_ROLE == "usine" else " par usine"),
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
            peak_fact = (fact_daily.groupby("Usine")["Tonnes/Jour"]
                         .max().reset_index().rename(columns={"Tonnes/Jour":"Pic/Jour"}))
            fig10 = px.bar(
                peak_fact, x="Usine", y="Pic/Jour",
                color="Usine", color_discrete_map=FACTORY_COLORS,
                title="Pic journalier max", template="plotly_dark",
                text_auto=".3s",
            )
            fig10.update_layout(paper_bgcolor="#161b22", showlegend=False, height=300)
            st.plotly_chart(fig10, use_container_width=True)
        with c2:
            # Use declared tonnage from agriculteurs (correct totals per usine)
            if GLOBAL_USINE_TONS:
                fact_tot = pd.DataFrame(list(GLOBAL_USINE_TONS.items()),
                                         columns=["Usine","Tonnes/Jour"])
            else:
                fact_tot = p.groupby("Usine")["Tonnes/Jour"].sum().reset_index()
            fig11 = px.bar(
                fact_tot, x="Tonnes/Jour", y="Usine", orientation="h",
                color="Usine", color_discrete_map=FACTORY_COLORS,
                title="Total tonnes déclarées (saison)", template="plotly_dark",
                text_auto=".3s",
            )
            fig11.update_layout(paper_bgcolor="#161b22", showlegend=False, height=300)
            st.plotly_chart(fig11, use_container_width=True)

        # Drill-down
        st.subheader("📋 Détail journalier")
        if CURRENT_ROLE == "usine":
            # Usine sees: date, total tonnes/jour, nb voyages — NO commercial column
            fact_one = p[p["Usine"] == CURRENT_NAME].sort_values("Date")
            daily_usine = (fact_one.groupby("Date")
                           .agg(Tonnes_Jour=("Tonnes/Jour","sum"),
                                Nb_Voyages=("Nb Voyages","sum"))
                           .reset_index())
            daily_usine["Pic"] = daily_usine["Date"].apply(
                lambda d: "⚡ PIC" if PEAK_START <= d.date() <= PEAK_END else ""
            )
            st.dataframe(daily_usine, use_container_width=True, height=300)
            st.download_button(
                "⬇️ Exporter (CSV)",
                data=df_to_csv(daily_usine),
                file_name=f"livraisons_{CURRENT_NAME}.csv",
                mime="text/csv",
            )
        else:
            sel_fact = st.selectbox("Choisir une usine", factories)
            fact_one = p[p["Usine"]==sel_fact].sort_values("Date")
            # Directeur/commercial sees commercial column too
            show_cols_usine = [c for c in ["Date","Commercial","Agriculteur","Tonnes/Jour",
                                            "Nb Voyages","Pic de Récolte"] if c in fact_one.columns]
            st.dataframe(
                fact_one[show_cols_usine].reset_index(drop=True),
                use_container_width=True, height=260,
            )

# ── TAB 4: TRANSPORT & ALERTES ───────────────────────────────
with tab4:
    # ALL 4 vehicle types — TRACTEUR shown even if 0
    ALL_VEH_COLS = ["Voyages TRACTEUR","Voyages PETIT POILOUR","Voyages POILOUR","Voyages SEMI"]
    VEH_COLORS   = {
        "TRACTEUR":      "#a16207",   # Tracteur caisses (COMOCAP)
        "PETIT POILOUR": "#f5a623",   # PPL (7-12t) = Petit Poilour
        "POILOUR":       "#3b82f6",   # PL  (13-25t) = Poilour
        "SEMI":          "#00e5a0",   # Semi (25-40t)
    }
    VEH_DISPLAY = {
        "TRACTEUR":      "TRACTEUR (caisses)",
        "PETIT POILOUR": "PPL / Petit Poilour (7-12t)",
        "POILOUR":       "PL / Poilour (13-25t)",
        "SEMI":          "SEMI (25-40t)",
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

    # ── Vehicle count summary ──
    st.subheader("📦 Total voyages par type — saison complète")
    veh_totals_all = {
        vc.replace("Voyages ",""): int(t[vc].sum())
        for vc in ALL_VEH_COLS
    }
    v1, v2, v3, v4 = st.columns(4)
    v1.metric("🚜 TRACTEUR",       f"{veh_totals_all.get('TRACTEUR',0):,} voyages")
    v2.metric("🚛 PETIT POILOUR",  f"{veh_totals_all.get('PETIT POILOUR',0):,} voyages")
    v3.metric("🚚 POILOUR",        f"{veh_totals_all.get('POILOUR',0):,} voyages")
    v4.metric("🚜 SEMI",           f"{veh_totals_all.get('SEMI',0):,} voyages")
    st.markdown("---")

    c1, c2 = st.columns(2)
    with c1:
        # Total unique vehicle counts needed for the season
        st.subheader("🚛 Véhicules nécessaires — saison")
        # Rotations par jour par type de véhicule (basé sur distances moyennes Tunisia)
        # SEMI: 3-5 rot/j (longs trajets KAIROUAN/GAFSA), PPL/PL: 5-7 rot/j (courts)
        ROT_PAR_VEH = {
            "TRACTEUR":      6,   # 6 rotations/j (courte distance, ferme locale)
            "PETIT POILOUR": 5,   # 5 rotations/j (moyen rayon)
            "POILOUR":       5,   # 5 rotations/j (moyen rayon)
            "SEMI":          4,   # 4 rotations/j (longs trajets RM + GAFSA)
        }
        fleet_inventory = {
            "TRACTEUR":      fl_trac,
            "PETIT POILOUR": fl_ppl,
            "POILOUR":       fl_pl,
            "SEMI":          fl_semi,
        }
        # Véhicules nécessaires = ceil(voyages_pic_journée / rotations_par_véhicule)
        import math
        for vc in ALL_VEH_COLS:
            vname = vc.replace("Voyages ", "")
            rot   = ROT_PAR_VEH.get(vname, 5)
            total_voyages = int(t[vc].sum())
            # Peak = max sur la période 1-15 juillet uniquement (cap s'applique là)
            t_peak = t[(t["Date"].dt.date >= PEAK_START) & (t["Date"].dt.date <= PEAK_END)]
            peak_day = int(t_peak[vc].max()) if not t_peak.empty and vc in t_peak.columns else 0
            if peak_day == 0 and not t.empty:
                peak_day = int(t[vc].max())  # fallback sur toute la période
            needed = math.ceil(peak_day / rot) if peak_day > 0 else 0
            owned  = fleet_inventory[vname]
            if total_voyages == 0:
                st.info(f"⚪ **{vname}** : non utilisé cette saison")
            elif needed > owned:
                st.error(f"🔴 **{vname}** : besoin **{needed} véhicules** × {rot} rot/j (pic={peak_day} voyages) | vous avez {owned} → manque {needed-owned}")
            else:
                st.success(f"✅ **{vname}** : besoin **{needed} véhicules** × {rot} rot/j (pic={peak_day} voyages) | vous avez {owned} → marge {owned-needed}")

    with c2:
        # Bar chart total voyages by type
        veh_totals_all = {vc.replace("Voyages ",""): int(t[vc].sum()) for vc in ALL_VEH_COLS}
        fig13 = px.bar(
            x=list(veh_totals_all.values()),
            y=list(veh_totals_all.keys()),
            orientation="h",
            color=list(veh_totals_all.keys()),
            color_discrete_map={k: VEH_COLORS[k] for k in VEH_COLORS},
            title="Total voyages saison",
            template="plotly_dark",
            text_auto=True,
        )
        fig13.update_layout(paper_bgcolor="#161b22", showlegend=False, height=300)
        st.plotly_chart(fig13, use_container_width=True)

    # ──── NOUVEAU : Tableau récapitulatif Transport par USINE ────
    st.markdown("---")
    st.subheader("🚛 Disponibilité Transport par Usine (caps PIC 1-15 Jul)")
    
    # Caps & transport confirmé (référence: transport_12_mai.xlsx)
    # Caps officiels + transport confirmé + joker + tracteur COMOCAP
    # COMOCAP : 100t/j en TRACTEUR (~10 voyages × 10t) en plus du transport confirmé
    TRANSPORT_DATA_USINE = {
        "SICAM":    {"cap": 1300, "conf": 1199, "bennes_conf": 58, "joker": 0,  "joker_bennes": 0, "tracteur": 0},
        "TUCAL":    {"cap": 750,  "conf": 348,  "bennes_conf": 19, "joker": 76, "joker_bennes": 4, "tracteur": 0},
        "COMOCAP":  {"cap": 700,  "conf": 298,  "bennes_conf": 21, "joker": 84, "joker_bennes": 6, "tracteur": 100},
        "ABIDA":    {"cap": 150,  "conf": 50,   "bennes_conf": 2,  "joker": 0,  "joker_bennes": 0, "tracteur": 0},
        "ELFALLEH": {"cap": 100,  "conf": 24,   "bennes_conf": 2,  "joker": 0,  "joker_bennes": 0, "tracteur": 0},
    }
    
    transport_summary = []
    for usine, data in TRANSPORT_DATA_USINE.items():
        cap_official = data["cap"]
        confirm_tons = data["conf"]
        bennes_conf  = data["bennes_conf"]
        joker_tons   = data["joker"]
        joker_bennes = data["joker_bennes"]
        tracteur_tons = data.get("tracteur", 0)
        total_dispo  = confirm_tons + joker_tons + tracteur_tons
        manque       = max(0, cap_official - total_dispo)
        # Estimation bennes nécessaires pour combler le manque (~25t/benne moyen)
        bennes_manquantes = math.ceil(manque / 25) if manque > 0 else 0
        couverture = round(total_dispo / cap_official * 100, 1)
        
        statut = "✅ Complet" if couverture >= 100 else (
                 "⚠️ Limité" if couverture >= 50 else "🔴 Critique")
        
        transport_summary.append({
            "Usine":            usine,
            "Cap officiel":     f"{cap_official}t/j",
            "Confirmé":         f"{confirm_tons}t/j",
            "Bennes confirmées":bennes_conf,
            "Joker":            f"{joker_tons}t/j" if joker_tons > 0 else "—",
            "Tracteur":         f"{tracteur_tons}t/j (~{tracteur_tons//10} voyages)" if tracteur_tons > 0 else "—",
            "Total disponible": f"{total_dispo}t/j",
            "Manque":           f"{manque}t/j" if manque > 0 else "—",
            "Bennes nécessaires en plus": bennes_manquantes if bennes_manquantes > 0 else "—",
            "Couverture":       f"{couverture}%",
            "Statut":           statut,
        })
    
    df_ts = pd.DataFrame(transport_summary)
    st.dataframe(df_ts, use_container_width=True, hide_index=True)
    
    # Stats globales
    total_cap    = sum(d["cap"] for d in TRANSPORT_DATA_USINE.values())
    total_dispo  = sum(d["conf"] + d["joker"] + d.get("tracteur", 0) for d in TRANSPORT_DATA_USINE.values())
    total_bennes = sum(d["bennes_conf"] + d["joker_bennes"] for d in TRANSPORT_DATA_USINE.values())
    total_manque = total_cap - total_dispo
    
    cs1, cs2, cs3, cs4 = st.columns(4)
    cs1.metric("Cap officiel total",      f"{total_cap:,} t/j")
    cs2.metric("Transport disponible",    f"{total_dispo:,} t/j",
               delta=f"{round(total_dispo/total_cap*100,1)}% couvert")
    cs3.metric("Bennes confirmées",       f"{total_bennes}")
    cs4.metric("Manque global",           f"{total_manque:,} t/j",
               delta=f"~{math.ceil(total_manque/25)} bennes manquantes",
               delta_color="inverse")
    
    st.caption("📌 Joker = camions polyvalents (BOURAK pour TUCAL, LUI-MEME pour COMOCAP). "
               "TRACTEUR = vrai transport 10t/voyage (uniquement COMOCAP, ~100t/j = 10 voyages). "
               "Source: transport_12_mai.xlsx (confirmé). Manque = caps officiel - transport disponible.")
    
    st.markdown("---")
    
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

# ── SECTION NÉCESSITÉ TRANSPORT (fin de tab4) ───────────────────
    st.markdown("---")
    st.subheader("📦 Disponibilité Transport & Besoin Restant par Usine")
    st.caption("Comparaison flotte propre disponible vs tonnage max planifié pendant le PIC")
    if _fleet_from_file if '_fleet_from_file' in dir() else False:
        st.success("✅ Flotte chargée depuis transport_disponible.xlsx")
    else:
        st.warning("⚠️ transport_disponible.xlsx non trouvé — valeurs par défaut utilisées. "
                   "Ajoutez le fichier dans le repo GitHub.")

    # ── Flotte propre : lecture depuis transport_disponible.xlsx ──
    @st.cache_data(ttl=300)
    def _load_fleet():
        import os
        FALLBACK = {
            "SICAM":    {"SEMI":330,"PL":825,"PPL": 44,"TRACTEUR":  0,"BOURAK":  0,"total":1199,"nb_bennes": 58},
            "COMOCAP":  {"SEMI": 90,"PL": 76,"PPL":132,"TRACTEUR":100,"BOURAK":  0,"total": 398,"nb_bennes": 21},
            "TUCAL":    {"SEMI": 30,"PL":318,"PPL":  0,"TRACTEUR":  0,"BOURAK": 76,"total": 424,"nb_bennes": 19},
            "ABIDA":    {"SEMI": 30,"PL": 20,"PPL":  0,"TRACTEUR":  0,"BOURAK":  0,"total":  50,"nb_bennes":  2},
            "ELFALLEH": {"SEMI":  0,"PL":  0,"PPL": 24,"TRACTEUR":  0,"BOURAK":  0,"total":  24,"nb_bennes":  2},
        }
        paths = ["transport_disponible.xlsx",
                 os.path.join(os.path.dirname(__file__), "transport_disponible.xlsx")]
        df_t = None
        for p in paths:
            if os.path.exists(p):
                try:
                    df_t = pd.read_excel(p, sheet_name=0)
                    break
                except Exception:
                    pass
        if df_t is None:
            return FALLBACK, False
        try:
            # Auto-détecter colonnes (case-insensitive)
            col_lower = {str(c).strip().lower(): str(c).strip() for c in df_t.columns}
            usine_col = col_lower.get("usine")
            ton_col   = col_lower.get("tonnage")
            type_col  = next((col_lower[k] for k in col_lower
                              if "type" in k and "vehicule" in k), None)
            conf_col  = col_lower.get("confirmation", col_lower.get("actif"))
            if not (usine_col and ton_col and type_col):
                return FALLBACK, False
            ALIASES = {"EL FALLEH":"ELFALLEH","FELLEH":"ELFALLEH","FELLA":"ELFALLEH",
                       "LUI-MEME":"LUIMEME","LUI-MÊME":"LUIMEME","TOTAL":"SKIP"}
            df_t["_u"] = df_t[usine_col].astype(str).str.strip().str.upper().map(
                lambda x: ALIASES.get(x, x))
            df_t["_t"] = pd.to_numeric(df_t[ton_col], errors="coerce")
            def _vt(t):
                t = str(t).strip().upper()
                if "SEMI" in t or "DOUBLE" in t: return "SEMI"
                if t.startswith("PPL"):           return "PPL"
                if t.startswith("PL"):            return "PL"
                return t
            df_t["_v"] = df_t[type_col].apply(_vt)
            df_t["_a"] = df_t[conf_col].astype(str).str.strip().str.lower() \
                         if conf_col else pd.Series("ok", index=df_t.index)
            df_ok = df_t[(df_t["_a"]=="ok") & df_t["_t"].notna() &
                         (df_t["_t"]>0) & (df_t["_u"]!="SKIP")]
            bour = int(df_ok[df_ok["_u"]=="BOURAK"]["_t"].sum())
            fleet = {}
            for usine in ["SICAM","COMOCAP","TUCAL","ABIDA","ELFALLEH"]:
                sub  = df_ok[df_ok["_u"]==usine]
                semi = int(sub[sub["_v"]=="SEMI"]["_t"].sum())
                pl   = int(sub[sub["_v"]=="PL"]["_t"].sum())
                ppl  = int(sub[sub["_v"]=="PPL"]["_t"].sum())
                trac = 100 if usine=="COMOCAP" else 0
                b    = bour if usine=="TUCAL" else 0
                fleet[usine] = {"SEMI":semi,"PL":pl,"PPL":ppl,
                                "TRACTEUR":trac,"BOURAK":b,
                                "total":semi+pl+ppl+trac+b,"nb_bennes":len(sub)}
            return fleet, True
        except Exception as e:
            return FALLBACK, False
    
    FLEET_DISPO, _fleet_from_file = _load_fleet()
    CAP_OFFICIEL = {"SICAM":1300,"COMOCAP":700,"TUCAL":750,"ABIDA":150,"ELFALLEH":100}
    CAP_VEH      = {"SEMI":(27,33),"PL":(15,25),"PPL":(6,14),"TRACTEUR":(9,11)}

    # Calculer le max planifié par usine depuis le planning chargé
    if not p.empty and "Usine" in p.columns:
        max_par_usine = (p.groupby(["Date","Usine"])["Tonnes/Jour"]
                         .sum().reset_index()
                         .groupby("Usine")["Tonnes/Jour"].max()
                         .to_dict())
    else:
        max_par_usine = {u: CAP_OFFICIEL[u] for u in CAP_OFFICIEL}

    # ── Tableau résumé ───────────────────────────────────────────────
    import plotly.graph_objects as go

    rows_nec = []
    for usine, cap_off in CAP_OFFICIEL.items():
        fleet    = FLEET_DISPO.get(usine, {})
        f_total  = fleet.get("total", 0)
        max_plan = max_par_usine.get(usine, cap_off)
        besoin   = max(0, max_plan - f_total)
        pct      = round(f_total / max_plan * 100, 1) if max_plan > 0 else 0

        # Estimer le besoin par type de véhicule
        rem = besoin
        b_semi = max(0, round((rem * 0.5) / 30)) if usine in ("SICAM","TUCAL") else 0
        rem -= b_semi * 30
        b_pl   = max(0, round((rem * 0.6) / 20)) if usine in ("SICAM","TUCAL","COMOCAP","ABIDA") else 0
        rem -= b_pl * 20
        b_ppl  = max(0, round(rem / 10))

        rows_nec.append({
            "Usine":          usine,
            "Cap. Officielle": cap_off,
            "Flotte propre":  f_total,
            "dont TRACTEUR":  fleet.get("TRACTEUR",0),
            "dont BOURAK":    fleet.get("BOURAK",0),
            "dont SEMI":      fleet.get("SEMI",0),
            "dont PL":        fleet.get("PL",0),
            "dont PPL":       fleet.get("PPL",0),
            "Max planifié":   int(max_plan),
            "Besoin restant": int(besoin),
            "Couverture %":   pct,
            "→ SEMI louer":   b_semi,
            "→ PL louer":     b_pl,
            "→ PPL louer":    b_ppl,
        })

    df_nec = pd.DataFrame(rows_nec)

    # ── 5 cartes KPI par usine ────────────────────────────────────
    cols_usine = st.columns(5)
    for i, row in df_nec.iterrows():
        usine = row["Usine"]
        besoin = row["Besoin restant"]
        pct    = row["Couverture %"]
        color  = "#1E8449" if pct >= 80 else ("#F39C12" if pct >= 50 else "#C0392B")
        with cols_usine[i]:
            st.markdown(f"""
            <div style="background:#1a2332;border-radius:8px;padding:10px;text-align:center;
                        border-left:4px solid {color};">
              <div style="font-size:13px;font-weight:bold;color:#ccc">{usine}</div>
              <div style="font-size:22px;font-weight:bold;color:{color}">{pct:.0f}%</div>
              <div style="font-size:11px;color:#aaa">couvert</div>
              <div style="font-size:12px;color:#e87;">{int(row['Flotte propre'])}t / {int(row['Max planifié'])}t</div>
              <div style="font-size:11px;color:#f66;font-weight:bold">
                {'⚠️ +' + str(besoin) + 't externe' if besoin > 0 else '✅ suffisant'}</div>
            </div>""", unsafe_allow_html=True)

    st.markdown(" ")

    # ── Tableau détaillé ──────────────────────────────────────────
    st.markdown("**Détail par usine**")
    col_disp = ["Usine","Flotte propre","dont TRACTEUR","dont BOURAK","dont SEMI","dont PL","dont PPL",
                "Max planifié","Besoin restant","Couverture %","→ SEMI louer","→ PL louer","→ PPL louer"]
    
    # Affichage simple sans style (compatible toutes versions pandas)
    df_show = df_nec[col_disp].copy()
    df_show["Couverture %"] = df_show["Couverture %"].apply(lambda x: f"{x:.0f}%")
    df_show["Besoin restant"] = df_show["Besoin restant"].apply(
        lambda x: f"✅ 0t" if x == 0 else f"⚠️ +{x}t")
    st.dataframe(df_show, use_container_width=True, hide_index=True, height=220)

    # ── Graphique barres groupées ──────────────────────────────────
    fig_nec = go.Figure()
    fig_nec.add_trace(go.Bar(
        name="Flotte propre (disponible)",
        x=df_nec["Usine"], y=df_nec["Flotte propre"],
        marker_color="#2E86C1", text=df_nec["Flotte propre"].astype(str)+"t",
        textposition="inside",
    ))
    fig_nec.add_trace(go.Bar(
        name="Besoin restant (à louer)",
        x=df_nec["Usine"], y=df_nec["Besoin restant"],
        marker_color="#E74C3C", text=df_nec["Besoin restant"].apply(lambda x: f"+{x}t" if x>0 else "✅"),
        textposition="inside",
    ))
    # Ligne cap officielle
    fig_nec.add_trace(go.Scatter(
        name="Capacité officielle (PIC)",
        x=df_nec["Usine"], y=df_nec["Cap. Officielle"],
        mode="markers+lines", marker_symbol="diamond",
        marker_size=10, line_dash="dash",
        marker_color="#F39C12", line_color="#F39C12",
    ))
    fig_nec.update_layout(
        barmode="stack",
        title="Transport disponible vs besoin par usine (PIC)",
        template="plotly_dark",
        paper_bgcolor="#161b22",
        plot_bgcolor="#0d1117",
        height=380,
        legend=dict(orientation="h", y=-0.2),
        yaxis_title="Tonnes/jour",
    )
    st.plotly_chart(fig_nec, use_container_width=True)

    # ── Export ────────────────────────────────────────────────────
    st.download_button(
        "⬇️ Exporter rapport transport nécessaire (CSV)",
        data=df_nec.to_csv(index=False).encode("utf-8-sig"),
        file_name="transport_necessaire.csv",
        mime="text/csv",
    )

# ── TAB 5: DÉCALAGE & OPTIMISATION ──────────────────────────
with tab5:

    # Explain clearly why these are empty
    st.success("✅ **OR-Tools a optimisé le planning automatiquement** — aucun décalage manuel ni double transport nécessaire.")

    st.markdown("""
    > Avec l'optimiseur OR-Tools, tous les conflits de transport sont résolus **pendant le calcul**.
    > L'algorithme distribue les tonnages sur les bons jours dès le départ,
    > respectant simultanément tous les caps. Il n'y a donc pas de "décalage après coup".
    """)

    st.divider()

    # ── Instead of empty charts: show constraint verification ──
    st.subheader("📊 Vérification des contraintes — résultat optimizer")

    # Commercial caps verification
    COMMERCIAL_CAPS = {
        "FEDI": 850, "MAKKI BEN SALAH": 800, "KHALIL": 800,
        "ACHREF AJLANI": 500, "JILANI OBAY": 50,
    }
    FACTORY_CAPS = {
        "SICAM": 1300, "TUCAL": 750, "COMOCAP": 700,
        "ABIDA": 150, "ELFALLEH": 100,
    }
    # Transport confirmé réel (source: transport_12_mai.xlsx — vérifié)
    # SICAM: 4 PPL(44t) + 43 PL(825t) + 11 SEMI(330t) = 1199t/j / 58 bennes
    # TUCAL: 18 PL(318t) + 1 SEMI(30t) = 348t/j / 19 bennes
    # COMOCAP: 13 PPL(132t) + 5 PL(76t) + 3 SEMI(90t) = 298t/j / 21 bennes
    # ABIDA: 1 PL(20t) + 1 SEMI(30t) = 50t/j / 2 bennes
    # ELFALLEH: 2 PPL(24t) = 24t/j / 2 bennes
    TRANSPORT_CONF = {
        "SICAM": 1199, "TUCAL": 348, "COMOCAP": 298,
        "ABIDA": 50,   "ELFALLEH": 24,
    }
    # Jokers: BOURAK=76t(PL) + LUI-MEME=84t(PPL+PL) = 160t total
    # BOURAK (PL, 4 bennes) → utilisable pour TUCAL principalement (zone CAP BON/NORD)
    # LUI-MEME (PPL+PL, 6 bennes) → utilisable pour COMOCAP principalement
    JOKER_ALLOC_DASH = {
        "TUCAL":    76,   # BOURAK 76t → renforce TUCAL
        "COMOCAP":  84,   # LUI-MEME 84t → renforce COMOCAP
        "ELFALLEH": 0,
        "SICAM":    0,
        "ABIDA":    0,
    }
    # Cap réel = min(cap_théorique, transport_confirmé + jokers)
    FACTORY_REAL_CAP = {
        u: min(FACTORY_CAPS[u], TRANSPORT_CONF.get(u, FACTORY_CAPS[u]) + JOKER_ALLOC_DASH.get(u, 0))
        for u in FACTORY_CAPS
    }
    # Manque de bennes (info)
    MANQUE_BENNES = {
        u: max(0, FACTORY_CAPS[u] - (TRANSPORT_CONF.get(u,0) + JOKER_ALLOC_DASH.get(u,0)))
        for u in FACTORY_CAPS
    }

    if not p.empty:
        c1, c2 = st.columns(2)

        with c1:
            st.markdown("**Caps commerciaux — max pendant PIC (1-15 Juillet)**")
            st.caption("Les caps ne s'appliquent que du 1 au 15 juillet (JILANI/KHALIL: 1-12 juil)")
            # Filtrer uniquement la période de pic pour la vérification des caps
            p_pic = p[(p["Date"].dt.date >= PEAK_START) & (p["Date"].dt.date <= PEAK_END)]
            if p_pic.empty:
                st.warning("⚠️ Aucune donnée pour le pic 1-15 juillet — planning peut-être incomplet")
                comm_max = p.groupby(["Date","Commercial"])["Tonnes/Jour"].sum().reset_index()
            else:
                comm_max = p_pic.groupby(["Date","Commercial"])["Tonnes/Jour"].sum().reset_index()
            comm_peak = comm_max.groupby("Commercial")["Tonnes/Jour"].max().reset_index()
            # Ajouter commerciaux absents pendant pic (ex: JILANI commence Jul 23)
            _all_c  = sorted(p["Commercial"].dropna().unique())
            _exist_c = set(comm_peak["Commercial"])
            _miss_c  = [cc for cc in _all_c if cc not in _exist_c]
            if _miss_c:
                comm_peak = pd.concat([comm_peak,
                    pd.DataFrame({"Commercial":_miss_c,"Tonnes/Jour":[0]*len(_miss_c)})],
                    ignore_index=True)
            comm_peak.columns = ["Commercial", "Max réel (t/j)"]
            comm_peak["Limite (t/j)"] = comm_peak["Commercial"].map(COMMERCIAL_CAPS).fillna(800)
            comm_peak["Marge (t/j)"]  = comm_peak["Limite (t/j)"] - comm_peak["Max réel (t/j)"]
            comm_peak["Statut"]       = comm_peak["Marge (t/j)"].apply(
                lambda m: "✅ OK" if m >= 0 else "❌ DÉPASSÉ"
            )
            comm_peak["Max réel (t/j)"]  = comm_peak["Max réel (t/j)"].round(0).astype(int)
            comm_peak["Limite (t/j)"]    = comm_peak["Limite (t/j)"].astype(int)
            comm_peak["Marge (t/j)"]     = comm_peak["Marge (t/j)"].round(0).astype(int)
            st.dataframe(comm_peak, use_container_width=True, hide_index=True,
                column_config={
                    "Statut": st.column_config.TextColumn(width="small"),
                    "Max réel (t/j)": st.column_config.ProgressColumn(
                        "Max réel (t/j)", min_value=0, max_value=1600, format="%d t"),
                })

            # Bar chart: actual vs cap
            import plotly.graph_objects as go
            fig_cap = go.Figure()
            fig_cap.add_trace(go.Bar(
                name="Max réel", x=comm_peak["Commercial"],
                y=comm_peak["Max réel (t/j)"],
                marker_color="#3b82f6", text=comm_peak["Max réel (t/j)"],
                textposition="outside",
            ))
            fig_cap.add_trace(go.Bar(
                name="Limite", x=comm_peak["Commercial"],
                y=comm_peak["Limite (t/j)"],
                marker_color="rgba(255,77,28,0.3)",
                marker_line_color="#e8543a", marker_line_width=2,
            ))
            fig_cap.update_layout(
                barmode="overlay", template="plotly_dark",
                paper_bgcolor="#161b22", height=280,
                title="Max journalier réel vs limite par commercial",
                legend=dict(orientation="h", yanchor="bottom", y=1.02),
            )
            st.plotly_chart(fig_cap, use_container_width=True)

        with c2:
            st.markdown("**Caps usines — PIC 1-15 Juillet uniquement ✅**")
            if "Usine" in p.columns:
                # ⚠️ OBLIGATOIRE: filtrer sur PIC seulement !
                # Les caps s'appliquent UNIQUEMENT 1-15 juillet.
                # Hors pic, OR-Tools libère les contraintes pour placer
                # le tonnage → les dépassements hors-pic sont NORMAUX et attendus.
                p_pic_u = p[(p["Date"].dt.date >= PEAK_START) &
                            (p["Date"].dt.date <= PEAK_END)]
                src_u = p_pic_u if not p_pic_u.empty else p
                usine_max  = src_u.groupby(["Date","Usine"])["Tonnes/Jour"].sum().reset_index()
                usine_peak = usine_max.groupby("Usine")["Tonnes/Jour"].max().reset_index()
                # Ajouter les usines à 0t pendant pic (ex: celles sans livraisons pic)
                _all_u = sorted(p["Usine"].dropna().unique())
                _exist_u = set(usine_peak["Usine"])
                _miss_u  = [u for u in _all_u if u not in _exist_u]
                if _miss_u:
                    usine_peak = pd.concat([usine_peak,
                        pd.DataFrame({"Usine":_miss_u,"Tonnes/Jour":[0]*len(_miss_u)})],
                        ignore_index=True)
                usine_peak.columns = ["Usine", "Max réel PIC (t/j)"]
                usine_peak["Limite (t/j)"]       = usine_peak["Usine"].map(FACTORY_CAPS).fillna(500)
                usine_peak["Cap transport (t/j)"] = usine_peak["Usine"].map(FACTORY_REAL_CAP).fillna(500)
                usine_peak["Marge (t/j)"]  = usine_peak["Limite (t/j)"] - usine_peak["Max réel PIC (t/j)"]
                usine_peak["Statut"]       = usine_peak["Marge (t/j)"].apply(
                    lambda m: "✅ OK" if m >= 0 else "❌ DÉPASSÉ"
                )
                usine_peak["Max réel PIC (t/j)"] = usine_peak["Max réel PIC (t/j)"].round(0).astype(int)
                usine_peak["Limite (t/j)"]   = usine_peak["Limite (t/j)"].astype(int)
                usine_peak["Marge (t/j)"]    = usine_peak["Marge (t/j)"].round(0).astype(int)
                st.dataframe(usine_peak, use_container_width=True, hide_index=True)

                fig_usine = go.Figure()
                fig_usine.add_trace(go.Bar(
                    name="Max réel PIC", x=usine_peak["Usine"],
                    y=usine_peak["Max réel PIC (t/j)"],
                    marker_color="#00e5a0", text=usine_peak["Max réel PIC (t/j)"],
                    textposition="outside",
                ))
                fig_usine.add_trace(go.Bar(
                    name="Limite", x=usine_peak["Usine"],
                    y=usine_peak["Limite (t/j)"],
                    marker_color="rgba(255,77,28,0.3)",
                    marker_line_color="#e8543a", marker_line_width=2,
                ))
                fig_usine.update_layout(
                    barmode="overlay", template="plotly_dark",
                    paper_bgcolor="#161b22", height=280,
                    title="Max journalier réel vs limite par usine",
                    legend=dict(orientation="h", yanchor="bottom", y=1.02),
                )
                st.plotly_chart(fig_usine, use_container_width=True)

    st.divider()

    # ── Peak period analysis ──
    st.subheader("📅 Analyse période de pic — 1–15 Juillet")
    if not p.empty:
        peak_p = p[(p["Date"].dt.date >= PEAK_START) & (p["Date"].dt.date <= PEAK_END)]
        if not peak_p.empty:
            peak_daily = peak_p.groupby("Date")["Tonnes/Jour"].sum().reset_index()
            fig_pk = px.area(
                peak_daily, x="Date", y="Tonnes/Jour",
                title="Tonnage journalier pendant le pic (1-15 Juillet)",
                color_discrete_sequence=["#f5a623"],
                template="plotly_dark",
            )
            fig_pk.update_layout(paper_bgcolor="#161b22", height=260)
            st.plotly_chart(fig_pk, use_container_width=True)

            c1, c2, c3 = st.columns(3)
            c1.metric("Total pic", f"{peak_p['Tonnes/Jour'].sum():,.0f} t")
            c2.metric("Max en 1 jour", f"{peak_daily['Tonnes/Jour'].max():,.0f} t")
            c3.metric("Moyenne/jour", f"{peak_daily['Tonnes/Jour'].mean():,.0f} t")

    # Summary table still shown
    if not resume.empty:
        st.divider()
        st.subheader("📊 Résumé par commercial")
        st.dataframe(resume.reset_index(drop=True), use_container_width=True)
        st.download_button(
            "⬇️ Exporter résumé commercial (CSV)",
            data=df_to_csv(resume.reset_index(drop=True)),
            file_name="resume_commercial.csv",
            mime="text/csv",
        )

# ── TAB 6: HISTORIQUE 2025 vs 2026 ───────────────────────────
with tab6:
    import plotly.graph_objects as go

    # Historical data embedded — NaN-free, validated
    HIST = {"global_2025": [["06-15", 67.1], ["06-16", 146.3], ["06-17", 156.2], ["06-18", 415.4], ["06-19", 607.1], ["06-20", 816.5], ["06-21", 947.3], ["06-22", 1122.9], ["06-23", 1416.7], ["06-24", 1800.8], ["06-25", 1963.5], ["06-26", 2338.1], ["06-27", 2445.0], ["06-28", 2311.4], ["06-29", 2766.7], ["06-30", 2323.9], ["07-01", 2471.9], ["07-02", 2532.5], ["07-03", 2529.7], ["07-04", 2978.2], ["07-05", 2644.6], ["07-06", 2483.8], ["07-07", 2201.7], ["07-08", 1939.0], ["07-09", 2390.0], ["07-10", 2499.8], ["07-11", 2684.6], ["07-12", 2577.2], ["07-13", 2352.2], ["07-14", 1936.9], ["07-15", 1745.5], ["07-16", 2003.1], ["07-17", 1578.8], ["07-18", 1347.7], ["07-19", 1284.6], ["07-20", 1091.0], ["07-21", 1119.4], ["07-22", 687.1], ["07-23", 879.3], ["07-24", 850.0], ["07-25", 1006.0], ["07-26", 783.6], ["07-27", 1002.1], ["07-28", 1154.4], ["07-29", 938.4], ["07-30", 1092.3], ["07-31", 989.7], ["08-01", 990.1], ["08-02", 924.6], ["08-03", 541.3], ["08-04", 1004.2], ["08-05", 845.6], ["08-06", 582.7], ["08-07", 754.9], ["08-08", 1148.1], ["08-09", 602.1], ["08-10", 866.5], ["08-11", 738.9], ["08-12", 868.1], ["08-13", 947.9], ["08-14", 888.4], ["08-15", 863.8], ["08-16", 881.8], ["08-17", 864.5], ["08-18", 651.4], ["08-19", 552.7], ["08-20", 443.6], ["08-21", 371.5], ["08-22", 458.2], ["08-23", 327.2], ["08-24", 334.4], ["08-25", 329.6], ["08-26", 310.3], ["08-27", 308.9], ["08-28", 368.2], ["08-29", 214.3], ["08-30", 148.8], ["08-31", 179.9], ["09-01", 258.8], ["09-02", 114.4], ["09-03", 117.1], ["09-04", 163.8], ["09-05", 66.3], ["09-06", 83.6], ["09-08", 43.2], ["09-09", 90.2], ["09-10", 79.8], ["09-12", 18.1], ["09-13", 66.0], ["09-14", 33.2], ["09-15", 34.1], ["09-16", 61.1]], "sicam_2025": [["06-15", 67.1], ["06-16", 146.3], ["06-17", 156.2], ["06-18", 336.3], ["06-19", 490.3], ["06-20", 601.9], ["06-21", 686.3], ["06-22", 794.7], ["06-23", 1062.9], ["06-24", 1200.1], ["06-25", 1026.3], ["06-26", 1360.5], ["06-27", 1424.7], ["06-28", 1270.7], ["06-29", 1557.8], ["06-30", 1182.5], ["07-01", 1382.0], ["07-02", 1345.8], ["07-03", 1410.4], ["07-04", 1602.1], ["07-05", 1359.5], ["07-06", 1089.8], ["07-07", 892.3], ["07-08", 874.6], ["07-09", 993.7], ["07-10", 1125.5], ["07-11", 1383.7], ["07-12", 1317.2], ["07-13", 1119.3], ["07-14", 858.0], ["07-15", 686.0], ["07-16", 925.3], ["07-17", 751.5], ["07-18", 511.9], ["07-19", 543.4], ["07-20", 424.4], ["07-21", 259.2], ["07-22", 185.9], ["07-23", 262.6], ["07-24", 143.4], ["07-25", 462.4], ["07-26", 346.3], ["07-27", 516.4], ["07-28", 441.2], ["07-29", 376.6], ["07-30", 514.8], ["07-31", 401.4], ["08-01", 302.3], ["08-02", 349.8], ["08-03", 295.3], ["08-04", 214.8], ["08-05", 342.2], ["08-06", 236.9], ["08-07", 282.1], ["08-08", 319.2], ["08-09", 163.6], ["08-10", 238.0], ["08-11", 362.6], ["08-12", 323.3], ["08-13", 416.8], ["08-14", 358.9], ["08-15", 363.9], ["08-16", 351.1], ["08-17", 366.0], ["08-18", 182.6], ["08-19", 216.6], ["08-20", 235.4], ["08-21", 311.8], ["08-22", 305.2], ["08-23", 229.4], ["08-24", 233.8], ["08-25", 260.7], ["08-26", 263.9], ["08-27", 247.7], ["08-28", 299.2], ["08-29", 214.3], ["08-30", 98.0], ["08-31", 125.3], ["09-01", 243.0], ["09-02", 88.4], ["09-03", 102.2], ["09-04", 163.8], ["09-05", 66.3], ["09-06", 83.6], ["09-08", 43.2], ["09-09", 90.2], ["09-10", 79.8]], "autres_2025": [["06-15", 0.0], ["06-16", 0.0], ["06-17", 0.0], ["06-18", 79.1], ["06-19", 116.8], ["06-20", 214.7], ["06-21", 261.0], ["06-22", 328.2], ["06-23", 353.8], ["06-24", 600.7], ["06-25", 937.2], ["06-26", 977.6], ["06-27", 1020.2], ["06-28", 1040.7], ["06-29", 1208.9], ["06-30", 1141.4], ["07-01", 1090.0], ["07-02", 1186.8], ["07-03", 1119.3], ["07-04", 1376.1], ["07-05", 1285.1], ["07-06", 1394.0], ["07-07", 1309.4], ["07-08", 1064.3], ["07-09", 1396.3], ["07-10", 1374.2], ["07-11", 1300.9], ["07-12", 1260.0], ["07-13", 1232.9], ["07-14", 1079.0], ["07-15", 1059.5], ["07-16", 1077.7], ["07-17", 827.2], ["07-18", 835.8], ["07-19", 741.2], ["07-20", 666.6], ["07-21", 860.3], ["07-22", 501.2], ["07-23", 616.7], ["07-24", 706.6], ["07-25", 543.6], ["07-26", 437.3], ["07-27", 485.7], ["07-28", 713.2], ["07-29", 561.8], ["07-30", 577.5], ["07-31", 588.3], ["08-01", 687.8], ["08-02", 574.8], ["08-03", 246.0], ["08-04", 789.5], ["08-05", 503.4], ["08-06", 345.8], ["08-07", 472.8], ["08-08", 828.9], ["08-09", 438.5], ["08-10", 628.5], ["08-11", 376.3], ["08-12", 544.8], ["08-13", 531.1], ["08-14", 529.5], ["08-15", 499.9], ["08-16", 530.7], ["08-17", 498.5], ["08-18", 468.8], ["08-19", 336.1], ["08-20", 208.3], ["08-21", 59.6], ["08-22", 153.0], ["08-23", 97.8], ["08-24", 100.6], ["08-25", 68.9], ["08-26", 46.4], ["08-27", 61.2], ["08-28", 69.0], ["08-29", 0.0], ["08-30", 50.8], ["08-31", 54.6], ["09-01", 15.9], ["09-02", 26.0], ["09-03", 15.0], ["09-04", 0.0], ["09-05", 0.0], ["09-06", 0.0], ["09-08", 0.0], ["09-09", 0.0], ["09-10", 0.0], ["09-12", 18.1], ["09-13", 66.0], ["09-14", 33.2], ["09-15", 34.1], ["09-16", 61.1]], "stats": {"global_total": 95962.5, "global_max": 2978.2, "global_max_date": "04/07/2025", "global_avg_peak": 2397.8, "sicam_total": 47342.3, "sicam_max": 1602.1, "autres_max": 1396.3, "autres_total": 48620.2, "plan_total": 87557.0, "plan_max": 2870.0, "plan_avg_peak": 2304.3}, "plan_2026": [["06-20", 140.0], ["06-21", 200.0], ["06-22", 280.0], ["06-23", 370.0], ["06-24", 400.0], ["06-25", 655.0], ["06-26", 675.0], ["06-27", 845.0], ["06-28", 975.0], ["06-29", 1143.0], ["06-30", 1273.0], ["07-01", 1523.0], ["07-02", 1560.0], ["07-03", 1785.0], ["07-04", 1920.0], ["07-05", 2153.0], ["07-06", 2043.0], ["07-07", 2140.0], ["07-08", 2240.0], ["07-09", 2660.0], ["07-10", 2800.0], ["07-11", 2790.0], ["07-12", 2775.0], ["07-13", 2870.0], ["07-14", 2790.0], ["07-15", 2515.0], ["07-16", 2438.0], ["07-17", 2310.0], ["07-18", 2195.0], ["07-19", 2002.0], ["07-20", 1900.0], ["07-21", 1801.0], ["07-22", 1645.0], ["07-23", 1693.0], ["07-24", 1618.0], ["07-25", 1833.0], ["07-26", 1870.0], ["07-27", 1885.0], ["07-28", 2010.0], ["07-29", 2036.0], ["07-30", 1878.0], ["07-31", 1745.0], ["08-01", 1687.0], ["08-02", 1627.0], ["08-03", 1450.0], ["08-04", 1368.0], ["08-05", 1265.0], ["08-06", 1235.0], ["08-07", 1100.0], ["08-08", 873.0], ["08-09", 898.0], ["08-10", 815.0], ["08-11", 790.0], ["08-12", 630.0], ["08-13", 470.0], ["08-14", 420.0], ["08-15", 240.0], ["08-16", 190.0], ["08-17", 120.0]], "factory_2026": {"COMOCAP": [["06-29", 20.0], ["06-30", 55.0], ["07-01", 55.0], ["07-02", 60.0], ["07-03", 135.0], ["07-04", 145.0], ["07-05", 160.0], ["07-06", 165.0], ["07-07", 165.0], ["07-08", 170.0], ["07-09", 400.0], ["07-10", 520.0], ["07-11", 500.0], ["07-12", 505.0], ["07-13", 585.0], ["07-14", 600.0], ["07-15", 515.0], ["07-16", 558.0], ["07-17", 545.0], ["07-18", 580.0], ["07-19", 542.0], ["07-20", 520.0], ["07-21", 421.0], ["07-22", 380.0], ["07-23", 463.0], ["07-24", 343.0], ["07-25", 293.0], ["07-26", 275.0], ["07-27", 245.0], ["07-28", 290.0], ["07-29", 306.0], ["07-30", 293.0], ["07-31", 250.0], ["08-01", 247.0], ["08-02", 197.0], ["08-03", 160.0], ["08-04", 118.0], ["08-05", 45.0], ["08-06", 40.0], ["08-07", 15.0], ["08-08", 15.0], ["08-09", 8.0]], "SICAM": [["06-20", 100.0], ["06-21", 120.0], ["06-22", 180.0], ["06-23", 250.0], ["06-24", 250.0], ["06-25", 420.0], ["06-26", 440.0], ["06-27", 500.0], ["06-28", 560.0], ["06-29", 653.0], ["06-30", 743.0], ["07-01", 898.0], ["07-02", 895.0], ["07-03", 1000.0], ["07-04", 1105.0], ["07-05", 1128.0], ["07-06", 1213.0], ["07-07", 1285.0], ["07-08", 1340.0], ["07-09", 1490.0], ["07-10", 1530.0], ["07-11", 1540.0], ["07-12", 1510.0], ["07-13", 1495.0], ["07-14", 1440.0], ["07-15", 1350.0], ["07-16", 1280.0], ["07-17", 1210.0], ["07-18", 1140.0], ["07-19", 1015.0], ["07-20", 970.0], ["07-21", 960.0], ["07-22", 965.0], ["07-23", 930.0], ["07-24", 930.0], ["07-25", 1155.0], ["07-26", 1180.0], ["07-27", 1220.0], ["07-28", 1270.0], ["07-29", 1280.0], ["07-30", 1180.0], ["07-31", 1090.0], ["08-01", 1045.0], ["08-02", 1040.0], ["08-03", 940.0], ["08-04", 860.0], ["08-05", 850.0], ["08-06", 835.0], ["08-07", 795.0], ["08-08", 578.0], ["08-09", 600.0], ["08-10", 540.0], ["08-11", 520.0], ["08-12", 380.0], ["08-13", 260.0], ["08-14", 210.0], ["08-15", 150.0], ["08-16", 90.0], ["08-17", 30.0]], "TUCAL": [["06-20", 20.0], ["06-21", 60.0], ["06-22", 80.0], ["06-23", 80.0], ["06-24", 110.0], ["06-25", 195.0], ["06-26", 195.0], ["06-27", 245.0], ["06-28", 315.0], ["06-29", 320.0], ["06-30", 325.0], ["07-01", 330.0], ["07-02", 365.0], ["07-03", 390.0], ["07-04", 420.0], ["07-05", 415.0], ["07-06", 395.0], ["07-07", 380.0], ["07-08", 350.0], ["07-09", 370.0], ["07-10", 340.0], ["07-11", 300.0], ["07-12", 290.0], ["07-13", 320.0], ["07-14", 300.0], ["07-15", 230.0], ["07-16", 210.0], ["07-17", 205.0], ["07-18", 185.0], ["07-19", 255.0], ["07-20", 240.0], ["07-21", 240.0], ["07-22", 140.0], ["07-23", 160.0], ["07-24", 235.0], ["07-25", 315.0], ["07-26", 345.0], ["07-27", 380.0], ["07-28", 410.0], ["07-29", 410.0], ["07-30", 385.0], ["07-31", 385.0], ["08-01", 375.0], ["08-02", 370.0], ["08-03", 330.0], ["08-04", 370.0], ["08-05", 350.0], ["08-06", 340.0], ["08-07", 270.0], ["08-08", 260.0], ["08-09", 270.0], ["08-10", 255.0], ["08-11", 250.0], ["08-12", 230.0], ["08-13", 190.0], ["08-14", 190.0], ["08-15", 90.0], ["08-16", 100.0], ["08-17", 90.0]], "ELFALLEH": [["07-01", 70.0], ["07-02", 70.0], ["07-03", 90.0], ["07-04", 90.0], ["07-05", 290.0], ["07-06", 110.0], ["07-07", 110.0], ["07-08", 120.0], ["07-09", 140.0], ["07-10", 150.0], ["07-11", 190.0], ["07-12", 210.0], ["07-13", 210.0], ["07-14", 210.0], ["07-15", 220.0], ["07-16", 220.0], ["07-17", 210.0], ["07-18", 180.0], ["07-19", 110.0], ["07-20", 90.0], ["07-21", 100.0], ["07-22", 100.0], ["07-23", 100.0], ["07-24", 70.0], ["07-25", 30.0], ["07-26", 30.0], ["07-27", 20.0], ["07-28", 20.0], ["07-29", 20.0], ["07-30", 20.0], ["07-31", 20.0], ["08-01", 20.0], ["08-02", 20.0], ["08-03", 20.0], ["08-04", 20.0], ["08-05", 20.0], ["08-06", 20.0], ["08-07", 20.0], ["08-08", 20.0], ["08-09", 20.0], ["08-10", 20.0], ["08-11", 20.0], ["08-12", 20.0], ["08-13", 20.0], ["08-14", 20.0]], "ABIDA": [["06-20", 20.0], ["06-21", 20.0], ["06-22", 20.0], ["06-23", 40.0], ["06-24", 40.0], ["06-25", 40.0], ["06-26", 40.0], ["06-27", 100.0], ["06-28", 100.0], ["06-29", 150.0], ["06-30", 150.0], ["07-01", 170.0], ["07-02", 170.0], ["07-03", 170.0], ["07-04", 160.0], ["07-05", 160.0], ["07-06", 160.0], ["07-07", 200.0], ["07-08", 260.0], ["07-09", 260.0], ["07-10", 260.0], ["07-11", 260.0], ["07-12", 260.0], ["07-13", 260.0], ["07-14", 240.0], ["07-15", 200.0], ["07-16", 170.0], ["07-17", 140.0], ["07-18", 110.0], ["07-19", 80.0], ["07-20", 80.0], ["07-21", 80.0], ["07-22", 60.0], ["07-23", 40.0], ["07-24", 40.0], ["07-25", 40.0], ["07-26", 40.0], ["07-27", 20.0], ["07-28", 20.0], ["07-29", 20.0]]}}

    st.subheader("Comparaison Historique 2025 (réel) vs Plan 2026")

    # ── KPI cards ──
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        delta = round(HIST["stats"]["plan_total"] - HIST["stats"]["global_total"], 0)
        st.metric("Total saison", f"{HIST['stats']['plan_total']:,.0f} t (2026)",
                  delta=f"{delta:+,.0f} t vs 2025", delta_color="inverse")
    with c2:
        delta = round(HIST["stats"]["plan_max"] - HIST["stats"]["global_max"], 0)
        st.metric("Pic max journalier", f"{HIST['stats']['plan_max']:,.0f} t (2026)",
                  delta=f"{delta:+,.0f} t vs 2025", delta_color="inverse")
    with c3:
        delta = round(HIST["stats"]["plan_avg_peak"] - HIST["stats"]["global_avg_peak"], 0)
        st.metric("Moy. 1–15 Juillet", f"{HIST['stats']['plan_avg_peak']:,.0f} t/jour (2026)",
                  delta=f"{delta:+,.0f} t/j vs 2025", delta_color="inverse")
    with c4:
        pct = round(HIST["stats"]["plan_total"] / HIST["stats"]["global_total"] * 100 - 100, 1)
        st.metric("Variation globale", f"{pct:+.1f}%", delta="vs saison 2025",
                  delta_color="inverse")

    st.markdown("---")

    # ── Chart 1: Global overlay 2025 vs 2026 ──
    st.subheader("Tonnage journalier global — 2025 réel vs 2026 plan")
    fig1 = go.Figure()
    fig1.add_trace(go.Scatter(
        x=[f"2025-{r[0]}" for r in HIST["global_2025"]],
        y=[r[1] for r in HIST["global_2025"]],
        name="2025 Réel", line=dict(color="#8b5cf6", width=2),
        fill="tozeroy", fillcolor="rgba(139,92,246,0.08)", mode="lines",
    ))
    fig1.add_trace(go.Scatter(
        x=[f"2026-{r[0]}" for r in HIST["plan_2026"]],
        y=[r[1] for r in HIST["plan_2026"]],
        name="2026 Plan", line=dict(color="#f5a623", width=2.5),
        fill="tozeroy", fillcolor="rgba(245,166,35,0.08)", mode="lines",
    ))
    fig1.update_layout(
        template="plotly_dark", plot_bgcolor="#0d1117", paper_bgcolor="#161b22",
        height=380, hovermode="x unified",
        legend=dict(orientation="h", yanchor="bottom", y=1.02),
        yaxis=dict(title="Tonnes/jour"),
    )
    st.plotly_chart(fig1, use_container_width=True)

    # ── Charts 2 & 3 side by side ──
    c1, c2 = st.columns(2)
    with c1:
        st.subheader("SICAM — 2025 vs 2026")
        fig2 = go.Figure()
        fig2.add_trace(go.Scatter(
            x=[f"2025-{r[0]}" for r in HIST["sicam_2025"]],
            y=[r[1] for r in HIST["sicam_2025"]],
            name="SICAM 2025", line=dict(color="#8b5cf6", width=2),
            fill="tozeroy", fillcolor="rgba(139,92,246,0.1)",
        ))
        if "SICAM" in HIST["factory_2026"]:
            fig2.add_trace(go.Scatter(
                x=[f"2026-{r[0]}" for r in HIST["factory_2026"]["SICAM"]],
                y=[r[1] for r in HIST["factory_2026"]["SICAM"]],
                name="SICAM 2026", line=dict(color="#f5a623", width=2.5),
                fill="tozeroy", fillcolor="rgba(245,166,35,0.1)",
            ))
        fig2.add_hline(y=HIST["stats"]["sicam_max"],
            line_dash="dot", line_color="#e8543a", line_width=1.5,
            annotation_text=f"Cap 2025 : {HIST['stats']['sicam_max']:,.0f}t",
            annotation_position="top right")
        fig2.update_layout(template="plotly_dark", plot_bgcolor="#0d1117",
            paper_bgcolor="#161b22", height=300, hovermode="x unified",
            legend=dict(orientation="h", yanchor="bottom", y=1.02),
            yaxis=dict(title="t/jour"))
        st.plotly_chart(fig2, use_container_width=True)

    with c2:
        st.subheader("Autres usines — 2025 vs 2026")
        fig3 = go.Figure()
        # Filter out zero-only tail from autres_2025
        autres_clean = [[d,t] for d,t in HIST["autres_2025"] if t > 0]
        fig3.add_trace(go.Scatter(
            x=[f"2025-{r[0]}" for r in autres_clean],
            y=[r[1] for r in autres_clean],
            name="Autres 2025", line=dict(color="#3b82f6", width=2),
            fill="tozeroy", fillcolor="rgba(59,130,246,0.1)",
        ))
        from collections import defaultdict
        day_tots = defaultdict(float)
        for fac in ["COMOCAP","TUCAL","ABIDA","ELFALLEH"]:
            for d, t in HIST["factory_2026"].get(fac, []):
                day_tots[f"2026-{d}"] += t
        sorted_days = sorted(day_tots.items())
        if sorted_days:
            fig3.add_trace(go.Scatter(
                x=[d for d,_ in sorted_days], y=[t for _,t in sorted_days],
                name="Autres 2026", line=dict(color="#00e5a0", width=2.5),
                fill="tozeroy", fillcolor="rgba(0,229,160,0.1)",
            ))
        fig3.add_hline(y=HIST["stats"]["autres_max"],
            line_dash="dot", line_color="#e8543a", line_width=1.5,
            annotation_text=f"Cap 2025 : {HIST['stats']['autres_max']:,.0f}t",
            annotation_position="top right")
        fig3.update_layout(template="plotly_dark", plot_bgcolor="#0d1117",
            paper_bgcolor="#161b22", height=300, hovermode="x unified",
            legend=dict(orientation="h", yanchor="bottom", y=1.02),
            yaxis=dict(title="t/jour"))
        st.plotly_chart(fig3, use_container_width=True)

    # ── Chart 4: Stacked by factory 2026 ──
    st.subheader("Répartition par usine — Plan 2026 (courbes empilées)")
    FACTORY_COLORS_MAP = {
        "SICAM":    "#f5a623",
        "COMOCAP":  "#3b82f6",
        "TUCAL":    "#8b5cf6",
        "ABIDA":    "#ff6b9d",
        "ELFALLEH": "#00e5a0",
    }
    # Plotly requires rgba() format for transparent fills — hex+alpha not supported
    FACTORY_FILL_MAP = {
        "SICAM":    "rgba(245,166,35,0.6)",
        "COMOCAP":  "rgba(59,130,246,0.6)",
        "TUCAL":    "rgba(139,92,246,0.6)",
        "ABIDA":    "rgba(255,107,157,0.6)",
        "ELFALLEH": "rgba(0,229,160,0.6)",
    }
    fig4 = go.Figure()
    for factory, rows in HIST["factory_2026"].items():
        if not rows:
            continue
        fig4.add_trace(go.Scatter(
            x=[f"2026-{r[0]}" for r in rows],
            y=[r[1] for r in rows],
            name=factory, stackgroup="one",
            line=dict(width=0.5),
            fillcolor=FACTORY_FILL_MAP.get(factory, "rgba(153,153,153,0.6)"),
            mode="lines",
        ))
    fig4.update_layout(
        template="plotly_dark", plot_bgcolor="#0d1117", paper_bgcolor="#161b22",
        height=340, hovermode="x unified",
        legend=dict(orientation="h", yanchor="bottom", y=1.02),
        yaxis=dict(title="Tonnes/jour"),
        title="Courbes empilées par usine (plan 2026)",
    )
    st.plotly_chart(fig4, use_container_width=True)

    # ── Comparison table — 2026 values from real TONNAGE column ──
    st.subheader("Tableau comparatif")

    def fmt(v):
        """Format tonnes: no decimal, comma only if >= 10000"""
        v = round(v, 0)
        if v >= 10000:
            return f"{v:,.0f} t"
        else:
            return f"{int(v):,} t" if v >= 1000 else f"{int(v)} t"

    # Real 2026 totals from TONNAGE column (source of truth)
    USINE_2026_REAL = {
        "SICAM":    44713.0,
        "COMOCAP":  17280.0,
        "TUCAL":    20915.0,
        "ELFALLEH":  4250.0,
        "ABIDA":     6530.0,
    }
    sicam_2026_total  = GLOBAL_USINE_TONS.get("SICAM", USINE_2026_REAL["SICAM"])
    autres_2026_total = sum(GLOBAL_USINE_TONS.get(f, USINE_2026_REAL[f])
                            for f in ["COMOCAP","TUCAL","ABIDA","ELFALLEH"])
    total_2026        = GLOBAL_TOTAL_TONS if GLOBAL_TOTAL_TONS > 0 else 86548.0  # fallback = current real
    avg_peak_2026     = GLOBAL_PEAK_TONS / 15 if GLOBAL_PEAK_TONS > 0 else 2304.0

    comp_data = {
        "Indicateur": [
            "Total saison (toutes usines)",
            "Pic max en 1 jour",
            "Moyenne journalière 1–15 Jul",
            "Total SICAM",
            "Total autres usines",
        ],
        "2025 Réel": [
            fmt(HIST["stats"]["global_total"]),
            fmt(HIST["stats"]["global_max"]) + f"  ({HIST['stats']['global_max_date']})",
            fmt(HIST["stats"]["global_avg_peak"]) + "/jour",
            fmt(HIST["stats"]["sicam_total"]),
            fmt(HIST["stats"]["autres_total"]),
        ],
        "2026 Plan": [
            fmt(total_2026),
            fmt(HIST["stats"]["plan_max"]),
            fmt(avg_peak_2026) + "/jour",
            fmt(sicam_2026_total),
            fmt(autres_2026_total),
        ],
        "Variation": [
            f"{round(total_2026 - HIST['stats']['global_total']):+,} t",
            f"{round(HIST['stats']['plan_max'] - HIST['stats']['global_max']):+,} t",
            f"{round(avg_peak_2026 - HIST['stats']['global_avg_peak']):+,} t/j",
            "—", "—",
        ],
    }
    import pandas as _pd
    st.dataframe(_pd.DataFrame(comp_data), use_container_width=True, hide_index=True)
    st.download_button(
        "⬇️ Exporter comparatif (CSV)",
        data=df_to_csv(_pd.DataFrame(comp_data)),
        file_name="comparatif_2025_2026.csv",
        mime="text/csv",
    )


# ── TAB 8 (new): PRÉVISIONS DÉC→MAI→JUIN ────────────────────
with tab8:
    st.subheader("📊 Suivi des Prévisions — Décembre 2025 → Mai 2026 → Juin 2026")
    st.caption("Évolution des besoins par région et usine sur les 3 prévisions de la saison")

    # ── DATA EXACTES — source: TONNAGE_PAR_REGION_anouer__2___1_.xlsx ─────
    import math
    def _c100(x):
        if not x or x == 0: return 0
        return math.ceil(float(x) / 100) * 100

    DEC25 = {
        'CAP BON':          {'tonnage':48000,'SICAM':22000,'TUCAL':10000,'COMOCAP':13000,'ABIDA':0,   'ELFALLEH':3000,'BESOIN':48000},
        'NORD':             {'tonnage':12000,'SICAM':3500, 'TUCAL':7000, 'COMOCAP':1500, 'ABIDA':500, 'ELFALLEH':500, 'BESOIN':13000},
        'GAFSA / KASSRINE': {'tonnage':15500,'SICAM':11500,'TUCAL':2500, 'COMOCAP':1500, 'ABIDA':2500,'ELFALLEH':0,   'BESOIN':18000},
        'KAIROUAN':         {'tonnage':9000, 'SICAM':4500, 'TUCAL':2000, 'COMOCAP':1500, 'ABIDA':2000,'ELFALLEH':0,   'BESOIN':10000},
        'SIDI BOUZID':      {'tonnage':5500, 'SICAM':2500, 'TUCAL':500,  'COMOCAP':1500, 'ABIDA':2000,'ELFALLEH':0,   'BESOIN':6500},
        'BOUFICHA':         {'tonnage':3500, 'SICAM':1000, 'TUCAL':1000, 'COMOCAP':1000, 'ABIDA':0,   'ELFALLEH':500, 'BESOIN':3500},
    }
    MAI26 = {
        'CAP BON':          {'tonnage':49000,'SICAM':22000,'TUCAL':11000,'COMOCAP':13000,'ABIDA':0,   'ELFALLEH':4000,'BESOIN':50000},
        'NORD':             {'tonnage':12500,'SICAM':3500, 'TUCAL':4500, 'COMOCAP':3000, 'ABIDA':1000,'ELFALLEH':500, 'BESOIN':12500},
        'GAFSA / KASSRINE': {'tonnage':17000,'SICAM':11000,'TUCAL':2000, 'COMOCAP':500,  'ABIDA':3500,'ELFALLEH':0,   'BESOIN':17000},
        'KAIROUAN':         {'tonnage':8000, 'SICAM':4500, 'TUCAL':1200, 'COMOCAP':1000, 'ABIDA':1500,'ELFALLEH':0,   'BESOIN':8200},
        'SIDI BOUZID':      {'tonnage':4000, 'SICAM':2500, 'TUCAL':500,  'COMOCAP':1500, 'ABIDA':2000,'ELFALLEH':0,   'BESOIN':6500},
        'BOUFICHA':         {'tonnage':3000, 'SICAM':1000, 'TUCAL':1000, 'COMOCAP':1000, 'ABIDA':0,   'ELFALLEH':0,   'BESOIN':3000},
    }
    JUN26 = {
        'CAP BON':          {'tonnage':49000,'SICAM':22000,'TUCAL':11000,'COMOCAP':12500,'ABIDA':0,   'ELFALLEH':4000,'BESOIN':49500},
        'NORD':             {'tonnage':12500,'SICAM':3500, 'TUCAL':4500, 'COMOCAP':3000, 'ABIDA':1000,'ELFALLEH':500, 'BESOIN':12500},
        'GAFSA / KASSRINE': {'tonnage':18000,'SICAM':11000,'TUCAL':2500, 'COMOCAP':500,  'ABIDA':3500,'ELFALLEH':0,   'BESOIN':17500},
        'KAIROUAN':         {'tonnage':9000, 'SICAM':4500, 'TUCAL':1200, 'COMOCAP':1500, 'ABIDA':1500,'ELFALLEH':500, 'BESOIN':9200},
        'SIDI BOUZID':      {'tonnage':6500, 'SICAM':2500, 'TUCAL':500,  'COMOCAP':1500, 'ABIDA':2000,'ELFALLEH':0,   'BESOIN':6500},
        'BOUFICHA':         {'tonnage':3000, 'SICAM':1000, 'TUCAL':1000, 'COMOCAP':1000, 'ABIDA':0,   'ELFALLEH':0,   'BESOIN':3000},
    }

    REGIONS_P = ['CAP BON','NORD','GAFSA / KASSRINE','KAIROUAN','SIDI BOUZID','BOUFICHA']
    PREVISIONS_DATES = {'DEC25':'16/01/2026','MAI26':'09/05/2026','JUN26':'03/06/2026'}
    USINES_P  = ['SICAM','TUCAL','COMOCAP','ABIDA','ELFALLEH']
    U_COLORS  = {'SICAM':'#F5A623','TUCAL':'#8B5CF6','COMOCAP':'#3B82F6','ABIDA':'#FF6B9D','ELFALLEH':'#00C896'}
    R_COLORS  = {'CAP BON':'#1F4E79','NORD':'#375623','GAFSA / KASSRINE':'#064E3B',
                 'KAIROUAN':'#7B2D8B','SIDI BOUZID':'#C00000','BOUFICHA':'#B45309'}

    # ── KPI résumé ──────────────────────────────────────────
    st.markdown("### 📌 Vue d'ensemble")
    k1,k2,k3,k4 = st.columns(4)
    k1.metric("Production Déc25",  f"{sum(_c100(v['tonnage']) for v in DEC25.values()):,} t")
    k2.metric("Production Mai26",  f"{sum(_c100(v['tonnage']) for v in MAI26.values()):,} t",
              delta=f"{sum(_c100(v['tonnage']) for v in MAI26.values())-sum(_c100(v['tonnage']) for v in DEC25.values()):+,} t")
    k3.metric("Production Juin26", f"{sum(_c100(v['tonnage']) for v in JUN26.values()):,} t",
              delta=f"{sum(_c100(v['tonnage']) for v in JUN26.values())-sum(_c100(v['tonnage']) for v in MAI26.values()):+,} t vs Mai")
    k4.metric("Besoin total usines Mai26", f"{sum(_c100(v['BESOIN']) for v in MAI26.values()):,} t")
    st.markdown("---")

    # ── TAB internes : Production / Usines / Décalages ──────
    pt1, pt2, pt3 = st.tabs([
        "📈 Évolution Production par Région",
        "🏭 Besoin Usines Déc vs Mai",
        "🔍 Analyse Décalages Mai26",
    ])

    # ─── Sous-tab 1 : Evolution production ───────────────────
    with pt1:
        # Bar chart grouped : Dec vs Mai par region
        import plotly.graph_objects as go

        fig_evol = go.Figure()
        fig_evol.add_trace(go.Bar(
            name='Déc 2025',
            x=REGIONS_P,
            y=[_c100(DEC25[r]['tonnage']) for r in REGIONS_P],
            marker_color='#4A90D9',
            text=[f"{_c100(DEC25[r]['tonnage']):,}t" for r in REGIONS_P],
            textposition='outside',
        ))
        fig_evol.add_trace(go.Bar(
            name='Mai 2026',
            x=REGIONS_P,
            y=[_c100(MAI26[r]['tonnage']) for r in REGIONS_P],
            marker_color='#00C896',
            text=[f"{_c100(MAI26[r]['tonnage']):,}t" for r in REGIONS_P],
            textposition='outside',
        ))
        # Jun26 — afficher TONNAGE PRODUIT (pas BESOIN)
        fig_evol.add_trace(go.Bar(
            name='Juin 2026 (Production prévue)',
            x=REGIONS_P,
            y=[_c100(JUN26[r]['tonnage']) for r in REGIONS_P],
            marker_color='#F5A623',
            text=[f"{_c100(JUN26[r]['tonnage']):,}t" for r in REGIONS_P],
            textposition='outside',
        ))
        fig_evol.update_layout(
            barmode='group', template='plotly_dark',
            paper_bgcolor='#161b22', plot_bgcolor='#0d1117',
            height=420, title='Évolution production prévue par région (3 prévisions)',
            yaxis_title='Tonnes', hovermode='x unified',
            legend=dict(orientation='h', yanchor='bottom', y=1.02),
        )
        st.plotly_chart(fig_evol, use_container_width=True)

        # Line evolution total
        fig_line = go.Figure()
        previsions = ['Déc 2025','Mai 2026']
        totaux = [sum(_c100(v['tonnage']) for v in DEC25.values()),
                  sum(_c100(v['tonnage']) for v in MAI26.values())]
        if JUN26:
            previsions.append('Juin 2026')
            totaux.append(sum(_c100(v['tonnage']) for v in JUN26.values()))

        fig_line.add_trace(go.Scatter(
            x=previsions, y=totaux, mode='lines+markers+text',
            text=[f"{t:,}t" for t in totaux],
            textposition='top center',
            line=dict(color='#F5A623', width=3),
            marker=dict(size=12, color='#F5A623'),
        ))
        fig_line.update_layout(
            template='plotly_dark', paper_bgcolor='#161b22',
            plot_bgcolor='#0d1117', height=280,
            title='Évolution du total production — toutes régions',
            yaxis_title='Tonnes',
        )
        st.plotly_chart(fig_line, use_container_width=True)

        # Table comparaison
        rows_comp = []
        for r in REGIONS_P:
            d_t = _c100(DEC25[r]['tonnage'])
            m_t = _c100(MAI26[r]['tonnage'])
            j_t = _c100(JUN26[r]['tonnage'])
            j_b = _c100(JUN26[r]['BESOIN'])
            evol = j_t - m_t
            rows_comp.append({
                'Région': r,
                'Déc 2025 (prod)': f"{d_t:,} t",
                'Mai 2026 (prod)': f"{m_t:,} t",
                'Juin 2026 (prod)': f"{j_t:,} t",
                'Évolution Mai→Juin': f"{evol:+,} t",
                'Besoin Usines Juin': f"{j_b:,} t",
                '% vs Déc': f"{(j_t/d_t-1)*100:+.1f}%" if d_t > 0 else '—',
            })
        st.dataframe(pd.DataFrame(rows_comp), use_container_width=True, hide_index=True)

    # ─── Sous-tab 2 : Besoin usines ──────────────────────────
    with pt2:
        col_sel = st.selectbox("Choisir une usine", USINES_P, key='usine_sel_prev')

        fig_usine = go.Figure()
        fig_usine.add_trace(go.Bar(
            name='Besoin Déc 2025',
            x=REGIONS_P,
            y=[DEC25[r][col_sel] for r in REGIONS_P],
            marker_color='#4A90D9',
            text=[f"{DEC25[r][col_sel]:,}t" if DEC25[r][col_sel] > 0 else '' for r in REGIONS_P],
            textposition='outside',
        ))
        fig_usine.add_trace(go.Bar(
            name='Besoin Mai 2026',
            x=REGIONS_P,
            y=[MAI26[r][col_sel] for r in REGIONS_P],
            marker_color=U_COLORS[col_sel],
            text=[f"{MAI26[r][col_sel]:,}t" if MAI26[r][col_sel] > 0 else '' for r in REGIONS_P],
            textposition='outside',
        ))
        fig_usine.update_layout(
            barmode='group', template='plotly_dark',
            paper_bgcolor='#161b22', plot_bgcolor='#0d1117',
            height=380, title=f'Besoin {col_sel} — Déc 2025 vs Mai 2026 par région',
            yaxis_title='Tonnes',
        )
        st.plotly_chart(fig_usine, use_container_width=True)

        # Totaux usine
        c1, c2, c3 = st.columns(3)
        tot_d = sum(DEC25[r][col_sel] for r in REGIONS_P)
        tot_m = sum(MAI26[r][col_sel] for r in REGIONS_P)
        c1.metric(f"Total besoin {col_sel} Déc25", f"{tot_d:,} t")
        c2.metric(f"Total besoin {col_sel} Mai26", f"{tot_m:,} t", delta=f"{tot_m-tot_d:+,} t")
        c3.metric(f"Variation", f"{(tot_m/tot_d-1)*100:+.1f}%" if tot_d > 0 else "—")

        # Radar chart all usines Dec vs Mai
        fig_radar = go.Figure()
        theta = USINES_P + [USINES_P[0]]
        fig_radar.add_trace(go.Scatterpolar(
            r=[sum(DEC25[r][u] for r in REGIONS_P) for u in USINES_P] +
              [sum(DEC25[r][USINES_P[0]] for r in REGIONS_P)],
            theta=theta, fill='toself', name='Déc 2025',
            line_color='#4A90D9',
        ))
        fig_radar.add_trace(go.Scatterpolar(
            r=[sum(MAI26[r][u] for r in REGIONS_P) for u in USINES_P] +
              [sum(MAI26[r][USINES_P[0]] for r in REGIONS_P)],
            theta=theta, fill='toself', name='Mai 2026',
            line_color='#00C896',
        ))
        fig_radar.update_layout(
            polar=dict(bgcolor='#161b22'),
            template='plotly_dark', paper_bgcolor='#161b22',
            height=380, title='Répartition totale par usine — Déc vs Mai',
        )
        st.plotly_chart(fig_radar, use_container_width=True)

    # ─── Sous-tab 3 : Analyse décalages ──────────────────────
    with pt3:
        st.markdown("#### Décalage = Besoin Usine − Production Allouée")
        st.caption("La production est allouée proportionnellement aux besoins de chaque usine dans la région.")

        rows_decal = []
        for region in REGIONS_P:
            m = MAI26[region]
            prod = m['tonnage']
            total_b = m['BESOIN']
            for u in USINES_P:
                b = m[u]
                if b == 0: continue
                share      = b / total_b if total_b > 0 else 0
                prod_alloc = round(prod * share)
                solde      = prod_alloc - b
                rows_decal.append({
                    'Région':        region,
                    'Usine':         u,
                    'Besoin (t)':    b,
                    'Prod. allouée': prod_alloc,
                    'Solde (t)':     solde,
                    'Statut':        '✅ OK' if solde >= 0 else ('⚠️ Léger' if solde >= -500 else '🔴 Déficit'),
                })

        df_decal = pd.DataFrame(rows_decal)

        # Heatmap-style visualization
        pivot_solde = df_decal.pivot_table(
            index='Région', columns='Usine', values='Solde (t)', fill_value=0)

        fig_heat = go.Figure(data=go.Heatmap(
            z=pivot_solde.values,
            x=list(pivot_solde.columns),
            y=list(pivot_solde.index),
            colorscale=[[0,'#C62828'],[0.5,'#F5F5F5'],[1,'#2E7D32']],
            zmid=0,
            text=[[f"{v:+,}t" for v in row] for row in pivot_solde.values],
            texttemplate='%{text}',
            textfont=dict(size=11, color='white'),
            colorbar=dict(title='Solde (t)'),
        ))
        fig_heat.update_layout(
            template='plotly_dark', paper_bgcolor='#161b22',
            height=350, title='Solde par région × usine (Mai 2026) — Vert=excédent Rouge=déficit',
        )
        st.plotly_chart(fig_heat, use_container_width=True)

        # Table détaillée
        st.dataframe(
            df_decal,
            use_container_width=True,
            hide_index=True,
            column_config={
                'Solde (t)': st.column_config.NumberColumn('Solde (t)', format='%+d t'),
                'Statut':    st.column_config.TextColumn(width='small'),
            }
        )

        # Summary déficits
        deficits = df_decal[df_decal['Solde (t)'] < 0].sort_values('Solde (t)')
        if not deficits.empty:
            st.markdown("#### ⚠️ Déficits à combler")
            for _, row in deficits.iterrows():
                color = "🔴" if row['Solde (t)'] < -500 else "⚠️"
                st.warning(f"{color} **{row['Région']}** → **{row['Usine']}** : manque **{abs(row['Solde (t)']):,}t** (besoin {row['Besoin (t)']:,}t, production allouée {row['Prod. allouée']:,}t)")
        else:
            st.success("✅ Aucun déficit détecté — production Mai26 couvre tous les besoins")

        st.success("✅ Prévision Juin 2026 intégrée — données du 03/06/2026")

# ── TAB 7: TONNAGE PAR RÉGION ────────────────────────────────
with tab7:
    st.subheader("🗺️ Tonnage par Région — Saison 2026")

    import math
    def ceil100(x):
        if not x or x <= 0: return 0
        return math.ceil(x / 100) * 100

    USINES_ORD = ["SICAM","COMOCAP","TUCAL","ELFALLEH","ABIDA"]
    COMMS_ORD  = ["FEDI","MAKKI BEN SALAH","KHALIL","ACHREF AJLANI","JILANI OBAY"]
    # Exact order requested — GAFSA+KASSRINE merged, no AUTRE
    REG_ORD    = ["CAP BON 2","CAP BON 1","NORD","GAFSA / KASSRINE",
                  "KAIROUAN","SIDI BOUZID","BOUFICHA"]

    REG_NORM = {
        "NABEUL":   "CAP BON 2",
        "BEJA":     "NORD",
        "MANOUBA":  "NORD",
        "GAFSA":    "GAFSA / KASSRINE",
        "KASSRINE": "GAFSA / KASSRINE",
    }

    # ── Detect if a date filter is active ──────────────────
    filter_active = False
    filter_label  = "Saison complète"
    d_filter_0    = None
    d_filter_1    = None

    if peak_only:
        filter_active = True
        filter_label  = "⚡ Pic 1–15 Juillet"
        d_filter_0    = PEAK_START
        d_filter_1    = PEAK_END
    elif isinstance(date_range, (list, tuple)) and len(date_range) == 2:
        d0, d1 = date_range
        season_start = datetime.date(2026, 6, 15)
        season_end   = datetime.date(2026, 8, 31)
        if d0 != season_start or d1 != season_end:
            filter_active = True
            filter_label  = f"{d0.strftime('%d/%m')} → {d1.strftime('%d/%m/%Y')}"
            d_filter_0    = d0
            d_filter_1    = d1

    st.caption(f"📅 Période affichée : **{filter_label}**")

    # ── Build df_reg — TOUJOURS depuis agriculteurs (region 100% normalisée) ──
    # NOTE: On N'utilise PAS planning.region car get_region(zone) de l'optimizer
    # ne reconnaît qu'un nombre limité de zones → la plupart deviennent
    # "AUTRE" → toutes les régions sauf KAIROUAN seraient filtrées.
    # Le tonnage par région est une vue SAISON COMPLÈTE — le filtre date 
    # ne s'applique pas ici.
    if False:  # if False → toujours aller dans else (load_region_data)
        pass
    else:
        # Use agriculteurs from Supabase — full season declared tonnage
        @st.cache_data(ttl=30)
        def load_region_data(_v=0):
            try:
                sb = get_supabase()

                # ── Source 1: Table agriculteurs ────────────────────────
                data_agri = sb.table("agriculteurs").select(
                    "commercial,nom,tonnage_total,usine,zone,region"
                ).execute().data
                df_agri = pd.DataFrame(data_agri) if data_agri else pd.DataFrame()

                # ── Source 2: Table planning (region aussi disponible) ──
                # Pagination obligatoire (Supabase limite à 1000/requête)
                data_plan = []
                _offset = 0
                while True:
                    _b = sb.table("planning").select(
                        "commercial,agriculteur,tonnes_jour,usine,region"
                    ).range(_offset, _offset+999).execute().data
                    if not _b:
                        break
                    data_plan.extend(_b)
                    if len(_b) < 1000:
                        break
                    _offset += 1000
                df_plan = pd.DataFrame(data_plan) if data_plan else pd.DataFrame()

                # ── Normalisation helper (étendue) ────────────────────────
                # Mapping complet incluant CAP BON sans numéro, casse mixte, etc.
                NORM = {
                    # Variantes Cap Bon (sans numéro = par défaut CAP BON 1)
                    "CAP BON":"CAP BON 1","cap bon":"CAP BON 1","Cap Bon":"CAP BON 1",
                    "CAPBON":"CAP BON 1","capbon":"CAP BON 1",
                    "CAP BON 1":"CAP BON 1","cap bon 1":"CAP BON 1",
                    "CAP BON 2":"CAP BON 2","cap bon 2":"CAP BON 2",
                    "CAPB1":"CAP BON 1","capb1":"CAP BON 1","CAP B1":"CAP BON 1",
                    "CAPB2":"CAP BON 2","capb2":"CAP BON 2","CAP B2":"CAP BON 2",
                    "CAPBON1":"CAP BON 1","CAPBON2":"CAP BON 2",
                    # Nabeul → CAP BON 2 (zone sud)
                    "NABEUL":"CAP BON 2","nabeul":"CAP BON 2","Nabeul":"CAP BON 2",
                    # NORD (Beja, Manouba, Tunis, Jandouba)
                    "BEJA":"NORD","beja":"NORD","Beja":"NORD","BÉJA":"NORD",
                    "MANOUBA":"NORD","manouba":"NORD","Manouba":"NORD",
                    "JANDOUBA":"NORD","jandouba":"NORD",
                    "TUNIS":"NORD","tunis":"NORD",
                    "Nord":"NORD","nord":"NORD",
                    # GAFSA / KASSRINE (toutes variantes)
                    "GAFSA":"GAFSA / KASSRINE","gafsa":"GAFSA / KASSRINE",
                    "KASSRINE":"GAFSA / KASSRINE","kassrine":"GAFSA / KASSRINE",
                    "KASSERINE":"GAFSA / KASSRINE","kasserine":"GAFSA / KASSRINE",
                    "GAFSA/KASSRINE":"GAFSA / KASSRINE",
                    "GAFSA / KASSERINE":"GAFSA / KASSRINE",
                    "Gafsa":"GAFSA / KASSRINE","Kassrine":"GAFSA / KASSRINE",
                    # KAIROUAN
                    "Kairouan":"KAIROUAN","kairouan":"KAIROUAN",
                    # SIDI BOUZID
                    "Sidi Bouzid":"SIDI BOUZID","sidi bouzid":"SIDI BOUZID",
                    "SIDI-BOUZID":"SIDI BOUZID","SIDIBOUZID":"SIDI BOUZID",
                    # BOUFICHA / Sousse
                    "Boufiche":"BOUFICHA","boufiche":"BOUFICHA",
                    "SOUSSE":"BOUFICHA","sousse":"BOUFICHA",
                    "Boufiche":"BOUFICHA","Bouficha":"BOUFICHA",
                }

                def normalize_reg(series):
                    s = series.fillna("").astype(str).str.strip()
                    # Premier passage: matching exact
                    s = s.replace(NORM)
                    # Deuxième passage: UPPER pour ce qui reste
                    mask = ~s.isin(REG_ORD) & (s != "")
                    if mask.any():
                        s_upper = s[mask].str.upper()
                        s[mask] = s_upper.replace(NORM)
                    # Troisième passage: contient "CAP BON" → CAP BON 1
                    mask2 = ~s.isin(REG_ORD) & (s != "")
                    if mask2.any():
                        cap_mask = s[mask2].str.upper().str.contains("CAP", na=False) & s[mask2].str.upper().str.contains("BON", na=False)
                        if cap_mask.any():
                            idx = s[mask2][cap_mask].index
                            s.loc[idx] = "CAP BON 1"
                    return s

                # ── Traiter agriculteurs ────────────────────────────────
                result_frames = []
                if not df_agri.empty:
                    df_agri["tonnage_total"] = pd.to_numeric(
                        df_agri["tonnage_total"], errors="coerce")
                    df_agri = df_agri[df_agri["tonnage_total"] > 0]
                    df_agri["REGION"] = normalize_reg(df_agri["region"])
                    df_agri["nom_key"] = df_agri["nom"].fillna("").astype(str).str.upper()
                    df_agri_ok = df_agri[df_agri["REGION"].isin(REG_ORD)].copy()
                    if not df_agri_ok.empty:
                        df_agri_ok = df_agri_ok.rename(columns={"nom":"nom"})
                        result_frames.append(df_agri_ok[
                            ["commercial","nom","tonnage_total","usine","REGION"]])

                # ── Traiter planning si agriculteurs vide ou sans régions ─
                if (not result_frames) and not df_plan.empty:
                    df_plan["tonnage_total"] = pd.to_numeric(
                        df_plan["tonnes_jour"], errors="coerce")
                    df_plan = df_plan[df_plan["tonnage_total"] > 0]
                    df_plan["REGION"] = normalize_reg(df_plan["region"])
                    df_plan_ok = df_plan[df_plan["REGION"].isin(REG_ORD)].copy()
                    if not df_plan_ok.empty:
                        df_plan_ok = df_plan_ok.rename(columns={"agriculteur":"nom"})
                        # Agréger par agriculteur (sum des tonnes/jour)
                        df_plan_agg = df_plan_ok.groupby(
                            ["commercial","nom","usine","REGION"],
                            as_index=False
                        )["tonnage_total"].sum()
                        result_frames.append(df_plan_agg)

                if not result_frames:
                    return None

                df_final = pd.concat(result_frames, ignore_index=True)
                df_final = df_final[df_final["REGION"].isin(REG_ORD)]
                return df_final if not df_final.empty else None

            except Exception as e:
                return None

        df_reg = load_region_data(_v=st.session_state["sb_refresh"])

        # Fallback to planning if agriculteurs empty
        if (df_reg is None or df_reg.empty) and not p.empty and "Région" in p.columns:
            p_all = p.copy()
            p_all["REGION"] = p_all["Région"].fillna("").astype(str).str.strip().str.upper()
            p_all["REGION"] = p_all["REGION"].replace(REG_NORM)
            p_all = p_all[p_all["REGION"].isin(REG_ORD)]
            p_all["usine"]       = p_all["Usine"].fillna("").astype(str).str.strip().str.upper()
            p_all["commercial"]  = p_all["Commercial"].fillna("").astype(str).str.strip()
            p_all["nom"]         = p_all["Agriculteur"].fillna("").astype(str).str.strip()
            p_all["tonnage_total"] = pd.to_numeric(p_all["Tonnes/Jour"], errors="coerce").fillna(0)
            df_reg = p_all[["commercial","nom","tonnage_total","usine","REGION"]].copy()

        data_source_label = "saison complète (tonnage déclaré)"

    # ── Display ─────────────────────────────────────────────
    if df_reg is None or df_reg.empty:
        st.warning("⚠️ Aucune donnée de région disponible dans Supabase.")
        st.info("""
**Causes possibles :**
- La colonne `region` dans Supabase est vide ou non normalisée
- Les commerciaux n'ont pas encore uploadé leurs fichiers

**Solution :**
1. Vérifiez l'onglet 🌾 Gestion Agriculteurs → Diagnostic Supabase
2. Re-uploadez les fichiers des commerciaux avec les régions correctes
3. Ou lancez `python migrate.py` après `python optimizer_v2.py`
        """)
    else:
        # KPI cards
        total_reg = df_reg["tonnage_total"].sum()
        k1, k2, k3, k4 = st.columns(4)
        k1.metric("Régions actives", df_reg["REGION"].nunique())
        k2.metric("Tonnage", f"{total_reg:,.0f} t")
        k3.metric("Région principale",
                  df_reg.groupby("REGION")["tonnage_total"].sum().idxmax()
                  if not df_reg.empty else "—")
        k4.metric("Usine principale",
                  df_reg.groupby("usine")["tonnage_total"].sum().idxmax()
                  if not df_reg.empty else "—")
        st.caption(f"Source: {data_source_label}")
        st.markdown("---")

        # ── Pivot Région × Usine ──────────────────────────
        st.subheader("📊 Tonnage par Région × Usine")
        pv_usine = df_reg.groupby(["REGION","usine"])["tonnage_total"].sum().round(0).unstack(fill_value=0)
        for u in USINES_ORD:
            if u not in pv_usine.columns: pv_usine[u] = 0
        pv_usine = pv_usine[[u for u in USINES_ORD if u in pv_usine.columns]]
        # Apply ceil100 to each cell
        for col in pv_usine.columns:
            pv_usine[col] = pv_usine[col].apply(ceil100)
        pv_usine["TOTAL"] = pv_usine.sum(axis=1)
        pv_usine = pv_usine.reindex([r for r in REG_ORD if r in pv_usine.index])
        total_u = pv_usine.sum(); total_u.name = "TOTAL"
        pv_display = pd.concat([pv_usine, total_u.to_frame().T]).astype(int)
        pv_display.index.name = "Région"
        st.dataframe(pv_display, use_container_width=True,
            column_config={"TOTAL": st.column_config.NumberColumn("TOTAL", format="%d t"),
                **{u: st.column_config.NumberColumn(u, format="%d t")
                   for u in USINES_ORD if u in pv_display.columns}})

        pv_long = pv_usine.drop(columns="TOTAL").reset_index().melt(
            id_vars="REGION", var_name="Usine", value_name="Tonnes")
        pv_long = pv_long[pv_long["Tonnes"] > 0]
        fig_ru = px.bar(pv_long, x="REGION", y="Tonnes", color="Usine",
            barmode="stack", color_discrete_map=FACTORY_COLORS,
            title=f"Tonnage par région × usine — {filter_label}",
            template="plotly_dark", text_auto=".2s")
        fig_ru.update_layout(paper_bgcolor="#161b22", plot_bgcolor="#0d1117",
                              height=400, hovermode="x unified")
        st.plotly_chart(fig_ru, use_container_width=True)
        st.markdown("---")

        # ── Pivot Région × Commercial ─────────────────────
        st.subheader("📊 Tonnage par Région × Commercial")
        pv_comm = df_reg.groupby(["REGION","commercial"])["tonnage_total"].sum().round(0).unstack(fill_value=0)
        for c in COMMS_ORD:
            if c not in pv_comm.columns: pv_comm[c] = 0
        pv_comm = pv_comm[[c for c in COMMS_ORD if c in pv_comm.columns]]
        for col in pv_comm.columns:
            pv_comm[col] = pv_comm[col].apply(ceil100)
        pv_comm["TOTAL"] = pv_comm.sum(axis=1)
        pv_comm = pv_comm.reindex([r for r in REG_ORD if r in pv_comm.index])
        total_c = pv_comm.sum(); total_c.name = "TOTAL"
        pv_comm_display = pd.concat([pv_comm, total_c.to_frame().T]).astype(int)
        pv_comm_display.index.name = "Région"
        st.dataframe(pv_comm_display, use_container_width=True)

        pv_comm_long = pv_comm.drop(columns="TOTAL").reset_index().melt(
            id_vars="REGION", var_name="Commercial", value_name="Tonnes")
        pv_comm_long = pv_comm_long[pv_comm_long["Tonnes"] > 0]
        fig_rc = px.bar(pv_comm_long, x="REGION", y="Tonnes", color="Commercial",
            barmode="stack", color_discrete_map=COMM_COLORS,
            title=f"Tonnage par région × commercial — {filter_label}",
            template="plotly_dark", text_auto=".2s")
        fig_rc.update_layout(paper_bgcolor="#161b22", plot_bgcolor="#0d1117",
                              height=400, hovermode="x unified")
        st.plotly_chart(fig_rc, use_container_width=True)
        st.markdown("---")

        # ── Pie + Summary table ───────────────────────────
        c1, c2 = st.columns(2)
        with c1:
            reg_tot = df_reg.groupby("REGION")["tonnage_total"].sum().reset_index()
            fig_pie = px.pie(reg_tot, names="REGION", values="tonnage_total",
                title=f"Répartition par région — {filter_label}", hole=0.4,
                template="plotly_dark", color_discrete_sequence=px.colors.qualitative.Set3)
            fig_pie.update_layout(paper_bgcolor="#161b22", height=350)
            st.plotly_chart(fig_pie, use_container_width=True)
        with c2:
            reg_ton  = df_reg.groupby("REGION")["tonnage_total"].sum().round(0).reset_index()
            reg_ton.columns = ["Région","Tonnage (t)"]
            nom_col = "nom" if "nom" in df_reg.columns else "Agriculteur"
            reg_uniq = df_reg.groupby("REGION")[nom_col].nunique().reset_index()
            reg_uniq.columns = ["Région","Agriculteurs"]
            reg_usine = df_reg.groupby("REGION").apply(
                lambda x: x.groupby("usine")["tonnage_total"].sum().idxmax()
                if not x.empty else "—").reset_index()
            reg_usine.columns = ["Région","Usine principale"]
            summary = reg_ton.merge(reg_uniq, on="Région").merge(reg_usine, on="Région")
            summary = summary.sort_values("Tonnage (t)", ascending=False)
            summary["% Total"] = (summary["Tonnage (t)"] / summary["Tonnage (t)"].sum() * 100).round(1).astype(str) + "%"
            st.dataframe(summary, use_container_width=True, hide_index=True, height=350)

        # ── Download ──────────────────────────────────────
        st.markdown("---")
        col_dl1, col_dl2 = st.columns(2)
        with col_dl1:
            st.download_button("⬇️ Région × Usine (CSV)",
                data=df_to_csv(pv_display.reset_index()),
                file_name=f"tonnage_region_usine_{filter_label.replace(' ','_')}.csv",
                mime="text/csv", use_container_width=True)
        with col_dl2:
            st.download_button("⬇️ Région × Commercial (CSV)",
                data=df_to_csv(pv_comm_display.reset_index()),
                file_name=f"tonnage_region_commercial_{filter_label.replace(' ','_')}.csv",
                mime="text/csv", use_container_width=True)

# ── TAB 9: GESTION AGRICULTEURS ──────────────────────────────
with tab9:

    # Directeur voit tout, commercial voit seulement ses agriculteurs
    if CURRENT_ROLE not in ("directeur", "commercial"):
        st.warning("🔒 Accès réservé au directeur et aux commerciaux.")
        st.stop()

    sb = get_supabase()
    IS_COMMERCIAL = CURRENT_ROLE == "commercial"

    # ── Constantes formulaire ──
    COMMERCIALS  = ["FEDI","MAKKI BEN SALAH","KHALIL","ACHREF AJLANI","JILANI OBAY"]
    USINES       = ["SICAM","COMOCAP","TUCAL","ABIDA","ELFALLEH"]
    ACCESS_CODES = ["PL/PPL","PL/SEMI","RM"]     # '/' uniquement, pas de '-'
    REGIONS      = ["CAP BON 1","CAP BON 2","NORD","GAFSA / KASSRINE",
                    "KAIROUAN","SIDI BOUZID","BOUFICHA"]

    # ── DIAGNOSTIC SUPABASE (directeur uniquement) ─────────────
    if CURRENT_ROLE == "directeur":
        with st.expander("🔍 Diagnostic Supabase — Qualité des données", expanded=False):
            try:
                sb_diag = get_supabase()
                diag_data = sb_diag.table("agriculteurs").select(
                    "commercial,region,usine,tonnage_total,date_debut,accessibilite"
                ).execute().data
                df_diag = pd.DataFrame(diag_data) if diag_data else pd.DataFrame()

                if df_diag.empty:
                    st.warning("Table agriculteurs vide.")
                else:
                    df_diag["tonnage_total"] = pd.to_numeric(df_diag["tonnage_total"], errors="coerce")
                    df_diag["date_debut"]    = pd.to_datetime(df_diag["date_debut"], errors="coerce")

                    d1, d2, d3, d4 = st.columns(4)
                    d1.metric("Total lignes", len(df_diag))
                    d2.metric("Total tonnage", f"{df_diag['tonnage_total'].sum():,.0f} t")
                    d3.metric("Agriculteurs uniques", df_diag["commercial"].nunique())

                    # Problèmes détectés
                    problems = []

                    # Régions invalides
                    GOOD_REGIONS = {"CAP BON 1","CAP BON 2","NORD","GAFSA / KASSRINE",
                                    "KAIROUAN","SIDI BOUZID","BOUFICHA"}
                    bad_reg = df_diag[~df_diag["region"].fillna("").str.strip().isin(GOOD_REGIONS)]
                    if len(bad_reg) > 0:
                        problems.append(f"⚠️ {len(bad_reg)} lignes avec région non normalisée: {bad_reg['region'].unique().tolist()}")

                    # Dates invalides (avant 2026)
                    bad_dates = df_diag[df_diag["date_debut"].dt.year < 2026]
                    if len(bad_dates) > 0:
                        problems.append(f"⚠️ {len(bad_dates)} lignes avec date_debut < 2026")

                    # Accessibilités invalides
                    GOOD_ACCESS = {"PL/PPL","PL/SEMI","RM","TRC/PPL","TRC/PPL/PL","PPL","PL","SEMI","TRAC+P,PLD","PL/PPL/SEMI"}
                    bad_acc = df_diag[~df_diag["accessibilite"].fillna("").isin(GOOD_ACCESS)]
                    if len(bad_acc) > 0:
                        vals = bad_acc['accessibilite'].unique().tolist()
                        problems.append(f"ℹ️ {len(bad_acc)} lignes avec accessibilité non répertoriée: {vals} — vérifier si valide")

                    # Tonnage nul
                    bad_ton = df_diag[df_diag["tonnage_total"].isna() | (df_diag["tonnage_total"] <= 0)]
                    if len(bad_ton) > 0:
                        problems.append(f"⚠️ {len(bad_ton)} lignes avec tonnage nul ou invalide")

                    d4.metric("Problèmes détectés", len(problems),
                              delta="OK" if len(problems)==0 else f"{len(problems)} anomalies",
                              delta_color="normal" if len(problems)==0 else "inverse")

                    if problems:
                        st.markdown("**Anomalies détectées :**")
                        for p in problems:
                            st.warning(p)
                        st.markdown("**ℹ️ Note :** Les codes TRC/PPL, TRC/PPL/PL sont valides. Si une autre valeur est signalée, vérifier avec le commercial concerné.")
                    else:
                        st.success("✅ Toutes les données Supabase sont propres et cohérentes.")

                    # Par commercial
                    st.markdown("**Par commercial :**")
                    comm_stats = df_diag.groupby("commercial").agg(
                        Lignes=("tonnage_total","count"),
                        Tonnage=("tonnage_total","sum"),
                    ).round(0).reset_index()
                    st.dataframe(comm_stats, use_container_width=True, hide_index=True)

                    # Nettoyage en un clic (directeur)
                    st.markdown("---")
                    if st.button("🧹 Nettoyer données corrompues (régions/dates invalides)",
                                 type="secondary"):
                        try:
                            # Supprimer beja, manouba et toutes régions invalides
                            for bad_region in ["beja","BEJA","manouba","MANOUBA","nabeul"]:
                                sb_diag.table("agriculteurs").delete().eq(
                                    "region", bad_region).execute()
                            # Supprimer dates invalides
                            sb_diag.table("agriculteurs").delete().lt(
                                "date_debut", "2026-01-01").execute()
                            st.success("✅ Données corrompues supprimées. Rechargez la page.")
                            st.cache_data.clear()
                            st.session_state["sb_refresh"] += 1
                            st.rerun()
                        except Exception as e:
                            st.error(f"Erreur nettoyage: {e}")

            except Exception as e:
                st.error(f"Erreur diagnostic Supabase: {e}")

    # ── Helper : charger tous les agriculteurs depuis Supabase ──
    @st.cache_data(ttl=10)
    def load_agriculteurs():
        try:
            q = sb.table("agriculteurs").select("*").order("commercial")
            data = q.execute().data
            df_a = pd.DataFrame(data) if data else pd.DataFrame()
            # Commercial voit seulement ses agriculteurs
            if IS_COMMERCIAL and not df_a.empty:
                df_a = df_a[df_a["commercial"] == CURRENT_NAME]
            return df_a
        except Exception as e:
            st.error(f"Erreur chargement agriculteurs : {e}")
            return pd.DataFrame()

    # ── Titre + bouton refresh ──
    col_title, col_refresh = st.columns([5, 1])
    with col_title:
        if IS_COMMERCIAL:
            st.subheader(f"🌾 Mes Agriculteurs — {CURRENT_NAME}")
            st.caption("Ajoutez, modifiez ou supprimez vos propres agriculteurs. "
                       "Contactez le directeur pour régénérer le planning.")
        else:
            st.subheader("🌾 Gestion des Agriculteurs")
            st.caption("Ajouter, modifier ou supprimer un agriculteur. "
                       "Les changements sont sauvegardés dans Supabase. "
                       "Cliquez 'Régénérer le planning' ensuite pour recalculer.")
    with col_refresh:
        if st.button("🔄", help="Rafraîchir la liste"):
            st.cache_data.clear()
            st.rerun()

    st.divider()

    # ── Sous-onglets : Ajouter / Modifier / Supprimer / Liste ──
    a1, a2, a3, a4 = st.tabs([
        "➕ Ajouter",
        "✏️ Modifier",
        "🗑️ Supprimer",
        "📋 Liste complète",
    ])

    # ════════════════════════════════════════════════════════
    # SOUS-ONGLET 1 : AJOUTER UN AGRICULTEUR
    # ════════════════════════════════════════════════════════
    with a1:
        st.subheader("Ajouter un nouvel agriculteur")

        with st.form("form_add", clear_on_submit=True):
            c1, c2 = st.columns(2)
            with c1:
                nom         = st.text_input("Nom de l'agriculteur *",
                                            placeholder="ex: MOHAMED BEN ALI")
                # Commercial voit seulement son propre nom dans la liste
                if IS_COMMERCIAL:
                    commercial = CURRENT_NAME
                    st.info(f"Commercial : **{CURRENT_NAME}**")
                else:
                    commercial = st.selectbox("Commercial *", COMMERCIALS)
                usine       = st.selectbox("Usine *", USINES)
                tonnage     = st.number_input("Tonnage total (tonnes) *",
                                              min_value=10.0, max_value=10000.0,
                                              value=200.0, step=25.0)
            with c2:
                region      = st.selectbox("Région", REGIONS)
                zone        = st.text_input("Zone / Localisation",
                                            placeholder="ex: dar allouch")
                access      = st.selectbox("Accessibilité véhicule *", ACCESS_CODES,
                                           help="PL/PPL = route normale | PL/SEMI = grande route")
                date_debut  = st.date_input("Date début récolte *",
                                            value=pd.Timestamp("2026-07-01"),
                                            min_value=pd.Timestamp("2026-06-15"),
                                            max_value=pd.Timestamp("2026-08-31"))
                date_fin    = st.date_input("Date fin récolte *",
                                            value=pd.Timestamp("2026-07-20"),
                                            min_value=pd.Timestamp("2026-06-15"),
                                            max_value=pd.Timestamp("2026-09-16"))

            st.markdown("")
            submitted = st.form_submit_button("✅ Enregistrer l'agriculteur",
                                              use_container_width=True,
                                              type="primary")

        if submitted:
            # Validation
            errors = []
            if not nom.strip():
                errors.append("Le nom est obligatoire.")
            if date_fin <= date_debut:
                errors.append("La date de fin doit être après la date de début.")
            if tonnage <= 0:
                errors.append("Le tonnage doit être positif.")

            if errors:
                for e in errors:
                    st.error(e)
            else:
                try:
                    row = {
                        "commercial":    commercial,
                        "nom":           nom.strip().upper(),
                        "region":        region,
                        "zone":          zone.strip() or None,
                        "usine":         usine,
                        "accessibilite": access,
                        "tonnage_total": float(tonnage),
                        "date_debut":    str(date_debut),
                        "date_fin":      str(date_fin),
                    }
                    sb.table("agriculteurs").insert(row).execute()
                    st.success(f"✅ Agriculteur **{nom.upper()}** ajouté avec succès !")
                    st.info("👉 Cliquez 'Régénérer le planning' dans la sidebar pour recalculer.")
                    st.cache_data.clear()
                except Exception as e:
                    st.error(f"Erreur lors de l'insertion : {e}")

    # ════════════════════════════════════════════════════════
    # SOUS-ONGLET 2 : MODIFIER UN AGRICULTEUR
    # ════════════════════════════════════════════════════════
    with a2:
        st.subheader("Modifier un agriculteur existant")

        df_agri = load_agriculteurs()
        if df_agri.empty:
            st.info("Aucun agriculteur dans la base.")
        else:
            # Search / filter
            search = st.text_input("🔍 Rechercher par nom ou commercial",
                                   placeholder="ex: MOHAMED ou FEDI")
            filtered = df_agri.copy()
            if search.strip():
                mask = (filtered["nom"].str.upper().str.contains(search.upper(), na=False) |
                        filtered["commercial"].str.upper().str.contains(search.upper(), na=False))
                filtered = filtered[mask]

            if filtered.empty:
                st.warning("Aucun résultat pour cette recherche.")
            else:
                # Select from filtered list
                options = [
                    f"[{r['id']}] {r['commercial']} — {r['nom']} ({r['tonnage_total']}t → {r['usine']})"
                    for _, r in filtered.iterrows()
                ]
                selected_label = st.selectbox("Sélectionner l'agriculteur à modifier",
                                               options)
                selected_id = int(selected_label.split("]")[0].replace("[",""))
                row_data = df_agri[df_agri["id"] == selected_id].iloc[0]

                st.markdown("---")
                st.caption(f"Modification de : **{row_data['nom']}** (ID {selected_id})")

                with st.form("form_edit"):
                    c1, c2 = st.columns(2)
                    with c1:
                        e_nom        = st.text_input("Nom", value=str(row_data["nom"]))
                        if IS_COMMERCIAL:
                            e_commercial = CURRENT_NAME
                            st.info(f"Commercial : **{CURRENT_NAME}**")
                        else:
                            e_commercial = st.selectbox("Commercial",
                                                         COMMERCIALS,
                                                         index=COMMERCIALS.index(row_data["commercial"])
                                                         if row_data["commercial"] in COMMERCIALS else 0)
                        e_usine      = st.selectbox("Usine",
                                                     USINES,
                                                     index=USINES.index(row_data["usine"])
                                                     if row_data["usine"] in USINES else 0)
                        e_tonnage    = st.number_input("Tonnage total",
                                                        min_value=10.0, max_value=10000.0,
                                                        value=float(row_data["tonnage_total"]),
                                                        step=25.0)
                    with c2:
                        e_region     = st.selectbox("Région",
                                                     REGIONS,
                                                     index=REGIONS.index(str(row_data["region"]).upper())
                                                     if str(row_data["region"]).upper() in REGIONS else 0)
                        e_zone       = st.text_input("Zone", value=str(row_data["zone"] or ""))
                        e_access     = st.selectbox("Accessibilité",
                                                     ACCESS_CODES,
                                                     index=ACCESS_CODES.index(str(row_data["accessibilite"]))
                                                     if str(row_data["accessibilite"]) in ACCESS_CODES else 0)
                        try:
                            dd_parsed = pd.to_datetime(row_data["date_debut"], errors="coerce")
                            dd = dd_parsed.date() if not pd.isna(dd_parsed) else pd.Timestamp("2026-07-01").date()
                            df_parsed = pd.to_datetime(row_data["date_fin"], errors="coerce")
                            df_ = df_parsed.date() if not pd.isna(df_parsed) else pd.Timestamp("2026-07-20").date()
                        except Exception:
                            dd  = pd.Timestamp("2026-07-01").date()
                            df_ = pd.Timestamp("2026-07-20").date()

                        e_debut = st.date_input("Date début", value=dd)
                        e_fin   = st.date_input("Date fin",   value=df_)

                    save_edit = st.form_submit_button("💾 Sauvegarder les modifications",
                                                      use_container_width=True,
                                                      type="primary")

                if save_edit:
                    if e_fin <= e_debut:
                        st.error("La date de fin doit être après la date de début.")
                    else:
                        try:
                            sb.table("agriculteurs").update({
                                "commercial":    e_commercial,
                                "nom":           e_nom.strip().upper(),
                                "region":        e_region,
                                "zone":          e_zone.strip() or None,
                                "usine":         e_usine,
                                "accessibilite": e_access,
                                "tonnage_total": float(e_tonnage),
                                "date_debut":    str(e_debut),
                                "date_fin":      str(e_fin),
                            }).eq("id", selected_id).execute()
                            st.success(f"✅ Agriculteur **{e_nom.upper()}** modifié !")
                            st.info("👉 Cliquez 'Régénérer le planning' dans la sidebar.")
                            st.cache_data.clear()
                        except Exception as e:
                            st.error(f"Erreur lors de la modification : {e}")

    # ════════════════════════════════════════════════════════
    # SOUS-ONGLET 3 : SUPPRIMER UN AGRICULTEUR
    # ════════════════════════════════════════════════════════
    with a3:
        st.subheader("Supprimer un agriculteur")
        st.warning("⚠️ La suppression est définitive. Le planning associé sera recalculé "
                   "à la prochaine régénération.")

        df_agri = load_agriculteurs()
        if df_agri.empty:
            st.info("Aucun agriculteur dans la base.")
        else:
            search_del = st.text_input("🔍 Rechercher l'agriculteur à supprimer",
                                        placeholder="ex: HSSINE BRINI")
            filtered_del = df_agri.copy()
            if search_del.strip():
                mask = (filtered_del["nom"].str.upper().str.contains(search_del.upper(), na=False) |
                        filtered_del["commercial"].str.upper().str.contains(search_del.upper(), na=False))
                filtered_del = filtered_del[mask]

            if filtered_del.empty:
                st.warning("Aucun résultat.")
            else:
                options_del = [
                    f"[{r['id']}] {r['commercial']} — {r['nom']} ({r['tonnage_total']}t)"
                    for _, r in filtered_del.iterrows()
                ]
                selected_del = st.selectbox("Sélectionner l'agriculteur à supprimer",
                                             options_del)
                sel_id_del = int(selected_del.split("]")[0].replace("[",""))
                sel_nom_del = df_agri[df_agri["id"] == sel_id_del].iloc[0]["nom"]

                st.markdown(f"**Tu vas supprimer :** {sel_nom_del} (ID {sel_id_del})")

                # Double confirmation
                confirm = st.checkbox(f"Je confirme la suppression de **{sel_nom_del}**")
                if confirm:
                    if st.button("🗑️ Supprimer définitivement",
                                 type="primary",
                                 use_container_width=True):
                        try:
                            sb.table("agriculteurs").delete().eq("id", sel_id_del).execute()
                            st.success(f"✅ Agriculteur **{sel_nom_del}** supprimé.")
                            st.info("👉 Cliquez 'Régénérer le planning' dans la sidebar.")
                            st.cache_data.clear()
                            st.rerun()
                        except Exception as e:
                            st.error(f"Erreur lors de la suppression : {e}")

    # ════════════════════════════════════════════════════════
    # SOUS-ONGLET 4 : LISTE COMPLÈTE
    # ════════════════════════════════════════════════════════
    with a4:
        st.subheader("Liste complète des agriculteurs")

        df_agri = load_agriculteurs()
        if df_agri.empty:
            st.info("Aucun agriculteur dans la base. Utilisez l'onglet 'Ajouter'.")
        else:
            # Stats rapides
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Total agriculteurs", len(df_agri))
            c2.metric("Tonnage total", f"{df_agri['tonnage_total'].sum():,.0f} t")
            c3.metric("Commercials", df_agri["commercial"].nunique())
            c4.metric("Usines", df_agri["usine"].nunique())

            st.markdown("")

            # Filtres
            col_f1, col_f2, col_f3 = st.columns(3)
            with col_f1:
                filt_comm = st.multiselect("Filtrer par commercial",
                                            sorted(df_agri["commercial"].unique()),
                                            default=sorted(df_agri["commercial"].unique()))
            with col_f2:
                filt_usine = st.multiselect("Filtrer par usine",
                                             sorted(df_agri["usine"].unique()),
                                             default=sorted(df_agri["usine"].unique()))
            with col_f3:
                sort_col = st.selectbox("Trier par",
                                         ["commercial","nom","tonnage_total","usine"])

            display_df = df_agri[
                df_agri["commercial"].isin(filt_comm) &
                df_agri["usine"].isin(filt_usine)
            ].sort_values(sort_col).reset_index(drop=True)

            # Show table
            show_cols = [c for c in
                         ["id","commercial","nom","usine","tonnage_total",
                          "accessibilite","region","zone","date_debut","date_fin"]
                         if c in display_df.columns]
            st.dataframe(
                display_df[show_cols],
                use_container_width=True,
                height=400,
                column_config={
                    "id":            st.column_config.NumberColumn("ID", width="small"),
                    "commercial":    st.column_config.TextColumn("Commercial"),
                    "nom":           st.column_config.TextColumn("Agriculteur"),
                    "usine":         st.column_config.TextColumn("Usine"),
                    "tonnage_total": st.column_config.NumberColumn("Tonnage (t)", format="%.0f t"),
                    "accessibilite": st.column_config.TextColumn("Accès"),
                    "region":        st.column_config.TextColumn("Région"),
                    "zone":          st.column_config.TextColumn("Zone"),
                    "date_debut":    st.column_config.TextColumn("Début récolte"),
                    "date_fin":      st.column_config.TextColumn("Fin récolte"),
                }
            )

            # Export
            st.download_button(
                "⬇️ Exporter la liste (CSV)",
                data=df_to_csv(display_df[show_cols]),
                file_name="agriculteurs_supabase.csv",
                mime="text/csv",
            )

# ── TAB 11: UPLOAD PLANNING ──────────────────────────────────
with tab10:
    if not UPLOAD_AVAILABLE:
        st.error("❌ Fichier `upload_tab.py` introuvable dans le dossier.")
        st.info("Mets `upload_tab.py` dans le même dossier que `dashboard_phase10.py`.")
    else:
        render_upload_tab(
            sb=get_supabase(),
            CURRENT_ROLE=CURRENT_ROLE,
            CURRENT_NAME=CURRENT_NAME,
            CURRENT_FILTER=CURRENT_FILTER,
            GLOBAL_COMMERCIAL_FARMERS=GLOBAL_COMMERCIAL_FARMERS,
            GLOBAL_COMMERCIAL_TONS=GLOBAL_COMMERCIAL_TONS,
            df_to_csv=df_to_csv,
        )