"""Tableau de bord CSSVD / Engagement producteurs — style "Crextio".

Lancer :  streamlit run app.py
Données : export KoboToolbox (.xlsx ou .csv). Déposer le fichier dans ./data/
          ou le charger via le bouton « ⚙ Paramètres ». Sans fichier, des
          données de démonstration sont utilisées.
"""
from __future__ import annotations

import glob
import io
import os
import re
import unicodedata
from datetime import timedelta

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

st.set_page_config(page_title="Ecam Track · CSSVD", page_icon="🌿", layout="wide")

GREEN = "#5aa55a"
GREEN_L = "#8fcf86"
DARK = "#2b2b2b"
GREY = "#9a9a9a"
FONT = "Outfit, sans-serif"

MOIS = ["Janvier", "Février", "Mars", "Avril", "Mai", "Juin", "Juillet", "Août",
        "Septembre", "Octobre", "Novembre", "Décembre"]
JOURS = ["Lun", "Mar", "Mer", "Jeu", "Ven", "Sam", "Dim"]

SYMPTOMES = {
    "sym_gonflement": ("Gonflement des rameaux", "🌱"),
    "sym_taches": ("Taches jaunes sur nervures", "🍂"),
    "sym_dessechement": ("Dessèchement des branches", "🪵"),
    "sym_decoloration": ("Décoloration rougeâtre", "🩸"),
    "sym_chute": ("Chute des feuilles", "🍃"),
    "sym_mortalite": ("Mortalité des cacaoyers", "🪦"),
}

# ----------------------------------------------------------------------------
# Style
# ----------------------------------------------------------------------------
CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Outfit:wght@300;400;500;600&family=DM+Serif+Display:ital@0;1&display=swap');
html, body, .stApp, [class*="css"], button, input, textarea { font-family: 'Outfit', sans-serif !important; }
.stApp { background: #6bb36a; }
header[data-testid="stHeader"] { background: transparent; }
.block-container {
  max-width: 1440px; margin: 1.6rem auto; padding: 1.8rem 2.2rem 2.4rem !important;
  background: radial-gradient(1200px 600px at 75% 10%, #e9f5e3 0%, transparent 60%),
              linear-gradient(140deg, #f1f7ee 0%, #dcefd5 55%, #c6e5be 100%);
  border-radius: 38px; box-shadow: 0 30px 80px rgba(0,0,0,.18);
}
h1,h2,h3,h4 { font-family: 'Outfit', sans-serif !important; color: #222; }

/* Cartes */
[class*="st-key-card"] {
  background: rgba(255,255,255,.55); border: 1px solid rgba(255,255,255,.75);
  border-radius: 28px; padding: 20px 22px; backdrop-filter: blur(12px);
  box-shadow: 0 6px 24px rgba(60,110,60,.08);
}
[class*="st-key-dark"] {
  background: #2c2c2c; border-radius: 28px; padding: 22px 20px; color: #fff;
  box-shadow: 0 14px 40px rgba(0,0,0,.25);
}
[class*="st-key-dark"] * { color: #fff; }
.st-key-card_profile {
  background: linear-gradient(170deg, #6f9a68 0%, #3e5c3a 55%, #1f2d1d 100%) !important;
  min-height: 300px; position: relative; overflow: hidden;
}
.h-equal { min-height: 300px; }

/* Typo */
.logo { display:inline-block; border:1.5px solid #3a3a3a; border-radius: 999px; padding: 6px 20px;
        font-size: 1.55rem; font-weight: 400; color:#222; letter-spacing:-.5px; }
.welcome { font-size: 3.1rem; font-weight: 400; letter-spacing: -1.5px; color:#1f1f1f; margin: 6px 0 4px; }
.card-title { font-size: 1.35rem; font-weight: 400; color:#222; margin:0; }
.muted { color:#7b7b7b; font-size:.82rem; }
.big { font-size: 2.3rem; font-weight: 300; letter-spacing:-1px; line-height:1; }
.arrow-btn { float:right; width:38px; height:38px; border-radius:50%; background:#fff; display:flex;
             align-items:center; justify-content:center; font-size:1rem; color:#333; box-shadow:0 2px 6px rgba(0,0,0,.06);}

/* Pills de progression */
.prow { display:flex; gap:8px; align-items:flex-end; }
.pcol { display:flex; flex-direction:column; gap:6px; }
.plabel { font-size:.85rem; color:#333; }
.pill { height:40px; border-radius:999px; display:flex; align-items:center; padding:0 16px; font-size:.8rem; }
.pill.dark { background:#2c2c2c; color:#fff; }
.pill.green { background:#5aa55a; color:#fff; }
.pill.outline { border:1.5px solid #9a9a9a; color:#333; background:transparent; }
.pill.striped { padding:0; overflow:hidden; background: repeating-linear-gradient(-55deg, rgba(255,255,255,.0) 0 6px, rgba(255,255,255,.75) 6px 9px); border:1px solid rgba(255,255,255,.6);}
.pill.striped > div { height:100%; border-radius:999px; display:flex; align-items:center; padding-left:16px; color:#fff;
  background: repeating-linear-gradient(-55deg, #5aa55a 0 8px, #6cb46a 8px 12px); }

/* Pills KPI (Engagement / Inspection / Promesse) */
.kp-row { display:flex; gap:42px; align-items:flex-end; margin-top:10px; }
.kp { display:flex; flex-direction:column; align-items:center; gap:8px; flex:1; max-width:190px; }
.kp-lab { font-style:italic; font-weight:500; font-size:1rem; color:#2a2a2a; white-space:nowrap; }
.kp-pill { position:relative; width:100%; height:36px; border-radius:999px; overflow:hidden; display:flex; align-items:center;
           box-shadow: 0 2px 6px rgba(0,0,0,.12); }
.kp-pill span { position:relative; z-index:1; padding-left:14px; font-weight:700; font-size:1.25rem; font-family: Arial, sans-serif; }
.kp-striped { background: repeating-linear-gradient(-55deg, #ffffff 0 7px, #e6efe3 7px 10px); }
.kp-striped span { color:#111; }
.kp-fill { position:absolute; left:0; top:0; bottom:0; min-width:7px; border-radius:999px 0 0 999px;
           background: linear-gradient(180deg, #4aa64f, #2e7d32); }
.kp-green { background: linear-gradient(180deg, #2f8a35 0%, #1f6a24 100%); }
.kp-green span, .kp-dark span { color:#fff; }
.kp-dark { background: linear-gradient(180deg, #3a3a3a 0%, #0f0f0f 100%); }

/* Compteurs x / objectif */
.cnt-row { display:grid; grid-template-columns: repeat(5, 1fr); gap:14px; margin: 18px 0 6px; }
.cnt { position:relative; text-align:center; padding:18px 12px 16px; border-radius:24px;
       background: linear-gradient(160deg, rgba(255,255,255,.75), rgba(255,255,255,.35));
       border:1px solid rgba(255,255,255,.85); box-shadow: 0 8px 24px rgba(50,100,50,.10); overflow:hidden; }
.cnt::before { content:""; position:absolute; left:0; top:0; height:4px; width:100%; background:var(--c); opacity:.9; }
.cnt .n { font-family:'DM Serif Display', Georgia, serif !important; font-size:2.6rem; line-height:1; color:#1d1d1d; letter-spacing:-.5px; }
.cnt .n small { font-size:1.35rem; color:#8a8a8a; margin-left:2px; font-family:'DM Serif Display', Georgia, serif !important; }
.cnt .lb { margin-top:8px; font-style:italic; font-weight:500; font-size:.95rem; color:#333; line-height:1.15; }
.cnt .tr { margin:12px auto 0; height:6px; width:82%; border-radius:99px; background:rgba(0,0,0,.07); overflow:hidden; }
.cnt .tr > div { height:100%; border-radius:99px; background:var(--c); }
.cnt .pc { margin-top:6px; font-size:.75rem; color:#6b6b6b; }
@media (max-width: 900px) { .cnt-row { grid-template-columns: repeat(2, 1fr); } }

/* Boutons symptômes (liste sombre cliquable) */
[class*="st-key-dark"] button { background:#3a3a3a !important; border:none !important; border-radius:999px !important;
  min-height:46px; margin:1px 0; transition: background .15s, transform .15s; }
[class*="st-key-dark"] button:hover { background:#5aa55a !important; transform: translateX(3px); }
[class*="st-key-dark"] button:disabled { background:#333 !important; opacity:.45; }
[class*="st-key-dark"] button > div { justify-content:flex-start !important; width:100%; }
[class*="st-key-dark"] button p { color:#fff !important; font-size:.86rem; text-align:left; }

/* Gros chiffres */
.stat { display:flex; align-items:flex-end; gap:10px; }
.stat .ic { width:30px; height:30px; border-radius:50%; background:rgba(255,255,255,.6); display:flex; align-items:center; justify-content:center; font-size:.85rem; margin-bottom:8px;}
.stat .num { font-size:3.6rem; font-weight:300; letter-spacing:-2px; line-height:1; color:#1f1f1f; }
.stat .lab { font-size:.8rem; color:#333; margin-top:-2px; }

/* Profil */
.prof-av { width:150px; height:150px; border-radius:50%; margin: 18px auto 0; display:flex; align-items:center; justify-content:center;
           font-size:3.4rem; font-weight:300; color:#fff; background: radial-gradient(circle at 30% 30%, rgba(255,255,255,.35), rgba(255,255,255,.05));
           border:1px solid rgba(255,255,255,.35); }
.prof-foot { display:flex; justify-content:space-between; align-items:flex-end; margin-top:34px; }
.prof-name { font-size:1.35rem; font-weight:300; color:#fff; line-height:1.1; }
.prof-role { font-size:.78rem; color:rgba(255,255,255,.75); }
.prof-pill { border:1.5px solid rgba(255,255,255,.85); border-radius:999px; padding:8px 16px; color:#fff; font-size:.95rem; white-space:nowrap; backdrop-filter: blur(6px); background:rgba(255,255,255,.08);}

/* Barres segmentées */
.seg { display:flex; gap:6px; margin-top:14px; }
.seg > div { display:flex; flex-direction:column; gap:6px; min-width:40px; }
.seg .bar { height:38px; border-radius:12px; display:flex; align-items:center; padding-left:10px; font-size:.78rem; color:#fff; }

/* Liste sombre */
.task { display:flex; align-items:center; gap:12px; padding:9px 0; }
.task .ic { width:42px; height:42px; border-radius:50%; display:flex; align-items:center; justify-content:center; font-size:1.05rem; flex:none;}
.task .ic.on { background:#fff; } .task .ic.off { background:#4a4a4a; }
.task .t1 { font-size:.9rem; } .task .t2 { font-size:.72rem; color:#a8a8a8 !important; }
.task .ck { margin-left:auto; width:18px; height:18px; border-radius:50%; display:flex; align-items:center; justify-content:center; font-size:.65rem; }
.task .ck.on { background:#5aa55a; } .task .ck.off { background:#5a5a5a; }
.task.dim .t1 { color:#9d9d9d !important; }

/* Calendrier */
.cal-head { display:flex; justify-content:space-between; align-items:center; margin-bottom:10px; }
.cal-chip { background:#fff; border-radius:999px; padding:4px 14px; font-size:.78rem; color:#333; }
.cal { display:grid; gap:0; position:relative; }
.cal .dh { text-align:center; font-size:.85rem; color:#888; padding-bottom:6px; }
.cal .dh b { display:block; font-weight:400; color:#888; }
.cal .dh.today, .cal .dh.today b { color:#222; font-weight:500; }
.cal .hr { font-size:.78rem; color:#333; padding-top:10px; }
.cal .cell { border-left:1.5px dashed rgba(0,0,0,.12); min-height:48px; padding:4px 5px; }
.ev { border-radius:14px; padding:6px 10px; font-size:.72rem; line-height:1.15; margin-bottom:4px; box-shadow:0 2px 8px rgba(0,0,0,.06);}
.ev.g { background:#5aa55a; color:#fff; } .ev.w { background:#fff; color:#222; }
.ev small { display:block; opacity:.8; font-size:.65rem; }

/* Liste repliable / mini barres */
[data-testid="stExpander"] details { border:none !important; background:transparent !important; }
[data-testid="stExpander"] summary { border-bottom:1.5px dashed rgba(0,0,0,.15); padding-left:0; }
[data-testid="stExpander"] summary p { font-size:1rem; color:#222; }
.mini { display:flex; align-items:center; gap:10px; margin:6px 0; font-size:.82rem; }
.mini .nm { width:46%; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; color:#333;}
.mini .tr { flex:1; height:8px; border-radius:99px; background:rgba(0,0,0,.07); overflow:hidden; }
.mini .tr > div { height:100%; border-radius:99px; background:#5aa55a; }
.mini .vl { width:38px; text-align:right; color:#555; }

/* Navigation */
.st-key-nav [data-testid="stButtonGroup"] { background: rgba(255,255,255,.6); border-radius:999px; padding:4px; display:inline-flex; }
.st-key-nav button { border:none !important; border-radius:999px !important; background:transparent !important; font-size:.85rem; padding: 4px 14px !important; min-height:34px; }
.st-key-nav button[kind="segmented_controlActive"], .st-key-nav button[data-testid="stBaseButton-segmented_controlActive"] { background:#2c2c2c !important; color:#fff !important; }
.st-key-nav button[kind="segmented_controlActive"] p, .st-key-nav button[data-testid="stBaseButton-segmented_controlActive"] p { color:#fff !important; }
[data-testid="stPopover"] button { border-radius:999px !important; background:rgba(255,255,255,.7) !important; border:none !important; }
.st-key-card_table [data-testid="stDataFrame"] { border-radius:16px; overflow:hidden; }
</style>
"""


def html(s: str) -> None:
    """Rend du HTML sans que Markdown n'interprète l'indentation."""
    st.markdown("\n".join(line.strip() for line in s.splitlines()), unsafe_allow_html=True)


def fmt(x: float, d: int = 0) -> str:
    return f"{x:,.{d}f}".replace(",", " ").replace(".", ",")


def pct(x: float) -> str:
    return "–" if pd.isna(x) else f"{x * 100:.0f}%"


# ----------------------------------------------------------------------------
# Chargement & normalisation de l'export Kobo
# ----------------------------------------------------------------------------
def _norm(s: str) -> str:
    s = unicodedata.normalize("NFKD", str(s)).encode("ascii", "ignore").decode()
    return re.sub(r"\s+", " ", s.lower()).strip()


# nom canonique -> (motif dans le nom de colonne normalisé, correspondance exacte ?)
SPEC = {
    "start": ("start", True),
    "today": ("today", True),
    "submission": ("_submission_time", False),
    "agent_id": ("agent/id de l'agent", False),
    "agent_nom": ("agent/nom_agent", False),
    "prod_code": ("producteur/code planteur", False),
    "prod_nom": ("producteur/nom_prod", False),
    "genre": ("producteur/genre", False),
    "section": ("producteur/section_prod", False),
    "localite": ("producteur/localite", False),
    "campement": ("producteur/campement", False),
    "annee": ("producteur/annee_naissance", False),
    "sup_parc": ("producteur/superficie_parcelle", False),
    "sup_tot": ("producteur/superficie_totale", False),
    "plantation": ("code plantation", False),
    "menage": ("type de menage", False),
    "engagement": ("donne son engagement", False),
    "doc_promesse": ("autres documents disponibles/formulaire de promesse", False),
    "doc_fiche": ("autres documents disponibles/fiche d'inspection", False),
    "doc_aucun": ("autres documents disponibles/aucun autre", False),
    "cssvd": ("signes de cssvd dans la parcelle", False),
    "sym_gonflement": ("observe(s)/gonflement", False),
    "sym_taches": ("observe(s)/taches jaunes", False),
    "sym_dessechement": ("observe(s)/dessechement", False),
    "sym_decoloration": ("observe(s)/decoloration", False),
    "sym_chute": ("observe(s)/chute des feuilles", False),
    "sym_mortalite": ("observe(s)/mortalite", False),
    "lat": ("_gps_parcelle_latitude", False),
    "lon": ("_gps_parcelle_longitude", False),
}


def _find(cols: list[str], pattern: str, exact: bool) -> str | None:
    hits = []
    for c in cols:
        n = _norm(c)
        if "${" in n or n.endswith("_url"):
            continue  # colonnes d'affichage / liens photo
        if (exact and n == pattern) or (not exact and pattern in n):
            hits.append(c)
    return min(hits, key=len) if hits else None


def _num(s: pd.Series) -> pd.Series:
    if not pd.api.types.is_numeric_dtype(s):
        s = s.astype(str).str.replace(",", ".", regex=False).str.strip()
    return pd.to_numeric(s, errors="coerce")


def _oui(s: pd.Series) -> pd.Series:
    t = s.astype("string").str.strip().str.lower()
    out = pd.Series(np.nan, index=s.index, dtype="object")
    out[t.str.startswith("oui").fillna(False).astype(bool)] = True
    out[t.str.startswith("non").fillna(False).astype(bool)] = False
    return out


def normalize(raw: pd.DataFrame) -> pd.DataFrame:
    cols = list(raw.columns)
    out = pd.DataFrame(index=raw.index)
    for key, (pat, exact) in SPEC.items():
        c = _find(cols, pat, exact)
        out[key] = raw[c] if c is not None else np.nan

    when = out["submission"].where(out["submission"].notna(), out["start"])
    out["date"] = pd.to_datetime(when, errors="coerce", utc=True).dt.tz_convert(None)
    for c in ["annee", "sup_parc", "sup_tot", "lat", "lon", "doc_promesse", "doc_fiche",
              "doc_aucun", *SYMPTOMES]:
        out[c] = _num(out[c])
    out["engagement"] = _oui(out["engagement"])
    out["cssvd"] = _oui(out["cssvd"])
    for c in ["agent_id", "agent_nom", "prod_code", "prod_nom", "genre", "section",
              "localite", "campement", "menage", "plantation"]:
        out[c] = out[c].astype("string").str.strip()
    out["section"] = out["section"].str.upper()
    out["localite"] = out["localite"].str.title()
    # liens des photos de chaque symptôme (colonnes « Photo : … _URL »)
    for k, pat in PHOTO_PAT.items():
        c = next((c for c in cols if pat in _norm(c) and _norm(c).endswith("_url")), None)
        out[f"photo_{k}"] = raw[c].astype("string").str.strip() if c is not None else pd.NA
    return out.drop(columns=["submission", "start", "today"])


PHOTO_PAT = {
    "sym_gonflement": "photo : gonflement",
    "sym_taches": "photo : taches jaunes",
    "sym_dessechement": "photo : dessechement",
    "sym_decoloration": "photo : decoloration",
    "sym_chute": "photo : chute des feuilles",
    "sym_mortalite": "photo : mortalite",
}


def kobo_token() -> str:
    tok = st.session_state.get("kobo_token") or os.environ.get("KOBO_TOKEN", "")
    if not tok:
        try:
            tok = st.secrets.get("KOBO_TOKEN", "")
        except Exception:  # pas de fichier secrets.toml
            tok = ""
    return tok.strip()


@st.cache_data(show_spinner=False, ttl=3600, max_entries=500)
def fetch_image(url: str, token: str) -> bytes | None:
    """Télécharge une photo Kobo et la renvoie en JPEG (None si inaccessible / format illisible)."""
    import requests
    from PIL import Image

    try:  # photos iPhone (.heic)
        from pillow_heif import register_heif_opener
        register_heif_opener()
    except ImportError:
        pass
    try:
        headers = {"User-Agent": "Mozilla/5.0 (dashboard)"}
        if token:
            headers["Authorization"] = f"Token {token}"
        r = requests.get(url, headers=headers, timeout=20)
        if r.status_code != 200:
            return None
        img = Image.open(io.BytesIO(r.content))
        img = img.convert("RGB")
        img.thumbnail((1200, 1200))
        buf = io.BytesIO()
        img.save(buf, format="JPEG", quality=85)
        return buf.getvalue()
    except Exception:
        return None


@st.dialog("Photos du symptôme", width="large")
def show_gallery(k: str, data: pd.DataFrame) -> None:
    name, icon = SYMPTOMES[k]
    col = f"photo_{k}"
    rows = data[data[k] == 1]
    photos = rows[rows[col].notna() & (rows[col] != "")] if col in rows else rows.iloc[0:0]
    photos = photos.sort_values("date", ascending=False)
    html(f"<div style='font-size:1.5rem'>{icon} {name}</div>"
         f"<div class='muted'>{len(rows)} parcelle(s) concernée(s) · {len(photos)} photo(s)</div>")
    if photos.empty:
        st.info("Aucune photo disponible pour ce symptôme dans les données filtrées.")
        return
    limit = 24
    token = kobo_token()
    fails = 0
    grid = st.columns(3)
    for i, r in enumerate(photos.head(limit).itertuples()):
        url = getattr(r, col)
        with grid[i % 3]:
            with st.spinner(""):
                img = fetch_image(url, token)
            if img:
                st.image(img, width="stretch")
            else:
                fails += 1
                st.link_button("📷 Ouvrir la photo ↗", url, width="stretch")
            d = r.date.strftime("%d/%m %H:%M") if pd.notna(r.date) else ""
            st.caption(f"**{r.prod_nom}** · {r.prod_code}  \n{r.section} · {r.agent_id} · {d}")
    if len(photos) > limit:
        st.caption(f"{limit} photos les plus récentes affichées sur {len(photos)}.")
    if fails:
        st.warning(f"{fails} photo(s) n'ont pas pu être affichées ici (accès Kobo protégé ou format HEIC). "
                   "Renseignez votre **jeton API Kobo** dans ⚙ Paramètres, ou cliquez sur « Ouvrir la photo ».")


def cssvd_table(data: pd.DataFrame) -> pd.DataFrame:
    """Liste des parcelles avec signes de swollen shoot (1 ligne = 1 parcelle enquêtée)."""
    d = data[data["cssvd"] == True].copy()  # noqa: E712
    sym_cols = list(SYMPTOMES)
    d["symptomes"] = d[sym_cols].apply(
        lambda r: ", ".join(SYMPTOMES[k][0] for k in sym_cols if r[k] == 1), axis=1) if len(d) else ""
    d["nb_symptomes"] = (d[sym_cols] == 1).sum(axis=1)
    out = pd.DataFrame({
        "Date enquête": d["date"].dt.strftime("%d/%m/%Y %H:%M"),
        "ID coach": d["agent_id"], "Nom coach": d["agent_nom"],
        "Code planteur": d["prod_code"], "Nom producteur": d["prod_nom"], "Genre": d["genre"],
        "Année naissance": d["annee"].astype("Int64"),
        "Section": d["section"], "Localité": d["localite"], "Lieu de résidence": d["campement"],
        "Code parcelle": d["plantation"],
        "Superficie parcelle (ha)": d["sup_parc"].round(2), "Superficie totale (ha)": d["sup_tot"].round(2),
        "Nb symptômes": d["nb_symptomes"], "Symptômes observés": d["symptomes"],
    })
    for k in sym_cols:
        out[SYMPTOMES[k][0]] = np.where(d[k] == 1, "Oui", "Non")
    out["Engagement"] = d["engagement"].map({True: "Oui", False: "Non"})
    out["Fiche inspection CSSVD"] = np.where(d["doc_fiche"] == 1, "Oui", "Non")
    out["Promesse de taille"] = np.where(d["doc_promesse"] == 1, "Oui", "Non")
    out["Latitude"], out["Longitude"] = d["lat"], d["lon"]
    for k in sym_cols:
        if f"photo_{k}" in d:
            out[f"Photo – {SYMPTOMES[k][0]}"] = d[f"photo_{k}"]
    out = out.dropna(axis=1, how="all")
    return out.sort_values(["Section", "ID coach", "Nom producteur"], na_position="last").reset_index(drop=True)


def to_xlsx(t: pd.DataFrame, sheet: str = "Producteurs CSSVD") -> bytes:
    from openpyxl.styles import Alignment, Font, PatternFill

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        t.to_excel(xw, index=False, sheet_name=sheet)
        ws = xw.sheets[sheet]
        head = PatternFill("solid", fgColor="2E7D32")
        for c in ws[1]:
            c.font, c.fill = Font(bold=True, color="FFFFFF"), head
            c.alignment = Alignment(wrap_text=True, vertical="center")
        ws.row_dimensions[1].height = 32
        ws.freeze_panes = "F2"
        ws.auto_filter.ref = ws.dimensions
        for i, col in enumerate(t.columns, start=1):
            letter = ws.cell(1, i).column_letter
            if str(col).startswith("Photo"):
                ws.column_dimensions[letter].width = 16
                for r in range(2, len(t) + 2):
                    cell = ws.cell(r, i)
                    if cell.value:
                        cell.hyperlink, cell.value, cell.style = cell.value, "Voir la photo", "Hyperlink"
            else:
                width = max([len(str(col))] + [len(str(v)) for v in t[col].head(500)]) + 2
                ws.column_dimensions[letter].width = min(max(width, 10), 45)
    return buf.getvalue()


def symptom_buttons(data: pd.DataFrame, prefix: str) -> None:
    """Un bouton par symptôme ; un clic ouvre la galerie de photos."""
    observed = {k: int((data[k] == 1).sum()) for k in SYMPTOMES}
    for k, v in sorted(observed.items(), key=lambda kv: -kv[1]):
        col = f"photo_{k}"
        n_ph = int((data[col].notna() & (data[k] == 1)).sum()) if col in data else 0
        name, icon = SYMPTOMES[k]
        label = f"{icon}  {name}  ·  {v} parcelle{'s' if v > 1 else ''}  ·  📷 {n_ph}"
        if st.button(label, key=f"{prefix}_{k}", width="stretch", disabled=v == 0):
            show_gallery(k, data)


@st.cache_data(show_spinner=False)
def read_bytes(data: bytes, name: str) -> pd.DataFrame:
    if name.lower().endswith((".xlsx", ".xls")):
        return pd.read_excel(io.BytesIO(data))
    return pd.read_csv(io.BytesIO(data), sep=None, engine="python", encoding_errors="replace")


@st.cache_data(show_spinner=False)
def demo_data(n: int = 160, seed: int = 7) -> pd.DataFrame:
    rng = np.random.default_rng(seed)
    coachs = {
        # noms fictifs (données de démonstration uniquement)
        "COACH901": "COACH DEMO ALPHA", "COACH902": "COACH DEMO BETA",
        "COACH903": "COACH DEMO GAMMA", "COACH904": "COACH DEMO DELTA",
        "COACH905": "COACH DEMO EPSILON", "COACH906": "COACH DEMO ZETA",
        "COACH907": "COACH DEMO ETA",
    }
    sections = {
        "CHEICKVILLE": ["Cheickville"], "GNANGOMIANKRO": ["Souangakro", "Gnangomiankro", "Robertkro", "Prefetkro"],
        "PETIT OUELLE": ["Yayakro", "Comikro", "Pogreagui", "N'zikro"], "BROUKRO": ["Méagui", "Broukro"],
        "GNANMIENKOUAKOUKRO": ["Ouattarakro", "Kouassikomenankro", "Pascalkro"], "KONANDANKRO": ["Dinankro"],
    }
    prenoms = ["KOFFI", "KOUADIO", "N'GUESSAN", "KONAN", "YAO", "KOUAME", "ADJOUA", "AYA", "AMOIN", "AFFOUE"]
    noms = ["N'DRI", "KOUAKOU", "N'ZUE", "BROU", "ZOUZOU", "OUATTARA", "KABORE", "KOMENAN", "ABIA", "YACOUBA"]
    start = pd.Timestamp("2026-11-14 07:30")
    rows = []
    for i in range(n):
        cid = rng.choice(list(coachs))
        sec = rng.choice(list(sections))
        fem = rng.random() < .22
        code = f"ECA{rng.integers(80, 4400):04d}"
        cssvd = rng.random() < .48
        sy = {k: int(cssvd and rng.random() < p) for k, p in zip(SYMPTOMES, [.55, .45, .6, .25, .5, .35])}
        promesse, fiche = int(rng.random() < .8), int(rng.random() < .82)
        has_gps = rng.random() < .6
        sup = round(float(rng.gamma(2.2, 2.2)) + .8, 4)
        rows.append({
            "agent_id": cid, "agent_nom": coachs[cid], "prod_code": code,
            "prod_nom": f"{rng.choice(noms)} {rng.choice(prenoms)}", "genre": "Femmes" if fem else "Hommes",
            "section": sec, "localite": rng.choice(sections[sec]),
            "campement": rng.choice(["Campement", "Village", "Commune"], p=[.7, .22, .08]),
            "annee": int(rng.integers(1950, 1998)), "sup_parc": sup, "sup_tot": sup,
            "plantation": f"{code}-1",
            "menage": "Femme inscrite au registre, conjoint" if fem else rng.choice(
                ["Homme inscrit au registre, première épouse", "Homme inscrit, première épouse"]),
            "engagement": bool(rng.random() < .93), "doc_promesse": promesse, "doc_fiche": fiche,
            "doc_aucun": int(not (promesse or fiche)), "cssvd": cssvd, **sy,
            "lat": 5.40 + rng.normal(0, .035) if has_gps else np.nan,
            "lon": -6.58 + rng.normal(0, .08) if has_gps else np.nan,
            "date": start + timedelta(days=int(rng.integers(0, 7)), hours=int(rng.integers(0, 11)),
                                      minutes=int(rng.integers(0, 60))),
        })
    df = pd.DataFrame(rows)
    for c in ["agent_id", "agent_nom", "prod_code", "prod_nom", "genre", "section", "localite",
              "campement", "menage", "plantation"]:
        df[c] = df[c].astype("string")
    return df


# ----------------------------------------------------------------------------
# Données partagées : l'administrateur publie, tout le monde voit la même version
#   1) GitHub (dépôt privé) si GITHUB_TOKEN + DATA_REPO sont configurés  → permanent
#   2) sinon disque du serveur (data/shared)                              → perdu au redémarrage
# ----------------------------------------------------------------------------
HERE = os.path.dirname(os.path.abspath(__file__))
SHARED_DIR = os.path.join(HERE, "data", "shared")
SHARED_FILE = os.path.join(SHARED_DIR, "export_kobo.bin")
SHARED_META = os.path.join(SHARED_DIR, "meta.json")


def secret(key: str, default: str = "") -> str:
    try:
        v = st.secrets.get(key, None)
    except Exception:  # pas de secrets.toml
        v = None
    return str(v if v is not None else os.environ.get(key, default)).strip()


def gh_cfg() -> tuple[str, str, str, str] | None:
    tok, repo = secret("GITHUB_TOKEN"), secret("DATA_REPO")
    if not (tok and repo):
        return None
    return tok, repo, secret("DATA_PATH", "export_kobo.xlsx"), secret("DATA_BRANCH", "main")


def _gh(method: str, path: str, **kw):
    import requests

    tok, repo, _, _ = gh_cfg()
    headers = {"Authorization": f"Bearer {tok}", "X-GitHub-Api-Version": "2022-11-28",
               "Accept": kw.pop("accept", "application/vnd.github+json")}
    return requests.request(method, f"https://api.github.com/repos/{repo}/{path}", headers=headers,
                            timeout=60, **kw)


@st.cache_data(ttl=20, show_spinner=False)
def remote_version() -> dict | None:
    """Dernière version publiée sur GitHub (sha + date + auteur), vérifiée toutes les 20 s au plus."""
    cfg = gh_cfg()
    if not cfg:
        return None
    _, _, path, branch = cfg
    r = _gh("GET", f"commits?path={path}&sha={branch}&per_page=1")
    if r.status_code != 200 or not r.json():
        return None
    c = r.json()[0]
    return {"sha": c["sha"], "date": c["commit"]["committer"]["date"], "message": c["commit"]["message"]}


def gh_download() -> bytes | None:
    _, _, path, branch = gh_cfg()
    r = _gh("GET", f"contents/{path}?ref={branch}", accept="application/vnd.github.raw")
    return r.content if r.status_code == 200 else None


def gh_upload(data: bytes, message: str) -> tuple[bool, str]:
    import base64

    _, _, path, branch = gh_cfg()
    cur = _gh("GET", f"contents/{path}?ref={branch}")
    body = {"message": message, "content": base64.b64encode(data).decode(), "branch": branch}
    if cur.status_code == 200:
        body["sha"] = cur.json()["sha"]
    r = _gh("PUT", f"contents/{path}", json=body)
    if r.status_code in (200, 201):
        return True, "ok"
    return False, f"GitHub {r.status_code} : {r.text[:200]}"


def _read_meta() -> dict:
    import json

    try:
        with open(SHARED_META, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def _write_local(data: bytes, meta: dict) -> None:
    import json

    os.makedirs(SHARED_DIR, exist_ok=True)
    tmp = SHARED_FILE + ".tmp"
    with open(tmp, "wb") as f:
        f.write(data)
    os.replace(tmp, SHARED_FILE)
    with open(SHARED_META, "w", encoding="utf-8") as f:
        json.dump(meta, f, ensure_ascii=False)


def publish(data: bytes, name: str, author: str = "administrateur") -> tuple[bool, str]:
    """Publie un nouvel export pour tous les utilisateurs."""
    now = pd.Timestamp.now(tz="UTC").isoformat()
    msg = f"Mise à jour des données : {name} ({author})"
    if gh_cfg():
        ok, err = gh_upload(data, msg)
        if not ok:
            return False, err
        remote_version.clear()
        v = remote_version() or {}
        _write_local(data, {"version": v.get("sha", now), "name": name, "date": now})
        return True, "Données publiées pour tous les utilisateurs (sauvegarde GitHub)."
    _write_local(data, {"version": now, "name": name, "date": now})
    return True, "Données publiées pour tous les utilisateurs (stockage temporaire du serveur)."


def current_version() -> str | None:
    """Identifiant de la version partagée actuelle (change à chaque publication)."""
    v = remote_version()
    if v:
        return v["sha"]
    return _read_meta().get("version")


def sync_shared() -> dict:
    """S'assure que la copie locale correspond à la dernière version publiée ; renvoie ses métadonnées."""
    v = remote_version()
    meta = _read_meta()
    if v and (meta.get("version") != v["sha"] or not os.path.exists(SHARED_FILE)):
        data = gh_download()
        if data:
            name = v["message"].split(":", 1)[-1].split("(")[0].strip() or "export_kobo.xlsx"
            meta = {"version": v["sha"], "name": name, "date": v["date"]}
            _write_local(data, meta)
    return meta if os.path.exists(SHARED_FILE) else {}


@st.cache_data(show_spinner="Chargement des données…", max_entries=3)
def load_shared(version: str) -> tuple[pd.DataFrame, pd.DataFrame]:
    with open(SHARED_FILE, "rb") as f:
        data = f.read()
    name = "x.xlsx" if data[:2] == b"PK" or data[:4] == b"\xd0\xcf\x11\xe0" else "x.csv"  # xlsx / xls / csv
    raw = read_bytes(data, name)
    return normalize(raw), raw


def load() -> tuple[pd.DataFrame, pd.DataFrame, str, dict]:
    meta = sync_shared()
    if meta:
        norm, raw = load_shared(meta["version"])
        return norm, raw, meta.get("name", "export Kobo"), meta
    files = sorted(glob.glob(os.path.join(HERE, "data", "*.xlsx")) + glob.glob(os.path.join(HERE, "data", "*.csv")),
                   key=os.path.getmtime, reverse=True)
    if files:
        with open(files[0], "rb") as f:
            raw = read_bytes(f.read(), files[0])
        return normalize(raw), raw, os.path.basename(files[0]), {}
    d = demo_data()
    return d, d, "Données de démonstration", {}


def is_admin() -> bool:
    return bool(st.session_state.get("is_admin"))


def admin_panel() -> None:
    import hmac

    pwd = secret("ADMIN_PASSWORD")
    st.markdown("**🔐 Espace administrateur**")
    if not pwd:
        st.caption("Ajoutez `ADMIN_PASSWORD` dans les Secrets pour activer la publication des données.")
        return
    if not is_admin():
        with st.form("login", border=False):
            p = st.text_input("Mot de passe", type="password", label_visibility="collapsed",
                              placeholder="Mot de passe administrateur")
            if st.form_submit_button("Se connecter", width="stretch"):
                if hmac.compare_digest(p.encode(), pwd.encode()):
                    st.session_state["is_admin"] = True
                    st.rerun()
                else:
                    st.error("Mot de passe incorrect.")
        return
    up = st.file_uploader("Nouvel export KoboToolbox (.xlsx / .csv)", type=["xlsx", "xls", "csv"], key="upload")
    if up is not None:
        try:
            test = normalize(read_bytes(up.getvalue(), up.name))
            st.caption(f"✓ {len(test)} lignes · {test['prod_code'].nunique()} producteurs détectés")
            ok_file = len(test) > 0
        except Exception as e:  # fichier illisible
            st.error(f"Fichier illisible : {e}")
            ok_file = False
        if st.button("📤 Publier pour tous les utilisateurs", type="primary", width="stretch", disabled=not ok_file):
            with st.spinner("Publication…"):
                ok, msg = publish(up.getvalue(), up.name)
            (st.success if ok else st.error)(msg)
            if ok:
                st.session_state.pop("data_version", None)
                st.rerun()
    st.caption("Stockage : " + ("**GitHub privé** ✓ (permanent)" if gh_cfg() else
                                "⚠ **serveur uniquement** — perdu au redémarrage. Configurez `GITHUB_TOKEN` et `DATA_REPO`."))
    st.text_input("Jeton API Kobo (photos)", type="password", key="kobo_token",
                  help="Facultatif si KOBO_TOKEN est défini dans les Secrets.")
    if st.button("Se déconnecter", width="stretch"):
        st.session_state["is_admin"] = False
        st.rerun()


@st.fragment(run_every=30)
def watch_updates() -> None:
    """Vérifie toutes les 30 s si l'administrateur a publié de nouvelles données ; recharge si oui."""
    try:
        v = current_version()
    except Exception:
        return
    seen = st.session_state.get("data_version")
    if v and seen and v != seen:
        st.session_state["data_version"] = v
        st.toast("🔄 Nouvelles données publiées — mise à jour du tableau de bord.")
        st.rerun(scope="app")


# ----------------------------------------------------------------------------
# Graphiques
# ----------------------------------------------------------------------------
def base_layout(fig: go.Figure, h: int) -> go.Figure:
    fig.update_layout(height=h, margin=dict(l=0, r=0, t=10, b=0), paper_bgcolor="rgba(0,0,0,0)",
                      plot_bgcolor="rgba(0,0,0,0)", font=dict(family=FONT, color="#555"), showlegend=False)
    return fig


CFG = {"displayModeBar": False}


def activity_chart(df: pd.DataFrame) -> tuple[go.Figure, int]:
    end = df["date"].max().normalize() if df["date"].notna().any() else pd.Timestamp.today().normalize()
    days = pd.date_range(end - timedelta(days=6), end)
    counts = df.groupby(df["date"].dt.normalize()).size().reindex(days, fill_value=0)
    x = [JOURS[d.weekday()][0] for d in days]
    xi = list(range(7))
    imax = int(np.argmax(counts.values))
    colors = [GREEN if i == imax else DARK for i in xi]
    fig = go.Figure()
    fig.add_bar(x=xi, y=counts.values, width=.13, marker=dict(color=colors, cornerradius=8),
                hovertemplate="%{customdata}<br>%{y} fiches<extra></extra>",
                customdata=[d.strftime("%d/%m") for d in days])
    fig.add_scatter(x=xi, y=[0] * 7, mode="markers", marker=dict(size=7, color=colors), hoverinfo="skip")
    top = counts.max() if counts.max() > 0 else 1
    fig.add_annotation(x=imax, y=top, yshift=16, text=f"{counts.values[imax]} fiches", showarrow=False,
                       bgcolor=GREEN, font=dict(color="white", size=11), borderpad=4)
    fig.update_xaxes(tickvals=xi, ticktext=x, showgrid=False, zeroline=False)
    fig.update_yaxes(visible=False, range=[-top * .08, top * 1.3])
    return base_layout(fig, 190), int(counts.sum())


def ring_chart(rate: float, label: str) -> go.Figure:
    r = 0 if pd.isna(rate) else float(rate)
    fig = go.Figure(go.Pie(values=[r, 1 - r], hole=.8, sort=False, direction="clockwise",
                           marker=dict(colors=[GREEN, "rgba(0,0,0,.07)"]), textinfo="none", hoverinfo="skip"))
    fig.add_annotation(text=f"<span style='font-size:40px'>{r * 100:.0f}%</span><br>"
                            f"<span style='font-size:12px;color:#666'>{label}</span>",
                       showarrow=False, font=dict(family=FONT, color="#222"))
    return base_layout(fig, 200)


def bar_h(series: pd.Series, h: int = 300, color: str = GREEN) -> go.Figure:
    s = series.sort_values()
    fig = go.Figure(go.Bar(x=s.values, y=s.index.astype(str), orientation="h",
                           marker=dict(color=color, cornerradius=6), text=s.values, textposition="outside"))
    fig.update_xaxes(visible=False)
    fig.update_yaxes(showgrid=False)
    return base_layout(fig, h)


# ----------------------------------------------------------------------------
# Composants HTML
# ----------------------------------------------------------------------------
def minibars(s: pd.Series, total: int | None = None) -> str:
    total = total or max(int(s.sum()), 1)
    out = []
    for k, v in s.items():
        w = v / total * 100
        out.append(f"<div class='mini'><div class='nm'>{k}</div><div class='tr'><div style='width:{w:.0f}%'></div></div>"
                   f"<div class='vl'>{v}</div></div>")
    return "".join(out)


def calendar_html(df: pd.DataFrame) -> str:
    d = df.dropna(subset=["date"])
    if d.empty:
        return "<p class='muted'>Aucune date de collecte.</p>"
    end = d["date"].max().normalize()
    days = list(pd.date_range(end - timedelta(days=5), end))
    d = d[d["date"] >= days[0]]
    d = d.assign(day=d["date"].dt.normalize(), hour=d["date"].dt.hour)
    g = d.groupby(["day", "hour", "agent_id"]).size().rename("n").reset_index()
    hours = g.groupby("hour")["n"].sum().nlargest(5).index.sort_values().tolist()
    prev_m, next_m = MOIS[(end.month - 2) % 12], MOIS[end.month % 12]

    cols = f"64px repeat({len(days)}, 1fr)"
    parts = [f"<div class='cal-head'><span class='cal-chip'>{prev_m}</span>"
             f"<span class='card-title' style='font-size:1.1rem'>{MOIS[end.month - 1]} {end.year}</span>"
             f"<span class='cal-chip'>{next_m}</span></div>",
             f"<div class='cal' style='grid-template-columns:{cols}'>", "<div></div>"]
    for day in days:
        cls = "dh today" if day == end else "dh"
        parts.append(f"<div class='{cls}'>{JOURS[day.weekday()]}<b>{day.day}</b></div>")
    gmax = g["n"].max() if not g.empty else 0
    for h in hours:
        parts.append(f"<div class='hr'>{h:02d}:00</div>")
        for day in days:
            ev = g[(g["day"] == day) & (g["hour"] == h)].sort_values("n", ascending=False).head(2)
            cell = "".join(
                f"<div class='ev {'g' if r.n == gmax else 'w'}'>{r.agent_id}<small>{r.n} fiche{'s' if r.n > 1 else ''}</small></div>"
                for r in ev.itertuples())
            parts.append(f"<div class='cell'>{cell}</div>")
    parts.append("</div>")
    return "".join(parts)


def initials(name: str) -> str:
    w = [x for x in re.split(r"[\s']+", str(name)) if x]
    return (w[0][0] + (w[-1][0] if len(w) > 1 else "")).upper() if w else "?"


# ----------------------------------------------------------------------------
# Application
# ----------------------------------------------------------------------------
html(CSS)

# --- En-tête ---------------------------------------------------------------
h1, h2, h3 = st.columns([1.1, 5.2, 1.4], vertical_alignment="center")
with h1:
    html("<span class='logo'>Ecam·Track</span>")
with h2:
    with st.container(key="nav"):
        page = st.segmented_control(
            "Navigation", ["Tableau de bord", "Producteurs", "Coachs", "Symptômes", "Carte", "Données"],
            default="Tableau de bord", label_visibility="collapsed") or "Tableau de bord"
with h3:
    pop = st.popover("⚙ Paramètres", width="stretch")

df_all, raw, source, meta = load()
st.session_state.setdefault("data_version", meta.get("version"))
if meta.get("version"):
    st.session_state["data_version"] = meta["version"]
watch_updates()

with pop:
    maj = ""
    if meta.get("date"):
        maj = " · publiées le " + pd.Timestamp(meta["date"]).tz_convert("Africa/Abidjan").strftime("%d/%m/%Y à %H:%M")
    st.caption(f"Données : **{source}** · {len(df_all)} lignes{maj}")
    sec_opts = sorted(df_all["section"].dropna().unique())
    coach_opts = sorted(df_all["agent_id"].dropna().unique())
    f_sec = st.multiselect("Sections", sec_opts)
    f_coach = st.multiselect("Coachs", coach_opts)
    if df_all["date"].notna().any():
        dmin, dmax = df_all["date"].min().date(), df_all["date"].max().date()
        f_dates = st.date_input("Période", (dmin, dmax), min_value=dmin, max_value=dmax)
    else:
        f_dates = ()
    objectif = st.number_input("Objectif (nombre de producteurs)", min_value=0, value=1000, step=50,
                               help="0 = pourcentage calculé sur les producteurs enquêtés.")
    st.divider()
    admin_panel()

df = df_all.copy()
if f_sec:
    df = df[df["section"].isin(f_sec)]
if f_coach:
    df = df[df["agent_id"].isin(f_coach)]
if isinstance(f_dates, (tuple, list)) and len(f_dates) == 2:
    df = df[df["date"].isna() | df["date"].dt.date.between(f_dates[0], f_dates[1])]

if df.empty:
    st.warning("Aucune donnée pour ces filtres.")
    st.stop()

# --- Indicateurs -----------------------------------------------------------
prod = (df.sort_values("date").groupby("prod_code", dropna=True)
        .agg(nom=("prod_nom", "last"), genre=("genre", "last"), section=("section", "last"),
             localite=("localite", "last"), campement=("campement", "last"), annee=("annee", "last"),
             superficie=("sup_tot", "max"), parcelles=("plantation", "nunique"),
             cssvd=("cssvd", "max"), engagement=("engagement", "max"), coach=("agent_id", "last"),
             fiches=("prod_nom", "size")))
n_fiches, n_prod, n_coach = len(df), len(prod), df["agent_id"].nunique()
hectares = prod["superficie"].sum()
r_eng = df["engagement"].astype(float).mean()
r_cssvd = df["cssvd"].astype(float).mean()
r_doc = ((df["doc_promesse"] == 1) | (df["doc_fiche"] == 1)).mean()
r_gps = df["lat"].notna().mean()


def prod_rate(mask: pd.Series) -> float:
    """Part des producteurs (uniques) remplissant la condition, sur l'objectif ou sur les enquêtés."""
    base = objectif or n_prod
    return df.loc[mask.fillna(False).astype(bool), "prod_code"].nunique() / base if base else np.nan


k_eng = prod_rate(df["engagement"] == True)  # noqa: E712
k_insp = prod_rate(df["doc_fiche"] == 1)
k_prom = prod_rate(df["doc_promesse"] == 1)

# --- Bloc « Bienvenue » ----------------------------------------------------
w1, w2 = st.columns([1.55, 1], vertical_alignment="bottom")
with w1:
    html(f"""
    <div class='welcome'>Bienvenue,</div>
    <div class='kp-row'>
      <div class='kp'><span class='kp-lab'>Engagement du ménage</span>
        <div class='kp-pill kp-striped'><div class='kp-fill' style='width:{min(max(k_eng * 100, 0), 100):.1f}%'></div>
        <span>{pct(k_eng)}</span></div></div>
      <div class='kp'><span class='kp-lab'>Inspection CSSVD</span>
        <div class='kp-pill kp-green'><span>{pct(k_insp)}</span></div></div>
      <div class='kp'><span class='kp-lab'>Promesse de taille</span>
        <div class='kp-pill kp-dark'><span>{pct(k_prom)}</span></div></div>
    </div>""")
with w2:
    html(f"""
    <div style='display:flex; justify-content:flex-end; gap:34px; padding-bottom:4px'>
      <div><div class='stat'><span class='ic'>📝</span><span class='num'>{n_fiches}</span></div><div class='lab' style='margin-left:40px'>Fiches</div></div>
      <div><div class='stat'><span class='ic'>🧑‍🌾</span><span class='num'>{n_coach}</span></div><div class='lab' style='margin-left:40px'>Coachs</div></div>
      <div><div class='stat'><span class='ic'>🌳</span><span class='num'>{fmt(hectares)}</span></div><div class='lab' style='margin-left:40px'>Hectares</div></div>
    </div>""")

# --- Compteurs « x / objectif » --------------------------------------------
def n_prod_where(mask: pd.Series) -> int:
    return int(df.loc[mask.fillna(False).astype(bool), "prod_code"].nunique())


cible = objectif or n_prod
compteurs = [
    ("Producteurs enquêtés", n_fiches, "#8fcf86"),  # = nombre total de lignes
    ("Refus d'engagement", n_prod_where(df["engagement"] == False), "#d9534f"),  # noqa: E712
    ("Engagement du ménage", n_prod_where(df["engagement"] == True), GREEN),  # noqa: E712
    ("Inspection CSSVD", n_prod_where(df["doc_fiche"] == 1), "#1f6a24"),
    ("Promesse de taille", n_prod_where(df["doc_promesse"] == 1), DARK),
]
cards = "".join(
    f"<div class='cnt' style='--c:{c}'><div class='n'>{v}<small>/{fmt(cible)}</small></div>"
    f"<div class='lb'>{lab}</div><div class='tr'><div style='width:{min(v / cible * 100, 100) if cible else 0:.1f}%'></div></div>"
    f"<div class='pc'>{(v / cible * 100 if cible else 0):.1f} % de l'objectif</div></div>"
    for lab, v, c in compteurs)
html(f"<div class='cnt-row'>{cards}</div>")

st.write("")

# ============================================================================
# PAGE : Tableau de bord
# ============================================================================
if page == "Tableau de bord":
    left, right = st.columns([3, 1.05], gap="medium")

    with left:
        a1, a2, a3 = st.columns(3, gap="medium")
        # Carte profil : coach le plus actif
        top = df.groupby(["agent_id", "agent_nom"]).size().sort_values(ascending=False)
        (cid, cname), cnt = top.index[0], int(top.iloc[0])
        with a1, st.container(key="card_profile"):
            html(f"""
            <div class='muted' style='color:rgba(255,255,255,.8)'>🏆 Coach le plus actif</div>
            <div class='prof-av'>{initials(cname)}</div>
            <div class='prof-foot'>
              <div><div class='prof-name'>{str(cname).title()}</div><div class='prof-role'>Coach · {cid}</div></div>
              <div class='prof-pill'>{cnt} fiches</div>
            </div>""")
        # Activité
        fig, week = activity_chart(df)
        with a2, st.container(key="card_activity"):
            html(f"""<span class='arrow-btn'>↗</span><p class='card-title'>Activité</p>
            <div style='display:flex; align-items:center; gap:12px; margin-top:8px'>
              <span class='big'>{week}</span><span class='muted'>fiches<br>7 derniers jours</span></div>""")
            st.plotly_chart(fig, config=CFG, width="stretch")
        # Taux d'engagement
        n_oui = int((df["engagement"] == True).sum())  # noqa: E712
        n_non = int((df["engagement"] == False).sum())  # noqa: E712
        with a3, st.container(key="card_ring"):
            html("<span class='arrow-btn'>↗</span><p class='card-title'>Engagement</p>")
            st.plotly_chart(ring_chart(r_eng, "Producteurs engagés"), config=CFG, width="stretch")
            html(f"""<div style='display:flex; gap:8px; justify-content:center'>
            <span class='cal-chip'>✓ Oui · {n_oui}</span><span class='cal-chip'>✕ Non · {n_non}</span>
            <span class='cal-chip' style='background:#2c2c2c;color:#fff'>{n_fiches} fiches</span></div>""")

        b1, b2 = st.columns([1, 2], gap="medium")
        with b1, st.container(key="card_accordion"):
            with st.expander("Répartition par section", expanded=True):
                html(minibars(prod["section"].value_counts().head(6), n_prod))
            with st.expander("Genre"):
                html(minibars(prod["genre"].value_counts(), n_prod))
            with st.expander("Type de ménage"):
                html(minibars(df["menage"].value_counts().head(5)))
            with st.expander("Lieu de résidence"):
                html(minibars(prod["campement"].value_counts(), n_prod))
            with st.expander("Plus grande exploitation"):
                if prod["superficie"].notna().any():
                    big = prod.sort_values("superficie", ascending=False).iloc[0]
                    html(f"<div class='mini'><div class='nm'><b>{big.nom}</b><br><span class='muted'>"
                         f"{big.name} · {big.section}</span></div><div class='vl' style='width:auto'>"
                         f"{fmt(big.superficie, 2)} ha</div></div>")
        with b2, st.container(key="card_calendar"):
            html(calendar_html(df))

    with right:
        docs = {
            "Promesse": (df["doc_promesse"] == 1).mean(),
            "Fiche CSSVD": (df["doc_fiche"] == 1).mean(),
            "Aucun": (df["doc_aucun"] == 1).mean(),
        }
        seg_cols = [GREEN, DARK, GREY]
        seg = "".join(
            f"<div style='flex:{max(v, .08):.3f}'><span class='muted' style='color:#333'>{pct(v)}</span>"
            f"<div class='bar' style='background:{c}'>{k if v > .25 else ''}</div></div>"
            for (k, v), c in zip(docs.items(), seg_cols))
        with st.container(key="card_docs"):
            html(f"""<div style='display:flex; justify-content:space-between; align-items:center'>
            <p class='card-title'>Documents</p><span class='big' style='font-size:2rem'>{pct(r_doc)}</span></div>
            <div class='seg'>{seg}</div>""")

        cs = df[df["cssvd"] == True]  # noqa: E712
        n_obs = sum(int((cs[k] == 1).sum()) > 0 for k in SYMPTOMES)
        with st.container(key="dark_symptoms"):
            html(f"""<div style='display:flex; justify-content:space-between; align-items:center; margin-bottom:6px'>
            <span style='font-size:1.25rem'>Symptômes CSSVD</span><span style='font-size:2rem;font-weight:300'>{n_obs}/6</span></div>
            <div class='t2' style='color:#a8a8a8;font-size:.78rem;margin-bottom:6px'>{len(cs)} parcelles avec signes ·
            cliquez sur un symptôme pour voir les photos</div>""")
            symptom_buttons(cs, "dash")
            st.download_button("⬇  Exporter la liste (Excel)", to_xlsx(cssvd_table(df)), key="dash_export",
                               file_name=f"producteurs_swollen_shoot_{pd.Timestamp.now():%Y%m%d}.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                               width="stretch", disabled=cs.empty)

# ============================================================================
# PAGE : Producteurs
# ============================================================================
elif page == "Producteurs":
    c1, c2, c3 = st.columns(3, gap="medium")
    with c1, st.container(key="card_p_sec"):
        html("<p class='card-title'>Producteurs par section</p>")
        st.plotly_chart(bar_h(prod["section"].value_counts()), config=CFG, width="stretch")
    with c2, st.container(key="card_p_sup"):
        html("<p class='card-title'>Superficie totale (ha)</p>")
        fig = px.histogram(prod, x="superficie", nbins=20, color_discrete_sequence=[GREEN])
        fig.update_traces(marker_cornerradius=6)
        fig.update_yaxes(title=None, gridcolor="rgba(0,0,0,.06)")
        fig.update_xaxes(title=None)
        st.plotly_chart(base_layout(fig, 300), config=CFG, width="stretch")
    with c3, st.container(key="card_p_age"):
        html("<p class='card-title'>Âge des producteurs</p>")
        ref_year = int(df["date"].max().year) if df["date"].notna().any() else 2026
        age = (ref_year - prod["annee"]).dropna()
        bins = pd.cut(age, [0, 30, 40, 50, 60, 120], labels=["< 30", "30-39", "40-49", "50-59", "60 +"], right=False)
        st.plotly_chart(bar_h(bins.value_counts().sort_index(ascending=False), color=DARK), config=CFG,
                        width="stretch")
    with st.container(key="card_table"):
        q = st.text_input("Rechercher un producteur", placeholder="Nom, code planteur, localité…")
        t = prod.reset_index().rename(columns={"prod_code": "code"})
        if q:
            m = t.astype(str).apply(lambda col: col.str.contains(q, case=False, regex=False)).any(axis=1)
            t = t[m]
        st.dataframe(t, width="stretch", hide_index=True, height=420, column_config={
            "superficie": st.column_config.NumberColumn("Superficie (ha)", format="%.2f"),
            "cssvd": st.column_config.CheckboxColumn("CSSVD"),
            "engagement": st.column_config.CheckboxColumn("Engagé"),
            "annee": st.column_config.NumberColumn("Naissance", format="%d"),
        })

# ============================================================================
# PAGE : Coachs
# ============================================================================
elif page == "Coachs":
    perf = (df.groupby(["agent_id", "agent_nom"])
            .agg(fiches=("prod_code", "size"), producteurs=("prod_code", "nunique"),
                 engagement=("engagement", lambda s: s.astype(float).mean()),
                 cssvd=("cssvd", lambda s: s.astype(float).mean()),
                 gps=("lat", lambda s: s.notna().mean()),
                 derniere=("date", "max"))
            .reset_index().sort_values("fiches", ascending=False))
    c1, c2 = st.columns([1.2, 1], gap="medium")
    with c1, st.container(key="card_c_bar"):
        html("<p class='card-title'>Fiches collectées par coach</p>")
        st.plotly_chart(bar_h(perf.set_index("agent_id")["fiches"], 340), config=CFG, width="stretch")
    with c2, st.container(key="card_c_time"):
        html("<p class='card-title'>Collectes par jour</p>")
        daily = df.dropna(subset=["date"]).groupby([df["date"].dt.date, "agent_id"]).size().rename("n").reset_index()
        fig = px.bar(daily, x="date", y="n", color="agent_id",
                     color_discrete_sequence=[GREEN, DARK, GREEN_L, GREY, "#3e7a3c", "#c9c9c9", "#7a7a7a"])
        fig.update_layout(showlegend=True, legend=dict(orientation="h", y=-.2, title=None))
        fig.update_xaxes(title=None)
        fig.update_yaxes(title=None, gridcolor="rgba(0,0,0,.06)")
        base_layout(fig, 340).update_layout(showlegend=True)
        st.plotly_chart(fig, config=CFG, width="stretch")
    with st.container(key="card_c_table"):
        st.dataframe(perf, width="stretch", hide_index=True, column_config={
            "agent_id": "ID", "agent_nom": "Nom",
            "engagement": st.column_config.ProgressColumn("Engagement", format="percent", min_value=0, max_value=1),
            "cssvd": st.column_config.ProgressColumn("CSSVD détecté", format="percent", min_value=0, max_value=1),
            "gps": st.column_config.ProgressColumn("GPS renseigné", format="percent", min_value=0, max_value=1),
            "derniere": st.column_config.DatetimeColumn("Dernière fiche", format="DD/MM HH:mm"),
        })

# ============================================================================
# PAGE : Symptômes
# ============================================================================
elif page == "Symptômes":
    cs = df[df["cssvd"] == True]  # noqa: E712
    c1, c2 = st.columns([1, 1.3], gap="medium")
    with c1, st.container(key="card_s_ring"):
        html("<p class='card-title'>Parcelles avec signes de CSSVD</p>")
        st.plotly_chart(ring_chart(r_cssvd, f"{len(cs)} / {df['cssvd'].notna().sum()} parcelles"),
                        config=CFG, width="stretch")
    with c2, st.container(key="card_s_bar"):
        html("<p class='card-title'>Fréquence des symptômes</p>")
        s = pd.Series({SYMPTOMES[k][0]: int((cs[k] == 1).sum()) for k in SYMPTOMES})
        st.plotly_chart(bar_h(s, 260), config=CFG, width="stretch")
    with st.container(key="dark_s_photos"):
        html("<span style='font-size:1.25rem'>📷 Photos par symptôme</span>"
             "<div class='t2' style='color:#a8a8a8;font-size:.78rem;margin-bottom:6px'>Cliquez pour ouvrir la galerie</div>")
        symptom_buttons(cs, "page")

    # --- Export de la liste des producteurs avec swollen shoot ---------------
    with st.container(key="card_s_export"):
        html("<p class='card-title'>🧾 Producteurs avec swollen shoot</p>"
             "<div class='muted'>1 ligne = 1 parcelle enquêtée avec signes de CSSVD · coach, parcelle, symptômes, GPS, photos</div>")
        tab = cssvd_table(df)
        e1, e2, e3 = st.columns(3)
        fs_sec = e1.multiselect("Section", sorted(tab["Section"].dropna().unique()), key="ex_sec")
        fs_coach = e2.multiselect("Coach", sorted(tab["ID coach"].dropna().unique()), key="ex_coach")
        fs_sym = e3.multiselect("Symptôme", [v[0] for v in SYMPTOMES.values()], key="ex_sym",
                                help="Parcelles présentant au moins un des symptômes choisis.")
        if fs_sec:
            tab = tab[tab["Section"].isin(fs_sec)]
        if fs_coach:
            tab = tab[tab["ID coach"].isin(fs_coach)]
        if fs_sym:
            tab = tab[(tab[fs_sym] == "Oui").any(axis=1)]
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Parcelles", len(tab))
        m2.metric("Producteurs", tab["Code planteur"].nunique())
        m3.metric("Coachs", tab["ID coach"].nunique())
        m4.metric("Superficie totale (ha)", fmt(tab.drop_duplicates("Code planteur")["Superficie totale (ha)"].sum(), 1))
        photo_cfg = {c: st.column_config.LinkColumn(c, display_text="📷 Voir") for c in tab if c.startswith("Photo")}
        st.dataframe(tab, width="stretch", hide_index=True, height=380, column_config=photo_cfg)
        stamp = pd.Timestamp.now().strftime("%Y%m%d")
        d1, d2, _ = st.columns([1, 1, 2])
        d1.download_button("⬇ Exporter Excel", to_xlsx(tab), file_name=f"producteurs_swollen_shoot_{stamp}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           type="primary", width="stretch", disabled=tab.empty)
        d2.download_button("⬇ Exporter CSV", tab.to_csv(index=False, sep=";").encode("utf-8-sig"),
                           file_name=f"producteurs_swollen_shoot_{stamp}.csv", mime="text/csv",
                           width="stretch", disabled=tab.empty)

    with st.container(key="card_s_sec"):
        html("<p class='card-title'>Taux de CSSVD par section</p>")
        t = (df.groupby("section").agg(parcelles=("cssvd", "size"),
                                       cssvd=("cssvd", lambda s: s.astype(float).mean()))
             .sort_values("cssvd", ascending=False).reset_index())
        fig = px.bar(t, x="section", y="cssvd", text=t["cssvd"].map(pct), color_discrete_sequence=[GREEN])
        fig.update_traces(marker_cornerradius=8, textposition="outside")
        fig.update_yaxes(visible=False, range=[0, 1.15])
        fig.update_xaxes(title=None)
        st.plotly_chart(base_layout(fig, 300), config=CFG, width="stretch")

# ============================================================================
# PAGE : Carte
# ============================================================================
elif page == "Carte":
    g = df.dropna(subset=["lat", "lon"]).copy()
    with st.container(key="card_map"):
        html(f"<p class='card-title'>Parcelles géolocalisées · {len(g)} / {len(df)}</p>")
        if g.empty:
            st.info("Aucune coordonnée GPS dans les données filtrées.")
        else:
            g["Statut"] = g["cssvd"].map({True: "CSSVD détecté", False: "Sain"}).fillna("Non renseigné")
            g["sup"] = g["sup_tot"].fillna(1).clip(.5, 30)
            fig = px.scatter_map(
                g, lat="lat", lon="lon", color="Statut", size="sup", size_max=18, zoom=10,
                hover_name="prod_nom", hover_data={"prod_code": True, "section": True, "agent_id": True,
                                                   "sup_tot": ":.2f", "lat": False, "lon": False, "sup": False},
                color_discrete_map={"CSSVD détecté": "#d9534f", "Sain": GREEN, "Non renseigné": GREY},
                map_style="open-street-map")
            fig.update_layout(height=600, margin=dict(l=0, r=0, t=0, b=0), paper_bgcolor="rgba(0,0,0,0)",
                              legend=dict(orientation="h", y=1.02, x=0, bgcolor="rgba(255,255,255,.7)"),
                              font=dict(family=FONT))
            st.plotly_chart(fig, width="stretch", config=CFG)

# ============================================================================
# PAGE : Données
# ============================================================================
else:
    with st.container(key="card_raw"):
        html(f"<p class='card-title'>Données brutes · {source}</p>")
        view = st.radio("Vue", ["Normalisée", "Export d'origine"], horizontal=True, label_visibility="collapsed")
        data = df if view == "Normalisée" else raw.loc[raw.index.intersection(df.index)]
        st.dataframe(data, width="stretch", height=560)
        st.download_button("⬇ Télécharger (CSV)", data.to_csv(index=False, sep=";").encode("utf-8-sig"),
                           file_name="cssvd_export.csv", mime="text/csv")
