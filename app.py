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
GREEN_L = "#9fd19c"
DARK = "#1d3622"  # vert profond (remplace le noir)
GREY = "#b9c4b6"
G1 = "#2e7d32"
AMBER = "#e3a33b"
RED = "#d0564f"
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
@import url('https://fonts.googleapis.com/css2?family=Outfit:wght@300;400;500;600;700&family=DM+Serif+Display:ital@0;1&display=swap');
:root { --ink:#18261a; --muted:#6a7b69; --line:rgba(31,58,36,.09); --soft:#f3f8f0; --hero:#1d3622; --hero2:#2f5a36;
        --g1:#2e7d32; --g2:#5aa55a; --g3:#9fd19c; --amber:#e3a33b; --red:#d0564f;
        --bg:#eef6ea; --bg2:#dcebd5; --card:rgba(255,255,255,.78); --card-solid:#fff; --btn:rgba(255,255,255,.85);
        --accent:#2f5a36; --warn-ink:#7a4d00;
        --shadow:0 1px 2px rgba(20,40,20,.04),0 8px 30px rgba(20,50,20,.07); }
html, body, .stApp, [class*="css"], button, input, textarea { font-family: 'Outfit', sans-serif !important; }
.stApp { background:var(--bg);
  background-image: radial-gradient(900px 500px at 8% -10%, var(--bg2), transparent 60%),
                    radial-gradient(700px 500px at 100% 0%, var(--bg2), transparent 55%);
  background-attachment: fixed; color: var(--ink); }
header[data-testid="stHeader"] { background: transparent; }
.block-container { max-width: 1400px; padding: 1.4rem 1.6rem 2.6rem !important; }
h1,h2,h3,h4 { font-family: 'Outfit', sans-serif !important; color: var(--ink); }
@keyframes rise { from { opacity:0; transform:translateY(8px); } to { opacity:1; transform:none; } }

/* Cartes */
[class*="st-key-card"] {
  background: var(--card); border: 1px solid var(--line);
  border-radius: 22px; padding: 20px 22px; backdrop-filter: blur(18px); -webkit-backdrop-filter: blur(18px);
  box-shadow: var(--shadow); animation: rise .5s ease both;
}
[class*="st-key-dark"] {
  background: linear-gradient(145deg, var(--hero2), var(--hero)); border-radius: 22px; padding: 22px 20px; color: #fff;
  box-shadow: 0 14px 40px rgba(29,54,34,.28); animation: rise .5s ease both; position: relative; overflow: hidden;
}
[class*="st-key-dark"] * { color: #fff; }
.st-key-card_profile {
  background: linear-gradient(160deg, #4f8a55 0%, #2f5a36 50%, #1d3622 100%) !important; border: none !important;
  min-height: 300px; position: relative; overflow: hidden;
}
.h-equal { min-height: 300px; }

/* Typo */
.brand { display:flex; align-items:center; gap:12px; }
.brand .lg { width:44px; height:44px; border-radius:14px; background:linear-gradient(135deg, var(--hero2), var(--hero));
             display:grid; place-items:center; font-size:22px; box-shadow: var(--shadow); flex:none; }
.brand b { white-space:nowrap; font-size:1.1rem; font-weight:700; color:var(--ink); letter-spacing:.2px; display:block; line-height:1.15; }
.brand small { color:var(--muted); font-size:.78rem; }
.welcome { font-family:'DM Serif Display', Georgia, serif !important; font-size: 3rem; font-weight: 400; letter-spacing: -.5px;
           color:var(--ink); margin: 6px 0 4px; line-height:1.05; }
.welcome span { font-family:inherit !important; }
.card-title { font-size: 1.15rem; font-weight: 600; color:var(--ink); margin:0; }
.muted { color:var(--muted); font-size:.82rem; }
.big { font-size: 2.3rem; font-weight: 600; letter-spacing:-1px; line-height:1; }
.arrow-btn { float:right; width:38px; height:38px; border-radius:50%; background:var(--soft); display:flex;
             align-items:center; justify-content:center; font-size:1rem; color:var(--ink); border:1px solid var(--line); }

/* Pills de progression */
.prow { display:flex; gap:8px; align-items:flex-end; }
.pcol { display:flex; flex-direction:column; gap:6px; }
.plabel { font-size:.85rem; color:var(--ink); }
.pill { height:40px; border-radius:999px; display:flex; align-items:center; padding:0 16px; font-size:.8rem; }
.pill.dark { background:var(--hero); color:#fff; }
.pill.green { background:#5aa55a; color:#fff; }
.pill.outline { border:1.5px solid var(--muted); color:var(--ink); background:transparent; }
.pill.striped { padding:0; overflow:hidden; background: repeating-linear-gradient(-55deg, rgba(255,255,255,.0) 0 6px, rgba(255,255,255,.75) 6px 9px); border:1px solid rgba(255,255,255,.6);}
.pill.striped > div { height:100%; border-radius:999px; display:flex; align-items:center; padding-left:16px; color:#fff;
  background: repeating-linear-gradient(-55deg, #5aa55a 0 8px, #6cb46a 8px 12px); }

/* Pills KPI (Engagement / Inspection / Promesse) */
.kp-row { display:flex; gap:22px; align-items:flex-end; margin-top:14px; }
.kp { display:flex; flex-direction:column; align-items:flex-start; gap:8px; flex:1; max-width:200px; }
.kp-lab { font-weight:500; font-size:.88rem; color:var(--muted); white-space:nowrap; }
.kp-pill { position:relative; width:100%; height:40px; border-radius:999px; overflow:hidden; display:flex; align-items:center;
           box-shadow: var(--shadow); border:1px solid var(--line); }
.kp-pill span { position:relative; z-index:1; padding-left:16px; font-weight:600; font-size:1.15rem; }
.kp-striped { background: var(--card); }
.kp-striped span { color:var(--ink); }
.kp-fill { position:absolute; left:0; top:0; bottom:0; min-width:8px; border-radius:999px;
           background: linear-gradient(90deg, var(--g3), var(--g2)); }
.kp-green { background: linear-gradient(135deg, var(--g2), var(--g1)); border-color:transparent; }
.kp-green span, .kp-dark span { color:#fff; }
.kp-dark { background: linear-gradient(135deg, var(--hero2), var(--hero)); border-color:transparent; }

/* Compteurs x / objectif : cartes avec anneau (même design que le rapport hors ligne) */
.kpis { display:grid; grid-template-columns: 1.35fr repeat(4, 1fr); gap:16px; margin: 18px 0 14px; }
.kpi { position:relative; overflow:hidden; background:var(--card); border:1px solid var(--line); border-radius:22px;
       padding:18px 20px; box-shadow:var(--shadow); backdrop-filter:blur(18px); animation: rise .5s ease both; min-width:0; }
.kpi .lab { font-size:.84rem; color:var(--muted); font-weight:500; display:flex; align-items:center; gap:8px; line-height:1.15; }
.kpi .ico { width:30px; height:30px; border-radius:10px; display:grid; place-items:center; background:var(--soft); font-size:15px; flex:none; }
.kpi .row { display:flex; align-items:center; justify-content:space-between; gap:8px; margin-top:12px; }
.kpi .val { font-size:2.3rem; font-weight:600; letter-spacing:-1px; line-height:1; color:var(--ink); white-space:nowrap; }
.kpi .val small { font-size:.9rem; color:var(--muted); font-weight:400; letter-spacing:0; margin-left:3px; }
.kpi .sub { font-size:.78rem; color:var(--muted); margin-top:10px; line-height:1.35; }
.kpi .sub b { color:var(--red); }
.ring { width:60px; height:60px; flex:none; }
.ring circle { fill:none; stroke-width:7; }
.ring .tr { stroke:var(--line); }
.ring .pr { stroke:var(--g2); stroke-linecap:round; transform:rotate(-90deg); transform-origin:50% 50%; }
.ring text { font:600 13px 'Outfit', sans-serif; fill:var(--ink); text-anchor:middle; dominant-baseline:central; }
.kpi.red .pr { stroke:var(--red); } .kpi.amber .pr { stroke:var(--amber); }
.kpi.hero { background:linear-gradient(145deg, var(--hero2), var(--hero)); border-color:transparent; }
.kpi.hero .lab, .kpi.hero .sub, .kpi.hero .val small { color:rgba(255,255,255,.75); }
.kpi.hero .sub b { color:#ffd0cc; }
.kpi.hero .val { color:#fff; font-size:3rem; }
.kpi.hero .ico { background:rgba(255,255,255,.14); }
.kpi.hero .ring { width:82px; height:82px; }
.kpi.hero .ring .tr { stroke:rgba(255,255,255,.16); } .kpi.hero .ring .pr { stroke:#c8f0c2; }
.kpi.hero .ring text { fill:#fff; font-size:14px; }
.kpi.hero::after { content:""; position:absolute; right:-60px; bottom:-80px; width:220px; height:220px; border-radius:50%;
                   background:radial-gradient(circle, rgba(255,255,255,.12), transparent 70%); }

/* Boutons symptômes (liste sombre cliquable) */
.stApp [class*="st-key-dark"] button { background:rgba(255,255,255,.1) !important; box-shadow:none !important; border:none !important; border-radius:999px !important;
  min-height:46px; margin:1px 0; transition: background .15s, transform .15s; }
.stApp [class*="st-key-dark"] button:hover { background:rgba(255,255,255,.22) !important; transform: translateX(3px); }
.stApp [class*="st-key-dark"] button:disabled { background:rgba(255,255,255,.05) !important; opacity:.45; }
[class*="st-key-dark"] button > div { justify-content:flex-start !important; width:100%; }
[class*="st-key-dark"] button p { color:#fff !important; font-size:.86rem; text-align:left; }

/* Entonnoir « Parcours producteur » (carte verte) */
.fn-row { margin: 10px 0 14px; }
.fn-lab { font-size:.85rem; color:rgba(255,255,255,.9); display:flex; justify-content:space-between; margin-bottom:5px; }
.fn-conv { font-size:.72rem; background:rgba(255,255,255,.18); border-radius:99px; padding:1px 8px; }
.fn-tr { position:relative; height:30px; border-radius:99px; background:rgba(255,255,255,.12); overflow:hidden; }
.fn-tr > div { height:100%; border-radius:99px; background:linear-gradient(90deg, #c8f0c2, rgba(255,255,255,.6)); }
.fn-tr b { position:absolute; right:12px; top:4px; color:#fff; font-weight:500; font-size:.95rem; }

/* Sélecteur de coopérative + navigation : pilule en verre, bouton actif vert profond */
.st-key-coop [data-testid="stButtonGroup"], .st-key-nav [data-testid="stButtonGroup"] {
  background: var(--card); backdrop-filter: blur(14px); border:1px solid var(--line); border-radius:999px;
  padding:4px; display:inline-flex; box-shadow: var(--shadow); }
.st-key-coop [data-testid="stButtonGroup"], .st-key-nav [data-testid="stButtonGroup"] { flex-wrap:nowrap; }
.st-key-coop button, .st-key-nav button { border:none !important; border-radius:999px !important; background:transparent !important;
  font-size:.88rem; padding: 4px 16px !important; min-height:36px; transition: background .2s, transform .15s; }
.st-key-coop button p, .st-key-nav button p { color: var(--muted) !important; font-weight:600; }
.st-key-coop button:hover p, .st-key-nav button:hover p { color: var(--ink) !important; }
.st-key-coop button[aria-checked="true"], .st-key-coop button[kind="segmented_controlActive"],
.st-key-nav button[aria-checked="true"], .st-key-nav button[kind="segmented_controlActive"] {
  background: var(--hero) !important; box-shadow: 0 4px 14px rgba(29,54,34,.3); }
.st-key-coop button[aria-checked="true"] p, .st-key-coop button[kind="segmented_controlActive"] p,
.st-key-nav button[aria-checked="true"] p, .st-key-nav button[kind="segmented_controlActive"] p { color:#fff !important; }

/* Alerte doublons */
.st-key-card_doublons { background: color-mix(in srgb, var(--amber) 13%, var(--card-solid)) !important; border:1px solid color-mix(in srgb, var(--amber) 40%, transparent) !important; padding: 6px 18px !important; }
.st-key-card_doublons summary p { color:var(--warn-ink) !important; font-weight:600; }

/* Gros chiffres */
.stat { display:flex; align-items:flex-end; gap:10px; }
.stat .ic { width:32px; height:32px; border-radius:10px; background:var(--btn); border:1px solid var(--line); display:flex; align-items:center; justify-content:center; font-size:.85rem; margin-bottom:8px;}
.stat .num { font-size:3.2rem; font-weight:600; letter-spacing:-1.5px; line-height:1; color:var(--ink); }
.stat .lab { font-size:.8rem; color:var(--muted); margin-top:2px; }

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
.cal-chip { background:var(--soft); border:1px solid var(--line); border-radius:999px; padding:4px 14px; font-size:.78rem; color:var(--ink); }
.cal { display:grid; gap:0; position:relative; }
.cal .dh { text-align:center; font-size:.85rem; color:var(--muted); padding-bottom:6px; }
.cal .dh b { display:block; font-weight:400; color:var(--muted); }
.cal .dh.today, .cal .dh.today b { color:var(--ink); font-weight:500; }
.cal .hr { font-size:.78rem; color:var(--ink); padding-top:10px; }
.cal .cell { border-left:1.5px dashed var(--line); min-height:48px; padding:4px 5px; }
.ev { border-radius:14px; padding:6px 10px; font-size:.72rem; line-height:1.15; margin-bottom:4px; box-shadow:0 2px 8px rgba(0,0,0,.06);}
.ev.g { background:linear-gradient(135deg, var(--g2), var(--g1)); color:#fff; } .ev.w { background:var(--card-solid); color:var(--ink); border:1px solid var(--line); }
.ev small { display:block; opacity:.8; font-size:.65rem; }

/* Liste repliable / mini barres */
[data-testid="stExpander"] details { border:none !important; background:transparent !important; }
[data-testid="stExpander"] summary { border-bottom:1px solid var(--line); padding-left:0; }
[data-testid="stExpander"] summary p { font-size:.98rem; color:var(--ink); font-weight:500; }
.mini { display:flex; align-items:center; gap:10px; margin:6px 0; font-size:.82rem; }
.mini .nm { width:46%; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; color:var(--ink);}
.mini .tr { flex:1; height:8px; border-radius:99px; background:var(--soft); overflow:hidden; }
.mini .tr > div { height:100%; border-radius:99px; background:linear-gradient(90deg, var(--g2), var(--g1)); }
.mini .vl { width:38px; text-align:right; color:var(--muted); }

[data-testid="stPopover"] button, .stDownloadButton button, .stButton button, .stLinkButton a {
  border-radius:999px !important; border:1px solid var(--line) !important; background:var(--btn) !important;
  box-shadow: var(--shadow); transition: transform .15s, border-color .15s; font-weight:500; }
[data-testid="stPopover"] button:hover, .stDownloadButton button:hover, .stButton button:hover { border-color: var(--g2) !important; transform: translateY(-1px); }
.stButton button[kind="primary"] { background: var(--hero) !important; color:#fff !important; border-color:transparent !important; }
.stButton button[kind="primary"] p { color:#fff !important; }
[data-testid="stDataFrame"] { border-radius:14px; overflow:hidden; border:1px solid var(--line); }
[data-testid="stTextInput"] input, [data-baseweb="select"] > div { border-radius:12px !important; }
.st-key-offline_dl button { border-radius:999px !important; border:none !important; color:#fff !important; font-weight:600 !important;
  background:linear-gradient(135deg,#3b7a45,#1d3622) !important; box-shadow:0 6px 18px rgba(29,54,34,.28); transition:transform .15s; }
.st-key-offline_dl button:hover { transform:translateY(-1px); }
.st-key-card_table [data-testid="stDataFrame"] { border-radius:16px; overflow:hidden; }
.stats .lab { margin-left: 40px; }



/* Bouton mode clair / sombre */
.tt-theme { width:42px; height:42px; border-radius:50%; border:1px solid var(--line); background:var(--btn);
  box-shadow:var(--shadow); cursor:pointer; font-size:18px; line-height:1; display:grid; place-items:center;
  transition: transform .15s, border-color .15s; padding:0; }
.tt-theme:hover { transform: translateY(-1px) rotate(-12deg); border-color: var(--g2); }
/* ===================== Mode sombre =====================
   - automatique : suit le réglage du téléphone / PC
   - forcé : bouton 🌙 / ☀️ de l'en-tête (attribut data-theme sur <html>) */
@media (prefers-color-scheme: dark) {
  :root:not([data-theme="light"]) {--ink:#e8f2e5; --muted:#9db29a; --line:rgba(220,240,215,.10); --soft:#1f3122; --hero:#2b5a33; --hero2:#3f8049; --g1:#62b866; --g2:#86cc82; --g3:#3f7a43; --amber:#f0b54f; --red:#ee7a72; --bg:#0f1a11; --bg2:#14291a; --card:rgba(28,42,30,.78); --card-solid:#1a291c; --btn:rgba(36,54,38,.9); --accent:#9fd89a; --warn-ink:#f3c977; --shadow:0 1px 2px rgba(0,0,0,.3),0 10px 30px rgba(0,0,0,.35);}
  :root:not([data-theme="light"]) .stApp .muted, :root:not([data-theme="light"]) .stApp [data-testid="stCaptionContainer"] p {color: var(--muted) !important;}
  :root:not([data-theme="light"]) .kpi.hero .ring .pr {stroke:#c8f0c2;}
  :root:not([data-theme="light"]) .stButton button p, :root:not([data-theme="light"]) .stDownloadButton button p, :root:not([data-theme="light"]) [data-testid="stPopover"] button p {color: var(--ink) !important;}
  :root:not([data-theme="light"]) .st-key-offline_dl button p, :root:not([data-theme="light"]) .stButton button[kind="primary"] p {color:#fff !important;}
  :root:not([data-theme="light"]) .kp-striped span {color: var(--ink);}
  :root:not([data-theme="light"]) .ev.w small {color: var(--muted);}
  :root:not([data-theme="light"]) .js-plotly-plot [style*="fill: rgb(29, :root:not([data-theme="light"]) 54, :root:not([data-theme="light"]) 34)"] {fill:#dcefd8 !important;}
  :root:not([data-theme="light"]) [style*="background:#1d3622"], :root:not([data-theme="light"]) [style*="background: #1d3622"], :root:not([data-theme="light"]) [style*="background: rgb(29, :root:not([data-theme="light"]) 54, :root:not([data-theme="light"]) 34)"] {background:#3f8049 !important;}
  :root:not([data-theme="light"]) .js-plotly-plot [style*="fill: rgb(34, :root:not([data-theme="light"]) 34, :root:not([data-theme="light"]) 34)"] {fill:var(--ink) !important;}
  :root:not([data-theme="light"]) .js-plotly-plot [style*="fill: rgb(102, :root:not([data-theme="light"]) 102, :root:not([data-theme="light"]) 102)"], :root:not([data-theme="light"]) .js-plotly-plot [style*="fill: rgb(106, :root:not([data-theme="light"]) 123, :root:not([data-theme="light"]) 105)"] {fill:var(--muted) !important;}
  :root:not([data-theme="light"]) .js-plotly-plot [style*="fill: rgb(0, :root:not([data-theme="light"]) 0, :root:not([data-theme="light"]) 0)"][style*="fill-opacity: 0.07"] {fill:#ffffff !important; fill-opacity:.1 !important;}
  :root:not([data-theme="light"]) .js-plotly-plot [style*="fill: rgb(185, :root:not([data-theme="light"]) 196, :root:not([data-theme="light"]) 182)"] {fill:#5c6e5a !important;}
}
:root[data-theme="dark"] {--ink:#e8f2e5; --muted:#9db29a; --line:rgba(220,240,215,.10); --soft:#1f3122; --hero:#2b5a33; --hero2:#3f8049; --g1:#62b866; --g2:#86cc82; --g3:#3f7a43; --amber:#f0b54f; --red:#ee7a72; --bg:#0f1a11; --bg2:#14291a; --card:rgba(28,42,30,.78); --card-solid:#1a291c; --btn:rgba(36,54,38,.9); --accent:#9fd89a; --warn-ink:#f3c977; --shadow:0 1px 2px rgba(0,0,0,.3),0 10px 30px rgba(0,0,0,.35);}
:root[data-theme="dark"] .stApp .muted, :root[data-theme="dark"] .stApp [data-testid="stCaptionContainer"] p {color: var(--muted) !important;}
:root[data-theme="dark"] .kpi.hero .ring .pr {stroke:#c8f0c2;}
:root[data-theme="dark"] .stButton button p, :root[data-theme="dark"] .stDownloadButton button p, :root[data-theme="dark"] [data-testid="stPopover"] button p {color: var(--ink) !important;}
:root[data-theme="dark"] .st-key-offline_dl button p, :root[data-theme="dark"] .stButton button[kind="primary"] p {color:#fff !important;}
:root[data-theme="dark"] .kp-striped span {color: var(--ink);}
:root[data-theme="dark"] .ev.w small {color: var(--muted);}
:root[data-theme="dark"] .js-plotly-plot [style*="fill: rgb(29, :root[data-theme="dark"] 54, :root[data-theme="dark"] 34)"] {fill:#dcefd8 !important;}
:root[data-theme="dark"] [style*="background:#1d3622"], :root[data-theme="dark"] [style*="background: #1d3622"], :root[data-theme="dark"] [style*="background: rgb(29, :root[data-theme="dark"] 54, :root[data-theme="dark"] 34)"] {background:#3f8049 !important;}
:root[data-theme="dark"] .js-plotly-plot [style*="fill: rgb(34, :root[data-theme="dark"] 34, :root[data-theme="dark"] 34)"] {fill:var(--ink) !important;}
:root[data-theme="dark"] .js-plotly-plot [style*="fill: rgb(102, :root[data-theme="dark"] 102, :root[data-theme="dark"] 102)"], :root[data-theme="dark"] .js-plotly-plot [style*="fill: rgb(106, :root[data-theme="dark"] 123, :root[data-theme="dark"] 105)"] {fill:var(--muted) !important;}
:root[data-theme="dark"] .js-plotly-plot [style*="fill: rgb(0, :root[data-theme="dark"] 0, :root[data-theme="dark"] 0)"][style*="fill-opacity: 0.07"] {fill:#ffffff !important; fill-opacity:.1 !important;}
:root[data-theme="dark"] .js-plotly-plot [style*="fill: rgb(185, :root[data-theme="dark"] 196, :root[data-theme="dark"] 182)"] {fill:#5c6e5a !important;}
/* ===================== Version tablette (< 1100 px) ===================== */
@media (max-width: 1100px) {
  .block-container { padding: 1.4rem 1.4rem 2rem !important; }
  .kp-row { gap: 16px; }
  .stat .num { font-size: 2.8rem; }
  .kpis { grid-template-columns: repeat(3, 1fr); }
  .kpi.hero { grid-column: span 3; }
}
/* ===================== Version téléphone (< 768 px) ===================== */
@media (max-width: 768px) {
  .block-container { margin: 0 auto !important; padding: .9rem .75rem 1.6rem !important; max-width: 100vw !important; overflow-x: hidden; }
  [data-testid="stHorizontalBlock"] { gap: .6rem !important; }
  .brand .lg { width:38px; height:38px; font-size:19px; }
  .welcome { font-size: 2.1rem; margin: 2px 0; }
  /* menus : défilement horizontal au doigt au lieu de déborder */
  .st-key-nav [data-testid="stButtonGroup"], .st-key-coop [data-testid="stButtonGroup"] {
    display: flex !important; flex-wrap: nowrap !important; overflow-x: auto; max-width: 100%;
    -webkit-overflow-scrolling: touch; scrollbar-width: none; }
  .st-key-nav [data-testid="stButtonGroup"]::-webkit-scrollbar { display: none; }
  .st-key-nav button, .st-key-coop button { flex: 0 0 auto; white-space: nowrap; }
  /* pastilles */
  .kp-row { gap: 8px; flex-wrap: nowrap; }
  .kp { min-width: 0; max-width: none; flex: 1 1 0; gap: 5px; align-items: stretch; }
  .kp-lab { font-size: .7rem; white-space: normal; text-align: center; line-height: 1.1; min-height: 2.2em;
            display: flex; align-items: flex-end; }
  .kp-pill { height: 30px; }
  .kp-pill span { font-size: .95rem; padding-left: 10px; }
  /* gros chiffres */
  .stats { justify-content: space-around !important; gap: 8px !important; padding-top: 10px; }
  .stat { justify-content: center; }
  .stat .ic { display: none; }
  .stat .num { font-size: 2.1rem; letter-spacing: -1px; }
  .stats .lab { margin-left: 0; text-align: center; font-size: .72rem; }
  /* compteurs / objectif : 1 grand + grille 2 x 2 */
  .kpis { grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 10px; margin-top: 12px; }
  .kpi.hero { grid-column: 1 / -1; }
  .kpi { padding: 14px; border-radius: 18px; }
  .kpi .val { font-size: 1.8rem; } .kpi.hero .val { font-size: 2.6rem; }
  .ring { width: 46px; height: 46px; } .kpi.hero .ring { width: 70px; height: 70px; }
  .kpi .lab { font-size: .76rem; } .kpi .ico { width: 26px; height: 26px; font-size: 13px; }
  .kpi .sub { font-size: .7rem; }
  /* cartes */
  [class*="st-key-card"], [class*="st-key-dark"] { padding: 14px 14px !important; border-radius: 20px; }
  .st-key-card_profile { min-height: 0; }
  .prof-av { width: 104px; height: 104px; font-size: 2.4rem; margin-top: 10px; }
  .prof-foot { margin-top: 18px; }
  .card-title { font-size: 1.1rem; }
  .big { font-size: 1.9rem; }
  .welcome + .kp-row { margin-top: 6px; }
  /* aucun bloc ne doit élargir la page au-delà de l'écran */
  /* colonnes empilées, chacune sur toute la largeur */
  [data-testid="stHorizontalBlock"] { flex-wrap: wrap !important; }
  [data-testid="stColumn"] { flex: 1 1 100% !important; width: 100% !important; min-width: 0 !important; max-width: 100% !important; }
  [data-testid="stVerticalBlock"], [data-testid="stElementContainer"], [data-testid="stMarkdownContainer"] { min-width: 0 !important; max-width: 100%; }
  /* calendrier : défilement horizontal à l'intérieur de sa carte uniquement */
  .st-key-card_calendar { overflow: hidden; }
  .st-key-card_calendar [data-testid="stMarkdownContainer"] { overflow-x: auto; width: 100%; }
  .cal { min-width: 520px; }
  /* alertes et listes */
  [data-testid="stExpander"] summary p { font-size: .88rem; }
  .mini .nm { width: 52%; }
  [class*="st-key-dark"] button p { font-size: .8rem; }
  .fn-tr { height: 26px; }
}
</style>
"""


THEME_TOGGLE = """
<button class="tt-theme" type="button" title="Passer en mode sombre / clair" aria-label="Changer de thème">🌙</button>
<script>
(() => {
  // Streamlit garde le thème choisi dans le navigateur sous cette clé (« Light », « Dark » ou « System »)
  const KEY = `stActiveTheme-${window.location.pathname}-v2`;
  const root = document.documentElement;
  const saved = () => { try { return JSON.parse(localStorage.getItem(KEY)); } catch (e) { return null; } };
  const sync = () => {
    const t = saved();
    if (t === "Dark") root.dataset.theme = "dark";
    else if (t === "Light") root.dataset.theme = "light";
    else delete root.dataset.theme;
  };
  const isDark = () => root.dataset.theme ? root.dataset.theme === "dark"
                                          : matchMedia("(prefers-color-scheme: dark)").matches;
  sync();
  document.querySelectorAll(".tt-theme").forEach((b) => {
    b.textContent = isDark() ? "☀️" : "🌙";
    b.title = isDark() ? "Passer en mode clair" : "Passer en mode sombre";
    if (b.dataset.bound) return;
    b.dataset.bound = "1";
    b.addEventListener("click", () => {
      try { localStorage.setItem(KEY, JSON.stringify(isDark() ? "Light" : "Dark")); } catch (e) { return; }
      window.location.reload();  // Streamlit applique son thème au chargement de la page
    });
  });
})();
</script>
"""


def html(s: str) -> None:
    """Rend du HTML sans que Markdown n'interprète l'indentation."""
    st.markdown("\n".join(line.strip() for line in s.splitlines()), unsafe_allow_html=True)


def ring_svg(p: float) -> str:
    """Anneau de progression SVG (p en %)."""
    p = max(0.0, min(float(p or 0), 100.0))
    c = 2 * np.pi * 26
    return (f"<svg class='ring' viewBox='0 0 64 64'><circle class='tr' cx='32' cy='32' r='26'/>"
            f"<circle class='pr' cx='32' cy='32' r='26' stroke-dasharray='{c:.1f}' stroke-dashoffset='{c * (1 - p / 100):.1f}'/>"
            f"<text x='32' y='32'>{(f'{p:.1f}'.replace('.', ',') if 0 < p < 10 else f'{p:.0f}')}%</text></svg>")


def kpi_cards(items: list[dict], cible: int) -> str:
    """Compteurs « x / objectif » : 1 grande carte verte + 4 cartes avec anneau.
    L'anneau = nombre ÷ objectif (le même calcul que « x / 1 000 ») ; le taux sur les enquêtés est écrit dessous."""
    out = []
    for i, k in enumerate(items):
        hero = i == 0
        out.append(
            f"<div class='kpi {'hero' if hero else ''} {k.get('cls', '')}'>"
            f"<div class='lab'><span class='ico'>{k['ico']}</span>{k['lab']}</div>"
            f"<div class='row'><div class='val'>{fmt(k['v'])}<small>/ {fmt(cible)}</small></div>{ring_svg(k['v'] / cible * 100 if cible else 0)}</div>"
            f"<div class='sub'>{k['sub']}</div></div>")
    return "<div class='kpis'>" + "".join(out) + "</div>"


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


def fetch_image(url: str, token: str) -> tuple[bytes | None, str]:
    """Photo d'un symptôme : copie archivée d'abord (permanente), sinon téléchargement depuis Kobo."""
    img = stored_photo(url)
    if img:
        return img, "ok"
    return download_kobo(url, token)


@st.cache_data(show_spinner=False, ttl=3600, max_entries=500)
def download_kobo(url: str, token: str) -> tuple[bytes | None, str]:
    """Télécharge une photo Kobo -> (JPEG, "ok") ou (None, cause : acces / introuvable / format / reseau)."""
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
    except Exception:
        return None, "reseau"
    ctype = r.headers.get("content-type", "")
    if r.status_code in (401, 403) or (r.status_code == 200 and "text/html" in ctype):
        return None, "acces"  # Kobo renvoie une page de connexion ou un refus
    if r.status_code == 404:  # Kobo masque les photos privées par un 404 quand la clé manque
        return None, "introuvable" if token else "acces"
    if r.status_code != 200:
        return None, f"erreur {r.status_code}"
    try:
        img = Image.open(io.BytesIO(r.content))
        img = img.convert("RGB")
        img.thumbnail((1200, 1200))
        buf = io.BytesIO()
        img.save(buf, format="JPEG", quality=85)
        return buf.getvalue(), "ok"
    except Exception:
        return None, "format"


FAIL_MSG = {
    "acces": "accès refusé par Kobo → la clé `KOBO_TOKEN` est absente, incorrecte, ou le compte n'a pas accès au projet",
    "introuvable": "introuvable sur Kobo → photo supprimée, ou la clé `KOBO_TOKEN` appartient à un compte sans accès au projet",
    "format": "format d'image non lisible",
    "reseau": "Kobo ne répond pas (réseau / délai dépassé)",
}


@st.dialog("Photos du symptôme", width="large")
def show_gallery(k: str, data: pd.DataFrame) -> None:
    name, icon = SYMPTOMES[k]
    col = f"photo_{k}"
    rows = data[data[k] == 1]
    photos = rows[rows[col].notna() & (rows[col] != "")] if col in rows else rows.iloc[0:0]
    photos = photos.sort_values("date", ascending=False)
    html(f"<div style='font-size:1.5rem'>{icon} {name}</div>"
         f"<div class='muted'>{len(rows)} producteur(s) concerné(s) · {len(photos)} photo(s)</div>")
    if photos.empty:
        st.info("Aucune photo disponible pour ce symptôme dans les données filtrées.")
        return
    limit = 24
    token = kobo_token()
    fails: dict[str, int] = {}
    grid = st.columns(3)
    for i, r in enumerate(photos.head(limit).itertuples()):
        url = getattr(r, col)
        with grid[i % 3]:
            with st.spinner(""):
                img, why = fetch_image(url, token)
            if img:
                st.image(img, width="stretch")
            else:
                fails[why] = fails.get(why, 0) + 1
                st.link_button("📷 Ouvrir la photo ↗", url, width="stretch")
            d = r.date.strftime("%d/%m %H:%M") if pd.notna(r.date) else ""
            st.caption(f"**{r.prod_nom}** · {r.prod_code}  \n{r.section} · {r.agent_id} · {d}")
    if len(photos) > limit:
        st.caption(f"{limit} photos les plus récentes affichées sur {len(photos)}.")
    if fails:
        lines = "\n".join(f"- **{n}** : {FAIL_MSG.get(w, w)}" for w, n in fails.items())
        conseil = ("\n\n➡ Ajoutez `KOBO_TOKEN` dans les Secrets (clé : https://kf.kobotoolbox.org/token/?format=json)."
                   if not token else "")
        st.warning(f"{sum(fails.values())} photo(s) non affichée(s) — utilisez « Ouvrir la photo » :\n{lines}{conseil}")


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


def _style_sheet(ws, t: pd.DataFrame, freeze: str = "F2") -> None:
    from openpyxl.styles import Alignment, Font, PatternFill

    head = PatternFill("solid", fgColor="2E7D32")
    for c in ws[1]:
        c.font, c.fill = Font(bold=True, color="FFFFFF"), head
        c.alignment = Alignment(wrap_text=True, vertical="center")
    ws.row_dimensions[1].height = 32
    ws.freeze_panes = freeze
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


def _sheet_name(name: str, used: set) -> str:
    """Nom de feuille Excel valide (31 caractères max, sans crochets : * ? / ni antislash) et unique."""
    base = re.sub(r"[\[\]:*?/\\]", "-", str(name)).strip()[:31] or "Feuille"
    out, i = base, 2
    while out.lower() in used:
        out = f"{base[:28]}-{i}"
        i += 1
    used.add(out.lower())
    return out


def to_xlsx_sheets(sheets: dict[str, pd.DataFrame]) -> bytes:
    """Classeur Excel mis en forme, une feuille par entrée du dictionnaire."""
    buf, used = io.BytesIO(), set()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        for name, t in sheets.items():
            sh = _sheet_name(name, used)
            t.to_excel(xw, index=False, sheet_name=sh)
            _style_sheet(xw.sheets[sh], t, "B2" if name == "Résumé" else "F2")
    return buf.getvalue()


def to_xlsx(t: pd.DataFrame, sheet: str = "Producteurs CSSVD") -> bytes:
    return to_xlsx_sheets({sheet: t})


def upas_xlsx_by_section(t: pd.DataFrame) -> bytes:
    """Liste de suivi UPAS : feuille « Résumé » + une feuille par section."""
    t = t.sort_values(["Section", "Étape", "Code producteur"], na_position="last")
    sec = t["Section"].fillna("Sans section")
    res = (t.assign(Section=sec)
           .groupby("Section")
           .agg(**{"Producteurs": ("Code producteur", "nunique"),
                   "Engagés": ("Engagé le", lambda c: int(c.notna().sum())),
                   "Inspectés": ("Inspecté le", lambda c: int(c.notna().sum())),
                   "Promesses": ("Promesse le", lambda c: int(c.notna().sum())),
                   "Parcours complet": ("Étape", lambda c: int(c.astype(str).str.startswith("✅").sum())),
                   "Superficie (ha)": ("Superficie (ha)", "sum")})
           .reset_index())
    tot = res.drop(columns="Section").sum(numeric_only=True)
    res = pd.concat([res, pd.DataFrame([{"Section": "TOTAL", **tot.to_dict()}])], ignore_index=True)
    res["% complet"] = (res["Parcours complet"] / res["Producteurs"].where(res["Producteurs"] > 0)).map(
        lambda x: f"{x * 100:.0f} %" if pd.notna(x) else "")
    res["Superficie (ha)"] = res["Superficie (ha)"].round(2)
    sheets = {"Résumé": res}
    for name, g in t.assign(Section=sec).groupby("Section", sort=True):
        sheets[name] = g.drop(columns="Section")
    return to_xlsx_sheets(sheets)


def ecam_xlsx_by_section(t: pd.DataFrame) -> bytes:
    """Liste ECAM des parcelles avec swollen shoot : feuille « Résumé » + une feuille par section."""
    t = t.sort_values(["Section", "ID coach", "Nom producteur"], na_position="last")
    t = t.assign(Section=t["Section"].fillna("Sans section"))
    oui = lambda c: int((c == "Oui").sum())  # noqa: E731
    sym_cols = [v[0] for v in SYMPTOMES.values() if v[0] in t]
    flag_cols = [c for c in ["Engagement", "Fiche inspection CSSVD", "Promesse de taille"] if c in t]

    def summarize(g: pd.DataFrame) -> dict:
        row = {"Parcelles": len(g), "Producteurs": g["Code planteur"].nunique(), "Coachs": g["ID coach"].nunique(),
               "Superficie totale (ha)": round(g.drop_duplicates("Code planteur")["Superficie totale (ha)"].sum(), 2)}
        row.update({c: oui(g[c]) for c in sym_cols})
        row.update({f"{c} (Oui)": oui(g[c]) for c in flag_cols})
        return row

    rows = [{"Section": name, **summarize(g)} for name, g in t.groupby("Section", sort=True)]
    rows.append({"Section": "TOTAL", **summarize(t)})  # producteurs / coachs comptés une seule fois au total
    sheets = {"Résumé": pd.DataFrame(rows)}
    for name, g in t.groupby("Section", sort=True):
        sheets[name] = g.drop(columns="Section")
    return to_xlsx_sheets(sheets)


def ecam_producteurs_table(pp: pd.DataFrame, fiches: pd.DataFrame) -> pd.DataFrame:
    """Liste ECAM : 1 ligne par code producteur (fiches en double fusionnées)."""
    f = fiches[fiches["prod_code"].notna()]
    par = f.groupby("prod_code").agg(
        parcelles=("plantation", lambda s: ", ".join(sorted(s.dropna().unique()))),
        nb_parcelles=("plantation", "nunique"), nb_fiches=("prod_code", "size"))
    p = pp[pp["prod_code"].notna()].set_index("prod_code").join(par)
    yn = {True: "Oui", False: "Non"}
    t = pd.DataFrame({
        "Code planteur": p.index, "Nom producteur": p["prod_nom"].values, "Genre": p["genre"].values,
        "Année naissance": p["annee"].astype("Int64").values,
        "Section": p["section"].values, "Localité": p["localite"].values, "Lieu de résidence": p["campement"].values,
        "ID coach": p["agent_id"].values, "Nom coach": p["agent_nom"].values,
        "Code(s) parcelle": p["parcelles"].values, "Nb parcelles": p["nb_parcelles"].values,
        "Nb fiches": p["nb_fiches"].values, "Superficie totale (ha)": p["sup_tot"].round(2).values,
        "Engagement": p["engagement"].map(yn).values,
        "Fiche inspection CSSVD": np.where(p["doc_fiche"] == 1, "Oui", "Non"),
        "Promesse de taille": np.where(p["doc_promesse"] == 1, "Oui", "Non"),
        "Signes CSSVD": p["cssvd"].map(yn).values,
        "Latitude": p["lat"].values, "Longitude": p["lon"].values,
        "Dernière fiche": p["date"].dt.strftime("%d/%m/%Y %H:%M").values,
    })
    return t.sort_values(["Section", "Nom producteur"], na_position="last").reset_index(drop=True)


def ecam_prod_xlsx_by_section(t: pd.DataFrame) -> bytes:
    """Liste des producteurs ECAM : feuille « Résumé » + une feuille par section."""
    t = t.assign(Section=t["Section"].fillna("Sans section"))

    def summarize(g: pd.DataFrame) -> dict:
        oui = lambda c: int((g[c] == "Oui").sum())  # noqa: E731
        return {"Producteurs": g["Code planteur"].nunique(),
                "Femmes": int(g["Genre"].astype(str).str.lower().str.startswith("f").sum()),
                "Hommes": int(g["Genre"].astype(str).str.lower().str.startswith("h").sum()),
                "Engagés": oui("Engagement"), "Refus": int((g["Engagement"] == "Non").sum()),
                "Fiche inspection CSSVD": oui("Fiche inspection CSSVD"),
                "Promesse de taille": oui("Promesse de taille"), "Signes CSSVD": oui("Signes CSSVD"),
                "Superficie totale (ha)": round(g["Superficie totale (ha)"].sum(), 2),
                "Parcelles": int(g["Nb parcelles"].sum()), "Fiches saisies": int(g["Nb fiches"].sum())}

    rows = [{"Section": name, **summarize(g)} for name, g in t.groupby("Section", sort=True)]
    rows.append({"Section": "TOTAL", **summarize(t)})
    res = pd.DataFrame(rows)
    res["% engagés"] = (res["Engagés"] / res["Producteurs"].where(res["Producteurs"] > 0)).map(
        lambda x: f"{x * 100:.0f} %" if pd.notna(x) else "")
    sheets = {"Résumé": res}
    for name, g in t.groupby("Section", sort=True):
        sheets[name] = g.drop(columns="Section")
    return to_xlsx_sheets(sheets)


def symptom_buttons(data: pd.DataFrame, prefix: str) -> None:
    """Un bouton par symptôme ; un clic ouvre la galerie de photos."""
    observed = {k: int((data[k] == 1).sum()) for k in SYMPTOMES}
    for k, v in sorted(observed.items(), key=lambda kv: -kv[1]):
        col = f"photo_{k}"
        n_ph = int((data[col].notna() & (data[k] == 1)).sum()) if col in data else 0
        name, icon = SYMPTOMES[k]
        label = f"{icon}  {name}  ·  {v} producteur{'s' if v > 1 else ''}  ·  📷 {n_ph}"
        if st.button(label, key=f"{prefix}_{k}", width="stretch", disabled=v == 0):
            show_gallery(k, data)


@st.cache_data(show_spinner=False)
def read_bytes(data: bytes, name: str) -> pd.DataFrame:
    if name.lower().endswith((".xlsx", ".xls")):
        return pd.read_excel(io.BytesIO(data))
    return pd.read_csv(io.BytesIO(data), sep=None, engine="python", encoding_errors="replace")


@st.cache_data(show_spinner=False, max_entries=4)
def merge_forms(raw: pd.DataFrame) -> tuple[pd.DataFrame, dict]:
    """Fusionne les fiches Kobo d'un même producteur ET d'une même parcelle en UNE fiche, au format Kobo.

    - options cochées (documents, symptômes) : réunies (case cochée si elle l'est sur au moins une fiche) ;
      « Aucun autre document » est décoché dès qu'un document est présent ;
    - réponses Oui / Non : « Oui » si au moins une fiche dit Oui ;
    - début de saisie : le plus ancien ; autres colonnes : la valeur la plus récente renseignée
      (photos et leurs liens compris) ;
    - colonne « _fusion » : nombre de fiches fusionnées et leurs identifiants Kobo (_id).
    Le fichier d'origine n'est pas modifié.
    """
    info = {"groupes": 0, "fiches": 0, "en_moins": 0}
    cols = list(raw.columns)
    c_code, c_plant = _find(cols, *SPEC["prod_code"]), _find(cols, *SPEC["plantation"])
    if c_code is None or raw.empty:
        return raw, info
    code = raw[c_code].astype("string").str.strip()
    plant = raw[c_plant].astype("string").str.strip().fillna("") if c_plant is not None else ""
    key = code + "§" + plant
    dup = code.notna() & key.duplicated(keep=False)
    if not dup.any():
        return raw, info
    c_sub = next((c for c in cols if _norm(c) == "_submission_time"), None)
    c_id = next((c for c in cols if _norm(c) == "_id"), None)
    children = {c: p for c in cols for p in cols if p != c and c.startswith(p + "/")}  # options à choix multiples
    parents = set(children.values())

    def last(v: pd.Series):
        v = v.dropna()
        v = v[v.astype(str).str.strip() != ""]
        return v.iloc[-1] if len(v) else np.nan

    merged = []
    for _, g in raw[dup].groupby(key[dup], sort=False):
        if c_sub is not None:
            g = g.iloc[pd.to_datetime(g[c_sub], errors="coerce").argsort(kind="stable")]
        row = g.iloc[-1].copy()
        for c in cols:
            v = g[c]
            nc = _norm(c)
            if c in children:
                row[c] = pd.to_numeric(v, errors="coerce").max()
            elif nc == "start":
                row[c] = v.dropna().min() if v.notna().any() else np.nan
            elif nc in ("_id", "_uuid", "_index", "meta/rootuuid", "meta/instanceid", "_submission_time") or c in parents:
                continue  # identifiants : ceux de la fiche la plus récente ; parents recalculés plus bas
            else:
                txt = v.dropna().astype(str).str.strip().str.lower()
                if len(txt) and txt.isin(["oui", "non"]).all():
                    row[c] = "Oui" if (txt == "oui").any() else "Non"
                else:
                    row[c] = last(v)
        for p in parents:  # libellé de la question à choix multiples = options cochées
            opts = [c for c, pp_ in children.items() if pp_ == p]
            on = [c for c in opts if row[c] == 1]
            aucun = [c for c in on if "aucun" in _norm(c)]
            if aucun and len(on) > len(aucun):
                for c in aucun:
                    row[c] = 0
                on = [c for c in on if c not in aucun]
            if on:
                row[p] = " ".join(c[len(p) + 1:] for c in on)
        ids = ", ".join(str(x) for x in g[c_id]) if c_id is not None else ""
        row["_fusion"] = f"{len(g)} fiches fusionnées" + (f" (_id {ids})" if ids else "")
        merged.append(row)
        info["groupes"] += 1
        info["fiches"] += len(g)
    out = pd.concat([raw[~dup], pd.DataFrame(merged)]).sort_index()
    info["en_moins"] = info["fiches"] - info["groupes"]
    return out, info


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
# Chaque fichier a son « emplacement » (slot) : ECAM = 1 export Kobo, UPAS = 3 fichiers.
# ----------------------------------------------------------------------------
HERE = os.path.dirname(os.path.abspath(__file__))
SHARED_DIR = os.path.join(HERE, "data", "shared")

SLOTS = {
    "ecam": {"coop": "ECAM", "label": "Export Kobo ECAM", "path": "export_kobo.xlsx",
             "file": "export_kobo.bin", "meta": "meta.json"},
    "upas_eng": {"coop": "UPAS", "label": "UPAS · Engagements des ménages", "path": "upas/engagements_menages.xlsx"},
    "upas_insp": {"coop": "UPAS", "label": "UPAS · Inspections CSSVD", "path": "upas/inspections_cssvd.xlsx"},
    "upas_prom": {"coop": "UPAS", "label": "UPAS · Promesses de taille", "path": "upas/promesses_taille.xlsx"},
}
COOP_SLOTS = {"ECAM": ["ecam"], "UPAS": ["upas_eng", "upas_insp", "upas_prom"]}


def classify(raw: pd.DataFrame) -> str | None:
    """Reconnaît le type de fichier d'après ses colonnes."""
    cols = {_norm(c) for c in raw.columns}
    if "date engagement" in cols:
        return "upas_eng"
    if "date observation" in cols:
        return "upas_insp"
    if "date promesse" in cols:
        return "upas_prom"
    if any(c.startswith("producteur/") or c == "_submission_time" for c in cols):
        return "ecam"
    return None


def secret(key: str, default: str = "") -> str:
    try:
        v = st.secrets.get(key, None)
    except Exception:  # pas de secrets.toml
        v = None
    return str(v if v is not None else os.environ.get(key, default)).strip()


def gh_cfg() -> tuple[str, str, str] | None:
    tok, repo = secret("GITHUB_TOKEN"), secret("DATA_REPO")
    if not (tok and repo):
        return None
    return tok, repo, secret("DATA_BRANCH", "main")


def slot_path(slot: str) -> str:
    if slot == "ecam":
        return secret("DATA_PATH", SLOTS["ecam"]["path"])
    return SLOTS[slot]["path"]


def _gh(method: str, path: str, **kw):
    import requests

    tok, repo, _ = gh_cfg()
    headers = {"Authorization": f"Bearer {tok}", "X-GitHub-Api-Version": "2022-11-28",
               "Accept": kw.pop("accept", "application/vnd.github+json")}
    return requests.request(method, f"https://api.github.com/repos/{repo}/{path}", headers=headers,
                            timeout=60, **kw)


@st.cache_data(ttl=20, show_spinner=False)
def remote_version(slot: str) -> dict | None:
    """Dernière version publiée sur GitHub pour cet emplacement (vérifiée toutes les 20 s au plus)."""
    cfg = gh_cfg()
    if not cfg:
        return None
    r = _gh("GET", f"commits?path={slot_path(slot)}&sha={cfg[2]}&per_page=1")
    if r.status_code != 200 or not r.json():
        return None
    c = r.json()[0]
    return {"sha": c["sha"], "date": c["commit"]["committer"]["date"], "message": c["commit"]["message"]}


def gh_download(slot: str) -> bytes | None:
    r = _gh("GET", f"contents/{slot_path(slot)}?ref={gh_cfg()[2]}", accept="application/vnd.github.raw")
    return r.content if r.status_code == 200 else None


def gh_upload(slot: str, data: bytes, message: str) -> tuple[bool, str]:
    import base64

    path, branch = slot_path(slot), gh_cfg()[2]
    cur = _gh("GET", f"contents/{path}?ref={branch}")
    body = {"message": message, "content": base64.b64encode(data).decode(), "branch": branch}
    if cur.status_code == 200:
        body["sha"] = cur.json()["sha"]
    r = _gh("PUT", f"contents/{path}", json=body)
    if r.status_code in (200, 201):
        return True, "ok"
    return False, f"GitHub {r.status_code} : {r.text[:200]}"


def _local(slot: str) -> tuple[str, str]:
    s = SLOTS[slot]
    return (os.path.join(SHARED_DIR, s.get("file", f"{slot}.bin")),
            os.path.join(SHARED_DIR, s.get("meta", f"{slot}.json")))


def _read_meta(slot: str) -> dict:
    import json

    try:
        with open(_local(slot)[1], encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def _write_local(slot: str, data: bytes, meta: dict) -> None:
    import json

    os.makedirs(SHARED_DIR, exist_ok=True)
    fpath, mpath = _local(slot)
    tmp = fpath + ".tmp"
    with open(tmp, "wb") as f:
        f.write(data)
    os.replace(tmp, fpath)
    with open(mpath, "w", encoding="utf-8") as f:
        json.dump(meta, f, ensure_ascii=False)


def publish(slot: str, data: bytes, name: str, author: str = "administrateur") -> tuple[bool, str]:
    """Publie un fichier pour tous les utilisateurs."""
    now = pd.Timestamp.now(tz="UTC").isoformat()
    msg = f"Mise à jour des données : {name} ({author})"
    if gh_cfg():
        ok, err = gh_upload(slot, data, msg)
        if not ok:
            return False, err
        remote_version.clear()
        v = remote_version(slot) or {}
        _write_local(slot, data, {"version": v.get("sha", now), "name": name, "date": now})
        return True, f"{SLOTS[slot]['label']} : publié (sauvegarde GitHub)."
    _write_local(slot, data, {"version": now, "name": name, "date": now})
    return True, f"{SLOTS[slot]['label']} : publié (stockage temporaire du serveur)."


def current_version(slots: tuple[str, ...]) -> str:
    """Identifiant combiné des versions publiées (change dès qu'un des fichiers est republié)."""
    out = []
    for s in slots:
        v = remote_version(s)
        out.append(v["sha"] if v else _read_meta(s).get("version", ""))
    return "|".join(out)


def sync_slot(slot: str) -> dict:
    """Aligne la copie locale sur la dernière version publiée ; renvoie ses métadonnées ({} si rien)."""
    v = remote_version(slot)
    meta = _read_meta(slot)
    fpath = _local(slot)[0]
    if v and (meta.get("version") != v["sha"] or not os.path.exists(fpath)):
        data = gh_download(slot)
        if data:
            name = v["message"].split(":", 1)[-1].split("(")[0].strip() or os.path.basename(slot_path(slot))
            meta = {"version": v["sha"], "name": name, "date": v["date"]}
            _write_local(slot, data, meta)
    return meta if os.path.exists(fpath) else {}


def _sniff_name(data: bytes) -> str:
    return "x.xlsx" if data[:2] == b"PK" or data[:4] == b"\xd0\xcf\x11\xe0" else "x.csv"  # xlsx / xls / csv


@st.cache_data(show_spinner="Chargement des données…", max_entries=8)
def load_slot_raw(slot: str, version: str) -> pd.DataFrame:
    with open(_local(slot)[0], "rb") as f:
        data = f.read()
    return read_bytes(data, _sniff_name(data))


@st.cache_data(show_spinner=False, max_entries=32)
def read_path(path: str, mtime: float) -> pd.DataFrame:
    with open(path, "rb") as f:
        return read_bytes(f.read(), path)


def local_files(slot: str) -> list[str]:
    """Fichiers du dossier data/ reconnus comme appartenant à cet emplacement (le plus récent d'abord)."""
    files = glob.glob(os.path.join(HERE, "data", "*.xlsx")) + glob.glob(os.path.join(HERE, "data", "*.csv"))
    hits = []
    for p in files:
        try:
            if classify(read_path(p, os.path.getmtime(p))) == slot:
                hits.append(p)
        except Exception:
            continue
    return sorted(hits, key=os.path.getmtime, reverse=True)


def load_slot(slot: str) -> tuple[pd.DataFrame | None, str, dict]:
    """Données d'un emplacement : version publiée, sinon fichier local du dossier data/, sinon None."""
    meta = sync_slot(slot)
    if meta:
        return load_slot_raw(slot, meta["version"]), meta.get("name", SLOTS[slot]["label"]), meta
    files = local_files(slot)
    if files:
        return read_path(files[0], os.path.getmtime(files[0])), os.path.basename(files[0]), {}
    return None, "", {}


def load() -> tuple[pd.DataFrame, pd.DataFrame, str, dict]:
    raw, name, meta = load_slot("ecam")
    if raw is not None:
        return normalize(raw), raw, name, meta
    d = demo_data()
    return d, d, "Données de démonstration", {}


# ----------------------------------------------------------------------------
# Archive des photos : chaque photo Kobo est copiée dans le dépôt de données (photos/<clé>.jpg).
# Elle reste visible même si la soumission est supprimée sur Kobo ou si la clé Kobo change.
# ----------------------------------------------------------------------------
PHOTO_DIR = os.path.join(SHARED_DIR, "photos")


def photo_key(url: str) -> str:
    import hashlib

    return hashlib.sha1(str(url).split("?")[0].rstrip("/").encode()).hexdigest()[:24]


def photo_urls(raw: pd.DataFrame) -> list[str]:
    """Tous les liens de photos de l'export (symptômes, documents, photo du producteur)."""
    cols = [c for c in raw.columns if _norm(c).endswith("_url")]
    if not cols:
        return []
    v = raw[cols].stack().astype(str).str.strip()
    return list(dict.fromkeys(v[v.str.startswith("http")]))


@st.cache_data(ttl=300, show_spinner=False)
def archived_keys() -> frozenset[str]:
    """Clés des photos déjà archivées (disque du serveur + dépôt GitHub)."""
    keys = set()
    if os.path.isdir(PHOTO_DIR):
        keys |= {f[:-4] for f in os.listdir(PHOTO_DIR) if f.endswith(".jpg")}
    if gh_cfg():
        try:
            r = _gh("GET", f"git/trees/{gh_cfg()[2]}?recursive=1")
            if r.status_code == 200:
                keys |= {t["path"][7:-4] for t in r.json().get("tree", [])
                         if t["path"].startswith("photos/") and t["path"].endswith(".jpg")}
        except Exception:
            pass
    return frozenset(keys)


def _save_photo_local(key: str, data: bytes) -> None:
    os.makedirs(PHOTO_DIR, exist_ok=True)
    with open(os.path.join(PHOTO_DIR, key + ".jpg"), "wb") as f:
        f.write(data)


def stored_photo(url: str) -> bytes | None:
    """Copie archivée d'une photo (disque du serveur, sinon dépôt GitHub), ou None."""
    key = photo_key(url)
    path = os.path.join(PHOTO_DIR, key + ".jpg")
    if os.path.exists(path):
        with open(path, "rb") as f:
            return f.read()
    if gh_cfg() and key in archived_keys():
        try:
            r = _gh("GET", f"contents/photos/{key}.jpg?ref={gh_cfg()[2]}", accept="application/vnd.github.raw")
        except Exception:
            return None
        if r.status_code == 200:
            _save_photo_local(key, r.content)
            return r.content
    return None


def gh_commit_files(files: dict[str, bytes], message: str) -> tuple[bool, str]:
    """Envoie plusieurs fichiers en UN seul commit (API Git de GitHub)."""
    import base64

    branch = gh_cfg()[2]
    ref = _gh("GET", f"git/ref/heads/{branch}")
    if ref.status_code != 200:
        return False, f"GitHub {ref.status_code} : branche « {branch} » introuvable (le dépôt est-il vide ?)"
    head = ref.json()["object"]["sha"]
    base_tree = _gh("GET", f"git/commits/{head}").json()["tree"]["sha"]
    items = []
    for path, data in files.items():
        b = _gh("POST", "git/blobs", json={"content": base64.b64encode(data).decode(), "encoding": "base64"})
        if b.status_code != 201:
            return False, f"GitHub {b.status_code} : {b.text[:200]}"
        items.append({"path": path, "mode": "100644", "type": "blob", "sha": b.json()["sha"]})
    tree = _gh("POST", "git/trees", json={"base_tree": base_tree, "tree": items})
    if tree.status_code != 201:
        return False, f"GitHub {tree.status_code} : {tree.text[:200]}"
    commit = _gh("POST", "git/commits", json={"message": message, "tree": tree.json()["sha"], "parents": [head]})
    if commit.status_code != 201:
        return False, f"GitHub {commit.status_code} : {commit.text[:200]}"
    upd = _gh("PATCH", f"git/refs/heads/{branch}", json={"sha": commit.json()["sha"]})
    return (True, "ok") if upd.status_code == 200 else (False, f"GitHub {upd.status_code} : {upd.text[:200]}")


def archive_photos(raw: pd.DataFrame, token: str, progress=None) -> dict:
    """Télécharge depuis Kobo les photos pas encore archivées et les enregistre (par lots de 20)."""
    urls = photo_urls(raw)
    have = archived_keys()
    todo = [u for u in urls if photo_key(u) not in have]
    res = {"total": len(urls), "deja": len(urls) - len(todo), "ok": 0, "echecs": {}, "erreur": ""}
    batch: dict[str, bytes] = {}

    def flush() -> None:
        if batch and gh_cfg() and not res["erreur"]:
            ok, err = gh_commit_files(batch, f"Archive de {len(batch)} photo(s)")
            if not ok:
                res["erreur"] = err
        batch.clear()

    for i, u in enumerate(todo, start=1):
        img, why = download_kobo(u, token)
        if img:
            key = photo_key(u)
            _save_photo_local(key, img)
            batch[f"photos/{key}.jpg"] = img
            res["ok"] += 1
        else:
            res["echecs"][why] = res["echecs"].get(why, 0) + 1
        if len(batch) >= 20:
            flush()
        if progress:
            progress(i / len(todo), f"Archivage des photos… {i}/{len(todo)}")
    flush()
    archived_keys.clear()
    return res


def archive_report(res: dict) -> str:
    txt = (f"🖼 Photos : **{res['ok']}** archivée(s) maintenant, {res['deja']} déjà archivée(s), "
           f"{res['total']} au total.")
    if res["echecs"]:
        txt += "\n\n⚠ Non archivées : " + " ; ".join(f"{n} — {FAIL_MSG.get(w, w)}" for w, n in res["echecs"].items())
    if res["erreur"]:
        txt += f"\n\n✗ Envoi vers GitHub impossible : {res['erreur']} (photos gardées sur le serveur seulement)."
    return txt


# ----------------------------------------------------------------------------
# Fusion ECAM : le nouvel export est ajouté aux données déjà publiées (rien n'est perdu).
# Une soumission est identifiée par `_uuid` (sinon `_id`) ; la version la plus récente l'emporte.
# ----------------------------------------------------------------------------
def _key_col(df: pd.DataFrame) -> str | None:
    for k in ("_uuid", "_id"):
        c = next((c for c in df.columns if _norm(c) == k), None)
        if c is not None:
            return c
    return None


def merge_ecam(old: pd.DataFrame, new: pd.DataFrame) -> tuple[pd.DataFrame | None, dict]:
    drop = [c for c in set(old.columns) | set(new.columns) if _norm(c) == "_index"]
    old = old.drop(columns=[c for c in drop if c in old])
    new = new.drop(columns=[c for c in drop if c in new])
    common = len(set(old.columns) & set(new.columns)) / max(len(set(new.columns)), 1)
    st_ = {"anciennes": len(old), "fichier": len(new)}
    if common < 0.5:
        st_["erreur"] = (f"les colonnes du nouveau fichier ne correspondent pas aux données publiées "
                         f"({common:.0%} en commun) : exportez de Kobo avec le même format (libellés / noms XML).")
        return None, st_
    k = _key_col(new) if _key_col(new) == _key_col(old) else None
    if k:
        new = new.drop_duplicates(subset=k, keep="last")
        nk, ok_ = new[k].astype(str), old[k].astype(str)
        kept = old[~ok_.isin(set(nk))]
        st_["maj"] = int(nk.isin(set(ok_)).sum())
        st_["ajoutees"] = len(new) - st_["maj"]
        merged = pd.concat([kept, new], ignore_index=True)
    else:  # pas d'identifiant Kobo : on retire seulement les lignes strictement identiques
        merged = pd.concat([old, new], ignore_index=True)
        merged = merged.loc[~merged.astype(str).duplicated(keep="last")].reset_index(drop=True)
        st_["maj"] = 0
        st_["ajoutees"] = len(merged) - len(old)
    st_["conservees"] = len(merged) - len(new)  # lignes absentes du nouveau fichier mais gardées
    sub = next((c for c in merged.columns if _norm(c) == "_submission_time"), None)
    if sub is not None:
        order = pd.to_datetime(merged[sub], errors="coerce", utc=True)
        merged = merged.iloc[order.argsort(kind="stable")].reset_index(drop=True)
    st_["total"] = len(merged)
    return merged, st_


def df_to_xlsx_bytes(df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    df.to_excel(buf, index=False)
    return buf.getvalue()


def published_raw(slot: str) -> pd.DataFrame | None:
    """Données actuellement publiées pour cet emplacement (None si rien n'a été publié)."""
    meta = sync_slot(slot)
    return load_slot_raw(slot, meta["version"]) if meta else None


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
    flash = st.session_state.pop("admin_flash", None)
    if flash:
        (st.success if flash[0] else st.error)(flash[1])
    ups = st.file_uploader("Fichiers de données (.xlsx / .csv) — ECAM ou UPAS, un ou plusieurs",
                           type=["xlsx", "xls", "csv"], key="upload", accept_multiple_files=True)
    merge = st.toggle("ECAM : ajouter aux données déjà publiées (fusion)", value=True, key="ecam_merge",
                      help="Les fiches déjà publiées sont conservées, les nouvelles sont ajoutées et les fiches "
                           "modifiées sur Kobo sont mises à jour (identifiant `_uuid`). Désactivez pour remplacer.")
    old_ecam = published_raw("ecam") if merge else None
    ready = []
    for up in ups or []:
        try:
            rawf = read_bytes(up.getvalue(), up.name)
            slot = classify(rawf)
        except Exception as e:  # fichier illisible
            st.error(f"{up.name} : illisible ({e})")
            continue
        if slot is None or rawf.empty:
            st.error(f"{up.name} : type de fichier non reconnu (ni export ECAM, ni fichier UPAS).")
            continue
        data, info = up.getvalue(), f"{len(rawf)} lignes"
        if slot == "ecam" and old_ecam is not None:
            merged, s = merge_ecam(old_ecam, rawf)
            if merged is None:
                st.error(f"{up.name} : fusion impossible — {s['erreur']}")
                continue
            data = df_to_xlsx_bytes(merged)
            info = (f"fusion : {s['anciennes']} déjà publiées + {s['fichier']} dans le fichier → "
                    f"**{s['ajoutees']} nouvelles**, {s['maj']} mises à jour, {s['conservees']} conservées · "
                    f"**{s['total']} au total**")
        st.caption(f"✓ **{up.name}** → {SLOTS[slot]['label']} · {info}")
        ready.append((slot, up.name, data))
    token = kobo_token()
    if ups:
        if st.button(f"📤 Publier {len(ready)} fichier(s) pour tous", type="primary", width="stretch",
                     disabled=not ready):
            msgs, all_ok = [], True
            with st.spinner("Publication…"):
                for slot, name, data in ready:
                    ok, msg = publish(slot, data, name)
                    all_ok &= ok
                    msgs.append(("✓ " if ok else "✗ ") + msg)
            if all_ok and token and any(s == "ecam" for s, _, _ in ready):
                bar = st.progress(0.0, "Archivage des photos…")
                res = archive_photos(read_bytes(next(d for s, _, d in ready if s == "ecam"), "x.xlsx"),
                                     token, bar.progress)
                bar.empty()
                msgs.append(archive_report(res))
            elif all_ok and any(s == "ecam" for s, _, _ in ready):
                msgs.append("⚠ Photos non archivées : clé Kobo absente (`KOBO_TOKEN`).")
            if all_ok:
                st.session_state["admin_flash"] = (True, "\n\n".join(msgs))
                st.rerun()
            st.error("\n\n".join(msgs))
    st.caption("Stockage : " + ("**GitHub privé** ✓ (permanent)" if gh_cfg() else
                                "⚠ **serveur uniquement** — perdu au redémarrage. Configurez `GITHUB_TOKEN` et `DATA_REPO`."))
    cur = published_raw("ecam")
    if cur is not None:
        urls = photo_urls(cur)
        have = archived_keys()
        missing = sum(photo_key(u) not in have for u in urls)
        st.caption(f"ECAM publié : **{len(cur)} fiches** · photos archivées : **{len(urls) - missing}/{len(urls)}**")
        c1, c2 = st.columns(2)
        if c1.button(f"🖼 Archiver {missing} photo(s)", width="stretch", disabled=not (missing and token),
                     help=None if token else "Clé Kobo (KOBO_TOKEN) nécessaire."):
            bar = st.progress(0.0, "Archivage des photos…")
            res = archive_photos(cur, token, bar.progress)
            bar.empty()
            st.session_state["admin_flash"] = (not res["erreur"], archive_report(res))
            st.rerun()
        c2.download_button("💾 Tout l'ECAM (.xlsx)", df_to_xlsx_bytes(cur), "ecam_toutes_les_fiches.xlsx",
                           width="stretch", help="Sauvegarde de toutes les fiches publiées (fusionnées).")
    st.text_input("Jeton API Kobo (photos)", type="password", key="kobo_token",
                  help="Facultatif si KOBO_TOKEN est défini dans les Secrets.")
    if st.button("Se déconnecter", width="stretch"):
        st.session_state["is_admin"] = False
        st.rerun()


@st.fragment(run_every=30)
def watch_updates(slots: tuple[str, ...], key: str) -> None:
    """Vérifie toutes les 30 s si l'administrateur a publié de nouvelles données ; recharge si oui."""
    try:
        v = current_version(slots)
    except Exception:
        return
    seen = st.session_state.get(key)
    if seen is None:
        st.session_state[key] = v
    elif v != seen:
        st.session_state[key] = v
        st.toast("🔄 Nouvelles données publiées — mise à jour du tableau de bord.")
        st.rerun(scope="app")


# ----------------------------------------------------------------------------
# Graphiques
# ----------------------------------------------------------------------------
def base_layout(fig: go.Figure, h: int) -> go.Figure:
    fig.update_layout(height=h, margin=dict(l=0, r=0, t=10, b=0), paper_bgcolor="rgba(0,0,0,0)",
                      plot_bgcolor="rgba(0,0,0,0)", font=dict(family=FONT, color="#6a7b69"), showlegend=False)
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
                            f"<span style='font-size:12px;color:#6a7b69'>{label}</span>",
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
# UPAS — 3 fichiers liés par le code producteur :
#   engagements des ménages → inspections CSSVD → promesses de taille
# (contacts téléphoniques et noms des conjointes volontairement non chargés)
# ----------------------------------------------------------------------------
UPAS_STEPS = {"eng": ("Engagement du ménage", GREEN_L), "insp": ("Inspection CSSVD", GREEN),
              "prom": ("Promesse de taille", DARK)}
MOIS_ABR = {f"{i + 1:02d}": m for i, m in enumerate(MOIS)}


def _col(raw: pd.DataFrame, name: str) -> pd.Series:
    """Colonne par nom (insensible aux accents / majuscules) ; série vide si absente."""
    want = _norm(name)
    for c in raw.columns:
        if _norm(c) == want:
            return raw[c]
    return pd.Series(pd.NA, index=raw.index)


def _txt(s: pd.Series) -> pd.Series:
    t = s.astype("string").str.strip()
    return t.mask(t.isin(["", "0", "nan", "None"]))


def _gps(s: pd.Series) -> tuple[pd.Series, pd.Series]:
    parts = s.astype("string").str.split(",", n=1, expand=True)
    if parts.shape[1] < 2:
        return pd.Series(np.nan, index=s.index), pd.Series(np.nan, index=s.index)
    return _num(parts[0]), _num(parts[1])


COACH_PAT = ("coach", "agent", "enqueteur", "technicien", "collecteur", "animateur")


def _coach(raw: pd.DataFrame) -> pd.Series:
    """Coach de la fiche (ID et/ou nom), si une colonne coach / agent / enquêteur existe dans le fichier."""
    cols = [c for c in raw.columns if any(k in _norm(c) for k in COACH_PAT) and "contact" not in _norm(c)
            and "tel" not in _norm(c).split()]
    if not cols:
        return pd.Series(pd.NA, index=raw.index, dtype="string")
    parts = [_txt(raw[c]) for c in cols[:2]]  # ex. « ID coach » + « Nom coach »
    out = parts[0]
    for x in parts[1:]:
        out = (out.fillna("") + " " + x.fillna("")).str.strip()
    return out.mask(out == "")


def _base(raw: pd.DataFrame, date_col: str) -> pd.DataFrame:
    d = pd.DataFrame(index=raw.index)
    d["code"] = _txt(_col(raw, "Code Producteur"))
    d["parcelle"] = _txt(_col(raw, "Code Parcelle"))
    d["section"] = _txt(_col(raw, "Section")).str.upper()
    d["localite"] = _txt(_col(raw, "Localite")).str.title()
    d["date"] = pd.to_datetime(_col(raw, date_col), errors="coerce")
    d["synchro"] = pd.to_datetime(_col(raw, "Synchro le"), errors="coerce")
    d["lat"], d["lon"] = _gps(_col(raw, "Position GPS"))
    d["coach"] = _coach(raw)
    return d


def upas_eng(raw: pd.DataFrame) -> pd.DataFrame:
    d = _base(raw, "Date Engagement")
    d["menage"] = _txt(_col(raw, "Type Menage"))
    d["superficie"] = _num(_col(raw, "Superficie Totale"))
    d["conjointe"] = _txt(_col(raw, "Nom Conjointe")).notna()  # seulement Oui/Non, pas le nom
    return d


def upas_insp(raw: pd.DataFrame) -> pd.DataFrame:
    d = _base(raw, "Date Observation")
    d["connait_signes"] = _oui(_col(raw, "Producteur Connait Signes"))
    d["signes"] = _oui(_col(raw, "Parcelle Presente Signes"))
    d["foyers"] = _oui(_col(raw, "Presence Foyers Cssvd"))
    d["nb_foyers"] = _num(_col(raw, "Nombre Foyers"))
    d["arbres_foyer"] = _num(_col(raw, "Arbres Infestes Par Foyer"))
    d["sympt_princ"] = _txt(_col(raw, "Symptomes Principaux"))
    d["sympt_sec"] = _txt(_col(raw, "Symptomes Secondaires"))
    d["annee_signes"] = _num(_col(raw, "Annee Apparition Signes"))
    d["coop_informee"] = _oui(_col(raw, "Cooperative Informee"))
    d["superficie"] = _num(_col(raw, "Superficie Totale Visitee"))
    d["age"] = _num(_col(raw, "Age Moyen Cacaoyers"))
    d["varietes"] = _txt(_col(raw, "Varietes Cultivees"))
    d["densite"] = _txt(_col(raw, "Densite Plantation"))
    d["dec_signes"] = _oui(_col(raw, "Decision Signes Visibles"))
    d["dec_voisines"] = _oui(_col(raw, "Decision Parcelles Voisines"))
    d["etat"] = _txt(_col(raw, "Etat Parcelle"))
    d["agrofo"] = _txt(_col(raw, "Pratiques Agroforesterie"))
    return d


def upas_prom(raw: pd.DataFrame) -> pd.DataFrame:
    d = _base(raw, "Date Promesse")
    d["sup_tailler"] = _num(_col(raw, "Superficie A Tailler"))
    d["periode"] = _txt(_col(raw, "Periode Souhaitee"))
    d["accepte"] = _num(_col(raw, "Engagement Accepte")) == 1
    d["contribution"] = _oui(_col(raw, "Contribution Expliquee"))
    d["inspection"] = _oui(_col(raw, "Inspection Cssvd Realisee"))
    d["eligible"] = _oui(_col(raw, "Parcelle Eligible"))
    return d


def periode_label(p) -> str:
    s = str(p)
    if re.fullmatch(r"\d{4}-\d{2}", s):
        return f"{MOIS_ABR.get(s[5:], s[5:])} {s[:4]}"
    return s


def split_counts(s: pd.Series) -> pd.Series:
    """Compte les réponses à choix multiples « a | b | c »."""
    items = s.dropna().astype(str).str.split(r"\s*\|\s*").explode().str.strip()
    return items[items != ""].value_counts()


def upas_suivi(eng: pd.DataFrame, insp: pd.DataFrame, prom: pd.DataFrame) -> pd.DataFrame:
    """Une ligne par producteur : où en est-il dans le parcours ?"""
    frames = [t[["code", "section", "localite", "lat", "lon"]] for t in (eng, insp, prom) if not t.empty]
    if not frames:
        return pd.DataFrame()
    base = pd.concat(frames).dropna(subset=["code"]).groupby("code").first()
    e = eng.dropna(subset=["code"]).groupby("code").agg(date_eng=("date", "min"), menage=("menage", "first"),
                                                        superficie=("superficie", "max"), conjointe=("conjointe", "max"))
    i = insp.dropna(subset=["code"]).groupby("code").agg(date_insp=("date", "min"), signes=("signes", "max"),
                                                         age=("age", "max"), densite=("densite", "first"))
    p = prom.dropna(subset=["code"]).groupby("code").agg(date_prom=("date", "min"), periode=("periode", "first"),
                                                         eligible=("eligible", "max"))
    s = base.join(e).join(i).join(p)
    s["engage"] = s.index.isin(e.index)
    s["inspecte"] = s.index.isin(i.index)
    s["promesse"] = s.index.isin(p.index)
    s["complet"] = s["engage"] & s["inspecte"] & s["promesse"]

    def etape(r) -> str:
        if r.complet:
            return "✅ Parcours complet"
        if r.promesse and not r.inspecte:
            return "⚠ Promesse sans inspection"
        if r.inspecte and not r.promesse:
            return "🔍 Inspecté — promesse à faire"
        if r.engage and not r.inspecte:
            return "📝 Engagé — inspection à faire"
        return "Autre"

    s["etape"] = s.apply(etape, axis=1)
    return s.reset_index()


def upas_filter(d: pd.DataFrame, secs: list, dates) -> pd.DataFrame:
    if d.empty:
        return d
    if secs:
        d = d[d["section"].isin(secs)]
    if isinstance(dates, (tuple, list)) and len(dates) == 2:
        d = d[d["date"].isna() | d["date"].dt.date.between(dates[0], dates[1])]
    return d


def funnel_html(steps: list[tuple[str, int]]) -> str:
    top = max(steps[0][1], 1)
    out = []
    prev = None
    for lab, v in steps:
        w = v / top * 100
        conv = f"<span class='fn-conv'>{v / prev * 100:.0f}%</span>" if prev else ""
        out.append(f"<div class='fn-row'><div class='fn-lab'>{lab}{conv}</div>"
                   f"<div class='fn-tr'><div style='width:{max(w, 3):.0f}%'></div><b>{v}</b></div></div>")
        prev = v or None
    return "".join(out)


def section_progress(suivi: pd.DataFrame) -> go.Figure:
    g = suivi.groupby("section")[["engage", "inspecte", "promesse"]].sum().sort_values("engage")
    fig = go.Figure()
    for col, (lab, colr) in zip(["engage", "inspecte", "promesse"], UPAS_STEPS.values()):
        fig.add_bar(y=g.index, x=g[col], name=lab, orientation="h", marker=dict(color=colr, cornerradius=5),
                    text=g[col], textposition="inside", insidetextanchor="end")
    fig.update_layout(barmode="group", bargap=.25, legend=dict(orientation="h", y=1.1, x=0))
    fig.update_xaxes(visible=False)
    base_layout(fig, 60 + 46 * len(g)).update_layout(showlegend=True, margin=dict(l=0, r=0, t=30, b=0))
    return fig


def upas_quality(eng: pd.DataFrame, insp: pd.DataFrame, prom: pd.DataFrame) -> pd.DataFrame:
    """Contrôle de la règle 1 producteur = 1 engagement = 1 inspection = 1 promesse."""
    files = {"Engagements": eng, "Inspections": insp, "Promesses": prom}
    rows = []
    # coach(s) de chaque producteur, fichier par fichier
    coachs = {f: (t.dropna(subset=["code", "coach"]).groupby("code")["coach"]
                  .agg(lambda s: ", ".join(sorted(s.unique()))).to_dict() if "coach" in t else {})
              for f, t in files.items()}

    def add(kind, code, fichier, detail):
        names = [coachs[f].get(code) for f in files if f in fichier]
        coach = ", ".join(dict.fromkeys(n for n in names if n))
        rows.append({"Anomalie": kind, "Code producteur": code, "Coach": coach or None,
                     "Fichier(s)": fichier, "Détail": detail})

    # 0) fiches sans code producteur (impossibles à relier aux autres fichiers)
    for fname, t in files.items():
        n = int(t["code"].isna().sum())
        if n:
            add("Fiche sans code producteur", "—", fname, f"{n} fiche(s) exclue(s) des indicateurs")
    # 1) doublons dans un même fichier
    for fname, t in files.items():
        d = t[t["code"].notna() & t["code"].duplicated(keep=False)]
        for code, g in d.groupby("code"):
            same = g.drop(columns=["synchro"]).astype(str).nunique().le(1).all()
            add("Doublon", code, fname,
                f"{len(g)} fiches {'identiques (envoi en double)' if same else 'différentes'} · synchronisées le "
                + ", ".join(g["synchro"].dt.strftime("%d/%m %H:%M:%S").dropna()))
    # 2) code parcelle qui ne reprend pas le code producteur
    for fname, t in files.items():
        for r in t.dropna(subset=["code", "parcelle"]).drop_duplicates("code").itertuples():
            if str(r.code) not in str(r.parcelle):
                add("Code parcelle incohérent", r.code, fname, f"parcelle « {r.parcelle} » ne correspond pas au producteur")
    # 3) même producteur, informations différentes d'un fichier à l'autre
    ref = eng.dropna(subset=["code"]).drop_duplicates("code").set_index("code")
    for fname, t in (("Inspections", insp), ("Promesses", prom)):
        o = t.dropna(subset=["code"]).drop_duplicates("code").set_index("code")
        for code in o.index.intersection(ref.index):
            for col, lab in (("parcelle", "Code parcelle"), ("section", "Section"), ("localite", "Localité")):
                a, b = ref.at[code, col], o.at[code, col]
                if pd.notna(a) and pd.notna(b) and str(a) != str(b):
                    add(f"{lab} différent(e)", code, f"Engagements ↔ {fname}", f"« {a} » ↔ « {b} »")
    # 4) ordre logique : Engagement → Inspection CSSVD → Promesse de taille
    #    (une inspection sans promesse n'est PAS une anomalie : c'est l'étape suivante à faire)
    e, i, p = (set(t["code"].dropna()) for t in (eng, insp, prom))
    for code in sorted(i - e):
        add("Inspection sans engagement", code, "Inspections", "absent du fichier des engagements")
    for code in sorted(p - e):
        add("Promesse sans engagement", code, "Promesses", "absent du fichier des engagements")
    declare = prom.dropna(subset=["code"]).drop_duplicates("code").set_index("code")["inspection"]
    for code in sorted(p - i):
        dit = " — pourtant déclarée « Oui » dans la promesse" if declare.get(code) == True else ""  # noqa: E712
        add("Promesse sans inspection", code, "Promesses", f"aucune inspection CSSVD dans le fichier{dit}")
    d_insp = insp.dropna(subset=["code"]).groupby("code")["date"].min()
    d_prom = prom.dropna(subset=["code"]).groupby("code")["date"].min()
    for code in sorted(set(d_insp.index) & set(d_prom.index)):
        if pd.notna(d_insp[code]) and pd.notna(d_prom[code]) and d_prom[code] < d_insp[code]:
            add("Promesse avant inspection", code, "Inspections ↔ Promesses",
                f"promesse le {d_prom[code]:%d/%m/%Y}, inspection le {d_insp[code]:%d/%m/%Y}")
    for code in sorted(i & p):
        if declare.get(code) == False:  # noqa: E712
            add("Déclaration incohérente", code, "Promesses",
                "« inspection CSSVD réalisée » = Non dans la promesse, alors qu'une inspection existe")
    out = pd.DataFrame(rows, columns=["Anomalie", "Code producteur", "Coach", "Fichier(s)", "Détail"])
    return out if out["Coach"].notna().any() else out.drop(columns="Coach")


def dedupe(t: pd.DataFrame) -> pd.DataFrame:
    """Une seule fiche par code producteur (la plus récemment synchronisée) ; fiches sans code exclues."""
    if t.empty:
        return t
    t = t.sort_values("synchro")
    return t[t["code"].notna()].drop_duplicates("code", keep="last")  # fiches sans code : non reliables, exclues


def render_upas(page: str, pop) -> None:
    raws = {s: load_slot(s) for s in COOP_SLOTS["UPAS"]}
    watch_updates(tuple(COOP_SLOTS["UPAS"]), "data_version_UPAS")
    eng_all = upas_eng(raws["upas_eng"][0]) if raws["upas_eng"][0] is not None else upas_eng(pd.DataFrame())
    insp_all = upas_insp(raws["upas_insp"][0]) if raws["upas_insp"][0] is not None else upas_insp(pd.DataFrame())
    prom_all = upas_prom(raws["upas_prom"][0]) if raws["upas_prom"][0] is not None else upas_prom(pd.DataFrame())

    with pop:
        for s, (r, name, meta) in raws.items():
            maj = ""
            if meta.get("date"):
                maj = " · " + pd.Timestamp(meta["date"]).tz_convert("Africa/Abidjan").strftime("%d/%m %H:%M")
            st.caption(f"{SLOTS[s]['label'].split('· ')[1]} : **{name or 'aucun fichier'}**"
                       f"{f' · {len(r)} lignes' if r is not None else ''}{maj}")
        secs = sorted(set(pd.concat([eng_all["section"], insp_all["section"], prom_all["section"]]).dropna()))
        f_sec = st.multiselect("Sections", secs, key="upas_sec")
        alld = pd.concat([eng_all["date"], insp_all["date"], prom_all["date"]]).dropna()
        f_dates = ()
        if not alld.empty:
            dmin, dmax = alld.min().date(), alld.max().date()
            f_dates = st.date_input("Période", (dmin, dmax), min_value=dmin, max_value=dmax, key="upas_dates")
        objectif = st.number_input("Objectif (nombre de producteurs)", min_value=0, value=1000, step=50,
                                   key="upas_obj")
        st.divider()
        admin_panel()

    if all(r is None for r, _, _ in raws.values()):
        st.info("Aucune donnée UPAS publiée pour le moment. L'administrateur peut charger les 3 fichiers "
                "(engagements, inspections CSSVD, promesses de taille) dans ⚙ Paramètres.")
        return

    eng, insp, prom = (upas_filter(t, f_sec, f_dates) for t in (eng_all, insp_all, prom_all))
    anomalies = upas_quality(eng, insp, prom)
    n_lignes = {"eng": len(eng), "insp": len(insp), "prom": len(prom)}
    eng, insp, prom = dedupe(eng), dedupe(insp), dedupe(prom)  # 1 fiche par producteur dans chaque fichier
    n_double = {k: n_lignes[k] - len(t) for k, t in (("eng", eng), ("insp", insp), ("prom", prom))}
    suivi = upas_suivi(eng, insp, prom)
    n_eng, n_insp, n_prom = eng["code"].nunique(), insp["code"].nunique(), prom["code"].nunique()
    n_suivis = len(suivi)
    n_complet = int(suivi["complet"].sum()) if n_suivis else 0
    cible = objectif or max(n_suivis, 1)
    hectares = eng.drop_duplicates("code")["superficie"].sum()

    # --- Bienvenue + pills ------------------------------------------------------
    w1, w2 = st.columns([1.55, 1], vertical_alignment="bottom")
    with w1:
        k_eng, k_insp, k_prom = n_eng / cible, n_insp / cible, n_prom / cible
        html(f"""
        <div class='welcome'>Coopérative <span style='color:var(--accent)'>UPAS</span></div>
        <div class='kp-row'>
          <div class='kp'><span class='kp-lab'>Engagement du ménage</span>
            <div class='kp-pill kp-striped'><div class='kp-fill' style='width:{min(k_eng * 100, 100):.1f}%'></div>
            <span>{pct(k_eng)}</span></div></div>
          <div class='kp'><span class='kp-lab'>Inspection CSSVD</span>
            <div class='kp-pill kp-green'><span>{pct(k_insp)}</span></div></div>
          <div class='kp'><span class='kp-lab'>Promesse de taille</span>
            <div class='kp-pill kp-dark'><span>{pct(k_prom)}</span></div></div>
        </div>""")
    with w2:
        n_sec = suivi["section"].nunique() if n_suivis else 0
        n_loc = suivi["localite"].nunique() if n_suivis else 0
        html(f"""
        <div class='stats' style='display:flex; justify-content:flex-end; gap:34px; padding-bottom:4px'>
          <div><div class='stat'><span class='ic'>🗺️</span><span class='num'>{n_sec}</span></div><div class='lab'>Sections</div></div>
          <div><div class='stat'><span class='ic'>🏘️</span><span class='num'>{n_loc}</span></div><div class='lab'>Localités</div></div>
          <div><div class='stat'><span class='ic'>🌳</span><span class='num'>{fmt(hectares)}</span></div><div class='lab'>Hectares</div></div>
        </div>""")

    def note(k: str) -> str:
        if n_double[k] > 0:
            return f"<b>{n_double[k]} en double</b>"
        return ""

    def r(a, b):
        return a / b * 100 if b else 0

    html(kpi_cards([
        {"lab": "Producteurs suivis", "ico": "👥", "v": n_suivis, "p": r(n_suivis, cible),
         "sub": f"{fmt(max(cible - n_suivis, 0))} restants pour atteindre l'objectif"},
        {"lab": "Engagement du ménage", "ico": "🤝", "v": n_eng, "p": r(n_eng, n_suivis),
         "sub": " · ".join(x for x in [f"{r(n_eng, n_suivis):.0f} % des suivis", note("eng")] if x)},
        {"lab": "Inspection CSSVD", "ico": "🔍", "v": n_insp, "p": r(n_insp, n_eng),
         "sub": " · ".join(x for x in [f"{r(n_insp, n_eng):.0f} % des engagés", note("insp")] if x)},
        {"lab": "Promesse de taille", "ico": "✂️", "v": n_prom, "p": r(n_prom, n_insp), "cls": "amber",
         "sub": " · ".join(x for x in [f"{r(n_prom, n_insp):.0f} % des inspectés", note("prom")] if x)},
        {"lab": "Parcours complet", "ico": "🏁", "v": n_complet, "p": r(n_complet, n_suivis),
         "sub": f"{r(n_complet, n_suivis):.0f} % des suivis"},
    ], cible))

    if not anomalies.empty:
        resume = " · ".join(f"{n} {k.lower()}" for k, n in anomalies["Anomalie"].value_counts().items())
        if "Coach" in anomalies:
            par_coach = (anomalies["Coach"].dropna().str.split(", ").explode().value_counts())
            resume += " · coach(s) : " + ", ".join(f"{c} ({n})" for c, n in par_coach.items())
        with st.container(key="card_doublons"):
            with st.expander(f"⚠ Contrôle qualité UPAS : {len(anomalies)} anomalie(s) — {resume}"):
                if "Coach" not in anomalies:
                    st.caption("ℹ Les fichiers UPAS ne contiennent pas le nom du coach : ajoutez une colonne "
                               "« Coach » (ou « Agent », « Enquêteur ») à l'export pour l'afficher ici.")
                st.caption("Règle attendue : 1 producteur = 1 engagement → 1 inspection CSSVD → 1 promesse de taille, "
                           "dans cet ordre (l'inspection précède la promesse). "
                           "Les doublons sont comptés une seule fois dans le tableau de bord. "
                           "Les fichiers sont reliés par le **code producteur** : un code parcelle mal saisi "
                           "ne fausse pas les compteurs, mais doit être corrigé à la source.")
                st.dataframe(anomalies, width="stretch", hide_index=True)
                st.download_button("⬇ Exporter les anomalies (Excel)", to_xlsx(anomalies, "Anomalies UPAS"),
                                   file_name=f"upas_anomalies_{pd.Timestamp.now():%Y%m%d}.xlsx",
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    st.write("")

    ins_ok = int(insp.loc[insp["signes"].notna(), "code"].nunique())
    n_signes = int(insp.loc[insp["signes"] == True, "code"].nunique())  # noqa: E712
    r_sain = (ins_ok - n_signes) / ins_ok if ins_ok else np.nan

    # ======================================================================
    if page == "Tableau de bord":
        left, right = st.columns([3, 1.05], gap="medium")
        with left:
            a1, a2, a3 = st.columns(3, gap="medium")
            with a1, st.container(key="card_profile"):
                html(f"""<div class='muted' style='color:rgba(255,255,255,.85)'>🧭 Parcours producteur</div>
                <div style='margin-top:14px'>{funnel_html([("Engagés", n_eng), ("Inspectés", n_insp),
                                                           ("Promesse", n_prom), ("Complet", n_complet)])}</div>""")
            events = pd.concat([eng[["date"]].assign(agent_id="Engagement"),
                                insp[["date"]].assign(agent_id="Inspection"),
                                prom[["date"]].assign(agent_id="Promesse")])
            fig, week = activity_chart(events)
            with a2, st.container(key="card_activity"):
                html(f"""<span class='arrow-btn'>↗</span><p class='card-title'>Activité</p>
                <div style='display:flex; align-items:center; gap:12px; margin-top:8px'>
                  <span class='big'>{week}</span><span class='muted'>fiches<br>7 derniers jours</span></div>""")
                st.plotly_chart(fig, config=CFG, width="stretch")
            with a3, st.container(key="card_ring"):
                html("<span class='arrow-btn'>↗</span><p class='card-title'>Parcelles saines</p>")
                st.plotly_chart(ring_chart(r_sain, "sans signes de CSSVD"), config=CFG, width="stretch")
                n_foy = int((insp["foyers"] == True).sum())  # noqa: E712
                html(f"""<div style='display:flex; gap:8px; justify-content:center; flex-wrap:wrap'>
                <span class='cal-chip'>✓ Saines · {ins_ok - n_signes}</span><span class='cal-chip'>⚠ Signes · {n_signes}</span>
                <span class='cal-chip' style='background:var(--hero);color:#fff;border-color:transparent'>Foyers · {n_foy}</span></div>""")

            b1, b2 = st.columns([1, 2], gap="medium")
            with b1, st.container(key="card_accordion"):
                with st.expander("Producteurs par section", expanded=True):
                    html(minibars(suivi["section"].value_counts().head(8), n_suivis))
                with st.expander("Type de ménage"):
                    html(minibars(eng["menage"].value_counts()))
                with st.expander("Variétés cultivées"):
                    html(minibars(split_counts(insp["varietes"])))
                with st.expander("Densité de plantation"):
                    html(minibars(insp["densite"].value_counts()))
                with st.expander("Ménages avec conjointe"):
                    html(minibars(eng.drop_duplicates("code")["conjointe"].map({True: "Avec conjointe",
                                                                                False: "Sans conjointe"}).value_counts()))
            with b2, st.container(key="card_sections"):
                html("<p class='card-title'>Avancement par section</p>")
                if n_suivis:
                    st.plotly_chart(section_progress(suivi), config=CFG, width="stretch")

        with right:
            per = prom["periode"].value_counts().sort_index()
            tot = max(int(per.sum()), 1)
            seg_cols = [GREEN, DARK, GREY, GREEN_L, "#3e7a3c"]
            seg = "".join(
                f"<div style='flex:{max(v / tot, .1):.3f}'><span class='muted' style='color:var(--ink)'>{v / tot * 100:.0f}%</span>"
                f"<div class='bar' style='background:{seg_cols[i % 5]}'>{periode_label(k).split()[0] if v / tot > .2 else ''}</div></div>"
                for i, (k, v) in enumerate(per.items()))
            with st.container(key="card_docs"):
                html(f"""<div style='display:flex; justify-content:space-between; align-items:center'>
                <p class='card-title'>Période de taille</p><span class='big' style='font-size:2rem'>{int(prom['sup_tailler'].sum())} ha</span></div>
                <div class='seg'>{seg}</div>""")
            etats = split_counts(insp["etat"]).head(7)
            icons = ["🌿", "🐛", "🌳", "☀️", "💧", "🍂", "🪓"]
            items = "".join(
                f"<div class='task'><div class='ic on'>{icons[i % 7]}</div><div><div class='t1'>{k}</div>"
                f"<div class='t2'>{v} producteur{'s' if v > 1 else ''} · {v / max(insp['code'].nunique(), 1) * 100:.0f}%</div></div>"
                f"<div class='ck on'>✓</div></div>" for i, (k, v) in enumerate(etats.items()))
            with st.container(key="dark_etat"):
                html(f"""<div style='display:flex; justify-content:space-between; align-items:center; margin-bottom:6px'>
                <span style='font-size:1.25rem'>État des parcelles</span><span style='font-size:2rem;font-weight:300'>{insp['code'].nunique()}</span></div>
                <div class='t2' style='color:#a8a8a8;font-size:.78rem;margin-bottom:6px'>observations lors des inspections CSSVD</div>
                {items}""")

    # ======================================================================
    elif page == "Suivi producteurs":
        c1, c2 = st.columns([1, 1.4], gap="medium")
        with c1, st.container(key="card_u_etape"):
            html("<p class='card-title'>Où en sont les producteurs ?</p>")
            if n_suivis:
                st.plotly_chart(bar_h(suivi["etape"].value_counts(), 260), config=CFG, width="stretch")
        with c2, st.container(key="card_u_sec"):
            html("<p class='card-title'>Avancement par section</p>")
            if n_suivis:
                st.plotly_chart(section_progress(suivi), config=CFG, width="stretch")
        with st.container(key="card_u_table"):
            html("<p class='card-title'>🧾 Liste de suivi des producteurs</p>"
                 "<div class='muted'>1 ligne = 1 producteur · ce qui est fait et ce qui reste à faire</div>")
            e1, e2 = st.columns(2)
            f_et = e1.multiselect("Étape", sorted(suivi["etape"].unique()) if n_suivis else [], key="u_et")
            q = e2.text_input("Rechercher", placeholder="Code producteur, localité…", key="u_q")
            t = pd.DataFrame({
                "Code producteur": suivi["code"], "Section": suivi["section"], "Localité": suivi["localite"],
                "Étape": suivi["etape"], "Type de ménage": suivi["menage"],
                "Superficie (ha)": suivi["superficie"].round(2),
                "Engagé le": suivi["date_eng"].dt.strftime("%d/%m/%Y"),
                "Inspecté le": suivi["date_insp"].dt.strftime("%d/%m/%Y"),
                "Signes CSSVD": suivi["signes"].map({True: "Oui", False: "Non"}),
                "Promesse le": suivi["date_prom"].dt.strftime("%d/%m/%Y"),
                "Période de taille": suivi["periode"].map(lambda p: periode_label(p) if pd.notna(p) else pd.NA),
                "Latitude": suivi["lat"], "Longitude": suivi["lon"],
            }) if n_suivis else pd.DataFrame()
            if f_et:
                t = t[t["Étape"].isin(f_et)]
            if q and not t.empty:
                t = t[t.astype(str).apply(lambda c: c.str.contains(q, case=False, regex=False)).any(axis=1)]
            st.dataframe(t, width="stretch", hide_index=True, height=420)
            stamp = pd.Timestamp.now().strftime("%Y%m%d")
            d1, d3, d2, _ = st.columns([1, 1.3, 1, 1])
            d1.download_button("⬇ Exporter Excel", to_xlsx(t, "Suivi UPAS"), file_name=f"upas_suivi_{stamp}.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                               type="primary", width="stretch", disabled=t.empty)
            d3.download_button("⬇ Excel par section", upas_xlsx_by_section(t) if not t.empty else b"",
                               file_name=f"upas_suivi_par_section_{stamp}.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                               type="primary", width="stretch", disabled=t.empty,
                               help="Une feuille « Résumé » + une feuille par section")
            d2.download_button("⬇ Exporter CSV", t.to_csv(index=False, sep=";").encode("utf-8-sig"),
                               file_name=f"upas_suivi_{stamp}.csv", mime="text/csv", width="stretch", disabled=t.empty)

    # ======================================================================
    elif page == "Inspections":
        if insp.empty:
            st.info("Aucune inspection CSSVD dans les données filtrées.")
            return
        c1, c2, c3 = st.columns(3, gap="medium")
        with c1, st.container(key="card_i_ring"):
            html("<p class='card-title'>Parcelles saines</p>")
            st.plotly_chart(ring_chart(r_sain, f"{ins_ok - n_signes} / {ins_ok} producteurs"), config=CFG, width="stretch")
        with c2, st.container(key="card_i_age"):
            html("<p class='card-title'>Âge moyen des cacaoyers</p>")
            bins = pd.cut(insp["age"], [0, 10, 20, 30, 40, 200], labels=["< 10 ans", "10-19", "20-29", "30-39", "40 ans +"],
                          right=False)
            st.plotly_chart(bar_h(bins.value_counts().sort_index(ascending=False), 260, DARK), config=CFG, width="stretch")
            html(f"<div class='muted' style='text-align:center'>Moyenne : <b>{insp['age'].mean():.0f} ans</b></div>")
        with c3, st.container(key="card_i_conn"):
            html("<p class='card-title'>Connaissances & décisions</p>")
            facts = [
                ("Producteur connaît les signes", insp["connait_signes"]),
                ("Signes visibles sur la parcelle", insp["signes"]),
                ("Foyers de CSSVD", insp["foyers"]),
                ("Décision : signes visibles", insp["dec_signes"]),
                ("Décision : parcelles voisines", insp["dec_voisines"]),
            ]
            html("".join(
                f"<div class='mini'><div class='nm'>{lab}</div><div class='tr'><div style='width:{(s == True).mean() * 100:.0f}%'></div></div>"  # noqa: E712
                f"<div class='vl'>{int((s == True).sum())}</div></div>" for lab, s in facts))  # noqa: E712
            html("<div class='muted' style='margin-top:8px'>Barre = part des inspections répondant « Oui »</div>")
        c4, c5 = st.columns(2, gap="medium")
        with c4, st.container(key="card_i_etat"):
            html("<p class='card-title'>État des parcelles</p>")
            st.plotly_chart(bar_h(split_counts(insp["etat"]), 320), config=CFG, width="stretch")
        with c5, st.container(key="card_i_agro"):
            html("<p class='card-title'>Pratiques d'agroforesterie</p>")
            st.plotly_chart(bar_h(split_counts(insp["agrofo"]), 320, "#3e7a3c"), config=CFG, width="stretch")
        c6, c7 = st.columns(2, gap="medium")
        with c6, st.container(key="card_i_var"):
            html("<p class='card-title'>Variétés cultivées</p>")
            st.plotly_chart(bar_h(split_counts(insp["varietes"]), 200, GREEN_L), config=CFG, width="stretch")
        with c7, st.container(key="card_i_dens"):
            html("<p class='card-title'>Densité de plantation</p>")
            st.plotly_chart(bar_h(insp["densite"].value_counts(), 200, DARK), config=CFG, width="stretch")
        with st.container(key="card_i_signes"):
            html("<p class='card-title'>⚠ Parcelles avec signes de CSSVD</p>")
            sig = insp[insp["signes"] == True]  # noqa: E712
            if sig.empty:
                st.success(f"Aucun des {ins_ok} producteurs inspectés ne présente de signes de CSSVD.")
            else:
                t = sig[["code", "parcelle", "section", "localite", "date", "sympt_princ", "sympt_sec",
                         "nb_foyers", "arbres_foyer", "annee_signes", "coop_informee", "lat", "lon"]].rename(columns={
                    "code": "Code producteur", "parcelle": "Code parcelle", "section": "Section", "localite": "Localité",
                    "date": "Date", "sympt_princ": "Symptômes principaux", "sympt_sec": "Symptômes secondaires",
                    "nb_foyers": "Nb foyers", "arbres_foyer": "Arbres infestés / foyer",
                    "annee_signes": "Année apparition", "coop_informee": "Coopérative informée",
                    "lat": "Latitude", "lon": "Longitude"})
                st.dataframe(t, width="stretch", hide_index=True)
                st.download_button("⬇ Exporter Excel", to_xlsx(t, "Parcelles CSSVD"),
                                   file_name="upas_parcelles_cssvd.xlsx", type="primary",
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    # ======================================================================
    elif page == "Promesses":
        if prom.empty:
            st.info("Aucune promesse de taille dans les données filtrées.")
            return
        c1, c2, c3 = st.columns(3, gap="medium")
        with c1, st.container(key="card_p_per"):
            html("<p class='card-title'>Période de taille souhaitée</p>")
            per = prom["periode"].value_counts().sort_index()
            per.index = [periode_label(p) for p in per.index]
            st.plotly_chart(bar_h(per[::-1], 240), config=CFG, width="stretch")
        with c2, st.container(key="card_p_insp"):
            html("<p class='card-title'>Inspection CSSVD réalisée</p>")
            st.plotly_chart(ring_chart(prom["code"].isin(set(insp["code"].dropna())).mean(), "des promesses"),
                            config=CFG, width="stretch")
        with c3, st.container(key="card_p_elig"):
            html("<p class='card-title'>Conditions</p>")
            facts = [("Engagement accepté", prom["accepte"]), ("Contribution expliquée", prom["contribution"]),
                     ("Parcelle éligible", prom["eligible"]), ("Inspection réalisée", prom["inspection"])]
            html("".join(
                f"<div class='mini'><div class='nm'>{lab}</div><div class='tr'><div style='width:{(s == True).mean() * 100:.0f}%'></div></div>"  # noqa: E712
                f"<div class='vl'>{int((s == True).sum())}</div></div>" for lab, s in facts))  # noqa: E712
            html(f"<div style='margin-top:14px' class='big'>{int(prom['sup_tailler'].sum())} ha</div>"
                 "<div class='muted'>superficie totale à tailler</div>")
        with st.container(key="card_p_todo"):
            todo = prom[~prom["code"].isin(set(insp["code"].dropna()))]  # absente du fichier des inspections
            html(f"<p class='card-title'>📋 Promesses sans inspection CSSVD · {len(todo)}</p>"
                 "<div class='muted'>Parcelles à inspecter avant la taille</div>")
            if todo.empty:
                st.success("Toutes les promesses de taille ont une inspection CSSVD.")
            else:
                t = todo[["code", "parcelle", "section", "localite", "date", "periode", "lat", "lon"]].assign(
                    periode=todo["periode"].map(periode_label), date=todo["date"].dt.strftime("%d/%m/%Y")).rename(columns={
                    "code": "Code producteur", "parcelle": "Code parcelle", "section": "Section",
                    "localite": "Localité", "date": "Date promesse", "periode": "Période de taille",
                    "lat": "Latitude", "lon": "Longitude"})
                st.dataframe(t, width="stretch", hide_index=True)
                st.download_button("⬇ Exporter la liste (Excel)", to_xlsx(t, "Inspections à faire"),
                                   file_name="upas_inspections_a_faire.xlsx", type="primary",
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    # ======================================================================
    elif page == "Carte":
        g = suivi.dropna(subset=["lat", "lon"]) if n_suivis else pd.DataFrame()
        with st.container(key="card_u_map"):
            html(f"<p class='card-title'>Parcelles géolocalisées · {len(g)} / {n_suivis}</p>")
            if g.empty:
                st.info("Aucune position GPS dans les données filtrées.")
            else:
                colors = {"✅ Parcours complet": "#c9a227", "⚠ Promesse sans inspection": "#d9534f",
                          "🔍 Inspecté — promesse à faire": GREEN, "📝 Engagé — inspection à faire": GREY,
                          "Autre": "#7a7a7a"}
                fig = px.scatter_map(g, lat="lat", lon="lon", color="etape", zoom=10, color_discrete_map=colors,
                                     hover_name="code", hover_data={"section": True, "localite": True,
                                                                    "lat": False, "lon": False},
                                     map_style="open-street-map")
                fig.update_traces(marker=dict(size=11))
                fig.update_layout(height=600, margin=dict(l=0, r=0, t=0, b=0), paper_bgcolor="rgba(0,0,0,0)",
                                  legend=dict(orientation="h", y=1.02, x=0, bgcolor="rgba(255,255,255,.75)", title=None),
                                  font=dict(family=FONT))
                st.plotly_chart(fig, width="stretch", config=CFG)

    # ======================================================================
    else:
        with st.container(key="card_u_raw"):
            html("<p class='card-title'>Données UPAS</p>"
                 "<div class='muted'>Contacts téléphoniques et noms des conjointes non affichés</div>")
            tabs = st.tabs(["Engagements des ménages", "Inspections CSSVD", "Promesses de taille"])
            for tab, slot in zip(tabs, COOP_SLOTS["UPAS"]):
                r = raws[slot][0]
                with tab:
                    if r is None:
                        st.info("Aucun fichier publié.")
                        continue
                    r = r[[c for c in r.columns if not re.search(r"contact|nom conjointe", _norm(c))]]
                    st.dataframe(r, width="stretch", height=480)
                    st.download_button("⬇ Télécharger (CSV)", r.to_csv(index=False, sep=";").encode("utf-8-sig"),
                                       file_name=f"{slot}.csv", mime="text/csv", key=f"dl_{slot}")


def _tri(s: pd.Series):
    """Oui si au moins une fiche dit Oui, Non si seulement des Non, vide sinon."""
    if (s == True).any():  # noqa: E712
        return True
    if (s == False).any():  # noqa: E712
        return False
    return np.nan


def consolidate(d: pd.DataFrame, keys: list[str]) -> pd.DataFrame:
    """Fusionne les fiches en double : une ligne par clé (producteur, ou producteur + parcelle)."""
    agg = {}
    for c in d.columns:
        if c in keys:
            continue
        if c in ("engagement", "cssvd"):
            agg[c] = _tri
        elif c.startswith(("doc_", "sym_")):
            agg[c] = "max"
        elif c == "date":
            agg[c] = "max"
        else:
            agg[c] = "last"  # dernière valeur renseignée
    out = d.sort_values("date").groupby(keys, dropna=False, sort=False).agg(agg).reset_index()
    out["doc_aucun"] = ((out["doc_fiche"] != 1) & (out["doc_promesse"] != 1)).astype(int)
    for c in ("engagement", "cssvd"):
        out[c] = out[c].astype("object")
    return out



# ----------------------------------------------------------------------------
# Rapport hors ligne : un fichier HTML autonome (ECAM + UPAS) à partager par WhatsApp / e-mail.
# Uniquement des codes producteurs : ni nom, ni téléphone, ni position GPS.
# ----------------------------------------------------------------------------
def _yn(v):
    return None if pd.isna(v) else bool(v)


def _day(v):
    return v.strftime("%Y-%m-%d") if pd.notna(v) else None


def report_version() -> str:
    """Change dès que des données ECAM ou UPAS sont republiées (sert de clé de cache)."""
    parts = [current_version(tuple(SLOTS))]
    for s in SLOTS:
        parts += [f"{p}:{os.path.getmtime(p)}" for p in local_files(s)[:1]]
    return "|".join(parts)


@st.cache_data(show_spinner=False, max_entries=4)
def offline_report_html(version: str, objectif: int = 1000) -> bytes:
    import json

    tz = "Africa/Abidjan"
    df, raw0, source, _ = load()
    n_saisies = len(df)
    if raw0 is not df and merge_forms(raw0)[1]["groupes"]:
        df = normalize(merge_forms(raw0)[0])  # fiches fusionnées, comme dans l'application
    parc = consolidate(df, ["prod_code", "plantation"])
    pp = consolidate(parc[parc["prod_code"].notna()], ["prod_code"])
    ecam = [{"code": r.prod_code, "section": r.section if pd.notna(r.section) else "Sans section",
             "coach": r.agent_id if pd.notna(r.agent_id) else "",
             "genre": str(r.genre)[:1].upper() if pd.notna(r.genre) else "",
             "eng": _yn(r.engagement), "cssvd": _yn(r.cssvd),
             "fiche": int(r.doc_fiche == 1), "prom": int(r.doc_promesse == 1),
             "sup": None if pd.isna(r.sup_tot) else round(float(r.sup_tot), 2),
             "sym": [k for k in SYMPTOMES if getattr(r, k) == 1], "date": _day(r.date)} for r in pp.itertuples()]
    fiches = df.dropna(subset=["date"])
    days = (fiches.assign(d=fiches["date"].dt.strftime("%Y-%m-%d"), section=fiches["section"].fillna("Sans section"))
            .groupby(["d", "section"]).size().rename("n").reset_index().to_dict("records"))

    raws = {s: load_slot(s)[0] for s in COOP_SLOTS["UPAS"]}
    eng = upas_eng(raws["upas_eng"] if raws["upas_eng"] is not None else pd.DataFrame())
    insp = upas_insp(raws["upas_insp"] if raws["upas_insp"] is not None else pd.DataFrame())
    prom = upas_prom(raws["upas_prom"] if raws["upas_prom"] is not None else pd.DataFrame())
    n_anom = len(upas_quality(eng, insp, prom)) if len(eng) or len(insp) or len(prom) else 0
    eng, insp, prom = dedupe(eng), dedupe(insp), dedupe(prom)
    suivi = upas_suivi(eng, insp, prom)
    upas = [] if suivi.empty else [
        {"code": r.code, "section": r.section if pd.notna(r.section) else "Sans section",
         "localite": r.localite if pd.notna(r.localite) else "",
         "eng": bool(r.engage), "insp": bool(r.inspecte), "prom": bool(r.promesse), "etape": r.etape,
         "sup": None if pd.isna(r.superficie) else round(float(r.superficie), 2),
         "signes": _yn(r.signes), "date": _day(r.date_eng)} for r in suivi.itertuples()]
    upas_dates = pd.concat([eng["date"], insp["date"], prom["date"]]).dropna()

    payload = {
        "ecam": ecam, "ecam_days": days, "ecam_fiches": int(n_saisies), "ecam_dup": int(n_saisies - len(df)),
        "ecam_demo": source == "Données de démonstration",
        "upas": upas, "upas_anomalies": int(n_anom), "objectif": int(objectif or 1000),
        "sym_names": {k: v[0] for k, v in SYMPTOMES.items()},
        "asof_ecam": df["date"].max().strftime("%d/%m/%Y") if df["date"].notna().any() else "",
        "asof_upas": upas_dates.max().strftime("%d/%m/%Y") if len(upas_dates) else "",
        "generated": pd.Timestamp.now(tz=tz).strftime("%d/%m/%Y à %H:%M"),
    }
    data = json.dumps(payload, ensure_ascii=False, default=str).replace("</", "<\\/")
    with open(os.path.join(HERE, "rapport_template.html"), encoding="utf-8") as f:
        return f.read().replace("__DATA__", data).encode("utf-8")


def offline_button() -> None:
    """Bouton « Rapport hors ligne » : le fichier n'est généré qu'au clic."""
    version = report_version()
    obj = int(st.session_state.get("upas_obj", 1000) or 1000)
    st.download_button("📥 Rapport hors ligne", lambda: offline_report_html(version, obj),
                       file_name=f"thrive_track_rapport_{pd.Timestamp.now(tz='Africa/Abidjan'):%Y-%m-%d}.html",
                       mime="text/html", width="stretch", key="offline_dl",
                       help="Un fichier unique (ECAM + UPAS) qui s'ouvre sans l'application : à envoyer par "
                            "WhatsApp ou e-mail. Seulement les codes producteurs, sans noms ni GPS.")


# ----------------------------------------------------------------------------
# Application
# ----------------------------------------------------------------------------
html(CSS)

# --- En-tête ---------------------------------------------------------------
PAGES = {
    "ECAM": ["Tableau de bord", "Producteurs", "Coachs", "Symptômes", "Carte", "Données"],
    "UPAS": ["Tableau de bord", "Suivi producteurs", "Inspections", "Promesses", "Carte", "Données"],
}
h1, h0, h2, h4, h5, h3 = st.columns([1.2, 1.25, 4.6, 1.3, .38, 1.1], vertical_alignment="center")
with h1:
    html("<div class='brand'><div class='lg'>🌿</div><div><b>Thrive·Track</b><small>Suivi CSSVD</small></div></div>")
with h0:
    with st.container(key="coop"):
        coop = st.segmented_control("Coopérative", list(PAGES), default="ECAM", key="coop_sel",
                                    label_visibility="collapsed") or "ECAM"
with h2:
    with st.container(key="nav"):
        page = st.segmented_control("Navigation", PAGES[coop], default="Tableau de bord", key=f"nav_{coop}",
                                    label_visibility="collapsed") or "Tableau de bord"
with h4:
    offline_button()
with h5:
    st.html(THEME_TOGGLE, unsafe_allow_javascript=True)
with h3:
    pop = st.popover("⚙ Paramètres", width="stretch")

if coop == "UPAS":
    render_upas(page, pop)
    st.stop()

df_all, raw, source, meta = load()
# Fiches d'un même producteur + parcelle fusionnées en une seule (format Kobo) : base de tous les indicateurs.
# Les fiches telles que saisies (df_saisie) servent uniquement au bandeau des doublons ; le fichier d'origine est intact.
df_saisie_all = df_all
raw_fus, fusion = (raw, {"groupes": 0, "fiches": 0, "en_moins": 0}) if raw is df_all else merge_forms(raw)
if fusion["groupes"]:
    df_all = normalize(raw_fus)
watch_updates(tuple(COOP_SLOTS["ECAM"]), "data_version_ECAM")

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

def _filtre(d: pd.DataFrame) -> pd.DataFrame:
    if f_sec:
        d = d[d["section"].isin(f_sec)]
    if f_coach:
        d = d[d["agent_id"].isin(f_coach)]
    if isinstance(f_dates, (tuple, list)) and len(f_dates) == 2:
        d = d[d["date"].isna() | d["date"].dt.date.between(f_dates[0], f_dates[1])]
    return d


df = _filtre(df_all.copy())               # fiches fusionnées → indicateurs
df_saisie = _filtre(df_saisie_all.copy())  # fiches telles que saisies → bandeau des doublons

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


# Indicateurs : 1 ligne par code producteur (pp). « parc » (1 ligne par parcelle) sert seulement à l'export terrain.
parc = consolidate(df, ["prod_code", "plantation"])
pp = consolidate(parc[parc["prod_code"].notna()], ["prod_code"])  # 1 ligne par code producteur
r_eng = pp["engagement"].astype(float).mean()
r_cssvd = pp["cssvd"].astype(float).mean()
r_doc = ((pp["doc_promesse"] == 1) | (pp["doc_fiche"] == 1)).mean()
r_gps = pp["lat"].notna().mean()


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
    <div class='welcome'>Coopérative <span style='color:var(--accent)'>ECAM</span></div>
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
    <div class='stats' style='display:flex; justify-content:flex-end; gap:34px; padding-bottom:4px'>
      <div><div class='stat'><span class='ic'>📝</span><span class='num'>{n_fiches}</span></div><div class='lab'>Fiches</div></div>
      <div><div class='stat'><span class='ic'>🧑‍🌾</span><span class='num'>{n_coach}</span></div><div class='lab'>Coachs</div></div>
      <div><div class='stat'><span class='ic'>🌳</span><span class='num'>{fmt(hectares)}</span></div><div class='lab'>Hectares</div></div>
    </div>""")

# --- Compteurs « x / objectif » --------------------------------------------
def n_prod_where(mask: pd.Series) -> int:
    return int(df.loc[mask.fillna(False).astype(bool), "prod_code"].nunique())


cible = objectif or n_prod
# Tout est compté par producteur unique, pour que Enquêtés = Refus + Engagés.
# Un producteur saisi sur plusieurs fiches est « engagé » si au moins une fiche dit Oui.
n_engages = n_prod_where(df["engagement"] == True)  # noqa: E712
engages = set(df.loc[df["engagement"] == True, "prod_code"].dropna())  # noqa: E712
n_refus = int(df.loc[(df["engagement"] == False) & ~df["prod_code"].isin(engages), "prod_code"].nunique())  # noqa: E712
n_saisies = len(df_saisie)
n_extra = n_saisies - n_fiches  # fiches en trop, fusionnées
n_fiche, n_promesse = n_prod_where(df["doc_fiche"] == 1), n_prod_where(df["doc_promesse"] == 1)
n_signes = n_prod_where(df["cssvd"] == True)  # noqa: E712


def _r(a, b):
    return a / b * 100 if b else 0


html(kpi_cards([
    {"lab": "Producteurs enquêtés", "ico": "👥", "v": n_prod, "p": _r(n_prod, cible),
     "sub": f"{fmt(max(cible - n_prod, 0))} restants · {n_saisies} fiches saisies"
            + (f" · <b>{n_extra} en double</b>" if n_extra > 0 else "")},
    {"lab": "Refus d'engagement", "ico": "✋", "v": n_refus, "p": _r(n_refus, n_prod), "cls": "red",
     "sub": f"{_r(n_refus, n_prod):.0f} % des enquêtés"},
    {"lab": "Engagement du ménage", "ico": "🤝", "v": n_engages, "p": _r(n_engages, n_prod),
     "sub": f"{_r(n_engages, n_prod):.0f} % des enquêtés"},
    {"lab": "Inspection CSSVD", "ico": "🔍", "v": n_fiche, "p": _r(n_fiche, n_prod),
     "sub": f"{_r(n_fiche, n_prod):.0f} % des enquêtés · {n_signes} avec signes de swollen shoot"},
    {"lab": "Promesse de taille", "ico": "✂️", "v": n_promesse, "p": _r(n_promesse, n_prod), "cls": "amber",
     "sub": f"{_r(n_promesse, n_prod):.0f} % des enquêtés"},
], cible))


def doublons_table(d: pd.DataFrame) -> pd.DataFrame:
    """Producteurs saisis sur plusieurs fiches, avec ce que contient chaque fiche."""
    dup = d[d["prod_code"].notna() & d["prod_code"].duplicated(keep=False)].sort_values(["prod_code", "date"])

    def contenu(r) -> str:
        docs = [n for n, ok in (("fiche CSSVD", r.doc_fiche == 1), ("promesse de taille", r.doc_promesse == 1)) if ok]
        return " + ".join(docs) if docs else "aucun document"

    rows = []
    for code, g in dup.groupby("prod_code", sort=True):
        rows.append({
            "Code planteur": code, "Nom producteur": g["prod_nom"].iloc[0],
            "Coach(s)": ", ".join(sorted(g["agent_id"].dropna().unique())),
            "Nom du coach": ", ".join(sorted(g["agent_nom"].dropna().unique())),
            "Parcelle(s)": ", ".join(sorted(g["plantation"].dropna().unique())),
            "Nb fiches": len(g),
            "Envoyées à Kobo le": ", ".join(g["date"].dt.strftime("%d/%m %H:%M").dropna()),
            "Contenu de chaque fiche": " | ".join(f"{i}) {contenu(r)}" for i, r in enumerate(g.itertuples(), 1)),
        })
    return pd.DataFrame(rows)


if n_extra > 0:
    dbl = doublons_table(df_saisie)
    # coach(s) concerné(s) : « COACH007 KOUAKOU N'GUESSAN EUGENE (3) »
    _dup = df_saisie[df_saisie["prod_code"].notna() & df_saisie["prod_code"].duplicated(keep=False)]
    _par_coach = (_dup.groupby(["agent_id", "agent_nom"], dropna=False)["prod_code"].nunique()
                  .sort_values(ascending=False))
    coachs_dbl = ", ".join(
        f"{cid if pd.notna(cid) else '?'} {cnom if pd.notna(cnom) else ''}".strip() + (f" ({n})" if len(_par_coach) > 1 else "")
        for (cid, cnom), n in _par_coach.items())
    with st.container(key="card_doublons"):
        with st.expander(f"⚠ {len(dbl)} producteur(s) saisi(s) sur plusieurs fiches — {n_extra} fiche(s) en trop "
                         f"(comptés une seule fois dans les compteurs) · coach : {coachs_dbl}"):
            st.caption("Cause habituelle : le coach a rempli un formulaire séparé pour chaque document "
                       "(fiche CSSVD, promesse de taille…) au lieu d'un seul formulaire par producteur. "
                       "Les informations des différentes fiches sont regroupées pour ces producteurs. "
                       "À corriger dans Kobo (supprimer ou fusionner les fiches en trop).")
            st.dataframe(dbl, width="stretch", hide_index=True)
            st.download_button("⬇ Exporter la liste des doublons (Excel)", to_xlsx(dbl, "Doublons"),
                               file_name=f"ecam_doublons_{pd.Timestamp.now():%Y%m%d}.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

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
        n_oui = int((pp["engagement"] == True).sum())  # noqa: E712
        n_non = int((pp["engagement"] == False).sum())  # noqa: E712
        with a3, st.container(key="card_ring"):
            html("<span class='arrow-btn'>↗</span><p class='card-title'>Engagement</p>")
            st.plotly_chart(ring_chart(r_eng, "Producteurs engagés"), config=CFG, width="stretch")
            html(f"""<div style='display:flex; gap:8px; justify-content:center'>
            <span class='cal-chip'>✓ Oui · {n_oui}</span><span class='cal-chip'>✕ Non · {n_non}</span>
            <span class='cal-chip' style='background:var(--hero);color:#fff;border-color:transparent'>{n_prod} producteurs</span></div>""")

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
            "Promesse": (pp["doc_promesse"] == 1).mean(),
            "Fiche CSSVD": (pp["doc_fiche"] == 1).mean(),
            "Aucun": (pp["doc_aucun"] == 1).mean(),
        }
        seg_cols = [GREEN, DARK, GREY]
        seg = "".join(
            f"<div style='flex:{max(v, .08):.3f}'><span class='muted' style='color:var(--ink)'>{pct(v)}</span>"
            f"<div class='bar' style='background:{c}'>{k if v > .25 else ''}</div></div>"
            for (k, v), c in zip(docs.items(), seg_cols))
        with st.container(key="card_docs"):
            html(f"""<div style='display:flex; justify-content:space-between; align-items:center'>
            <p class='card-title'>Documents</p><span class='big' style='font-size:2rem'>{pct(r_doc)}</span></div>
            <div class='seg'>{seg}</div>""")

        cs = pp[pp["cssvd"] == True]  # noqa: E712
        n_obs = sum(int((cs[k] == 1).sum()) > 0 for k in SYMPTOMES)
        with st.container(key="dark_symptoms"):
            html(f"""<div style='display:flex; justify-content:space-between; align-items:center; margin-bottom:6px'>
            <span style='font-size:1.25rem'>Symptômes CSSVD</span><span style='font-size:2rem;font-weight:300'>{n_obs}/6</span></div>
            <div class='t2' style='color:#a8a8a8;font-size:.78rem;margin-bottom:6px'>{len(cs)} producteurs avec signes ·
            cliquez sur un symptôme pour voir les photos</div>""")
            symptom_buttons(cs, "dash")
            st.download_button("⬇  Exporter la liste (Excel)", to_xlsx(cssvd_table(parc)), key="dash_export",
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
        html("<p class='card-title'>🧾 Liste des producteurs</p>"
             "<div class='muted'>1 ligne = 1 code producteur (fiches en double fusionnées)</div>")
        t = ecam_producteurs_table(pp, df)
        e1, e2 = st.columns([1, 2])
        fp_sec = e1.multiselect("Section", sorted(t["Section"].dropna().unique()), key="prod_sec")
        q = e2.text_input("Rechercher un producteur", placeholder="Nom, code planteur, localité, coach…")
        if fp_sec:
            t = t[t["Section"].isin(fp_sec)]
        if q:
            m = t.astype(str).apply(lambda col: col.str.contains(q, case=False, regex=False)).any(axis=1)
            t = t[m]
        st.dataframe(t, width="stretch", hide_index=True, height=420, column_config={
            "Superficie totale (ha)": st.column_config.NumberColumn(format="%.2f"),
            "Année naissance": st.column_config.NumberColumn(format="%d"),
        })
        stamp = pd.Timestamp.now().strftime("%Y%m%d")
        d1, d3, d2, _ = st.columns([1, 1.3, 1, 1])
        d1.download_button("⬇ Exporter Excel", to_xlsx(t, "Producteurs ECAM"),
                           file_name=f"ecam_producteurs_{stamp}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           type="primary", width="stretch", disabled=t.empty, key="prod_xlsx")
        d3.download_button("⬇ Excel par section", ecam_prod_xlsx_by_section(t) if not t.empty else b"",
                           file_name=f"ecam_producteurs_par_section_{stamp}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           type="primary", width="stretch", disabled=t.empty, key="prod_xlsx_sec",
                           help="Une feuille « Résumé » + une feuille par section")
        d2.download_button("⬇ Exporter CSV", t.to_csv(index=False, sep=";").encode("utf-8-sig"),
                           file_name=f"ecam_producteurs_{stamp}.csv", mime="text/csv",
                           width="stretch", disabled=t.empty, key="prod_csv")

# ============================================================================
# PAGE : Coachs
# ============================================================================
elif page == "Coachs":
    activite = df.groupby(["agent_id", "agent_nom"]).agg(fiches=("prod_code", "size"), derniere=("date", "max"))
    taux = (pp.groupby(["agent_id", "agent_nom"])
            .agg(producteurs=("prod_code", "nunique"),
                 engagement=("engagement", lambda s: s.astype(float).mean()),
                 cssvd=("cssvd", lambda s: s.astype(float).mean()),
                 promesse=("doc_promesse", lambda s: (s == 1).mean()),
                 gps=("lat", lambda s: s.notna().mean())))
    perf = (activite.join(taux, how="outer").reset_index()
            [["agent_id", "agent_nom", "fiches", "producteurs", "engagement", "cssvd", "promesse", "gps", "derniere"]]
            .sort_values("fiches", ascending=False))
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
            "promesse": st.column_config.ProgressColumn("Promesse de taille", format="percent", min_value=0, max_value=1),
            "gps": st.column_config.ProgressColumn("GPS renseigné", format="percent", min_value=0, max_value=1),
            "derniere": st.column_config.DatetimeColumn("Dernière fiche", format="DD/MM HH:mm"),
        })

# ============================================================================
# PAGE : Symptômes
# ============================================================================
elif page == "Symptômes":
    cs = pp[pp["cssvd"] == True]  # noqa: E712
    c1, c2 = st.columns([1, 1.3], gap="medium")
    with c1, st.container(key="card_s_ring"):
        html("<p class='card-title'>Producteurs avec signes de CSSVD</p>")
        st.plotly_chart(ring_chart(r_cssvd, f"{len(cs)} / {pp['cssvd'].notna().sum()} producteurs"),
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
        tab = cssvd_table(parc)
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
        d1, d3, d2, _ = st.columns([1, 1.3, 1, 1])
        d1.download_button("⬇ Exporter Excel", to_xlsx(tab), file_name=f"producteurs_swollen_shoot_{stamp}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           type="primary", width="stretch", disabled=tab.empty)
        d3.download_button("⬇ Excel par section", ecam_xlsx_by_section(tab) if not tab.empty else b"",
                           file_name=f"producteurs_swollen_shoot_par_section_{stamp}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           type="primary", width="stretch", disabled=tab.empty,
                           help="Une feuille « Résumé » + une feuille par section")
        d2.download_button("⬇ Exporter CSV", tab.to_csv(index=False, sep=";").encode("utf-8-sig"),
                           file_name=f"producteurs_swollen_shoot_{stamp}.csv", mime="text/csv",
                           width="stretch", disabled=tab.empty)

    with st.container(key="card_s_sec"):
        html("<p class='card-title'>Taux de CSSVD par section</p>")
        t = (pp.groupby("section").agg(producteurs=("cssvd", "size"),
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
    g = pp.dropna(subset=["lat", "lon"]).copy()
    with st.container(key="card_map"):
        html(f"<p class='card-title'>Producteurs géolocalisés · {len(g)} / {len(pp)}</p>")
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
        vues = ["Normalisée", "Fiches fusionnées (format Kobo)", "Export d'origine"]
        view = st.radio("Vue", vues if fusion["groupes"] else [vues[0], vues[2]], horizontal=True,
                        label_visibility="collapsed")
        if view == vues[0]:
            data = df
        elif view == vues[1]:
            data = raw_fus.loc[raw_fus.index.intersection(df.index)]
            st.caption(f"🔗 {fusion['fiches']} fiches de {fusion['groupes']} producteur(s) fusionnées en "
                       f"{fusion['groupes']} fiche(s) — colonne « _fusion ». C'est la base des indicateurs ; "
                       "le fichier Kobo d'origine n'est pas modifié.")
        else:
            data = raw.loc[raw.index.intersection(df_saisie.index)]
        st.dataframe(data, width="stretch", height=560)
        d1, d2 = st.columns(2)
        d1.download_button("⬇ Télécharger (CSV)", data.to_csv(index=False, sep=";").encode("utf-8-sig"),
                           file_name="cssvd_export.csv", mime="text/csv", width="stretch")
        if view != vues[0]:
            d2.download_button("⬇ Télécharger (Excel)", df_to_xlsx_bytes(data),
                               file_name="ecam_fiches_fusionnees.xlsx" if view == vues[1] else "ecam_export_origine.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", width="stretch")
