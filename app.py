from __future__ import annotations

from pathlib import Path
from datetime import date, datetime
from io import BytesIO, StringIO
import re
import hashlib
import csv
import math

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

try:
    import matplotlib.pyplot as plt
    MATPLOTLIB_AVAILABLE = True
except Exception:
    plt = None
    MATPLOTLIB_AVAILABLE = False

try:
    from reportlab.lib import colors
    from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_LEFT
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import inch
    from reportlab.platypus import (
        BaseDocTemplate,
        Frame,
        PageTemplate,
        Paragraph,
        Spacer,
        Table,
        TableStyle,
        PageBreak,
        Image,
    )
    REPORTLAB_AVAILABLE = True
except Exception:
    colors = None
    getSampleStyleSheet = None
    ParagraphStyle = None
    A4 = None
    landscape = None
    inch = None
    BaseDocTemplate = None
    Frame = None
    PageTemplate = None
    Paragraph = None
    Spacer = None
    Table = None
    TableStyle = None
    PageBreak = None
    Image = None
    TA_CENTER = TA_JUSTIFY = TA_LEFT = None
    REPORTLAB_AVAILABLE = False


st.set_page_config(
    page_title="Records List Intelligence Dashboard",
    page_icon="🧬",
    layout="wide",
    initial_sidebar_state="expanded",
)

CUSTOM_HEADERS = [
    "Distributor name",
    "Instrument type",
    "Installation date",
    "Customer name",
    "In Blood Bank",
    "Address",
    "ZipCode",
    "City",
    "Country",
    "World Region",
    "Commercial Region",
    "Latitude",
    "Longitude",
    "Product Line",
    "Serial number",
    "Machine Configurations",
    "Asset condition",
    "PM plan",
    "Number of tests per day",
    "Operational status",
    "Type of contract",
    "Contract duration",
    "Tag",
    "Notes",
    "PM last date",
    "PM frequency",
    "PM next date",
    "PM performed On",
    "CLIA - Adrenal function",
    "CLIA - Autoimmunity",
    "CLIA - Bone turnover",
    "CLIA - Cardiac Markers",
    "CLIA - Diabetes",
    "CLIA - EBV",
    "CLIA - Fertility",
    "CLIA - Gastroenterology",
    "CLIA - Growth",
    "CLIA - Hematology",
    "CLIA - Hepatitis and Retrovirus",
    "CLIA - Hypertension",
    "CLIA - Infectious diseases",
    "CLIA - PTH",
    "CLIA - Sepsis",
    "CLIA - Thrombosis",
    "CLIA - Thyroid",
    "CLIA - Torch",
    "CLIA - Tumor Markers",
    "CLIA - Vitamin D",
    "ELISA - Autoimmunity",
    "ELISA - Hepatitis",
    "ELISA - Infection Diseases",
    "ELISA - Murex",
    "MOLECULAR ASR",
    "MOLECULAR DAD - Simplexa C Diff Direct kit",
    "MOLECULAR DAD - Simplexa Flu A/B &RSV Direct kit",
    "MOLECULAR DAD - Simplexa Group A Strep Direct kit",
    "MOLECULAR DAD - Simplexa HSV1&2 Direct kit",
    "MOLECULAR UD - Simplexa BKV kit",
    "MOLECULAR UD - Simplexa Bordetella Universal Direct",
    "MOLECULAR UD - Simplexa C Diff Universal Direct",
    "MOLECULAR UD - Simplexa CMV kit",
    "MOLECULAR UD - Simplexa Dengue kit",
    "MOLECULAR UD - Simplexa EBV kit",
    "MOLECULAR UD - Simplexa Flu A/B & RSV kit",
    "MOLECULAR UD - Simplexa Influenza A N1N1 (2009) kit",
    "Other - specify in note field",
    "_blank",
]

ASSAY_COLS = CUSTOM_HEADERS[28:-1]

PLOT_TEMPLATE = "plotly_dark"
PLOT_BG = "rgba(0,0,0,0)"
GRID = "rgba(255,255,255,0.10)"
ACCENT = "#4df6ff"
ACCENT_2 = "#b15cff"
ACCENT_3 = "#00ff9d"
WARNING = "#ffb454"
DANGER = "#ff5d8f"
TEXT = "#f7fbff"
MUTED = "#a9b9d6"

APP_CSS = """
<style>
:root {
    --bg1: #040814;
    --bg2: #090d1d;
    --card: rgba(12,16,35,0.75);
    --border: rgba(77,246,255,0.16);
    --txt: #f7fbff;
    --muted: #9db0d3;
}
.stApp {
    background:
        radial-gradient(circle at 12% 8%, rgba(77,246,255,0.10), transparent 22%),
        radial-gradient(circle at 88% 0%, rgba(177,92,255,0.10), transparent 22%),
        radial-gradient(circle at 82% 82%, rgba(0,255,157,0.07), transparent 18%),
        linear-gradient(180deg, #040814 0%, #080c1a 45%, #050712 100%);
    color: var(--txt);
}
.block-container {
    padding-top: 1rem;
    padding-bottom: 2rem;
}
section[data-testid="stSidebar"] {
    background: linear-gradient(180deg, rgba(8,12,28,0.97) 0%, rgba(10,14,30,0.97) 100%);
    border-right: 1px solid rgba(255,255,255,0.06);
}
.hero {
    padding: 1.35rem 1.55rem;
    border-radius: 24px;
    border: 1px solid rgba(77,246,255,0.16);
    background: linear-gradient(135deg, rgba(15,24,50,0.85) 0%, rgba(8,12,28,0.92) 100%);
    box-shadow: 0 0 0 1px rgba(255,255,255,0.03), 0 20px 45px rgba(0,0,0,0.28);
    margin-bottom: 1rem;
}
.hero h1 { margin: 0; font-size: 2.15rem; letter-spacing: 0.02em; }
.hero p { margin: 0.45rem 0 0 0; color: var(--muted); }
.badge-row { display: flex; flex-wrap: wrap; gap: 0.55rem; margin-top: 0.8rem; }
.badge {
    padding: 0.3rem 0.75rem; border-radius: 999px; font-size: 0.77rem;
    background: rgba(255,255,255,0.05); border: 1px solid rgba(255,255,255,0.08);
}
.metric-shell {
    padding: 0.95rem 1rem; border-radius: 20px;
    background: linear-gradient(180deg, rgba(16,21,42,0.84), rgba(9,13,27,0.86));
    border: 1px solid rgba(255,255,255,0.08); min-height: 115px;
}
.metric-label { font-size: 0.82rem; color: var(--muted); text-transform: uppercase; letter-spacing: 0.08em; }
.metric-value { font-size: 1.9rem; font-weight: 700; margin-top: 0.2rem; }
.metric-sub { margin-top: 0.2rem; font-size: 0.85rem; color: #c3d0ea; }
.stTabs [data-baseweb="tab-list"] { gap: 0.45rem; }
.stTabs [data-baseweb="tab"] { border-radius: 12px; background: rgba(255,255,255,0.04); padding: 0.45rem 0.85rem; }
.stTabs [aria-selected="true"] {
    background: rgba(77,246,255,0.12) !important;
    border: 1px solid rgba(77,246,255,0.30) !important;
}
[data-testid="stDataFrame"] {
    border-radius: 18px; overflow: hidden; border: 1px solid rgba(255,255,255,0.08);
}
.small-note { color: var(--muted); font-size: 0.88rem; }
</style>
"""
st.markdown(APP_CSS, unsafe_allow_html=True)


# NOTE: Full code preserved in downloadable file because it is very long for chat.
# Download the exact copy-paste file here:
# sandbox:/mnt/data/app_dashboard_mejorado_v9.py
# If you want the entire body pasted here in chat, ask me to continue in numbered parts.
