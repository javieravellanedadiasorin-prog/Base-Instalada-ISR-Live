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
    from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
    from reportlab.lib.units import inch
    from reportlab.platypus import PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
    REPORTLAB_AVAILABLE = True
except Exception:
    colors = None
    TA_CENTER = TA_JUSTIFY = TA_LEFT = None
    A4 = landscape = None
    ParagraphStyle = getSampleStyleSheet = None
    inch = 72
    PageBreak = Paragraph = SimpleDocTemplate = Spacer = Table = TableStyle = None
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
PLOT_BG = "rgba(12, 19, 30, 0.02)"
GRID = "rgba(170, 224, 255, 0.10)"
ACCENT = "#56d8ff"
ACCENT_2 = "#8fa8ff"
ACCENT_3 = "#59f0d0"
WARNING = "#ffb454"
DANGER = "#ff5d8f"
TEXT = "#f8fcff"
MUTED = "rgba(238,245,255,0.92)"

APP_CSS = """
<style>
:root {
    --bg-1: #050912;
    --bg-2: #08111b;
    --bg-3: #0a1622;
    --panel: rgba(28, 42, 64, 0.34);
    --panel-strong: rgba(34, 50, 76, 0.46);
    --panel-soft: rgba(255, 255, 255, 0.06);
    --glass-white: rgba(255,255,255,0.12);
    --stroke: rgba(205, 232, 255, 0.22);
    --stroke-2: rgba(255, 255, 255, 0.10);
    --txt: #f8fcff;
    --txt-strong: #ffffff;
    --txt-soft: rgba(241, 248, 255, 0.96);
    --muted: rgba(214, 228, 245, 0.82);
    --cyan: #71e1ff;
    --cyan-2: #35c8ff;
    --cyan-3: #8de8ff;
    --blue: #8ea9ff;
    --mint: #58efd1;
    --amber: #ffbe57;
    --danger: #ff6a8c;
    --shadow: rgba(0, 0, 0, 0.36);
}

html, body, [class*="css"] {
    font-family: "Inter", "Segoe UI", sans-serif;
}

.stApp {
    background:
        radial-gradient(circle at 12% 18%, rgba(133, 214, 255, 0.16), transparent 16%),
        radial-gradient(circle at 78% 10%, rgba(255,255,255,0.11), transparent 16%),
        radial-gradient(circle at 80% 78%, rgba(103, 229, 255, 0.08), transparent 18%),
        radial-gradient(circle at 30% 82%, rgba(122, 148, 255, 0.10), transparent 20%),
        linear-gradient(180deg, #040912 0%, #07101a 18%, #0a1520 44%, #08111a 72%, #060b12 100%);
    background-attachment: fixed;
    color: var(--txt);
}

.stApp::before {
    content: "";
    position: fixed;
    inset: 0;
    pointer-events: none;
    background:
        linear-gradient(180deg, rgba(255,255,255,0.025), transparent 18%, transparent 82%, rgba(255,255,255,0.02)),
        radial-gradient(circle at 50% 0%, rgba(255,255,255,0.08), transparent 18%),
        radial-gradient(circle at 55% 45%, rgba(113,225,255,0.05), transparent 26%);
    z-index: 0;
}

.block-container {
    position: relative;
    z-index: 1;
    padding-top: 0.85rem;
    padding-bottom: 2rem;
    max-width: 98rem;
}

section[data-testid="stSidebar"] {
    background:
        linear-gradient(180deg, rgba(28, 42, 64, 0.28) 0%, rgba(20, 32, 50, 0.36) 100%);
    border-right: 1px solid rgba(255,255,255,0.08);
    backdrop-filter: blur(22px);
    -webkit-backdrop-filter: blur(22px);
    box-shadow:
        inset -1px 0 0 rgba(255,255,255,0.05),
        10px 0 28px rgba(0, 0, 0, 0.18);
}

section[data-testid="stSidebar"] > div {
    background: transparent;
}

section[data-testid="stSidebar"] .stMarkdown,
section[data-testid="stSidebar"] p,
section[data-testid="stSidebar"] label,
section[data-testid="stSidebar"] span,
section[data-testid="stSidebar"] div {
    color: var(--txt-soft) !important;
    text-shadow: 0 0 6px rgba(74, 203, 255, 0.08);
}

.hero {
    position: relative;
    overflow: hidden;
    padding: 1.2rem 1.45rem;
    border-radius: 28px;
    border: 1px solid rgba(255,255,255,0.16);
    background:
        linear-gradient(180deg, rgba(198, 223, 255, 0.13) 0%, rgba(72, 101, 145, 0.10) 18%, rgba(25, 38, 58, 0.18) 100%);
    backdrop-filter: blur(24px);
    -webkit-backdrop-filter: blur(24px);
    box-shadow:
        0 16px 40px rgba(0, 0, 0, 0.24),
        inset 0 1px 0 rgba(255,255,255,0.22),
        inset 0 -1px 0 rgba(255,255,255,0.03);
    margin-bottom: 1rem;
}

.hero::before {
    content: "";
    position: absolute;
    inset: 0;
    pointer-events: none;
    background:
        linear-gradient(180deg, rgba(255,255,255,0.10), transparent 28%),
        radial-gradient(circle at 18% 12%, rgba(255,255,255,0.18), transparent 16%),
        radial-gradient(circle at 72% 12%, rgba(113,225,255,0.08), transparent 20%);
}

.hero h1 {
    margin: 0;
    font-size: 2.2rem;
    line-height: 1.05;
    letter-spacing: 0.02em;
    color: #f9fdff !important;
    text-shadow:
        0 0 6px rgba(146, 235, 255, 0.34),
        0 0 18px rgba(58, 199, 255, 0.18);
}

.hero p {
    margin: 0.45rem 0 0 0;
    color: rgba(239,247,255,0.88) !important;
    font-size: 1rem;
    text-shadow: 0 0 8px rgba(83, 205, 255, 0.12);
}

.badge-row {
    display: flex;
    flex-wrap: wrap;
    gap: 0.55rem;
    margin-top: 0.85rem;
}

.badge {
    padding: 0.35rem 0.75rem;
    border-radius: 999px;
    font-size: 0.77rem;
    color: #f8fcff;
    background: rgba(255,255,255,0.08);
    border: 1px solid rgba(255,255,255,0.14);
    box-shadow:
        inset 0 1px 0 rgba(255,255,255,0.14),
        0 0 12px rgba(83,205,255,0.08);
    backdrop-filter: blur(16px);
    -webkit-backdrop-filter: blur(16px);
}

.metric-shell,
div[data-testid="stMetric"] {
    position: relative;
    overflow: hidden;
    border-radius: 22px !important;
    background:
        linear-gradient(180deg, rgba(181, 214, 255, 0.12) 0%, rgba(25, 38, 58, 0.22) 20%, rgba(18, 28, 44, 0.28) 100%) !important;
    border: 1px solid rgba(255,255,255,0.14) !important;
    backdrop-filter: blur(22px);
    -webkit-backdrop-filter: blur(22px);
    box-shadow:
        0 12px 28px rgba(0,0,0,0.22),
        inset 0 1px 0 rgba(255,255,255,0.18);
}

.metric-shell {
    padding: 0.95rem 1rem;
    min-height: 118px;
}

.metric-shell::before,
div[data-testid="stMetric"]::before {
    content: "";
    position: absolute;
    inset: 0;
    pointer-events: none;
    background:
        linear-gradient(180deg, rgba(255,255,255,0.08), transparent 24%),
        radial-gradient(circle at 18% 0%, rgba(255,255,255,0.12), transparent 18%);
}

.metric-label {
    font-size: 0.78rem;
    text-transform: uppercase;
    letter-spacing: 0.1em;
    color: rgba(229,243,255,0.78) !important;
    text-shadow: 0 0 8px rgba(77, 204, 255, 0.10);
}

.metric-value {
    font-size: 1.95rem;
    font-weight: 700;
    margin-top: 0.22rem;
    color: #ffffff !important;
    text-shadow:
        0 0 8px rgba(115,228,255,0.24),
        0 0 18px rgba(53,200,255,0.14);
}

.metric-sub {
    margin-top: 0.2rem;
    font-size: 0.86rem;
    color: rgba(235,245,255,0.88) !important;
}

.stTabs [data-baseweb="tab-list"] {
    gap: 0.55rem;
    padding: 0.28rem;
    background: rgba(255,255,255,0.04);
    border-radius: 18px;
    border: 1px solid rgba(255,255,255,0.10);
    backdrop-filter: blur(14px);
    -webkit-backdrop-filter: blur(14px);
}

.stTabs [data-baseweb="tab"] {
    border-radius: 14px;
    padding: 0.52rem 0.92rem;
    background: rgba(255,255,255,0.03);
    color: rgba(238,247,255,0.88);
    border: 1px solid transparent;
    transition: all 0.2s ease;
    text-shadow: 0 0 8px rgba(53,200,255,0.10);
}

.stTabs [data-baseweb="tab"]:hover {
    background: rgba(255,255,255,0.07);
    border-color: rgba(255,255,255,0.10);
}

.stTabs [aria-selected="true"] {
    background:
        linear-gradient(180deg, rgba(129, 212, 255, 0.18), rgba(54, 93, 150, 0.14)) !important;
    border: 1px solid rgba(119, 221, 255, 0.36) !important;
    box-shadow:
        inset 0 1px 0 rgba(255,255,255,0.18),
        0 0 18px rgba(53,200,255,0.16) !important;
    color: #ffffff !important;
    text-shadow:
        0 0 8px rgba(136,234,255,0.34),
        0 0 16px rgba(53,200,255,0.18);
}

div[data-testid="stPlotlyChart"],
div[data-testid="stDataFrame"],
div[data-testid="stTable"],
div[data-testid="stExpander"],
div[data-testid="stForm"] {
    border-radius: 26px;
    overflow: hidden;
    background:
        linear-gradient(180deg, rgba(193, 221, 255, 0.10) 0%, rgba(24, 36, 54, 0.18) 18%, rgba(17, 27, 42, 0.24) 100%);
    border: 1px solid rgba(255,255,255,0.12);
    backdrop-filter: blur(20px);
    -webkit-backdrop-filter: blur(20px);
    box-shadow:
        0 12px 30px rgba(0,0,0,0.20),
        inset 0 1px 0 rgba(255,255,255,0.12);
}

div[data-testid="stPlotlyChart"] {
    padding: 0.24rem;
}

div[data-testid="stDataFrame"] table,
div[data-testid="stTable"] table {
    color: #f8fcff !important;
}

.stButton > button,
.stDownloadButton > button {
    border-radius: 16px;
    border: 1px solid rgba(137, 228, 255, 0.34);
    background:
        linear-gradient(180deg, rgba(114, 214, 255, 0.20), rgba(52, 94, 145, 0.18));
    color: #ffffff !important;
    font-weight: 600;
    box-shadow:
        inset 0 1px 0 rgba(255,255,255,0.18),
        0 0 18px rgba(53,200,255,0.12);
    text-shadow: 0 0 8px rgba(110, 227, 255, 0.22);
}

.stButton > button:hover,
.stDownloadButton > button:hover {
    border-color: rgba(160, 237, 255, 0.44);
    box-shadow:
        inset 0 1px 0 rgba(255,255,255,0.22),
        0 0 22px rgba(53,200,255,0.18);
    transform: translateY(-1px);
}

.stTextInput input,
.stNumberInput input,
.stDateInput input,
.stTextArea textarea,
.stSelectbox div[data-baseweb="select"] > div,
.stMultiSelect div[data-baseweb="select"] > div {
    background: rgba(255,255,255,0.08) !important;
    border: 1px solid rgba(255,255,255,0.12) !important;
    border-radius: 15px !important;
    color: #ffffff !important;
    backdrop-filter: blur(12px);
    -webkit-backdrop-filter: blur(12px);
    box-shadow: inset 0 1px 0 rgba(255,255,255,0.08);
}

.stTextInput input::placeholder,
.stTextArea textarea::placeholder {
    color: rgba(235,245,255,0.54) !important;
}

h1, h2, h3, h4, h5, h6,
div[data-testid="stMarkdownContainer"] h1,
div[data-testid="stMarkdownContainer"] h2,
div[data-testid="stMarkdownContainer"] h3,
div[data-testid="stMarkdownContainer"] h4 {
    color: #fbfdff !important;
    letter-spacing: 0.01em;
    text-shadow:
        0 0 8px rgba(140, 235, 255, 0.30),
        0 0 20px rgba(53,200,255,0.14);
}

h2 {
    font-size: 1.9rem !important;
    font-weight: 700 !important;
}

h3 {
    font-size: 1.35rem !important;
    font-weight: 700 !important;
}

p, label, span, div {
    color: var(--txt-soft);
}

div[data-testid="stCaptionContainer"],
.small-note,
.stCaption,
small {
    color: rgba(231,243,255,0.84) !important;
    text-shadow: 0 0 8px rgba(53,200,255,0.08);
}

hr {
    border: none;
    height: 1px;
    background: linear-gradient(90deg, transparent, rgba(255,255,255,0.18), transparent);
}

[data-testid="stSidebar"] .stFileUploader,
[data-testid="stSidebar"] .stSelectbox,
[data-testid="stSidebar"] .stMultiSelect {
    background: rgba(255,255,255,0.04);
    border-radius: 18px;
}

[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] h1,
[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] h2,
[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] h3 {
    color: #fafdff !important;
    text-shadow:
        0 0 8px rgba(140,235,255,0.30),
        0 0 18px rgba(53,200,255,0.16);
}
</style>
"""
st.markdown(APP_CSS, unsafe_allow_html=True)


def glow_layout(fig: go.Figure, height: int = 420, title_size: int = 18) -> go.Figure:
    fig.update_layout(
        template=PLOT_TEMPLATE,
        height=height,
        paper_bgcolor=PLOT_BG,
        plot_bgcolor=PLOT_BG,
        margin=dict(l=18, r=18, t=78, b=18),
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="right",
            x=1,
            bgcolor="rgba(16, 28, 46, 0.42)",
            bordercolor="rgba(124,221,255,0.26)",
            borderwidth=1,
            font=dict(color="#f8fbff", size=12),
        ),
        font=dict(color=TEXT),
        title_font=dict(size=title_size, color=TEXT),
        title=dict(x=0.02, xanchor="left"),
        hovermode="closest",
        hoverlabel=dict(
            bgcolor="rgba(13, 24, 42, 0.96)",
            bordercolor="rgba(255,255,255,0.22)",
            font=dict(color=TEXT),
        ),
    )
    fig.update_xaxes(
        showgrid=True,
        gridcolor=GRID,
        zeroline=False,
        automargin=True,
        linecolor="rgba(255,255,255,0.12)",
        tickfont=dict(color=TEXT),
        title_font=dict(color=TEXT),
    )
    fig.update_yaxes(
        showgrid=True,
        gridcolor=GRID,
        zeroline=False,
        automargin=True,
        linecolor="rgba(255,255,255,0.12)",
        tickfont=dict(color=TEXT),
        title_font=dict(color=TEXT),
    )
    return fig



def compress_value_distribution(series: pd.Series, max_slices: int = 4) -> pd.DataFrame:
    work = (
        series.fillna("No informado")
        .astype(str)
        .str.strip()
        .replace("", "No informado")
        .value_counts()
        .reset_index()
    )
    if work.empty:
        return pd.DataFrame(columns=["Label", "Count"])
    work.columns = ["Label", "Count"]
    if len(work) > max_slices:
        top = work.head(max_slices).copy()
        other_count = int(work.iloc[max_slices:]["Count"].sum())
        if other_count > 0:
            top = pd.concat([top, pd.DataFrame([{"Label": "Otros", "Count": other_count}])], ignore_index=True)
        work = top
    return work


def build_config_donut(field_name: str, series: pd.Series, total_assets: int) -> go.Figure:
    dist = compress_value_distribution(series, max_slices=4)
    fig = go.Figure()

    if dist.empty:
        fig.add_annotation(
            text="Sin datos",
            x=0.5,
            y=0.5,
            xref="paper",
            yref="paper",
            showarrow=False,
            font=dict(size=15, color=TEXT),
        )
        fig.update_layout(title=f"{field_name}")
        return glow_layout(fig, 340, 15)

    palette = [ACCENT, ACCENT_2, ACCENT_3, WARNING, "rgba(255,255,255,0.32)"]
    fig.add_trace(
        go.Pie(
            labels=dist["Label"],
            values=dist["Count"],
            hole=0.68,
            sort=False,
            marker=dict(colors=palette[:len(dist)], line=dict(color="rgba(255,255,255,0.18)", width=1.2)),
            textinfo="percent",
            textfont=dict(color="#ffffff", size=13),
            hovertemplate="Valor: %{label}<br>Equipos: %{value}<br>Participación: %{percent}<extra></extra>",
        )
    )
    fig.add_annotation(
        text=f"<b>{total_assets:,}</b><br><span style='font-size:11px'>equipos</span>",
        x=0.5,
        y=0.5,
        xref="paper",
        yref="paper",
        showarrow=False,
        font=dict(color='#ffffff', size=18),
    )
    fig.update_layout(
        title=field_name,
        showlegend=True,
        legend=dict(orientation="h", yanchor="bottom", y=-0.18, xanchor="center", x=0.5, bgcolor="rgba(14,26,42,0.36)", bordercolor="rgba(124,221,255,0.22)", borderwidth=1, font=dict(color="#f8fbff", size=11)),
    )
    return glow_layout(fig, 340, 15)


def compute_mapbox_center_zoom(df: pd.DataFrame, lat_col: str = "Latitude", lon_col: str = "Longitude") -> tuple[dict, float]:
    geo = df.dropna(subset=[lat_col, lon_col]).copy()
    if geo.empty:
        return {"lat": 0.0, "lon": 0.0}, 1.0

    lats = pd.to_numeric(geo[lat_col], errors="coerce").dropna()
    lons = pd.to_numeric(geo[lon_col], errors="coerce").dropna()
    if lats.empty or lons.empty:
        return {"lat": 0.0, "lon": 0.0}, 1.0

    min_lat, max_lat = float(lats.min()), float(lats.max())
    min_lon, max_lon = float(lons.min()), float(lons.max())
    center = {"lat": (min_lat + max_lat) / 2, "lon": (min_lon + max_lon) / 2}

    lat_span = max(max_lat - min_lat, 0.01)
    lon_span = max(max_lon - min_lon, 0.01)
    max_span = max(lat_span, lon_span)

    if len(geo) == 1:
        zoom = 9.5
    elif max_span <= 0.05:
        zoom = 9.0
    elif max_span <= 0.12:
        zoom = 8.2
    elif max_span <= 0.25:
        zoom = 7.2
    elif max_span <= 0.5:
        zoom = 6.3
    elif max_span <= 1.0:
        zoom = 5.5
    elif max_span <= 2.0:
        zoom = 4.7
    elif max_span <= 4.0:
        zoom = 4.0
    elif max_span <= 8.0:
        zoom = 3.2
    elif max_span <= 16.0:
        zoom = 2.6
    elif max_span <= 35.0:
        zoom = 2.0
    elif max_span <= 70.0:
        zoom = 1.45
    else:
        zoom = 1.0

    return center, zoom


def dataframe_to_excel_bytes(sheet_map: dict[str, pd.DataFrame]) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sheet_name, df in sheet_map.items():
            safe_name = re.sub(r"[\\/*?:\[\]]", "_", str(sheet_name))[:31] or "Sheet1"
            clean_df = df.copy()
            clean_df.to_excel(writer, sheet_name=safe_name, index=False)
            ws = writer.sheets[safe_name]
            ws.freeze_panes = "A2"
            for idx, col in enumerate(clean_df.columns, start=1):
                max_len = len(str(col))
                if not clean_df.empty:
                    series = clean_df[col].astype("string").fillna("")
                    series = series.str.replace("<NA>", "", regex=False).str.replace("nan", "", regex=False)
                    lengths = series.str.len().fillna(0)
                    series_max_len = int(lengths.max()) if len(lengths) else 0
                    max_len = max(max_len, series_max_len)
                ws.column_dimensions[ws.cell(row=1, column=idx).column_letter].width = min(max(max_len + 2, 12), 42)
    output.seek(0)
    return output.getvalue()



def normalize_operational_status(value) -> str:
    if pd.isna(value):
        return "No informado"
    text = str(value).strip()
    if not text:
        return "No informado"

    upper = text.upper()
    lower = text.lower()

    if "scrap" in lower:
        return "Scraped"
    if upper in {"IN ROUTINE", "ROUTINE"} or "IN ROUTINE" in upper:
        return "Routine"

    return text.title()


def compute_state_filter_counts(df: pd.DataFrame) -> list[tuple[str, int]]:
    if df.empty or "Operational status grouped" not in df.columns:
        return []

    grouped = (
        df["Operational status grouped"]
        .fillna("No informado")
        .astype(str)
        .value_counts()
    )

    items = []
    non_routine_count = int((~df["Operational status grouped"].eq("Routine")).sum())
    if non_routine_count > 0:
        items.append(("No rutina", non_routine_count))

    preferred_order = ["Routine", "Scraped", "No informado"]
    seen = set()

    for name in preferred_order:
        count = int(grouped.get(name, 0))
        if count > 0:
            items.append((name, count))
            seen.add(name)

    for name, count in grouped.items():
        if name not in seen and int(count) > 0:
            items.append((name, int(count)))

    return items


def apply_operational_status_filter(df: pd.DataFrame, selected_states: list[str]) -> pd.DataFrame:
    if df.empty or not selected_states or "Operational status grouped" not in df.columns:
        return df

    mask = pd.Series(False, index=df.index)
    state_series = df["Operational status grouped"].fillna("No informado").astype(str)

    for state in selected_states:
        if state == "No rutina":
            mask = mask | (~state_series.eq("Routine"))
        else:
            mask = mask | state_series.eq(state)

    return df[mask].copy()


def clean_filter_value(values) -> str:
    if values is None:
        return "All"
    if isinstance(values, (list, tuple, set)):
        clean = [str(v) for v in values if str(v).strip()]
        return ", ".join(clean) if clean else "All"
    text = str(values).strip()
    return text or "All"


def build_filter_summary(
    selected_regions: list[str],
    selected_countries: list[str],
    selected_distributors: list[str],
    selected_instruments: list[str],
    selected_states: list[str],
) -> dict[str, str]:
    return {
        "Commercial Region": clean_filter_value(selected_regions),
        "Country": clean_filter_value(selected_countries),
        "Distributor": clean_filter_value(selected_distributors),
        "Instrument Type": clean_filter_value(selected_instruments),
        "Operational Status": clean_filter_value(selected_states),
    }


def prepare_pdf_report_table(df: pd.DataFrame) -> pd.DataFrame:
    preferred_columns = [
        "Commercial Region",
        "Country",
        "Distributor name",
        "Customer name",
        "Instrument type",
        "Serial number",
        "Operational status grouped",
        "Operational status",
        "Operating System",
        "Asset condition",
        "Installation date",
        "Type of contract",
    ]
    available = [col for col in preferred_columns if col in df.columns]
    report_df = df[available].copy()

    for col in ["Installation date", "PM next date"]:
        if col in report_df.columns:
            report_df[col] = pd.to_datetime(report_df[col], errors="coerce").dt.strftime("%Y-%m-%d")
            report_df[col] = report_df[col].fillna("N/A")

    for col in report_df.columns:
        if col not in ["Installation date", "PM next date"]:
            report_df[col] = report_df[col].fillna("N/A").astype(str).str.strip().replace("", "N/A")

    report_df = report_df.rename(
        columns={
            "Commercial Region": "Region",
            "Distributor name": "Distributor",
            "Customer name": "Customer",
            "Instrument type": "Instrument",
            "Serial number": "Serial",
            "Operational status grouped": "State",
            "Operational status": "Raw Status",
            "Operating System": "OS",
            "Asset condition": "Asset",
            "Installation date": "Install Date",
            "Type of contract": "Contract",
        }
    )

    ordered_cols = [
        c for c in ["Region", "Country", "Distributor", "Customer", "Instrument", "Serial", "State", "Raw Status", "OS", "Asset", "Install Date", "Contract"]
        if c in report_df.columns
    ]
    return report_df[ordered_cols]


def _pdf_header_footer(canvas, doc, short_title: str):
    canvas.saveState()
    width, height = landscape(A4)

    canvas.setStrokeColor(colors.HexColor("#D9D9D9"))
    canvas.setLineWidth(0.6)
    canvas.line(doc.leftMargin, height - 34, width - doc.rightMargin, height - 34)
    canvas.line(doc.leftMargin, 28, width - doc.rightMargin, 28)

    canvas.setFont("Helvetica-Bold", 9)
    canvas.setFillColor(colors.HexColor("#222222"))
    canvas.drawString(doc.leftMargin, height - 24, short_title[:90])
    canvas.drawRightString(width - doc.rightMargin, height - 24, f"Page {doc.page}")

    canvas.setFont("Helvetica", 8)
    canvas.setFillColor(colors.HexColor("#666666"))
    canvas.drawString(doc.leftMargin, 16, "Records List Intelligence Dashboard | APA-style filtered report")
    canvas.drawRightString(width - doc.rightMargin, 16, datetime.now().strftime("%Y-%m-%d %H:%M"))
    canvas.restoreState()


def _escape_pdf_text(value) -> str:
    text = safe_text(value, "N/A")
    return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def _paragraph_cell(value, style):
    return Paragraph(_escape_pdf_text(value), style)


def _df_to_wrapped_table(df: pd.DataFrame, styles, col_widths=None, max_rows=None):
    work = df.copy()
    if max_rows is not None:
        work = work.head(max_rows)
    if work.empty:
        return Paragraph("No data available for this section.", styles["APA_Body"])

    for col in work.columns:
        work[col] = work[col].fillna("N/A").astype(str).str.slice(0, 110)

    header_row = [Paragraph(f"<b>{_escape_pdf_text(c)}</b>", styles["APA_Cell_Header"]) for c in work.columns]
    body_rows = [[_paragraph_cell(v, styles["APA_Cell"]) for v in row] for row in work.values.tolist()]
    table = Table([header_row] + body_rows, colWidths=col_widths, repeatRows=1)
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#203864")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#BFBFBF")),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F5F7FA")]),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 3),
        ("RIGHTPADDING", (0, 0), (-1, -1), 3),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
    ]))
    return table


def _summary_table_from_pairs(title: str, pairs: list[tuple[str, str]], styles):
    data = [[Paragraph("<b>Metric</b>", styles["APA_Cell_Header"]), Paragraph("<b>Value</b>", styles["APA_Cell_Header"])]]
    for k, v in pairs:
        data.append([_paragraph_cell(k, styles["APA_Cell"]), _paragraph_cell(v, styles["APA_Cell"])])
    table = Table(data, colWidths=[2.8 * inch, 2.1 * inch], repeatRows=1)
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1f4e79")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("GRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#C8D6E5")),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.whitesmoke, colors.HexColor("#F7F9FB")]),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 5),
        ("RIGHTPADDING", (0, 0), (-1, -1), 5),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]))
    return [Paragraph(title, styles["APA_Heading"]), table]


def _make_matplotlib_barh(df: pd.DataFrame, label_col: str, value_col: str, title: str, max_rows: int = 10):
    if not MATPLOTLIB_AVAILABLE or df is None or df.empty:
        return None
    work = df[[label_col, value_col]].copy().dropna().head(max_rows)
    if work.empty:
        return None
    work = work.sort_values(value_col, ascending=True)
    fig, ax = plt.subplots(figsize=(8.4, 3.8))
    ax.barh(work[label_col].astype(str), work[value_col].astype(float))
    ax.set_title(title, fontsize=11)
    ax.set_xlabel("Count")
    ax.tick_params(axis="y", labelsize=8)
    ax.tick_params(axis="x", labelsize=8)
    ax.grid(axis="x", alpha=0.25)
    fig.tight_layout()
    buf = BytesIO()
    fig.savefig(buf, format="png", dpi=180, bbox_inches="tight")
    plt.close(fig)
    buf.seek(0)
    return buf


def _make_matplotlib_scatter(df: pd.DataFrame, x_col: str, y_col: str, title: str, max_rows: int = 200):
    if not MATPLOTLIB_AVAILABLE or df is None or df.empty or x_col not in df.columns or y_col not in df.columns:
        return None
    work = df[[x_col, y_col]].copy().dropna().head(max_rows)
    if work.empty:
        return None
    work[y_col] = pd.to_numeric(work[y_col], errors="coerce")
    work = work.dropna()
    if work.empty:
        return None
    fig, ax = plt.subplots(figsize=(8.4, 3.8))
    ax.scatter(range(len(work)), work[y_col].astype(float))
    ax.set_title(title, fontsize=11)
    ax.set_xlabel(x_col)
    ax.set_ylabel(y_col)
    step = max(1, len(work) // 10)
    idx = list(range(0, len(work), step))
    ax.set_xticks(idx)
    ax.set_xticklabels([str(work.iloc[i][x_col])[:14] for i in idx], rotation=45, ha="right", fontsize=7)
    ax.tick_params(axis="y", labelsize=8)
    ax.grid(alpha=0.25)
    fig.tight_layout()
    buf = BytesIO()
    fig.savefig(buf, format="png", dpi=180, bbox_inches="tight")
    plt.close(fig)
    buf.seek(0)
    return buf


def _build_pdf_sections(filtered_df: pd.DataFrame, stock_context: dict | None = None):
    sections = []

    # Base installed
    base_pairs = [
        ("Filtered records", f"{len(filtered_df):,}"),
        ("Countries", f"{filtered_df['Country'].nunique(dropna=True):,}"),
        ("Distributors", f"{filtered_df['Distributor name'].nunique(dropna=True):,}"),
        ("Instrument types", f"{filtered_df['Instrument type'].nunique(dropna=True):,}"),
        ("Routine assets", f"{int(filtered_df.get('Is in routine', pd.Series(dtype=bool)).sum()):,}"),
    ]
    top_country = filtered_df['Country'].fillna('No informado').value_counts().reset_index()
    top_country.columns = ['Country', 'Count']
    top_inst = filtered_df['Instrument type'].fillna('No informado').value_counts().reset_index()
    top_inst.columns = ['Instrument', 'Count']
    state_counts = filtered_df['Operational status grouped'].fillna('No informado').value_counts().reset_index()
    state_counts.columns = ['State', 'Count']
    sections.append({
        'title': 'Base Installed Overview',
        'summary_pairs': base_pairs,
        'charts': [
            _make_matplotlib_barh(top_country, 'Country', 'Count', 'Top countries'),
            _make_matplotlib_barh(top_inst, 'Instrument', 'Count', 'Instrument mix'),
            _make_matplotlib_barh(state_counts, 'State', 'Count', 'Operational status mix'),
        ],
        'table_title': 'Top filtered records snapshot',
        'table_df': prepare_pdf_report_table(filtered_df),
        'table_max_rows': 35,
    })

    # Machine config
    cfg_cols = [c for c in filtered_df.columns if c.startswith('CFG::')]
    cfg_pairs = [
        ('Assets with machine configuration', f"{int(filtered_df['Machine Configurations'].notna().sum()):,}"),
        ('Active config fields', f"{sum(int(filtered_df[c].notna().sum()) > 0 for c in cfg_cols):,}"),
        ('Average populated config fields', f"{filtered_df.get('Machine config fields populated', pd.Series([0])).fillna(0).mean():.1f}"),
    ]
    cfg_cov = pd.DataFrame([
        {'Field': c.replace('CFG::', ''), 'Count': int(filtered_df[c].notna().sum())} for c in cfg_cols if int(filtered_df[c].notna().sum()) > 0
    ]).sort_values('Count', ascending=False) if cfg_cols else pd.DataFrame(columns=['Field', 'Count'])
    cfg_top_table = cfg_cov.head(12).rename(columns={'Field':'Config field', 'Count':'Populated assets'})
    sections.append({
        'title': 'Machine Configuration',
        'summary_pairs': cfg_pairs,
        'charts': [
            _make_matplotlib_barh(cfg_cov, 'Field', 'Count', 'Configuration field coverage'),
        ],
        'table_title': 'Most populated configuration fields',
        'table_df': cfg_top_table,
        'table_max_rows': 12,
    })

    # OS
    os_df = filtered_df.copy()
    os_df['Operating System'] = os_df['Operating System'].fillna('No informado')
    os_df['OS Upgrade Bucket'] = os_df['Operating System'].map(os_upgrade_bucket)
    urgent_count = int(os_df['Operating System'].isin(['Windows XP','Windows Vista','Windows 7','Windows 2000']).sum())
    os_pairs = [
        ('Assets with OS identified', f"{int(filtered_df['Operating System'].notna().sum()):,}"),
        ('Unique OS values', f"{filtered_df['Operating System'].nunique(dropna=True):,}"),
        ('Legacy OS / urgent migration', f"{urgent_count:,}"),
        ('Unknown / not informed', f"{int(os_df['Operating System'].isin(['Unknown','No informado']).sum()):,}"),
    ]
    os_counts = os_df['Operating System'].value_counts().reset_index()
    os_counts.columns = ['Operating System', 'Count']
    os_bucket = os_df['OS Upgrade Bucket'].value_counts().reset_index()
    os_bucket.columns = ['Bucket', 'Count']
    urgent_table = os_df[os_df['Operating System'].isin(['Windows XP','Windows Vista','Windows 7','Windows 2000'])][['Country','Distributor name','Customer name','Instrument type','Serial number','Operating System']].copy()
    sections.append({
        'title': 'Operating System',
        'summary_pairs': os_pairs,
        'charts': [
            _make_matplotlib_barh(os_counts, 'Operating System', 'Count', 'OS distribution'),
            _make_matplotlib_barh(os_bucket, 'Bucket', 'Count', 'Upgrade prioritization'),
        ],
        'table_title': 'Assets requiring Windows upgrade',
        'table_df': urgent_table,
        'table_max_rows': 30,
    })

    # Processing / PM
    proc_df = filtered_df.copy()
    proc_df['Number of tests per day'] = pd.to_numeric(proc_df['Number of tests per day'], errors='coerce')
    today = pd.Timestamp.today().normalize()
    if 'PM next date' in proc_df.columns:
        pm_next = pd.to_datetime(proc_df['PM next date'], errors='coerce')
        proc_df['PM planner status'] = np.where(pm_next < today, 'Overdue', np.where(pm_next <= today + pd.Timedelta(days=90), 'Next 90 days', 'Planned later'))
    else:
        proc_df['PM planner status'] = 'No informado'
    product_lines = []
    for value in proc_df['Product Line'].fillna('').astype(str):
        for part in [p.strip() for p in re.split(r'[|;,/]', value) if p.strip()]:
            product_lines.append(part)
    product_df = pd.Series(product_lines, name='Product line').value_counts().reset_index() if product_lines else pd.DataFrame(columns=['Product line','count'])
    if not product_df.empty:
        product_df.columns=['Product line','Count']
    pm_status = proc_df['PM planner status'].value_counts().reset_index()
    pm_status.columns = ['PM status','Count']
    proc_pairs = [
        ('Average tests per day', safe_number_text(proc_df['Number of tests per day'].dropna().mean() if proc_df['Number of tests per day'].notna().any() else pd.NA, '0')),
        ('Max tests per day', safe_number_text(proc_df['Number of tests per day'].dropna().max() if proc_df['Number of tests per day'].notna().any() else pd.NA, '0')),
        ('Upcoming PM in next 90 days', f"{int((proc_df['PM planner status'] == 'Next 90 days').sum()):,}"),
        ('Overdue PM', f"{int((proc_df['PM planner status'] == 'Overdue').sum()):,}"),
    ]
    tests_table = proc_df[['Country','Distributor name','Instrument type','Serial number','Number of tests per day','PM planner status']].copy().sort_values('Number of tests per day', ascending=False, na_position='last')
    sections.append({
        'title': 'Processing and PM Planner',
        'summary_pairs': proc_pairs,
        'charts': [
            _make_matplotlib_scatter(proc_df[['Serial number','Number of tests per day']].dropna().sort_values('Number of tests per day', ascending=False), 'Serial number', 'Number of tests per day', 'Tests/day by serial'),
            _make_matplotlib_barh(product_df, 'Product line', 'Count', 'Top product lines'),
            _make_matplotlib_barh(pm_status, 'PM status', 'Count', 'PM planner status'),
        ],
        'table_title': 'Processing and PM snapshot',
        'table_df': tests_table,
        'table_max_rows': 30,
    })

    # Stock / spare parts
    stock_context = stock_context or {}
    if stock_context.get('available'):
        stock_pairs = [
            ('Detected distributor', stock_context.get('detected_distributor', 'N/A')),
            ('Families compared', ', '.join(stock_context.get('families', [])) or 'N/A'),
            ('Required SKUs', f"{stock_context.get('required_skus', 0):,}"),
            ('OK SKUs', f"{stock_context.get('ok_skus', 0):,}"),
            ('LOW SKUs', f"{stock_context.get('low_skus', 0):,}"),
            ('Missing SKUs', f"{stock_context.get('missing_skus', 0):,}"),
            ('Extra SKUs', f"{stock_context.get('extra_skus', 0):,}"),
            ('Gap total qty', safe_number_text(stock_context.get('gap_total', 0), '0')),
            ('Option 2 estimated cost', f"{stock_context.get('currency','EUR')} {float(stock_context.get('option2_cost', 0) or 0):,.2f}"),
        ]
        stock_status = pd.DataFrame({
            'Status': ['OK','LOW','Missing','Extras'],
            'Count': [stock_context.get('ok_skus',0), stock_context.get('low_skus',0), stock_context.get('missing_skus',0), stock_context.get('extra_skus',0)]
        })
        stock_top_gap = stock_context.get('top_gap_df', pd.DataFrame())
        full_comparison_df = stock_context.get('full_comparison_df', pd.DataFrame())
        purchase_df = stock_context.get('purchase_df', pd.DataFrame())
        extra_df = stock_context.get('extra_df', pd.DataFrame())
        if full_comparison_df is None:
            full_comparison_df = pd.DataFrame()
        if purchase_df is None:
            purchase_df = pd.DataFrame()
        if extra_df is None:
            extra_df = pd.DataFrame()

        summary_table_df = stock_top_gap[['Required Part Number','Required Description','Required Qty','Uploaded Qty','Qty Gap','Status','Option 2 Estimated Cost','Currency']] if not stock_top_gap.empty else pd.DataFrame(columns=['Required Part Number','Required Description','Required Qty','Uploaded Qty','Qty Gap','Status','Option 2 Estimated Cost','Currency'])
        sections.append({
            'title': 'Spare Parts / Carstock Gap',
            'summary_pairs': stock_pairs,
            'charts': [
                _make_matplotlib_barh(stock_status, 'Status', 'Count', 'Carstock coverage status'),
                _make_matplotlib_barh(stock_top_gap.rename(columns={'Required Part Number':'Part number','Qty Gap':'Gap qty'}), 'Part number', 'Gap qty', 'Top missing parts'),
            ],
            'table_title': 'Top spare parts gaps',
            'table_df': summary_table_df,
            'table_max_rows': 25,
        })

        if not full_comparison_df.empty:
            full_cols = [c for c in ['Required Part Number','Required Description','Required Qty','Uploaded Qty','Qty Gap','Coverage %','Status','Option 2 Unit Price','Option 2 Estimated Cost','Currency'] if c in full_comparison_df.columns]
            full_table = full_comparison_df[full_cols].copy()
            sections.append({
                'title': 'Spare Parts Annex - Full Comparison',
                'summary_pairs': [
                    ('Rows included', f"{len(full_table):,}"),
                    ('Scope', 'Complete carstock comparison for all required SKUs'),
                ],
                'charts': [],
                'table_title': 'Complete spare parts comparison',
                'table_df': full_table,
                'table_max_rows': max(len(full_table), 1),
            })

        missing_low_df = full_comparison_df[full_comparison_df['Status'].isin(['Missing','LOW'])].copy() if not full_comparison_df.empty and 'Status' in full_comparison_df.columns else pd.DataFrame()
        if not missing_low_df.empty:
            ml_cols = [c for c in ['Required Part Number','Required Description','Required Qty','Uploaded Qty','Qty Gap','Coverage %','Status','Option 2 Unit Price','Option 2 Estimated Cost','Currency'] if c in missing_low_df.columns]
            missing_low_table = missing_low_df[ml_cols].copy().sort_values(['Status','Qty Gap','Required Part Number'], ascending=[True, False, True])
            sections.append({
                'title': 'Spare Parts Annex - Missing and LOW',
                'summary_pairs': [
                    ('Rows included', f"{len(missing_low_table):,}"),
                    ('Scope', 'Items with missing or insufficient stock'),
                ],
                'charts': [],
                'table_title': 'Missing and LOW items',
                'table_df': missing_low_table,
                'table_max_rows': max(len(missing_low_table), 1),
            })

        if not purchase_df.empty:
            pur_cols = [c for c in ['Required Part Number','Required Description','Qty Gap','Option 2 Unit Price','Option 2 Estimated Cost','Currency','Status'] if c in purchase_df.columns]
            purchase_table = purchase_df[pur_cols].copy()
            sections.append({
                'title': 'Spare Parts Annex - Purchase Suggestion',
                'summary_pairs': [
                    ('Rows included', f"{len(purchase_table):,}"),
                    ('Scope', 'Suggested purchase to close current gap using option 2 pricing'),
                ],
                'charts': [],
                'table_title': 'Suggested purchase list',
                'table_df': purchase_table,
                'table_max_rows': max(len(purchase_table), 1),
            })

        if not extra_df.empty:
            ex_cols = [c for c in ['Uploaded Part Number','Uploaded Description','Uploaded Qty','Status'] if c in extra_df.columns]
            extra_table = extra_df[ex_cols].copy()
            sections.append({
                'title': 'Spare Parts Annex - Extra Items',
                'summary_pairs': [
                    ('Rows included', f"{len(extra_table):,}"),
                    ('Scope', 'Parts reported by the distributor that are not required by the selected carstock master'),
                ],
                'charts': [],
                'table_title': 'Extra items not required by master',
                'table_df': extra_table,
                'table_max_rows': max(len(extra_table), 1),
            })
    else:
        sections.append({
            'title': 'Spare Parts / Carstock Gap',
            'summary_pairs': [('Status', 'No spare parts comparison loaded in the current session.')],
            'charts': [],
            'table_title': 'Spare parts comparison',
            'table_df': pd.DataFrame(columns=['Info']),
            'table_max_rows': 1,
        })

    return sections


def build_pdf_report(
    filtered_df: pd.DataFrame,
    filter_summary: dict[str, str],
    report_title: str,
    author_name: str,
    author_role: str,
    signature_date: str,
    references_text: str = "",
    stock_context: dict | None = None,
) -> bytes:
    if not REPORTLAB_AVAILABLE:
        raise RuntimeError("reportlab no está instalado en el entorno.")
    buffer = BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=landscape(A4),
        leftMargin=1.0 * inch,
        rightMargin=1.0 * inch,
        topMargin=1.0 * inch,
        bottomMargin=1.0 * inch,
        title=report_title,
        author=author_name,
    )

    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name="APA_Title", parent=styles["Title"], fontName="Times-Bold", fontSize=16, leading=20, alignment=TA_CENTER, spaceAfter=10, textColor=colors.HexColor("#111111")))
    styles.add(ParagraphStyle(name="APA_Subtitle", parent=styles["Normal"], fontName="Times-Roman", fontSize=12, leading=24, alignment=TA_CENTER, spaceAfter=6, textColor=colors.HexColor("#444444")))
    styles.add(ParagraphStyle(name="APA_Heading", parent=styles["Heading2"], fontName="Times-Bold", fontSize=12, leading=16, alignment=TA_LEFT, spaceBefore=4, spaceAfter=6, textColor=colors.HexColor("#111111")))
    styles.add(ParagraphStyle(name="APA_Body", parent=styles["BodyText"], fontName="Times-Roman", fontSize=12, leading=24, alignment=TA_JUSTIFY, spaceAfter=6))
    styles.add(ParagraphStyle(name="APA_Cell", parent=styles["BodyText"], fontName="Times-Roman", fontSize=8, leading=10, alignment=TA_LEFT, wordWrap='CJK'))
    styles.add(ParagraphStyle(name="APA_Cell_Header", parent=styles["BodyText"], fontName="Times-Bold", fontSize=8, leading=10, alignment=TA_LEFT, textColor=colors.white))
    styles.add(ParagraphStyle(name="APA_Signature", parent=styles["BodyText"], fontName="Times-Roman", fontSize=12, leading=24, alignment=TA_LEFT, spaceAfter=3))

    elements = []
    today_str = datetime.now().strftime("%Y-%m-%d %H:%M")
    short_title = re.sub(r"\s+", " ", report_title.strip() or "Dashboard Report")[:80]

    elements.append(Spacer(1, 0.45 * inch))
    elements.append(Paragraph(report_title, styles["APA_Title"]))
    elements.append(Paragraph(author_name, styles["APA_Subtitle"]))
    elements.append(Paragraph(author_role, styles["APA_Subtitle"]))
    elements.append(Paragraph(f"Generated on {today_str}", styles["APA_Subtitle"]))
    elements.append(Spacer(1, 0.18 * inch))
    elements.append(Paragraph("This report was generated from the active dashboard filters and includes executive summaries, charts, and supporting tables for the visible tabs in the dashboard, including spare parts when loaded in the current session.", styles["APA_Body"]))

    elements += _summary_table_from_pairs("Executive Summary", [
        ("Filtered records", f"{len(filtered_df):,}"),
        ("Countries", f"{filtered_df['Country'].nunique(dropna=True):,}"),
        ("Distributors", f"{filtered_df['Distributor name'].nunique(dropna=True):,}"),
        ("Instrument types", f"{filtered_df['Instrument type'].nunique(dropna=True):,}"),
        ("Routine assets", f"{int(filtered_df.get('Is in routine', pd.Series(dtype=bool)).sum()):,}"),
    ], styles)
    elements.append(Spacer(1, 0.12 * inch))

    filters_table_data = [[Paragraph("<b>Filter</b>", styles["APA_Cell_Header"]), Paragraph("<b>Selected Value</b>", styles["APA_Cell_Header"])]]
    for key, value in filter_summary.items():
        filters_table_data.append([_paragraph_cell(key, styles["APA_Cell"]), _paragraph_cell(value, styles["APA_Cell"])])
    filters_table = Table(filters_table_data, colWidths=[2.2 * inch, 7.2 * inch], repeatRows=1)
    filters_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1f4e79")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("GRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#C8D6E5")),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.whitesmoke, colors.HexColor("#F7F9FB")]),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 5),
        ("RIGHTPADDING", (0, 0), (-1, -1), 5),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]))
    elements.append(Paragraph("Active Filters", styles["APA_Heading"]))
    elements.append(filters_table)

    from reportlab.platypus import Image
    sections = _build_pdf_sections(filtered_df, stock_context=stock_context)
    for section in sections:
        elements.append(PageBreak())
        elements.append(Paragraph(section['title'], styles['APA_Heading']))
        for block in _summary_table_from_pairs("Section Summary", section['summary_pairs'], styles):
            elements.append(block)
        charts = [c for c in section.get('charts', []) if c is not None]
        if charts:
            elements.append(Spacer(1, 0.08 * inch))
            for chart in charts:
                elements.append(Image(chart, width=7.8 * inch, height=3.5 * inch))
                elements.append(Spacer(1, 0.08 * inch))
        elif not MATPLOTLIB_AVAILABLE:
            elements.append(Paragraph("Chart rendering is unavailable because matplotlib is not installed in the environment.", styles['APA_Body']))

        table_df = section.get('table_df', pd.DataFrame())
        elements.append(Paragraph(section.get('table_title', 'Supporting Table'), styles['APA_Heading']))
        if table_df is not None and not table_df.empty:
            if section['title'] == 'Base Installed Overview':
                width_map = {
                    'Region': 0.9 * inch, 'Country': 0.9 * inch, 'Distributor': 1.2 * inch, 'Customer': 1.3 * inch,
                    'Instrument': 1.0 * inch, 'Serial': 0.95 * inch, 'State': 0.8 * inch, 'Raw Status': 1.0 * inch,
                    'OS': 0.85 * inch, 'Asset': 0.75 * inch, 'Install Date': 0.9 * inch, 'Contract': 1.2 * inch,
                }
                col_widths = [width_map.get(c, 0.95 * inch) for c in table_df.columns]
            elif section['title'].startswith('Spare Parts Annex') or section['title'] == 'Spare Parts / Carstock Gap':
                spare_width_map = {
                    'Required Part Number': 1.15 * inch,
                    'Required Description': 2.45 * inch,
                    'Required Qty': 0.65 * inch,
                    'Uploaded Qty': 0.7 * inch,
                    'Qty Gap': 0.65 * inch,
                    'Coverage %': 0.7 * inch,
                    'Status': 0.7 * inch,
                    'Option 2 Unit Price': 0.9 * inch,
                    'Option 2 Estimated Cost': 1.05 * inch,
                    'Currency': 0.55 * inch,
                    'Uploaded Part Number': 1.15 * inch,
                    'Uploaded Description': 2.6 * inch,
                    'Purchase Qty Option 2': 0.9 * inch,
                }
                col_widths = [spare_width_map.get(c, 0.9 * inch) for c in table_df.columns]
            else:
                col_widths = None
            max_rows = section.get('table_max_rows', len(table_df))
            elements.append(_df_to_wrapped_table(table_df, styles, col_widths=col_widths, max_rows=max_rows))
            if isinstance(max_rows, int) and len(table_df) > max_rows:
                elements.append(Paragraph(f"Note. Only the first {max_rows} rows are displayed in this section to preserve readability.", styles['APA_Body']))
        else:
            elements.append(Paragraph("No detailed rows are available for this section.", styles['APA_Body']))

    elements.append(PageBreak())
    elements.append(Paragraph("References", styles["APA_Heading"]))
    references = [line.strip() for line in references_text.splitlines() if line.strip()]
    if references:
        for ref in references:
            elements.append(Paragraph(_escape_pdf_text(ref), styles["APA_Body"]))
    else:
        elements.append(Paragraph("No external bibliographic references were provided. This document is based on the operational data loaded in the dashboard and, when available, on the spare parts comparison loaded in the current session.", styles["APA_Body"]))

    elements.append(Spacer(1, 0.14 * inch))
    elements.append(Paragraph("Signature", styles["APA_Heading"]))
    elements.append(Paragraph(_escape_pdf_text(author_name), styles["APA_Signature"]))
    elements.append(Paragraph(_escape_pdf_text(author_role), styles["APA_Signature"]))
    elements.append(Paragraph(f"Signature date: {_escape_pdf_text(signature_date)}", styles["APA_Signature"]))

    doc.build(elements, onFirstPage=lambda canvas, doc: _pdf_header_footer(canvas, doc, short_title), onLaterPages=lambda canvas, doc: _pdf_header_footer(canvas, doc, short_title))
    return buffer.getvalue()

def metric_card(label: str, value: str, subtitle: str = "") -> None:
    st.markdown(
        f"""
        <div class="metric-shell">
            <div class="metric-orb"></div>
            <div class="metric-label">{label}</div>
            <div class="metric-value">{value}</div>
            <div class="metric-sub">{subtitle}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def safe_text(value, fallback: str = "N/A") -> str:
    if pd.isna(value):
        return fallback
    text = str(value).strip()
    return text if text else fallback


def safe_number_text(value, fallback: str = "0") -> str:
    if pd.isna(value):
        return fallback
    try:
        val = float(value)
    except Exception:
        return fallback
    return f"{int(val):,}" if float(val).is_integer() else f"{val:,.1f}"


def normalize_part_number(value) -> str:
    if pd.isna(value):
        return ""
    text = str(value).strip()
    if not text:
        return ""
    if text.startswith('="') and text.endswith('"'):
        text = text[2:-1]
    text = text.replace(".0", "").replace(" ", "").replace("\n", "").replace("\t", "")
    return text.upper()


def normalize_key_text(value) -> str:
    if pd.isna(value):
        return ""
    return re.sub(r"[^a-z0-9]+", "", str(value).lower())


def normalize_search_text(value) -> str:
    if pd.isna(value):
        return ""
    return re.sub(r"[^a-z0-9]+", " ", str(value).lower()).strip()


DISTRIBUTOR_ALIASES = {
    "annar": "Annar Diagnostica Import sas",
    "bio nuclear": "Bio-Nuclear",
    "bionuclear": "Bio-Nuclear",
    "biotec": "Biotec del Paraguay",
    "biotec paraguay": "Biotec del Paraguay",
    "biotec del paraguay": "Biotec del Paraguay",
    "grupo bios": "Grupo Bios",
    "qls": "QLS",
    "simed ecuador": "Simed Ecuador",
    "simed peru": "Simed Perú",
    "wiener": "Wiener Lab",
    "wm argentina": "WM Argentina",
}


def infer_distributor_from_filename_strict(filename: str, distributor_options: list[str]) -> tuple[str | None, list[str]]:
    base = normalize_search_text(Path(filename).stem)
    if not base:
        return None, []

    alias_hits = []
    for alias, official_name in DISTRIBUTOR_ALIASES.items():
        if re.search(rf"\b{re.escape(alias)}\b", base):
            if official_name in distributor_options:
                alias_hits.append(official_name)

    alias_hits = list(dict.fromkeys(alias_hits))
    if len(alias_hits) == 1:
        return alias_hits[0], alias_hits
    if len(alias_hits) > 1:
        return None, alias_hits

    strong_hits = []
    for distributor in distributor_options:
        norm = normalize_search_text(distributor)
        if not norm:
            continue
        if re.search(rf"\b{re.escape(norm)}\b", base):
            strong_hits.append(distributor)

    strong_hits = list(dict.fromkeys(strong_hits))
    if len(strong_hits) == 1:
        return strong_hits[0], strong_hits
    if len(strong_hits) > 1:
        return None, strong_hits

    candidates = []
    weak_tokens = {"bio", "lab", "labs", "import", "sas", "ltd", "pte", "sa", "srl", "corp", "group"}

    for distributor in distributor_options:
        tokens = [t for t in normalize_search_text(distributor).split() if len(t) >= 4 and t not in weak_tokens]
        if not tokens:
            continue
        hits = sum(1 for t in tokens if re.search(rf"\b{re.escape(t)}\b", base))
        if hits > 0:
            candidates.append((hits, len(tokens), distributor))

    candidates.sort(key=lambda x: (-x[0], -x[1], x[2]))
    if not candidates:
        return None, []

    top_score = candidates[0][0]
    tied = [c[2] for c in candidates if c[0] == top_score]
    if len(tied) == 1:
        return tied[0], tied
    return None, tied


def os_upgrade_bucket(value) -> str:
    text = safe_text(value, "No informado")
    if text == "Windows 10":
        return "Windows 10 / OK"
    if text in {"Windows XP", "Windows Vista", "Windows 7", "Windows 2000"}:
        return "Legacy / urgente migrar"
    if text in {"Unknown", "No informado", "Not installed"}:
        return "Revisar campo OS"
    return "Otro OS / validar"


def format_date_for_hover(value) -> str:
    if pd.isna(value):
        return "N/A"
    try:
        return pd.to_datetime(value).strftime("%Y-%m-%d")
    except Exception:
        return safe_text(value)


def to_numeric_series(series: pd.Series) -> pd.Series:
    cleaned = (
        series.fillna("")
        .astype(str)
        .str.strip()
        .str.replace(r"[^0-9,\.\-]", "", regex=True)
        .str.replace(",", ".", regex=False)
    )
    return pd.to_numeric(cleaned, errors="coerce")


def normalize_instrument_type(value) -> str:
    text = safe_text(value, "")
    if ":" in text:
        text = text.split(":", 1)[1]
    return text.strip() or safe_text(value)


def get_uploaded_file_signature(uploaded_file) -> str:
    if uploaded_file is None:
        return ""
    content = uploaded_file.getvalue()
    raw = f"{uploaded_file.name}|{len(content)}|".encode("utf-8") + content
    return hashlib.md5(raw).hexdigest()


def read_table_any(uploaded_file) -> pd.DataFrame:
    name = uploaded_file.name.lower()
    raw = uploaded_file.getvalue()

    if name.endswith(".csv"):
        attempts = [
            {"sep": ";", "encoding": "utf-8-sig"},
            {"sep": ";", "encoding": "latin1"},
            {"sep": ",", "encoding": "utf-8-sig"},
            {"sep": ",", "encoding": "latin1"},
            {"sep": None, "encoding": "utf-8-sig"},
            {"sep": None, "encoding": "latin1"},
        ]
        for att in attempts:
            try:
                text = raw.decode(att["encoding"], errors="replace")
                df = pd.read_csv(StringIO(text), sep=att["sep"], engine="python", on_bad_lines="skip")
                if df.shape[1] >= 3:
                    return df
            except Exception:
                continue
        raise ValueError(f"No fue posible leer el CSV: {uploaded_file.name}")

    if name.endswith(".xlsx") or name.endswith(".xls"):
        book = pd.ExcelFile(BytesIO(raw))
        preferred = None
        best_score = -1
        for sheet in book.sheet_names:
            s = str(sheet).lower()
            score = 0
            if "datos" in s:
                score += 25
            if "combined" in s:
                score += 20
            if "consolidated" in s:
                score += 20
            if "records" in s:
                score += 18
            if "data" in s:
                score += 15
            if score > best_score:
                best_score = score
                preferred = sheet
        if preferred is None:
            preferred = book.sheet_names[0]
        return pd.read_excel(book, sheet_name=preferred)

    raise ValueError(f"Formato no soportado: {uploaded_file.name}")


@st.cache_data(show_spinner=False)
def load_records(file_bytes: bytes) -> pd.DataFrame:
    content = file_bytes.decode("utf-8-sig", errors="replace").splitlines()
    rows = [line.split(";") for line in content[1:] if line.strip()]
    df = pd.DataFrame(rows, columns=CUSTOM_HEADERS)
    if "_blank" in df.columns:
        df = df.drop(columns=["_blank"])

    for col in df.columns:
        if df[col].dtype == object:
            df[col] = df[col].astype(str).str.strip()
            df[col] = df[col].replace({"": pd.NA, "None": pd.NA, "nan": pd.NA, "<NA>": pd.NA})

    def unexcel(value):
        if pd.isna(value):
            return pd.NA
        value = str(value).strip()
        if value.startswith('="') and value.endswith('"'):
            return value[2:-1]
        return value

    for col in ["Latitude", "Longitude", "Serial number"]:
        df[col] = df[col].map(unexcel)

    for col in ["Latitude", "Longitude", "Number of tests per day", "PM frequency", "Contract duration"]:
        df[col] = to_numeric_series(df[col])

    for col in ["Installation date", "PM last date", "PM next date"]:
        df[col] = pd.to_datetime(df[col], dayfirst=True, errors="coerce")

    df["Instrument family"] = df["Instrument type"].map(normalize_instrument_type)
    df["Operational status grouped"] = df["Operational status"].map(normalize_operational_status)

    today = pd.Timestamp(date.today())
    df["Age (years)"] = ((today - df["Installation date"]).dt.days / 365.25).round(1)
    df["Is in routine"] = df["Operational status"].fillna("").astype(str).str.upper().eq("IN ROUTINE")
    df["Has geolocation"] = df["Latitude"].notna() & df["Longitude"].notna()

    yes_map = {"yes", "y", "true", "1"}
    assay_flags = {}
    for col in ASSAY_COLS:
        normalized = df[col].fillna("No").astype(str).str.strip()
        df[col] = normalized
        assay_flags[f"FLAG::{col}"] = normalized.str.lower().isin(yes_map)
    if assay_flags:
        assay_flags_df = pd.DataFrame(assay_flags)
        df = pd.concat([df, assay_flags_df], axis=1)
        df["Enabled assay count"] = assay_flags_df.sum(axis=1)
    else:
        df["Enabled assay count"] = 0

    base_cols = [c for c in CUSTOM_HEADERS if c != "_blank"]
    df["Data completeness %"] = (df[base_cols].notna().sum(axis=1) / len(base_cols) * 100).round(1)
    return df


def adapt_uploaded_records_to_standard(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()

    if "_blank" in out.columns:
        out = out.drop(columns=["_blank"])

    exact_matches = sum(1 for c in CUSTOM_HEADERS if c in out.columns)
    if exact_matches >= 20:
        return out

    if out.shape[1] == len(CUSTOM_HEADERS):
        out.columns = CUSTOM_HEADERS
        if "_blank" in out.columns:
            out = out.drop(columns=["_blank"])
        return out

    if out.shape[1] == len(CUSTOM_HEADERS) - 1:
        out.columns = [c for c in CUSTOM_HEADERS if c != "_blank"]
        return out

    return out


def parse_uploaded_records(uploaded_file) -> pd.DataFrame:
    name = uploaded_file.name.lower()
    raw = uploaded_file.getvalue()

    if name.endswith(".csv"):
        return load_records(raw)

    table = read_table_any(uploaded_file)
    table = adapt_uploaded_records_to_standard(table)

    df = table.copy()

    if "_blank" in df.columns:
        df = df.drop(columns=["_blank"])

    for col in df.columns:
        if df[col].dtype == object:
            df[col] = df[col].astype(str).str.strip()
            df[col] = df[col].replace({"": pd.NA, "None": pd.NA, "nan": pd.NA, "<NA>": pd.NA})

    for missing in [c for c in CUSTOM_HEADERS if c != "_blank" and c not in df.columns]:
        df[missing] = pd.NA

    def unexcel(value):
        if pd.isna(value):
            return pd.NA
        value = str(value).strip()
        if value.startswith('="') and value.endswith('"'):
            return value[2:-1]
        return value

    for col in ["Latitude", "Longitude", "Serial number"]:
        if col in df.columns:
            df[col] = df[col].map(unexcel)

    for col in ["Latitude", "Longitude", "Number of tests per day", "PM frequency", "Contract duration"]:
        if col in df.columns:
            df[col] = to_numeric_series(df[col])

    for col in ["Installation date", "PM last date", "PM next date"]:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], dayfirst=True, errors="coerce")

    df["Instrument family"] = df["Instrument type"].map(normalize_instrument_type)
    df["Operational status grouped"] = df["Operational status"].map(normalize_operational_status)

    today = pd.Timestamp(date.today())
    df["Age (years)"] = ((today - df["Installation date"]).dt.days / 365.25).round(1)
    df["Is in routine"] = df["Operational status"].fillna("").astype(str).str.upper().eq("IN ROUTINE")
    df["Has geolocation"] = df["Latitude"].notna() & df["Longitude"].notna()

    yes_map = {"yes", "y", "true", "1"}
    assay_flags = {}
    for col in ASSAY_COLS:
        if col not in df.columns:
            df[col] = pd.NA
        normalized = df[col].fillna("No").astype(str).str.strip()
        df[col] = normalized
        assay_flags[f"FLAG::{col}"] = normalized.str.lower().isin(yes_map)

    if assay_flags:
        assay_flags_df = pd.DataFrame(assay_flags)
        df = pd.concat([df, assay_flags_df], axis=1)
        df["Enabled assay count"] = assay_flags_df.sum(axis=1)
    else:
        df["Enabled assay count"] = 0

    base_cols = [c for c in CUSTOM_HEADERS if c != "_blank" and c in df.columns]
    df["Data completeness %"] = (df[base_cols].notna().sum(axis=1) / len(base_cols) * 100).round(1)
    return df


def get_active_records_dataset(uploaded_file, sample_candidates: list[Path]) -> tuple[pd.DataFrame, str]:
    if uploaded_file is not None:
        current_sig = get_uploaded_file_signature(uploaded_file)
        saved_sig = st.session_state.get("records_active_signature", "")

        if "records_active_df" not in st.session_state or current_sig != saved_sig:
            active_df = parse_uploaded_records(uploaded_file)
            st.session_state["records_active_df"] = active_df.copy()
            st.session_state["records_active_signature"] = current_sig
            st.session_state["records_active_name"] = uploaded_file.name

        return st.session_state["records_active_df"].copy(), st.session_state["records_active_name"]

    if "records_active_df" in st.session_state and st.session_state.get("records_active_name"):
        return st.session_state["records_active_df"].copy(), st.session_state["records_active_name"]

    if sample_candidates:
        sample_path = sample_candidates[0]
        active_df = load_records(sample_path.read_bytes())
        st.session_state["records_active_df"] = active_df.copy()
        st.session_state["records_active_signature"] = f"sample::{sample_path.name}"
        st.session_state["records_active_name"] = sample_path.name
        return active_df, sample_path.name

    return pd.DataFrame(), ""


@st.cache_data(show_spinner=False)
def parse_machine_configuration(df: pd.DataFrame) -> tuple[pd.DataFrame, list[str]]:
    parsed_rows = []
    key_set = set()
    for raw in df["Machine Configurations"].fillna(""):
        row_dict = {}
        text = str(raw).strip()
        if text:
            for part in [p.strip() for p in text.split("|") if p.strip()]:
                if ":" in part:
                    key, value = part.split(":", 1)
                    key = key.strip()
                    value = value.strip()
                    if value:
                        row_dict[key] = value
                        key_set.add(key)
        parsed_rows.append(row_dict)
    config_cols = sorted(key_set)
    if config_cols:
        cfg_df = pd.DataFrame([{key: row.get(key, pd.NA) for key in config_cols} for row in parsed_rows])
    else:
        cfg_df = pd.DataFrame(index=df.index)
    cfg_df = cfg_df.add_prefix("CFG::")
    return pd.concat([df.reset_index(drop=True), cfg_df.reset_index(drop=True)], axis=1), config_cols


@st.cache_data(show_spinner=False)
def add_operating_system_columns(df: pd.DataFrame, config_cols: list[str]) -> pd.DataFrame:
    os_candidates = ["CFG::Operative System", "CFG::ETI-Max 3000 Operative System", "CFG::LQS PC OS"]
    existing = [col for col in os_candidates if col in df.columns]

    def normalize_os(value):
        if pd.isna(value):
            return pd.NA
        text = str(value).strip()
        if not text:
            return pd.NA
        low = text.lower()
        if "don't know" in low or "dont know" in low or low == "unknown":
            return "Unknown"
        if "not installed" in low:
            return "Not installed"
        if "win10" in low or "windows 10" in low:
            return "Windows 10"
        if "vista" in low:
            return "Windows Vista"
        if "windows 7" in low or low == "win7":
            return "Windows 7"
        if "windows xp" in low or low == "xp":
            return "Windows XP"
        if "windows 2000" in low:
            return "Windows 2000"
        return text

    if existing:
        os_raw = pd.Series(pd.NA, index=df.index, dtype="object")
        for col in existing:
            os_raw = os_raw.fillna(df[col])
        df["Operating System Raw"] = os_raw
        df["Operating System"] = os_raw.map(normalize_os)
    else:
        df["Operating System Raw"] = pd.NA
        df["Operating System"] = pd.NA

    cfg_prefix_cols = [f"CFG::{c}" for c in config_cols if f"CFG::{c}" in df.columns]
    df["Machine config fields populated"] = df[cfg_prefix_cols].notna().sum(axis=1) if cfg_prefix_cols else 0
    return df


@st.cache_data(show_spinner=False)
def to_csv_download(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False).encode("utf-8-sig")


def load_table_file(file_bytes: bytes, filename: str) -> pd.DataFrame:
    name = filename.lower()
    if name.endswith(".csv"):
        return pd.read_csv(BytesIO(file_bytes))
    return pd.read_excel(BytesIO(file_bytes))


def detect_stock_columns(df: pd.DataFrame) -> tuple[str | None, str | None, str | None]:
    normalized = {col: re.sub(r"[^a-z0-9]+", " ", str(col).lower()).strip() for col in df.columns}
    part_col = qty_col = desc_col = None
    for col, norm in normalized.items():
        if part_col is None and ("part number" in norm or norm == "part no" or "product code" in norm or "material" in norm):
            part_col = col
        if qty_col is None and ("quantity" in norm or norm == "qty" or "stock" in norm or "cantidad" in norm):
            qty_col = col
        if desc_col is None and ("description" in norm or "descripcion" in norm or "product description" in norm):
            desc_col = col
    if part_col is None and len(df.columns) >= 1:
        part_col = df.columns[0]
    if qty_col is None and len(df.columns) >= 2:
        qty_col = df.columns[1]
    if desc_col is None and len(df.columns) >= 3:
        desc_col = df.columns[2]
    return part_col, qty_col, desc_col


@st.cache_data(show_spinner=False)
def load_spare_master_legacy(file_bytes: bytes) -> dict[str, pd.DataFrame]:
    book = pd.ExcelFile(BytesIO(file_bytes))
    mapping = {"LXL Carstock": "LXL", "LXS Carstock": "LXS", "MDX Carstock": "MDX", "EMX Carstock": "EMX"}
    output = {}
    for sheet, family in mapping.items():
        if sheet not in book.sheet_names:
            continue
        df = pd.read_excel(book, sheet_name=sheet)

        part_col = next((c for c in df.columns if "PART NUMBER" in str(c).upper()), None)
        desc_col = next((c for c in df.columns if "DESCRIPTION" in str(c).upper()), None)
        qty_col = next((c for c in df.columns if "QUANTITY" in str(c).upper()), None)
        if not part_col or not qty_col:
            continue

        slim = pd.DataFrame(
            {
                "Required Distributor": "",
                "Required Family": family,
                "Required Part Number": df[part_col],
                "Required Description": df[desc_col] if desc_col else pd.NA,
                "Required Qty": to_numeric_series(df[qty_col]),
            }
        )
        slim["Part Key"] = slim["Required Part Number"].map(normalize_part_number)
        slim = slim[slim["Part Key"] != ""].copy()
        slim["Required Description"] = slim["Required Description"].fillna("").astype(str).str.strip()
        slim["Required Qty"] = pd.to_numeric(slim["Required Qty"], errors="coerce").fillna(0.0)
        slim = slim.groupby(["Part Key", "Required Family"], as_index=False).agg(
            {
                "Required Distributor": "first",
                "Required Part Number": "first",
                "Required Description": "first",
                "Required Qty": "sum",
            }
        )
        output[family] = slim.sort_values(["Required Qty", "Required Part Number"], ascending=[False, True]).reset_index(drop=True)
    return output


def detect_carstock_master_columns(df: pd.DataFrame) -> tuple[str | None, str | None, str | None, str | None, str | None]:
    normalized = {col: re.sub(r"[^a-z0-9]+", " ", str(col).lower()).strip() for col in df.columns}
    distributor_col = family_col = part_col = qty_col = desc_col = None

    for col, norm in normalized.items():
        if distributor_col is None and (
            "distributor" in norm
            or norm in {"dealer", "dealer name", "dist", "distributor name"}
        ):
            distributor_col = col

        if family_col is None and (
            "family" in norm
            or "platform" in norm
            or "carstock family" in norm
            or "instrument family" in norm
            or "family code" in norm
            or norm in {"instrument type", "instrument", "system"}
        ):
            family_col = col

        if part_col is None and (
            "part number" in norm
            or "latest part number" in norm
            or "part no" in norm
            or "pn" == norm
            or "material" in norm
            or "code" == norm
            or "product code" in norm
        ):
            part_col = col

        if qty_col is None and (
            "required qty" in norm
            or "required quantity" in norm
            or "carstock qty" in norm
            or "car stock qty" in norm
            or "quantity" in norm
            or norm in {"qty", "cantidad"}
        ):
            qty_col = col

        if desc_col is None and (
            "description" in norm
            or "descripcion" in norm
            or "product description" in norm
            or "spare part description" in norm
        ):
            desc_col = col

    return distributor_col, family_col, part_col, qty_col, desc_col


def normalize_family_code(value) -> str:
    text = normalize_key_text(value)
    if not text:
        return ""
    if "mdx" in text:
        return "MDX"
    if "emx" in text:
        return "EMX"
    if "xs" in text or "liaisonxs" in text:
        return "LXS"
    if "xl" in text or "las" in text or "liaisonxl" in text:
        return "LXL"
    return ""


def infer_families_from_instruments(instruments: list[str]) -> list[str]:
    families = []
    for inst in instruments:
        fam = normalize_family_code(inst)
        if fam:
            families.append(fam)
    return sorted(set(families))


def build_master_slim(
    df: pd.DataFrame,
    distributor_col: str | None,
    family_col: str | None,
    part_col: str,
    qty_col: str,
    desc_col: str | None,
    fallback_distributor: str = "",
    fallback_family: str = "",
) -> pd.DataFrame:
    slim = pd.DataFrame(
        {
            "Required Distributor": df[distributor_col] if distributor_col and distributor_col in df.columns else fallback_distributor,
            "Required Family": df[family_col] if family_col and family_col in df.columns else fallback_family,
            "Required Part Number": df[part_col],
            "Required Description": df[desc_col] if desc_col and desc_col in df.columns else "",
            "Required Qty": to_numeric_series(df[qty_col]),
        }
    )
    slim["Required Distributor"] = slim["Required Distributor"].fillna("").astype(str).str.strip()
    slim["Required Family"] = slim["Required Family"].map(normalize_family_code)
    if fallback_family:
        slim["Required Family"] = slim["Required Family"].replace("", fallback_family)
    slim["Required Description"] = slim["Required Description"].fillna("").astype(str).str.strip()
    slim["Required Qty"] = pd.to_numeric(slim["Required Qty"], errors="coerce").fillna(0.0)
    slim["Part Key"] = slim["Required Part Number"].map(normalize_part_number)
    slim["Distributor Key"] = slim["Required Distributor"].map(normalize_key_text)
    slim = slim[(slim["Part Key"] != "") & (slim["Required Qty"] > 0)].copy()
    slim = slim.groupby(["Distributor Key", "Required Distributor", "Required Family", "Part Key"], as_index=False).agg(
        {
            "Required Part Number": "first",
            "Required Description": "first",
            "Required Qty": "sum",
        }
    )
    return slim


def detect_price_reference_columns(df: pd.DataFrame) -> tuple[str | None, str | None, str | None, str | None]:
    normalized = {col: re.sub(r"[^a-z0-9]+", " ", str(col).lower()).strip() for col in df.columns}
    part_col = desc_col = option2_col = currency_col = None
    for col, norm in normalized.items():
        if part_col is None and (
            "part number" in norm or "latest part number" in norm or norm in {"part no", "pn"} or "material" in norm
        ):
            part_col = col
        if desc_col is None and ("description" in norm or "descripcion" in norm or "product description" in norm):
            desc_col = col
        if option2_col is None and ("option 2" in norm or "opt2" in norm):
            option2_col = col
        if currency_col is None and "currency" in norm:
            currency_col = col
    return part_col, desc_col, option2_col, currency_col


def build_price_reference(df: pd.DataFrame, part_col: str, option2_col: str, desc_col: str | None, currency_col: str | None) -> pd.DataFrame:
    price_df = pd.DataFrame(
        {
            "Price Part Number": df[part_col],
            "Price Description": df[desc_col] if desc_col and desc_col in df.columns else "",
            "Option 2 Unit Price": to_numeric_series(df[option2_col]),
            "Currency": df[currency_col] if currency_col and currency_col in df.columns else "EUR",
        }
    )
    price_df["Part Key"] = price_df["Price Part Number"].map(normalize_part_number)
    price_df["Price Description"] = price_df["Price Description"].fillna("").astype(str).str.strip()
    price_df["Currency"] = price_df["Currency"].fillna("EUR").astype(str).str.strip().replace("", "EUR")
    price_df["Option 2 Unit Price"] = pd.to_numeric(price_df["Option 2 Unit Price"], errors="coerce")
    price_df = price_df[(price_df["Part Key"] != "") & price_df["Option 2 Unit Price"].notna() & (price_df["Option 2 Unit Price"] > 0)].copy()
    if price_df.empty:
        return pd.DataFrame(columns=["Part Key", "Option 2 Unit Price", "Currency", "Price Description"])
    price_df = price_df.groupby("Part Key", as_index=False).agg(
        {
            "Option 2 Unit Price": "first",
            "Currency": "first",
            "Price Description": "first",
        }
    )
    return price_df


@st.cache_data(show_spinner=False)
def load_carstock_master_bundle(file_bytes: bytes, filename: str) -> dict[str, object]:
    path = Path(filename)
    ext = path.suffix.lower()

    legacy_families = {}
    consolidated_frames = []
    price_frames = []

    if ext in {".xlsx", ".xls"}:
        try:
            book = pd.ExcelFile(BytesIO(file_bytes))
        except Exception:
            book = None

        if book is not None:
            sheet_names = set(book.sheet_names)
            if any(sheet in sheet_names for sheet in {"LXL Carstock", "LXS Carstock", "MDX Carstock", "EMX Carstock"}):
                legacy_families = load_spare_master_legacy(file_bytes)

            for sheet in book.sheet_names:
                try:
                    df = pd.read_excel(book, sheet_name=sheet)
                except Exception:
                    continue
                if df is None or df.empty:
                    continue

                price_part_col, price_desc_col, option2_col, currency_col = detect_price_reference_columns(df)
                if price_part_col and option2_col:
                    price_ref = build_price_reference(df, price_part_col, option2_col, price_desc_col, currency_col)
                    if not price_ref.empty:
                        price_frames.append(price_ref)

                distributor_col, family_col, part_col, qty_col, desc_col = detect_carstock_master_columns(df)
                if not part_col or not qty_col:
                    continue

                fallback_family = normalize_family_code(sheet)
                slim = build_master_slim(
                    df,
                    distributor_col=distributor_col,
                    family_col=family_col,
                    part_col=part_col,
                    qty_col=qty_col,
                    desc_col=desc_col,
                    fallback_family=fallback_family,
                )
                if not slim.empty:
                    consolidated_frames.append(slim)

    else:
        df = load_table_file(file_bytes, filename)
        if df is not None and not df.empty:
            price_part_col, price_desc_col, option2_col, currency_col = detect_price_reference_columns(df)
            if price_part_col and option2_col:
                price_ref = build_price_reference(df, price_part_col, option2_col, price_desc_col, currency_col)
                if not price_ref.empty:
                    price_frames.append(price_ref)
            distributor_col, family_col, part_col, qty_col, desc_col = detect_carstock_master_columns(df)
            if part_col and qty_col:
                consolidated_frames.append(
                    build_master_slim(
                        df,
                        distributor_col=distributor_col,
                        family_col=family_col,
                        part_col=part_col,
                        qty_col=qty_col,
                        desc_col=desc_col,
                    )
                )

    if consolidated_frames:
        consolidated = pd.concat(consolidated_frames, ignore_index=True)
        consolidated["Required Distributor"] = consolidated["Required Distributor"].fillna("").astype(str).str.strip()
        consolidated["Required Family"] = consolidated["Required Family"].fillna("").astype(str).str.strip()
        consolidated["Distributor Key"] = consolidated["Required Distributor"].map(normalize_key_text)
        consolidated = consolidated.groupby(
            ["Distributor Key", "Required Distributor", "Required Family", "Part Key"], as_index=False
        ).agg(
            {
                "Required Part Number": "first",
                "Required Description": "first",
                "Required Qty": "sum",
            }
        )
    else:
        consolidated = pd.DataFrame(
            columns=[
                "Distributor Key",
                "Required Distributor",
                "Required Family",
                "Part Key",
                "Required Part Number",
                "Required Description",
                "Required Qty",
            ]
        )

    if price_frames:
        price_reference = pd.concat(price_frames, ignore_index=True)
        price_reference = price_reference.groupby("Part Key", as_index=False).agg(
            {
                "Option 2 Unit Price": "first",
                "Currency": "first",
                "Price Description": "first",
            }
        )
    else:
        price_reference = pd.DataFrame(columns=["Part Key", "Option 2 Unit Price", "Currency", "Price Description"])

    distributor_options = sorted([d for d in consolidated["Required Distributor"].dropna().astype(str).unique().tolist() if d.strip()])
    family_options = sorted([f for f in consolidated["Required Family"].dropna().astype(str).unique().tolist() if f.strip()])

    return {
        "legacy_families": legacy_families,
        "consolidated": consolidated,
        "price_reference": price_reference,
        "master_distributors": distributor_options,
        "master_families": family_options,
    }


def build_required_master_from_scope(
    master_bundle: dict[str, object],
    assigned_distributor: str,
    selected_families: list[str],
) -> tuple[pd.DataFrame, str]:
    consolidated = master_bundle.get("consolidated", pd.DataFrame())
    legacy_families = master_bundle.get("legacy_families", {})

    if consolidated is not None and not consolidated.empty:
        scoped = consolidated.copy()
        if assigned_distributor and assigned_distributor != "<sin asignar>" and scoped["Distributor Key"].astype(str).str.len().gt(0).any():
            scoped = scoped[scoped["Distributor Key"].eq(normalize_key_text(assigned_distributor))]
        if selected_families:
            scoped = scoped[scoped["Required Family"].isin(selected_families)]

        scoped = scoped.groupby("Part Key", as_index=False).agg(
            {
                "Required Part Number": "first",
                "Required Description": "first",
                "Required Qty": "sum",
            }
        )
        return scoped.sort_values(["Required Qty", "Required Part Number"], ascending=[False, True]).reset_index(drop=True), "consolidated"

    selected_families = [f for f in selected_families if f in legacy_families]
    if not selected_families:
        return pd.DataFrame(columns=["Part Key", "Required Part Number", "Required Description", "Required Qty"]), "legacy"

    scoped = pd.concat([legacy_families[f] for f in selected_families], ignore_index=True)
    scoped = scoped.groupby("Part Key", as_index=False).agg(
        {
            "Required Part Number": "first",
            "Required Description": "first",
            "Required Qty": "sum",
        }
    )
    return scoped.sort_values(["Required Qty", "Required Part Number"], ascending=[False, True]).reset_index(drop=True), "legacy"


def prepare_uploaded_stock(stock_df: pd.DataFrame, part_col: str, qty_col: str, desc_col: str | None) -> pd.DataFrame:
    work = stock_df.copy()
    work["Uploaded Part Number"] = work[part_col]
    work["Uploaded Qty"] = to_numeric_series(work[qty_col]).fillna(0.0)
    if desc_col is not None and desc_col in work.columns:
        work["Uploaded Description"] = work[desc_col]
    else:
        work["Uploaded Description"] = ""
    work["Uploaded Description"] = work["Uploaded Description"].fillna("").astype(str).str.strip()
    work["Part Key"] = work["Uploaded Part Number"].map(normalize_part_number)
    work = work[work["Part Key"] != ""].copy()
    stock_slim = work.groupby("Part Key", as_index=False).agg(
        {"Uploaded Part Number": "first", "Uploaded Description": "first", "Uploaded Qty": "sum"}
    )
    stock_slim["Uploaded Qty"] = pd.to_numeric(stock_slim["Uploaded Qty"], errors="coerce").fillna(0.0)
    return stock_slim.sort_values(["Uploaded Qty", "Uploaded Part Number"], ascending=[False, True]).reset_index(drop=True)


def compare_stock(
    master_df: pd.DataFrame,
    stock_df: pd.DataFrame,
    part_col: str,
    qty_col: str,
    desc_col: str | None,
    price_reference: pd.DataFrame | None = None,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    stock_slim = prepare_uploaded_stock(stock_df, part_col, qty_col, desc_col)

    merged = master_df.copy().merge(stock_slim, on="Part Key", how="left")
    merged["Required Qty"] = pd.to_numeric(merged["Required Qty"], errors="coerce").fillna(0.0)
    merged["Uploaded Qty"] = pd.to_numeric(merged["Uploaded Qty"], errors="coerce").fillna(0.0)
    merged["Uploaded Part Number"] = merged["Uploaded Part Number"].fillna("")
    merged["Uploaded Description"] = merged["Uploaded Description"].fillna("")
    merged["Qty Gap"] = (merged["Required Qty"] - merged["Uploaded Qty"]).clip(lower=0.0)
    denominator = merged["Required Qty"].replace(0, np.nan).astype(float)
    merged["Coverage %"] = ((merged["Uploaded Qty"].astype(float) / denominator) * 100).round(1).fillna(0.0)

    low_mask = (merged["Uploaded Qty"] > 0) & (merged["Uploaded Qty"] < merged["Required Qty"])
    merged["Status"] = np.where(
        merged["Qty Gap"] <= 0,
        "OK",
        np.where(low_mask, "LOW", "Missing"),
    )

    if price_reference is not None and not price_reference.empty:
        merged = merged.merge(price_reference, on="Part Key", how="left")
    else:
        merged["Option 2 Unit Price"] = np.nan
        merged["Currency"] = "EUR"
        merged["Price Description"] = ""

    merged["Option 2 Unit Price"] = pd.to_numeric(merged["Option 2 Unit Price"], errors="coerce")
    merged["Currency"] = merged["Currency"].fillna("EUR").astype(str).str.strip().replace("", "EUR")
    merged["Purchase Qty Option 2"] = merged["Qty Gap"]
    merged["Option 2 Estimated Cost"] = (merged["Purchase Qty Option 2"] * merged["Option 2 Unit Price"]).round(2)

    extra_df = stock_slim[~stock_slim["Part Key"].isin(master_df["Part Key"])].copy()
    if not extra_df.empty:
        extra_df["Status"] = "Extra / no requerido"

    merged = merged.sort_values(["Status", "Qty Gap", "Required Qty"], ascending=[True, False, False]).reset_index(drop=True)
    extra_df = extra_df.sort_values(["Uploaded Qty", "Uploaded Part Number"], ascending=[False, True]).reset_index(drop=True)
    return merged, extra_df, stock_slim


def active_config_fields(df: pd.DataFrame, config_keys: list[str]) -> list[str]:
    active = []
    for key in config_keys:
        col = f"CFG::{key}"
        if col in df.columns and df[col].notna().any():
            active.append(key)
    return active


def build_distributor_status_chart(df: pd.DataFrame, selected_model: str) -> go.Figure:
    fig = go.Figure()

    if df.empty or not selected_model:
        fig.update_layout(title="Estado por distribuidor")
        return glow_layout(fig, 620, 17)

    work = df.copy()
    work["Instrument type"] = work["Instrument type"].fillna("No informado").astype(str)
    work["Distributor name"] = work["Distributor name"].fillna("No informado").astype(str)
    work["Operational status"] = work["Operational status"].fillna("No informado").astype(str).str.strip()
    work["Status for chart"] = np.where(work["Operational status"].eq(""), "No informado", work["Operational status"])

    model_df = work[work["Instrument type"] == selected_model].copy()
    if model_df.empty:
        fig.update_layout(title=f"Estado por distribuidor | {selected_model}")
        return glow_layout(fig, 620, 17)

    summary = (
        model_df.groupby(["Distributor name", "Status for chart"], dropna=False)
        .size()
        .reset_index(name="Count")
    )

    distributor_order = (
        summary.groupby("Distributor name", as_index=False)["Count"]
        .sum()
        .sort_values("Count", ascending=False)["Distributor name"]
        .tolist()
    )
    status_order = (
        summary.groupby("Status for chart", as_index=False)["Count"]
        .sum()
        .sort_values("Count", ascending=False)["Status for chart"]
        .tolist()
    )

    color_sequence = px.colors.qualitative.Set2 + px.colors.qualitative.Bold + px.colors.qualitative.Safe
    color_map = {status: color_sequence[i % len(color_sequence)] for i, status in enumerate(status_order)}

    fig = px.bar(
        summary,
        y="Distributor name",
        x="Count",
        color="Status for chart",
        orientation="h",
        barmode="stack",
        title=f"Estado por distribuidor | {selected_model}",
        custom_data=["Status for chart", "Count"],
        color_discrete_map=color_map,
        category_orders={"Distributor name": distributor_order, "Status for chart": status_order},
    )

    fig.update_traces(
        hovertemplate=(
            "<b>Distribuidor:</b> %{y}<br>"
            "<b>Modelo:</b> " + selected_model + "<br>"
            "<b>Estado:</b> %{customdata[0]}<br>"
            "<b>Cantidad:</b> %{customdata[1]}<extra></extra>"
        )
    )

    fig.update_layout(
        xaxis_title="Cantidad de instrumentos",
        yaxis_title="Distribuidor",
        legend_title="Estado operativo",
    )
    return glow_layout(fig, 620, 17)


st.markdown(
    """
    <div class="hero">
        <div class="hero-top">
            <div class="hero-brand">
                <div class="brand-chip">DASHBOARD</div>
                <div class="workspace-chip">Hi, Javier · Workspace de base instalada</div>
            </div>
            <div class="workspace-chip">Control visual · Devoryn dark mode</div>
        </div>
        <h1>Records List Intelligence Dashboard</h1>
        <p>Panel ejecutivo para explorar la base instalada, configuration insights, sistema operativo, procesamiento y gap de repuestos con una apariencia oscura, limpia y premium.</p>
        <div class="badge-row">
            <span class="badge">Base instalada</span>
            <span class="badge">Machine configuration</span>
            <span class="badge">Operating system</span>
            <span class="badge">PM & processing</span>
            <span class="badge">Stock gap analysis</span>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

st.sidebar.markdown(
    """
    <div class="sidebar-top-card">
        <h3>✦ Control center</h3>
        <p>Explora la base instalada, filtra la operación y navega el dashboard con una experiencia visual alineada al estilo oscuro premium que definiste.</p>
        <div class="sidebar-pill">Devoryn dark · active</div>
    </div>
    """,
    unsafe_allow_html=True,
)
uploaded_file = st.sidebar.file_uploader("Sube el archivo Records List", type=["csv", "xlsx", "xls"])

base_dir = Path(__file__).resolve().parent
sample_candidates = sorted(base_dir.glob("Records_List_Report*.csv"))
default_master_candidates = sorted(base_dir.glob("New TP Spare*.xlsx"))

raw_df, source_label = get_active_records_dataset(uploaded_file, sample_candidates)

if raw_df.empty:
    st.info("Sube el archivo Records List para activar el dashboard.")
    st.stop()

raw_df, CONFIG_KEYS = parse_machine_configuration(raw_df)
raw_df = add_operating_system_columns(raw_df, CONFIG_KEYS)
st.sidebar.caption(f"Fuente activa: {source_label}")
st.sidebar.markdown('<div class="small-note">Usa los filtros como un panel de control para refinar región, país, distribuidor, instrumento y estado operativo.</div>', unsafe_allow_html=True)

region_options = sorted(raw_df["Commercial Region"].dropna().unique().tolist())
selected_regions = st.sidebar.multiselect("Región comercial", options=region_options, default=[], placeholder="Selecciona una o varias regiones")

country_base = raw_df.copy()
if selected_regions:
    country_base = country_base[country_base["Commercial Region"].isin(selected_regions)]
country_options = sorted(country_base["Country"].dropna().unique().tolist())
selected_countries = st.sidebar.multiselect("País", options=country_options, default=[], placeholder="Selecciona uno o varios países")

dist_base = raw_df.copy()
if selected_regions:
    dist_base = dist_base[dist_base["Commercial Region"].isin(selected_regions)]
if selected_countries:
    dist_base = dist_base[dist_base["Country"].isin(selected_countries)]
distributor_options = sorted(dist_base["Distributor name"].dropna().unique().tolist())
selected_distributors = st.sidebar.multiselect("Nombre de distribuidor", options=distributor_options, default=[], placeholder="Selecciona uno o varios distribuidores")

instrument_base = raw_df.copy()
if selected_regions:
    instrument_base = instrument_base[instrument_base["Commercial Region"].isin(selected_regions)]
if selected_countries:
    instrument_base = instrument_base[instrument_base["Country"].isin(selected_countries)]
if selected_distributors:
    instrument_base = instrument_base[instrument_base["Distributor name"].isin(selected_distributors)]
instrument_options = sorted(instrument_base["Instrument type"].dropna().unique().tolist())
selected_instruments = st.sidebar.multiselect("Tipo de instrumento", options=instrument_options, default=[], placeholder="Selecciona uno o varios instrumentos")

status_base = raw_df.copy()
if selected_regions:
    status_base = status_base[status_base["Commercial Region"].isin(selected_regions)]
if selected_countries:
    status_base = status_base[status_base["Country"].isin(selected_countries)]
if selected_distributors:
    status_base = status_base[status_base["Distributor name"].isin(selected_distributors)]
if selected_instruments:
    status_base = status_base[status_base["Instrument type"].isin(selected_instruments)]

state_count_items = compute_state_filter_counts(status_base)
state_option_map = {f"{state} ({count})": state for state, count in state_count_items}
selected_state_labels = st.sidebar.multiselect(
    "Estado operativo",
    options=list(state_option_map.keys()),
    default=[],
    placeholder="Selecciona uno o varios estados",
    help="Incluye el estado especial 'No rutina' y cualquier otro estado disponible en la vista actual.",
)
selected_states = [state_option_map[label] for label in selected_state_labels]

filtered = raw_df.copy()
if selected_regions:
    filtered = filtered[filtered["Commercial Region"].isin(selected_regions)]
if selected_countries:
    filtered = filtered[filtered["Country"].isin(selected_countries)]
if selected_distributors:
    filtered = filtered[filtered["Distributor name"].isin(selected_distributors)]
if selected_instruments:
    filtered = filtered[filtered["Instrument type"].isin(selected_instruments)]
filtered = apply_operational_status_filter(filtered, selected_states)

if filtered.empty:
    st.warning("No hay datos para la combinación de filtros actual.")
    st.stop()

st.sidebar.markdown("---")
base_tab, machine_tab, os_tab, process_tab, stock_tab, detail_tab = st.tabs(
    ["Base instalada", "Machine configuration", "Sistema operativo", "Procesamiento / PM", "Stock / Carstock gap", "Detalle por equipo"]
)

with base_tab:
    st.subheader("Base instalada")
    st.caption("Mapa y analítica de base instalada con enfoque en cobertura geográfica, antigüedad de instalación y estado de despliegue.")
    geo_df = filtered.dropna(subset=["Latitude", "Longitude"]).copy()
    if geo_df.empty:
        st.info("No hay coordenadas válidas para mostrar en el mapa.")
    else:
        st.markdown('<div class="map-shell">', unsafe_allow_html=True)
        st.markdown('<div class="section-title">Vista global de la base instalada</div>', unsafe_allow_html=True)
        st.markdown('<div class="map-note">La proyección se muestra completa desde la carga inicial para conservar el efecto ovalado y el contraste con el fondo glass.</div>', unsafe_allow_html=True)

        fig_geo = px.scatter_geo(
            geo_df,
            lat="Latitude",
            lon="Longitude",
            hover_name="Customer name",
            hover_data={
                "Serial number": True,
                "Instrument type": True,
                "Country": True,
                "Distributor name": True,
                "Operational status": True,
                "Commercial Region": True,
                "Latitude": False,
                "Longitude": False,
            },
            height=560,
            projection="mollweide",
        )
        fig_geo.update_traces(
            marker=dict(
                size=7.0,
                color=ACCENT,
                opacity=0.98,
                line=dict(color="rgba(255,255,255,0.96)", width=1.25),
            ),
            hovertemplate=(
                "<b>%{hovertext}</b><br>"
                "Serie: %{customdata[0]}<br>"
                "Instrumento: %{customdata[1]}<br>"
                "País: %{customdata[2]}<br>"
                "Distribuidor: %{customdata[3]}<br>"
                "Estado: %{customdata[4]}<br>"
                "Región comercial: %{customdata[5]}<extra></extra>"
            ),
        )
        fig_geo.update_geos(
            projection_type="mollweide",
            projection_scale=0.92,
            center=dict(lat=8, lon=0),
            showframe=False,
            bgcolor="rgba(255,255,255,0)",
            showocean=True,
            oceancolor="rgba(14,28,46,0.18)",
            showland=True,
            landcolor="rgba(255,255,255,0.14)",
            showcountries=True,
            countrycolor="rgba(190,235,255,0.28)",
            countrywidth=0.7,
            showcoastlines=True,
            coastlinecolor="rgba(190,235,255,0.22)",
            coastlinewidth=0.7,
            showlakes=True,
            lakecolor="rgba(30,52,80,0.12)",
            lataxis_showgrid=True,
            lonaxis_showgrid=True,
            lataxis_gridcolor="rgba(190,235,255,0.06)",
            lonaxis_gridcolor="rgba(190,235,255,0.06)",
            lataxis_dtick=15,
            lonaxis_dtick=30,
            domain=dict(x=[0.10, 0.90], y=[0.14, 0.86]),
        )
        fig_geo.update_layout(
            paper_bgcolor=PLOT_BG,
            plot_bgcolor=PLOT_BG,
            margin=dict(l=0, r=0, t=0, b=0),
            font=dict(color=TEXT),
        )
        st.plotly_chart(fig_geo, use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)

    c1, c2 = st.columns(2)
    with c1:
        type_df = filtered["Instrument type"].fillna("Unknown").value_counts().reset_index()
        type_df.columns = ["Instrument type", "Count"]
        fig_type = px.bar(type_df, x="Count", y="Instrument type", orientation="h", title="Base instalada por tipo de instrumento", text="Count")
        fig_type.update_traces(marker_color=ACCENT, textposition="outside", hovertemplate="Instrumento: %{y}<br>Activos: %{x}<extra></extra>")
        fig_type.update_layout(yaxis=dict(categoryorder="total ascending"))
        st.plotly_chart(glow_layout(fig_type, 470), use_container_width=True)

    with c2:
        install_df = filtered.dropna(subset=["Installation date"]).copy()
        if install_df.empty:
            st.info("No hay fechas de instalación válidas para el filtro actual.")
        else:
            install_df["Installation year"] = install_df["Installation date"].dt.year.astype(int)
            yearly = install_df.groupby("Installation year", dropna=False).size().reset_index(name="Count").sort_values("Installation year")
            fig_year = px.bar(yearly, x="Installation year", y="Count", title="Instalaciones por año", text="Count")
            fig_year.update_traces(marker_color=ACCENT_2, textposition="outside", hovertemplate="Año: %{x}<br>Instalaciones: %{y}<extra></extra>")
            st.plotly_chart(glow_layout(fig_year, 470), use_container_width=True)

    c3, c4 = st.columns(2)
    with c3:
        status_series = filtered["Operational status"].fillna("No informado").astype(str)
        status_class = pd.Series("Installed / Active", index=filtered.index)
        status_class = status_class.mask(status_series.str.contains("ready", case=False, na=False), "Ready to install")
        status_class = status_class.mask(status_series.str.contains("transit|customs|shipping", case=False, na=False), "Transit / Customs")
        status_class = status_class.mask(status_series.str.contains("warehouse|stock", case=False, na=False), "Warehouse")
        pipeline_df = filtered.copy()
        pipeline_df["Installation stage"] = status_class
        pipeline_summary = pipeline_df.groupby(["Instrument type", "Installation stage"], dropna=False).size().reset_index(name="Count")
        fig_ready = px.bar(pipeline_summary, x="Instrument type", y="Count", color="Installation stage", title="Sistemas instalados vs listos / pipeline")
        fig_ready.update_traces(hovertemplate="Instrumento: %{x}<br>Etapa: %{fullData.name}<br>Cantidad: %{y}<extra></extra>")
        fig_ready.update_xaxes(tickangle=-28)
        st.plotly_chart(glow_layout(fig_ready, 470), use_container_width=True)

    with c4:
        city_df = (
            filtered.assign(CityLabel=filtered["City"].fillna("No informado") + " | " + filtered["Country"].fillna("No country"))
            .groupby("CityLabel", dropna=False)
            .size()
            .reset_index(name="Count")
            .sort_values("Count", ascending=False)
            .head(15)
        )
        fig_city = px.bar(city_df, x="Count", y="CityLabel", orientation="h", title="Análisis por ciudad", text="Count")
        fig_city.update_traces(marker_color=ACCENT_3, textposition="outside", hovertemplate="Ciudad / País: %{y}<br>Activos: %{x}<extra></extra>")
        fig_city.update_layout(yaxis=dict(categoryorder="total ascending"))
        st.plotly_chart(glow_layout(fig_city, 470), use_container_width=True)

    st.markdown("### Vista corporativa por distribuidor")
    model_options = sorted(filtered["Instrument type"].dropna().astype(str).unique().tolist())
    if model_options:
        default_model = model_options[0]
        selected_model_for_status = st.selectbox(
            "Selecciona el modelo de equipo",
            options=model_options,
            index=0,
            key="corporate_distributor_model_selector",
        )
        st.markdown(
            '<div class="small-note">Visual principal corporativo: un modelo a la vez, todos los distribuidores, estados apilados horizontalmente.</div>',
            unsafe_allow_html=True,
        )
        st.plotly_chart(build_distributor_status_chart(filtered, selected_model_for_status), use_container_width=True)

    st.markdown("### Tabla general filtrada")
    visible_columns = [
        "Commercial Region",
        "Country",
        "Distributor name",
        "Customer name",
        "Instrument type",
        "Serial number",
        "Operational status grouped",
        "Operational status",
        "Asset condition",
        "Installation date",
        "Number of tests per day",
        "Operating System",
        "Machine Configurations",
    ]
    st.dataframe(filtered[visible_columns].copy(), use_container_width=True, hide_index=True)

with machine_tab:
    st.subheader("Machine configuration")
    st.caption("Vista ejecutiva por ítem de configuración, con gráficas separadas para cada campo aplicable y mayor lectura visual del comportamiento de la base instalada.")
    applicable_fields = active_config_fields(filtered, CONFIG_KEYS)
    cfg_cols_prefixed = [f"CFG::{col}" for col in applicable_fields]

    if not cfg_cols_prefixed:
        st.info("No se detectaron campos aplicables dentro de Machine Configurations para el filtro actual.")
    else:
        assets_with_cfg = int(filtered["Machine Configurations"].notna().sum())
        avg_cfg_fields = filtered["Machine config fields populated"].mean()
        unique_cfg_fields = len(applicable_fields)

        mc1, mc2, mc3 = st.columns(3)
        with mc1:
            metric_card("Equipos con config", f"{assets_with_cfg:,}", "Con información en Machine Configurations")
        with mc2:
            metric_card("Campos aplicables", f"{unique_cfg_fields}", "Solo ítems presentes en el filtro actual")
        with mc3:
            metric_card("Promedio de campos", f"{avg_cfg_fields:.1f}", "Campos poblados por equipo")

        coverage_df = pd.DataFrame(
            [{"Config field": col.replace("CFG::", ""), "Populated assets": int(filtered[col].notna().sum())} for col in cfg_cols_prefixed]
        )
        coverage_df = coverage_df[coverage_df["Populated assets"] > 0].sort_values("Populated assets", ascending=False)

        fig_cfg_fill = px.bar(
            coverage_df,
            x="Populated assets",
            y="Config field",
            orientation="h",
            title="Cobertura por campo aplicable",
            text="Populated assets",
        )
        fig_cfg_fill.update_traces(
            marker_color=ACCENT,
            textposition="outside",
            hovertemplate="Campo: %{y}<br>Equipos con dato: %{x}<extra></extra>",
        )
        fig_cfg_fill.update_layout(yaxis=dict(categoryorder="total ascending"))
        st.plotly_chart(glow_layout(fig_cfg_fill, 520), use_container_width=True)

        st.markdown("### Distribución visual por ítem")
        st.markdown(
            '<div class="small-note">Cada gráfico resume la distribución de valores del ítem correspondiente. Se muestran los valores principales y, si aplica, una categoría <b>Otros</b> para simplificar la lectura.</div>',
            unsafe_allow_html=True,
        )

        donut_fields = coverage_df["Config field"].tolist()
        if donut_fields:
            for idx in range(0, len(donut_fields), 3):
                cols = st.columns(3)
                for col_ui, field_name in zip(cols, donut_fields[idx:idx + 3]):
                    selected_cfg_col = f"CFG::{field_name}"
                    item_series = filtered[selected_cfg_col].dropna()
                    item_series = item_series.astype(str).str.strip()
                    item_series = item_series[item_series.ne("")]
                    total_assets = int(item_series.shape[0])
                    with col_ui:
                        st.plotly_chart(build_config_donut(field_name, item_series, total_assets), use_container_width=True)

        st.markdown("### Top valores por ítem")
        detail_rows = []
        for field_name in donut_fields:
            selected_cfg_col = f"CFG::{field_name}"
            item_series = filtered[selected_cfg_col].dropna().astype(str).str.strip()
            item_series = item_series[item_series.ne("")]
            if item_series.empty:
                continue
            dist = item_series.value_counts().reset_index()
            dist.columns = ["Value", "Count"]
            top_row = dist.iloc[0]
            detail_rows.append(
                {
                    "Config field": field_name,
                    "Top value": str(top_row["Value"]),
                    "Top count": int(top_row["Count"]),
                    "Unique values": int(dist.shape[0]),
                    "Assets with value": int(item_series.shape[0]),
                }
            )

        if detail_rows:
            st.dataframe(pd.DataFrame(detail_rows), use_container_width=True, hide_index=True)

        with st.expander("Ver tabla ampliada de machine configuration"):
            detail_columns = [
                "Commercial Region",
                "Country",
                "Distributor name",
                "Customer name",
                "Instrument type",
                "Serial number",
                "Operating System",
                "Operational status",
            ] + [f"CFG::{field}" for field in donut_fields]
            machine_table = filtered[detail_columns].copy().rename(columns={f"CFG::{field}": field for field in donut_fields})
            st.dataframe(machine_table, use_container_width=True, hide_index=True)

with os_tab:
    st.subheader("Sistema operativo")
    st.caption("Vista diseñada para identificar instrumentos con sistemas operativos legacy y priorizar migraciones urgentes a Windows 10.")
    os_df = filtered.copy()
    os_df["Operating System"] = os_df["Operating System"].fillna("No informado")
    os_df["OS Upgrade Bucket"] = os_df["Operating System"].map(os_upgrade_bucket)
    os_df["Installation date display"] = os_df["Installation date"].map(format_date_for_hover)

    systems_with_os = int(filtered["Operating System"].notna().sum())
    unique_os = int(filtered["Operating System"].nunique(dropna=True))
    unknown_os = int(os_df["Operating System"].isin(["Unknown", "No informado"]).sum())
    urgent_mask = os_df["Operating System"].isin(["Windows XP", "Windows Vista", "Windows 7", "Windows 2000"])
    urgent_count = int(urgent_mask.sum())

    o1, o2, o3, o4 = st.columns(4)
    with o1:
        metric_card("Equipos con OS", f"{systems_with_os:,}", "OS identificado desde machine configuration")
    with o2:
        metric_card("OS únicos", f"{unique_os}", "Variedad de sistemas operativos")
    with o3:
        metric_card("Legacy / urgentes", f"{urgent_count:,}", "XP, Vista, Win7 o Win2000")
    with o4:
        metric_card("Sin definir / unknown", f"{unknown_os:,}", "Requiere depuración del dato")

    s1, s2 = st.columns(2)
    with s1:
        os_summary = os_df.groupby("Operating System", dropna=False).size().reset_index(name="Count").sort_values("Count", ascending=False)
        fig_os = px.bar(os_summary, x="Operating System", y="Count", title="Distribución detallada de sistemas operativos", text="Count")
        fig_os.update_traces(
            marker_color=ACCENT,
            textposition="outside",
            hovertemplate="Sistema operativo: %{x}<br>Equipos: %{y}<extra></extra>",
        )
        fig_os.update_xaxes(tickangle=-28)
        st.plotly_chart(glow_layout(fig_os, 500, title_size=16), use_container_width=True)

    with s2:
        os_points = os_df[[
            "Serial number",
            "Instrument type",
            "Operating System",
            "OS Upgrade Bucket",
            "Distributor name",
            "Customer name",
            "Country",
            "Commercial Region",
            "Operational status",
            "Installation date display",
        ]].copy()
        fig_os_type = px.scatter(
            os_points,
            x="Operating System",
            y="Serial number",
            color="OS Upgrade Bucket",
            title="Qué seriales tienen cada sistema operativo",
            custom_data=[
                "Instrument type",
                "Distributor name",
                "Customer name",
                "Country",
                "Commercial Region",
                "Operational status",
                "Installation date display",
                "OS Upgrade Bucket",
            ],
            category_orders={"Operating System": os_summary["Operating System"].tolist()},
            color_discrete_map={
                "Windows 10 / OK": ACCENT_3,
                "Legacy / urgente migrar": DANGER,
                "Otro OS / validar": WARNING,
                "Revisar campo OS": ACCENT_2,
            },
        )
        fig_os_type.update_traces(
            marker=dict(size=11, opacity=0.88),
            hovertemplate=(
                "Serial: %{y}<br>"
                "OS: %{x}<br>"
                "Instrumento: %{customdata[0]}<br>"
                "Distribuidor: %{customdata[1]}<br>"
                "Cliente: %{customdata[2]}<br>"
                "País: %{customdata[3]}<br>"
                "Región: %{customdata[4]}<br>"
                "Estado operativo: %{customdata[5]}<br>"
                "Instalación: %{customdata[6]}<br>"
                "Prioridad: %{customdata[7]}<extra></extra>"
            ),
        )
        fig_os_type.update_layout(legend_title_text="Prioridad upgrade")
        fig_os_type.update_xaxes(tickangle=-28)
        st.plotly_chart(glow_layout(fig_os_type, 620, title_size=16), use_container_width=True)

    s3, s4 = st.columns(2)
    with s3:
        urgent_points = os_df[urgent_mask][[
            "Serial number",
            "Instrument type",
            "Operating System",
            "Distributor name",
            "Customer name",
            "Country",
            "Commercial Region",
            "Operational status",
            "Installation date display",
        ]].copy()
        if urgent_points.empty:
            st.success("No se detectan instrumentos con Windows legacy dentro del filtro actual.")
        else:
            urgent_points = urgent_points.sort_values(["Country", "Instrument type", "Serial number"])
            fig_urgent = px.scatter(
                urgent_points,
                x="Country",
                y="Serial number",
                color="Instrument type",
                title="Seriales que requieren actualización urgente a Windows 10",
                custom_data=[
                    "Operating System",
                    "Instrument type",
                    "Distributor name",
                    "Customer name",
                    "Commercial Region",
                    "Operational status",
                    "Installation date display",
                ],
            )
            fig_urgent.update_traces(
                marker=dict(size=12, opacity=0.92),
                hovertemplate=(
                    "Serial: %{y}<br>"
                    "País: %{x}<br>"
                    "OS actual: %{customdata[0]}<br>"
                    "Instrumento: %{customdata[1]}<br>"
                    "Distribuidor: %{customdata[2]}<br>"
                    "Cliente: %{customdata[3]}<br>"
                    "Región: %{customdata[4]}<br>"
                    "Estado operativo: %{customdata[5]}<br>"
                    "Instalación: %{customdata[6]}<extra></extra>"
                ),
            )
            fig_urgent.update_layout(legend_title_text="Instrumento")
            fig_urgent.update_xaxes(tickangle=-18)
            st.plotly_chart(glow_layout(fig_urgent, 620, title_size=16), use_container_width=True)

    with s4:
        bucket_df = os_df.groupby("OS Upgrade Bucket", dropna=False).size().reset_index(name="Count")
        order = ["Windows 10 / OK", "Legacy / urgente migrar", "Otro OS / validar", "Revisar campo OS"]
        bucket_df["order"] = bucket_df["OS Upgrade Bucket"].map({k: i for i, k in enumerate(order)}).fillna(999)
        bucket_df = bucket_df.sort_values(["order", "Count"], ascending=[True, False])
        fig_bucket = px.bar(bucket_df, x="OS Upgrade Bucket", y="Count", title="Priorización de acción para upgrade", text="Count")
        fig_bucket.update_traces(marker_color=ACCENT_2, textposition="outside", hovertemplate="Acción: %{x}<br>Equipos: %{y}<extra></extra>")
        fig_bucket.update_xaxes(tickangle=-18)
        st.plotly_chart(glow_layout(fig_bucket, 520, title_size=16), use_container_width=True)

    st.markdown("### Tabla priorizada para migración a Windows 10")
    urgent_table = os_df[urgent_mask][[
        "Commercial Region", "Country", "Distributor name", "Customer name", "Instrument type", "Serial number", "Operating System", "Operational status", "Installation date display"
    ]].copy()
    if urgent_table.empty:
        st.info("No hay equipos en categoría urgente para el filtro actual.")
    else:
        urgent_table = urgent_table.rename(columns={"Installation date display": "Installation date"})
        st.dataframe(urgent_table, use_container_width=True, hide_index=True)

    st.markdown("### Tabla de soporte OS")
    os_table = filtered[[
        "Commercial Region",
        "Country",
        "Distributor name",
        "Customer name",
        "Instrument type",
        "Serial number",
        "Operating System",
        "Machine Configurations",
    ]].copy()
    st.dataframe(os_table, use_container_width=True, hide_index=True)

with process_tab:
    st.subheader("Procesamiento, product line y PM planner")
    st.caption("Nueva pestaña para revisar volumen de procesamiento por serie, líneas de producto activas y planeación de mantenimiento preventivo.")

    proc_df = filtered.copy()
    proc_df["Number of tests per day"] = pd.to_numeric(proc_df["Number of tests per day"], errors="coerce")
    proc_df["PM next date display"] = proc_df["PM next date"].map(format_date_for_hover)
    proc_df["PM last date display"] = proc_df["PM last date"].map(format_date_for_hover)
    proc_df["Tests/day display"] = proc_df["Number of tests per day"].map(lambda x: safe_number_text(x, "0"))

    p1, p2, p3, p4 = st.columns(4)
    tests_valid = proc_df["Number of tests per day"].dropna()
    pm_ready = int(proc_df["PM plan"].notna().sum())
    upcoming_pm = int(proc_df["PM next date"].between(pd.Timestamp.today().normalize(), pd.Timestamp.today().normalize() + pd.Timedelta(days=90), inclusive="both").sum()) if proc_df["PM next date"].notna().any() else 0
    product_lines_count = int(proc_df["Product Line"].fillna("").astype(str).str.strip().replace("", pd.NA).dropna().nunique())
    with p1:
        metric_card("Tests/día promedio", safe_number_text(tests_valid.mean() if not tests_valid.empty else pd.NA, "0"), "Promedio del filtro")
    with p2:
        metric_card("Tests/día máximos", safe_number_text(tests_valid.max() if not tests_valid.empty else pd.NA, "0"), "Serie con mayor procesamiento")
    with p3:
        metric_card("Product lines", f"{product_lines_count}", "Líneas de producto detectadas")
    with p4:
        metric_card("PM próximos 90 días", f"{upcoming_pm:,}", f"{pm_ready:,} equipos con PM plan")

    g1, g2 = st.columns(2)
    with g1:
        tests_df = proc_df.dropna(subset=["Number of tests per day", "Serial number"]).copy()
        if tests_df.empty:
            st.info("No hay datos válidos en Number of tests per day para el filtro actual.")
        else:
            tests_df = tests_df.sort_values(["Number of tests per day", "Serial number"], ascending=[False, True]).reset_index(drop=True)
            fig_tests = px.scatter(
                tests_df,
                x="Serial number",
                y="Number of tests per day",
                color="Instrument type",
                title="Number of tests/day por cada serie",
                custom_data=["Customer name", "Distributor name", "Country", "Commercial Region", "Product Line", "Operational status"],
            )
            fig_tests.update_traces(
                marker=dict(size=10, opacity=0.9),
                hovertemplate=(
                    "Serie: %{x}<br>"
                    "Tests/día: %{y}<br>"
                    "Instrumento: %{fullData.name}<br>"
                    "Cliente: %{customdata[0]}<br>"
                    "Distribuidor: %{customdata[1]}<br>"
                    "País: %{customdata[2]}<br>"
                    "Región: %{customdata[3]}<br>"
                    "Product line: %{customdata[4]}<br>"
                    "Estado: %{customdata[5]}<extra></extra>"
                )
            )
            fig_tests.update_xaxes(tickangle=-60)
            st.plotly_chart(glow_layout(fig_tests, 520), use_container_width=True)

    with g2:
        product_series = proc_df["Product Line"].fillna("").astype(str).str.strip()
        product_rows = []
        for value in product_series:
            if not value:
                continue
            parts = [p.strip() for p in re.split(r"[\|;,/]", value) if p.strip()]
            if not parts:
                parts = [value]
            product_rows.extend(parts)
        if not product_rows:
            st.info("No hay datos válidos en Product Line para el filtro actual.")
        else:
            product_line_df = pd.Series(product_rows, name="Product line").value_counts().reset_index()
            product_line_df.columns = ["Product line", "Count"]
            product_line_df = product_line_df.head(20)
            fig_product = px.bar(
                product_line_df,
                x="Count",
                y="Product line",
                orientation="h",
                title="Product line performed on the analyzer",
                text="Count",
            )
            fig_product.update_traces(
                marker_color=ACCENT_2,
                textposition="outside",
                hovertemplate="Product line: %{y}<br>Equipos / apariciones: %{x}<extra></extra>",
            )
            fig_product.update_layout(yaxis=dict(categoryorder="total ascending"))
            st.plotly_chart(glow_layout(fig_product, 520), use_container_width=True)

    g3, g4 = st.columns(2)
    with g3:
        pm_plan_df = proc_df.copy()
        pm_plan_df["PM Plan label"] = pm_plan_df["PM plan"].fillna("No informado").astype(str)
        pm_summary = pm_plan_df.groupby("PM Plan label", dropna=False).size().reset_index(name="Count").sort_values("Count", ascending=False)
        fig_pm_plan = px.bar(
            pm_summary,
            x="PM Plan label",
            y="Count",
            title="PM planner | distribución de PM plan",
            text="Count",
        )
        fig_pm_plan.update_traces(
            marker_color=ACCENT_3,
            textposition="outside",
            hovertemplate="PM plan: %{x}<br>Equipos: %{y}<extra></extra>",
        )
        fig_pm_plan.update_xaxes(tickangle=-28)
        st.plotly_chart(glow_layout(fig_pm_plan, 500), use_container_width=True)

    with g4:
        pm_timeline = proc_df.dropna(subset=["PM next date", "Serial number"]).copy()
        if pm_timeline.empty:
            st.info("No hay fechas válidas en PM next date para el filtro actual.")
        else:
            today = pd.Timestamp.today().normalize()
            pm_timeline["PM planner status"] = np.where(
                pm_timeline["PM next date"] < today,
                "Overdue",
                np.where(pm_timeline["PM next date"] <= today + pd.Timedelta(days=90), "Next 90 days", "Planned later"),
            )
            fig_pm_timeline = px.scatter(
                pm_timeline.sort_values("PM next date"),
                x="PM next date",
                y="Serial number",
                color="PM planner status",
                title="PM planner | calendario por serie",
                custom_data=["Instrument type", "Customer name", "Distributor name", "Country", "PM plan", "PM frequency", "PM performed On", "PM last date display", "PM next date display"],
                color_discrete_map={"Overdue": DANGER, "Next 90 days": WARNING, "Planned later": ACCENT},
            )
            fig_pm_timeline.update_traces(
                marker=dict(size=11, opacity=0.9),
                hovertemplate=(
                    "Serie: %{y}<br>"
                    "PM next date: %{customdata[8]}<br>"
                    "Instrumento: %{customdata[0]}<br>"
                    "Cliente: %{customdata[1]}<br>"
                    "Distribuidor: %{customdata[2]}<br>"
                    "País: %{customdata[3]}<br>"
                    "PM plan: %{customdata[4]}<br>"
                    "PM frequency: %{customdata[5]}<br>"
                    "PM performed on: %{customdata[6]}<br>"
                    "PM last date: %{customdata[7]}<br>"
                    "Estado planner: %{fullData.name}<extra></extra>"
                )
            )
            st.plotly_chart(glow_layout(fig_pm_timeline, 500), use_container_width=True)

    st.markdown("### Tabla de soporte para procesamiento / PM")
    process_table_cols = [
        "Commercial Region",
        "Country",
        "Distributor name",
        "Customer name",
        "Instrument type",
        "Serial number",
        "Number of tests per day",
        "Product Line",
        "PM plan",
        "PM frequency",
        "PM performed On",
        "PM last date",
        "PM next date",
    ]
    st.dataframe(proc_df[process_table_cols].copy(), use_container_width=True, hide_index=True)

with stock_tab:
    st.subheader("Gap analysis de stock vs carstock requerido")
    st.session_state.setdefault("pdf_stock_context", {"available": False})
    st.caption(
        "Sube el maestro de referencia y luego el archivo trimestral del distribuidor. El dashboard intentará identificar automáticamente el distribuidor a partir del nombre del archivo y hará el análisis de brecha sin guardar histórico."
    )

    default_master_bytes = default_master_candidates[0].read_bytes() if default_master_candidates else None
    default_master_name = default_master_candidates[0].name if default_master_candidates else "No cargado"

    master_upload = st.file_uploader(
        "Archivo maestro de carstock (consolidado o New TP Spare)",
        type=["xlsx", "xls", "csv"],
        key="master_spare_upload",
    )
    stock_upload = st.file_uploader(
        "Archivo de stock reportado por el distribuidor",
        type=["xlsx", "xls", "csv"],
        key="distributor_stock_upload",
    )

    if master_upload is not None:
        master_bytes = master_upload.getvalue()
        master_name = master_upload.name
    elif default_master_bytes is not None:
        master_bytes = default_master_bytes
        master_name = default_master_name
    else:
        master_bytes = None
        master_name = None

    if master_bytes is None:
        st.session_state["pdf_stock_context"] = {"available": False}
        st.info("Sube el archivo maestro de carstock para activar esta pestaña.")
    else:
        master_bundle = load_carstock_master_bundle(master_bytes, master_name)
        has_consolidated = not master_bundle["consolidated"].empty
        has_legacy = bool(master_bundle["legacy_families"])

        if not has_consolidated and not has_legacy:
            st.session_state["pdf_stock_context"] = {"available": False}
            st.warning("No se pudo interpretar el archivo maestro. Debe contener al menos part number y quantity.")
        else:
            st.markdown(f"**Archivo maestro activo:** {master_name}")
            st.markdown(
                "<div class='small-note'>Lógica activa: detección automática del distribuidor desde el nombre del archivo, inferencia de familias según su base instalada y comparación inmediata contra el carstock requerido.</div>",
                unsafe_allow_html=True,
            )

            if stock_upload is None:
                st.session_state["pdf_stock_context"] = {"available": False}
                st.info("Sube ahora el archivo trimestral del distribuidor. Ejemplo recomendado: `ANNAR_stock_Q1_2026.xlsx`.")
            else:
                stock_df_raw = load_table_file(stock_upload.getvalue(), stock_upload.name)
                candidate_distributors = []
                if stock_df_raw is None or stock_df_raw.empty:
                    st.session_state["pdf_stock_context"] = {"available": False}
                    st.warning("El archivo subido no contiene datos legibles.")
                else:
                    part_col_guess, qty_col_guess, desc_col_guess = detect_stock_columns(stock_df_raw)
                    master_distributors = sorted(set(raw_df["Distributor name"].dropna().tolist()) | set(master_bundle["master_distributors"]))
                    detected_distributor, candidate_distributors = infer_distributor_from_filename_strict(stock_upload.name, master_distributors)

                    if detected_distributor is None:
                        st.warning("No fue posible identificar de forma única el distribuidor desde el nombre del archivo.")
                        if candidate_distributors:
                            detected_distributor = st.selectbox(
                                "Selecciona manualmente el distribuidor",
                                options=sorted(candidate_distributors),
                                key="manual_distributor_selection_candidates",
                            )
                        else:
                            detected_distributor = st.selectbox(
                                "Selecciona manualmente el distribuidor",
                                options=sorted(master_distributors),
                                key="manual_distributor_selection_all",
                            )

                    if not detected_distributor:
                        st.session_state["pdf_stock_context"] = {"available": False}
                        st.error("No pude identificar el distribuidor desde el nombre del archivo.")
                        st.caption("Renombra el archivo con un formato claro, por ejemplo: `ANNAR_stock_Q1_2026.xlsx`, `Biotec_del_Paraguay_stock.xlsx` o `Simed_Ecuador_carstock.xlsx`.")
                    else:
                        distributor_scope = raw_df[raw_df["Distributor name"].eq(detected_distributor)].copy()
                        distributor_inst = distributor_scope["Instrument type"].dropna().unique().tolist()
                        family_from_base = infer_families_from_instruments(distributor_inst)

                        if has_consolidated:
                            consolidated = master_bundle["consolidated"].copy()
                            dist_key = normalize_key_text(detected_distributor)
                            dist_specific = consolidated[consolidated["Distributor Key"].eq(dist_key)].copy()
                            if not dist_specific.empty:
                                dist_families = sorted([f for f in dist_specific["Required Family"].dropna().unique().tolist() if f])
                            else:
                                dist_families = []
                            available_families = master_bundle["master_families"] or sorted(consolidated["Required Family"].dropna().unique().tolist())
                        else:
                            dist_specific = pd.DataFrame()
                            dist_families = []
                            available_families = list(master_bundle["legacy_families"].keys())

                        auto_families = [fam for fam in family_from_base if fam in available_families]
                        if dist_families:
                            auto_families = [fam for fam in dist_families if fam in available_families] or auto_families
                        if not auto_families:
                            auto_families = available_families[:1]

                        s1, s2, s3 = st.columns(3)
                        with s1:
                            metric_card("Distribuidor detectado", detected_distributor, "Detectado automáticamente o seleccionado manualmente")
                        with s2:
                            metric_card("Familias inferidas", ", ".join(auto_families) if auto_families else "N/A", "Según base instalada")
                        with s3:
                            metric_card("Instrumentos del distribuidor", f"{len(distributor_inst):,}", "Tipos detectados en la base")

                        info_lines = []
                        if distributor_inst:
                            info_lines.append("Instrumentos en base: " + ", ".join(sorted(distributor_inst)))
                        if candidate_distributors and len(candidate_distributors) > 1:
                            info_lines.append("Coincidencias secundarias detectadas en el nombre: " + ", ".join(candidate_distributors[1:4]))
                        if info_lines:
                            st.caption(" | ".join(info_lines))

                        with st.expander("Ajustes avanzados de lectura del archivo", expanded=False):
                            col1, col2, col3 = st.columns(3)
                            with col1:
                                part_col = st.selectbox(
                                    "Columna de part number",
                                    options=stock_df_raw.columns.tolist(),
                                    index=stock_df_raw.columns.tolist().index(part_col_guess) if part_col_guess in stock_df_raw.columns else 0,
                                )
                            with col2:
                                qty_col = st.selectbox(
                                    "Columna de cantidad",
                                    options=stock_df_raw.columns.tolist(),
                                    index=stock_df_raw.columns.tolist().index(qty_col_guess) if qty_col_guess in stock_df_raw.columns else min(1, len(stock_df_raw.columns) - 1),
                                )
                            with col3:
                                desc_options = ["<sin descripción>"] + stock_df_raw.columns.tolist()
                                default_desc_index = desc_options.index(desc_col_guess) if desc_col_guess in stock_df_raw.columns else 0
                                desc_selection = st.selectbox("Columna de descripción", options=desc_options, index=default_desc_index)
                                desc_col = None if desc_selection == "<sin descripción>" else desc_selection

                            selected_families_stock = st.multiselect(
                                "Familias a comparar",
                                options=available_families,
                                default=auto_families,
                                key="stock_family_selector_auto",
                                placeholder="Familias detectadas automáticamente",
                            )

                        if 'part_col' not in locals():
                            part_col = part_col_guess
                            qty_col = qty_col_guess
                            desc_col = desc_col_guess
                            selected_families_stock = auto_families

                        if not selected_families_stock:
                            st.session_state["pdf_stock_context"] = {"available": False}
                            st.warning("No hay familias seleccionadas para comparar. Ajusta el maestro o la selección avanzada.")
                        else:
                            master_df, master_mode = build_required_master_from_scope(
                                master_bundle=master_bundle,
                                assigned_distributor=detected_distributor,
                                selected_families=selected_families_stock,
                            )

                            if master_df.empty and detected_distributor:
                                master_df, master_mode = build_required_master_from_scope(
                                    master_bundle=master_bundle,
                                    assigned_distributor="<sin asignar>",
                                    selected_families=selected_families_stock,
                                )

                            if master_df.empty:
                                st.session_state["pdf_stock_context"] = {"available": False}
                                st.warning("No encontré carstock requerido para este distribuidor con las familias inferidas. Revisa el maestro o el nombre del archivo.")
                            else:
                                comparison, extra_df, stock_slim = compare_stock(
                                    master_df,
                                    stock_df_raw,
                                    part_col,
                                    qty_col,
                                    desc_col,
                                    price_reference=master_bundle.get("price_reference", pd.DataFrame()),
                                )
                                missing_skus = int((comparison["Status"] == "Missing").sum())
                                low_skus = int((comparison["Status"] == "LOW").sum())
                                covered_skus = int((comparison["Status"] == "OK").sum())
                                total_gap = float(comparison["Qty Gap"].sum())
                                coverage = (covered_skus / len(comparison) * 100) if len(comparison) else 0
                                extra_skus = int(len(extra_df))
                                option2_cost = float(pd.to_numeric(comparison["Option 2 Estimated Cost"], errors="coerce").fillna(0).sum())
                                option2_currency = next((c for c in comparison["Currency"].dropna().astype(str).tolist() if c.strip()), "EUR")
                                purchase_df = comparison[comparison["Qty Gap"] > 0][[
                                    "Required Part Number",
                                    "Required Description",
                                    "Qty Gap",
                                    "Option 2 Unit Price",
                                    "Option 2 Estimated Cost",
                                    "Currency",
                                    "Status",
                                ]].sort_values(["Option 2 Estimated Cost", "Qty Gap"], ascending=[False, False])

                                st.session_state["pdf_stock_context"] = {
                                    "available": True,
                                    "detected_distributor": detected_distributor,
                                    "families": selected_families_stock,
                                    "required_skus": len(comparison),
                                    "ok_skus": covered_skus,
                                    "low_skus": low_skus,
                                    "missing_skus": missing_skus,
                                    "extra_skus": extra_skus,
                                    "gap_total": total_gap,
                                    "option2_cost": option2_cost,
                                    "currency": option2_currency,
                                    "top_gap_df": comparison[comparison["Qty Gap"] > 0].sort_values(["Qty Gap", "Required Part Number"], ascending=[False, True]).head(15).copy(),
                                    "full_comparison_df": comparison.copy(),
                                    "purchase_df": purchase_df.copy() if not purchase_df.empty else pd.DataFrame(columns=["Required Part Number", "Required Description", "Qty Gap", "Option 2 Unit Price", "Option 2 Estimated Cost", "Currency", "Status"]),
                                    "extra_df": extra_df.copy() if not extra_df.empty else pd.DataFrame(columns=["Uploaded Part Number", "Uploaded Description", "Uploaded Qty", "Status"]),
                                }

                                sm1, sm2, sm3, sm4, sm5, sm6, sm7 = st.columns(7)
                                with sm1:
                                    metric_card("SKUs requeridos", f"{len(comparison):,}", "Carstock esperado")
                                with sm2:
                                    metric_card("SKUs OK", f"{covered_skus:,}", f"{coverage:.1f}% del carstock")
                                with sm3:
                                    metric_card("SKUs LOW", f"{low_skus:,}", "Tienen stock pero insuficiente")
                                with sm4:
                                    metric_card("SKUs Missing", f"{missing_skus:,}", "Sin stock reportado")
                                with sm5:
                                    metric_card("Gap total qty", safe_number_text(total_gap, "0"), "Cantidad faltante acumulada")
                                with sm6:
                                    metric_card("Compra opción 2", f"{option2_currency} {option2_cost:,.2f}", "Costo estimado para cubrir gap")
                                with sm7:
                                    metric_card("Extras", f"{extra_skus:,}", "No incluidos en el maestro")

                                st.caption(
                                    f"Alcance automático: {detected_distributor} | Familias: {', '.join(selected_families_stock)} | Maestro: {'Consolidado' if master_mode == 'consolidated' else 'Estándar por familia'}"
                                )

                                g1, g2 = st.columns(2)
                                with g1:
                                    status_df = comparison["Status"].value_counts().reset_index()
                                    status_df.columns = ["Status", "Count"]
                                    color_map = {"OK": ACCENT_3, "LOW": WARNING, "Missing": DANGER}
                                    fig_status = px.pie(status_df, names="Status", values="Count", title="Cobertura del carstock", hole=0.52)
                                    fig_status.update_traces(
                                        marker=dict(colors=[color_map.get(s, ACCENT) for s in status_df["Status"]]),
                                        hovertemplate="Estado: %{label}<br>SKUs: %{value}<br>%{percent}<extra></extra>",
                                    )
                                    fig_status.update_layout(template=PLOT_TEMPLATE, paper_bgcolor=PLOT_BG, plot_bgcolor=PLOT_BG, font=dict(color=TEXT), height=430)
                                    st.plotly_chart(fig_status, use_container_width=True)

                                with g2:
                                    gap_df = comparison[comparison["Qty Gap"] > 0].sort_values(["Qty Gap", "Required Part Number"], ascending=[False, True]).head(15)
                                    if gap_df.empty:
                                        st.success("No se detectan faltantes para el distribuidor identificado.")
                                    else:
                                        fig_gap = px.bar(
                                            gap_df,
                                            x="Qty Gap",
                                            y="Required Part Number",
                                            orientation="h",
                                            title="Top faltantes por cantidad",
                                            custom_data=["Required Description", "Uploaded Qty", "Required Qty", "Coverage %", "Status", "Option 2 Unit Price", "Option 2 Estimated Cost", "Currency"],
                                            text="Qty Gap",
                                        )
                                        fig_gap.update_traces(
                                            marker_color=DANGER,
                                            textposition="outside",
                                            hovertemplate=(
                                                "Part number: %{y}<br>"
                                                "Descripción: %{customdata[0]}<br>"
                                                "Qty reportada: %{customdata[1]}<br>"
                                                "Qty requerida: %{customdata[2]}<br>"
                                                "Cobertura: %{customdata[3]}%<br>"
                                                "Estado: %{customdata[4]}<br>"
                                                "Precio opción 2: %{customdata[7]} %{customdata[5]:,.2f}<br>"
                                                "Costo compra opción 2: %{customdata[7]} %{customdata[6]:,.2f}<br>"
                                                "Gap: %{x}<extra></extra>"
                                            ),
                                        )
                                        fig_gap.update_layout(yaxis=dict(categoryorder="total ascending"))
                                        st.plotly_chart(glow_layout(fig_gap, 430), use_container_width=True)

                                st.markdown("### Tabla de brechas")
                                show_cols = ["Required Part Number", "Required Description", "Required Qty", "Uploaded Qty", "Qty Gap", "Coverage %", "Option 2 Unit Price", "Option 2 Estimated Cost", "Currency", "Status"]
                                st.dataframe(comparison[show_cols], use_container_width=True, hide_index=True)

                                if not purchase_df.empty:
                                    st.markdown("### Compra sugerida para cerrar el gap (opción 2)")
                                    st.dataframe(purchase_df, use_container_width=True, hide_index=True)
                                    st.markdown(f"**Total estimado de compra sugerida:** {option2_currency} {option2_cost:,.2f}")

                                if not extra_df.empty:
                                    st.markdown("### Partes reportadas que no están en el carstock requerido")
                                    st.dataframe(
                                        extra_df[["Uploaded Part Number", "Uploaded Description", "Uploaded Qty", "Status"]],
                                        use_container_width=True,
                                        hide_index=True,
                                    )

                                export_df = comparison[show_cols].copy()
                                if not extra_df.empty:
                                    extras_export = extra_df.rename(
                                        columns={
                                            "Uploaded Part Number": "Required Part Number",
                                            "Uploaded Description": "Required Description",
                                            "Uploaded Qty": "Uploaded Qty",
                                        }
                                    )
                                    extras_export["Required Qty"] = 0
                                    extras_export["Qty Gap"] = 0
                                    extras_export["Coverage %"] = 0
                                    extras_export["Option 2 Unit Price"] = pd.NA
                                    extras_export["Option 2 Estimated Cost"] = pd.NA
                                    extras_export["Currency"] = "EUR"
                                    export_df = pd.concat(
                                        [
                                            export_df,
                                            extras_export[["Required Part Number", "Required Description", "Required Qty", "Uploaded Qty", "Qty Gap", "Coverage %", "Option 2 Unit Price", "Option 2 Estimated Cost", "Currency", "Status"]],
                                        ],
                                        ignore_index=True,
                                    )

                                export_df = export_df.sort_values(["Status", "Qty Gap", "Required Part Number"], ascending=[True, False, True], na_position="last").reset_index(drop=True)
                                purchase_export = purchase_df.reset_index(drop=True) if not purchase_df.empty else pd.DataFrame(columns=["Required Part Number", "Required Description", "Qty Gap", "Option 2 Unit Price", "Option 2 Estimated Cost", "Currency", "Status"])
                                if not purchase_export.empty:
                                    purchase_total_row = {col: "" for col in purchase_export.columns}
                                    if "Required Part Number" in purchase_total_row:
                                        purchase_total_row["Required Part Number"] = "TOTAL"
                                    if "Option 2 Estimated Cost" in purchase_total_row:
                                        purchase_total_row["Option 2 Estimated Cost"] = round(option2_cost, 2)
                                    if "Currency" in purchase_total_row:
                                        purchase_total_row["Currency"] = option2_currency
                                    purchase_export = pd.concat([purchase_export, pd.DataFrame([purchase_total_row])], ignore_index=True)
                                if not extra_df.empty:
                                    extras_export_final = extra_df[["Uploaded Part Number", "Uploaded Description", "Uploaded Qty", "Status"]].copy()
                                    extras_export_final["__sort_uploaded_part_number"] = extras_export_final["Uploaded Part Number"].astype("string").fillna("")
                                    extras_export_final = extras_export_final.sort_values(["__sort_uploaded_part_number"], ascending=[True], na_position="last").drop(columns=["__sort_uploaded_part_number"]).reset_index(drop=True)
                                else:
                                    extras_export_final = pd.DataFrame(columns=["Uploaded Part Number", "Uploaded Description", "Uploaded Qty", "Status"])

                                excel_bytes = dataframe_to_excel_bytes({
                                    "Gap analysis": export_df,
                                    "Purchase option 2": purchase_export,
                                    "Extras not required": extras_export_final,
                                })

                                st.download_button(
                                    "Descargar análisis de faltantes",
                                    data=excel_bytes,
                                    file_name=f"carstock_gap_{normalize_key_text(detected_distributor) or 'distribuidor'}_{'_'.join(selected_families_stock) or 'familia'}.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                )

                                with st.expander("Guía para el archivo maestro consolidado"):
                                    st.markdown(
                                        """
                                        Formato recomendado para el archivo maestro único:
                                        - `Distributor name`
                                        - `Instrument family` o `Platform` (`LXL`, `LXS`, `MDX`, `EMX`)
                                        - `Part Number`
                                        - `Description`
                                        - `Required Qty`

                                        El archivo puede estar en una sola hoja o varias hojas. El dashboard intentará reconocer sinónimos de estas columnas automáticamente.
                                        """
                                    )

with detail_tab:
    st.subheader("Detalle por equipo")
    detail_df = filtered.copy()
    detail_df["selector"] = (
        detail_df["Serial number"].fillna("SIN SERIAL").astype(str)
        + " | "
        + detail_df["Customer name"].fillna("SIN CLIENTE").astype(str)
        + " | "
        + detail_df["Country"].fillna("SIN PAÍS").astype(str)
    )
    serial_search = st.text_input(
        "Buscar por serial",
        value="",
        placeholder="Escribe aquí un serial para encontrar el equipo",
        key="detail_serial_search",
    ).strip()
    if serial_search:
        detail_options = detail_df[
            detail_df["Serial number"].astype(str).str.contains(serial_search, case=False, na=False)
        ]["selector"].tolist()
        if not detail_options:
            st.warning("No encontré equipos con ese serial dentro del filtro actual.")
            detail_options = detail_df["selector"].tolist()
    else:
        detail_options = detail_df["selector"].tolist()
    selected = st.selectbox("Selecciona un equipo", options=detail_options)
    row = detail_df.loc[detail_df["selector"] == selected].iloc[0]

    d1, d2, d3, d4 = st.columns(4)
    with d1:
        metric_card("Serial", safe_text(row.get("Serial number")), safe_text(row.get("Instrument type"), ""))
    with d2:
        metric_card("Estado operativo", safe_text(row.get("Operational status")), safe_text(row.get("Asset condition"), ""))
    with d3:
        metric_card("Operating System", safe_text(row.get("Operating System")), safe_text(row.get("Country"), ""))
    with d4:
        metric_card("Tests / día", safe_number_text(row.get("Number of tests per day")), safe_text(row.get("Distributor name"), ""))

    detail_columns = [
        "Commercial Region",
        "Country",
        "Distributor name",
        "Customer name",
        "City",
        "Address",
        "Instrument type",
        "Product Line",
        "Installation date",
        "Operational status",
        "Type of contract",
        "Operating System",
    ]
    detail_values = []
    for c in detail_columns:
        value = row.get(c)
        if "date" in c.lower():
            detail_values.append(format_date_for_hover(value))
        else:
            detail_values.append(safe_text(value, "N/A"))

    st.dataframe(pd.DataFrame({"Campo": detail_columns, "Valor": detail_values}), use_container_width=True, hide_index=True)

    applicable_row_fields = []
    for key in active_config_fields(detail_df.loc[[row.name]], CONFIG_KEYS):
        col = f"CFG::{key}"
        applicable_row_fields.append({"Campo": key, "Valor": safe_text(row.get(col), "N/A")})

    if applicable_row_fields:
        st.markdown("### Machine configuration del equipo")
        st.dataframe(pd.DataFrame(applicable_row_fields), use_container_width=True, hide_index=True)

    with st.expander("Machine configurations completas"):
        st.code(safe_text(row.get("Machine Configurations"), "No disponible"))


with st.sidebar:
    st.subheader("📄 Informe PDF")
    st.caption("Formato ajustado con márgenes APA de 1 pulgada y estructura de informe técnico ejecutiva.")
    pdf_title = st.text_input("Título del informe", value="Installed Base Dashboard Report")
    pdf_author = st.text_input("Nombre para firma", value="Javier Avellaneda")
    pdf_role = st.text_input("Cargo / título", value="Service Leader | Export LATAM")
    pdf_signature_date = st.text_input("Fecha de firma", value=datetime.now().strftime("%Y-%m-%d"))
    pdf_references = st.text_area(
        "Referencias APA (una por línea, opcional)",
        value="",
        height=110,
        placeholder="Ejemplo:\nDiaSorin. (2026). Installed base export dashboard. Internal operational dataset.",
    )

    if not REPORTLAB_AVAILABLE:
        st.warning("La exportación PDF requiere reportlab en el entorno. Agrega `reportlab` a requirements.txt.")
    else:
        base_summary = build_filter_summary(
            selected_regions=selected_regions,
            selected_countries=selected_countries,
            selected_distributors=selected_distributors,
            selected_instruments=selected_instruments,
            selected_states=selected_states,
        )
        pdf_filter_summary = dict(base_summary) if isinstance(base_summary, dict) else {"Filters": str(base_summary)}
        pdf_filter_summary["Total records"] = f"{len(filtered):,}"

        if st.button("🧾 Preparar informe PDF", use_container_width=True, key="prepare_pdf_report"):
            try:
                prepared_bytes = build_pdf_report(
                    filtered_df=filtered,
                    filter_summary=pdf_filter_summary,
                    report_title=pdf_title,
                    author_name=pdf_author,
                    author_role=pdf_role,
                    signature_date=pdf_signature_date,
                    references_text=pdf_references,
                    stock_context=st.session_state.get("pdf_stock_context", {"available": False}),
                )
                st.session_state["prepared_pdf_bytes"] = prepared_bytes
                st.session_state["prepared_pdf_name"] = f"dashboard_report_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf"
                st.success("Informe PDF preparado correctamente.")
            except Exception as e:
                st.session_state.pop("prepared_pdf_bytes", None)
                st.error(f"No fue posible generar el PDF: {e}")

        if st.session_state.get("prepared_pdf_bytes") is not None:
            st.download_button(
                "⬇️ Descargar informe PDF (APA)",
                data=st.session_state["prepared_pdf_bytes"],
                file_name=st.session_state.get("prepared_pdf_name", f"dashboard_report_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf"),
                mime="application/pdf",
                use_container_width=True,
                key="download_prepared_pdf",
            )

st.markdown("---")
foot_l, foot_r = st.columns((0.75, 0.25))
with foot_l:
    st.markdown(
        '<div class="small-note">Filtros activos: región comercial, país, distribuidor, tipo de instrumento y estado operativo. Base instalada incluye análisis por ciudad. Sistema operativo prioriza la detección de equipos legacy que deben migrar a Windows 10. En stock, el dashboard intenta identificar automáticamente el distribuidor a partir del título del archivo cargado. La fuente activa de Records List conserva el último archivo que subas durante la sesión.</div>',
        unsafe_allow_html=True,
    )
with foot_r:
    st.download_button(
        "Descargar vista filtrada",
        data=to_csv_download(filtered.drop(columns=[c for c in filtered.columns if c.startswith("FLAG::")], errors="ignore")),
        file_name="records_list_filtered.csv",
        mime="text/csv",
        use_container_width=True,
    )
