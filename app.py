from __future__ import annotations

from pathlib import Path
from datetime import date, datetime
from io import BytesIO, StringIO
import io
import re
import textwrap
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
        return "Todos"
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
        "Región comercial": clean_filter_value(selected_regions),
        "País": clean_filter_value(selected_countries),
        "Distribuidor": clean_filter_value(selected_distributors),
        "Tipo de instrumento": clean_filter_value(selected_instruments),
        "Estado operativo": clean_filter_value(selected_states),
    }


def translate_status_value(value: str) -> str:
    mapping = {
        'Routine': 'En rutina',
        'NOT IN ROUTINE': 'No en rutina',
        'IN ROUTINE': 'En rutina',
        'Scrapped': 'Descartado',
        'Warehouse To Be Refurbished': 'Bodega por reacondicionar',
        'WAREHOUSE to be refurbished': 'Bodega por reacondicionar',
        'Warehouse Ready To Be Installed': 'Bodega lista para instalar',
        'WAREHOUSE ready to be installed': 'Bodega lista para instalar',
        'WAREHOUSE to be scrapped': 'Bodega por descartar',
        'Refurbisched': 'Reacondicionado',
        'Refurbished': 'Reacondicionado',
        'Unknown': 'No informado',
        'Not installed': 'No informado',
        'Missing': 'Faltante',
        'LOW': 'Bajo',
        'OK': 'OK',
    }
    txt = safe_text(value, 'No informado')
    return mapping.get(txt, txt)


def format_pdf_numeric_value(value):
    if pd.isna(value):
        return 'No informado'
    if isinstance(value, (int, np.integer)):
        return f"{int(value):,}".replace(',', '.')
    if isinstance(value, (float, np.floating)):
        if abs(value - round(value)) < 1e-9:
            return f"{int(round(value)):,}".replace(',', '.')
        return f"{float(value):,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
    text_value = str(value).strip()
    return text_value if text_value else 'No informado'


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

    if "Installation date" in report_df.columns:
        report_df["Installation date"] = pd.to_datetime(report_df["Installation date"], errors="coerce").dt.strftime("%Y-%m-%d").fillna("No informado")

    for col in report_df.columns:
        if col == "Installation date":
            continue
        report_df[col] = report_df[col].fillna("No informado").astype(str).str.strip().replace("", "No informado")

    for col in ["Operational status grouped", "Operational status", "Asset condition", "Operating System"]:
        if col in report_df.columns:
            report_df[col] = report_df[col].map(translate_status_value)

    report_df = report_df.rename(columns={
        "Commercial Region": "Región",
        "Country": "País",
        "Distributor name": "Distribuidor",
        "Customer name": "Cliente",
        "Instrument type": "Instrumento",
        "Serial number": "Serial",
        "Operational status grouped": "Estado",
        "Operational status": "Estado detallado",
        "Operating System": "Sistema operativo",
        "Asset condition": "Condición",
        "Installation date": "Fecha de instalación",
        "Type of contract": "Tipo de contrato",
    })

    ordered_cols = [c for c in ["Región", "País", "Distribuidor", "Cliente", "Instrumento", "Serial", "Estado", "Estado detallado", "Sistema operativo", "Condición", "Fecha de instalación", "Tipo de contrato"] if c in report_df.columns]
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
    text = safe_text(value, "No informado")
    return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def _paragraph_cell(value, style):
    return Paragraph(_escape_pdf_text(value), style)


def _df_to_wrapped_table(df: pd.DataFrame, styles, col_widths=None, max_rows=None):
    work = df.copy()
    if max_rows is not None:
        work = work.head(max_rows)
    if work.empty:
        return Paragraph("No hay datos disponibles para esta sección.", styles["APA_Body"])

    for col in work.columns:
        work[col] = work[col].map(format_pdf_numeric_value).astype(str).str.slice(0, 140)

    cell_style = styles["APA_Cell_Tiny"] if len(work.columns) >= 8 else styles["APA_Cell"]
    header_style = styles["APA_Cell_Header_Tiny"] if len(work.columns) >= 8 else styles["APA_Cell_Header"]
    header_row = [Paragraph(f"<b>{_escape_pdf_text(c)}</b>", header_style) for c in work.columns]
    body_rows = [[_paragraph_cell(v, cell_style) for v in row] for row in work.values.tolist()]
    if col_widths is None:
        usable_width = 10.9 * inch
        per_col = usable_width / max(len(work.columns), 1)
        col_widths = [per_col] * len(work.columns)
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



def _summary_table_from_pairs(title: str, pairs: list[tuple[str, str]], styles, col1='Métrica', col2='Valor'):
    data = [[Paragraph(f"<b>{_escape_pdf_text(col1)}</b>", styles["APA_Cell_Header"]), Paragraph(f"<b>{_escape_pdf_text(col2)}</b>", styles["APA_Cell_Header"])]]
    for k, v in pairs:
        data.append([_paragraph_cell(k, styles["APA_Cell"]), _paragraph_cell(v, styles["APA_Cell"])])
    table = Table(data, colWidths=[2.8 * inch, 2.8 * inch], repeatRows=1)
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1f3b64")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("LINEBELOW", (0, 0), (-1, 0), 0.8, colors.HexColor("#1f3b64")),
        ("LINEABOVE", (0, 1), (-1, -1), 0.25, colors.HexColor("#D9D9D9")),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F6F8FB")]),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 5),
        ("RIGHTPADDING", (0, 0), (-1, -1), 5),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]))
    return [Paragraph(title, styles["APA_Heading"]), table]


def _wrap_label(text_value, max_chars: int = 28) -> str:
    text_value = safe_text(text_value, "No informado")
    words = str(text_value).split()
    lines, current = [], ""
    for word in words:
        candidate = f"{current} {word}".strip()
        if len(candidate) <= max_chars:
            current = candidate
        else:
            if current:
                lines.append(current)
            current = word
    if current:
        lines.append(current)
    return "\n".join(lines[:3])


def _clean_spare_qty(series: pd.Series) -> pd.Series:
    vals = pd.to_numeric(series, errors="coerce").fillna(0.0)
    vals = vals.clip(lower=0.0)
    return vals


def _safe_share_pct(part, total) -> float:
    try:
        total = float(total)
        part = float(part)
        if total <= 0:
            return 0.0
        return round(part * 100.0 / total, 1)
    except Exception:
        return 0.0


def _make_pdf_barh(df: pd.DataFrame, label_col: str, value_col: str, title: str, xlabel: str = "Cantidad", max_rows: int = 10, color: str = "#2F80ED"):
    if not MATPLOTLIB_AVAILABLE or df is None or df.empty or label_col not in df.columns or value_col not in df.columns:
        return None
    work = df[[label_col, value_col]].copy().dropna()
    if work.empty:
        return None
    work[value_col] = pd.to_numeric(work[value_col], errors="coerce")
    work = work.dropna()
    if work.empty:
        return None
    work = work.sort_values(value_col, ascending=False).head(max_rows).sort_values(value_col, ascending=True)
    work[label_col] = work[label_col].map(lambda x: _wrap_label(x, 30))
    height = max(2.8, 0.48 * len(work) + 1.25)
    fig, ax = plt.subplots(figsize=(8.6, height))
    bars = ax.barh(work[label_col].astype(str), work[value_col].astype(float), color=color)
    ax.set_title(title, fontsize=12, fontweight='bold')
    ax.set_xlabel(xlabel, fontsize=9)
    ax.tick_params(axis='y', labelsize=8)
    ax.tick_params(axis='x', labelsize=8)
    ax.grid(axis='x', alpha=0.22)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    for bar in bars:
        value = bar.get_width()
        ax.text(value + max(work[value_col].max() * 0.01, 0.1), bar.get_y() + bar.get_height()/2, safe_number_text(value, '0'), va='center', fontsize=8)
    fig.tight_layout()
    buf = BytesIO()
    fig.savefig(buf, format='png', dpi=180, bbox_inches='tight', facecolor='white')
    plt.close(fig)
    buf.seek(0)
    return buf




def _make_pdf_stacked_barh(df: pd.DataFrame, category_col: str, segment_col: str, value_col: str, title: str, max_categories: int = 8, max_segments: int = 6):
    if df is None or df.empty or category_col not in df.columns or segment_col not in df.columns or value_col not in df.columns:
        return None

    work = df.copy()
    work[category_col] = work[category_col].fillna('No informado').astype(str)
    work[segment_col] = work[segment_col].fillna('No informado').astype(str)
    work[value_col] = pd.to_numeric(work[value_col], errors='coerce').fillna(0)
    work = work[work[value_col] > 0]
    if work.empty:
        return None

    cat_order = (
        work.groupby(category_col, as_index=False)[value_col]
        .sum()
        .sort_values(value_col, ascending=False)[category_col]
        .tolist()
    )[:max_categories]
    work = work[work[category_col].isin(cat_order)].copy()

    seg_order = (
        work.groupby(segment_col, as_index=False)[value_col]
        .sum()
        .sort_values(value_col, ascending=False)[segment_col]
        .tolist()
    )
    if len(seg_order) > max_segments:
        keep = seg_order[: max_segments - 1]
        work[segment_col] = np.where(work[segment_col].isin(keep), work[segment_col], 'Otros')
        seg_order = keep + ['Otros']
        work = work.groupby([category_col, segment_col], as_index=False)[value_col].sum()

    pivot = work.pivot_table(index=category_col, columns=segment_col, values=value_col, aggfunc='sum', fill_value=0)
    pivot = pivot.reindex(index=cat_order, fill_value=0)
    ordered_cols = [c for c in seg_order if c in pivot.columns] + [c for c in pivot.columns if c not in seg_order]
    pivot = pivot[ordered_cols]

    fig, ax = plt.subplots(figsize=(8.6, 4.6))
    left = np.zeros(len(pivot))
    palette = ['#5BC0EB', '#9BB1FF', '#63E0C9', '#FDBA5A', '#6B7280', '#C084FC']

    for i, col in enumerate(pivot.columns):
        values = pivot[col].to_numpy(dtype=float)
        bars = ax.barh(pivot.index.astype(str), values, left=left, color=palette[i % len(palette)], label=str(col))
        for b, v, l in zip(bars, values, left):
            if v >= 1:
                ax.text(l + v + 0.3, b.get_y() + b.get_height()/2, f'{int(v)}', va='center', ha='left', fontsize=8, color='#0F172A')
        left = left + values

    ax.set_title(title, fontsize=13, fontweight='bold')
    ax.set_xlabel('Cantidad')
    ax.set_ylabel('Modelo')
    ax.grid(axis='x', alpha=0.22)
    ax.invert_yaxis()
    ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.14), ncol=min(3, max(1, len(pivot.columns))), frameon=False, fontsize=8)
    fig.tight_layout()
    buf = io.BytesIO()
    fig.savefig(buf, format='png', dpi=220, bbox_inches='tight')
    plt.close(fig)
    buf.seek(0)
    return buf
def _make_pdf_donut(df: pd.DataFrame, label_col: str, value_col: str, title: str, max_rows: int = 5):
    if not MATPLOTLIB_AVAILABLE or df is None or df.empty or label_col not in df.columns or value_col not in df.columns:
        return None
    work = df[[label_col, value_col]].copy().dropna()
    if work.empty:
        return None
    work[value_col] = pd.to_numeric(work[value_col], errors='coerce')
    work = work.dropna()
    work = work[work[value_col] > 0]
    if work.empty:
        return None
    work = work.sort_values(value_col, ascending=False)
    if len(work) > max_rows:
        top = work.head(max_rows - 1).copy()
        others = work.iloc[max_rows - 1:][value_col].sum()
        top = pd.concat([top, pd.DataFrame({label_col: ['Otros'], value_col: [others]})], ignore_index=True)
        work = top
    labels = work[label_col].astype(str).map(lambda x: _wrap_label(x, 22)).tolist()
    values = work[value_col].astype(float).tolist()
    total = sum(values)
    colors_list = ['#2F80ED', '#56CCF2', '#27AE60', '#F2C94C', '#EB5757', '#9B51E0']
    fig, ax = plt.subplots(figsize=(4.8, 3.6))
    wedges, texts, autotexts = ax.pie(
        values,
        labels=None,
        colors=colors_list[:len(values)],
        startangle=90,
        counterclock=False,
        wedgeprops=dict(width=0.35, edgecolor='white'),
        autopct=lambda pct: f'{pct:.1f}%' if pct >= 8 else ''
    )
    ax.text(0, 0.05, safe_number_text(total, '0'), ha='center', va='center', fontsize=16, fontweight='bold')
    ax.text(0, -0.15, 'equipos', ha='center', va='center', fontsize=9)
    ax.set_title(_wrap_label(title, 34), fontsize=10.5, fontweight='bold', pad=10)
    ax.legend(wedges, labels, loc='lower center', bbox_to_anchor=(0.5, -0.22), ncol=2, fontsize=7, frameon=False, columnspacing=1.2, handletextpad=0.6)
    fig.tight_layout()
    buf = BytesIO()
    fig.savefig(buf, format='png', dpi=180, bbox_inches='tight', facecolor='white')
    plt.close(fig)
    buf.seek(0)
    return buf


def _make_pdf_hist_categories(df: pd.DataFrame, label_col: str, order: list[str], title: str, xlabel: str = 'Cantidad'):
    counts = df[label_col].fillna('No informado').astype(str).value_counts()
    chart_df = pd.DataFrame({label_col: order, 'Count': [int(counts.get(v, 0)) for v in order]})
    chart_df = chart_df[chart_df['Count'] > 0]
    return _make_pdf_barh(chart_df, label_col, 'Count', title, xlabel=xlabel, max_rows=len(chart_df))


def _pdf_image_flowables(image_buffers: list, max_per_row: int = 2, image_width: float = 4.7 * inch, image_height: float = 2.9 * inch):
    from reportlab.platypus import Image
    flowables = []
    valid = [img for img in image_buffers if img is not None]
    if not valid:
        return flowables
    for idx in range(0, len(valid), max_per_row):
        row = valid[idx:idx + max_per_row]
        imgs = [Image(img, width=image_width, height=image_height) for img in row]
        if len(imgs) == 1:
            flowables.append(imgs[0])
        else:
            tbl = Table([[imgs[0], imgs[1]]], colWidths=[image_width, image_width])
            tbl.setStyle(TableStyle([('VALIGN', (0, 0), (-1, -1), 'TOP'), ('LEFTPADDING', (0, 0), (-1, -1), 0), ('RIGHTPADDING', (0, 0), (-1, -1), 8)]))
            flowables.append(tbl)
        flowables.append(Spacer(1, 0.10 * inch))
    return flowables


def _build_machine_config_summary(filtered_df: pd.DataFrame):
    cfg_cols = [c for c in filtered_df.columns if c.startswith('CFG::')]
    cfg_cov_rows = []
    value_summary_rows = []
    chart_buffers = []
    for col in cfg_cols:
        non_null = filtered_df[col].dropna()
        count_non_null = int(non_null.shape[0])
        if count_non_null <= 0:
            continue
        field_name = col.replace('CFG::', '')
        if field_name.strip().lower() in {'operative system', 'operating system'}:
            continue
        vc = non_null.astype(str).str.strip()
        vc = vc[vc != '']
        if vc.empty:
            continue
        counts = vc.value_counts().reset_index()
        counts.columns = ['Value', 'Count']
        counts['Share %'] = counts['Count'].map(lambda x: _safe_share_pct(x, counts['Count'].sum()))
        top_value = safe_text(counts.iloc[0]['Value'])
        top_count = int(counts.iloc[0]['Count'])
        cfg_cov_rows.append({'Campo de configuración': field_name, 'Equipos con dato': count_non_null})
        value_summary_rows.append({
            'Campo de configuración': field_name,
            'Equipos con dato': count_non_null,
            'Valores únicos': int(counts.shape[0]),
            'Valor principal': top_value,
            'Conteo principal': top_count,
        })
        for _, row in counts.head(5).iterrows():
            value_summary_rows.append({
                'Campo de configuración': f"{field_name} — valor",
                'Equipos con dato': '',
                'Valores únicos': '',
                'Valor principal': safe_text(row['Value']),
                'Conteo principal': f"{int(row['Count'])} ({row['Share %']:.1f}%)",
            })
        if counts.shape[0] <= 5 and counts.iloc[0]['Count'] / counts['Count'].sum() < 0.86:
            chart = _make_pdf_donut(counts.rename(columns={'Value': 'Categoría'}), 'Categoría', 'Count', field_name, max_rows=5)
        else:
            chart = _make_pdf_barh(counts.rename(columns={'Value': 'Categoría'}), 'Categoría', 'Count', field_name, xlabel='Equipos', max_rows=6, color='#1f77b4')
        chart_buffers.append((count_non_null, chart))
    cov_df = pd.DataFrame(cfg_cov_rows).sort_values('Equipos con dato', ascending=False) if cfg_cov_rows else pd.DataFrame(columns=['Campo de configuración', 'Equipos con dato'])
    value_df = pd.DataFrame(value_summary_rows)
    charts = [c for _, c in sorted(chart_buffers, key=lambda x: x[0], reverse=True)[:6]]
    return cov_df, value_df, charts


def _build_executive_insights(filtered_df: pd.DataFrame, stock_context: dict | None = None) -> tuple[list[str], list[str]]:
    insights = []
    recommendations = []
    total_records = len(filtered_df)
    countries = filtered_df['Country'].fillna('No informado').value_counts()
    if not countries.empty:
        top_country = countries.index[0]
        top_country_pct = _safe_share_pct(countries.iloc[0], total_records)
        insights.append(f"La base instalada filtrada contiene {total_records} equipos y se concentra principalmente en {top_country} ({top_country_pct}%).")
    routine_assets = int(filtered_df.get('Is in routine', pd.Series(dtype=bool)).sum())
    insights.append(f"Se identificaron {routine_assets} equipos en rutina dentro del universo filtrado.")
    os_series = filtered_df.get('Operating System', pd.Series(dtype=object)).fillna('No informado').astype(str)
    legacy_count = int(os_series.isin(['Windows XP', 'Windows Vista', 'Windows 7', 'Windows 2000']).sum())
    unknown_os = int(os_series.isin(['Unknown', 'No informado', 'Not installed']).sum())
    if legacy_count > 0:
        insights.append(f"Existen {legacy_count} equipos con sistemas operativos legados que deben priorizarse en el plan de actualización.")
    if unknown_os > 0:
        insights.append(f"Hay {unknown_os} equipos sin visibilidad clara del sistema operativo, lo que limita la planeación técnica.")
    if 'PM next date' in filtered_df.columns:
        pm_next = pd.to_datetime(filtered_df['PM next date'], errors='coerce')
        overdue_pm = int((pm_next < pd.Timestamp.today().normalize()).fillna(False).sum())
        if overdue_pm > 0:
            insights.append(f"Se detectaron {overdue_pm} mantenimientos preventivos vencidos en la vista actual.")
    if stock_context and stock_context.get('available'):
        missing = int(stock_context.get('missing_skus', 0))
        low = int(stock_context.get('low_skus', 0))
        cost = float(stock_context.get('option2_cost', 0) or 0)
        insights.append(f"La revisión de carstock identificó {missing} SKUs faltantes y {low} SKUs en nivel bajo, con una exposición estimada de EUR {cost:,.2f}.")

        if missing > 0:
            recommendations.append('Priorizar la compra de repuestos faltantes y de bajo stock con base en el costo estimado y la criticidad operativa.')
    if legacy_count > 0:
        recommendations.append('Ejecutar un plan de migración para equipos con Windows Vista/XP/7 y validar de inmediato los activos sin dato de sistema operativo.')
    if unknown_os > 0:
        recommendations.append('Completar los campos vacíos de sistema operativo y configuración de equipo para mejorar la trazabilidad del parque instalado.')
    if 'PM next date' in filtered_df.columns:
        pm_next = pd.to_datetime(filtered_df['PM next date'], errors='coerce')
        overdue_pm = int((pm_next < pd.Timestamp.today().normalize()).fillna(False).sum())
        if overdue_pm > 0:
            recommendations.append('Reprogramar los mantenimientos preventivos vencidos y ordenar la ejecución por volumen de pruebas y criticidad del cliente.')
    recommendations.append('Usar este informe como base para una revisión ejecutiva del distribuidor, combinando base instalada, OS, PM y cobertura de repuestos.')
    return insights[:6], recommendations[:5]


def _build_pdf_sections(filtered_df: pd.DataFrame, stock_context: dict | None = None):
    sections = []
    annexes = []

    base_pairs = [
        ('Registros filtrados', f"{len(filtered_df):,}"),
        ('Países', f"{filtered_df['Country'].nunique(dropna=True):,}"),
        ('Distribuidores', f"{filtered_df['Distributor name'].nunique(dropna=True):,}"),
        ('Tipos de instrumento', f"{filtered_df['Instrument type'].nunique(dropna=True):,}"),
        ('Equipos en rutina', f"{int(filtered_df.get('Is in routine', pd.Series(dtype=bool)).sum()):,}"),
    ]
    top_country = filtered_df['Country'].fillna('No informado').value_counts().reset_index()
    top_country.columns = ['País', 'Cantidad']
    top_inst = filtered_df['Instrument type'].fillna('No informado').value_counts().reset_index()
    top_inst.columns = ['Instrumento', 'Cantidad']
    state_counts = filtered_df['Operational status grouped'].fillna('No informado').value_counts().reset_index()
    state_counts.columns = ['Estado', 'Cantidad']
    age_df = filtered_df.copy()
    age_df['Rango de antigüedad'] = pd.cut(
        pd.to_numeric(age_df.get('Age (years)', pd.Series(dtype=float)), errors='coerce'),
        bins=[-1, 5, 8, 10, 100],
        labels=['0-5 años', '5-8 años', '8-10 años', '10+ años']
    )
    age_counts = age_df['Rango de antigüedad'].value_counts().reset_index()
    age_counts.columns = ['Rango', 'Cantidad']
    age_counts['Rango'] = pd.Categorical(age_counts['Rango'], categories=['0-5 años', '5-8 años', '8-10 años', '10+ años'], ordered=True)
    age_counts = age_counts.sort_values('Rango')

    corporate_model_charts = []
    corporate_model_df = filtered_df.copy()
    corporate_model_df['Instrument type'] = corporate_model_df['Instrument type'].fillna('No informado').astype(str)
    corporate_model_df['Distributor name'] = corporate_model_df['Distributor name'].fillna('No informado').astype(str)
    model_rank = corporate_model_df['Instrument type'].value_counts().index.tolist()
    global_dist = (
        corporate_model_df.groupby(['Instrument type', 'Distributor name'], dropna=False)
        .size()
        .reset_index(name='Cantidad')
        .rename(columns={'Instrument type': 'Modelo', 'Distributor name': 'Distribuidor'})
    )
    if not global_dist.empty:
        top_global = global_dist.groupby('Distribuidor', as_index=False)['Cantidad'].sum().sort_values(['Cantidad', 'Distribuidor'], ascending=[False, True]).head(5)['Distribuidor'].tolist()
        global_dist_main = global_dist[global_dist['Distribuidor'].isin(top_global)].copy()
        global_dist_main['Distribuidor'] = global_dist_main['Distribuidor'].astype(str).map(lambda x: distributor_display_name(x, 18))
        corporate_model_charts.append(_make_pdf_stacked_barh(global_dist_main, 'Modelo', 'Distribuidor', 'Cantidad', 'Vista global por distribuidor | resumen (Top 5)', max_categories=8, max_segments=6))
    detail_corporate_rows = []
    for model_name in model_rank[:6]:
        model_slice = corporate_model_df[corporate_model_df['Instrument type'] == model_name].copy()
        counts = model_slice['Distributor name'].value_counts().reset_index()
        counts.columns = ['Distribuidor', 'Cantidad']
        counts = counts.sort_values(['Cantidad', 'Distribuidor'], ascending=[False, True]).reset_index(drop=True)
        if counts.empty:
            continue
        top_counts = counts.head(5).copy()
        top_counts['Distribuidor'] = top_counts['Distribuidor'].astype(str).map(lambda x: distributor_display_name(x, 20))
        corporate_model_charts.append(_make_pdf_donut(top_counts, 'Distribuidor', 'Cantidad', f'Distribución por distribuidor | {model_name} | Top 5', max_rows=5))
        full_counts = counts.copy()
        full_counts['Modelo'] = model_name
        detail_corporate_rows.append(full_counts)

    sections.append({
        'title': 'Resumen de base instalada',
        'intro': 'Esta sección resume la base instalada filtrada y destaca la concentración geográfica, el mix de instrumentos, el estado operativo, el perfil de antigüedad y la distribución por distribuidor para cada modelo visible.',
        'summary_pairs': base_pairs,
        'charts': [
            _make_pdf_barh(top_country, 'País', 'Cantidad', 'Países con mayor concentración', max_rows=10),
            _make_pdf_barh(top_inst, 'Instrumento', 'Cantidad', 'Mix de instrumentos', max_rows=10),
            _make_pdf_barh(state_counts, 'Estado', 'Cantidad', 'Distribución por estado operativo', max_rows=10),
            _make_pdf_barh(age_counts, 'Rango', 'Cantidad', 'Perfil de antigüedad', max_rows=4),
        ] + corporate_model_charts,
        'table_title': 'Muestra resumida de equipos filtrados',
        'table_df': prepare_pdf_report_table(filtered_df),
        'table_max_rows': 10,
    })
    annexes.append({
        'title': 'Anexo A. Base instalada detallada',
        'intro': 'Detalle tabular de la base instalada filtrada.',
        'summary_pairs': [('Filas incluidas', f"{len(filtered_df):,}"), ('Alcance', 'Detalle completo de la base instalada filtrada')],
        'charts': [],
        'table_title': 'Detalle completo de equipos filtrados',
        'table_df': prepare_pdf_report_table(filtered_df),
        'table_max_rows': max(len(filtered_df), 1),
    })
    if detail_corporate_rows:
        detail_corporate_df = pd.concat(detail_corporate_rows, ignore_index=True)
        detail_charts = []
        for model_name in model_rank[:6]:
            model_detail = detail_corporate_df[detail_corporate_df['Modelo'].eq(model_name)].copy()
            if model_detail.empty:
                continue
            model_detail['Distribuidor'] = model_detail['Distribuidor'].astype(str).map(lambda x: distributor_display_name(x, 28))
            detail_charts.append(_make_pdf_barh(model_detail, 'Distribuidor', 'Cantidad', f'Detalle completo | {model_name}', xlabel='Cantidad de equipos', max_rows=max(12, len(model_detail)), color='#2F80ED'))
        annexes.append({
            'title': 'Anexo B. Distribución completa por distribuidor y modelo',
            'intro': 'Detalle completo de distribuidores por modelo. En el cuerpo principal solo se muestra el resumen Top 5 para mantener la lectura ejecutiva.',
            'summary_pairs': [('Filas incluidas', f"{len(detail_corporate_df):,}"), ('Alcance', 'Detalle completo de distribuidores por modelo')],
            'charts': detail_charts,
            'table_title': 'Detalle completo por modelo',
            'table_df': detail_corporate_df[['Modelo', 'Distribuidor', 'Cantidad']].sort_values(['Modelo', 'Cantidad', 'Distribuidor'], ascending=[True, False, True]),
            'table_max_rows': max(len(detail_corporate_df), 1),
        })

    cfg_pairs = [
        ('Equipos con configuración', f"{int(filtered_df['Machine Configurations'].notna().sum()):,}"),
        ('Campos activos de configuración', f"{sum(int(filtered_df[c].notna().sum()) > 0 for c in filtered_df.columns if c.startswith('CFG::')):,}"),
        ('Promedio de campos poblados', f"{filtered_df.get('Machine config fields populated', pd.Series([0])).fillna(0).mean():.1f}"),
    ]
    cfg_cov, cfg_value_df, cfg_charts = _build_machine_config_summary(filtered_df)
    cfg_charts = cfg_charts[:4]
    sections.append({
        'title': 'Configuración de equipo',
        'intro': 'Se consolidan los campos detectados en configuración de equipo y se muestran las distribuciones de los ítems con mayor visibilidad en el filtro activo.',
        'summary_pairs': cfg_pairs,
        'charts': [_make_pdf_barh(cfg_cov, 'Campo de configuración', 'Equipos con dato', 'Cobertura de campos de configuración', max_rows=10)] + cfg_charts,
        'table_title': 'Resumen de configuración de equipo',
        'table_df': cfg_value_df,
        'table_max_rows': 12,
    })
    if not cfg_value_df.empty:
        annexes.append({
            'title': 'Anexo C. Valores de configuración',
            'intro': 'Valores principales por campo de configuración.',
            'summary_pairs': [('Filas incluidas', f"{len(cfg_value_df):,}"), ('Alcance', 'Resumen ampliado de campos y valores de configuración')],
            'charts': [],
            'table_title': 'Valores principales por campo',
            'table_df': cfg_value_df,
            'table_max_rows': max(len(cfg_value_df), 1),
        })

    os_df = filtered_df.copy()
    os_df['Operating System'] = os_df['Operating System'].fillna('No informado').replace({'Unknown':'No informado','Not installed':'No informado'})
    os_df['Bucket de actualización'] = os_df['Operating System'].map(os_upgrade_bucket).replace({
        'Windows 10 / OK': 'Windows 10 / OK',
        'Legacy / urgente migrar': 'Legado / migración urgente',
        'Revisar campo OS': 'Revisar campo OS',
        'Otro OS / validar': 'Otro OS / validar',
    })
    urgent_table = os_df[os_df['Operating System'].isin(['Windows XP', 'Windows Vista', 'Windows 7', 'Windows 2000'])][['Country','Distributor name','Customer name','Instrument type','Serial number','Operating System']].copy()
    urgent_table.columns = ['País', 'Distribuidor', 'Cliente', 'Instrumento', 'Serial', 'Sistema operativo']
    os_pairs = [
        ('Equipos con OS identificado', f"{int(filtered_df['Operating System'].notna().sum()):,}"),
        ('Valores únicos de OS', f"{filtered_df['Operating System'].nunique(dropna=True):,}"),
        ('OS legado / migración urgente', f"{int(os_df['Operating System'].isin(['Windows XP','Windows Vista','Windows 7','Windows 2000']).sum()):,}"),
        ('OS no informado', f"{int(os_df['Operating System'].isin(['Unknown','No informado','Not installed']).sum()):,}"),
    ]
    os_counts = os_df['Operating System'].value_counts().reset_index()
    os_counts.columns = ['Sistema operativo', 'Cantidad']
    os_bucket = os_df['Bucket de actualización'].value_counts().reset_index()
    os_bucket.columns = ['Prioridad', 'Cantidad']
    sections.append({
        'title': 'Sistema operativo',
        'intro': 'Esta sección identifica equipos con sistemas operativos legados, visibilidad incompleta y prioridades de actualización.',
        'summary_pairs': os_pairs,
        'charts': [
            _make_pdf_barh(os_counts, 'Sistema operativo', 'Cantidad', 'Distribución de sistema operativo', max_rows=10),
            _make_pdf_barh(os_bucket, 'Prioridad', 'Cantidad', 'Priorización de actualización', max_rows=10, color='#1f77b4'),
        ],
        'table_title': 'Equipos que requieren actualización de Windows',
        'table_df': urgent_table,
        'table_max_rows': 10,
    })
    if not urgent_table.empty:
        annexes.append({
            'title': 'Anexo D. Equipos con OS legado',
            'intro': 'Detalle de equipos con sistema operativo legado.',
            'summary_pairs': [('Filas incluidas', f"{len(urgent_table):,}"), ('Alcance', 'Equipos con Windows XP/Vista/7/2000')],
            'charts': [],
            'table_title': 'Detalle de equipos con OS legado',
            'table_df': urgent_table,
            'table_max_rows': max(len(urgent_table), 1),
        })

    proc_df = filtered_df.copy()
    proc_df['Pruebas por día'] = pd.to_numeric(proc_df['Number of tests per day'], errors='coerce').fillna(0)
    today = pd.Timestamp.today().normalize()
    if 'PM next date' in proc_df.columns:
        pm_next = pd.to_datetime(proc_df['PM next date'], errors='coerce')
        proc_df['Estado PM'] = np.where(pm_next < today, 'Vencido', np.where(pm_next <= today + pd.Timedelta(days=90), 'Próximos 90 días', 'Planificado más adelante'))
    else:
        proc_df['Estado PM'] = 'No informado'
    pm_status = proc_df['Estado PM'].value_counts().reset_index()
    pm_status.columns = ['Estado PM', 'Cantidad']
    top_tests = proc_df[['Serial number', 'Pruebas por día', 'Instrument type']].copy()
    top_tests = top_tests.sort_values('Pruebas por día', ascending=False).head(10)
    top_tests['Equipo'] = top_tests['Serial number'].astype(str) + ' | ' + top_tests['Instrument type'].astype(str)
    proc_pairs = [
        ('Promedio de pruebas por día', safe_number_text(proc_df['Pruebas por día'].mean(), '0')),
        ('Máximo de pruebas por día', safe_number_text(proc_df['Pruebas por día'].max(), '0')),
        ('PM próximos 90 días', f"{int((proc_df['Estado PM'] == 'Próximos 90 días').sum()):,}"),
        ('PM vencidos', f"{int((proc_df['Estado PM'] == 'Vencido').sum()):,}"),
    ]
    proc_table = proc_df[['Country', 'Distributor name', 'Instrument type', 'Serial number', 'Pruebas por día', 'Estado PM']].copy()
    proc_table.columns = ['País', 'Distribuidor', 'Instrumento', 'Serial', 'Pruebas por día', 'Estado PM']
    sections.append({
        'title': 'Procesamiento y planificación de PM',
        'intro': 'Se prioriza la carga operativa y el estado del mantenimiento preventivo mediante visuales ejecutivas más legibles.',
        'summary_pairs': proc_pairs,
        'charts': [
            _make_pdf_barh(top_tests.rename(columns={'Equipo': 'Equipo'}), 'Equipo', 'Pruebas por día', 'Top 10 equipos por pruebas por día', xlabel='Pruebas/día', max_rows=10),
            _make_pdf_barh(pm_status, 'Estado PM', 'Cantidad', 'Estado del plan de mantenimiento preventivo', max_rows=10, color='#2D9CDB'),
        ],
        'table_title': 'Resumen de equipos con mayor volumen y estado PM',
        'table_df': proc_table.sort_values('Pruebas por día', ascending=False),
        'table_max_rows': 10,
    })
    annexes.append({
        'title': 'Anexo E. Detalle de procesamiento y PM',
        'intro': 'Detalle ampliado de pruebas por día y estado de PM.',
        'summary_pairs': [('Filas incluidas', f"{len(proc_table):,}"), ('Alcance', 'Detalle ampliado de procesamiento y mantenimiento preventivo')],
        'charts': [],
        'table_title': 'Detalle ampliado de procesamiento y PM',
        'table_df': proc_table.sort_values('Pruebas por día', ascending=False),
        'table_max_rows': max(len(proc_table), 1),
    })

    stock_context = stock_context or {}
    if stock_context.get('available'):
        full_comparison_df = stock_context.get('full_comparison_df', pd.DataFrame()).copy()
        purchase_df = stock_context.get('purchase_df', pd.DataFrame()).copy()
        extra_df = stock_context.get('extra_df', pd.DataFrame()).copy()
        stock_top_gap = stock_context.get('top_gap_df', pd.DataFrame()).copy()
        for df_ in [full_comparison_df, purchase_df, extra_df, stock_top_gap]:
            if df_ is not None and not df_.empty:
                if 'Uploaded Qty' in df_.columns:
                    df_['Uploaded Qty'] = _clean_spare_qty(df_['Uploaded Qty'])
                if 'Coverage %' in df_.columns:
                    df_['Coverage %'] = pd.to_numeric(df_['Coverage %'], errors='coerce').fillna(0.0).clip(lower=0.0, upper=999.0)
        main_status = pd.DataFrame({'Estado': ['OK', 'Bajo', 'Faltante'], 'Cantidad': [int(stock_context.get('ok_skus', 0)), int(stock_context.get('low_skus', 0)), int(stock_context.get('missing_skus', 0))]})
        extras_status = pd.DataFrame({'Estado': ['Extras'], 'Cantidad': [int(stock_context.get('extra_skus', 0))]})
        stock_pairs = [
            ('Distribuidor detectado', stock_context.get('detected_distributor', 'N/A')),
            ('Familias comparadas', ', '.join(stock_context.get('families', [])) or 'N/A'),
            ('SKUs requeridos', f"{stock_context.get('required_skus', 0):,}"),
            ('SKUs OK', f"{stock_context.get('ok_skus', 0):,}"),
            ('SKUs LOW', f"{stock_context.get('low_skus', 0):,}"),
            ('SKUs faltantes', f"{stock_context.get('missing_skus', 0):,}"),
            ('Costo estimado opción 2', f"EUR {float(stock_context.get('option2_cost', 0) or 0):,.2f}"),
        ]
        if not stock_top_gap.empty:
            stock_top_gap = stock_top_gap.copy()
            stock_top_gap['Parte'] = stock_top_gap['Required Part Number'].astype(str) + ' | ' + stock_top_gap['Required Description'].fillna('').astype(str).str.slice(0, 28)
        gap_table = pd.DataFrame()
        if not stock_top_gap.empty:
            gap_table = stock_top_gap[[c for c in ['Required Part Number','Required Description','Required Qty','Uploaded Qty','Qty Gap','Status','Option 2 Estimated Cost','Currency'] if c in stock_top_gap.columns]].copy()
            gap_table.columns = ['Parte requerida', 'Descripción', 'Cant. requerida', 'Cant. cargada', 'Brecha', 'Estado', 'Costo estimado opción 2', 'Moneda']
            if 'Estado' in gap_table.columns:
                gap_table['Estado'] = gap_table['Estado'].map(translate_status_value)
        sections.append({
            'title': 'Repuestos y brecha de carstock',
            'intro': 'Se resume la cobertura del stock requerido y la brecha estimada de compra. Los ítems extra se muestran por separado para no distorsionar la lectura principal del gap.',
            'summary_pairs': stock_pairs,
            'charts': [
                _make_pdf_barh(main_status, 'Estado', 'Cantidad', 'Cobertura del carstock requerido', max_rows=3, color='#2F80ED'),
                _make_pdf_barh(extras_status, 'Estado', 'Cantidad', 'Ítems extra no requeridos por el maestro', max_rows=1, color='#56CCF2') if int(stock_context.get('extra_skus', 0)) > 0 else None,
                _make_pdf_barh(stock_top_gap.rename(columns={'Parte': 'Parte', 'Qty Gap': 'Brecha'}), 'Parte', 'Brecha', 'Principales repuestos faltantes', xlabel='Brecha de cantidad', max_rows=10, color='#EB5757') if not stock_top_gap.empty else None,
            ],
            'table_title': 'Principales brechas de repuestos',
            'table_df': gap_table,
            'table_max_rows': 10,
        })
        if not full_comparison_df.empty:
            annex_table = full_comparison_df[[c for c in ['Required Part Number','Required Description','Required Qty','Uploaded Qty','Qty Gap','Coverage %','Status','Option 2 Unit Price','Option 2 Estimated Cost','Currency'] if c in full_comparison_df.columns]].copy()
            annex_table.columns = ['Parte requerida', 'Descripción', 'Cant. requerida', 'Cant. cargada', 'Brecha', 'Cobertura %', 'Estado', 'Precio unitario opción 2', 'Costo estimado opción 2', 'Moneda']
            if 'Estado' in annex_table.columns:
                annex_table['Estado'] = annex_table['Estado'].map(translate_status_value)
            annexes.append({
                'title': 'Anexo F. Comparación completa de repuestos',
                'intro': 'Comparación completa entre el maestro de carstock y el stock cargado por el distribuidor.',
                'summary_pairs': [('Filas incluidas', f"{len(annex_table):,}"), ('Alcance', 'Comparación completa de repuestos')],
                'charts': [],
                'table_title': 'Comparación completa de repuestos',
                'table_df': annex_table,
                'table_max_rows': max(len(annex_table), 1),
            })
        if not purchase_df.empty:
            pur_cols = [c for c in ['Required Part Number','Required Description','Qty Gap','Option 2 Unit Price','Option 2 Estimated Cost','Currency','Status'] if c in purchase_df.columns]
            pur_table = purchase_df[pur_cols].copy()
            pur_table.columns = ['Parte requerida', 'Descripción', 'Brecha', 'Precio unitario opción 2', 'Costo estimado opción 2', 'Moneda', 'Estado']
            if 'Estado' in pur_table.columns:
                pur_table['Estado'] = pur_table['Estado'].map(translate_status_value)
            annexes.append({
                'title': 'Anexo G. Lista sugerida de compra',
                'intro': 'Compra sugerida para cerrar la brecha actual de carstock.',
                'summary_pairs': [('Filas incluidas', f"{len(pur_table):,}"), ('Alcance', 'Lista sugerida de compra basada en opción 2')],
                'charts': [],
                'table_title': 'Lista sugerida de compra',
                'table_df': pur_table,
                'table_max_rows': max(len(pur_table), 1),
            })
        if not extra_df.empty:
            ex_cols = [c for c in ['Uploaded Part Number','Uploaded Description','Uploaded Qty','Status'] if c in extra_df.columns]
            ex_table = extra_df[ex_cols].copy()
            ex_table.columns = ['Parte cargada', 'Descripción cargada', 'Cantidad cargada', 'Estado']
            annexes.append({
                'title': 'Anexo H. Ítems extra no requeridos',
                'intro': 'Repuestos reportados por el distribuidor que no pertenecen al maestro de carstock seleccionado.',
                'summary_pairs': [('Filas incluidas', f"{len(ex_table):,}"), ('Alcance', 'Ítems extra no requeridos por el maestro')],
                'charts': [],
                'table_title': 'Ítems extra no requeridos',
                'table_df': ex_table,
                'table_max_rows': max(len(ex_table), 1),
            })
    return sections, annexes


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
        leftMargin=0.85 * inch,
        rightMargin=0.85 * inch,
        topMargin=0.8 * inch,
        bottomMargin=0.65 * inch,
        title=report_title,
        author=author_name,
    )

    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name="APA_Title", parent=styles["Title"], fontName="Times-Bold", fontSize=18, leading=24, alignment=TA_CENTER, spaceAfter=12, textColor=colors.HexColor("#111111")))
    styles.add(ParagraphStyle(name="APA_Subtitle", parent=styles["Normal"], fontName="Times-Roman", fontSize=12, leading=16, alignment=TA_CENTER, spaceAfter=6, textColor=colors.HexColor("#222222")))
    styles.add(ParagraphStyle(name="APA_Heading", parent=styles["Heading2"], fontName="Times-Bold", fontSize=13, leading=16, alignment=TA_LEFT, spaceBefore=4, spaceAfter=6, textColor=colors.HexColor("#111111")))
    styles.add(ParagraphStyle(name="APA_Body", parent=styles["BodyText"], fontName="Times-Roman", fontSize=11, leading=16, alignment=TA_JUSTIFY, spaceAfter=6))
    styles.add(ParagraphStyle(name="APA_Cell", parent=styles["BodyText"], fontName="Times-Roman", fontSize=8, leading=10, alignment=TA_LEFT, wordWrap='CJK'))
    styles.add(ParagraphStyle(name="APA_Cell_Header", parent=styles["BodyText"], fontName="Times-Bold", fontSize=8, leading=10, alignment=TA_LEFT, textColor=colors.white))
    styles.add(ParagraphStyle(name="APA_Cell_Tiny", parent=styles["BodyText"], fontName="Times-Roman", fontSize=7, leading=8.5, alignment=TA_LEFT, wordWrap='CJK'))
    styles.add(ParagraphStyle(name="APA_Cell_Header_Tiny", parent=styles["BodyText"], fontName="Times-Bold", fontSize=7, leading=8.5, alignment=TA_LEFT, textColor=colors.white))
    styles.add(ParagraphStyle(name="APA_Signature", parent=styles["BodyText"], fontName="Times-Roman", fontSize=11, leading=16, alignment=TA_LEFT, spaceAfter=3))

    elements = []
    short_title = re.sub(r"\s+", " ", (report_title.strip() or "Informe de base instalada"))[:80]
    generated_date = datetime.now().strftime("%Y-%m-%d %H:%M")
    org_name = "DiaSorin S.p.A."
    title_for_cover = report_title or "Informe de base instalada"

    def page_header_footer(canvas, doc):
        canvas.saveState()
        width, height = landscape(A4)
        canvas.setFont("Helvetica", 8)
        canvas.setFillColor(colors.HexColor("#555555"))
        canvas.drawString(doc.leftMargin, height - 24, short_title)
        canvas.drawRightString(width - doc.rightMargin, height - 24, f"Página {doc.page}")
        canvas.setStrokeColor(colors.HexColor("#D7D7D7"))
        canvas.setLineWidth(0.5)
        canvas.line(doc.leftMargin, height - 30, width - doc.rightMargin, height - 30)
        canvas.line(doc.leftMargin, 24, width - doc.rightMargin, 24)
        canvas.drawString(doc.leftMargin, 12, generated_date)
        canvas.drawRightString(width - doc.rightMargin, 12, "Formato APA")
        canvas.restoreState()

    elements.append(Spacer(1, 1.3 * inch))
    elements.append(Paragraph(title_for_cover, styles["APA_Title"]))
    elements.append(Spacer(1, 0.18 * inch))
    elements.append(Paragraph(author_name, styles["APA_Subtitle"]))
    elements.append(Paragraph(author_role, styles["APA_Subtitle"]))
    elements.append(Paragraph(org_name, styles["APA_Subtitle"]))
    elements.append(Paragraph(signature_date, styles["APA_Subtitle"]))
    elements.append(PageBreak())

    insights, recommendations = _build_executive_insights(filtered_df, stock_context=stock_context)
    elements.append(Paragraph("Resumen ejecutivo", styles["APA_Heading"]))
    elements.append(Paragraph(
        "Este informe consolida la información visible en el dashboard filtrado y resume la base instalada, la configuración de equipo, el sistema operativo, la planificación de mantenimiento preventivo y la cobertura de repuestos del distribuidor seleccionado.",
        styles["APA_Body"],
    ))
    for text_line in insights:
        elements.append(Paragraph(f"• {_escape_pdf_text(text_line)}", styles["APA_Body"]))
    elements.append(Spacer(1, 0.06 * inch))
    elements.append(Paragraph("Acciones recomendadas", styles["APA_Heading"]))
    for text_line in recommendations:
        elements.append(Paragraph(f"• {_escape_pdf_text(text_line)}", styles["APA_Body"]))

    elements.append(Spacer(1, 0.08 * inch))
    filters_pairs = [(k, v) for k, v in filter_summary.items()]
    for block in _summary_table_from_pairs("Filtros aplicados", filters_pairs, styles):
        elements.append(block)

    sections, annexes = _build_pdf_sections(filtered_df, stock_context=stock_context)
    from reportlab.platypus import Image
    for section in sections:
        elements.append(PageBreak())
        elements.append(Paragraph(section['title'], styles['APA_Heading']))
        if section.get('intro'):
            elements.append(Paragraph(section['intro'], styles['APA_Body']))
        for block in _summary_table_from_pairs("Resumen de la sección", section['summary_pairs'], styles):
            elements.append(block)
        charts = [c for c in section.get('charts', []) if c is not None]
        if charts:
            elements.append(Spacer(1, 0.05 * inch))
            for fl in _pdf_image_flowables(charts, max_per_row=2):
                elements.append(fl)
        table_df = section.get('table_df', pd.DataFrame())
        if table_df is not None and not table_df.empty:
            elements.append(Paragraph(section.get('table_title', 'Tabla de apoyo'), styles['APA_Heading']))
            col_widths = None
            if section['title'] == 'Resumen de base instalada':
                width_map = {'Región': 0.75 * inch, 'País': 0.75 * inch, 'Distribuidor': 1.1 * inch, 'Cliente': 1.25 * inch, 'Instrumento': 1.0 * inch, 'Serial': 0.8 * inch, 'Estado': 0.8 * inch, 'Estado detallado': 0.95 * inch, 'Sistema operativo': 0.85 * inch, 'Condición': 0.75 * inch, 'Fecha de instalación': 0.9 * inch, 'Tipo de contrato': 1.25 * inch}
                col_widths = [width_map.get(c, 0.9 * inch) for c in table_df.columns]
            elif section['title'] == 'Repuestos y brecha de carstock':
                width_map = {'Parte requerida': 1.0 * inch, 'Descripción': 1.7 * inch, 'Cant. requerida': 0.8 * inch, 'Cant. cargada': 0.8 * inch, 'Brecha': 0.7 * inch, 'Estado': 0.8 * inch, 'Costo estimado opción 2': 1.0 * inch, 'Moneda': 0.55 * inch}
                col_widths = [width_map.get(c, 0.9 * inch) for c in table_df.columns]
            max_rows = section.get('table_max_rows', len(table_df))
            elements.append(_df_to_wrapped_table(table_df, styles, col_widths=col_widths, max_rows=max_rows))
            if isinstance(max_rows, int) and len(table_df) > max_rows:
                elements.append(Paragraph(f"Nota. En el cuerpo principal solo se muestran las primeras {max_rows} filas. El detalle completo se encuentra en los anexos.", styles['APA_Body']))

    elements.append(PageBreak())
    elements.append(Paragraph("Conclusiones", styles["APA_Heading"]))
    for text_line in insights[:4]:
        elements.append(Paragraph(f"• {_escape_pdf_text(text_line)}", styles["APA_Body"]))
    elements.append(Spacer(1, 0.06 * inch))
    elements.append(Paragraph("Fuente de datos", styles["APA_Heading"]))
    elements.append(Paragraph("Fuente de datos: registros filtrados del dashboard y, cuando aplica, archivo de stock cargado en la sesión actual.", styles["APA_Body"]))
    elements.append(Spacer(1, 0.08 * inch))
    elements.append(Paragraph("Firma", styles["APA_Heading"]))
    elements.append(Paragraph(_escape_pdf_text(author_name), styles["APA_Signature"]))
    elements.append(Paragraph(_escape_pdf_text(author_role), styles["APA_Signature"]))
    elements.append(Paragraph(_escape_pdf_text(org_name), styles["APA_Signature"]))
    elements.append(Paragraph(f"Fecha: {_escape_pdf_text(signature_date)}", styles["APA_Signature"]))

    for annex in annexes:
        elements.append(PageBreak())
        elements.append(Paragraph(annex['title'], styles['APA_Heading']))
        if annex.get('intro'):
            elements.append(Paragraph(annex['intro'], styles['APA_Body']))
        for block in _summary_table_from_pairs("Resumen del anexo", annex['summary_pairs'], styles):
            elements.append(block)
        if annex.get('charts'):
            for fl in _pdf_image_flowables(annex['charts'], max_per_row=2):
                elements.append(fl)
        table_df = annex.get('table_df', pd.DataFrame())
        if table_df is not None and not table_df.empty:
            elements.append(Paragraph(annex.get('table_title', 'Tabla del anexo'), styles['APA_Heading']))
            elements.append(_df_to_wrapped_table(table_df, styles, max_rows=annex.get('table_max_rows', len(table_df))))

    def cover_page(canvas, doc):
        canvas.saveState()
        canvas.restoreState()

    doc.build(elements, onFirstPage=cover_page, onLaterPages=page_header_footer)
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



def shorten_distributor_name(name: str, max_len: int = 22) -> str:
    text_name = safe_text(name, "No informado")
    exact_map = {
        "annar diagnostica import sas": "Annar",
        "laboratorios cienvar s a": "Cienvar",
        "wm argentina s a": "WM Argentina",
        "grupo bios": "Grupo Bios",
        "bio nuclear": "Bio-Nuclear",
        "diagnostico ual": "Diag. UAL",
        "biotec del paraguay s r l": "Biotec Paraguay",
        "biotec del paraguay": "Biotec Paraguay",
        "islalab products llc": "IslaLab",
        "capris médica": "Capris",
        "capris medica": "Capris",
        "dimex medica": "Dimex",
        "caribbean medical supplies inc": "Caribbean Medical",
        "simed ecuador": "Simed Ecuador",
        "simed ecuador": "Simed Ecuador",
    }
    if text_name in exact_map:
        short = exact_map[text_name]
        return short if len(short) <= max_len else short[: max_len - 1] + "…"

    cleaned = re.sub(r"\b(s\.a\.?|s\.a\.s\.?|s\.r\.l\.?|ltd\.?|llc|inc\.?|corp\.?|corporation|company|import|imports|diagnostica|diagnostics|medical|medica|laboratorios|laboratorio)\b", "", text_name, flags=re.IGNORECASE)
    cleaned = re.sub(r"\s+", " ", cleaned).strip(" ,.-")
    if not cleaned:
        cleaned = text_name

    words = cleaned.split()
    if len(cleaned) <= max_len:
        return cleaned
    if len(words) >= 2:
        candidate = " ".join(words[:2]).strip()
        if len(candidate) <= max_len:
            return candidate
    return cleaned[: max_len - 1].rstrip() + "…"


def wrap_chart_title(text_value: str, width: int = 28) -> str:
    return "<br>".join(textwrap.wrap(safe_text(text_value, ""), width=width)) if safe_text(text_value, "") else ""

def build_long_palette(n: int) -> list[str]:
    base = [ACCENT, ACCENT_2, ACCENT_3, WARNING, "#9BB1FF", "#C084FC", "#F472B6", "#60A5FA", "#34D399", "#F59E0B", "#A78BFA", "#F87171", "#22D3EE", "#4ADE80"]
    if n <= len(base):
        return base[:n]
    repeats = (n // len(base)) + 1
    return (base * repeats)[:n]


def distributor_display_name(name: str, max_len: int = 22) -> str:
    text_name = safe_text(name, "No informado")
    return shorten_distributor_name(text_name, max_len=max_len)


def summarize_distributor_counts(summary_df: pd.DataFrame, top_n: int = 5) -> pd.DataFrame:
    """
    Mantiene solo los distribuidores más relevantes para la vista ejecutiva.
    No agrupa en "Otros"; simplemente corta el dataset al top_n para que la
    visual principal permanezca legible y ordenada.
    """
    if summary_df is None or summary_df.empty:
        return pd.DataFrame(columns=["Distributor name", "Count"])

    work = summary_df.copy()
    if "Distributor name" not in work.columns or "Count" not in work.columns:
        return work

    work["Distributor name"] = work["Distributor name"].fillna("No informado").astype(str)
    work["Count"] = pd.to_numeric(work["Count"], errors="coerce").fillna(0)

    work = (
        work.sort_values(["Count", "Distributor name"], ascending=[False, True])
        .reset_index(drop=True)
    )

    if isinstance(top_n, int) and top_n > 0:
        work = work.head(top_n).copy()

    return work.reset_index(drop=True)

def build_distributor_detail_table(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame(
            columns=[
                "Modelo",
                "Distribuidor",
                "Cantidad",
                "% del modelo",
                "% del total filtrado",
            ]
        )

    work = df.copy()
    work["Instrument type"] = work["Instrument type"].fillna("No informado").astype(str).str.strip()
    work["Distributor name"] = work["Distributor name"].fillna("No informado").astype(str).str.strip()

    summary = (
        work.groupby(["Instrument type", "Distributor name"], dropna=False)
        .size()
        .reset_index(name="Cantidad")
    )

    if summary.empty:
        return pd.DataFrame(
            columns=[
                "Modelo",
                "Distribuidor",
                "Cantidad",
                "% del modelo",
                "% del total filtrado",
            ]
        )

    total_filtered = int(summary["Cantidad"].sum())

    model_totals = (
        summary.groupby("Instrument type", as_index=False)["Cantidad"]
        .sum()
        .rename(columns={"Cantidad": "Total modelo"})
    )

    summary = summary.merge(model_totals, on="Instrument type", how="left")
    summary["% del modelo"] = (summary["Cantidad"] / summary["Total modelo"] * 100).round(1)
    summary["% del total filtrado"] = (summary["Cantidad"] / total_filtered * 100).round(1)

    summary = summary.rename(columns={"Instrument type": "Modelo", "Distributor name": "Distribuidor"})
    summary = summary.sort_values(by=["Modelo", "Cantidad", "Distribuidor"], ascending=[True, False, True]).reset_index(drop=True)

    return summary[["Modelo", "Distribuidor", "Cantidad", "% del modelo", "% del total filtrado"]]

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


def build_distributor_global_overview(df: pd.DataFrame, top_n: int = 5) -> go.Figure:
    fig = go.Figure()

    if df.empty:
        fig.update_layout(title='Vista global por distribuidor')
        return glow_layout(fig, 520, 16)

    work = df.copy()
    work['Instrument type'] = work['Instrument type'].fillna('No informado').astype(str)
    work['Distributor name'] = work['Distributor name'].fillna('No informado').astype(str)

    model_order = work['Instrument type'].value_counts().index.tolist()
    summary = work.groupby(['Instrument type', 'Distributor name'], dropna=False).size().reset_index(name='Count')
    if summary.empty:
        fig.update_layout(title='Vista global por distribuidor')
        return glow_layout(fig, 520, 16)

    dist_order = summary.groupby('Distributor name', as_index=False)['Count'].sum().sort_values(['Count', 'Distributor name'], ascending=[False, True])
    top_distributors = dist_order['Distributor name'].tolist()[:top_n]
    summary = summary[summary['Distributor name'].isin(top_distributors)].copy()
    label_map = {name: distributor_display_name(name, 18) for name in top_distributors}
    legend_order = [label_map[name] for name in top_distributors]
    summary['Legend label'] = summary['Distributor name'].astype(str).map(label_map)
    summary['Legend label'] = pd.Categorical(summary['Legend label'], categories=legend_order, ordered=True)
    summary['Instrument type'] = pd.Categorical(summary['Instrument type'], categories=model_order, ordered=True)
    summary = summary.sort_values(['Instrument type', 'Legend label', 'Count'], ascending=[True, True, False])
    palette = build_long_palette(len(top_distributors))

    fig = px.bar(
        summary,
        y='Instrument type',
        x='Count',
        color='Legend label',
        orientation='h',
        barmode='stack',
        text='Count',
        title='Vista global por distribuidor | resumen ejecutivo (Top 5)',
        category_orders={'Instrument type': model_order, 'Legend label': legend_order},
        color_discrete_sequence=palette,
        custom_data=['Instrument type', 'Distributor name', 'Count'],
    )
    fig.update_traces(
        textposition='inside',
        insidetextanchor='middle',
        hovertemplate='<b>Modelo:</b> %{customdata[0]}<br><b>Distribuidor:</b> %{customdata[1]}<br><b>Cantidad:</b> %{customdata[2]}<extra></extra>'
    )
    fig.update_layout(
        legend_title='Distribuidor',
        xaxis_title='Cantidad de equipos',
        yaxis_title='Modelo',
        margin=dict(t=72, b=48, l=8, r=8),
        height=520,
    )
    fig.update_yaxes(categoryorder='array', categoryarray=model_order[::-1])
    return glow_layout(fig, 520, 16)


def build_distributor_model_donut(df: pd.DataFrame, selected_model: str, top_n: int = 5) -> go.Figure:
    fig = go.Figure()

    if df.empty or not selected_model:
        fig.update_layout(title="Distribución por distribuidor")
        return glow_layout(fig, 430, 15)

    work = df.copy()
    work["Instrument type"] = work["Instrument type"].fillna("No informado").astype(str)
    work["Distributor name"] = work["Distributor name"].fillna("No informado").astype(str)
    model_df = work[work["Instrument type"] == selected_model].copy()

    if model_df.empty:
        fig.update_layout(title=selected_model)
        return glow_layout(fig, 430, 15)

    summary = (
        model_df.groupby("Distributor name", dropna=False)
        .size()
        .reset_index(name="Count")
        .sort_values(["Count", "Distributor name"], ascending=[False, True])
        .reset_index(drop=True)
    )

    if summary.empty:
        fig.update_layout(title=selected_model)
        return glow_layout(fig, 430, 15)

    summary = summarize_distributor_counts(summary, top_n=top_n)
    summary["Legend label"] = summary["Distributor name"].astype(str).map(lambda x: distributor_display_name(x, 20))
    summary = summary.sort_values(["Count", "Legend label"], ascending=[False, True]).reset_index(drop=True)
    palette = build_long_palette(len(summary))

    fig.add_trace(
        go.Pie(
            labels=summary["Legend label"],
            values=summary["Count"],
            hole=0.68,
            sort=False,
            marker=dict(colors=palette[:len(summary)], line=dict(color="rgba(255,255,255,0.20)", width=1.2)),
            textinfo="percent",
            textfont=dict(color="#ffffff", size=12),
            customdata=np.column_stack([summary["Distributor name"], summary["Count"]]),
            hovertemplate="<b>Modelo:</b> "
            + selected_model
            + "<br><b>Distribuidor:</b> %{customdata[0]}<br><b>Cantidad:</b> %{customdata[1]}<br><b>Participación:</b> %{percent}<extra></extra>",
        )
    )
    total_assets = int(model_df.shape[0])
    fig.add_annotation(
        text=f"<b>{total_assets:,}</b><br><span style='font-size:11px'>equipos</span>",
        x=0.5,
        y=0.52,
        xref="paper",
        yref="paper",
        showarrow=False,
        font=dict(color="#ffffff", size=17),
    )
    fig.update_layout(
        title=dict(text=wrap_chart_title(f"{selected_model} | Top 5", 26), x=0.03, y=0.96, xanchor="left", yanchor="top", font=dict(size=14, color="#f9fdff")),
        showlegend=True,
        height=430,
        margin=dict(t=72, b=96, l=8, r=8),
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=-0.10,
            xanchor="center",
            x=0.5,
            bgcolor="rgba(14,26,42,0.18)",
            bordercolor="rgba(124,221,255,0.16)",
            borderwidth=1,
            font=dict(color="#f8fbff", size=10),
            itemwidth=90,
            itemsizing="constant",
        ),
    )
    return glow_layout(fig, 430, 15)


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
    model_options = (
        filtered["Instrument type"]
        .fillna("No informado")
        .astype(str)
        .value_counts()
        .index
        .tolist()
    )
    if model_options:
        st.markdown(
            '<div class="small-note">Primero se muestra un resumen ejecutivo con los distribuidores más relevantes. Debajo aparecen gráficos circulares Top 5 por modelo. El detalle completo de todos los distribuidores se puede abrir más abajo sin saturar la vista principal.</div>',
            unsafe_allow_html=True,
        )
        st.plotly_chart(
            build_distributor_global_overview(filtered, top_n=5),
            use_container_width=True,
            key="global_distributor_overview_bar",
        )

        cards_per_row = 3
        for start in range(0, len(model_options), cards_per_row):
            row_models = model_options[start:start + cards_per_row]
            cols = st.columns(len(row_models))
            for col, model_name in zip(cols, row_models):
                with col:
                    st.plotly_chart(
                        build_distributor_model_donut(filtered, model_name, top_n=5),
                        use_container_width=True,
                        key=f"donut_model_distributor_{normalize_key_text(model_name)}",
                    )

        with st.expander("Ver detalle completo de todos los distribuidores por modelo", expanded=False):
            st.markdown(
                '<div class="small-note">Esta vista despliega el detalle completo sin resumir. Se usa una barra horizontal por modelo porque comunica mejor que un donut cuando hay muchos distribuidores.</div>',
                unsafe_allow_html=True,
            )
            st.dataframe(build_distributor_detail_table(filtered), use_container_width=True, hide_index=True)
            for model_name in model_options:
                st.plotly_chart(
                    build_distributor_detail_bar(filtered, model_name),
                    use_container_width=True,
                    key=f"detail_bar_model_{normalize_key_text(model_name)}",
                )

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
