from __future__ import annotations

from io import BytesIO, StringIO
from pathlib import Path
from datetime import datetime
import csv
import hashlib
import re

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st


# =========================================================
# CONFIG
# =========================================================
st.set_page_config(
    page_title="Records List Intelligence Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

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

DEFAULT_FOLDER = r"C:\Users\javier.avellaneda\OneDrive - Diasorin-Luminex\Documentos\Service DashBoard Backup\Python"

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

APP_CSS = """
<style>
.block-container {
    padding-top: 1rem;
    padding-bottom: 1rem;
    max-width: 98rem;
}
[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #0e1528 0%, #0b1220 100%);
}
.main-title {
    font-size: 2rem;
    font-weight: 800;
    color: #EAF2FF;
    margin-bottom: 0.25rem;
}
.subtle {
    color: #A9B9D6;
    font-size: 0.95rem;
}
.kpi-card {
    background: linear-gradient(180deg, #111827 0%, #0f172a 100%);
    border: 1px solid rgba(255,255,255,0.08);
    border-radius: 18px;
    padding: 16px 18px;
    box-shadow: 0 10px 28px rgba(0,0,0,0.22);
}
.kpi-label {
    color: #A9B9D6;
    font-size: 0.88rem;
    margin-bottom: 4px;
}
.kpi-value {
    color: #EAF2FF;
    font-size: 2rem;
    font-weight: 800;
    line-height: 1.1;
}
.kpi-subtitle {
    color: #7c93b7;
    font-size: 0.85rem;
    margin-top: 5px;
}
.section-title {
    color: #EAF2FF;
    font-size: 1.20rem;
    font-weight: 700;
    margin-top: 0.2rem;
    margin-bottom: 0.55rem;
}
hr {
    border-color: rgba(255,255,255,0.08);
}
.small-note {
    color: #93a7c9;
    font-size: 0.85rem;
}
</style>
"""
st.markdown(APP_CSS, unsafe_allow_html=True)


# =========================================================
# HELPERS
# =========================================================
def glow_layout(fig: go.Figure, height: int = 420, title_size: int = 18) -> go.Figure:
    fig.update_layout(
        template=PLOT_TEMPLATE,
        height=height,
        paper_bgcolor=PLOT_BG,
        plot_bgcolor=PLOT_BG,
        margin=dict(l=20, r=20, t=82, b=20),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        font=dict(color=TEXT),
        title_font=dict(size=title_size),
        hovermode="closest",
        hoverlabel=dict(bgcolor="#0d1228", font=dict(color=TEXT)),
    )
    fig.update_xaxes(showgrid=True, gridcolor=GRID, zeroline=False, automargin=True)
    fig.update_yaxes(showgrid=True, gridcolor=GRID, zeroline=False, automargin=True)
    return fig


def metric_card(label: str, value: str, subtitle: str = "") -> None:
    st.markdown(
        f"""
        <div class="kpi-card">
            <div class="kpi-label">{label}</div>
            <div class="kpi-value">{value}</div>
            <div class="kpi-subtitle">{subtitle}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def dataframe_to_excel_bytes(sheet_map: dict[str, pd.DataFrame]) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sheet_name, df in sheet_map.items():
            safe_name = re.sub(r"[\\/*?:\\[\\]]", "_", str(sheet_name))[:31] or "Sheet1"
            clean_df = df.copy()
            clean_df.to_excel(writer, sheet_name=safe_name, index=False)
            ws = writer.sheets[safe_name]
            ws.freeze_panes = "A2"

            for idx, col in enumerate(clean_df.columns, start=1):
                max_len = len(str(col))
                if not clean_df.empty:
                    series = clean_df[col].astype(str).replace("", "").replace("nan", "").replace("<NA>", "")
                    try:
                        max_len = max(max_len, int(series.map(len).max()))
                    except Exception:
                        pass
                ws.column_dimensions[ws.cell(row=1, column=idx).column_letter].width = min(max(max_len + 2, 12), 42)

    output.seek(0)
    return output.getvalue()


def safe_ratio(n: float, d: float) -> float:
    if d in (0, None) or pd.isna(d):
        return 0.0
    return float(n) / float(d)


def normalize_text_series(series: pd.Series) -> pd.Series:
    return (
        series.astype("string")
        .str.strip()
        .replace({"": pd.NA, "nan": pd.NA, "None": pd.NA, "<NA>": pd.NA})
    )


def safe_unique(df: pd.DataFrame, col: str) -> list[str]:
    if col not in df.columns:
        return []
    vals = (
        df[col]
        .dropna()
        .astype("string")
        .str.strip()
        .replace("", pd.NA)
        .dropna()
        .unique()
        .tolist()
    )
    return sorted(vals, key=lambda x: str(x).lower())


def apply_multiselect_filter(df: pd.DataFrame, col: str, selected: list[str]) -> pd.DataFrame:
    if col not in df.columns or not selected:
        return df.copy()
    return df[df[col].isin(selected)].copy()


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


def gauge_figure(value: int, total: int, title: str, color: str = ACCENT) -> go.Figure:
    total = max(int(total), 1)
    fig = go.Figure(
        go.Indicator(
            mode="gauge+number",
            value=int(value),
            number={"font": {"size": 38, "color": TEXT}},
            title={"text": title, "font": {"size": 18, "color": TEXT}},
            gauge={
                "axis": {"range": [0, total], "tickcolor": MUTED},
                "bar": {"color": color},
                "bgcolor": "rgba(255,255,255,0.06)",
                "borderwidth": 0,
                "steps": [{"range": [0, total], "color": "rgba(255,255,255,0.08)"}],
            },
        )
    )
    fig.update_layout(
        height=320,
        paper_bgcolor=PLOT_BG,
        plot_bgcolor=PLOT_BG,
        margin=dict(l=20, r=20, t=60, b=20),
        font=dict(color=TEXT),
    )
    return fig


# =========================================================
# FILE LOGIC
# =========================================================
def discover_files(folder: str) -> list[Path]:
    base = Path(folder)
    if not base.exists():
        return []

    files: list[Path] = []
    for ext in ("*.xlsx", "*.xls", "*.csv"):
        files.extend(base.glob(ext))

    return sorted(
        [f for f in files if f.is_file() and not f.name.startswith("~$")],
        key=lambda x: x.name.lower(),
    )


def score_records_file(path: Path) -> int:
    name = path.name.lower()
    score = 0

    if path.suffix.lower() in [".xlsx", ".xls"]:
        score += 80
    elif path.suffix.lower() == ".csv":
        score += 20

    if "records" in name:
        score += 35
    if "list" in name:
        score += 25
    if "report" in name:
        score += 20
    if "installed" in name:
        score += 20
    if "isr" in name:
        score += 22
    if "base" in name:
        score += 15
    if "fecha_fabricacion" in name:
        score += 12

    if "spare" in name or "parts" in name or "price" in name or "carstock" in name:
        score -= 40

    return score


def pick_best_records_file(files: list[Path]) -> Path | None:
    if not files:
        return None
    ranked = sorted(files, key=score_records_file, reverse=True)
    return ranked[0] if ranked else None


def read_csv_robust_from_path(path: Path) -> pd.DataFrame:
    encodings = ["utf-8-sig", "utf-8", "latin1", "cp1252"]
    seps: list[str | None] = [None, ";", ",", "\t", "|"]

    for enc in encodings:
        try:
            sample = path.read_text(encoding=enc, errors="ignore")[:5000]
            if sample.strip():
                try:
                    dialect = csv.Sniffer().sniff(sample, delimiters=";,|\t,")
                    sniffed_sep = dialect.delimiter
                    df = pd.read_csv(
                        path,
                        sep=sniffed_sep,
                        encoding=enc,
                        engine="python",
                        on_bad_lines="skip",
                    )
                    if df.shape[1] >= 3:
                        return df
                except Exception:
                    pass
        except Exception:
            pass

    for enc in encodings:
        for sep in seps:
            try:
                df = pd.read_csv(
                    path,
                    sep=sep,
                    encoding=enc,
                    engine="python",
                    on_bad_lines="skip",
                )
                if df.shape[1] >= 3:
                    return df
            except Exception:
                continue

    for enc in encodings:
        try:
            df = pd.read_csv(
                path,
                encoding=enc,
                engine="python",
                sep=None,
                on_bad_lines="skip",
            )
            return df
        except Exception:
            continue

    raise ValueError(f"No fue posible leer el CSV: {path.name}")


def read_table_from_path(path: Path) -> pd.DataFrame:
    suffix = path.suffix.lower()

    if suffix == ".csv":
        return read_csv_robust_from_path(path)

    if suffix in [".xlsx", ".xls"]:
        xls = pd.ExcelFile(path)
        best_sheet = None
        best_score = -1

        for sheet in xls.sheet_names:
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
            if "sheet1" in s:
                score += 2
            if score > best_score:
                best_score = score
                best_sheet = sheet

        if best_sheet is None:
            best_sheet = xls.sheet_names[0]

        return pd.read_excel(path, sheet_name=best_sheet)

    raise ValueError(f"Formato no soportado: {path.name}")


def get_uploaded_file_signature(uploaded_file) -> str:
    if uploaded_file is None:
        return ""
    content = uploaded_file.getvalue()
    raw = f"{uploaded_file.name}|{len(content)}|".encode("utf-8") + content
    return hashlib.md5(raw).hexdigest()


def read_uploaded_table_robust(uploaded_file) -> pd.DataFrame:
    if uploaded_file is None:
        return pd.DataFrame()

    file_name = uploaded_file.name.lower()

    if file_name.endswith(".xlsx") or file_name.endswith(".xls"):
        xls = pd.ExcelFile(uploaded_file)
        best_sheet = None
        best_score = -1

        for sheet in xls.sheet_names:
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
                best_sheet = sheet

        if best_sheet is None:
            best_sheet = xls.sheet_names[0]

        return pd.read_excel(uploaded_file, sheet_name=best_sheet)

    if file_name.endswith(".csv"):
        raw_bytes = uploaded_file.getvalue()
        encodings = ["utf-8-sig", "utf-8", "latin1", "cp1252"]
        seps = [None, ";", ",", "\t", "|"]

        for enc in encodings:
            try:
                sample = raw_bytes.decode(enc, errors="ignore")[:5000]
                if sample.strip():
                    try:
                        dialect = csv.Sniffer().sniff(sample, delimiters=";,|\t,")
                        sniffed_sep = dialect.delimiter
                        df = pd.read_csv(
                            StringIO(raw_bytes.decode(enc, errors="ignore")),
                            sep=sniffed_sep,
                            engine="python",
                            on_bad_lines="skip",
                        )
                        if df.shape[1] >= 3:
                            return df
                    except Exception:
                        pass
            except Exception:
                pass

        for enc in encodings:
            for sep in seps:
                try:
                    text = raw_bytes.decode(enc, errors="ignore")
                    df = pd.read_csv(
                        StringIO(text),
                        sep=sep,
                        engine="python",
                        on_bad_lines="skip",
                    )
                    if df.shape[1] >= 3:
                        return df
                except Exception:
                    continue

        raise ValueError(f"No fue posible leer el CSV subido: {uploaded_file.name}")

    raise ValueError(f"Formato no soportado: {uploaded_file.name}")


def load_default_records_df(folder: str) -> tuple[pd.DataFrame, str]:
    files = discover_files(folder)
    chosen = pick_best_records_file(files)
    if chosen is None:
        return pd.DataFrame(), ""
    df = read_table_from_path(chosen)
    return df, chosen.name


def get_active_records_df(folder: str) -> tuple[pd.DataFrame, str]:
    uploaded_file = st.session_state.get("records_list_upload_widget")

    if uploaded_file is not None:
        current_signature = get_uploaded_file_signature(uploaded_file)
        saved_signature = st.session_state.get("records_list_upload_signature", "")

        if "records_list_upload_df" not in st.session_state or current_signature != saved_signature:
            uploaded_df = read_uploaded_table_robust(uploaded_file)
            st.session_state["records_list_upload_df"] = uploaded_df.copy()
            st.session_state["records_list_upload_signature"] = current_signature
            st.session_state["records_list_upload_name"] = uploaded_file.name

        return (
            st.session_state["records_list_upload_df"].copy(),
            f"{st.session_state['records_list_upload_name']} (uploaded)",
        )

    default_df, default_name = load_default_records_df(folder)
    return default_df, default_name


# =========================================================
# DATA PREP
# =========================================================
COLUMN_ALIASES = {
    "Distributor name": ["Distributor name", "Distributor", "Distributor Name"],
    "Instrument type": ["Instrument type", "Instrument Type", "Type"],
    "Installation date": ["Installation date", "Install date", "Installation Date"],
    "Customer name": ["Customer name", "Customer Name", "Customer"],
    "In Blood Bank": ["In Blood Bank"],
    "Address": ["Address"],
    "ZipCode": ["ZipCode", "Zip Code"],
    "City": ["City", "Ciudad"],
    "Country": ["Country", "País", "Pais"],
    "World Region": ["World Region"],
    "Commercial Region": ["Commercial Region", "Commercial region", "Region"],
    "Latitude": ["Latitude", "Lat", "LATITUDE"],
    "Longitude": ["Longitude", "Lon", "Long", "LONGITUDE"],
    "Product Line": ["Product Line", "Product line"],
    "Serial number": ["Serial number", "Serial Number", "SN", "Serial"],
    "Machine Configurations": ["Machine Configurations", "Machine Configuration", "Configuration"],
    "Asset condition": ["Asset condition", "Asset Condition"],
    "PM plan": ["PM plan", "PM Plan"],
    "Number of tests per day": ["Number of tests per day", "Tests/day", "Tests per day"],
    "Operational status": ["Operational status", "Status", "Operational Status"],
    "Type of contract": ["Type of contract", "Contract type", "Type Contract"],
    "Contract duration": ["Contract duration"],
    "Tag": ["Tag"],
    "Notes": ["Notes", "Note"],
    "PM last date": ["PM last date", "PM Last Date"],
    "PM frequency": ["PM frequency", "PM Frequency"],
    "PM next date": ["PM next date", "PM Next Date"],
    "PM performed On": ["PM performed On", "PM Performed On"],
}

BASE_REQUIRED_COLS = [
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
]


def harmonize_columns(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    rename_map = {}

    for canonical, aliases in COLUMN_ALIASES.items():
        for alias in aliases:
            if alias in out.columns:
                if alias != canonical:
                    rename_map[alias] = canonical
                break

    if rename_map:
        out = out.rename(columns=rename_map)

    return out


def prepare_records_df(df: pd.DataFrame) -> pd.DataFrame:
    out = harmonize_columns(df.copy())

    for col in BASE_REQUIRED_COLS:
        if col not in out.columns:
            out[col] = pd.NA

    text_cols = [
        "Distributor name",
        "Instrument type",
        "Customer name",
        "In Blood Bank",
        "Address",
        "ZipCode",
        "City",
        "Country",
        "World Region",
        "Commercial Region",
        "Product Line",
        "Serial number",
        "Machine Configurations",
        "Asset condition",
        "PM plan",
        "Operational status",
        "Type of contract",
        "Contract duration",
        "Tag",
        "Notes",
    ]
    for col in text_cols:
        out[col] = normalize_text_series(out[col])

    date_cols = ["Installation date", "PM last date", "PM next date", "PM performed On"]
    for col in date_cols:
        out[col] = pd.to_datetime(out[col], errors="coerce")

    num_cols = ["Latitude", "Longitude", "Number of tests per day", "PM frequency"]
    for col in num_cols:
        out[col] = pd.to_numeric(out[col], errors="coerce")

    out["Distributor Clean"] = out["Distributor name"].fillna("Data not available")
    out["Instrument Clean"] = out["Instrument type"].fillna("Data not available")
    out["Customer Clean"] = out["Customer name"].fillna("Data not available")
    out["City Clean"] = out["City"].fillna("Data not available")
    out["Country Clean"] = out["Country"].fillna("Data not available")
    out["Region Clean"] = out["Commercial Region"].fillna("Data not available")
    out["Status Clean"] = out["Operational status"].fillna("Data not available")
    out["Installation Year"] = out["Installation date"].dt.year

    return out


# =========================================================
# FILTERS
# =========================================================
def build_global_filters(df: pd.DataFrame) -> pd.DataFrame:
    c1, c2, c3, c4 = st.columns([1.2, 1.2, 1.2, 1.0])

    with c1:
        region_options = safe_unique(df, "Region Clean")
        selected_region = st.multiselect(
            "Commercial Region",
            options=region_options,
            default=region_options,
            key="flt_region",
        )
    df1 = apply_multiselect_filter(df, "Region Clean", selected_region)

    with c2:
        country_options = safe_unique(df1, "Country Clean")
        selected_country = st.multiselect(
            "Country",
            options=country_options,
            default=country_options,
            key="flt_country",
        )
    df2 = apply_multiselect_filter(df1, "Country Clean", selected_country)

    with c3:
        distributor_options = safe_unique(df2, "Distributor Clean")
        selected_distributor = st.multiselect(
            "Distributor name",
            options=distributor_options,
            default=distributor_options,
            key="flt_distributor",
        )
    df3 = apply_multiselect_filter(df2, "Distributor Clean", selected_distributor)

    with c4:
        city_options = ["Todas"] + safe_unique(df3, "City Clean")
        city = st.selectbox("City", city_options, index=0, key="flt_city")

    if city != "Todas":
        df3 = df3[df3["City Clean"] == city].copy()

    return df3


# =========================================================
# CHARTS
# =========================================================
def fig_instrument_by_type(df: pd.DataFrame) -> go.Figure:
    counts = (
        df["Instrument Clean"]
        .value_counts(dropna=False)
        .rename_axis("Instrument type")
        .reset_index(name="Count")
        .sort_values("Count", ascending=True)
    )
    fig = px.bar(counts, x="Count", y="Instrument type", orientation="h", text="Count")
    fig.update_traces(marker_color=ACCENT, textposition="inside")
    fig.update_layout(title="Instrumentos por tipo", showlegend=False)
    return glow_layout(fig, height=360)


def fig_instrument_by_status(df: pd.DataFrame) -> go.Figure:
    counts = (
        df["Status Clean"]
        .value_counts(dropna=False)
        .rename_axis("Status")
        .reset_index(name="Count")
        .sort_values("Count", ascending=True)
    )
    fig = px.bar(counts, x="Count", y="Status", orientation="h", text="Count")
    fig.update_traces(marker_color=ACCENT_2, textposition="outside")
    fig.update_layout(title="Instrumentos por status", showlegend=False)
    return glow_layout(fig, height=360)


def fig_installations_per_year(df: pd.DataFrame) -> go.Figure:
    tmp = df.dropna(subset=["Installation Year"]).copy()
    if tmp.empty:
        fig = go.Figure()
        fig.update_layout(title="Instalaciones por año")
        return glow_layout(fig, height=320)

    counts = (
        tmp["Installation Year"]
        .astype(int)
        .value_counts()
        .sort_index()
        .rename_axis("Year")
        .reset_index(name="Installations")
    )
    fig = px.bar(counts, x="Year", y="Installations", text="Installations")
    fig.update_traces(marker_color=ACCENT_3)
    fig.update_layout(title="Instalaciones por año", showlegend=False)
    return glow_layout(fig, height=320)


def fig_ib_map(df: pd.DataFrame) -> go.Figure:
    geo = df.dropna(subset=["Latitude", "Longitude"]).copy()
    if geo.empty:
        fig = go.Figure()
        fig.update_layout(title="Mapa de base instalada")
        return glow_layout(fig, height=420)

    map_df = (
        geo.groupby(["Country Clean", "Latitude", "Longitude"], dropna=False)
        .agg(
            Instruments=("Serial number", "count"),
            Distributor=("Distributor Clean", lambda s: ", ".join(sorted(set(s.dropna().astype(str).tolist()))[:4])),
        )
        .reset_index()
    )

    center, zoom = compute_mapbox_center_zoom(map_df, "Latitude", "Longitude")

    fig = px.scatter_mapbox(
        map_df,
        lat="Latitude",
        lon="Longitude",
        size="Instruments",
        hover_name="Country Clean",
        hover_data={"Distributor": True, "Instruments": True, "Latitude": False, "Longitude": False},
        zoom=zoom,
        center=center,
        size_max=36,
    )
    fig.update_traces(marker=dict(color=ACCENT, opacity=0.78))
    fig.update_layout(
        title="Mapa de base instalada",
        mapbox_style="carto-positron",
        paper_bgcolor=PLOT_BG,
        plot_bgcolor=PLOT_BG,
        font=dict(color=TEXT),
        margin=dict(l=20, r=20, t=55, b=10),
        height=420,
    )
    return fig


def fig_distributors_instruments_hover(df: pd.DataFrame) -> go.Figure:
    if df.empty:
        fig = go.Figure()
        fig.update_layout(title="Instrumentos por distribuidor")
        return glow_layout(fig, height=560)

    work = df.copy()

    for col in [
        "Distributor Clean",
        "Instrument Clean",
        "Serial number",
        "Customer Clean",
        "City Clean",
        "Status Clean",
        "Machine Configurations",
    ]:
        if col not in work.columns:
            work[col] = "Data not available"

    grouped = (
        work.groupby(["Distributor Clean", "Instrument Clean"], dropna=False)
        .agg(
            Count=("Serial number", "count"),
            Serials=("Serial number", lambda s: "<br>".join(sorted(set(s.dropna().astype(str).tolist()))[:20])),
            Customers=("Customer Clean", lambda s: "<br>".join(sorted(set(s.dropna().astype(str).tolist()))[:12])),
            Cities=("City Clean", lambda s: "<br>".join(sorted(set(s.dropna().astype(str).tolist()))[:12])),
            Statuses=("Status Clean", lambda s: "<br>".join(sorted(set(s.dropna().astype(str).tolist()))[:12])),
            Configs=("Machine Configurations", lambda s: "<br>".join(sorted(set(s.dropna().astype(str).tolist()))[:10])),
        )
        .reset_index()
    )

    totals = (
        grouped.groupby("Distributor Clean", as_index=False)["Count"]
        .sum()
        .sort_values("Count", ascending=False)
    )

    distributor_order = totals["Distributor Clean"].tolist()
    grouped["Distributor Clean"] = pd.Categorical(
        grouped["Distributor Clean"],
        categories=distributor_order,
        ordered=True,
    )
    grouped = grouped.sort_values(["Distributor Clean", "Instrument Clean"])

    fig = px.bar(
        grouped,
        x="Distributor Clean",
        y="Count",
        color="Instrument Clean",
        barmode="stack",
        custom_data=[
            "Instrument Clean",
            "Count",
            "Serials",
            "Customers",
            "Cities",
            "Statuses",
            "Configs",
        ],
        title="Instrumentos por distribuidor",
    )

    fig.update_traces(
        hovertemplate=(
            "<b>Distributor:</b> %{x}<br>"
            "<b>Instrument type:</b> %{customdata[0]}<br>"
            "<b>Total instruments:</b> %{customdata[1]}<br><br>"
            "<b>Serial numbers:</b><br>%{customdata[2]}<br><br>"
            "<b>Customers:</b><br>%{customdata[3]}<br><br>"
            "<b>Cities:</b><br>%{customdata[4]}<br><br>"
            "<b>Status:</b><br>%{customdata[5]}<br><br>"
            "<b>Machine Configurations:</b><br>%{customdata[6]}<extra></extra>"
        )
    )

    fig.update_layout(
        xaxis_title="Distributor",
        yaxis_title="Cantidad de instrumentos",
        legend_title="Instrument type",
        xaxis=dict(tickangle=-35),
    )

    return glow_layout(fig, height=560, title_size=16)


# =========================================================
# UI
# =========================================================
def show_header(active_file_name: str) -> None:
    st.markdown('<div class="main-title">Records List Intelligence Dashboard</div>', unsafe_allow_html=True)
    st.markdown(
        f'<div class="subtle">Archivo activo: <b>{active_file_name or "No detectado"}</b></div>',
        unsafe_allow_html=True,
    )
    st.markdown("<hr>", unsafe_allow_html=True)


def sidebar_controls() -> str:
    st.sidebar.markdown("## Fuente de datos")
    folder = st.sidebar.text_input(
        "Carpeta de trabajo",
        value=DEFAULT_FOLDER,
        help="Si no subes un archivo, la app buscará automáticamente el mejor Records List en esta carpeta.",
    )

    st.sidebar.markdown("### Base instalada activa")
    records_upload = st.sidebar.file_uploader(
        "Sube el último Records List / ISR Live",
        type=["xlsx", "xls", "csv"],
        key="records_list_upload_widget",
        help="La app mantendrá este archivo como fuente activa mientras la sesión siga abierta.",
    )

    if records_upload is not None:
        st.sidebar.success(f"Usando último archivo subido: {records_upload.name}")
    elif st.session_state.get("records_list_upload_name"):
        st.sidebar.info(f"Archivo activo en sesión: {st.session_state['records_list_upload_name']}")

    st.sidebar.markdown("---")
    st.sidebar.caption("Primer paquete: último archivo subido + gráfica por distribuidor")
    return folder


# =========================================================
# MAIN PAGE
# =========================================================
def render_main_dashboard(df: pd.DataFrame, active_name: str) -> None:
    show_header(active_name)

    filtered = build_global_filters(df)

    total_ib = len(filtered)
    total_distributors = filtered["Distributor Clean"].nunique() if "Distributor Clean" in filtered.columns else 0
    total_countries = filtered["Country Clean"].nunique() if "Country Clean" in filtered.columns else 0
    total_types = filtered["Instrument Clean"].nunique() if "Instrument Clean" in filtered.columns else 0

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        metric_card("TOTAL IB", f"{total_ib:,}", "Filtrado actual")
    with c2:
        metric_card("DISTRIBUTORS", f"{total_distributors:,}", "Visibles en el filtro")
    with c3:
        metric_card("COUNTRIES", f"{total_countries:,}", "Visibles en el filtro")
    with c4:
        metric_card("INSTRUMENT TYPES", f"{total_types:,}", "Tipos visibles")

    st.markdown("<div class='section-title'>Vista general</div>", unsafe_allow_html=True)
    r1c1, r1c2, r1c3 = st.columns([1.0, 1.15, 1.0])
    with r1c1:
        st.plotly_chart(fig_instrument_by_type(filtered), use_container_width=True)
    with r1c2:
        st.plotly_chart(fig_ib_map(filtered), use_container_width=True)
    with r1c3:
        st.plotly_chart(fig_instrument_by_status(filtered), use_container_width=True)

    st.markdown("<div class='section-title'>Instrumentos por distribuidor</div>", unsafe_allow_html=True)
    st.markdown(
        "<div class='small-note'>Pasa el mouse sobre cada segmento para ver seriales, clientes, ciudades, status y machine configurations.</div>",
        unsafe_allow_html=True,
    )
    st.plotly_chart(fig_distributors_instruments_hover(filtered), use_container_width=True)

    st.markdown("<div class='section-title'>Instalaciones por año</div>", unsafe_allow_html=True)
    st.plotly_chart(fig_installations_per_year(filtered), use_container_width=True)

    st.markdown("<div class='section-title'>Detalle filtrado</div>", unsafe_allow_html=True)
    detail_cols = [
        "Distributor name",
        "Instrument type",
        "Customer name",
        "City",
        "Country",
        "Serial number",
        "Installation date",
        "Operational status",
        "Machine Configurations",
        "Number of tests per day",
    ]
    detail_cols = [c for c in detail_cols if c in filtered.columns]

    export_bytes = dataframe_to_excel_bytes({"Filtered Records": filtered[detail_cols]})
    st.download_button(
        "⬇️ Exportar detalle filtrado a Excel",
        data=export_bytes,
        file_name=f"filtered_records_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    st.dataframe(filtered[detail_cols], use_container_width=True, height=420)


def main() -> None:
    folder = sidebar_controls()

    try:
        raw_df, active_name = get_active_records_df(folder)
    except Exception as e:
        st.error(f"Error cargando la base instalada: {e}")
        st.stop()

    if raw_df.empty:
        st.error("No encontré una base instalada válida.")
        st.info("Sube un Records List / ISR Live o deja un archivo Excel/CSV válido en la carpeta configurada.")
        st.stop()

    records_df = prepare_records_df(raw_df)
    render_main_dashboard(records_df, active_name)


if __name__ == "__main__":
    main()
