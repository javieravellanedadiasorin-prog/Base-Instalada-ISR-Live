from __future__ import annotations

from pathlib import Path
from io import BytesIO
from datetime import datetime
import csv
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
    page_title="Diasorin Installed Base Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

BASE_DIR = Path(__file__).resolve().parent

PLOT_BG = "rgba(0,0,0,0)"
GRID = "rgba(255,255,255,0.10)"
TEXT = "#EAF2FF"
MUTED = "#A9B9D6"
ACCENT = "#3BA3FF"
ACCENT_2 = "#7A5CFF"
ACCENT_3 = "#00C389"
WARNING = "#FFB454"
DANGER = "#FF5D8F"

PAGE_INSTALLED_BASE = "Installed Base"
PAGE_IB_COUNTRY_DISTRIBUTOR = "IB by Country/Distributor"
PAGE_CLIA_AGE = "CLIA Age Analysis"
PAGE_LXL_MACHINE_CONFIG = "LXL Machine Configuration"
PAGE_WIN10 = "WIN 10"
PAGE_LXS_MACHINE_CONFIG = "LXS Machine Configuration"
PAGE_SP_PRICE_LIST = "SP Price List"
PAGE_PARAMETERS_CLIA = "Parameters CLIA"

ALL_PAGES = [
    PAGE_INSTALLED_BASE,
    PAGE_IB_COUNTRY_DISTRIBUTOR,
    PAGE_CLIA_AGE,
    PAGE_LXL_MACHINE_CONFIG,
    PAGE_WIN10,
    PAGE_LXS_MACHINE_CONFIG,
    PAGE_SP_PRICE_LIST,
    PAGE_PARAMETERS_CLIA,
]

APP_CSS = """
<style>
.block-container {
    padding-top: 1.0rem;
    padding-bottom: 1.0rem;
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
</style>
"""
st.markdown(APP_CSS, unsafe_allow_html=True)


# =========================================================
# HELPERS
# =========================================================
def glow_layout(fig: go.Figure, height: int = 420, title_size: int = 18) -> go.Figure:
    fig.update_layout(
        height=height,
        paper_bgcolor=PLOT_BG,
        plot_bgcolor=PLOT_BG,
        margin=dict(l=20, r=20, t=70, b=20),
        font=dict(color=TEXT),
        title_font=dict(size=title_size),
        legend=dict(orientation="v", font=dict(color=TEXT)),
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
                    series = clean_df[col].astype(str).replace("nan", "").replace("<NA>", "")
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


def first_existing_col(df: pd.DataFrame, candidates: list[str]) -> str | None:
    for c in candidates:
        if c in df.columns:
            return c
    return None


def normalize_text(series: pd.Series) -> pd.Series:
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
    if not selected or col not in df.columns:
        return df.copy()
    return df[df[col].isin(selected)].copy()


# =========================================================
# COLUMN MAPPING
# =========================================================
COLUMN_ALIASES = {
    "Distributor name": ["Distributor name", "Distributor", "Distributor Name"],
    "Instrument type": ["Instrument type", "Instrument Type", "Type"],
    "Installation date": ["Installation date", "Install date", "Installation Date"],
    "Customer name": ["Customer name", "Customer Name", "Customer"],
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
    "User Software": ["User Software", "SW Version", "Software Version", "Software"],
    "LXL User Software": ["LXL User Software"],
    "LXS User Software": ["LXS User Software"],
    "Operating System": ["Operating System", "OS", "OS (Original)", "OS Original"],
    "PC Model": ["PC Model", "PC Model (Original)", "PC Model Original"],
    "Win10 Status": ["Win10 Status", "Win10 Status (Original)"],
    "EDID Emulator": ["EDID Emulator", "EDID Emulator Status"],
    "System in Altitude": ["System in Altitude"],
    "Automation partner": ["Automation partner"],
    "LXS Tank configuration": ["LXS Tank configuration", "XS Tank configuration"],
    "Carstock": ["Carstock"],
    "Part Number": ["Part Number", "PN", "Part number"],
    "PART NUMBER DESCRIPTION": ["PART NUMBER DESCRIPTION", "Part Description", "Description"],
    "PN Revision": ["PN Revision", "Revision"],
    "Type": ["Type"],
    "SP Price (Option 1)": ["SP Price (Option 1)", "SP Price Option 1"],
    "SP Price (Option 2)": ["SP Price (Option 2)", "SP Price Option 2"],
    "SP Price (Option 3)": ["SP Price (Option 3)", "SP Price Option 3"],
    "Parts per system (12 months)": ["Parts per system (12 months)"],
    "Minimum Stock Level Required": ["Minimum Stock Level Required"],
    "On Hands": ["On Hands", "On Hand"],
    "QTY Delta": ["QTY Delta", "Qty Delta"],
    "Delta QTY Filter": ["Delta QTY Filter"],
    "PO Frequency (weeks)": ["PO Frequency (weeks)", "PO Frequency"],
    "Instrument Type": ["Instrument Type"],
}

REQUIRED_BASE_COLS = [
    "Distributor name",
    "Instrument type",
    "Installation date",
    "Customer name",
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
    "User Software",
    "LXL User Software",
    "LXS User Software",
    "Operating System",
    "PC Model",
    "Win10 Status",
    "EDID Emulator",
    "System in Altitude",
    "Automation partner",
    "LXS Tank configuration",
]

READY_TO_INSTALL_STATUSES = {
    "WAREHOUSE ready to be installed",
    "WAREHOUSE new system",
    "WAREHOUSE to be refurbished",
    "WAREHOUSE TRANSIT",
}

XL_TYPES = {"LIAISON XL", "LIAISON XL LAS"}
XS_TYPES = {"LIAISON XS"}


# =========================================================
# FILE DISCOVERY / LOADING
# =========================================================
@st.cache_data(show_spinner=False)
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


def score_installed_base_file(path: Path) -> int:
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
        score += 25
    if "installed" in name:
        score += 22
    if "isr" in name:
        score += 22
    if "base" in name:
        score += 18
    if "fecha_fabricacion" in name:
        score += 12

    if "spare" in name or "parts" in name or "price" in name or "carstock" in name:
        score -= 40

    return score


def score_spare_parts_file(path: Path) -> int:
    name = path.name.lower()
    score = 0

    if path.suffix.lower() in [".xlsx", ".xls"]:
        score += 60
    elif path.suffix.lower() == ".csv":
        score += 15

    if "spare" in name:
        score += 35
    if "parts" in name:
        score += 25
    if "price" in name:
        score += 25
    if "carstock" in name:
        score += 25
    if "tp" in name:
        score += 15
    if "new tp" in name:
        score += 15

    if "records" in name or "installed" in name or "isr" in name:
        score -= 30

    return score


def read_csv_robust(path: Path) -> pd.DataFrame:
    encodings = ["utf-8-sig", "utf-8", "latin1", "cp1252"]
    seps_to_try: list[str | None] = [None, ";", ",", "\t", "|"]

    # Intento con sniffing real
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

    # Intentos controlados
    for enc in encodings:
        for sep in seps_to_try:
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

    # Último recurso
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

    raise ValueError(f"No fue posible leer el CSV de forma robusta: {path.name}")


def load_any_table(path: Path) -> pd.DataFrame:
    suffix = path.suffix.lower()

    if suffix == ".csv":
        return read_csv_robust(path)

    if suffix in [".xlsx", ".xls"]:
        xls = pd.ExcelFile(path)
        best_sheet = None
        best_score = -1

        for sheet in xls.sheet_names:
            s = sheet.lower()
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
                score += 16
            if "price" in s:
                score += 12
            if "carstock" in s:
                score += 12
            if "sheet1" in s:
                score += 2
            if score > best_score:
                best_score = score
                best_sheet = sheet

        if best_sheet is None:
            best_sheet = xls.sheet_names[0]

        return pd.read_excel(path, sheet_name=best_sheet)

    raise ValueError(f"Formato no soportado: {path.name}")


def pick_best_file(files: list[Path], scorer) -> Path | None:
    if not files:
        return None
    ranked = sorted(files, key=scorer, reverse=True)
    return ranked[0] if ranked else None


@st.cache_data(show_spinner=False)
def load_installed_base_df(folder: str) -> tuple[pd.DataFrame, str]:
    files = discover_files(folder)
    chosen = pick_best_file(files, score_installed_base_file)
    if chosen is None:
        return pd.DataFrame(), ""

    df = load_any_table(chosen)
    return df, chosen.name


@st.cache_data(show_spinner=False)
def load_spare_parts_df(folder: str) -> tuple[pd.DataFrame, str]:
    files = discover_files(folder)
    chosen = pick_best_file(files, score_spare_parts_file)
    if chosen is None:
        return pd.DataFrame(), ""

    if score_spare_parts_file(chosen) <= 0:
        return pd.DataFrame(), ""

    df = load_any_table(chosen)
    return df, chosen.name


# =========================================================
# DATA PREPARATION
# =========================================================
def harmonize_columns(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    rename_map = {}

    for canonical, aliases in COLUMN_ALIASES.items():
        existing = first_existing_col(out, aliases)
        if existing is not None and existing != canonical:
            rename_map[existing] = canonical

    if rename_map:
        out = out.rename(columns=rename_map)

    return out


def prepare_installed_base_df(df: pd.DataFrame) -> pd.DataFrame:
    out = harmonize_columns(df.copy())

    for col in REQUIRED_BASE_COLS:
        if col not in out.columns:
            out[col] = pd.NA

    text_cols = [
        "Distributor name",
        "Instrument type",
        "Customer name",
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
        "User Software",
        "LXL User Software",
        "LXS User Software",
        "Operating System",
        "PC Model",
        "Win10 Status",
        "EDID Emulator",
        "System in Altitude",
        "Automation partner",
        "LXS Tank configuration",
    ]
    for col in text_cols:
        out[col] = normalize_text(out[col])

    date_cols = ["Installation date", "PM last date", "PM next date", "PM performed On"]
    for col in date_cols:
        out[col] = pd.to_datetime(out[col], errors="coerce")

    num_cols = ["Latitude", "Longitude", "Number of tests per day", "PM frequency"]
    for col in num_cols:
        out[col] = pd.to_numeric(out[col], errors="coerce")

    out["Status Clean"] = out["Operational status"].fillna("Data not available")
    out["Region Clean"] = out["Commercial Region"].fillna("Data not available")
    out["Country Clean"] = out["Country"].fillna("Data not available")
    out["Distributor Clean"] = out["Distributor name"].fillna("Data not available")
    out["City Clean"] = out["City"].fillna("Data not available")
    out["Customer Clean"] = out["Customer name"].fillna("Data not available")
    out["Instrument Clean"] = out["Instrument type"].fillna("Data not available")
    out["Installation year"] = out["Installation date"].dt.year

    out["Platform Group"] = np.where(
        out["Instrument Clean"].isin(list(XL_TYPES)),
        "XL",
        np.where(out["Instrument Clean"].isin(list(XS_TYPES)), "XS", "OTHER"),
    )

    out["Is In Routine"] = out["Status Clean"].eq("IN ROUTINE")
    out["Ready to Install"] = out["Status Clean"].isin(READY_TO_INSTALL_STATUSES)

    out["LXL User Software Final"] = out["LXL User Software"]
    mask_missing_lxl = out["LXL User Software Final"].isna() & out["Instrument Clean"].isin(list(XL_TYPES))
    out.loc[mask_missing_lxl, "LXL User Software Final"] = out.loc[mask_missing_lxl, "User Software"]

    out["LXS User Software Final"] = out["LXS User Software"]
    mask_missing_lxs = out["LXS User Software Final"].isna() & out["Instrument Clean"].isin(list(XS_TYPES))
    out.loc[mask_missing_lxs, "LXS User Software Final"] = out.loc[mask_missing_lxs, "User Software"]

    out["SW Version Status"] = np.where(
        out["User Software"].astype("string").str.contains("4.2.5|4.2.6|1.5", na=False),
        "Up to date",
        np.where(out["User Software"].isna(), "Data not available", "Outdated"),
    )

    out["OS Clean"] = out["Operating System"].fillna("Data not available")
    out["PC Model Clean"] = out["PC Model"].fillna("Data not available")
    out["Win10 Status Clean"] = out["Win10 Status"].fillna("Data not available")
    out["EDID Emulator Clean"] = out["EDID Emulator"].fillna("Data not available")
    out["System in Altitude Clean"] = out["System in Altitude"].fillna("Data not available")
    out["Automation partner Clean"] = out["Automation partner"].fillna("Data not available")
    out["LXS Tank configuration Clean"] = out["LXS Tank configuration"].fillna("Data not available")

    return out


def prepare_spare_parts_df(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df.copy()

    out = harmonize_columns(df.copy())

    needed = [
        "Commercial Region",
        "Distributor name",
        "Type",
        "PN Revision",
        "Carstock",
        "Delta QTY Filter",
        "Instrument Type",
        "PO Frequency (weeks)",
        "Part Number",
        "PART NUMBER DESCRIPTION",
        "Parts per system (12 months)",
        "Minimum Stock Level Required",
        "On Hands",
        "QTY Delta",
        "SP Price (Option 1)",
        "SP Price (Option 2)",
        "SP Price (Option 3)",
    ]
    for col in needed:
        if col not in out.columns:
            out[col] = pd.NA

    text_cols = [
        "Commercial Region",
        "Distributor name",
        "Type",
        "PN Revision",
        "Carstock",
        "Delta QTY Filter",
        "Instrument Type",
        "Part Number",
        "PART NUMBER DESCRIPTION",
    ]
    for col in text_cols:
        out[col] = normalize_text(out[col])

    num_cols = [
        "Parts per system (12 months)",
        "Minimum Stock Level Required",
        "On Hands",
        "QTY Delta",
        "SP Price (Option 1)",
        "SP Price (Option 2)",
        "SP Price (Option 3)",
        "PO Frequency (weeks)",
    ]
    for col in num_cols:
        out[col] = pd.to_numeric(out[col], errors="coerce")

    for opt in ["1", "2", "3"]:
        price_col = f"SP Price (Option {opt})"
        total_col = f"Total Carstock Option {opt}"
        if price_col in out.columns:
            out[total_col] = out[price_col].fillna(0) * out["QTY Delta"].fillna(0)

    return out


# =========================================================
# FILTERS
# =========================================================
def build_global_filters(df: pd.DataFrame, key_prefix: str = "main") -> pd.DataFrame:
    c1, c2, c3, c4 = st.columns([1.2, 1.2, 1.2, 1.0])

    with c1:
        region_options = safe_unique(df, "Region Clean")
        selected_region = st.multiselect(
            "Commercial Region",
            options=region_options,
            default=region_options,
            key=f"{key_prefix}_region",
        )
    df1 = apply_multiselect_filter(df, "Region Clean", selected_region)

    with c2:
        country_options = safe_unique(df1, "Country Clean")
        selected_country = st.multiselect(
            "Country",
            options=country_options,
            default=country_options,
            key=f"{key_prefix}_country",
        )
    df2 = apply_multiselect_filter(df1, "Country Clean", selected_country)

    with c3:
        distributor_options = safe_unique(df2, "Distributor Clean")
        selected_distributor = st.multiselect(
            "Distributor name",
            options=distributor_options,
            default=distributor_options,
            key=f"{key_prefix}_distributor",
        )
    df3 = apply_multiselect_filter(df2, "Distributor Clean", selected_distributor)

    with c4:
        city_options = ["Todas"] + safe_unique(df3, "City Clean")
        city = st.selectbox("City", city_options, index=0, key=f"{key_prefix}_city")

    if city != "Todas":
        df3 = df3[df3["City Clean"] == city].copy()

    return df3


def build_customer_filter(df: pd.DataFrame, key_prefix: str = "customer") -> pd.DataFrame:
    customers = ["Todas"] + safe_unique(df, "Customer Clean")
    chosen = st.selectbox("Customer Name", customers, index=0, key=f"{key_prefix}_customer")
    if chosen == "Todas":
        return df.copy()
    return df[df["Customer Clean"] == chosen].copy()


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
    fig.update_layout(title="Instrument por Type", showlegend=False)
    return glow_layout(fig, height=380)


def fig_instrument_by_status(df: pd.DataFrame) -> go.Figure:
    counts = (
        df["Status Clean"]
        .value_counts(dropna=False)
        .rename_axis("Status")
        .reset_index(name="Count")
        .sort_values("Count", ascending=True)
    )
    fig = px.bar(counts, x="Count", y="Status", orientation="h", text="Count")
    fig.update_traces(marker_color=ACCENT, textposition="outside")
    fig.update_layout(title="Instrument por Status", showlegend=False)
    return glow_layout(fig, height=380)


def fig_installations_per_year(df: pd.DataFrame) -> go.Figure:
    tmp = df.dropna(subset=["Installation year"]).copy()
    if tmp.empty:
        fig = go.Figure()
        fig.update_layout(title="Installations por Year")
        return glow_layout(fig, height=320)

    counts = (
        tmp["Installation year"]
        .astype(int)
        .value_counts()
        .sort_index()
        .rename_axis("Year")
        .reset_index(name="Installations")
    )
    fig = px.bar(counts, x="Year", y="Installations", text="Installations")
    fig.update_traces(marker_color=ACCENT)
    fig.update_layout(title="Installations por Year", showlegend=False)
    return glow_layout(fig, height=320)


def fig_ib_map(df: pd.DataFrame) -> go.Figure:
    geo = df.dropna(subset=["Latitude", "Longitude"]).copy()
    if geo.empty:
        fig = go.Figure()
        fig.update_layout(title="Installed Base Map")
        return glow_layout(fig, height=380)

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
        title="Installed Base Map",
        mapbox_style="carto-positron",
        paper_bgcolor=PLOT_BG,
        plot_bgcolor=PLOT_BG,
        font=dict(color=TEXT),
        margin=dict(l=20, r=20, t=55, b=10),
        height=380,
    )
    return fig


def fig_donut(df: pd.DataFrame, col: str, title: str, top_n: int = 20) -> go.Figure:
    counts = (
        df[col]
        .fillna("Data not available")
        .astype(str)
        .value_counts()
        .head(top_n)
        .rename_axis("Category")
        .reset_index(name="Count")
    )
    fig = px.pie(counts, names="Category", values="Count", hole=0.58)
    fig.update_traces(textposition="outside", textinfo="percent+value")
    fig.update_layout(title=title)
    return glow_layout(fig, height=420)


def fig_simple_donut(df: pd.DataFrame, col: str, title: str) -> go.Figure:
    counts = (
        df[col]
        .fillna("Data not available")
        .astype(str)
        .value_counts()
        .rename_axis("Category")
        .reset_index(name="Count")
    )
    fig = px.pie(counts, names="Category", values="Count", hole=0.55)
    fig.update_traces(textposition="outside", textinfo="percent+value")
    fig.update_layout(title=title)
    return glow_layout(fig, height=340)


# =========================================================
# PAGES
# =========================================================
def show_header(data_file: str, sp_file: str) -> None:
    st.markdown('<div class="main-title">Diasorin Installed Base Dashboard</div>', unsafe_allow_html=True)
    st.markdown(
        f'<div class="subtle">Base file: <b>{data_file or "No detectado"}</b> &nbsp;&nbsp;|&nbsp;&nbsp; Spare parts file: <b>{sp_file or "No detectado"}</b></div>',
        unsafe_allow_html=True,
    )
    st.markdown("<hr>", unsafe_allow_html=True)


def render_installed_base_page(df: pd.DataFrame) -> None:
    filtered = build_global_filters(df, "ib")
    total_ib = len(filtered)
    in_routine = int(filtered["Is In Routine"].sum())
    ready = int(filtered["Ready to Install"].sum())
    xl_total = int(filtered["Instrument Clean"].isin(list(XL_TYPES)).sum())
    xs_total = int(filtered["Instrument Clean"].isin(list(XS_TYPES)).sum())

    c1, c2, c3, c4, c5 = st.columns(5)
    with c1:
        metric_card("TOTAL IB", f"{total_ib:,}", "Installed instruments")
    with c2:
        metric_card("IN ROUTINE", f"{in_routine:,}", f"{safe_ratio(in_routine, total_ib) * 100:.1f}%")
    with c3:
        metric_card("READY TO INSTALL", f"{ready:,}", f"{safe_ratio(ready, total_ib) * 100:.1f}%")
    with c4:
        metric_card("TOTAL XL", f"{xl_total:,}", "XL + XL LAS")
    with c5:
        metric_card("TOTAL XS", f"{xs_total:,}", "LIAISON XS")

    st.markdown("<div class='section-title'>Installed Base Overview</div>", unsafe_allow_html=True)

    row1_c1, row1_c2, row1_c3 = st.columns([1.1, 1.0, 1.0])
    with row1_c1:
        st.plotly_chart(fig_instrument_by_type(filtered), key="ib_type", width="stretch")
    with row1_c2:
        st.plotly_chart(fig_ib_map(filtered), key="ib_map", width="stretch")
    with row1_c3:
        st.plotly_chart(fig_instrument_by_status(filtered), key="ib_status", width="stretch")

    row2_c1, row2_c2, row2_c3 = st.columns([1.0, 1.1, 1.0])
    with row2_c1:
        st.plotly_chart(gauge_figure(in_routine, total_ib, "In Routine vs Total", ACCENT), key="ib_gauge_1", width="stretch")
    with row2_c2:
        st.plotly_chart(fig_installations_per_year(filtered), key="ib_years", width="stretch")
    with row2_c3:
        st.plotly_chart(gauge_figure(ready, total_ib, "Systems ready to be installed", ACCENT_3), key="ib_gauge_2", width="stretch")

    st.markdown("<div class='section-title'>Installed Base Detail</div>", unsafe_allow_html=True)

    detail_cols = [
        "Distributor name", "Instrument type", "City", "Customer name",
        "Serial number", "Installation date", "Operational status",
        "Commercial Region", "Country", "Machine Configurations",
        "User Software", "Operating System", "PC Model",
    ]
    detail_cols = [c for c in detail_cols if c in filtered.columns]

    export_bytes = dataframe_to_excel_bytes({"Installed Base": filtered[detail_cols]})
    st.download_button(
        "⬇️ Export Installed Base (Excel)",
        data=export_bytes,
        file_name=f"Installed_Base_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    st.dataframe(filtered[detail_cols], width="stretch", height=420)


def render_ib_country_distributor_page(df: pd.DataFrame) -> None:
    filtered = build_global_filters(df, "ib_cd")

    c1, c2 = st.columns(2)
    with c1:
        st.plotly_chart(fig_donut(filtered, "Country Clean", "IB by Country", top_n=30), key="country_donut", width="stretch")
    with c2:
        st.plotly_chart(fig_donut(filtered, "Distributor Clean", "IB by Distributor", top_n=30), key="dist_donut", width="stretch")

    detail_cols = [
        "Distributor name", "Country", "City", "Customer name",
        "Instrument type", "Serial number", "Operational status"
    ]
    detail_cols = [c for c in detail_cols if c in filtered.columns]
    st.dataframe(filtered[detail_cols], width="stretch", height=420)


def render_clia_age_page(df: pd.DataFrame) -> None:
    filtered = build_global_filters(df, "clia_age")
    filtered = build_customer_filter(filtered, "clia_age")

    sw_options = safe_unique(filtered, "SW Version Status")
    selected_sw = st.multiselect("SW Version", options=sw_options, default=sw_options, key="clia_age_sw")
    if selected_sw:
        filtered = filtered[filtered["SW Version Status"].isin(selected_sw)].copy()

    c1, c2, c3 = st.columns(3)
    with c1:
        st.plotly_chart(fig_simple_donut(filtered, "EDID Emulator Clean", "EDID Emulator Status"), key="clia_edid", width="stretch")
    with c2:
        st.plotly_chart(fig_instrument_by_status(filtered), key="clia_status", width="stretch")
    with c3:
        st.plotly_chart(fig_simple_donut(filtered, "System in Altitude Clean", "System in Altitude"), key="clia_alt", width="stretch")

    d1, d2 = st.columns(2)
    with d1:
        st.plotly_chart(fig_simple_donut(filtered, "LXL User Software Final", "User Software"), key="clia_lxl_sw", width="stretch")
    with d2:
        st.plotly_chart(fig_simple_donut(filtered, "SW Version Status", "SW Version"), key="clia_sw_status", width="stretch")

    detail_cols = [
        "Distributor name", "Instrument type", "City", "Customer name",
        "Serial number", "Installation date", "Operational status",
        "EDID Emulator", "System in Altitude", "LXL User Software Final",
        "User Software", "SW Version Status"
    ]
    detail_cols = [c for c in detail_cols if c in filtered.columns]
    st.dataframe(filtered[detail_cols], width="stretch", height=420)


def render_lxl_machine_config_page(df: pd.DataFrame) -> None:
    filtered = build_global_filters(df, "lxl_mc")
    filtered = build_customer_filter(filtered, "lxl_mc")
    filtered = filtered[filtered["Instrument Clean"].isin(list(XL_TYPES))].copy()

    sw_options = safe_unique(filtered, "SW Version Status")
    selected_sw = st.multiselect("SW Version", options=sw_options, default=sw_options, key="lxl_mc_sw")
    if selected_sw:
        filtered = filtered[filtered["SW Version Status"].isin(selected_sw)].copy()

    c1, c2, c3 = st.columns(3)
    with c1:
        st.plotly_chart(fig_simple_donut(filtered, "EDID Emulator Clean", "EDID Emulator Status"), key="lxl_edid", width="stretch")
    with c2:
        st.plotly_chart(fig_instrument_by_status(filtered), key="lxl_status", width="stretch")
    with c3:
        st.plotly_chart(fig_simple_donut(filtered, "System in Altitude Clean", "System in Altitude"), key="lxl_altitude", width="stretch")

    d1, d2, d3 = st.columns(3)
    with d1:
        st.plotly_chart(fig_simple_donut(filtered, "LXL User Software Final", "User Software"), key="lxl_user_sw", width="stretch")
    with d2:
        st.plotly_chart(fig_simple_donut(filtered, "SW Version Status", "SW Version"), key="lxl_sw_ver", width="stretch")
    with d3:
        st.plotly_chart(fig_simple_donut(filtered, "Automation partner Clean", "Automation partner"), key="lxl_auto", width="stretch")

    detail_cols = [
        "Distributor name", "Instrument type", "City", "Customer name",
        "Serial number", "Installation date", "Operational status",
        "EDID Emulator", "System in Altitude", "Automation partner",
        "LXL User Software Final", "Machine Configurations"
    ]
    detail_cols = [c for c in detail_cols if c in filtered.columns]
    st.dataframe(filtered[detail_cols], width="stretch", height=420)


def render_win10_page(df: pd.DataFrame) -> None:
    filtered = build_global_filters(df, "win10")
    filtered = build_customer_filter(filtered, "win10")

    status_options = safe_unique(filtered, "Status Clean")
    selected_status = st.multiselect("Status", options=status_options, default=status_options, key="win10_status")
    if selected_status:
        filtered = filtered[filtered["Status Clean"].isin(selected_status)].copy()

    c1, c2, c3 = st.columns(3)
    with c1:
        st.plotly_chart(fig_simple_donut(filtered, "OS Clean", "Operative System (ISR Live)"), key="win_os", width="stretch")
    with c2:
        st.plotly_chart(fig_simple_donut(filtered, "PC Model Clean", "PC Model (ISR Live)"), key="win_pc", width="stretch")
    with c3:
        compare_df = filtered.copy()
        compare_df["ISR Live vs SN Analysis"] = np.where(
            compare_df["OS Clean"].astype(str).str.contains("win10", case=False, na=False),
            "Win10 Original",
            "Inconsistency / review"
        )
        st.plotly_chart(fig_simple_donut(compare_df, "ISR Live vs SN Analysis", "ISR Live vs SN Analysis"), key="win_analysis", width="stretch")

    d1, d2 = st.columns(2)
    with d1:
        st.plotly_chart(fig_simple_donut(filtered, "Win10 Status Clean", "Win10 Status (Original)"), key="win_status_orig", width="stretch")
    with d2:
        fully = int(filtered["Win10 Status Clean"].astype(str).str.contains("fully", case=False, na=False).sum())
        total = len(filtered)
        st.plotly_chart(gauge_figure(fully, total, "Upgrade Progress", ACCENT_3), key="win_upgrade", width="stretch")

    detail_cols = [
        "Distributor name", "City", "Customer name", "Serial number",
        "Operational status", "PC Model", "Operating System", "Win10 Status"
    ]
    detail_cols = [c for c in detail_cols if c in filtered.columns]
    st.dataframe(filtered[detail_cols], width="stretch", height=420)


def render_lxs_machine_config_page(df: pd.DataFrame) -> None:
    filtered = build_global_filters(df, "lxs_mc")
    filtered = build_customer_filter(filtered, "lxs_mc")
    filtered = filtered[filtered["Instrument Clean"].isin(list(XS_TYPES))].copy()

    c1, c2, c3 = st.columns(3)
    with c1:
        st.plotly_chart(fig_simple_donut(filtered, "Region Clean", "Commercial Region"), key="lxs_region", width="stretch")
    with c2:
        st.plotly_chart(fig_instrument_by_status(filtered), key="lxs_status", width="stretch")
    with c3:
        st.plotly_chart(fig_simple_donut(filtered, "System in Altitude Clean", "System in Altitude"), key="lxs_alt", width="stretch")

    d1, d2 = st.columns(2)
    with d1:
        st.plotly_chart(fig_simple_donut(filtered, "LXS User Software Final", "User Software"), key="lxs_sw", width="stretch")
    with d2:
        st.plotly_chart(fig_simple_donut(filtered, "LXS Tank configuration Clean", "LXS Tank configuration"), key="lxs_tank", width="stretch")

    detail_cols = [
        "Distributor name", "Instrument type", "City", "Customer name",
        "Serial number", "Installation date", "Operational status",
        "LXS User Software Final", "LXS Tank configuration", "System in Altitude"
    ]
    detail_cols = [c for c in detail_cols if c in filtered.columns]
    st.dataframe(filtered[detail_cols], width="stretch", height=420)


def render_sp_price_list_page(sp_df: pd.DataFrame) -> None:
    if sp_df.empty:
        st.warning("No encontré un archivo de spare parts / price list en esta carpeta.")
        return

    df = sp_df.copy()

    c1, c2, c3, c4 = st.columns([1.0, 1.2, 1.2, 1.0])

    with c1:
        reg = safe_unique(df, "Commercial Region")
        reg_sel = st.multiselect("Commercial Region", reg, default=reg, key="sp_reg")
    if reg_sel:
        df = df[df["Commercial Region"].isin(reg_sel)].copy()

    with c2:
        dist = safe_unique(df, "Distributor name")
        dist_sel = st.multiselect("Distributor name", dist, default=dist, key="sp_dist")
    if dist_sel:
        df = df[df["Distributor name"].isin(dist_sel)].copy()

    with c3:
        tps = safe_unique(df, "Type")
        tps_sel = st.multiselect("Type", tps, default=tps, key="sp_type")
    if tps_sel:
        df = df[df["Type"].isin(tps_sel)].copy()

    with c4:
        pnrev = ["Todas"] + safe_unique(df, "PN Revision")
        pnrev_sel = st.selectbox("PN Revision", pnrev, index=0, key="sp_pnrev")
    if pnrev_sel != "Todas":
        df = df[df["PN Revision"] == pnrev_sel].copy()

    c5, c6, c7, c8 = st.columns(4)
    with c5:
        carstock_vals = ["Todas"] + safe_unique(df, "Carstock")
        carstock_sel = st.selectbox("Carstock", carstock_vals, index=0, key="sp_carstock")
    if carstock_sel != "Todas":
        df = df[df["Carstock"] == carstock_sel].copy()

    with c6:
        delta_vals = ["Todas"] + safe_unique(df, "Delta QTY Filter")
        delta_sel = st.selectbox("Delta QTY Filter", delta_vals, index=0, key="sp_delta")
    if delta_sel != "Todas":
        df = df[df["Delta QTY Filter"] == delta_sel].copy()

    with c7:
        inst_vals = ["Todas"] + safe_unique(df, "Instrument Type")
        inst_sel = st.selectbox("Instrument Type", inst_vals, index=0, key="sp_inst")
    if inst_sel != "Todas":
        df = df[df["Instrument Type"] == inst_sel].copy()

    with c8:
        po_vals = ["Todas"] + [str(v) for v in sorted(df["PO Frequency (weeks)"].dropna().unique().tolist())]
        po_sel = st.selectbox("PO Frequency (weeks)", po_vals, index=0, key="sp_po")
    if po_sel != "Todas":
        df = df[df["PO Frequency (weeks)"] == pd.to_numeric(po_sel, errors="coerce")].copy()

    search = st.text_input("Search part number or description", value="", key="sp_search").strip().lower()
    if search:
        df = df[
            df["Part Number"].astype("string").str.lower().str.contains(search, na=False) |
            df["PART NUMBER DESCRIPTION"].astype("string").str.lower().str.contains(search, na=False)
        ].copy()

    total_xl = int(df["Instrument Type"].astype("string").str.contains("xl", case=False, na=False).sum()) if "Instrument Type" in df.columns else 0
    total_emx = int(df["Instrument Type"].astype("string").str.contains("eti|max|emx", case=False, na=False).sum()) if "Instrument Type" in df.columns else 0
    total_xs = int(df["Instrument Type"].astype("string").str.contains("xs", case=False, na=False).sum()) if "Instrument Type" in df.columns else 0

    k1, k2, k3 = st.columns(3)
    with k1:
        metric_card("TOTAL XL", f"{total_xl:,}", "Rows in current filter")
    with k2:
        metric_card("TOTAL EMX", f"{total_emx:,}", "Rows in current filter")
    with k3:
        metric_card("TOTAL XS", f"{total_xs:,}", "Rows in current filter")

    display_cols = [
        "Instrument Type",
        "Part Number",
        "PART NUMBER DESCRIPTION",
        "Parts per system (12 months)",
        "Minimum Stock Level Required",
        "On Hands",
        "QTY Delta",
        "SP Price (Option 1)",
        "Total Carstock Option 1",
        "SP Price (Option 2)",
        "Total Carstock Option 2",
        "SP Price (Option 3)",
        "Total Carstock Option 3",
        "Distributor name",
        "Commercial Region",
        "Type",
        "PN Revision",
        "Carstock",
    ]
    display_cols = [c for c in display_cols if c in df.columns]

    export_bytes = dataframe_to_excel_bytes({"SP Price List": df[display_cols]})
    st.download_button(
        "⬇️ Export SP Price List (Excel)",
        data=export_bytes,
        file_name=f"SP_Price_List_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    st.dataframe(df[display_cols], width="stretch", height=500)


def render_parameters_clia_page(df: pd.DataFrame) -> None:
    filtered = build_global_filters(df, "params")
    filtered = build_customer_filter(filtered, "params")

    clia_cols = [c for c in filtered.columns if c.upper().startswith("CLIA -")]
    if not clia_cols:
        st.warning("No detecté columnas CLIA - en la base actual.")
        st.dataframe(filtered.head(100), width="stretch", height=420)
        return

    melted = filtered[["Serial number", "Customer name", "Distributor name", "Instrument type"] + clia_cols].copy()
    long_df = melted.melt(
        id_vars=["Serial number", "Customer name", "Distributor name", "Instrument type"],
        value_vars=clia_cols,
        var_name="CLIA Parameter",
        value_name="Value"
    )
    long_df["Value"] = long_df["Value"].astype("string").fillna("Data not available")

    summary = (
        long_df[long_df["Value"].str.lower().isin(["yes", "x", "true", "1"], na=False)]
        .groupby("CLIA Parameter")
        .size()
        .reset_index(name="Count")
        .sort_values("Count", ascending=True)
    )

    if summary.empty:
        st.info("Las columnas CLIA existen, pero no encontré valores afirmativos estándar.")
        st.dataframe(long_df.head(300), width="stretch", height=420)
        return

    fig = px.bar(summary, x="Count", y="CLIA Parameter", orientation="h", text="Count")
    fig.update_traces(marker_color=ACCENT)
    fig.update_layout(title="Parameters CLIA")
    st.plotly_chart(glow_layout(fig, height=720), key="clia_params", width="stretch")
    st.dataframe(long_df, width="stretch", height=420)


# =========================================================
# SIDEBAR
# =========================================================
def sidebar_controls(default_folder: str) -> tuple[str, str]:
    st.sidebar.markdown("## Páginas")
    page = st.sidebar.radio("Ir a", ALL_PAGES, index=0, label_visibility="collapsed")

    st.sidebar.markdown("---")
    folder = st.sidebar.text_input(
        "Carpeta de trabajo",
        value=default_folder,
        help="La app buscará Excel/CSV en esta carpeta para base instalada y spare parts.",
    )

    st.sidebar.markdown("---")
    st.sidebar.caption("Ambiente de pruebas Streamlit")
    return page, folder


# =========================================================
# MAIN
# =========================================================
def main():
    default_folder = r"C:\Users\javier.avellaneda\OneDrive - Diasorin-Luminex\Documentos\Service DashBoard Backup\Python"
    page, folder = sidebar_controls(default_folder)

    try:
        raw_df, data_file = load_installed_base_df(folder)
    except Exception as e:
        st.error(f"Error cargando base instalada: {e}")
        st.stop()

    try:
        sp_raw_df, sp_file = load_spare_parts_df(folder)
    except Exception as e:
        sp_raw_df, sp_file = pd.DataFrame(), ""
        st.warning(f"No pude cargar spare parts: {e}")

    show_header(data_file, sp_file)

    if raw_df.empty:
        st.error("No encontré una base instalada válida en la carpeta indicada.")
        st.info("Pon en esta carpeta el archivo Records / Installed Base / ISR Live en formato Excel o CSV.")
        st.stop()

    ib_df = prepare_installed_base_df(raw_df)
    sp_df = prepare_spare_parts_df(sp_raw_df)

    if page == PAGE_INSTALLED_BASE:
        render_installed_base_page(ib_df)
    elif page == PAGE_IB_COUNTRY_DISTRIBUTOR:
        render_ib_country_distributor_page(ib_df)
    elif page == PAGE_CLIA_AGE:
        render_clia_age_page(ib_df)
    elif page == PAGE_LXL_MACHINE_CONFIG:
        render_lxl_machine_config_page(ib_df)
    elif page == PAGE_WIN10:
        render_win10_page(ib_df)
    elif page == PAGE_LXS_MACHINE_CONFIG:
        render_lxs_machine_config_page(ib_df)
    elif page == PAGE_SP_PRICE_LIST:
        render_sp_price_list_page(sp_df)
    elif page == PAGE_PARAMETERS_CLIA:
        render_parameters_clia_page(ib_df)


if __name__ == "__main__":
    main()
