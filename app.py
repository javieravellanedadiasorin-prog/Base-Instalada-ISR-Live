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




# =============================================================================
# MÓDULO ADICIONAL: ANTIGÜEDAD SEGÚN FECHA DE FABRICACIÓN
# Este bloque es independiente y no altera la lógica existente del dashboard.
# =============================================================================

def normalize_serial_match(value) -> str:
    """Normaliza seriales para cruzar Records List con archivos de fabricación."""
    if pd.isna(value):
        return ""
    text_value = str(value).strip()
    if text_value.startswith('="') and text_value.endswith('"'):
        text_value = text_value[2:-1]
    text_value = re.sub(r"\.0$", "", text_value)
    return re.sub(r"[^A-Za-z0-9]+", "", text_value).upper()


def normalize_column_label(value) -> str:
    return re.sub(r"[^a-z0-9]+", " ", str(value).lower()).strip()


def guess_column_index(columns, keywords: list[str]) -> int:
    normalized = [normalize_column_label(col) for col in columns]
    for keyword in keywords:
        key = normalize_column_label(keyword)
        for idx, column_name in enumerate(normalized):
            if column_name == key:
                return idx
    for keyword in keywords:
        key = normalize_column_label(keyword)
        for idx, column_name in enumerate(normalized):
            if key and key in column_name:
                return idx
    return 0


def manufacturing_date_candidate_columns(df: pd.DataFrame) -> list[str]:
    candidates = []
    preferred_terms = [
        "manufacturing date",
        "manufacture date",
        "production date",
        "mfg date",
        "build date",
        "fecha de fabricacion",
        "fecha fabricación",
        "po date",
        "invoice date",
    ]
    for term in preferred_terms:
        term_norm = normalize_column_label(term)
        for col in df.columns:
            col_norm = normalize_column_label(col)
            if (col_norm == term_norm or term_norm in col_norm) and col not in candidates:
                candidates.append(col)
    for col in df.columns:
        if "date" in normalize_column_label(col) or "fecha" in normalize_column_label(col):
            if col not in candidates:
                candidates.append(col)
    return candidates or df.columns.tolist()


def parse_date_flexible(series: pd.Series) -> pd.Series:
    direct = pd.to_datetime(series, errors="coerce", dayfirst=False)
    alternate = pd.to_datetime(series, errors="coerce", dayfirst=True)
    return direct.fillna(alternate)


@st.cache_data(show_spinner=False)
def load_manufacturing_source(file_bytes: bytes, filename: str) -> dict[str, pd.DataFrame]:
    """Lee uno o varios archivos de fabricación subidos manualmente en cada sesión."""
    name = filename.lower()
    try:
        if name.endswith(".csv"):
            raw_text = file_bytes.decode("utf-8-sig", errors="replace")
            for sep in [";", ",", None]:
                try:
                    frame = pd.read_csv(StringIO(raw_text), sep=sep, engine="python")
                    if frame.shape[1] >= 2:
                        return {filename: frame}
                except Exception:
                    continue
            raise ValueError("No se pudo reconocer la estructura del CSV.")
        if name.endswith(".xlsx") or name.endswith(".xls"):
            workbook = pd.read_excel(BytesIO(file_bytes), sheet_name=None)
            return {str(sheet): frame for sheet, frame in workbook.items() if frame is not None and not frame.empty}
    except ImportError as exc:
        if name.endswith(".xls"):
            raise RuntimeError(
                "Para leer archivos .XLS en Streamlit se requiere incluir `xlrd>=2.0.1` en requirements.txt."
            ) from exc
        raise
    raise ValueError(f"Formato no soportado para fecha de fabricación: {filename}")


def build_age_bucket(series: pd.Series) -> pd.Series:
    return pd.cut(
        series,
        bins=[-np.inf, 3, 5, 8, 10, 15, np.inf],
        labels=["0–3 años", "3–5 años", "5–8 años", "8–10 años", "10–15 años", "15+ años"],
        right=False,
    )


def build_manufacturing_match(
    filtered_assets: pd.DataFrame,
    manufacturing_frames: list[pd.DataFrame],
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Cruza la vista filtrada actual con las referencias de fabricación por serial."""
    assets = filtered_assets.copy()
    assets["Serial match key"] = assets["Serial number"].map(normalize_serial_match)

    if not manufacturing_frames:
        assets["Manufacturing Date"] = pd.NaT
        assets["Manufacturing Source"] = pd.NA
        assets["Manufacturing Product"] = pd.NA
        assets["Manufacturing age (years)"] = np.nan
        assets["Manufacturing age bucket"] = pd.NA
        return assets, pd.DataFrame()

    reference = pd.concat(manufacturing_frames, ignore_index=True)
    reference["Serial match key"] = reference["Manufacturing Serial"].map(normalize_serial_match)
    reference["Manufacturing Date"] = pd.to_datetime(reference["Manufacturing Date"], errors="coerce")
    reference = reference[(reference["Serial match key"] != "") & reference["Manufacturing Date"].notna()].copy()

    duplicate_check = (
        reference.groupby("Serial match key", dropna=False)["Manufacturing Date"]
        .nunique()
        .reset_index(name="Different manufacturing dates")
    )
    conflicting_keys = set(
        duplicate_check.loc[duplicate_check["Different manufacturing dates"] > 1, "Serial match key"].tolist()
    )

    reference = reference.sort_values(["Serial match key", "Manufacturing Date"], ascending=[True, True])
    reference = reference.drop_duplicates(subset=["Serial match key"], keep="first").copy()
    reference["Manufacturing date conflict"] = reference["Serial match key"].isin(conflicting_keys)

    keep_cols = [
        "Serial match key",
        "Manufacturing Serial",
        "Manufacturing Date",
        "Manufacturing Source",
        "Manufacturing Sheet",
        "Manufacturing Product",
        "Manufacturing date conflict",
    ]
    merged = assets.merge(reference[keep_cols], on="Serial match key", how="left")
    today = pd.Timestamp(date.today())
    merged["Manufacturing age (years)"] = ((today - merged["Manufacturing Date"]).dt.days / 365.25).round(1)
    merged["Manufacturing year"] = merged["Manufacturing Date"].dt.year.astype("Int64")
    merged["Manufacturing age bucket"] = build_age_bucket(merged["Manufacturing age (years)"])
    merged["Manufacturing matched"] = merged["Manufacturing Date"].notna()
    return merged, reference


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
@@ -3349,88 +3506,88 @@
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
base_tab, machine_tab, os_tab, process_tab, stock_tab, manufacturing_tab, detail_tab = st.tabs(
    ["Base instalada", "Machine configuration", "Sistema operativo", "Procesamiento / PM", "Stock / Carstock gap", "Antigüedad / fabricación", "Detalle por equipo"]
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
@@ -4429,86 +4586,417 @@
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


with manufacturing_tab:
    st.subheader("Antigüedad de la base instalada por fecha de fabricación")
    st.caption(
        "Esta pestaña cruza el serial de la vista filtrada actual con los archivos de fabricación que cargues manualmente. "
        "Los filtros globales de región, país, distribuidor, tipo de instrumento y estado operativo se aplican automáticamente."
    )

    manufacturing_uploads = st.file_uploader(
        "Carga manualmente los archivos actualizados de fabricación",
        type=["xls", "xlsx", "csv"],
        accept_multiple_files=True,
        key="manufacturing_date_uploads",
        help="Puedes cargar simultáneamente los archivos LXL/LXL LAS, LXS y ETI-MAX. No quedan almacenados al cerrar la sesión.",
    )

    if not manufacturing_uploads:
        st.info(
            "Carga los archivos de fechas de fabricación para activar el cruce por serial y visualizar la edad real de los equipos."
        )
    else:
        prepared_sources = []
        reading_errors = []

        st.markdown("### Configuración de las fuentes cargadas")
        st.markdown(
            '<div class="small-note">Selecciona la hoja, el serial y la columna que corporativamente corresponde a la fecha de fabricación. '
            'Esta confirmación evita interpretar automáticamente <b>PO Date</b> o <b>Invoice Date</b> como fabricación cuando no corresponda.</div>',
            unsafe_allow_html=True,
        )

        for source_idx, uploaded_manufacturing in enumerate(manufacturing_uploads):
            try:
                source_sheets = load_manufacturing_source(uploaded_manufacturing.getvalue(), uploaded_manufacturing.name)
                sheet_names = list(source_sheets.keys())
                if not sheet_names:
                    reading_errors.append(f"{uploaded_manufacturing.name}: no contiene hojas con datos.")
                    continue

                sheet_scores = []
                for sheet_name in sheet_names:
                    frame = source_sheets[sheet_name]
                    score = 0
                    normalized_cols = [normalize_column_label(c) for c in frame.columns]
                    if any("serial" in c for c in normalized_cols):
                        score += 10
                    if any("date" in c or "fecha" in c for c in normalized_cols):
                        score += 5
                    score += min(len(frame), 1000) / 10000
                    sheet_scores.append((score, sheet_name))
                suggested_sheet = sorted(sheet_scores, reverse=True)[0][1]

                with st.expander(f"Fuente: {uploaded_manufacturing.name}", expanded=True):
                    selected_sheet = st.selectbox(
                        "Hoja con información de seriales",
                        options=sheet_names,
                        index=sheet_names.index(suggested_sheet),
                        key=f"manufacturing_sheet_{source_idx}",
                    )
                    source_frame = source_sheets[selected_sheet].copy()
                    columns_available = source_frame.columns.tolist()

                    if source_frame.empty or not columns_available:
                        st.warning("La hoja seleccionada no contiene registros.")
                        continue

                    serial_default = guess_column_index(
                        columns_available,
                        ["Serial No", "Serial number", "Serial", "Número de serie", "Serie"],
                    )
                    date_options = manufacturing_date_candidate_columns(source_frame)
                    date_default = guess_column_index(
                        date_options,
                        ["Manufacturing Date", "Manufacture Date", "Production Date", "MFG Date", "Fecha de fabricación", "PO Date", "Invoice Date"],
                    )
                    product_default = 0
                    product_terms = ["Product", "Product description", "Material", "Instrument", "Description", "Model"]
                    normalized_product_columns = [normalize_column_label(col) for col in columns_available]
                    for term in product_terms:
                        term_normalized = normalize_column_label(term)
                        matched_product_columns = [
                            idx for idx, column_name in enumerate(normalized_product_columns)
                            if column_name == term_normalized or term_normalized in column_name
                        ]
                        if matched_product_columns:
                            product_default = matched_product_columns[0] + 1
                            break

                    c_source_1, c_source_2, c_source_3 = st.columns(3)
                    with c_source_1:
                        serial_column = st.selectbox(
                            "Columna de serial",
                            options=columns_available,
                            index=serial_default,
                            key=f"manufacturing_serial_col_{source_idx}",
                        )
                    with c_source_2:
                        date_column = st.selectbox(
                            "Columna de fecha de fabricación",
                            options=date_options,
                            index=date_default if date_default < len(date_options) else 0,
                            key=f"manufacturing_date_col_{source_idx}",
                        )
                    with c_source_3:
                        product_choices = ["<sin columna de producto>"] + columns_available
                        selected_product = st.selectbox(
                            "Columna de producto / modelo (opcional)",
                            options=product_choices,
                            index=product_default,
                            key=f"manufacturing_product_col_{source_idx}",
                        )

                    prepared = pd.DataFrame(
                        {
                            "Manufacturing Serial": source_frame[serial_column],
                            "Manufacturing Date": parse_date_flexible(source_frame[date_column]),
                            "Manufacturing Source": uploaded_manufacturing.name,
                            "Manufacturing Sheet": selected_sheet,
                            "Manufacturing Product": (
                                source_frame[selected_product]
                                if selected_product != "<sin columna de producto>"
                                else pd.Series(pd.NA, index=source_frame.index)
                            ),
                        }
                    )
                    prepared["Serial match key"] = prepared["Manufacturing Serial"].map(normalize_serial_match)
                    valid_rows = int(((prepared["Serial match key"] != "") & prepared["Manufacturing Date"].notna()).sum())
                    st.caption(
                        f"Registros válidos para cruce: {valid_rows:,} de {len(prepared):,} | "
                        f"Fecha utilizada: {date_column}"
                    )
                    if valid_rows:
                        prepared_sources.append(prepared)
            except Exception as exc:
                reading_errors.append(f"{uploaded_manufacturing.name}: {exc}")

        if reading_errors:
            for message in reading_errors:
                st.warning(message)

        if not prepared_sources:
            st.info("Aún no hay una fuente válida con serial y fecha para realizar la comparación.")
        else:
            manufacturing_df, manufacturing_reference = build_manufacturing_match(filtered, prepared_sources)
            matched_df = manufacturing_df[manufacturing_df["Manufacturing matched"]].copy()
            unmatched_df = manufacturing_df[~manufacturing_df["Manufacturing matched"]].copy()
            future_date_count = int((matched_df["Manufacturing Date"] > pd.Timestamp(date.today())).sum()) if not matched_df.empty else 0
            conflicting_date_count = int(matched_df.get("Manufacturing date conflict", pd.Series(dtype=bool)).fillna(False).sum()) if not matched_df.empty else 0
            match_rate = _safe_share_pct(len(matched_df), len(manufacturing_df))

            st.markdown("### Cobertura y edad del parque filtrado")
            a1, a2, a3, a4, a5 = st.columns(5)
            with a1:
                metric_card("Equipos filtrados", f"{len(manufacturing_df):,}", "Vista actual del dashboard")
            with a2:
                metric_card("Seriales cruzados", f"{len(matched_df):,}", f"{match_rate:.1f}% con fecha de fabricación")
            with a3:
                average_age = matched_df["Manufacturing age (years)"].mean() if not matched_df.empty else pd.NA
                metric_card("Edad promedio", f"{average_age:.1f} años" if pd.notna(average_age) else "N/A", "Según fecha seleccionada")
            with a4:
                oldest_age = matched_df["Manufacturing age (years)"].max() if not matched_df.empty else pd.NA
                metric_card("Equipo más antiguo", f"{oldest_age:.1f} años" if pd.notna(oldest_age) else "N/A", "Mayor edad identificada")
            with a5:
                newest_age = matched_df["Manufacturing age (years)"].min() if not matched_df.empty else pd.NA
                metric_card("Equipo más nuevo", f"{newest_age:.1f} años" if pd.notna(newest_age) else "N/A", "Menor edad identificada")

            if future_date_count > 0:
                st.warning(f"Se detectaron {future_date_count} equipos con fecha de fabricación futura; revisa la columna de fecha seleccionada.")
            if conflicting_date_count > 0:
                st.warning(f"Se detectaron {conflicting_date_count} seriales con más de una fecha distinta en los archivos cargados.")

            if matched_df.empty:
                st.warning("No se encontraron coincidencias por serial entre la base filtrada y las fuentes de fabricación.")
            else:
                matched_df["Age label"] = (
                    matched_df["Serial number"].fillna("SIN SERIAL").astype(str)
                    + " | "
                    + matched_df["Instrument type"].fillna("Sin modelo").astype(str)
                )
                matched_df["Manufacturing date display"] = matched_df["Manufacturing Date"].map(format_date_for_hover)
                oldest = matched_df.sort_values(["Manufacturing age (years)", "Serial number"], ascending=[False, True]).head(15)
                annual = (
                    matched_df.groupby("Manufacturing year", dropna=False)
                    .size()
                    .reset_index(name="Count")
                    .dropna(subset=["Manufacturing year"])
                    .sort_values("Manufacturing year")
                )
                age_distribution = (
                    matched_df["Manufacturing age bucket"]
                    .value_counts(sort=False)
                    .reset_index()
                )
                age_distribution.columns = ["Rango de edad", "Count"]
                age_distribution = age_distribution[age_distribution["Count"] > 0]

                chart_left, chart_right = st.columns(2)
                with chart_left:
                    fig_age_distribution = px.bar(
                        age_distribution,
                        x="Rango de edad",
                        y="Count",
                        text="Count",
                        title="Estado de la base instalada por rango de edad",
                        category_orders={"Rango de edad": ["0–3 años", "3–5 años", "5–8 años", "8–10 años", "10–15 años", "15+ años"]},
                    )
                    fig_age_distribution.update_traces(
                        marker_color=ACCENT,
                        textposition="outside",
                        hovertemplate="Rango: %{x}<br>Equipos: %{y}<extra></extra>",
                    )
                    st.plotly_chart(glow_layout(fig_age_distribution, 475, title_size=16), use_container_width=True)

                with chart_right:
                    fig_oldest = px.bar(
                        oldest.sort_values("Manufacturing age (years)", ascending=True),
                        x="Manufacturing age (years)",
                        y="Age label",
                        orientation="h",
                        text="Manufacturing age (years)",
                        title="Top 15 equipos más antiguos",
                        custom_data=["Manufacturing date display", "Customer name", "Distributor name", "Country", "Operational status"],
                    )
                    fig_oldest.update_traces(
                        marker_color=WARNING,
                        texttemplate="%{text:.1f} años",
                        textposition="outside",
                        hovertemplate=(
                            "Equipo: %{y}<br>"
                            "Edad: %{x:.1f} años<br>"
                            "Fabricación: %{customdata[0]}<br>"
                            "Cliente: %{customdata[1]}<br>"
                            "Distribuidor: %{customdata[2]}<br>"
                            "País: %{customdata[3]}<br>"
                            "Estado: %{customdata[4]}<extra></extra>"
                        ),
                    )
                    st.plotly_chart(glow_layout(fig_oldest, 475, title_size=16), use_container_width=True)

                timeline_left, timeline_right = st.columns(2)
                with timeline_left:
                    fig_timeline = px.scatter(
                        matched_df.sort_values("Manufacturing Date"),
                        x="Manufacturing Date",
                        y="Serial number",
                        color="Manufacturing age bucket",
                        title="Línea de tiempo de fabricación por serial",
                        custom_data=["Instrument type", "Manufacturing age (years)", "Customer name", "Distributor name", "Country", "Operational status", "Manufacturing Source"],
                        color_discrete_map={
                            "0–3 años": ACCENT_3,
                            "3–5 años": ACCENT,
                            "5–8 años": ACCENT_2,
                            "8–10 años": WARNING,
                            "10–15 años": "#ff8b55",
                            "15+ años": DANGER,
                        },
                    )
                    fig_timeline.update_traces(
                        marker=dict(size=10, opacity=0.90),
                        hovertemplate=(
                            "Serial: %{y}<br>"
                            "Fabricación: %{x|%Y-%m-%d}<br>"
                            "Modelo: %{customdata[0]}<br>"
                            "Edad: %{customdata[1]:.1f} años<br>"
                            "Cliente: %{customdata[2]}<br>"
                            "Distribuidor: %{customdata[3]}<br>"
                            "País: %{customdata[4]}<br>"
                            "Estado: %{customdata[5]}<br>"
                            "Fuente: %{customdata[6]}<extra></extra>"
                        ),
                    )
                    fig_timeline.update_layout(legend_title="Rango de edad")
                    st.plotly_chart(glow_layout(fig_timeline, 600, title_size=16), use_container_width=True)

                with timeline_right:
                    fig_annual = px.bar(
                        annual,
                        x="Manufacturing year",
                        y="Count",
                        text="Count",
                        title="Equipos por año de fabricación",
                    )
                    fig_annual.update_traces(
                        marker_color=ACCENT_2,
                        textposition="outside",
                        hovertemplate="Año de fabricación: %{x}<br>Equipos: %{y}<extra></extra>",
                    )
                    st.plotly_chart(glow_layout(fig_annual, 600, title_size=16), use_container_width=True)

                st.markdown("### Equipos ordenados del más antiguo al más nuevo")
                age_table_columns = [
                    "Commercial Region",
                    "Country",
                    "Distributor name",
                    "Customer name",
                    "Instrument type",
                    "Serial number",
                    "Manufacturing Date",
                    "Manufacturing age (years)",
                    "Manufacturing age bucket",
                    "Installation date",
                    "Operational status",
                    "Manufacturing Source",
                ]
                age_table = matched_df[age_table_columns].sort_values(
                    ["Manufacturing age (years)", "Manufacturing Date"],
                    ascending=[False, True],
                )
                st.dataframe(age_table, use_container_width=True, hide_index=True)

                st.download_button(
                    "Descargar cruce de fechas de fabricación",
                    data=to_csv_download(manufacturing_df.drop(columns=["Serial match key"], errors="ignore")),
                    file_name="installed_base_manufacturing_age_filtered.csv",
                    mime="text/csv",
                    use_container_width=False,
                    key="download_manufacturing_age_analysis",
                )

            if not unmatched_df.empty:
                with st.expander(f"Seriales sin coincidencia en los archivos de fabricación ({len(unmatched_df):,})", expanded=False):
                    missing_columns = [
                        "Country",
                        "Distributor name",
                        "Customer name",
                        "Instrument type",
                        "Serial number",
                        "Operational status",
                    ]
                    st.dataframe(unmatched_df[missing_columns], use_container_width=True, hide_index=True)

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
