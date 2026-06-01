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
@@ -3389,8 +3546,8 @@ def build_distributor_model_donut(df: pd.DataFrame, selected_model: str, top_n:
st.stop()

st.sidebar.markdown("---")
base_tab, machine_tab, os_tab, process_tab, stock_tab, detail_tab = st.tabs(
    ["Base instalada", "Machine configuration", "Sistema operativo", "Procesamiento / PM", "Stock / Carstock gap", "Detalle por equipo"]
base_tab, machine_tab, os_tab, process_tab, stock_tab, manufacturing_tab, detail_tab = st.tabs(
    ["Base instalada", "Machine configuration", "Sistema operativo", "Procesamiento / PM", "Stock / Carstock gap", "Antigüedad / fabricación", "Detalle por equipo"]
)

with base_tab:
@@ -4469,6 +4626,337 @@ def build_distributor_model_donut(df: pd.DataFrame, selected_model: str, top_n:
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
