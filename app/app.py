"""Entrada Streamlit. No se exige usuario ni contraseña en la aplicación."""
from modules.core import *  # noqa: F401,F403 - compatibilidad con app original

# =========================
# APP
# =========================

st.title(
    "📌 Reporteador de Incidencias 2026"
)

st.caption(
    "Sistema operativo IMSS-BIENESTAR para seguimiento de incidencias."
)

# Estilos ligeros: no se cargan fuentes, imágenes ni bibliotecas externas.
st.markdown("""
<style>
.block-container {padding-top: 1.25rem; padding-bottom: 2rem; max-width: 1450px;}
[data-testid="stSidebar"] {background: #f7faf9;}
[data-testid="stSidebar"] h1 {font-size: 1.2rem;}
h1 {letter-spacing: -.025em;}
div[data-testid="stMetric"] {border: 1px solid #e3ebe8; border-radius: 10px; padding: .65rem .9rem;}
</style>
""", unsafe_allow_html=True)

st.sidebar.title(
    "Panel de control"
)

st.sidebar.success(
    "Base principal conectada a Supabase."
)
st.sidebar.divider()


# =========================
# MENÚ
# =========================

menu = st.sidebar.radio(
    "Menú",
    [
        "Registrar incidencia",
        "Urgencias",
        "Resumen Ejecutivo",
        "Seguimiento",
        "Cédulas de rechazo",
        "Base Supabase"
    ]
)


# =========================
# CARGA INTELIGENTE SUPABASE
# =========================

# No se modifican filas en Supabase. Se excluyen urgencias históricas
# de los paneles de Incidencias únicamente a nivel de presentación.
# Urgencias trabaja con Google Sheets: no descargar toda la tabla de
# incidencias desde Supabase cuando se abre este módulo.
incidencias = pd.DataFrame() if menu == "Urgencias" else cargar_incidencias()
if menu != "Urgencias" and not incidencias.empty and "TIPO_INCIDENCIA" in incidencias.columns:
    incidencias = incidencias.loc[
        ~incidencias["TIPO_INCIDENCIA"].fillna("").astype(str).str.strip().str.upper().eq("URGENCIA")
    ].copy()

if menu == "Seguimiento":

    with st.spinner(
        "Revisando actualización diaria de seguimiento..."
    ):

        resultado_auto = actualizar_seguimiento_diario_si_corresponde(
            limite=10000,
            forzar=False
        )

    if resultado_auto.get(
        "ejecutado",
        False
    ):

        st.success(
            f"Seguimiento actualizado automáticamente: {resultado_auto.get('actualizadas', 0)} registros."
        )

        if resultado_auto.get(
            "errores",
            []
        ):

            with st.expander(
                "Ver errores de actualización"
            ):

                st.write(
                    resultado_auto.get(
                        "errores",
                        []
                    )
                )

        st.cache_data.clear()

        incidencias = cargar_incidencias()

    else:

        st.info(
            "El seguimiento completo ya fue actualizado hoy. Las nuevas incidencias se cruzan al guardarse."
        )

st.sidebar.markdown("---")

total_real_sidebar = contar_incidencias_reales_supabase()

st.sidebar.write(
    "Incidencias Supabase filas:",
    total_real_sidebar or len(incidencias)
)

st.sidebar.caption(
    "Modo rápido: sin CPM, inventario ni agenda de citas"
)


st.sidebar.divider()

st.sidebar.caption(
    f"Incidencias registradas: {(total_real_sidebar or len(incidencias)):,}"
)

st.sidebar.caption(
    "Fuente: Supabase + Google Drive"
)


# =========================
# RESUMEN EJECUTIVO
# =========================

if menu == "Resumen Ejecutivo":

    st.subheader(
        "📊 Resumen Ejecutivo"
    )


    st.markdown(
        "### 📄 Reporte Ejecutivo PDF"
    )

    if st.button(
        "📄 Generar reporte ejecutivo PDF",
        use_container_width=True
    ):

        pdf = generar_reporte_ejecutivo_pdf(
            incidencias
        )

        st.download_button(
            label="⬇️ Descargar reporte ejecutivo",
            data=pdf,
            file_name=f"REPORTE_EJECUTIVO_INCIDENCIAS_{fecha_hoy_sistema()}.pdf",
            mime="application/pdf",
            use_container_width=True
        )

    st.divider()
    df_resumen = incidencias.copy()

    if df_resumen.empty:

        st.info(
            "Aún no hay incidencias registradas."
        )

    else:

        if "ESTATUS_SEGUIMIENTO" not in df_resumen.columns:

            df_resumen["ESTATUS_SEGUIMIENTO"] = ""

        df_resumen["ESTATUS_SEGUIMIENTO"] = (
            df_resumen["ESTATUS_SEGUIMIENTO"]
            .astype(str)
            .str.strip()
            .replace(
                {
                    "": "Sin estatus",
                    "nan": "Sin estatus",
                    "None": "Sin estatus"
                }
            )
        )

        if "SEMAFORO" not in df_resumen.columns:

            df_resumen["SEMAFORO"] = df_resumen[
                "ESTATUS_SEGUIMIENTO"
            ].apply(
                semaforo_seguimiento
            )

        total_real_supabase = contar_incidencias_reales_supabase()

        total = total_real_supabase or len(
            df_resumen
        )

        completo_entregado = df_resumen[
            df_resumen["ESTATUS_SEGUIMIENTO"]
            .astype(str)
            .str.upper()
            .eq("COMPLETO-ENTREGADO")
        ].shape[0]

        incompleta_sin_entregar = df_resumen[
            df_resumen["ESTATUS_SEGUIMIENTO"]
            .astype(str)
            .str.upper()
            .eq("INCOMPLETA-SIN ENTREGAR")
        ].shape[0]

        incompleta_cancelada = df_resumen[
            df_resumen["ESTATUS_SEGUIMIENTO"]
            .astype(str)
            .str.upper()
            .eq("INCOMPLETA-CANCELADA")
        ].shape[0]

        sin_estatus = df_resumen[
            df_resumen["ESTATUS_SEGUIMIENTO"]
            .astype(str)
            .str.upper()
            .isin(
                [
                    "",
                    "SIN ESTATUS",
                    "NAN",
                    "NONE"
                ]
            )
        ].shape[0]

        resueltas = 0

        if "ESTATUS_INCIDENCIA" in df_resumen.columns:

            resueltas = df_resumen[
                df_resumen["ESTATUS_INCIDENCIA"]
                .astype(str)
                .str.upper()
                .isin(
                    [
                        "RESUELTA",
                        "RESUELTO"
                    ]
                )
            ].shape[0]

        porcentaje_entregado = 0

        if total > 0:

            porcentaje_entregado = round(
                (completo_entregado / total) * 100,
                2
            )

        c1, c2, c3, c4, c5 = st.columns(
            5
        )

        c1.metric(
            "Total incidencias",
            f"{total:,}"
        )

        c2.metric(
            "🟢 Completo-Entregado",
            f"{completo_entregado:,}"
        )

        c3.metric(
            "🟡 Incompleta-sin Entregar",
            f"{incompleta_sin_entregar:,}"
        )

        c4.metric(
            "🔴 Incompleta-Cancelada",
            f"{incompleta_cancelada:,}"
        )

        c5.metric(
            "% entregado",
            f"{porcentaje_entregado}%"
        )

        st.divider()

        st.subheader(
            "🚦 Semáforo de seguimiento"
        )

        resumen_seguimiento = (
            df_resumen
            .groupby(
                "ESTATUS_SEGUIMIENTO",
                dropna=False
            )
            .size()
            .reset_index(
                name="TOTAL"
            )
            .sort_values(
                "TOTAL",
                ascending=False
            )
        )

        resumen_seguimiento["SEMAFORO"] = resumen_seguimiento[
            "ESTATUS_SEGUIMIENTO"
        ].apply(
            semaforo_seguimiento
        )

        chart_seguimiento = (
            alt.Chart(
                resumen_seguimiento
            )
            .mark_bar()
            .encode(
                x=alt.X(
                    "TOTAL:Q",
                    title="Total"
                ),
                y=alt.Y(
                    "SEMAFORO:N",
                    sort="-x",
                    title="Estatus seguimiento"
                ),
                tooltip=[
                    "SEMAFORO",
                    "TOTAL"
                ]
            )
        )

        st.altair_chart(
            chart_seguimiento,
            use_container_width=True
        )

        dataframe_limpio(
            resumen_seguimiento[
                [
                    "SEMAFORO",
                    "ESTATUS_SEGUIMIENTO",
                    "TOTAL"
                ]
            ]
        )

        st.divider()

        st.subheader(
            "🔎 Filtros ejecutivos"
        )

        f1, f2, f3, f4 = st.columns(
            4
        )

        entidad_filtro = f1.selectbox(
            "Entidad",
            ["Todas"] + sorted(
                df_resumen["ENTIDAD"]
                .dropna()
                .astype(str)
                .unique()
                .tolist()
            )
            if "ENTIDAD" in df_resumen.columns
            else ["Todas"]
        )

        estatus_filtro = f2.selectbox(
            "Estatus seguimiento",
            ["Todos"] + sorted(
                df_resumen["ESTATUS_SEGUIMIENTO"]
                .dropna()
                .astype(str)
                .unique()
                .tolist()
            )
        )

        atribuible_filtro = f3.selectbox(
            "Atribuible a",
            ["Todos"] + sorted(
                df_resumen["ATRIBUIBLE A"]
                .dropna()
                .astype(str)
                .unique()
                .tolist()
            )
            if "ATRIBUIBLE A" in df_resumen.columns
            else ["Todos"]
        )

        responsable_filtro = f4.selectbox(
            "Responsable",
            ["Todos"] + sorted(
                df_resumen["RESPONSABLE"]
                .dropna()
                .astype(str)
                .unique()
                .tolist()
            )
            if "RESPONSABLE" in df_resumen.columns
            else ["Todos"]
        )

        df_filtrado = df_resumen.copy()

        if entidad_filtro != "Todas" and "ENTIDAD" in df_filtrado.columns:

            df_filtrado = df_filtrado[
                df_filtrado["ENTIDAD"].astype(str) == entidad_filtro
            ]

        if estatus_filtro != "Todos":

            df_filtrado = df_filtrado[
                df_filtrado["ESTATUS_SEGUIMIENTO"].astype(str) == estatus_filtro
            ]

        if atribuible_filtro != "Todos" and "ATRIBUIBLE A" in df_filtrado.columns:

            df_filtrado = df_filtrado[
                df_filtrado["ATRIBUIBLE A"].astype(str) == atribuible_filtro
            ]

        if responsable_filtro != "Todos" and "RESPONSABLE" in df_filtrado.columns:

            df_filtrado = df_filtrado[
                df_filtrado["RESPONSABLE"].astype(str) == responsable_filtro
            ]

        st.caption(
            f"Registros filtrados: {len(df_filtrado):,}"
        )

        st.divider()

        col_izq, col_der = st.columns(
            2
        )

        with col_izq:

            st.subheader(
                "🏥 Estados con más incidencias"
            )

            if "ENTIDAD" in df_filtrado.columns:

                top_entidades = (
                    df_filtrado
                    .groupby(
                        "ENTIDAD",
                        dropna=False
                    )
                    .size()
                    .reset_index(
                        name="TOTAL"
                    )
                    .sort_values(
                        "TOTAL",
                        ascending=False
                    )
                    .head(
                        10
                    )
                )

                if not top_entidades.empty:

                    chart_entidades = (
                        alt.Chart(
                            top_entidades
                        )
                        .mark_bar()
                        .encode(
                            x=alt.X(
                                "TOTAL:Q",
                                title="Total"
                            ),
                            y=alt.Y(
                                "ENTIDAD:N",
                                sort="-x",
                                title="Entidad"
                            ),
                            tooltip=[
                                "ENTIDAD",
                                "TOTAL"
                            ]
                        )
                    )

                    st.altair_chart(
                        chart_entidades,
                        use_container_width=True
                    )

                    dataframe_limpio(
                        top_entidades
                    )

        with col_der:

            st.subheader(
                "🏢 Proveedores con más incidencias"
            )

            if "PROVEEDOR" in df_filtrado.columns:

                top_proveedores = (
                    df_filtrado
                    .groupby(
                        "PROVEEDOR",
                        dropna=False
                    )
                    .size()
                    .reset_index(
                        name="TOTAL"
                    )
                    .sort_values(
                        "TOTAL",
                        ascending=False
                    )
                    .head(
                        10
                    )
                )

                if not top_proveedores.empty:

                    chart_proveedores = (
                        alt.Chart(
                            top_proveedores
                        )
                        .mark_bar()
                        .encode(
                            x=alt.X(
                                "TOTAL:Q",
                                title="Total"
                            ),
                            y=alt.Y(
                                "PROVEEDOR:N",
                                sort="-x",
                                title="Proveedor"
                            ),
                            tooltip=[
                                "PROVEEDOR",
                                "TOTAL"
                            ]
                        )
                    )

                    st.altair_chart(
                        chart_proveedores,
                        use_container_width=True
                    )

                    dataframe_limpio(
                        top_proveedores
                    )

        st.divider()

        col_a, col_b = st.columns(
            2
        )

        with col_a:

            st.subheader(
                "📌 Tipos de incidencia"
            )

            if "TIPO_INCIDENCIA" in df_filtrado.columns:

                top_tipos = (
                    df_filtrado
                    .groupby(
                        "TIPO_INCIDENCIA",
                        dropna=False
                    )
                    .size()
                    .reset_index(
                        name="TOTAL"
                    )
                    .sort_values(
                        "TOTAL",
                        ascending=False
                    )
                    .head(
                        10
                    )
                )

                if not top_tipos.empty:

                    st.altair_chart(
                        alt.Chart(
                            top_tipos
                        )
                        .mark_bar()
                        .encode(
                            x=alt.X(
                                "TOTAL:Q",
                                title="Total"
                            ),
                            y=alt.Y(
                                "TIPO_INCIDENCIA:N",
                                sort="-x",
                                title="Tipo incidencia"
                            ),
                            tooltip=[
                                "TIPO_INCIDENCIA",
                                "TOTAL"
                            ]
                        ),
                        use_container_width=True
                    )

                    dataframe_limpio(
                        top_tipos
                    )

        with col_b:

            st.subheader(
                "👤 Atribuible a"
            )

            if "ATRIBUIBLE A" in df_filtrado.columns:

                top_atribuible = (
                    df_filtrado
                    .groupby(
                        "ATRIBUIBLE A",
                        dropna=False
                    )
                    .size()
                    .reset_index(
                        name="TOTAL"
                    )
                    .sort_values(
                        "TOTAL",
                        ascending=False
                    )
                    .head(
                        10
                    )
                )

                if not top_atribuible.empty:

                    st.altair_chart(
                        alt.Chart(
                            top_atribuible
                        )
                        .mark_bar()
                        .encode(
                            x=alt.X(
                                "TOTAL:Q",
                                title="Total"
                            ),
                            y=alt.Y(
                                "ATRIBUIBLE A:N",
                                sort="-x",
                                title="Atribuible a"
                            ),
                            tooltip=[
                                "ATRIBUIBLE A",
                                "TOTAL"
                            ]
                        ),
                        use_container_width=True
                    )

                    dataframe_limpio(
                        top_atribuible
                    )

        st.divider()

        st.subheader(
            "⚠️ Riesgo operativo"
        )

        riesgo = df_filtrado[
            df_filtrado["ESTATUS_SEGUIMIENTO"]
            .astype(str)
            .str.upper()
            .isin(
                [
                    "INCOMPLETA-SIN ENTREGAR",
                    "INCOMPLETA-CANCELADA"
                ]
            )
        ].copy()

        if riesgo.empty:

            st.success(
                "No hay registros de riesgo operativo con los filtros seleccionados."
            )

        else:

            riesgo["SEMAFORO"] = riesgo[
                "ESTATUS_SEGUIMIENTO"
            ].apply(
                semaforo_seguimiento
            )

            columnas_riesgo = [
                "SEMAFORO",
                "ENTIDAD",
                "ORDEN",
                "CLAVE_CNIS",
                "PROVEEDOR",
                "TIPO_INCIDENCIA",
                "ATRIBUIBLE A",
                "ESTATUS_SEGUIMIENTO",
                "ESTATUS_INCIDENCIA",
                "RESPONSABLE"
            ]

            columnas_riesgo = [
                c for c in columnas_riesgo
                if c in riesgo.columns
            ]

            dataframe_limpio(
                riesgo[columnas_riesgo].head(
                    300
                )
            )

        st.divider()

        with st.expander(
            "⬇️ Exportar resumen filtrado"
        ):

            excel = convertir_excel(
                df_filtrado
            )

            st.download_button(
                label="⬇️ Descargar resumen ejecutivo en Excel",
                data=excel,
                file_name="resumen_ejecutivo_incidencias.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )


# =========================
# REGISTRAR INCIDENCIA
# =========================

if menu == "Registrar incidencia":

    st.subheader(
        "📝 Registrar incidencia"
    )

    st.caption(
        "Modo rápido: compendio e incidencias se consultan directo en Supabase."
    )

    modo_registro = st.radio(
        "Tipo de captura",
        [
            "Individual",
            "Masiva"
        ],
        horizontal=True
    )

    if modo_registro == "Individual":

        valor_busqueda = st.text_input(
            "Orden de suministro",
            placeholder="Ejemplo: CIMB-28-01-2025-28030776-U013"
        )

        if valor_busqueda:

            with st.spinner(
                "Buscando orden en Supabase..."
            ):

                datos_orden, resultado = obtener_datos_orden_para_registro(
                    valor_busqueda
                )

            if datos_orden is None:

                st.warning(
                    "No encontré esa orden exacta."
                )

                sugerencias = sugerir_ordenes(
                    valor_busqueda
                )

                if len(sugerencias) > 0:

                    st.info(
                        "Posibles coincidencias:"
                    )

                    dataframe_limpio(
                        sugerencias
                    )

                else:

                    st.error(
                        "No encontré coincidencias en Supabase."
                    )

            else:

                estatus_base = datos_orden["estatus_base"]
                origen_compendio = datos_orden["origen_compendio"]
                orden = datos_orden["orden"]
                tipo_entrega = datos_orden["tipo_entrega"]
                entidad = datos_orden["entidad"]
                clues_destino = datos_orden["clues_destino"]
                unidad_destino = datos_orden["unidad_destino"]
                almacen = datos_orden["almacen"]
                proveedor = datos_orden["proveedor"]
                clave = datos_orden["clave"]
                descripcion = datos_orden["descripcion"]
                piezas_emitidas = datos_orden["piezas_emitidas"]
                piezas_recibidas_ol = datos_orden["piezas_recibidas_ol"]
                piezas_entregadas = datos_orden["piezas_entregadas"]
                operador = datos_orden["operador"]
                tipo_red = datos_orden["tipo_red"]
                grupo_terapeutico = datos_orden["grupo_terapeutico"]
                estatus_orden = datos_orden["estatus_orden"]
                estatus_recepcion_ol = datos_orden.get(
                    "estatus_recepcion_ol",
                    ""
                )
                estatus_entrega_estado = datos_orden.get(
                    "estatus_entrega_estado",
                    ""
                )
                estatus_completa = datos_orden.get(
                    "estatus_completa",
                    ""
                )
                incidencias_previas = obtener_incidencias_previas_supabase(
                    orden
                )

                if incidencias_previas.empty:

                    incidencias_previas = obtener_incidencias_previas(
                        incidencias,
                        orden
                    )

                st.divider()

                c_estado, c_previas = st.columns(
                    2
                )

                with c_estado:

                    if str(
                        estatus_base
                    ).upper().strip() == "INACTIVA":

                        st.error(
                            "🚫 Orden CANCELADA / INACTIVA"
                        )

                    else:

                        st.success(
                            "✅ Orden ACTIVA"
                        )

                with c_previas:

                    if len(
                        incidencias_previas
                    ) > 0:

                        st.warning(
                            f"⚠️ Ya tiene {len(incidencias_previas)} incidencia(s)"
                        )

                    else:

                        st.success(
                            "🟢 Sin incidencias previas"
                        )

                if len(
                    incidencias_previas
                ) > 0:

                    with st.expander(
                        "Ver incidencias previas"
                    ):

                        dataframe_limpio(
                            incidencias_previas
                        )

                st.subheader(
                    "📋 Datos encontrados en Supabase"
                )

                dataframe_limpio(
                    resultado.head(
                        50
                    )
                )

                st.divider()

                st.subheader(
                    "🔒 Información de la orden"
                )

                c1, c2, c3 = st.columns(
                    3
                )

                c1.text_input(
                    "Entidad",
                    entidad,
                    disabled=True
                )

                c2.text_input(
                    "Almacén / CLUES destino",
                    almacen,
                    disabled=True
                )

                c3.text_input(
                    "Proveedor",
                    proveedor,
                    disabled=True
                )

                c4, c5, c6 = st.columns(
                    3
                )

                c4.text_input(
                    "Orden",
                    orden,
                    disabled=True
                )

                c5.text_input(
                    "Clave CNIS",
                    clave,
                    disabled=True
                )

                c6.text_input(
                    "Tipo de entrega",
                    tipo_entrega,
                    disabled=True
                )

                c7, c8, c9 = st.columns(
                    3
                )

                c7.text_input(
                    "Estatus base",
                    estatus_base,
                    disabled=True
                )

                c8.text_input(
                    "Origen compendio",
                    origen_compendio,
                    disabled=True
                )

                c9.text_input(
                    "Estatus orden",
                    estatus_orden,
                    disabled=True
                )

                c10, c11 = st.columns(
                    2
                )

                c10.text_input(
                    "Tipo de red",
                    tipo_red,
                    disabled=True
                )

                c11.text_input(
                    "Grupo terapéutico",
                    grupo_terapeutico,
                    disabled=True
                )

                st.text_area(
                    "Descripción",
                    descripcion,
                    disabled=True
                )

                st.divider()

                st.subheader(
                    "🚚 Validación logística"
                )

                c12, c13, c14 = st.columns(
                    3
                )

                c12.metric(
                    "Piezas emitidas",
                    piezas_emitidas
                )

                c13.metric(
                    "Piezas recibidas OL",
                    piezas_recibidas_ol
                )

                c14.metric(
                    "Piezas entregadas CLUES",
                    piezas_entregadas
                )

                c15, c16 = st.columns(
                    2
                )

                c15.text_input(
                    "Estatus recepción OL",
                    estatus_recepcion_ol,
                    disabled=True
                )

                c16.text_input(
                    "Estatus entrega Estado / CLUES",
                    estatus_entrega_estado,
                    disabled=True
                )

                st.text_input(
                    "Incidencia automática",
                    estatus_completa,
                    disabled=True
                )

                st.divider()

                st.subheader(
                    "✍️ Captura de incidencia"
                )

                c17, c18, c19 = st.columns(
                    3
                )

                with c17:

                    atribuible = st.selectbox(
                        "Atribuible a",
                        ATRIBUIBLES
                    )

                with c18:

                    tipo = st.selectbox(
                        "Tipo de incidencia",
                        TIPOS_INCIDENCIA_GENERAL
                    )

                with c19:

                    estatus = st.selectbox(
                        "Estatus incidencia",
                        [
                            "En proceso",
                            "Rechazado",
                            "Resuelto"
                        ]
                    )

                responsable = st.selectbox(
                    "Responsable",
                    MONITORES
                )

                fecha_incidencia = st.date_input(
                    "Fecha de la incidencia",
                    value=datetime.now().date(),
                    format="DD/MM/YYYY",
                    help="Fecha en que ocurrió la incidencia. La fecha de captura se conserva por separado."
                )

                observaciones = st.text_area(
                    "Observaciones"
                )

                st.subheader(
                    "📎 Evidencias"
                )

                c20, c21 = st.columns(
                    2
                )

                with c20:

                    cedula_rechazo = st.file_uploader(
                        "Cédula rechazo PDF",
                        type=[
                            "pdf"
                        ]
                    )

                    fecha_rechazo = st.date_input(
                        "Fecha del rechazo",
                        value=datetime.now().date(),
                        format="DD/MM/YYYY",
                        help="Fecha en que ocurrió el rechazo, independiente de la fecha de captura."
                    )

                with c21:

                    correo_seguimiento = st.file_uploader(
                        "Correo seguimiento PDF",
                        type=[
                            "pdf"
                        ]
                    )

                guardar = st.button(
                    "💾 Guardar incidencia",
                    use_container_width=True
                )

                if guardar:

                    ruta_cedula = subir_pdf_evidencia_drive(
                        cedula_rechazo,
                        orden,
                        "cedula",
                        entidad,
                        clues_destino
                    )

                    ruta_correo = subir_pdf_evidencia_drive(
                        correo_seguimiento,
                        orden,
                        "correo",
                        entidad,
                        clues_destino
                    )

                    observaciones_guardar = (
                        agregar_fecha_rechazo_observaciones(observaciones, fecha_rechazo)
                        if cedula_rechazo is not None
                        else observaciones
                    )

                    nueva = construir_registro_incidencia(
                        valor_busqueda,
                        datos_orden,
                        atribuible,
                        tipo,
                        estatus,
                        responsable,
                        observaciones_guardar,
                        fecha_incidencia,
                        ruta_cedula,
                        ruta_correo
                    )

                    duplicada = existe_incidencia_duplicada(
                        orden,
                        tipo,
                        atribuible
                    )

                    if duplicada:

                        st.warning(
                            "Esta orden ya tiene una incidencia igual registrada."
                        )

                    else:

                        guardar_incidencia(
                            nueva
                        )

                        st.cache_data.clear()

                        link_respaldo = generar_respaldo_drive()

                        if link_respaldo:

                            st.success(
                                "Incidencia guardada correctamente. Evidencias en Drive y Google Sheets actualizado."
                            )

                        else:

                            st.success(
                                "Incidencia guardada correctamente. Las evidencias se enviaron a Drive."
                            )

                        st.cache_data.clear()

                        st.rerun()

    else:

        st.info(
            "Pega una orden por línea. Se aplicará la misma incidencia a todas."
        )

        texto_ordenes = st.text_area(
            "Órdenes de suministro",
            height=220,
            placeholder="IMBB-16-02-2025-16244042-U013\\nIMBB-16-02-2025-16244043-U013"
        )

        ordenes = extraer_ordenes_masivas(
            texto_ordenes
        )

        st.caption(
            f"Órdenes detectadas: {len(ordenes)}"
        )

        c1, c2, c3 = st.columns(
            3
        )

        with c1:

            atribuible_m = st.selectbox(
                "Atribuible a",
                ATRIBUIBLES,
                key="masivo_atribuible"
            )

        with c2:

            tipo_m = st.selectbox(
                "Tipo de incidencia",
                TIPOS_INCIDENCIA_GENERAL,
                key="masivo_tipo"
            )

        with c3:

            estatus_m = st.selectbox(
                "Estatus incidencia",
                [
                    "En proceso",
                    "Rechazado",
                    "Resuelto"
                ],
                key="masivo_estatus"
            )

        responsable_m = st.selectbox(
            "Responsable",
            MONITORES,
            key="masivo_responsable"
        )

        fecha_incidencia_m = st.date_input(
            "Fecha de la incidencia",
            value=datetime.now().date(),
            format="DD/MM/YYYY",
            key="masivo_fecha_incidencia"
        )

        observaciones_m = st.text_area(
            "Observaciones",
            key="masivo_observaciones"
        )

        validar_masivo = st.button(
            "🔎 Validar órdenes",
            use_container_width=True
        )

        if validar_masivo and ordenes:

            registros_preview = []

            for orden_m in ordenes:

                datos_m, resultado_m = obtener_datos_orden_para_registro(
                    orden_m
                )

                if datos_m is None:

                    registros_preview.append(
                        {
                            "ORDEN": orden_m,
                            "ESTATUS": "NO ENCONTRADA",
                            "ENTIDAD": "",
                            "CLUES": "",
                            "PROVEEDOR": "",
                            "INCIDENCIAS_PREVIAS": 0
                        }
                    )

                    continue

                previas_m = obtener_incidencias_previas_supabase(
                    datos_m["orden"]
                )

                registros_preview.append(
                    {
                        "ORDEN": datos_m["orden"],
                        "ESTATUS": datos_m["estatus_base"],
                        "ENTIDAD": datos_m["entidad"],
                        "CLUES": datos_m["clues_destino"],
                        "PROVEEDOR": datos_m["proveedor"],
                        "INCIDENCIAS_PREVIAS": len(previas_m)
                    }
                )

            st.session_state["preview_masivo"] = registros_preview

        if "preview_masivo" in st.session_state:

            preview_df = pd.DataFrame(
                st.session_state["preview_masivo"]
            )

            dataframe_limpio(
                preview_df
            )

            encontradas = preview_df[
                preview_df["ESTATUS"] != "NO ENCONTRADA"
            ].shape[0]

            st.success(
                f"Listas para guardar: {encontradas} de {len(preview_df)}"
            )

        guardar_masivo = st.button(
            "💾 Guardar incidencia para todas las órdenes encontradas",
            use_container_width=True
        )

        if guardar_masivo:

            if not ordenes:

                st.warning(
                    "Primero pega al menos una orden."
                )

            else:

                guardadas = 0

                no_encontradas = []

                errores = []

                with st.spinner(
                    "Guardando incidencias masivas..."
                ):

                    for orden_m in ordenes:

                        try:

                            datos_m, resultado_m = obtener_datos_orden_para_registro(
                                orden_m
                            )

                            if datos_m is None:

                                no_encontradas.append(
                                    orden_m
                                )

                                continue

                            duplicada_m = existe_incidencia_duplicada(
                                datos_m["orden"],
                                tipo_m,
                                atribuible_m
                            )

                            if duplicada_m:

                                continue

                            nueva_m = construir_registro_incidencia(
                                orden_m,
                                datos_m,
                                atribuible_m,
                                tipo_m,
                                estatus_m,
                                responsable_m,
                                observaciones_m,
                                fecha_incidencia_m,
                                "",
                                ""
                            )

                            guardar_incidencia(
                                nueva_m
                            )

                            guardadas += 1

                        except Exception as e:

                            errores.append(
                                f"{orden_m}: {e}"
                            )

                link_respaldo = ""

                if guardadas > 0:

                    st.cache_data.clear()

                    link_respaldo = generar_respaldo_drive()

                if link_respaldo:

                    st.success(
                        f"Guardadas correctamente: {guardadas}. Google Sheets actualizado."
                    )

                else:

                    st.success(
                        f"Guardadas correctamente: {guardadas}"
                    )

                if no_encontradas:

                    st.warning(
                        f"No encontradas: {len(no_encontradas)}"
                    )

                    st.write(
                        no_encontradas
                    )

                if errores:

                    st.error(
                        f"Errores: {len(errores)}"
                    )

                    st.write(
                        errores
                    )


# =========================
# URGENCIAS
# =========================

elif menu == "Urgencias":

    st.subheader(
        "🚨 Registro de urgencias"
    )

    st.caption(
        "Captura la clave, la entidad y la fecha de solicitud. "
        "El sistema localizará las órdenes existentes y conservará únicamente "
        "las que no tengan entrega registrada en la entidad."
    )

    with st.expander(
        "📊 Historial de urgencias y órdenes pendientes de entrega",
        expanded=True
    ):

        datos_exportacion_urgencias = cargar_urgencias_para_exportacion()
        urgencias_capturadas_export = datos_exportacion_urgencias[
            "claves"
        ]
        ordenes_pendientes_export = datos_exportacion_urgencias[
            "pendientes"
        ]
        urgencias_atendidas_export = datos_exportacion_urgencias[
            "atendidas"
        ]
        urgencias_sin_orden_export = datos_exportacion_urgencias[
            "sin_orden"
        ]

        m_export_1, m_export_2, m_export_3, m_export_4 = st.columns(4)

        m_export_1.metric(
            "Urgencias (claves)",
            len(urgencias_capturadas_export)
        )

        m_export_2.metric(
            "Órdenes pendientes",
            len(ordenes_pendientes_export)
        )

        m_export_3.metric(
            "Atendidas (claves)",
            len(urgencias_atendidas_export)
        )

        m_export_4.metric(
            "Sin orden / revisar emisión",
            len(urgencias_sin_orden_export)
        )

        st.caption(
            "Una urgencia equivale a una clave CNIS por entidad; las filas "
            "inactivas o canceladas no se cuentan."
        )

        st.caption("Urgencias utiliza exclusivamente BASE_URGENCIAS en Google Sheets. "
                   "No se trasladan ni eliminan registros de Supabase.")

        c_export_1, c_export_2 = st.columns(2)

        with c_export_1:
            if st.button(
                "🔄 Actualizar capturas",
                use_container_width=True,
                key="urgencias_actualizar_exportacion"
            ):
                st.cache_data.clear()
                st.rerun()

        with c_export_2:
            excel_urgencias = convertir_excel_urgencias(
                datos_exportacion_urgencias
            )

            st.download_button(
                label="⬇️ Descargar Excel de urgencias (pendientes y atendidas)",
                data=excel_urgencias,
                file_name=(
                    f"urgencias_pendientes_atendidas_{fecha_hoy_sistema()}.xlsx"
                ),
                mime=(
                    "application/vnd.openxmlformats-officedocument."
                    "spreadsheetml.sheet"
                ),
                use_container_width=True,
                key="urgencias_descargar_excel"
            )

        if not ordenes_pendientes_export.empty:
            st.caption(
                "Vista previa de las órdenes con piezas pendientes de entrega "
                "para enviarlas a programación y surtimiento."
            )

            columnas_pendientes_export = [
                "FECHA_URGENCIA",
                "ENTIDAD",
                "CLAVE_CNIS",
                "ORDEN_SUMINISTRO",
                "CLUES_DESTINO",
                "UNIDAD_DESTINO",
                "PROVEEDOR",
                "DESCRIPCION",
                "PIEZAS_EMITIDAS",
                "PIEZAS_ENTREGADAS_CLUES",
                "PENDIENTE_ENTREGA",
                "ESTATUS_ENTREGA_ENTIDAD",
                "RESPONSABLE"
            ]

            dataframe_limpio(
                ordenes_pendientes_export[
                    [
                        columna for columna in columnas_pendientes_export
                        if columna in ordenes_pendientes_export.columns
                    ]
                ].head(300)
            )
        else:
            st.info(
                "No hay órdenes pendientes de entrega en las urgencias capturadas."
            )

    c1, c2, c3 = st.columns(
        [
            1,
            2,
            1
        ]
    )

    with c1:

        clave_urgencia = st.text_input(
            "Clave CNIS",
            placeholder="Ejemplo: 010.000.0012.00",
            key="urgencias_clave"
        ).strip()

    with c2:

        entidades_urgencia = obtener_entidades_compendio_urgencia()

        opciones_entidades_urgencia = [
            "Selecciona una entidad"
        ] + entidades_urgencia

        entidad_urgencia = st.selectbox(
            "Estado / entidad",
            opciones_entidades_urgencia,
            key="urgencias_entidad"
        )

        if entidad_urgencia == "Selecciona una entidad":

            entidad_urgencia = ""

    with c3:

        fecha_urgencia = st.date_input(
            "Fecha de solicitud",
            value=datetime.now().date(),
            key="urgencias_fecha"
        )

    responsable_urgencia = st.selectbox(
        "Monitora / responsable",
        MONITORES,
        key="urgencias_responsable"
    )

    observaciones_urgencia_captura = st.text_area(
        "Observaciones de la urgencia",
        placeholder="Describe brevemente la necesidad o el seguimiento requerido.",
        key="urgencias_observaciones"
    )

    pdf_urgencia = st.file_uploader(
        "📎 Adjuntar PDF de correo o evidencia",
        type=[
            "pdf"
        ],
        key="urgencias_pdf",
        help="Selecciona el correo o la evidencia que respalda la urgencia."
    )

    buscar_urgencia = st.button(
        "🔎 Consultar órdenes de la clave",
        use_container_width=True,
        key="urgencias_buscar"
    )

    if buscar_urgencia:

        if not clave_urgencia or not entidad_urgencia:

            st.warning(
                "Captura la clave CNIS y el estado / entidad antes de consultar."
            )

        else:

            with st.spinner(
                "Consultando órdenes y validando entregas en la entidad..."
            ):

                resultado_urgencia = buscar_ordenes_urgencia(
                    clave_urgencia,
                    entidad_urgencia
                )

            st.session_state["resultado_urgencias"] = {
                "clave": normalizar_clave_urgencia(
                    clave_urgencia
                ),
                "entidad": normalizar_entidad_urgencia(
                    entidad_urgencia
                ),
                "fecha": fecha_urgencia_iso(
                    fecha_urgencia
                ),
                "resultado": resultado_urgencia
            }

    resultado_guardado = st.session_state.get(
        "resultado_urgencias"
    )

    consulta_actual = (
        resultado_guardado
        and resultado_guardado.get("clave") == normalizar_clave_urgencia(
            clave_urgencia
        )
        and resultado_guardado.get("entidad") == normalizar_entidad_urgencia(
            entidad_urgencia
        )
        and resultado_guardado.get("fecha") == fecha_urgencia_iso(
            fecha_urgencia
        )
    )

    if consulta_actual:

        resultado_urgencia = resultado_guardado["resultado"]
        ordenes_urgencia = resultado_urgencia.get(
            "ordenes",
            pd.DataFrame()
        )
        ordenes_entregadas_urgencia = resultado_urgencia.get(
            "ordenes_entregadas",
            pd.DataFrame()
        )
        ordenes_emitidas_urgencia = resultado_urgencia.get(
            "ordenes_emitidas",
            pd.DataFrame()
        )
        revision_pendiente_urgencia = None

        if not ordenes_emitidas_urgencia.empty:

            revision_pendiente_urgencia = obtener_urgencia_revision_pendiente(
                incidencias,
                clave_urgencia,
                entidad_urgencia
            )

        st.divider()
        st.subheader(
            "Resultado de la consulta"
        )

        m1, m2, m3, m4 = st.columns(
            4
        )

        m1.metric(
            "Registros de la clave",
            resultado_urgencia.get("total_clave", 0)
        )

        m2.metric(
            "Órdenes en la entidad",
            resultado_urgencia.get("total_con_orden", 0)
        )

        m3.metric(
            "No entregadas elegibles",
            len(ordenes_urgencia)
        )

        m4.metric(
            "Excluidas por entrega",
            len(ordenes_entregadas_urgencia)
        )

        if not resultado_urgencia.get(
            "clave_valida",
            False
        ):

            detalle_consulta = resultado_urgencia.get(
                "error_consulta",
                ""
            )

            if detalle_consulta:

                st.error(
                    "No se pudo consultar el compendio desde la aplicación. "
                    "La clave no se marcó como inexistente."
                )

                st.caption(
                    "Detalle técnico: "
                    f"{detalle_consulta}"
                )

            else:

                st.error(
                    "La clave no fue localizada en la base consultada. "
                    "Verifica el formato de la clave o la sincronización del compendio."
                )

        elif not resultado_urgencia.get(
            "entidad_homologada",
            ""
        ):

            if resultado_urgencia.get("sin_orden", False):

                st.warning(
                    "La clave existe en el compendio, pero todavía no tiene "
                    "emisión para la entidad seleccionada. Se puede registrar "
                    "como REVISAR CON EMISIÓN."
                )

            else:

                st.error(
                    "La clave existe en el compendio, pero no tiene registros "
                    "para la entidad seleccionada."
                )

        else:

            st.success(
                "Entidad localizada en el compendio: "
                f"{resultado_urgencia['entidad_homologada']}"
            )

        if not ordenes_urgencia.empty:

            st.success(
                "Estas son las órdenes que sí se pueden registrar como urgencia "
                "porque no tienen entrega registrada en la entidad."
            )

            columnas_preview = [
                "CLAVE_CNIS",
                "ENTIDAD",
                "ORDEN_SUMINISTRO",
                "ESTATUS_BASE",
                "TIPO_ENTREGA",
                "CLUES_DESTINO",
                "UNIDAD_DESTINO",
                "PROVEEDOR",
                "PIEZAS_EMITIDAS",
                "PIEZAS_ENTREGADAS_CLUES",
                "ESTATUS_ENTREGA_ENTIDAD"
            ]

            dataframe_limpio(
                ordenes_urgencia[
                    [
                        columna for columna in columnas_preview
                        if columna in ordenes_urgencia.columns
                    ]
                ]
            )

        elif (
            resultado_urgencia.get("sin_orden", False)
            and resultado_urgencia.get("entidad_homologada", "")
        ):

            st.warning(
                "No se encontró una orden de suministro para la clave y entidad. "
                "La solicitud se guardará como REVISAR CON EMISIÓN en la hoja "
                "Sin_Orden_Urgencias."
            )

        elif (
            not ordenes_entregadas_urgencia.empty
            and revision_pendiente_urgencia is None
        ):

            st.warning(
                "La clave sí tiene órdenes en la entidad, pero todas tienen "
                "entrega registrada. No se guardará ninguna orden para la urgencia."
            )

        if not ordenes_entregadas_urgencia.empty:

            with st.expander(
                "Ver órdenes excluidas por tener entrega en la entidad"
            ):

                columnas_excluidas = [
                    "ORDEN_SUMINISTRO",
                    "PIEZAS_ENTREGADAS_CLUES",
                    "ESTATUS_ENTREGA_ENTIDAD",
                    "MOTIVO_EXCLUSION"
                ]

                dataframe_limpio(
                    ordenes_entregadas_urgencia[
                        [
                            columna for columna in columnas_excluidas
                            if columna in ordenes_entregadas_urgencia.columns
                        ]
                    ]
                )

        puede_guardar_urgencia = (
            resultado_urgencia.get("clave_valida", False)
            and bool(entidad_urgencia)
            and (
                not ordenes_urgencia.empty
                or resultado_urgencia.get("sin_orden", False)
                or revision_pendiente_urgencia is not None
            )
        )

        if puede_guardar_urgencia:

            st.divider()

            texto_boton_guardar = (
                "💾 Registrar como REVISAR CON EMISIÓN"
                if resultado_urgencia.get("requiere_revision_emision", False)
                else "💾 Actualizar solicitud pendiente con emisión"
                if revision_pendiente_urgencia is not None
                else "💾 Registrar urgencia en base independiente"
            )

            guardar_urgencia = st.button(
                texto_boton_guardar,
                use_container_width=True,
                key="urgencias_guardar"
            )

            if guardar_urgencia:

                with st.spinner(
                    "Registrando la urgencia únicamente en BASE_URGENCIAS..."
                ):

                    resultado_guardado_urgencia = registrar_urgencia_completa(
                        clave_urgencia,
                        resultado_urgencia.get(
                            "entidad_homologada",
                            ""
                        ) or entidad_urgencia,
                        fecha_urgencia,
                        responsable_urgencia,
                        observaciones_urgencia_captura,
                        resultado_urgencia,
                        incidencias,
                        pdf_urgencia
                    )

                guardadas_urgencia = resultado_guardado_urgencia.get(
                    "guardadas",
                    0
                )
                actualizadas_urgencia = resultado_guardado_urgencia.get(
                    "actualizadas",
                    0
                )

                if guardadas_urgencia - actualizadas_urgencia > 0:

                    st.success(
                        "Urgencia registrada en BASE_URGENCIAS; no se creó una incidencia."
                    )

                if actualizadas_urgencia > 0:

                    st.success(
                        "La solicitud REVISAR CON EMISIÓN se actualizó con "
                        "los datos de la emisión localizada en el compendio."
                    )

                if resultado_guardado_urgencia.get("duplicadas", 0) > 0:

                    st.info(
                        "Se omitieron registros que ya estaban capturados "
                        "para la misma clave, entidad, orden y fecha."
                    )

                drive_urgencia = resultado_guardado_urgencia.get(
                    "drive"
                )

                if drive_urgencia:

                    st.success(
                        "Base de Urgencias actualizada en Drive: "
                        f"{drive_urgencia.get('ordenes', 0)} órdenes y "
                        f"{drive_urgencia.get('sin_orden', 0)} solicitudes sin orden."
                    )

                    st.link_button(
                        "🔗 Abrir base de Urgencias en Drive",
                        drive_urgencia.get("link", ""),
                        use_container_width=True
                    )

                if resultado_guardado_urgencia.get("errores"):

                    st.error(
                        "Se presentaron errores durante el registro."
                    )

                    st.write(
                        resultado_guardado_urgencia.get("errores")
                    )

                st.cache_data.clear()
                st.session_state.pop(
                    "resultado_urgencias",
                    None
                )

    st.divider()

    st.subheader(
        "📋 Carga masiva de claves"
    )

    st.caption(
        "Pega una clave por línea o varias separadas por coma, punto y coma "
        "o tabulador. Todas se validarán contra el compendio."
    )

    texto_claves_urgencias = st.text_area(
        "Claves CNIS para carga masiva",
        height=180,
        placeholder="010.000.0012.00\n010.000.0045.00\n010.000.0099.00",
        key="urgencias_claves_masivas"
    )

    claves_urgencias_masivas = extraer_claves_urgencia_masivas(
        texto_claves_urgencias
    )

    st.caption(
        f"Claves detectadas: {len(claves_urgencias_masivas)}"
    )

    consultar_masivo_urgencias = st.button(
        "🔎 Validar claves masivas",
        use_container_width=True,
        key="urgencias_validar_masivo"
    )

    if consultar_masivo_urgencias:

        if not claves_urgencias_masivas:

            st.warning(
                "Pega al menos una clave para realizar la validación."
            )

        elif not entidad_urgencia:

            st.warning(
                "Selecciona una entidad antes de validar las claves."
            )

        else:

            resultados_masivos = []

            with st.spinner(
                "Validando claves contra el compendio y revisando entregas..."
            ):

                for clave_masiva in claves_urgencias_masivas:

                    resultado_masivo = buscar_ordenes_urgencia(
                        clave_masiva,
                        entidad_urgencia
                    )

                    resultados_masivos.append(
                        {
                            "clave": clave_masiva,
                            "resultado": resultado_masivo
                        }
                    )

            st.session_state[
                "resultados_urgencias_masivas"
            ] = {
                "entidad": normalizar_entidad_urgencia(
                    entidad_urgencia
                ),
                "fecha": fecha_urgencia_iso(
                    fecha_urgencia
                ),
                "resultados": resultados_masivos
            }

    resultados_masivos_guardados = st.session_state.get(
        "resultados_urgencias_masivas"
    )

    consulta_masiva_actual = (
        resultados_masivos_guardados
        and resultados_masivos_guardados.get("entidad") == normalizar_entidad_urgencia(
            entidad_urgencia
        )
        and resultados_masivos_guardados.get("fecha") == fecha_urgencia_iso(
            fecha_urgencia
        )
        and resultados_masivos_guardados.get("resultados")
    )

    if consulta_masiva_actual:

        resumen_masivo = []

        for item_masivo in resultados_masivos_guardados["resultados"]:

            resultado_masivo = item_masivo["resultado"]
            ordenes_masivas = resultado_masivo.get(
                "ordenes",
                pd.DataFrame()
            )
            ordenes_entregadas_masivas = resultado_masivo.get(
                "ordenes_entregadas",
                pd.DataFrame()
            )

            if not resultado_masivo.get("clave_valida", False):

                estatus_masivo = "CLAVE NO LOCALIZADA EN LA BASE DEL SISTEMA"

            elif resultado_masivo.get("sin_orden", False):

                estatus_masivo = "REVISAR CON EMISIÓN"

            elif not ordenes_masivas.empty:

                estatus_masivo = "LISTA PARA REGISTRAR"

            elif not ordenes_entregadas_masivas.empty:

                estatus_masivo = "TODAS ENTREGADAS"

            else:

                estatus_masivo = "SIN RESULTADO"

            resumen_masivo.append(
                {
                    "CLAVE_CNIS": item_masivo["clave"],
                    "ESTATUS": estatus_masivo,
                    "ORDENES_NO_ENTREGADAS": len(ordenes_masivas),
                    "ORDENES_EXCLUIDAS_ENTREGADAS": len(
                        ordenes_entregadas_masivas
                    ),
                    "ENTIDAD_EN_COMPENDIO": resultado_masivo.get(
                        "entidad_homologada",
                        ""
                    )
                }
            )

        st.subheader(
            "Resultado de validación masiva"
        )

        dataframe_limpio(
            pd.DataFrame(
                resumen_masivo
            )
        )

        guardar_masivo_urgencias = st.button(
            "💾 Registrar todas las urgencias válidas",
            use_container_width=True,
            key="urgencias_guardar_masivo"
        )

        if guardar_masivo_urgencias:

            total_guardadas_masivo = 0
            total_actualizadas_masivo = 0
            total_duplicadas_masivo = 0
            errores_masivos = []

            with st.spinner(
                "Registrando las urgencias válidas en BASE_URGENCIAS..."
            ):

                for item_masivo in resultados_masivos_guardados["resultados"]:

                    clave_masiva = item_masivo["clave"]
                    resultado_masivo = item_masivo["resultado"]

                    if not resultado_masivo.get("clave_valida", False):

                        errores_masivos.append(
                            f"{clave_masiva}: no localizada en la base del sistema"
                        )
                        continue

                    entidad_masiva = resultado_masivo.get(
                        "entidad_homologada",
                        ""
                    ) or entidad_urgencia

                    if not entidad_masiva:

                        errores_masivos.append(
                            f"{clave_masiva}: no se pudo determinar la entidad seleccionada"
                        )
                        continue

                    resultado_registro_masivo = registrar_urgencia_completa(
                        clave_masiva,
                        entidad_masiva,
                        fecha_urgencia,
                        responsable_urgencia,
                        observaciones_urgencia_captura,
                        resultado_masivo,
                        incidencias,
                        pdf_urgencia
                    )

                    total_guardadas_masivo += resultado_registro_masivo.get(
                        "guardadas",
                        0
                    )
                    total_actualizadas_masivo += resultado_registro_masivo.get(
                        "actualizadas",
                        0
                    )
                    total_duplicadas_masivo += resultado_registro_masivo.get(
                        "duplicadas",
                        0
                    )
                    errores_masivos.extend(
                        resultado_registro_masivo.get(
                            "errores",
                            []
                        )
                    )

            if total_guardadas_masivo - total_actualizadas_masivo:

                st.success(
                    f"Se registraron {total_guardadas_masivo - total_actualizadas_masivo} "
                    "urgencias correctamente en BASE_URGENCIAS."
                )

            if total_actualizadas_masivo:

                st.success(
                    f"Se actualizaron {total_actualizadas_masivo} solicitudes con "
                    "los datos de emisión del compendio."
                )

            if total_duplicadas_masivo:

                st.info(
                    f"Se omitieron {total_duplicadas_masivo} registros duplicados."
                )

            if errores_masivos:

                st.warning(
                    "Algunas claves no se registraron:"
                )
                st.write(
                    errores_masivos
                )

            st.cache_data.clear()
            st.session_state.pop(
                "resultados_urgencias_masivas",
                None
            )


# =========================
# SEGUIMIENTO
# =========================

elif menu == "Seguimiento":

    st.subheader(
        "📋 Seguimiento de incidencias"
    )

    st.caption(
        "El cruce con compendio se actualiza automáticamente solo una vez al día. "
        "Las incidencias nuevas se cruzan al momento de guardarse."
    )

    df_seg = incidencias.copy()

    if df_seg.empty:

        st.info(
            "Aún no hay incidencias registradas."
        )

    else:

        if "ESTATUS_SEGUIMIENTO" not in df_seg.columns:

            df_seg["ESTATUS_SEGUIMIENTO"] = ""

        if "ATRIBUIBLE A" in df_seg.columns:

            df_seg["ATRIBUIBLE A"] = df_seg["ATRIBUIBLE A"].apply(
                homologar_atribuible
            )

        df_seg["SEMAFORO"] = df_seg["ESTATUS_SEGUIMIENTO"].apply(
            semaforo_seguimiento
        )

        # Rango para consultar/descargar incidencias previas por fecha de captura.
        fechas_captura = pd.to_datetime(df_seg["FECHA_REGISTRO"], errors="coerce", dayfirst=True)
        fechas_validas = fechas_captura.dropna()
        if not fechas_validas.empty:
            fmin = fechas_validas.min().date()
            fmax = fechas_validas.max().date()
            rango_fechas = st.date_input(
                "Rango de fecha de captura",
                value=(fmin, fmax),
                min_value=fmin,
                max_value=fmax,
                format="DD/MM/YYYY",
                key="seguimiento_rango_fecha_captura"
            )
            if isinstance(rango_fechas, (tuple, list)) and len(rango_fechas) == 2:
                desde, hasta = rango_fechas
                mascara_fecha = fechas_captura.dt.date.between(desde, hasta)
                df_seg = df_seg[mascara_fecha.fillna(False)].copy()

        c1, c2, c3, c4 = st.columns(
            4
        )

        entidad_filtro = c1.selectbox(
            "Entidad",
            ["Todas"] + sorted(
                df_seg["ENTIDAD"]
                .dropna()
                .astype(str)
                .unique()
                .tolist()
            )
        )

        estatus_filtro = c2.selectbox(
            "Estatus seguimiento",
            ["Todos"] + sorted(
                df_seg["ESTATUS_SEGUIMIENTO"]
                .dropna()
                .astype(str)
                .unique()
                .tolist()
            )
        )

        atribuible_filtro = c3.selectbox(
            "Atribuible a",
            ["Todos"] + sorted(
                df_seg["ATRIBUIBLE A"]
                .dropna()
                .astype(str)
                .unique()
                .tolist()
            )
        )

        buscar_orden = c4.text_input(
            "Buscar orden"
        )

        if entidad_filtro != "Todas":

            df_seg = df_seg[
                df_seg["ENTIDAD"].astype(str) == entidad_filtro
            ]

        if estatus_filtro != "Todos":

            df_seg = df_seg[
                df_seg["ESTATUS_SEGUIMIENTO"].astype(str) == estatus_filtro
            ]

        if atribuible_filtro != "Todos":

            df_seg = df_seg[
                df_seg["ATRIBUIBLE A"].astype(str) == atribuible_filtro
            ]

        if buscar_orden.strip():

            orden_norm = normalizar_orden(
                buscar_orden
            )

            df_seg = df_seg[
                df_seg["ORDEN"]
                .astype(str)
                .apply(normalizar_orden)
                .str.contains(
                    orden_norm,
                    na=False,
                    regex=False
                )
            ]

        st.caption(
            f"Registros mostrados: {len(df_seg):,}"
        )

        columnas_mostrar = [
            "SEMAFORO",
            "FECHA_REGISTRO",
            "FECHA_INCIDENCIA",
            "FECHA_RECHAZO",
            "ENTIDAD",
            "ORDEN",
            "CLAVE_CNIS",
            "PROVEEDOR",
            "TIPO_INCIDENCIA",
            "ATRIBUIBLE A",
            "ESTATUS_SEGUIMIENTO",
            "ESTATUS_INCIDENCIA",
            "RESPONSABLE",
            "OBSERVACIONES",
            "PDF_CEDULA_RECHAZO",
            "PDF_CORREO_SEGUIMIENTO"
        ]

        columnas_mostrar = [
            c for c in columnas_mostrar
            if c in df_seg.columns
        ]

        dataframe_limpio(
            df_seg[columnas_mostrar]
        )

        st.divider()

        with st.expander(
            "📎 Agregar archivo a una incidencia previa"
        ):

            if "ID" not in df_seg.columns:

                st.warning(
                    "No encontré la columna ID para actualizar evidencias."
                )

            else:

                df_archivos = df_seg.copy()

                df_archivos["_OPCION_ARCHIVO"] = df_archivos.apply(
                    lambda x: f"{x.get('ID', '')} | {x.get('ORDEN', '')} | {x.get('CLAVE_CNIS', '')} | {x.get('TIPO_INCIDENCIA', '')}",
                    axis=1
                )

                opcion_archivo = st.selectbox(
                    "Incidencia a actualizar",
                    df_archivos["_OPCION_ARCHIVO"].astype(str).tolist()
                )

                fila_archivo = df_archivos[
                    df_archivos["_OPCION_ARCHIVO"].astype(str) == str(opcion_archivo)
                ].iloc[0]

                tipo_archivo = st.selectbox(
                    "Tipo de archivo",
                    [
                        "Cédula rechazo",
                        "Correo seguimiento"
                    ]
                )

                archivo_nuevo = st.file_uploader(
                    "Archivo PDF",
                    type=[
                        "pdf"
                    ],
                    key="archivo_incidencia_previa"
                )

                fecha_rechazo_previa = None
                if tipo_archivo == "Cédula rechazo":
                    fecha_rechazo_previa = st.date_input(
                        "Fecha del rechazo",
                        value=datetime.now().date(),
                        format="DD/MM/YYYY",
                        key="fecha_rechazo_incidencia_previa"
                    )

                if st.button(
                    "📎 Subir archivo a incidencia previa",
                    use_container_width=True
                ):

                    tipo_pdf = (
                        "cedula"
                        if tipo_archivo == "Cédula rechazo"
                        else "correo"
                    )

                    ok = actualizar_evidencia_incidencia(
                        fila_archivo.get("ID", ""),
                        archivo_nuevo,
                        tipo_pdf,
                        fila_archivo.get("ORDEN", fila_archivo.get("ORDEN_BUSCADA", "")),
                        fila_archivo.get("ENTIDAD", "SIN_ESTADO"),
                        fila_archivo.get("CLUES_DESTINO", "SIN_CLUES")
                    )

                    if ok:

                        if tipo_pdf == "cedula" and fecha_rechazo_previa is not None:
                            obs_actual = fila_archivo.get("OBSERVACIONES", "")
                            obs_guardar = agregar_fecha_rechazo_observaciones(obs_actual, fecha_rechazo_previa)
                            supabase.table("incidencias").update({"observaciones": obs_guardar}).eq(
                                "id", fila_archivo.get("ID", "")
                            ).execute()
                            st.cache_data.clear()

                        st.success(
                            "Archivo agregado correctamente a la incidencia previa."
                        )

                        st.rerun()

                    else:

                        st.error(
                            "No se pudo subir el archivo. Revisa que seleccionaste un PDF."
                        )

        excel = convertir_excel(
            df_seg
        )

        st.download_button(
            label="⬇️ Descargar incidencias filtradas en Excel",
            data=excel,
            file_name=(
                f"reporte_incidencias_{desde.strftime('%Y%m%d')}_{hasta.strftime('%Y%m%d')}.xlsx"
                if 'desde' in locals() and 'hasta' in locals()
                else "reporte_incidencias_filtrado.xlsx"
            ),
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )


# =========================
# CÉDULAS DE RECHAZO
# =========================

elif menu == "Cédulas de rechazo":

    st.subheader("📄 Cédulas de rechazo")
    st.caption("Consulta las cédulas registradas, abre cada PDF con un clic o descarga el concentrado en Excel.")

    df_cedulas = incidencias.copy()

    if "PDF_CEDULA_RECHAZO" not in df_cedulas.columns:
        st.info("Aún no existe información de cédulas de rechazo en la base.")
    else:
        enlaces = (
            df_cedulas["PDF_CEDULA_RECHAZO"]
            .fillna("")
            .astype(str)
            .str.strip()
        )
        df_cedulas = df_cedulas[
            ~enlaces.str.lower().isin(["", "nan", "none", "null"])
        ].copy()

        if df_cedulas.empty:
            st.info("No hay cédulas de rechazo registradas para mostrar.")
        else:
            c1, c2, c3 = st.columns(3)

            with c1:
                if "ENTIDAD" in df_cedulas.columns:
                    entidades = sorted(df_cedulas["ENTIDAD"].fillna("").astype(str).unique().tolist())
                    entidad = st.selectbox("Entidad", ["Todas"] + [x for x in entidades if x.strip()], key="cedulas_entidad")
                    if entidad != "Todas":
                        df_cedulas = df_cedulas[df_cedulas["ENTIDAD"].astype(str) == entidad]

            with c2:
                if "TIPO_INCIDENCIA" in df_cedulas.columns:
                    tipos = sorted(df_cedulas["TIPO_INCIDENCIA"].fillna("").astype(str).unique().tolist())
                    tipo = st.selectbox("Tipo de incidencia", ["Todos"] + [x for x in tipos if x.strip()], key="cedulas_tipo")
                    if tipo != "Todos":
                        df_cedulas = df_cedulas[df_cedulas["TIPO_INCIDENCIA"].astype(str) == tipo]

            with c3:
                orden_buscar = st.text_input("Buscar orden", key="cedulas_orden").strip()
                if orden_buscar:
                    col_orden = "ORDEN" if "ORDEN" in df_cedulas.columns else "ORDEN_BUSCADA"
                    if col_orden in df_cedulas.columns:
                        df_cedulas = df_cedulas[df_cedulas[col_orden].fillna("").astype(str).str.contains(orden_buscar, case=False, na=False)]

            st.metric("Cédulas encontradas", len(df_cedulas))

            columnas = [
                "FECHA_REGISTRO", "FECHA_RECHAZO", "ENTIDAD", "ORDEN", "ORDEN_BUSCADA",
                "CLAVE_CNIS", "PROVEEDOR", "TIPO_INCIDENCIA",
                "ESTATUS_INCIDENCIA", "RESPONSABLE", "PDF_CEDULA_RECHAZO"
            ]
            columnas = [c for c in columnas if c in df_cedulas.columns]
            vista = formatear_fechas_cortas_df(df_cedulas[columnas].copy())

            st.dataframe(
                vista,
                use_container_width=True,
                hide_index=True,
                column_config={
                    "PDF_CEDULA_RECHAZO": st.column_config.LinkColumn(
                        "📄 Abrir PDF",
                        display_text="Ver PDF"
                    )
                }
            )

            excel_cedulas = convertir_excel_cedulas_rechazo(vista)
            st.download_button(
                "⬇️ Descargar cédulas de rechazo en Excel",
                data=excel_cedulas,
                file_name="cedulas_de_rechazo.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )


# =========================
# BASE SUPABASE
# =========================

elif menu == "Base Supabase":

    from modules.editor_incidencias import mostrar_editor
    vista_base = st.radio("Herramienta", ["Consultar compendio", "Editar incidencia"], horizontal=True)
    if vista_base == "Editar incidencia":
        mostrar_editor(supabase)
        st.stop()

    st.subheader(
        "⚡ Base Supabase"
    )

    st.info(
        "Vista rápida de los primeros 100 registros desde Supabase."
    )

    respuesta = (
        supabase
        .table(
            "compendio"
        )
        .select(
            "*"
        )
        .limit(
            100
        )
        .execute()
    )

    muestra = pd.DataFrame(
        respuesta.data
    )

    dataframe_limpio(
        muestra
    )

# version seguimiento 2026	
