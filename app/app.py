def generar_reporte_ejecutivo_pdf(incidencias):

    buffer = BytesIO()

    doc = SimpleDocTemplate(
        buffer,
        pagesize=landscape(letter),
        rightMargin=1 * cm,
        leftMargin=1 * cm,
        topMargin=1.3 * cm,
        bottomMargin=1 * cm
    )

    styles = getSampleStyleSheet()

    estilo_titulo = ParagraphStyle(
        "TituloPrincipalReporte",
        parent=styles["Title"],
        textColor=colors.HexColor(COLOR_ROJO),
        fontSize=17,
        alignment=1,
        leading=20,
        spaceAfter=6
    )

    estilo_institucional = ParagraphStyle(
        "InstitucionalReporte",
        parent=styles["Normal"],
        fontSize=9,
        alignment=1,
        leading=11,
        spaceAfter=4
    )

    estilo_subtitulo = ParagraphStyle(
        "SubtituloReporte",
        parent=styles["Heading2"],
        textColor=colors.HexColor(COLOR_VERDE),
        fontSize=12,
        leading=14,
        spaceBefore=4,
        spaceAfter=6
    )

    normal = styles["Normal"]

    elementos = []

    elementos.append(
        Paragraph(
            "UNIDAD DE ADMINISTRACIÓN Y FINANZAS<br/>"
            "COORDINACIÓN DE RECURSOS MATERIALES<br/>"
            "COORDINACIÓN TÉCNICA DE ABASTO DE INSUMOS PARA LA SALUD",
            estilo_institucional
        )
    )

    elementos.append(
        Paragraph(
            "REPORTE EJECUTIVO DE INCIDENCIAS",
            estilo_titulo
        )
    )

    hoy = fecha_hoy_sistema()

    elementos.append(
        Paragraph(
            f"Fecha del reporte: {hoy}",
            normal
        )
    )

    elementos.append(
        Spacer(1, 0.25 * cm)
    )

    if incidencias is None or incidencias.empty:

        elementos.append(
            Paragraph(
                "Sin información disponible.",
                normal
            )
        )

        doc.build(
            elementos,
            onFirstPage=agregar_encabezado_reporte,
            onLaterPages=agregar_encabezado_reporte
        )

        buffer.seek(0)

        return buffer

    df = incidencias.copy()

    for columna in [
        "FECHA_REGISTRO",
        "CREADO_EN",
        "RESPONSABLE",
        "ENTIDAD",
        "ORDEN",
        "CLUES_DESTINO",
        "TIPO_INCIDENCIA",
        "ATRIBUIBLE A",
        "ESTATUS_INCIDENCIA",
        "ESTATUS_SEGUIMIENTO"
    ]:

        if columna not in df.columns:

            df[columna] = ""

    df["FECHA_REGISTRO_DT"] = pd.to_datetime(
        df["FECHA_REGISTRO"],
        errors="coerce",
        utc=True
    )

    fechas_creado_en = pd.to_datetime(
        df["CREADO_EN"],
        errors="coerce",
        utc=True
    )

    df["FECHA_REGISTRO_DT"] = df["FECHA_REGISTRO_DT"].fillna(
        fechas_creado_en
    )

    df = df[
        df["FECHA_REGISTRO_DT"].notna()
    ].copy()

    if df.empty:

        elementos.append(
            Paragraph(
                "No existen fechas válidas para generar el reporte.",
                normal
            )
        )

        doc.build(
            elementos,
            onFirstPage=agregar_encabezado_reporte,
            onLaterPages=agregar_encabezado_reporte
        )

        buffer.seek(0)

        return buffer

    df["FECHA_REGISTRO_MX"] = (
        df["FECHA_REGISTRO_DT"]
        .dt.tz_convert("America/Mexico_City")
    )

    df["FECHA_DIA"] = df["FECHA_REGISTRO_MX"].dt.strftime(
        "%Y-%m-%d"
    )

    df["FECHA_DIA_TEXTO"] = df["FECHA_REGISTRO_MX"].dt.strftime(
        "%d/%m/%Y"
    )

    df["MES_REPORTE"] = df["FECHA_REGISTRO_MX"].dt.strftime(
        "%Y-%m"
    )

    df["MONITORA"] = df["RESPONSABLE"].apply(
        extraer_monitora
    )

    df["MONITORA"] = df["MONITORA"].replace(
        "",
        "SIN RESPONSABLE"
    )

    hoy_dt = pd.to_datetime(
        hoy
    )

    mes_actual = hoy_dt.strftime(
        "%Y-%m"
    )

    df_mes = df[
        df["MES_REPORTE"] == mes_actual
    ].copy()

    df_hoy = df[
        df["FECHA_DIA"] == hoy
    ].copy()

    total_hoy = len(df_hoy)

    monitoras_hoy = df_hoy["MONITORA"].replace(
        "SIN RESPONSABLE",
        pd.NA
    ).dropna().nunique()

    entidades_hoy = df_hoy["ENTIDAD"].replace(
        "",
        pd.NA
    ).dropna().nunique()

    ordenes_hoy = df_hoy["ORDEN"].replace(
        "",
        pd.NA
    ).dropna().nunique()

    elementos.append(
        Paragraph(
            "1. Resumen del día",
            estilo_subtitulo
        )
    )

    indicadores_dia = pd.DataFrame(
        [
            ["Incidencias capturadas hoy", total_hoy],
            ["Monitoras que capturaron hoy", monitoras_hoy],
            ["Entidades atendidas hoy", entidades_hoy],
            ["Órdenes registradas hoy", ordenes_hoy],
        ],
        columns=[
            "Indicador",
            "Total"
        ]
    )

    elementos += tabla_pdf(
        "Indicadores del día",
        indicadores_dia,
        max_filas=10,
        anchos=[
            9 * cm,
            4 * cm
        ]
    )

    if not df_hoy.empty:

        por_monitora_hoy = (
            df_hoy
            .groupby("MONITORA", dropna=False)
            .size()
            .reset_index(name="Capturas")
            .sort_values("Capturas", ascending=False)
        )

        elementos += tabla_pdf(
            "Capturas del día por responsable",
            por_monitora_hoy,
            max_filas=15,
            anchos=[
                12 * cm,
                4 * cm
            ]
        )

    else:

        elementos.append(
            Paragraph(
                "No se registraron capturas durante el día del reporte.",
                normal
            )
        )

    elementos.append(
        PageBreak()
    )

    elementos.append(
        Paragraph(
            "2. Incidencias cargadas durante el mes",
            estilo_subtitulo
        )
    )

    if df_mes.empty:

        elementos.append(
            Paragraph(
                "No existen incidencias registradas en el mes actual.",
                normal
            )
        )

    else:

        total_mes = len(df_mes)

        dias_con_captura = df_mes["FECHA_DIA"].nunique()

        promedio_diario = round(
            total_mes / max(dias_con_captura, 1),
            2
        )

        por_dia_base = (
            df_mes
            .groupby("FECHA_DIA")
            .size()
            .reset_index(name="Incidencias")
        )

        rango_dias = pd.date_range(
            start=hoy_dt.replace(day=1),
            end=hoy_dt,
            freq="D"
        )

        calendario_mes = pd.DataFrame(
            {
                "FECHA_DIA": rango_dias.strftime("%Y-%m-%d")
            }
        )

        por_dia_mes = calendario_mes.merge(
            por_dia_base,
            on="FECHA_DIA",
            how="left"
        )

        por_dia_mes["Incidencias"] = por_dia_mes["Incidencias"].fillna(
            0
        ).astype(int)

        por_dia_mes["Fecha"] = pd.to_datetime(
            por_dia_mes["FECHA_DIA"]
        ).dt.strftime("%d/%m/%Y")

        por_dia_mes = por_dia_mes[
            [
                "Fecha",
                "Incidencias"
            ]
        ]

        dia_mayor = por_dia_mes.sort_values(
            "Incidencias",
            ascending=False
        ).iloc[0]

        por_monitora_mes = (
            df_mes
            .groupby("MONITORA", dropna=False)
            .size()
            .reset_index(name="Incidencias")
            .sort_values("Incidencias", ascending=False)
        )

        monitora_mayor = por_monitora_mes.iloc[0]

        indicadores_mes = pd.DataFrame(
            [
                ["Total de incidencias del mes", total_mes],
                ["Días con captura en el mes", dias_con_captura],
                ["Promedio diario con captura", promedio_diario],
                ["Día con mayor carga", f"{dia_mayor['Fecha']} - {dia_mayor['Incidencias']} incidencias"],
                ["Responsable con más capturas", f"{monitora_mayor['MONITORA']} - {monitora_mayor['Incidencias']} incidencias"],
            ],
            columns=[
                "Indicador",
                "Resultado"
            ]
        )

        elementos += tabla_pdf(
            "Indicadores mensuales",
            indicadores_mes,
            max_filas=10,
            anchos=[
                8 * cm,
                12 * cm
            ]
        )

        elementos += tabla_pdf(
            "Incidencias cargadas por día durante el mes",
            por_dia_mes,
            max_filas=31,
            anchos=[
                6 * cm,
                4 * cm
            ]
        )

        elementos.append(
            PageBreak()
        )

        elementos.append(
            Paragraph(
                "3. Quién carga las incidencias",
                estilo_subtitulo
            )
        )

        por_monitora_mes.insert(
            0,
            "Lugar",
            range(1, len(por_monitora_mes) + 1)
        )

        elementos += tabla_pdf(
            "Ranking mensual por responsable",
            por_monitora_mes,
            max_filas=25,
            anchos=[
                2 * cm,
                13 * cm,
                4 * cm
            ]
        )

        cruce_dia_monitora = pd.pivot_table(
            df_mes,
            index="FECHA_DIA_TEXTO",
            columns="MONITORA",
            values="ORDEN",
            aggfunc="count",
            fill_value=0
        ).reset_index()

        cruce_dia_monitora = cruce_dia_monitora.rename(
            columns={
                "FECHA_DIA_TEXTO": "Fecha"
            }
        )

        cruce_dia_monitora["Total"] = cruce_dia_monitora.drop(
            columns=[
                "Fecha"
            ],
            errors="ignore"
        ).sum(axis=1)

        columnas_cruce = [
            "Fecha"
        ] + [
            c for c in cruce_dia_monitora.columns
            if c not in [
                "Fecha",
                "Total"
            ]
        ] + [
            "Total"
        ]

        cruce_dia_monitora = cruce_dia_monitora[
            columnas_cruce
        ]

        columnas_mostrar = cruce_dia_monitora.columns.tolist()

        if len(columnas_mostrar) > 8:

            top_monitoras = por_monitora_mes["MONITORA"].head(
                6
            ).tolist()

            columnas_mostrar = [
                "Fecha"
            ] + top_monitoras + [
                "Total"
            ]

            cruce_dia_monitora = cruce_dia_monitora[
                [
                    c for c in columnas_mostrar
                    if c in cruce_dia_monitora.columns
                ]
            ]

        ancho_fecha = 3 * cm
        columnas_restantes = max(
            len(cruce_dia_monitora.columns) - 1,
            1
        )

        anchos_cruce = [
            ancho_fecha
        ] + [
            22 * cm / columnas_restantes
        ] * columnas_restantes

        elementos += tabla_pdf(
            "Cruce de capturas por día y responsable",
            cruce_dia_monitora,
            max_filas=31,
            anchos=anchos_cruce
        )

    elementos.append(
        PageBreak()
    )

    elementos.append(
        Paragraph(
            "4. Resumen general acumulado",
            estilo_subtitulo
        )
    )

    total_general = len(df)

    resueltas = df[
        df["ESTATUS_INCIDENCIA"]
        .astype(str)
        .str.upper()
        .isin(
            [
                "RESUELTA",
                "RESUELTO"
            ]
        )
    ].shape[0]

    en_proceso = df[
        df["ESTATUS_INCIDENCIA"]
        .astype(str)
        .str.upper()
        .eq("EN PROCESO")
    ].shape[0]

    rechazadas = df[
        df["ESTATUS_INCIDENCIA"]
        .astype(str)
        .str.upper()
        .eq("RECHAZADO")
    ].shape[0]

    monitoras_general = df["MONITORA"].replace(
        "SIN RESPONSABLE",
        pd.NA
    ).dropna().nunique()

    indicadores_general = pd.DataFrame(
        [
            ["Total acumulado", total_general],
            ["Resueltas", resueltas],
            ["En proceso", en_proceso],
            ["Rechazadas", rechazadas],
            ["Responsables con capturas", monitoras_general],
        ],
        columns=[
            "Indicador",
            "Total"
        ]
    )

    elementos += tabla_pdf(
        "Indicadores generales",
        indicadores_general,
        max_filas=10,
        anchos=[
            9 * cm,
            4 * cm
        ]
    )

    por_tipo = (
        df
        .groupby("TIPO_INCIDENCIA", dropna=False)
        .size()
        .reset_index(name="Total")
        .sort_values("Total", ascending=False)
    )

    elementos += tabla_pdf(
        "Principales tipos de incidencia",
        por_tipo,
        max_filas=10,
        anchos=[
            14 * cm,
            4 * cm
        ]
    )

    por_entidad = (
        df
        .groupby("ENTIDAD", dropna=False)
        .size()
        .reset_index(name="Total")
        .sort_values("Total", ascending=False)
    )

    elementos += tabla_pdf(
        "Entidades con más incidencias",
        por_entidad,
        max_filas=10,
        anchos=[
            10 * cm,
            4 * cm
        ]
    )

    if not df_hoy.empty:

        elementos.append(
            PageBreak()
        )

        elementos.append(
            Paragraph(
                "5. Detalle breve del día",
                estilo_subtitulo
            )
        )

        detalle_hoy = df_hoy[
            [
                "ORDEN",
                "ENTIDAD",
                "CLUES_DESTINO",
                "TIPO_INCIDENCIA",
                "ATRIBUIBLE A",
                "ESTATUS_INCIDENCIA",
                "MONITORA"
            ]
        ].copy()

        elementos += tabla_pdf(
            "Detalle de incidencias capturadas hoy",
            detalle_hoy,
            max_filas=25,
            anchos=[
                4.2 * cm,
                3.0 * cm,
                3.0 * cm,
                5.0 * cm,
                3.0 * cm,
                3.0 * cm,
                4.0 * cm
            ]
        )

    doc.build(
        elementos,
        onFirstPage=agregar_encabezado_reporte,
        onLaterPages=agregar_encabezado_reporte
    )

    buffer.seek(0)

    return buffer