import pandas as pd


RUTA = (
    r"C:\Users\guillermo.ortega\OneDrive - IMSS-BIENESTAR"
    r"\Escritorio\Incidencias fase_1\INCIDENCIAS 2026.xlsx"
)

SALIDA = (
    r"C:\Users\guillermo.ortega\OneDrive - IMSS-BIENESTAR"
    r"\Escritorio\Incidencias fase_1\incidencias_supabase.csv"
)


df = pd.read_excel(
    RUTA
)

df.columns = (
    df.columns
    .astype(str)
    .str.strip()
)


def tomar(
    col_opciones
):

    for col in col_opciones:

        if col in df.columns:

            return df[col]

    return ""


def limpiar_texto(
    serie
):

    if isinstance(
        serie,
        str
    ):

        return serie

    return (
        serie
        .fillna("")
        .astype(str)
        .replace(
            {
                "nan": "",
                "None": "",
                "NaT": ""
            }
        )
    )


nuevo = pd.DataFrame()


fecha = tomar(
    [
        "FECHA_REGISTRO",
        "FECHA",
        "Fecha",
        "fecha"
    ]
)

fecha = pd.to_datetime(
    fecha,
    errors="coerce",
    dayfirst=True
)

nuevo["fecha_registro"] = fecha.dt.strftime(
    "%Y-%m-%d"
)


nuevo["origen_registro"] = "CARGA EXCEL"

nuevo["orden_buscada"] = limpiar_texto(
    tomar(
        [
            "ORDEN_BUSCADA",
            "ORDEN",
            "Orden de Suministro",
            "ORDEN DE SUMINISTRO"
        ]
    )
)

nuevo["orden_suministro"] = limpiar_texto(
    tomar(
        [
            "orden_suministro",
            "ORDEN",
            "Orden de Suministro",
            "ORDEN DE SUMINISTRO"
        ]
    )
)

nuevo["tipo_entrega"] = limpiar_texto(
    tomar(
        [
            "TIPO_ENTREGA",
            "TIPO DE ENTREGA",
            "Tipo de entrega"
        ]
    )
)

nuevo["entidad"] = limpiar_texto(
    tomar(
        [
            "ENTIDAD",
            "Estado",
            "ESTADO"
        ]
    )
)

nuevo["almacen_clues_destino"] = limpiar_texto(
    tomar(
        [
            "ALMACEN_CLUES_DESTINO",
            "LUGAR DE ENTREGA",
            "ALMACÉN",
            "ALMACEN"
        ]
    )
)

nuevo["clues_destino"] = limpiar_texto(
    tomar(
        [
            "CLUES_DESTINO",
            "CLUES",
            "CLUES DESTINO"
        ]
    )
)

nuevo["unidad_destino"] = limpiar_texto(
    tomar(
        [
            "UNIDAD_DESTINO",
            "UNIDAD DESTINO",
            "UNIDAD"
        ]
    )
)

nuevo["proveedor"] = limpiar_texto(
    tomar(
        [
            "PROVEEDOR",
            "Proveedor"
        ]
    )
)

nuevo["orden"] = limpiar_texto(
    tomar(
        [
            "ORDEN",
            "orden",
            "orden_suministro",
            "Orden de Suministro",
            "ORDEN DE SUMINISTRO"
        ]
    )
)

nuevo["clave_cnis"] = limpiar_texto(
    tomar(
        [
            "CLAVE_CNIS",
            "CLAVE CNIS",
            "Clave CNIS"
        ]
    )
)

nuevo["descripcion"] = limpiar_texto(
    tomar(
        [
            "DESCRIPCION",
            "DESCRIPCIÓN",
            "Descripción"
        ]
    )
)

nuevo["piezas_emitidas"] = limpiar_texto(
    tomar(
        [
            "PIEZAS_EMITIDAS",
            "NO. DE PZAS. EMITIDAS"
        ]
    )
)

nuevo["piezas_recibidas_ol"] = limpiar_texto(
    tomar(
        [
            "PIEZAS_RECIBIDAS_OL",
            "PZAS. RECIBIDAS POR O.L."
        ]
    )
)

nuevo["piezas_entregadas_clues"] = limpiar_texto(
    tomar(
        [
            "PIEZAS_ENTREGADAS_CLUES",
            "PIEZAS REPORTADAS COMO ENTREGADAS CLUES DESTINO"
        ]
    )
)

nuevo["tipo_red"] = limpiar_texto(
    tomar(
        [
            "TIPO_RED",
            "TIPO DE RED"
        ]
    )
)

nuevo["grupo_terapeutico"] = limpiar_texto(
    tomar(
        [
            "GRUPO_TERAPEUTICO",
            "GRUPO TERAPÉUTICO",
            "GPO TER"
        ]
    )
)

nuevo["estatus_operativo"] = limpiar_texto(
    tomar(
        [
            "ESTATUS_OPERATIVO",
            "ESTATUS"
        ]
    )
)

nuevo["estatus_base"] = limpiar_texto(
    tomar(
        [
            "ESTATUS_BASE"
        ]
    )
)

nuevo["origen_compendio"] = limpiar_texto(
    tomar(
        [
            "ORIGEN_COMPENDIO"
        ]
    )
)

nuevo["operador_logistico"] = limpiar_texto(
    tomar(
        [
            "OPERADOR_LOGISTICO",
            "OPERADOR LOGÍSTICO"
        ]
    )
)

nuevo["estatus_recepcion_ol"] = limpiar_texto(
    tomar(
        [
            "ESTATUS_RECEPCION_OL"
        ]
    )
)

nuevo["estatus_entrega_estado"] = limpiar_texto(
    tomar(
        [
            "ESTATUS_ENTREGA_ESTADO"
        ]
    )
)

nuevo["estatus_incidencia_completa"] = limpiar_texto(
    tomar(
        [
            "ESTATUS_INCIDENCIA_COMPLETA"
        ]
    )
)

nuevo["tipo_incidencia"] = limpiar_texto(
    tomar(
        [
            "TIPO_INCIDENCIA",
            "TIPO DE INCIDENCIA"
        ]
    )
)

nuevo["atribuible_a"] = limpiar_texto(
    tomar(
        [
            "ATRIBUIBLE A",
            "ATRIBUIBLE_A",
            "ATRIBUIBLE"
        ]
    )
)

nuevo["estatus_incidencia"] = limpiar_texto(
    tomar(
        [
            "ESTATUS_INCIDENCIA",
            "ESTATUS INCIDENCIA"
        ]
    )
)

nuevo["responsable"] = limpiar_texto(
    tomar(
        [
            "RESPONSABLE"
        ]
    )
)

nuevo["observaciones"] = limpiar_texto(
    tomar(
        [
            "OBSERVACIONES"
        ]
    )
)

nuevo["pdf_cedula_rechazo"] = limpiar_texto(
    tomar(
        [
            "PDF_CEDULA_RECHAZO"
        ]
    )
)

nuevo["pdf_correo_seguimiento"] = limpiar_texto(
    tomar(
        [
            "PDF_CORREO_SEGUIMIENTO"
        ]
    )
)


nuevo = nuevo.fillna(
    ""
)

nuevo.to_csv(
    SALIDA,
    index=False,
    encoding="utf-8-sig"
)

print(
    "CSV incidencias generado correctamente"
)

print(
    "Filas:",
    len(nuevo)
)

print(
    "Columnas:",
    nuevo.columns.tolist()
)

print(
    nuevo.head()
)