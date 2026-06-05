import pandas as pd


RUTA = (
    r"C:\Users\guillermo.ortega\OneDrive - IMSS-BIENESTAR"
    r"\Escritorio\Incidencias fase_1\agenda_citas.xlsx"
)

SALIDA = (
    r"C:\Users\guillermo.ortega\OneDrive - IMSS-BIENESTAR"
    r"\Escritorio\Incidencias fase_1\agenda_citas_supabase.csv"
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
    opciones
):

    for col in opciones:

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


agenda = pd.DataFrame()


agenda["orden_suministro"] = limpiar_texto(
    tomar(
        [
            "orden_suministro",
            "ORDEN_SUMINISTRO",
            "ORDEN DE SUMINISTRO",
            "Orden de Suministro",
            "ORDEN",
            "orden",
            "NO. ORDEN",
            "no_orden"
        ]
    )
)


fecha = tomar(
    [
        "fecha_cita",
        "FECHA CITA",
        "Fecha cita",
        "FECHA  DE CITA AGENDA",
        "FECHA DE CITA AGENDA",
        "Fecha  de cita agenda",
        "Fecha de cita agenda",
        "fecha_de_cita_agenda",
        "FECHA",
        "Fecha"
    ]
)

fecha = pd.to_datetime(
    fecha,
    errors="coerce",
    dayfirst=True
)

agenda["fecha_cita"] = fecha.dt.strftime(
    "%Y-%m-%d"
)


agenda["estatus"] = limpiar_texto(
    tomar(
        [
            "estatus",
            "ESTATUS",
            "Estatus"
        ]
    )
)

agenda["entidad"] = limpiar_texto(
    tomar(
        [
            "entidad",
            "ENTIDAD",
            "Estado",
            "ESTADO"
        ]
    )
)

agenda["clues"] = limpiar_texto(
    tomar(
        [
            "clues",
            "CLUES",
            "CLUES DESTINO",
            "clues_destino",
            "CLUES_DESTINO"
        ]
    )
)

agenda["unidad"] = limpiar_texto(
    tomar(
        [
            "unidad",
            "UNIDAD",
            "UNIDAD DESTINO",
            "unidad_destino",
            "UNIDAD_DESTINO"
        ]
    )
)


agenda = agenda.fillna(
    ""
)

agenda = agenda[
    agenda["orden_suministro"]
    .astype(str)
    .str.strip()
    != ""
].copy()

agenda.to_csv(
    SALIDA,
    index=False,
    encoding="utf-8-sig"
)

print(
    "CSV agenda generado correctamente"
)

print(
    "Filas:",
    len(agenda)
)

print(
    "Columnas:",
    agenda.columns.tolist()
)

print(
    agenda.head()
)