import pandas as pd


RUTA = (
    r"C:\Users\guillermo.ortega\OneDrive - IMSS-BIENESTAR"
    r"\Escritorio\Incidencias fase_1\inv 29 de mayo 2026.xlsx"
)

SALIDA = (
    r"C:\Users\guillermo.ortega\OneDrive - IMSS-BIENESTAR"
    r"\Escritorio\Incidencias fase_1\inventario_clues_supabase.csv"
)


df = pd.read_excel(
    RUTA,
    sheet_name="Resultado consulta"
)

df.columns = (
    df.columns
    .astype(str)
    .str.strip()
)


def tomar(opciones):

    for col in opciones:

        if col in df.columns:

            return df[col]

    return pd.Series([""] * len(df))


def limpiar_texto(serie):

    return (
        serie
        .fillna("")
        .astype(str)
        .str.strip()
        .replace(
            {
                "nan": "",
                "None": "",
                "NaT": ""
            }
        )
    )


inventario = pd.DataFrame()


inventario["entidad"] = limpiar_texto(
    tomar(
        [
            "ENTIDAD",
            "Entidad",
            "ESTADO",
            "Estado"
        ]
    )
)

inventario["clues"] = limpiar_texto(
    tomar(
        [
            "CLUES",
            "Clues",
            "CLUES DESTINO",
            "clues"
        ]
    )
)

inventario["unidad"] = limpiar_texto(
    tomar(
        [
            "UNIDAD",
            "Unidad",
            "NOMBRE UNIDAD",
            "Nombre Unidad"
        ]
    )
)

inventario["clave_cnis"] = limpiar_texto(
    tomar(
        [
            "CLAVE",
            "Clave",
            "CLAVE CNIS",
            "Clave CNIS",
            "clave_cnis"
        ]
    )
)

inventario["descripcion"] = limpiar_texto(
    tomar(
        [
            "DESCRIPCIÓN",
            "DESCRIPCION",
            "Descripción",
            "Descripcion"
        ]
    )
)


# =========================
# PIEZAS NUMÉRICO
# =========================

piezas = pd.to_numeric(
    limpiar_texto(
        tomar(
            [
                "PIEZAS",
                "Piezas",
                "piezas",
                "EXISTENCIA",
                "Existencia"
            ]
        )
    )
    .str.replace(",", "", regex=False),
    errors="coerce"
)

inventario["piezas"] = (
    piezas
    .fillna(0)
    .round(0)
    .astype("int64")
)


inventario["lote"] = limpiar_texto(
    tomar(
        [
            "LOTE",
            "Lote",
            "lote"
        ]
    )
)

inventario["estatus"] = limpiar_texto(
    tomar(
        [
            "ESTATUS",
            "Estatus",
            "estatus"
        ]
    )
)


# =========================
# FECHA CADUCIDAD
# =========================

fecha = tomar(
    [
        "CADUCIDAD",
        "Caducidad",
        "FECHA CADUCIDAD",
        "Fecha Caducidad"
    ]
)

fecha = pd.to_datetime(
    fecha,
    errors="coerce",
    dayfirst=True
)

inventario["caducidad"] = fecha.dt.strftime(
    "%Y-%m-%d"
).fillna("")


# =========================
# LIMPIEZA FINAL
# =========================

inventario = inventario.fillna("")

inventario = inventario[
    inventario["clave_cnis"]
    .astype(str)
    .str.strip() != ""
].copy()


# =========================
# EXPORTAR CSV
# =========================

inventario.to_csv(
    SALIDA,
    index=False,
    encoding="utf-8-sig"
)

print("\nCSV inventario generado correctamente")
print("Filas:", len(inventario))
print("Columnas:", inventario.columns.tolist())

print("\nPrimeras filas:")
print(inventario.head())