import pandas as pd

RUTA_XLSB = r"C:\Users\guillermo.ortega\OneDrive - IMSS-BIENESTAR\Escritorio\Incidencias fase_1\CompendioAbasto25-26_22.05.2026.xlsb"

HOJAS = [
    ("DATA", "ACTIVA"),
    ("CANCELADAS", "INACTIVA")
]

bases = []

for hoja, estatus_base in HOJAS:

    print(f"Leyendo hoja: {hoja}")

    df = pd.read_excel(
        RUTA_XLSB,
        sheet_name=hoja,
        engine="pyxlsb"
    )

    df.columns = (
        df.columns
        .astype(str)
        .str.strip()
    )

    df["ESTATUS_BASE"] = estatus_base
    df["ORIGEN_COMPENDIO"] = hoja

    bases.append(df)

base = pd.concat(
    bases,
    ignore_index=True
)

base = base.astype(str)

base.to_parquet(
    "base_compendio_ligera.parquet",
    index=False
)

print("Parquet generado correctamente.")
print("Filas totales:", len(base))
print("Columnas:", len(base.columns))
print(base["ESTATUS_BASE"].value_counts())