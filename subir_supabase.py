import pandas as pd

RUTA = r"C:\Users\guillermo.ortega\OneDrive - IMSS-BIENESTAR\Escritorio\Incidencias fase_1\base_compendio_ligera.parquet"

df = pd.read_parquet(
    RUTA
)

df.columns = (
    df.columns
    .astype(str)
    .str.strip()
)

renombres = {
    "Orden de Suministro": "orden_suministro",
    "Número de orden de suministro": "no_orden",
    "ESTATUS_BASE": "estatus_base",
    "ORIGEN_COMPENDIO": "origen_compendio",
    "Entidad": "entidad",
    "Entidad de destino": "estado",
    "CLUES Destino": "clues_destino",
    "CLUES de destino": "clues_destino",
    "Unidad Destino": "unidad_destino",
    "Nombre de la unidad": "unidad_destino",
    "Almacén de entrega": "almacen",
    "Proveedor": "proveedor",
    "Razón social": "proveedor",
    "Clave CNIS": "clave_cnis",
    "Clave del medicamento": "clave_cnis",
    "Descripción": "descripcion",
    "Medicamento": "descripcion",
    "Tipo de entrega": "tipo_entrega",
    "No. de pzas. Emitidas": "piezas_emitidas",
    "Cantidad solicitada": "piezas_emitidas",
    "Pzas. Recibidas por O.L.": "piezas_recibidas_ol",
    "Piezas Reportadas como entregadas CLUES Destino": "piezas_entregadas_clues",
    "Tipo de Red": "tipo_red",
    "Grupo Terapéutico": "grupo_terapeutico",
    "Estatus de la orden de suministro": "estatus",
    "Descripción del estatus de la orden de suministro": "estatus",
}

df = df.rename(
    columns=renombres
)

# Si hay columnas duplicadas por venir de DATA + CANCELADAS,
# conserva la primera con información útil.
df = df.loc[
    :,
    ~df.columns.duplicated()
].copy()

columnas_finales = [
    "orden_suministro",
    "no_orden",
    "estatus_base",
    "origen_compendio",
    "entidad",
    "estado",
    "clues_destino",
    "unidad_destino",
    "almacen",
    "proveedor",
    "clave_cnis",
    "descripcion",
    "tipo_entrega",
    "piezas_emitidas",
    "piezas_recibidas_ol",
    "piezas_entregadas_clues",
    "tipo_red",
    "grupo_terapeutico",
    "estatus"
]

for col in columnas_finales:

    if col not in df.columns:

        df[col] = ""

df["orden"] = df["orden_suministro"]

df = df[
    [
        "orden",
        "orden_suministro",
        "no_orden",
        "estatus_base",
        "origen_compendio",
        "entidad",
        "estado",
        "clues_destino",
        "unidad_destino",
        "almacen",
        "proveedor",
        "clave_cnis",
        "descripcion",
        "tipo_entrega",
        "piezas_emitidas",
        "piezas_recibidas_ol",
        "piezas_entregadas_clues",
        "tipo_red",
        "grupo_terapeutico",
        "estatus"
    ]
].copy()

df.to_csv(
    "compendio_supabase.csv",
    index=False,
    encoding="utf-8-sig"
)

print("CSV generado correctamente")
print("Filas:", len(df))
print("Columnas:", list(df.columns))
print(df.head(5))