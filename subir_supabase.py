import pandas as pd

RUTA = r"C:\Users\guillermo.ortega\OneDrive - IMSS-BIENESTAR\Escritorio\Incidencias fase_1\base_compendio_ligera.parquet"

df = pd.read_parquet(RUTA)

columnas = [
    "ORDEN",
    "ORDEN DE SUMINISTRO",
    "NO. ORDEN",
    "ESTATUS_BASE",
    "ORIGEN_COMPENDIO",
    "ENTIDAD",
    "ESTADO",
    "CLUES DESTINO",
    "UNIDAD DESTINO",
    "ALMACÉN",
    "PROVEEDOR",
    "CLAVE CNIS",
    "DESCRIPCIÓN",
    "TIPO DE ENTREGA",
    "NO. DE PZAS. EMITIDAS",
    "PZAS. RECIBIDAS POR O.L.",
    "PIEZAS REPORTADAS COMO ENTREGADAS CLUES DESTINO",
    "OPERADOR LOGÍSTICO",
    "TIPO DE RED",
    "GRUPO TERAPÉUTICO"
]

columnas_existentes = [
    c for c in columnas
    if c in df.columns
]

df = df[columnas_existentes].copy()

renombres = {
    "ORDEN": "orden",
    "ORDEN DE SUMINISTRO": "orden_suministro",
    "NO. ORDEN": "no_orden",
    "ESTATUS_BASE": "estatus_base",
    "ORIGEN_COMPENDIO": "origen_compendio",
    "ENTIDAD": "entidad",
    "ESTADO": "estado",
    "CLUES DESTINO": "clues_destino",
    "UNIDAD DESTINO": "unidad_destino",
    "ALMACÉN": "almacen",
    "PROVEEDOR": "proveedor",
    "CLAVE CNIS": "clave_cnis",
    "DESCRIPCIÓN": "descripcion",
    "TIPO DE ENTREGA": "tipo_entrega",
    "NO. DE PZAS. EMITIDAS": "piezas_emitidas",
    "PZAS. RECIBIDAS POR O.L.": "piezas_recibidas_ol",
    "PIEZAS REPORTADAS COMO ENTREGADAS CLUES DESTINO": "piezas_entregadas_clues",
    "OPERADOR LOGÍSTICO": "operador_logistico",
    "TIPO DE RED": "tipo_red",
    "GRUPO TERAPÉUTICO": "grupo_terapeutico"
}

df = df.rename(columns=renombres)

df.to_csv(
    "compendio_supabase.csv",
    index=False
)

print("CSV generado correctamente")