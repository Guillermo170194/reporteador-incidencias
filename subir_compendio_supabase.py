import os
import math
import pandas as pd
from supabase import create_client


# =========================
# CONFIGURACIÓN
# =========================

ARCHIVO_CSV = "compendio_supabase.csv"
TABLA_SUPABASE = "compendio"
BLOQUE = 500


# =========================
# SUPABASE
# =========================

SUPABASE_URL = os.environ.get(
    "SUPABASE_URL",
    "https://zstcpebhdnxtoudampck.supabase.co"
)

SUPABASE_KEY = os.environ.get(
    "SUPABASE_KEY",
    ""
)

if not SUPABASE_URL or not SUPABASE_KEY:

    raise Exception(
        "Falta SUPABASE_URL o SUPABASE_KEY. Revisa variables de entorno."
    )

supabase = create_client(
    SUPABASE_URL,
    SUPABASE_KEY
)


# =========================
# FUNCIONES
# =========================

def limpiar_valor(valor):

    if pd.isna(valor):

        return None

    texto = str(valor).strip()

    if texto.lower() in [
        "nan",
        "none",
        "null",
        ""
    ]:

        return None

    return texto


def normalizar_orden(valor):

    if pd.isna(valor):

        return None

    return (
        str(valor)
        .upper()
        .strip()
        .replace(" ", "")
        .replace("_", "-")
        .replace("–", "-")
        .replace("—", "-")
    )


def limpiar_numero(valor):

    if pd.isna(valor):

        return None

    try:

        texto = (
            str(valor)
            .replace(",", "")
            .replace("$", "")
            .strip()
        )

        if texto.lower() in [
            "nan",
            "none",
            "null",
            ""
        ]:

            return None

        return float(texto)

    except Exception:

        return None


def preparar_registros(df):

    registros = []

    for _, fila in df.iterrows():

        orden_suministro = normalizar_orden(
            fila.get(
                "orden_suministro"
            )
        )

        orden = normalizar_orden(
            fila.get(
                "orden"
            )
        )

        no_orden = normalizar_orden(
            fila.get(
                "no_orden"
            )
        )

        registro = {
            "orden": orden,
            "orden_suministro": orden_suministro,
            "no_orden": no_orden,
            "estatus_base": limpiar_valor(
                fila.get("estatus_base")
            ),
            "origen_compendio": limpiar_valor(
                fila.get("origen_compendio")
            ),
            "entidad": limpiar_valor(
                fila.get("entidad")
            ),
            "estado": limpiar_valor(
                fila.get("estado")
            ),
            "clues_destino": limpiar_valor(
                fila.get("clues_destino")
            ),
            "unidad_destino": limpiar_valor(
                fila.get("unidad_destino")
            ),
            "almacen": limpiar_valor(
                fila.get("almacen")
            ),
            "proveedor": limpiar_valor(
                fila.get("proveedor")
            ),
            "clave_cnis": limpiar_valor(
                fila.get("clave_cnis")
            ),
            "descripcion": limpiar_valor(
                fila.get("descripcion")
            ),
            "tipo_entrega": limpiar_valor(
                fila.get("tipo_entrega")
            ),
            "piezas_emitidas": limpiar_numero(
                fila.get("piezas_emitidas")
            ),
            "piezas_recibidas_ol": limpiar_numero(
                fila.get("piezas_recibidas_ol")
            ),
            "piezas_entregadas_clues": limpiar_numero(
                fila.get("piezas_entregadas_clues")
            ),
            "operador_logistico": limpiar_valor(
                fila.get("operador_logistico")
            ),
            "tipo_red": limpiar_valor(
                fila.get("tipo_red")
            ),
            "grupo_terapeutico": limpiar_valor(
                fila.get("grupo_terapeutico")
            )
        }

        registros.append(
            registro
        )

    return registros


# =========================
# LEER CSV
# =========================

print("Leyendo CSV del compendio...")

df = pd.read_csv(
    ARCHIVO_CSV,
    dtype=str,
    low_memory=False
)

df.columns = (
    df.columns
    .astype(str)
    .str.strip()
)

print(
    f"Filas CSV: {len(df)}"
)

print(
    "Columnas:",
    list(df.columns)
)


# =========================
# LIMPIAR TABLA
# =========================

print(
    "Borrando compendio anterior en Supabase..."
)

supabase.table(
    TABLA_SUPABASE
).delete().neq(
    "orden_suministro",
    "__NO_EXISTE__"
).execute()

print(
    "Compendio anterior borrado."
)


# =========================
# PREPARAR Y SUBIR
# =========================

registros = preparar_registros(
    df
)

total = len(
    registros
)

print(
    f"Registros preparados: {total}"
)

for i in range(
    0,
    total,
    BLOQUE
):

    bloque = registros[
        i:i + BLOQUE
    ]

    supabase.table(
        TABLA_SUPABASE
    ).insert(
        bloque
    ).execute()

    print(
        f"Subidas {i + len(bloque)} de {total} filas"
    )

print(
    "Compendio actualizado correctamente en Supabase."
)