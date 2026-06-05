import os
import math
import pandas as pd
from datetime import datetime
from supabase import create_client


# =========================
# CONFIGURACIÓN
# =========================

ARCHIVO_EXCEL = "incidencias.xlsx"
TABLA_SUPABASE = "incidencias"


# =========================
# SUPABASE
# =========================

SUPABASE_URL = "https://zstcpebhdnxtoudampck.supabase.co"
SUPABASE_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InpzdGNwZWJoZG54dG91ZGFtcGNrIiwicm9sZSI6ImFub24iLCJpYXQiOjE3Nzk2MzkxNTIsImV4cCI6MjA5NTIxNTE1Mn0.hSh8EoDM2BD0uQcbK_3oQGrEbqUiRd69NRXhPNy6FGU"

if not SUPABASE_URL or not SUPABASE_KEY:
    raise Exception(
        "Faltan SUPABASE_URL o SUPABASE_KEY en variables de entorno."
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

    if texto.lower() in ["nan", "none", "null", ""]:
        return None

    return texto


def obtener(fila, opciones):

    for col in opciones:
        if col in fila.index:
            valor = limpiar_valor(fila.get(col))
            if valor:
                return valor

    return None


def limpiar_fecha(valor):

    if pd.isna(valor):
        return None

    try:
        fecha = pd.to_datetime(
            valor,
            errors="coerce"
        )

        if pd.isna(fecha):
            return None

        return fecha.isoformat()

    except Exception:
        return None


# =========================
# LEER EXCEL
# =========================

df = pd.read_excel(
    ARCHIVO_EXCEL
)

df.columns = (
    df.columns
    .astype(str)
    .str.strip()
)

print(f"Filas encontradas: {len(df)}")


# =========================
# ARMAR REGISTROS
# =========================

registros = []

for _, fila in df.iterrows():

    registro = {
        "fecha_registro": limpiar_fecha(
            obtener(
                fila,
                [
                    "FECHA_REGISTRO",
                    "Fecha Registro",
                    "fecha_registro"
                ]
            )
        ),
        "origen_registro": obtener(
            fila,
            [
                "ORIGEN_REGISTRO",
                "origen_registro"
            ]
        ),
        "orden_buscada": obtener(
            fila,
            [
                "ORDEN_BUSCADA",
                "orden_buscada"
            ]
        ),
        "orden_suministro": obtener(
            fila,
            [
                "orden_suministro",
                "ORDEN",
                "Orden de Suministro",
                "ORDEN DE SUMINISTRO"
            ]
        ),
        "tipo_entrega": obtener(
            fila,
            [
                "TIPO_ENTREGA",
                "tipo_entrega"
            ]
        ),
        "entidad": obtener(
            fila,
            [
                "ENTIDAD",
                "entidad",
                "Estado",
                "ESTADO"
            ]
        ),
        "almacen_clues_destino": obtener(
            fila,
            [
                "ALMACEN_CLUES_DESTINO",
                "almacen_clues_destino"
            ]
        ),
        "clues_destino": obtener(
            fila,
            [
                "CLUES_DESTINO",
                "clues_destino"
            ]
        ),
        "unidad_destino": obtener(
            fila,
            [
                "UNIDAD_DESTINO",
                "unidad_destino"
            ]
        ),
        "proveedor": obtener(
            fila,
            [
                "PROVEEDOR",
                "proveedor"
            ]
        ),
        "orden": obtener(
            fila,
            [
                "ORDEN",
                "orden",
                "orden_suministro"
            ]
        ),
        "clave_cnis": obtener(
            fila,
            [
                "CLAVE_CNIS",
                "clave_cnis",
                "Clave CNIS"
            ]
        ),
        "descripcion": obtener(
            fila,
            [
                "DESCRIPCION",
                "descripcion",
                "Descripción"
            ]
        ),
        "piezas_emitidas": obtener(
            fila,
            [
                "PIEZAS_EMITIDAS",
                "piezas_emitidas"
            ]
        ),
        "piezas_recibidas_ol": obtener(
            fila,
            [
                "PIEZAS_RECIBIDAS_OL",
                "piezas_recibidas_ol"
            ]
        ),
        "piezas_entregadas_clues": obtener(
            fila,
            [
                "PIEZAS_ENTREGADAS_CLUES",
                "piezas_entregadas_clues"
            ]
        ),
        "tipo_red": obtener(
            fila,
            [
                "TIPO_RED",
                "tipo_red"
            ]
        ),
        "grupo_terapeutico": obtener(
            fila,
            [
                "GRUPO_TERAPEUTICO",
                "grupo_terapeutico",
                "GRUPO TERAPEUTICO",
                "GRUPO TERAPÉUTICO"
            ]
        ),
        "estatus_operativo": obtener(
            fila,
            [
                "ESTATUS_OPERATIVO",
                "estatus_operativo"
            ]
        ),
        "estatus_base": obtener(
            fila,
            [
                "ESTATUS_BASE",
                "estatus_base"
            ]
        ),
        "origen_compendio": obtener(
            fila,
            [
                "ORIGEN_COMPENDIO",
                "origen_compendio"
            ]
        ),
        "operador_logistico": obtener(
            fila,
            [
                "OPERADOR_LOGISTICO",
                "operador_logistico"
            ]
        ),
        "estatus_recepcion_ol": obtener(
            fila,
            [
                "ESTATUS_RECEPCION_OL",
                "estatus_recepcion_ol"
            ]
        ),
        "estatus_entrega_estado": obtener(
            fila,
            [
                "ESTATUS_ENTREGA_ESTADO",
                "estatus_entrega_estado"
            ]
        ),
        "estatus_incidencia_completa": obtener(
            fila,
            [
                "ESTATUS_INCIDENCIA_COMPLETA",
                "estatus_incidencia_completa"
            ]
        ),
        "tipo_incidencia": obtener(
            fila,
            [
                "TIPO_INCIDENCIA",
                "tipo_incidencia"
            ]
        ),
        "atribuible_a": obtener(
            fila,
            [
                "ATRIBUIBLE A",
                "ATRIBUIBLE_A",
                "atribuible_a"
            ]
        ),
        "estatus_incidencia": obtener(
            fila,
            [
                "ESTATUS_INCIDENCIA",
                "estatus_incidencia"
            ]
        ),
        "responsable": obtener(
            fila,
            [
                "RESPONSABLE",
                "responsable"
            ]
        ),
        "observaciones": obtener(
            fila,
            [
                "OBSERVACIONES",
                "observaciones"
            ]
        ),
        "pdf_cedula_rechazo": obtener(
            fila,
            [
                "PDF_CEDULA_RECHAZO",
                "pdf_cedula_rechazo"
            ]
        ),
        "pdf_correo_seguimiento": obtener(
            fila,
            [
                "PDF_CORREO_SEGUIMIENTO",
                "pdf_correo_seguimiento"
            ]
        )
    }

    registros.append(registro)


# =========================
# SUBIR EN BLOQUES
# =========================

BLOQUE = 500

for i in range(0, len(registros), BLOQUE):

    bloque = registros[i:i + BLOQUE]

    supabase.table(
        TABLA_SUPABASE
    ).insert(
        bloque
    ).execute()

    print(
        f"Subidas {i + len(bloque)} de {len(registros)} filas"
    )

print("Carga histórica terminada correctamente.")