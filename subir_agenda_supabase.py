import os
import pandas as pd
from supabase import create_client


ARCHIVO_EXCEL = "agenda_citas.xlsx"
TABLA_SUPABASE = "agenda_citas"

SUPABASE_URL = "https://zstcpebhdnxtoudampck.supabase.co"
SUPABASE_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InpzdGNwZWJoZG54dG91ZGFtcGNrIiwicm9sZSI6ImFub24iLCJpYXQiOjE3Nzk2MzkxNTIsImV4cCI6MjA5NTIxNTE1Mn0.hSh8EoDM2BD0uQcbK_3oQGrEbqUiRd69NRXhPNy6FGU"

supabase = create_client(
    SUPABASE_URL,
    SUPABASE_KEY
)


def limpiar(valor):

    if pd.isna(valor):
        return None

    texto = str(valor).strip()

    if texto.lower() in ["nan", "none", "null", ""]:
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


def obtener(fila, opciones):

    for col in opciones:
        if col in fila.index:
            valor = limpiar(fila.get(col))
            if valor:
                return valor

    return None


def fecha_texto(valor):

    if pd.isna(valor):
        return None

    fecha = pd.to_datetime(
        valor,
        errors="coerce"
    )

    if pd.isna(fecha):
        return limpiar(valor)

    return fecha.strftime("%d/%m/%Y")


df = pd.read_excel(
    ARCHIVO_EXCEL
)

df.columns = (
    df.columns
    .astype(str)
    .str.strip()
)

print("Filas agenda encontradas:", len(df))

registros = []

for _, fila in df.iterrows():

    orden = obtener(
        fila,
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

    registro = {
        "orden_suministro": normalizar_orden(orden),
        "fecha_cita": fecha_texto(
            obtener(
                fila,
                [
                    "FECHA  DE CITA AGENDA",
                    "FECHA DE CITA AGENDA",
                    "Fecha  de cita agenda",
                    "Fecha de cita agenda",
                    "fecha_de_cita_agenda",
                    "fecha_cita",
                    "FECHA CITA",
                    "Fecha cita",
                    "FECHA",
                    "fecha"
                ]
            )
        ),
        "estatus": obtener(
            fila,
            [
                "ESTATUS",
                "estatus",
                "Estatus"
            ]
        ),
        "entidad": obtener(
            fila,
            [
                "ENTIDAD",
                "entidad",
                "ESTADO",
                "estado"
            ]
        ),
        "clues": obtener(
            fila,
            [
                "CLUES",
                "clues",
                "CLUES DESTINO",
                "clues_destino"
            ]
        ),
        "almacen": obtener(
            fila,
            [
                "ALMACEN",
                "ALMACÉN",
                "almacen",
                "UNIDAD",
                "Unidad",
                "UNIDAD DESTINO"
            ]
        ),
        "proveedor": obtener(
            fila,
            [
                "PROVEEDOR",
                "proveedor"
            ]
        ),
        "observaciones": obtener(
            fila,
            [
                "OBSERVACIONES",
                "observaciones",
                "OBS"
            ]
        )
    }

    if registro["orden_suministro"]:
        registros.append(registro)

print("Registros válidos:", len(registros))

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

print("Agenda cargada correctamente en Supabase.")