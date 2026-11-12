import streamlit as st
import pandas as pd
import numpy as np
import altair as alt
import re
import os
import json
import tempfile
import unicodedata
import uuid
from collections.abc import Mapping

from dotenv import load_dotenv

load_dotenv()

from io import BytesIO
from datetime import datetime
from zoneinfo import ZoneInfo
from supabase import create_client
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

from reportlab.lib import colors
from reportlab.lib.pagesizes import letter, landscape
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
from reportlab.lib.units import cm


# =========================
# CONFIG STREAMLIT
# =========================

st.set_page_config(
    page_title="Reporteador de Incidencias",
    layout="wide"
)


# =========================
# IDS DRIVE
# =========================

FOLDER_ID_EVIDENCIAS = (
    "1Fbxzc1SC-c5yaLh7z1h4qIG8sY5W78B0"
)

FOLDER_ID_INCIDENCIAS = (
    "1bMw-Un3KZHQds0zsRZAKSXuAggSuL4PG"
)

FOLDER_ID_BASES = (
    "1J1hHDZDTt8CMVBJ6TW8uPd_EU6laYbhG"
)

# Carpeta principal nueva de Drive para incidencias.
# Dentro de esta carpeta el sistema crea:
# - EVIDENCIAS_PDF / ESTADO / CLUES
# - RESPALDOS_EXCEL
FOLDER_ID_INCIDENCIAS_DRIVE = (
    "1zbxeS-iaKrZdlXn2Ua_xxcMf_ppC4SJ8"
)

NOMBRE_CARPETA_EVIDENCIAS = (
    "EVIDENCIAS_PDF"
)

NOMBRE_CARPETA_RESPALDOS = (
    "RESPALDOS_EXCEL"
)

NOMBRE_GOOGLE_SHEET_INCIDENCIAS = (
    "BASE_INCIDENCIAS_SUPABASE"
)

# Libro maestro independiente para el módulo de urgencias.
# Se crea dentro de la misma carpeta de Drive de incidencias.
NOMBRE_GOOGLE_SHEET_URGENCIAS = (
    "BASE_URGENCIAS"
)

NOMBRE_HOJA_URGENCIAS_ORDENES = (
    "Ordenes_Urgencias"
)

NOMBRE_HOJA_URGENCIAS_SIN_ORDEN = (
    "Sin_Orden_Urgencias"
)


# =========================
# ARCHIVOS DRIVE
# =========================

ARCHIVO_INCIDENCIAS = (
    "incidencias.xlsx"
)

# =========================
# RUTAS TEMPORALES
# =========================

TEMP_DIR = tempfile.gettempdir()

RUTA_INCIDENCIAS = os.path.join(
    TEMP_DIR,
    ARCHIVO_INCIDENCIAS
)

# =========================
# GOOGLE DRIVE
# =========================

SCOPES = [
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/spreadsheets"
]


@st.cache_resource
def obtener_google_credentials():

    credenciales_env = os.getenv(
        "GOOGLE_CREDENTIALS",
        ""
    )

    if credenciales_env:

        try:

            credenciales = json.loads(
                credenciales_env
            )

            credentials = (
                service_account.Credentials
                .from_service_account_info(
                    credenciales,
                    scopes=SCOPES
                )
            )

            return credentials

        except Exception as e:

            st.error(
                "GOOGLE_CREDENTIALS no tiene formato JSON válido. Revisa la configuración del servidor."
            )

            st.stop()

    ruta_json = os.getenv(
        "GOOGLE_CREDENTIALS_FILE",
        ""
    )

    if not ruta_json:

        ruta_json = os.path.join(
            os.path.dirname(__file__),
            "credenciales_google.json"
        )

    if not os.path.exists(
        ruta_json
    ):

        st.error(
            "No encontré credenciales de Google. En Render define GOOGLE_CREDENTIALS "
            "como variable de entorno. En local usa credenciales_google.json o GOOGLE_CREDENTIALS_FILE."
        )

        st.stop()

    credentials = (
        service_account.Credentials
        .from_service_account_file(
            ruta_json,
            scopes=SCOPES
        )
    )

    return credentials


@st.cache_resource
def obtener_drive_service():

    credentials = obtener_google_credentials()

    service = build(
        "drive",
        "v3",
        credentials=credentials
    )

    return service


@st.cache_resource
def obtener_sheets_service():

    credentials = obtener_google_credentials()

    service = build(
        "sheets",
        "v4",
        credentials=credentials
    )

    return service


drive_service = obtener_drive_service()
sheets_service = obtener_sheets_service()


# =========================
# SUPABASE
# =========================

@st.cache_resource
def obtener_supabase():

    url = os.environ[
        "SUPABASE_URL"
    ]

    key = os.environ[
        "SUPABASE_KEY"
    ]

    return create_client(
        url,
        key
    )


supabase = obtener_supabase()


# =========================
# DRIVE HELPERS
# =========================

def buscar_archivo_drive(
    nombre_archivo,
    folder_id
):

    resultado = (
        drive_service.files()
        .list(
            q=(
                f"'{folder_id}' in parents "
                f"and trashed = false"
            ),
            fields="files(id, name, parents, mimeType)",
            corpora="allDrives",
            supportsAllDrives=True,
            includeItemsFromAllDrives=True
        )
        .execute()
    )

    archivos = resultado.get(
        "files",
        []
    )

    nombre_busqueda = (
        nombre_archivo
        .lower()
        .strip()
        .replace(".xlsx", "")
    )

    for archivo in archivos:

        nombre_drive = (
            archivo.get(
                "name",
                ""
            )
            .lower()
            .strip()
            .replace(".xlsx", "")
        )

        if (
            nombre_busqueda == nombre_drive
            or nombre_busqueda in nombre_drive
            or nombre_drive in nombre_busqueda
        ):

            return archivo

    return None


def descargar_archivo_drive(
    file_id,
    ruta_destino
):

    info = (
        drive_service.files()
        .get(
            fileId=file_id,
            fields="id, name, mimeType",
            supportsAllDrives=True
        )
        .execute()
    )

    mime = info.get(
        "mimeType",
        ""
    )

    if mime == "application/vnd.google-apps.spreadsheet":

        request = (
            drive_service.files()
            .export_media(
                fileId=file_id,
                mimeType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        )

    else:

        request = (
            drive_service.files()
            .get_media(
                fileId=file_id,
                supportsAllDrives=True
            )
        )

    with open(
        ruta_destino,
        "wb"
    ) as archivo:

        from googleapiclient.http import MediaIoBaseDownload

        downloader = MediaIoBaseDownload(
            archivo,
            request
        )

        terminado = False

        while not terminado:

            status, terminado = downloader.next_chunk()


def descargar_por_nombre(
    nombre_archivo,
    folder_id,
    ruta_destino,
    obligatorio=False
):



    archivo = buscar_archivo_drive(
        nombre_archivo,
        folder_id
    )


    if archivo is None:

        st.sidebar.error(
            f"No encontrado: {nombre_archivo}"
        )

        if obligatorio:

            st.error(
                f"No se encontró en Drive: {nombre_archivo}"
            )

            st.stop()

        return False

    descargar_archivo_drive(
        archivo["id"],
        ruta_destino
    )

    st.sidebar.success(
        f"Descargado: {nombre_archivo}"
    )

    return True


def subir_archivo_drive(
    ruta_archivo,
    nombre_archivo,
    folder_id
):

    archivo_existente = buscar_archivo_drive(
        nombre_archivo,
        folder_id
    )

    media = MediaFileUpload(
        ruta_archivo,
        resumable=True
    )

    if archivo_existente:

        drive_service.files().update(
            fileId=archivo_existente["id"],
            media_body=media
        ).execute()

    else:

        metadata = {
            "name": nombre_archivo,
            "parents": [
                folder_id
            ]
        }

        drive_service.files().create(
            body=metadata,
            media_body=media,
            fields="id"
        ).execute()


def sincronizar_archivos_drive():

    # Optimizado: ya no se descarga agenda ni bases auxiliares desde Drive.
    # Las incidencias y el compendio principal viven en Supabase.
    return True


# =========================
# COLORES
# =========================

COLOR_VINO = "#7A1F3D"
COLOR_VINO_OSCURO = "#5F1830"
COLOR_DORADO = "#BC955C"
COLOR_VERDE = "#235B4E"
COLOR_GRIS = "#DDC9A3"
COLOR_ROJO = "#9F2241"
COLOR_FONDO = "#F7F7F7"

PALETA_INSTITUCIONAL = [
    COLOR_VINO,
    COLOR_DORADO,
    COLOR_VERDE,
    COLOR_ROJO,
    COLOR_GRIS,
    "#8A6F3D",
    "#3D6F8A",
    "#6F3D8A"
]

st.markdown(
    f"""
    <style>
    .stApp {{
        background-color: {COLOR_FONDO};
    }}

    h1, h2, h3 {{
        color: {COLOR_VINO};
        font-weight: 800;
    }}

    section[data-testid="stSidebar"] {{
        background-color: #ffffff;
        border-right: 4px solid {COLOR_VINO};
    }}

    .stButton>button {{
        background-color: {COLOR_VINO};
        color: white;
        border-radius: 10px;
        border: none;
        font-weight: bold;
    }}

    .stButton>button:hover {{
        background-color: {COLOR_VINO_OSCURO};
        color: white;
    }}

    div[data-testid="metric-container"] {{
        background-color: white;
        border-left: 6px solid {COLOR_DORADO};
        padding: 16px;
        border-radius: 14px;
        box-shadow: 0px 2px 8px rgba(0,0,0,0.08);
    }}

    div[data-testid="stDataFrame"] {{
        background-color: white;
        border-radius: 12px;
    }}
    </style>
    """,
    unsafe_allow_html=True
)


# =========================
# CATÁLOGOS
# =========================

ATRIBUIBLES = [
    "Estado",
    "Proveedor",
    "Operador Logístico",
    "IMSS-BIENESTAR",
    "Otro"
]


def homologar_atribuible(valor):

    texto = limpiar_valor_visual(valor).upper().strip()

    equivalencias = {
        "ESTADO": "Estado",
        "PROVEEDOR": "Proveedor",
        "OPERADOR LOGISTICO": "Operador Logístico",
        "OPERADOR LOGÍSTICO": "Operador Logístico",
        "OPERADOR": "Operador Logístico",
        "O.L.": "Operador Logístico",
        "OL": "Operador Logístico",
        "IMSS-BIENESTAR": "IMSS-BIENESTAR",
        "IMSS BIENESTAR": "IMSS-BIENESTAR",
        "IMSSB": "IMSS-BIENESTAR",
        "OTRO": "Otro"
    }

    return equivalencias.get(texto, limpiar_valor_visual(valor).title() if texto else "")



RESPONSABLES_FIJOS = {
    "YAIR OSWALDO GONZÁLEZ GARCÍA": [
        "YENNY ZORAID ARROYO GONZÁLEZ",
        "EMILY ESTEFANÍA SÁNCHEZ HERNÁNDEZ",
        "AMAIRANI DE GUADALUPE IRALA CERVANTES"
    ],
    "OSCAR IVÁN FERNÁNDEZ JIMÉNEZ": [
        "MARÍA FERNANDA IRALA CERVANTES",
        "MARIBEL RIVERA LUNA",
        "FABIOLA GÓMEZ RAMÍREZ"
    ]
}

MONITORES = []

for subdirector, lista_monitoras in RESPONSABLES_FIJOS.items():

    # Primero se agrega el supervisor solo.
    MONITORES.append(
        subdirector
    )

    for monitora in lista_monitoras:

        MONITORES.append(
            f"{subdirector} | {monitora}"
        )

TIPOS_INCIDENCIA_GENERAL = [
    "CORTA CADUCIDAD",
    "DOCUMENTACIÓN ERRÓNEA O INCOMPLETA",
    "ENTREGA DUPLICADA",
    "ENTREGA EN CLUES DIFERENTE",
    "ENTREGA FUERA DEL EJERCICIO FISCAL",
    "FALTA DE CITA",
    "Estado Rechaza Cita",
    "No Acudió a su cita",
    "INSUMO EN MAL ESTADO",
    "INSUMO INCOMPLETO",
    "MAL ETIQUETADO",
    "ORDENES CANCELADAS",
    "RECHAZO POR CAPACIDAD DEL ALMACÉN",
    "RECHAZO INJUSTIFICADO",
    "URGENCIA",
    "OTRO. ESPECIFICAR EN OBSERVACIONES"
]


# =========================
# FUNCIONES BASE
# =========================

def normalizar_texto(
    valor
):

    return (
        str(valor)
        .strip()
        .replace("None", "")
        .replace("nan", "")
    )


def normalizar_orden(
    valor
):

    return (
        str(valor)
        .upper()
        .strip()
        .replace(" ", "")
        .replace("_", "-")
        .replace("–", "-")
        .replace("—", "-")
    )


def limpiar_nombre_archivo(
    texto
):

    texto = str(texto)

    texto = re.sub(
        r"[^A-Za-z0-9_\-]+",
        "_",
        texto
    )

    return texto[:80]


def limpiar_nombre_carpeta(
    texto
):

    texto = str(texto).strip()

    if (
        not texto
        or texto.lower() == "nan"
        or texto.lower() == "none"
    ):

        texto = "SIN_DATO"

    texto = re.sub(
        r'[\\/:*?"<>|]+',
        "_",
        texto
    )

    texto = re.sub(
        r"\s+",
        " ",
        texto
    )

    return texto[:100]


def obtener_valor(
    fila,
    opciones
):

    if fila is None:
        return ""

    for col in opciones:

        try:
            # Las filas de pandas exponen ``index``; las filas que vienen de
            # Supabase/Google pueden llegar como diccionarios.  No asumimos
            # un solo tipo porque la migracion combina ambos origenes.
            if isinstance(fila, Mapping):
                if col not in fila:
                    continue
                valor = fila.get(col, "")
            elif hasattr(fila, "index"):
                if col not in fila.index:
                    continue
                valor = fila.get(col, "")
            else:
                valor = fila[col]
        except (AttributeError, KeyError, IndexError, TypeError):
            continue

        if (
            pd.notna(valor)
            and str(valor).strip()
            and str(valor).strip().lower() != "nan"
            and str(valor).strip().lower() != "none"
        ):

            return valor

    return ""


def convertir_numero(
    valor
):

    try:

        return float(
            str(valor)
            .replace(",", "")
            .replace("$", "")
            .strip()
        )

    except:

        return 0


def calcular_estatus_piezas(
    piezas_emitidas,
    piezas_validar
):

    emitidas = convertir_numero(
        piezas_emitidas
    )

    validadas = convertir_numero(
        piezas_validar
    )

    if validadas == 0:

        return "NO ENTREGADA"

    if validadas < emitidas:

        return "ENTREGA PARCIAL"

    return "ENTREGADA COMPLETA"


def calcular_estatus_incidencia_completa(
    estatus_entrega_estado
):

    if (
        str(estatus_entrega_estado)
        .upper()
        .strip()
        == "ENTREGADA COMPLETA"
    ):

        return "COMPLETA"

    return "INCOMPLETA"


def es_operador_logistico(
    tipo_entrega
):

    tipo = str(
        tipo_entrega
    ).upper().strip()

    return (
        "OPERADOR" in tipo
        or "LOGISTICO" in tipo
        or "LOGÍSTICO" in tipo
    )


def construir_almacen(
    clues_destino,
    unidad_destino,
    almacen_original
):

    clues = normalizar_texto(
        clues_destino
    )

    unidad = normalizar_texto(
        unidad_destino
    )

    almacen = normalizar_texto(
        almacen_original
    )

    if (
        clues
        and unidad
    ):

        return f"{clues} - {unidad}"

    if unidad:

        return unidad

    if clues:

        return clues

    return almacen


def fecha_a_texto(
    valor
):

    if (
        valor is None
        or str(valor).strip() == ""
        or str(valor).lower() == "nan"
        or str(valor).lower() == "none"
    ):

        return ""

    try:

        fecha = pd.to_datetime(
            valor,
            errors="coerce"
        )

        if pd.isna(
            fecha
        ):

            return str(
                valor
            )

        return fecha.strftime(
            "%d/%m/%Y"
        )

    except:

        return str(
            valor
        )


def agregar_fecha_rechazo_observaciones(observaciones, fecha_rechazo):
    """Guarda la fecha de rechazo sin requerir cambios en el esquema de Supabase."""
    texto = limpiar_valor_visual(observaciones)
    if not fecha_rechazo:
        return texto
    fecha = fecha_a_texto(fecha_rechazo)
    texto = re.sub(r"\s*\[FECHA_RECHAZO:[^\]]*\]", "", texto).strip()
    return f"{texto} [FECHA_RECHAZO:{fecha}]".strip()


def extraer_fecha_rechazo(observaciones):
    match = re.search(r"\[FECHA_RECHAZO:([^\]]+)\]", str(observaciones or ""))
    return fecha_a_texto(match.group(1)) if match else ""


def observaciones_sin_fecha_rechazo(observaciones):
    return re.sub(r"\s*\[FECHA_RECHAZO:[^\]]*\]", "", str(observaciones or "")).strip()


def agregar_fecha_incidencia_observaciones(observaciones, fecha_incidencia):
    """Guarda la fecha de incidencia sin requerir cambios en el esquema de Supabase."""
    texto = limpiar_valor_visual(observaciones)
    if not fecha_incidencia:
        return texto
    fecha = fecha_a_texto(fecha_incidencia)
    texto = re.sub(r"\s*\[FECHA_INCIDENCIA:[^\]]*\]", "", texto).strip()
    return f"{texto} [FECHA_INCIDENCIA:{fecha}]".strip()


def extraer_fecha_incidencia(observaciones):
    match = re.search(r"\[FECHA_INCIDENCIA:([^\]]+)\]", str(observaciones or ""))
    return fecha_a_texto(match.group(1)) if match else ""


def observaciones_sin_fechas_sistema(observaciones):
    texto = re.sub(r"\s*\[FECHA_RECHAZO:[^\]]*\]", "", str(observaciones or ""))
    texto = re.sub(r"\s*\[FECHA_INCIDENCIA:[^\]]*\]", "", texto)
    return texto.strip()


def limpiar_valor_visual(
    valor
):

    if valor is None:

        return ""

    try:

        if pd.isna(
            valor
        ):

            return ""

    except Exception:

        pass

    texto = str(
        valor
    ).strip()

    if texto.lower() in [
        "nan",
        "none",
        "null",
        "nat",
        "<na>"
    ]:

        return ""

    return texto


def formatear_fechas_cortas_df(df):
    """Formatea columnas de fecha para presentación/exportación como DD/MM/AAAA.

    No altera las fechas internas usadas por filtros, cálculos o controles del sistema.
    """
    if df is None or not isinstance(df, pd.DataFrame) or df.empty:
        return df

    resultado = df.copy()
    for columna in resultado.columns:
        nombre = str(columna).upper().strip()
        # Solo columnas cuyo encabezado identifica explícitamente una fecha.
        if "FECHA" not in nombre:
            continue

        serie_original = resultado[columna]
        fechas = pd.to_datetime(serie_original, errors="coerce", dayfirst=False)
        mascara = fechas.notna()
        if mascara.any():
            serie = serie_original.copy().astype("object")
            serie.loc[mascara] = fechas.loc[mascara].dt.strftime("%d/%m/%Y")
            resultado[columna] = serie

    return resultado


def limpiar_df_visual(
    df
):

    if df is None:

        return pd.DataFrame()

    if not isinstance(
        df,
        pd.DataFrame
    ):

        try:

            df = pd.DataFrame(
                df
            )

        except Exception:

            return pd.DataFrame()

    if df.empty:

        return df

    limpio = formatear_fechas_cortas_df(df.copy())

    limpio = limpio.replace(
        {
            None: np.nan,
            "None": np.nan,
            "none": np.nan,
            "nan": np.nan,
            "NaN": np.nan,
            "NULL": np.nan,
            "null": np.nan,
            "NaT": np.nan,
            "<NA>": np.nan
        }
    )

    columnas_utiles = []

    for columna in limpio.columns:

        serie = limpio[
            columna
        ]

        tiene_datos = (
            serie
            .astype(str)
            .str.strip()
            .replace(
                {
                    "nan": "",
                    "None": "",
                    "none": "",
                    "NULL": "",
                    "null": "",
                    "NaT": "",
                    "<NA>": ""
                }
            )
            .ne("")
            .any()
        )

        if tiene_datos:

            columnas_utiles.append(
                columna
            )

    limpio = limpio[
        columnas_utiles
    ].copy()

    limpio = limpio.fillna(
        ""
    )

    return limpio


def dataframe_limpio(
    df,
    use_container_width=True,
    hide_index=True
):

    st.dataframe(
        limpiar_df_visual(
            df
        ),
        use_container_width=use_container_width,
        hide_index=hide_index
    )


def texto_limpio(
    valor,
    default=""
):

    texto = limpiar_valor_visual(
        valor
    )

    if texto == "":

        return default

    return texto


# =========================
# SUPABASE HELPERS
# =========================

def buscar_orden_fuerte(
    valor_busqueda
):

    valor = normalizar_orden(
        valor_busqueda
    )

    if (
        not valor
        or valor == "NAN"
    ):

        return pd.DataFrame()

    respuesta = (
        supabase
        .table(
            "compendio"
        )
        .select(
            "*"
        )
        .eq(
            "orden_suministro",
            valor
        )
        .limit(
            100
        )
        .execute()
    )

    datos = respuesta.data

    if not datos:

        respuesta = (
            supabase
            .table(
                "compendio"
            )
            .select(
                "*"
            )
            .eq(
                "orden",
                valor
            )
            .limit(
                100
            )
            .execute()
        )

        datos = respuesta.data

    if not datos:

        respuesta = (
            supabase
            .table(
                "compendio"
            )
            .select(
                "*"
            )
            .eq(
                "no_orden",
                valor
            )
            .limit(
                100
            )
            .execute()
        )

        datos = respuesta.data

    if not datos:

        return pd.DataFrame()

    resultado = pd.DataFrame(
        datos
    )

    if "estatus_base" in resultado.columns:

        resultado["_prioridad"] = resultado[
            "estatus_base"
        ].apply(
            lambda x:
            0 if str(x).upper().strip() == "INACTIVA"
            else 1
        )

        resultado = resultado.sort_values(
            "_prioridad"
        )

        resultado = resultado.drop(
            columns=[
                "_prioridad"
            ],
            errors="ignore"
        )

    return resultado


def buscar_orden_compendio_rapida(
    valor_busqueda
):

    valor = normalizar_orden(
        valor_busqueda
    )

    if (
        not valor
        or valor == "NAN"
    ):

        return pd.DataFrame()

    columnas = (
        "orden_suministro, orden, no_orden, estatus_base, "
        "piezas_entregadas_clues"
    )

    for columna in [
        "orden_suministro",
        "orden",
        "no_orden"
    ]:

        try:

            respuesta = (
                supabase
                .table(
                    "compendio"
                )
                .select(
                    columnas
                )
                .eq(
                    columna,
                    valor
                )
                .limit(
                    1
                )
                .execute()
            )

            if respuesta.data:

                return pd.DataFrame(
                    respuesta.data
                )

        except Exception as e:

            mensaje = str(
                e
            )

            if (
                "57014" in mensaje
                or "statement timeout" in mensaje.lower()
                or "canceling statement" in mensaje.lower()
            ):

                return pd.DataFrame()

            continue

    return pd.DataFrame()


def sugerir_ordenes(
    valor_busqueda,
    limite=10
):

    valor = normalizar_orden(
        valor_busqueda
    )

    if (
        not valor
        or len(valor) < 8
    ):

        return pd.DataFrame()

    columnas = (
        "orden_suministro, orden, no_orden, estatus_base, "
        "entidad, proveedor, clave_cnis, descripcion"
    )

    consultas = [
        (
            "orden_suministro",
            valor
        ),
        (
            "orden",
            valor
        ),
        (
            "no_orden",
            valor
        )
    ]

    resultados = []

    for columna, texto in consultas:

        try:

            respuesta = (
                supabase
                .table(
                    "compendio"
                )
                .select(
                    columnas
                )
                .ilike(
                    columna,
                    f"{texto}%"
                )
                .limit(
                    limite
                )
                .execute()
            )

            datos = respuesta.data or []

            if datos:

                resultados.extend(
                    datos
                )

        except Exception as e:

            mensaje = str(
                e
            )

            if (
                "57014" in mensaje
                or "statement timeout" in mensaje.lower()
                or "canceling statement" in mensaje.lower()
            ):

                st.warning(
                    "La búsqueda de sugerencias tardó demasiado en Supabase. "
                    "Escribe más caracteres de la orden o usa la orden completa."
                )

                return pd.DataFrame()

            st.warning(
                f"No se pudieron cargar sugerencias desde Supabase: {e}"
            )

            return pd.DataFrame()

    if not resultados:

        return pd.DataFrame()

    sugerencias = pd.DataFrame(
        resultados
    )

    columnas_dedup = [
        c for c in [
            "orden_suministro",
            "orden",
            "no_orden"
        ]
        if c in sugerencias.columns
    ]

    if columnas_dedup:

        sugerencias = sugerencias.drop_duplicates(
            subset=columnas_dedup,
            keep="first"
        )

    return sugerencias.head(
        limite
    )


# =========================
# CARGAR DATOS
# =========================

@st.cache_data(
    ttl=300,
    show_spinner="Cargando incidencias desde Supabase..."
)
def cargar_incidencias():

    columnas_necesarias = {
        "ID": "",
        "FECHA_REGISTRO": "",
        "ORIGEN_REGISTRO": "BASE SUPABASE",
        "ORDEN_BUSCADA": "",
        "orden_suministro": "",
        "TIPO_ENTREGA": "",
        "ENTIDAD": "",
        "ALMACEN_CLUES_DESTINO": "",
        "CLUES_DESTINO": "",
        "UNIDAD_DESTINO": "",
        "PROVEEDOR": "",
        "ORDEN": "",
        "CLAVE_CNIS": "",
        "DESCRIPCION": "",
        "PIEZAS_EMITIDAS": "",
        "PIEZAS_RECIBIDAS_OL": "",
        "PIEZAS_ENTREGADAS_CLUES": "",
        "TIPO_RED": "",
        "GRUPO_TERAPEUTICO": "",
        "ESTATUS_OPERATIVO": "",
        "ESTATUS_BASE": "",
        "ORIGEN_COMPENDIO": "",
        "OPERADOR_LOGISTICO": "",
        "ESTATUS_RECEPCION_OL": "",
        "ESTATUS_ENTREGA_ESTADO": "",
        "ESTATUS_INCIDENCIA_COMPLETA": "INCOMPLETA",
        "ESTATUS_SEGUIMIENTO": "",
        "TIPO_INCIDENCIA": "",
        "ATRIBUIBLE A": "",
        "ESTATUS_INCIDENCIA": "En proceso",
        "RESPONSABLE": "",
        "OBSERVACIONES": "",
        "PDF_CEDULA_RECHAZO": "",
        "PDF_CORREO_SEGUIMIENTO": "",
        "CREADO_EN": ""
    }

    try:

        datos = []

        bloque = 900

        # Primero preguntamos cuántas filas existen realmente.
        # Así evitamos pedir rangos fuera de la tabla y evitamos el error PGRST103.
        respuesta_total = (
            supabase
            .table(
                "incidencias"
            )
            .select(
                "id",
                count="exact"
            )
            .limit(
                1
            )
            .execute()
        )

        total_conocido = int(
            getattr(
                respuesta_total,
                "count",
                0
            ) or 0
        )

        if total_conocido == 0:

            datos = []

        else:

            for inicio in range(
                0,
                total_conocido,
                bloque
            ):

                fin = min(
                    inicio + bloque - 1,
                    total_conocido - 1
                )

                respuesta = (
                    supabase
                    .table(
                        "incidencias"
                    )
                    .select(
                        "*"
                    )
                    .order(
                        "id",
                        desc=True
                    )
                    .range(
                        inicio,
                        fin
                    )
                    .execute()
                )

                parte = respuesta.data or []

                if parte:

                    datos.extend(
                        parte
                    )

    except Exception as e:

        st.error(
            f"No se pudieron cargar incidencias desde Supabase: {e}"
        )

        datos = []

    if not datos:

        return pd.DataFrame(
            columns=list(
                columnas_necesarias.keys()
            )
        )

    incidencias = pd.DataFrame(
        datos
    )

    renombres = {
        "id": "ID",
        "fecha_registro": "FECHA_REGISTRO",
        "origen_registro": "ORIGEN_REGISTRO",
        "orden_buscada": "ORDEN_BUSCADA",
        "tipo_entrega": "TIPO_ENTREGA",
        "entidad": "ENTIDAD",
        "almacen_clues_destino": "ALMACEN_CLUES_DESTINO",
        "clues_destino": "CLUES_DESTINO",
        "unidad_destino": "UNIDAD_DESTINO",
        "proveedor": "PROVEEDOR",
        "orden": "ORDEN",
        "clave_cnis": "CLAVE_CNIS",
        "descripcion": "DESCRIPCION",
        "piezas_emitidas": "PIEZAS_EMITIDAS",
        "piezas_recibidas_ol": "PIEZAS_RECIBIDAS_OL",
        "piezas_entregadas_clues": "PIEZAS_ENTREGADAS_CLUES",
        "tipo_red": "TIPO_RED",
        "grupo_terapeutico": "GRUPO_TERAPEUTICO",
        "estatus_operativo": "ESTATUS_OPERATIVO",
        "estatus_base": "ESTATUS_BASE",
        "origen_compendio": "ORIGEN_COMPENDIO",
        "operador_logistico": "OPERADOR_LOGISTICO",
        "estatus_recepcion_ol": "ESTATUS_RECEPCION_OL",
        "estatus_entrega_estado": "ESTATUS_ENTREGA_ESTADO",
        "estatus_incidencia_completa": "ESTATUS_INCIDENCIA_COMPLETA",
        "estatus_seguimiento": "ESTATUS_SEGUIMIENTO",
        "tipo_incidencia": "TIPO_INCIDENCIA",
        "atribuible_a": "ATRIBUIBLE A",
        "estatus_incidencia": "ESTATUS_INCIDENCIA",
        "responsable": "RESPONSABLE",
        "observaciones": "OBSERVACIONES",
        "pdf_cedula_rechazo": "PDF_CEDULA_RECHAZO",
        "pdf_correo_seguimiento": "PDF_CORREO_SEGUIMIENTO",
        "creado_en": "CREADO_EN"
    }

    incidencias = incidencias.rename(
        columns=renombres
    )

    if "OBSERVACIONES" in incidencias.columns:
        incidencias["FECHA_RECHAZO"] = incidencias["OBSERVACIONES"].apply(extraer_fecha_rechazo)
        incidencias["FECHA_INCIDENCIA"] = incidencias["OBSERVACIONES"].apply(extraer_fecha_incidencia)
        incidencias["OBSERVACIONES"] = incidencias["OBSERVACIONES"].apply(observaciones_sin_fechas_sistema)
    else:
        incidencias["FECHA_RECHAZO"] = ""
        incidencias["FECHA_INCIDENCIA"] = ""

    if "FECHA_REGISTRO" in incidencias.columns:
        incidencias["FECHA_REGISTRO"] = incidencias["FECHA_REGISTRO"].apply(fecha_a_texto)
        # Históricas: si no tenían fecha de incidencia, se toma su fecha de captura.
        mascara_sin_fecha = incidencias["FECHA_INCIDENCIA"].astype(str).str.strip().eq("")
        incidencias.loc[mascara_sin_fecha, "FECHA_INCIDENCIA"] = incidencias.loc[mascara_sin_fecha, "FECHA_REGISTRO"]

    # Homologación para registros importados desde Excel con encabezados distintos.
    columnas_alternas = {
        "CLAVE_CNIS": ["CLAVE CNIS", "CLAVE", "Clave CNIS", "clave", "clave cnis"],
        "ORDEN": ["ORDEN SUMINISTRO", "ORDEN DE SUMINISTRO", "Orden de Suministro", "NO. ORDEN"],
        "ATRIBUIBLE A": ["ATRIBUIBLE_A", "ATRIBUIBLE", "Atribuible a", "atribuible"],
        "PDF_CEDULA_RECHAZO": ["PDF CEDULA RECHAZO", "PDF CÉDULA RECHAZO", "CEDULA_RECHAZO", "CÉDULA RECHAZO"],
        "PDF_CORREO_SEGUIMIENTO": ["PDF CORREO SEGUIMIENTO", "CORREO_SEGUIMIENTO", "CORREO SEGUIMIENTO"]
    }

    for destino, alternas in columnas_alternas.items():

        if destino not in incidencias.columns:

            incidencias[destino] = ""

        for alterna in alternas:

            if alterna in incidencias.columns:

                incidencias[destino] = incidencias[destino].where(
                    incidencias[destino].astype(str).str.strip().ne(""),
                    incidencias[alterna]
                )

    if "ATRIBUIBLE A" in incidencias.columns:

        incidencias["ATRIBUIBLE A"] = incidencias["ATRIBUIBLE A"].apply(
            homologar_atribuible
        )

    for columna, valor_default in columnas_necesarias.items():

        if columna not in incidencias.columns:

            incidencias[columna] = valor_default

    columnas_ordenadas = list(
        columnas_necesarias.keys()
    )

    otras_columnas = [
        c for c in incidencias.columns
        if c not in columnas_ordenadas
    ]

    return incidencias[
        columnas_ordenadas + otras_columnas
    ]


@st.cache_data(
    ttl=300,
    show_spinner=False
)
def contar_incidencias_reales_supabase():

    try:

        respuesta = (
            supabase
            .table(
                "incidencias"
            )
            .select(
                "id",
                count="exact"
            )
            .limit(
                1
            )
            .execute()
        )

        return int(
            getattr(
                respuesta,
                "count",
                0
            ) or 0
        )

    except Exception:

        return 0



def obtener_incidencias_previas_supabase(
    orden
):

    orden_norm = normalizar_orden(
        orden
    )

    if not orden_norm:

        return pd.DataFrame()

    columnas_orden = [
        "orden_suministro",
        "orden",
        "orden_buscada"
    ]

    resultados = []

    # Primero intenta búsquedas exactas con índice.
    for columna in columnas_orden:

        try:

            respuesta = (
                supabase
                .table(
                    "incidencias"
                )
                .select(
                    "*"
                )
                .eq(
                    columna,
                    orden_norm
                )
                .order(
                    "id",
                    desc=True
                )
                .limit(
                    100
                )
                .execute()
            )

            if respuesta.data:

                resultados.extend(
                    respuesta.data
                )

        except Exception:

            continue

    if resultados:

        previas = pd.DataFrame(
            resultados
        )

        if "id" in previas.columns:

            previas = previas.drop_duplicates(
                subset=[
                    "id"
                ]
            )

            previas = previas.sort_values(
                "id",
                ascending=False
            )

        return previas

    # Fallback: si la orden se guardó con espacios, guiones raros o formato distinto,
    # trae las últimas incidencias y compara normalizado. La tabla es chica y evita falsos negativos.
    try:

        respuesta = (
            supabase
            .table(
                "incidencias"
            )
            .select(
                "*"
            )
            .order(
                "id",
                desc=True
            )
            .limit(
                10000
            )
            .execute()
        )

        datos = respuesta.data

        if not datos:

            return pd.DataFrame()

        previas = pd.DataFrame(
            datos
        )

        mascara = pd.Series(
            False,
            index=previas.index
        )

        for columna in columnas_orden:

            if columna in previas.columns:

                mascara = (
                    mascara
                    | previas[columna]
                    .astype(str)
                    .apply(normalizar_orden)
                    .eq(orden_norm)
                )

        previas = previas[
            mascara
        ].copy()

        if "id" in previas.columns:

            previas = previas.sort_values(
                "id",
                ascending=False
            )

        return previas

    except Exception as e:

        st.warning(
            f"No se pudieron consultar incidencias previas: {e}"
        )

        return pd.DataFrame()


def existe_incidencia_duplicada(
    orden,
    tipo_incidencia,
    atribuible_a
):

    previas = obtener_incidencias_previas_supabase(
        orden
    )

    if previas.empty:

        return False

    tipo_norm = str(
        tipo_incidencia
    ).upper().strip()

    atribuible_norm = str(
        atribuible_a
    ).upper().strip()

    if "tipo_incidencia" not in previas.columns:

        previas["tipo_incidencia"] = ""

    if "atribuible_a" not in previas.columns:

        previas["atribuible_a"] = ""

    duplicada = previas[
        (
            previas["tipo_incidencia"]
            .astype(str)
            .str.upper()
            .str.strip()
            == tipo_norm
        )
        &
        (
            previas["atribuible_a"]
            .astype(str)
            .str.upper()
            .str.strip()
            == atribuible_norm
        )
    ]

    return len(
        duplicada
    ) > 0


@st.cache_data(
    ttl=300,
    show_spinner="Cargando resumen ejecutivo..."
)
def cargar_resumen_incidencias():

    incidencias_temp = cargar_incidencias()

    total = len(
        incidencias_temp
    )

    if total == 0:

        return {
            "total": 0,
            "completas": 0,
            "incompletas": 0,
            "resueltas": 0,
            "porcentaje": 0
        }

    completas = incidencias_temp[
        incidencias_temp["ESTATUS_INCIDENCIA_COMPLETA"]
        .astype(str)
        .str.upper()
        .eq("COMPLETA")
    ].shape[0]

    incompletas = incidencias_temp[
        incidencias_temp["ESTATUS_INCIDENCIA_COMPLETA"]
        .astype(str)
        .str.upper()
        .eq("INCOMPLETA")
    ].shape[0]

    resueltas = incidencias_temp[
        incidencias_temp["ESTATUS_INCIDENCIA"]
        .astype(str)
        .str.upper()
        .isin(["RESUELTA", "RESUELTO"])
    ].shape[0]

    porcentaje = 0

    if total > 0:

        porcentaje = round(
            (resueltas / total) * 100,
            2
        )

    return {
        "total": total,
        "completas": completas,
        "incompletas": incompletas,
        "resueltas": resueltas,
        "porcentaje": porcentaje
    }


@st.cache_data(
    ttl=300,
    show_spinner="Cargando últimos registros..."
)
def cargar_incidencias_recientes(
    limite=300
):

    try:

        respuesta = (
            supabase
            .table(
                "incidencias"
            )
            .select(
                "*"
            )
            .order(
                "id",
                desc=True
            )
            .limit(
                limite
            )
            .execute()
        )

        datos = respuesta.data

        if not datos:

            return pd.DataFrame()

        return pd.DataFrame(
            datos
        )

    except Exception as e:

        st.error(
            f"No se pudieron cargar recientes desde Supabase: {e}"
        )

        return pd.DataFrame()


def obtener_incidencias_previas(
    incidencias,
    orden
):

    if incidencias.empty:

        return pd.DataFrame()

    incidencias = incidencias.copy()

    incidencias.columns = (
        incidencias.columns
        .astype(str)
        .str.strip()
    )

    columna_orden = None

    posibles_orden = [
        "orden_suministro",
        "ORDEN",
        "ORDEN_BUSCADA",
        "Orden de Suministro",
        "ORDEN DE SUMINISTRO",
        "NO. ORDEN",
        "no_orden"
    ]

    for col in posibles_orden:

        if col in incidencias.columns:

            columna_orden = col

            break

    if columna_orden is None:

        return pd.DataFrame()

    incidencias["_ORDEN_TEMP"] = (
        incidencias[columna_orden]
        .astype(str)
        .apply(normalizar_orden)
    )

    orden_norm = normalizar_orden(
        orden
    )

    previas = incidencias[
        incidencias["_ORDEN_TEMP"]
        ==
        orden_norm
    ].copy()

    previas = previas.drop(
        columns=[
            "_ORDEN_TEMP"
        ],
        errors="ignore"
    )

    return previas

def preparar_valor_supabase(
    valor
):

    if pd.isna(
        valor
    ):

        return None

    if isinstance(
        valor,
        (
            pd.Timestamp,
            datetime
        )
    ):

        return valor.isoformat()

    texto = str(
        valor
    ).strip()

    if texto.lower() in [
        "nan",
        "none",
        "null",
        ""
    ]:

        return None

    return texto


def guardar_incidencia(
    nueva,
    incidencia_id=""
):
    if str(nueva.get("TIPO_INCIDENCIA", "")).strip().upper() == "URGENCIA":
        raise ValueError("Las urgencias solo pueden registrarse en BASE_URGENCIAS de Google Sheets.")


    registro = {
        "fecha_registro": preparar_valor_supabase(nueva.get("FECHA_REGISTRO", datetime.now())),
        "origen_registro": preparar_valor_supabase(nueva.get("ORIGEN_REGISTRO", "SISTEMA")),
        "orden_buscada": preparar_valor_supabase(nueva.get("ORDEN_BUSCADA", "")),
        "orden_suministro": preparar_valor_supabase(nueva.get("orden_suministro", nueva.get("ORDEN", ""))),
        "tipo_entrega": preparar_valor_supabase(nueva.get("TIPO_ENTREGA", "")),
        "entidad": preparar_valor_supabase(nueva.get("ENTIDAD", "")),
        "almacen_clues_destino": preparar_valor_supabase(nueva.get("ALMACEN_CLUES_DESTINO", "")),
        "clues_destino": preparar_valor_supabase(nueva.get("CLUES_DESTINO", "")),
        "unidad_destino": preparar_valor_supabase(nueva.get("UNIDAD_DESTINO", "")),
        "proveedor": preparar_valor_supabase(nueva.get("PROVEEDOR", "")),
        "orden": preparar_valor_supabase(nueva.get("ORDEN", "")),
        "clave_cnis": preparar_valor_supabase(nueva.get("CLAVE_CNIS", "")),
        "descripcion": preparar_valor_supabase(nueva.get("DESCRIPCION", "")),
        "piezas_emitidas": preparar_valor_supabase(nueva.get("PIEZAS_EMITIDAS", "")),
        "piezas_recibidas_ol": preparar_valor_supabase(nueva.get("PIEZAS_RECIBIDAS_OL", "")),
        "piezas_entregadas_clues": preparar_valor_supabase(nueva.get("PIEZAS_ENTREGADAS_CLUES", "")),
        "tipo_red": preparar_valor_supabase(nueva.get("TIPO_RED", "")),
        "grupo_terapeutico": preparar_valor_supabase(nueva.get("GRUPO_TERAPEUTICO", "")),
        "estatus_operativo": preparar_valor_supabase(nueva.get("ESTATUS_OPERATIVO", "")),
        "estatus_base": preparar_valor_supabase(nueva.get("ESTATUS_BASE", "")),
        "origen_compendio": preparar_valor_supabase(nueva.get("ORIGEN_COMPENDIO", "")),
        "operador_logistico": preparar_valor_supabase(nueva.get("OPERADOR_LOGISTICO", "")),
        "estatus_recepcion_ol": preparar_valor_supabase(nueva.get("ESTATUS_RECEPCION_OL", "")),
        "estatus_entrega_estado": preparar_valor_supabase(nueva.get("ESTATUS_ENTREGA_ESTADO", "")),
        "estatus_incidencia_completa": preparar_valor_supabase(nueva.get("ESTATUS_INCIDENCIA_COMPLETA", "")),
        "estatus_seguimiento": preparar_valor_supabase(nueva.get("ESTATUS_SEGUIMIENTO", "")),
        "tipo_incidencia": preparar_valor_supabase(nueva.get("TIPO_INCIDENCIA", "")),
        "atribuible_a": preparar_valor_supabase(homologar_atribuible(nueva.get("ATRIBUIBLE A", ""))),
        "estatus_incidencia": preparar_valor_supabase(nueva.get("ESTATUS_INCIDENCIA", "")),
        "responsable": preparar_valor_supabase(nueva.get("RESPONSABLE", "")),
        "observaciones": preparar_valor_supabase(nueva.get("OBSERVACIONES", "")),
        "pdf_cedula_rechazo": preparar_valor_supabase(nueva.get("PDF_CEDULA_RECHAZO", "")),
        "pdf_correo_seguimiento": preparar_valor_supabase(nueva.get("PDF_CORREO_SEGUIMIENTO", ""))
    }

    if incidencia_id:

        respuesta = (
            supabase
            .table(
                "incidencias"
            )
            .update(
                registro
            )
            .eq(
                "id",
                incidencia_id
            )
            .execute()
        )

    else:

        respuesta = (
            supabase
            .table(
                "incidencias"
            )
            .insert(
                registro
            )
            .execute()
        )

    # Limpia el cache inmediatamente después de guardar para que
    # Resumen Ejecutivo, Seguimiento y Google Sheets lean datos nuevos.
    st.cache_data.clear()

    datos_insertados = getattr(
        respuesta,
        "data",
        None
    ) or []

    if datos_insertados and isinstance(
        datos_insertados[0],
        dict
    ):

        return datos_insertados[0].get(
            "id",
            incidencia_id
        )

    return incidencia_id or ""


# =========================
# EVIDENCIAS DRIVE
# =========================

def buscar_carpeta_drive(
    nombre_carpeta,
    parent_id
):

    nombre_carpeta = limpiar_nombre_carpeta(
        nombre_carpeta
    )

    query = (
        f"name = '{nombre_carpeta}' "
        f"and mimeType = 'application/vnd.google-apps.folder' "
        f"and '{parent_id}' in parents "
        f"and trashed = false"
    )

    resultado = (
        drive_service.files()
        .list(
            q=query,
            fields="files(id,name)",
            corpora="allDrives",
            supportsAllDrives=True,
            includeItemsFromAllDrives=True
        )
        .execute()
    )

    carpetas = resultado.get(
        "files",
        []
    )

    if carpetas:

        return carpetas[0]["id"]

    return None


def crear_carpeta_drive(
    nombre_carpeta,
    parent_id
):

    nombre_carpeta = limpiar_nombre_carpeta(
        nombre_carpeta
    )

    metadata = {
        "name": nombre_carpeta,
        "mimeType": "application/vnd.google-apps.folder",
        "parents": [
            parent_id
        ]
    }

    carpeta = (
        drive_service.files()
        .create(
            body=metadata,
            fields="id",
            supportsAllDrives=True
        )
        .execute()
    )

    return carpeta["id"]


def obtener_o_crear_carpeta_drive(
    nombre_carpeta,
    parent_id
):

    carpeta_id = buscar_carpeta_drive(
        nombre_carpeta,
        parent_id
    )

    if carpeta_id:

        return carpeta_id

    return crear_carpeta_drive(
        nombre_carpeta,
        parent_id
    )


def obtener_carpeta_evidencias_raiz():

    return obtener_o_crear_carpeta_drive(
        NOMBRE_CARPETA_EVIDENCIAS,
        FOLDER_ID_INCIDENCIAS_DRIVE
    )


def obtener_carpeta_respaldos_raiz():

    return obtener_o_crear_carpeta_drive(
        NOMBRE_CARPETA_RESPALDOS,
        FOLDER_ID_INCIDENCIAS_DRIVE
    )


def obtener_carpeta_evidencia(
    estado,
    clues
):

    carpeta_evidencias_id = obtener_carpeta_evidencias_raiz()

    carpeta_estado_id = obtener_o_crear_carpeta_drive(
        estado,
        carpeta_evidencias_id
    )

    carpeta_clues_id = obtener_o_crear_carpeta_drive(
        clues,
        carpeta_estado_id
    )

    return carpeta_clues_id


def subir_pdf_evidencia_drive(archivo, orden, tipo_pdf, estado, clues):
    """Sube un PDF a la carpeta existente; cada solicitud usa un temporal único."""
    if archivo is None:
        return ""
    ruta_local = None
    try:
        datos = archivo.getvalue()
        if not datos.startswith(b"%PDF-"):
            st.error("La evidencia no tiene una cabecera PDF válida.")
            return ""
        if len(datos) > 25 * 1024 * 1024:
            st.error("El PDF supera el límite de 25 MB.")
            return ""
        carpeta_destino_id = obtener_carpeta_evidencia(estado, clues)
        fecha = datetime.now().strftime("%Y%m%d_%H%M%S")
        nombre_archivo = (f"{fecha}_{uuid.uuid4().hex[:8]}_"
                          f"{limpiar_nombre_archivo(orden)}_"
                          f"{limpiar_nombre_archivo(tipo_pdf)}.pdf")
        with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
            ruta_local = tmp.name
            tmp.write(datos)
        media = MediaFileUpload(ruta_local, mimetype="application/pdf", resumable=True)
        nuevo = drive_service.files().create(
            body={"name": nombre_archivo, "parents": [carpeta_destino_id]},
            media_body=media, fields="id, webViewLink", supportsAllDrives=True
        ).execute()
        link = nuevo.get("webViewLink") or ("https://drive.google.com/file/d/" + nuevo["id"] + "/view")
        st.success("Evidencia PDF guardada en Google Drive.")
        return link
    except Exception:
        st.error("No se pudo subir la evidencia. Verifica Drive y vuelve a intentar.")
        return ""
    finally:
        if ruta_local:
            try:
                os.remove(ruta_local)
            except OSError:
                pass


def actualizar_evidencia_incidencia(
    incidencia_id,
    archivo,
    tipo_pdf,
    orden,
    estado,
    clues
):

    if archivo is None:

        return False

    link = subir_pdf_evidencia_drive(
        archivo,
        orden,
        tipo_pdf,
        estado,
        clues
    )

    if not link:

        return False

    columna = (
        "pdf_cedula_rechazo"
        if tipo_pdf == "cedula"
        else "pdf_correo_seguimiento"
    )

    supabase.table(
        "incidencias"
    ).update(
        {
            columna: link
        }
    ).eq(
        "id",
        incidencia_id
    ).execute()

    st.cache_data.clear()

    return True


def convertir_excel(
    df
):

    salida = BytesIO()

    with pd.ExcelWriter(
        salida,
        engine="openpyxl"
    ) as writer:

        df.to_excel(
            writer,
            index=False,
            sheet_name="Incidencias"
        )

    salida.seek(0)

    return salida


def preparar_datos_exportacion_urgencias(incidencias):
    """Prepara las urgencias capturadas y separa las órdenes pendientes.

    La fuente operativa es BASE_URGENCIAS. El formato histórico de la tabla
    ``incidencias`` se conserva únicamente para leer y migrar capturas previas.
    BASE_URGENCIAS no guarda una columna independiente para
    ``PENDIENTE_ENTREGA``; por eso se calcula al exportar a partir de piezas
    emitidas y piezas entregadas en la CLUES.

    La unidad de una urgencia es la pareja clave CNIS + entidad. Las filas
    inactivas, canceladas o recolectadas no forman parte de las urgencias
    vigentes. Una clave pasa a ``Atendida`` cuando sus órdenes activas ya no
    tienen piezas pendientes y existe una entrega registrada.
    """

    columnas_salida = [
        "FECHA_URGENCIA",
        "FECHA_REGISTRO",
        "ENTIDAD",
        "CLAVE_CNIS",
        "ORDEN_SUMINISTRO",
        "CLUES_DESTINO",
        "UNIDAD_DESTINO",
        "ALMACEN_CLUES_DESTINO",
        "PROVEEDOR",
        "DESCRIPCION",
        "TIPO_ENTREGA",
        "TIPO_RED",
        "GRUPO_TERAPEUTICO",
        "OPERADOR_LOGISTICO",
        "ESTATUS_BASE",
        "PIEZAS_EMITIDAS",
        "PIEZAS_RECIBIDAS_OL",
        "PIEZAS_ENTREGADAS_CLUES",
        "PENDIENTE_ENTREGA",
        "ESTATUS_ENTREGA_ENTIDAD",
        "ESTATUS_URGENCIA",
        "FECHA_ATENCION",
        "NUM_ORDENES_URGENCIA",
        "ORDENES_PENDIENTES",
        "ESTATUS_SEGUIMIENTO",
        "ESTATUS_INCIDENCIA",
        "RESPONSABLE",
        "OBSERVACIONES",
        "PDF_CORREO_SEGUIMIENTO",
        "ID"
    ]

    vacio = pd.DataFrame(columns=columnas_salida)

    if incidencias is None or not isinstance(incidencias, pd.DataFrame):
        return {
            "capturadas": vacio,
            "claves": vacio,
            "pendientes": vacio,
            "atendidas": vacio,
            "sin_orden": vacio,
            "columnas": columnas_salida
        }

    df = incidencias.copy()

    if "TIPO_INCIDENCIA" not in df.columns:
        return {
            "capturadas": vacio,
            "claves": vacio,
            "pendientes": vacio,
            "atendidas": vacio,
            "sin_orden": vacio,
            "columnas": columnas_salida
        }

    df = df[
        df["TIPO_INCIDENCIA"]
        .fillna("")
        .astype(str)
        .str.strip()
        .str.upper()
        .eq("URGENCIA")
    ].copy()

    if df.empty:
        return {
            "capturadas": vacio,
            "claves": vacio,
            "pendientes": vacio,
            "atendidas": vacio,
            "sin_orden": vacio,
            "columnas": columnas_salida
        }

    # Asegura encabezados que pueden faltar en registros antiguos.
    defaults = {
        "ID": "",
        "FECHA_REGISTRO": "",
        "ENTIDAD": "",
        "CLAVE_CNIS": "",
        "ORDEN": "",
        "orden_suministro": "",
        "ORDEN_BUSCADA": "",
        "CLUES_DESTINO": "",
        "UNIDAD_DESTINO": "",
        "ALMACEN_CLUES_DESTINO": "",
        "PROVEEDOR": "",
        "DESCRIPCION": "",
        "TIPO_ENTREGA": "",
        "TIPO_RED": "",
        "GRUPO_TERAPEUTICO": "",
        "OPERADOR_LOGISTICO": "",
        "ESTATUS_BASE": "",
        "ORIGEN_COMPENDIO": "",
        "ORIGEN": "",
        "PIEZAS_EMITIDAS": "",
        "PIEZAS_RECIBIDAS_OL": "",
        "PIEZAS_ENTREGADAS_CLUES": "",
        "ESTATUS_ENTREGA_ESTADO": "",
        "FECHA_ENTREGA": "",
        "FECHA_ENTREGA_CLUES": "",
        "FECHA_RECEPCION": "",
        "FECHA_RECEPCION_CLUES": "",
        "FECHA_ATENCION": "",
        "ESTATUS_SEGUIMIENTO": "",
        "ESTATUS_INCIDENCIA": "",
        "RESPONSABLE": "",
        "OBSERVACIONES": "",
        "PDF_CORREO_SEGUIMIENTO": ""
    }

    for columna, valor in defaults.items():
        if columna not in df.columns:
            df[columna] = valor

    # Conserva la fecha real de solicitud, que se escribe dentro de
    # OBSERVACIONES como FECHA_URGENCIA=AAAA-MM-DD.
    df["FECHA_URGENCIA"] = df.apply(
        obtener_fecha_urgencia_registro,
        axis=1
    )

    def primer_valor_fila(fila, columnas):
        for columna in columnas:
            valor = limpiar_valor_visual(fila.get(columna, ""))
            if valor:
                return valor
        return ""

    df["ORDEN_SUMINISTRO"] = df.apply(
        lambda fila: primer_valor_fila(
            fila,
            ["ORDEN", "orden_suministro", "ORDEN_BUSCADA"]
        ),
        axis=1
    )

    # Normaliza cantidades sin perder los valores visuales que se muestran
    # en el Excel.
    emitidas_num = df["PIEZAS_EMITIDAS"].apply(convertir_numero)
    entregadas_num = df["PIEZAS_ENTREGADAS_CLUES"].apply(convertir_numero)
    df["PENDIENTE_ENTREGA"] = (
        emitidas_num - entregadas_num
    ).clip(lower=0)

    estatus_existente = df["ESTATUS_ENTREGA_ESTADO"].apply(
        limpiar_valor_visual
    )
    df["ESTATUS_ENTREGA_ENTIDAD"] = estatus_existente

    faltan_estatus = df["ESTATUS_ENTREGA_ENTIDAD"].eq("")
    if faltan_estatus.any():
        df.loc[
            faltan_estatus,
            "ESTATUS_ENTREGA_ENTIDAD"
        ] = df.loc[
            faltan_estatus
        ].apply(
            lambda fila: calcular_estatus_piezas(
                fila.get("PIEZAS_EMITIDAS", ""),
                fila.get("PIEZAS_ENTREGADAS_CLUES", "")
            ),
            axis=1
        )

    # La combinación de estatus permite detectar pendientes aun cuando un
    # registro antiguo no tenga capturadas las cantidades.
    estatus_combinado = (
        df[
            [
                "ESTATUS_ENTREGA_ENTIDAD",
                "ESTATUS_SEGUIMIENTO",
                "ESTATUS_BASE",
                "ESTATUS_INCIDENCIA",
                "ORIGEN_COMPENDIO",
                "ORIGEN"
            ]
        ]
        .fillna("")
        .astype(str)
        .agg(" | ".join, axis=1)
        .str.upper()
    )

    tiene_orden = df["ORDEN_SUMINISTRO"].apply(
        limpiar_valor_visual
    ).ne("")

    cancelada = estatus_combinado.str.contains(
        r"CANCEL|INACTIV|RECOLECT",
        regex=True,
        na=False
    )

    texto_pendiente = estatus_combinado.str.contains(
        r"NO ENTREG|SIN ENTREG|PENDIENTE|FALTANTE|PARCIAL|PROGRAMADA|INCOMPLETA",
        regex=True,
        na=False
    )

    # El seguimiento puede actualizarse después de la captura y dejar el
    # texto histórico ``NO ENTREGADA`` en otra columna. Un estatus explícito
    # de entrega completa o piezas entregadas sin saldo pendiente prevalece
    # sobre ese texto anterior.
    entrega_confirmada = (
        entregadas_num.gt(0)
        & ~df["PENDIENTE_ENTREGA"].gt(0)
    ) | estatus_combinado.str.contains(
        r"ENTREGADA COMPLETA|COMPLETA[- ]ENTREGADO|COMPLETO[- ]ENTREGADO",
        regex=True,
        na=False
    )
    texto_pendiente_efectivo = texto_pendiente & ~entrega_confirmada

    # Algunas versiones del compendio incluyen fecha de entrega o recepción.
    # Se conserva cuando existe para documentar la atención posterior a la
    # captura. Si no existe, el estatus y las piezas son la fuente de verdad:
    # la orden sólo pudo capturarse como urgencia cuando aún no tenía entrega.
    if "FECHA_ATENCION" not in df.columns:
        df["FECHA_ATENCION"] = ""
    for columna_fecha in [
        "FECHA_ENTREGA",
        "FECHA_ENTREGA_CLUES",
        "FECHA_RECEPCION",
        "FECHA_RECEPCION_CLUES",
        "FECHA_ATENCION"
    ]:
        if columna_fecha not in df.columns:
            continue
        valores_fecha = df[columna_fecha].apply(limpiar_valor_visual)
        vacias = df["FECHA_ATENCION"].apply(limpiar_valor_visual).eq("")
        df.loc[vacias, "FECHA_ATENCION"] = valores_fecha.loc[vacias]

    pendiente_cuantificado = df["PENDIENTE_ENTREGA"].gt(0)
    entrega_completa = (
        ~cancelada
        & ~texto_pendiente_efectivo
        & ~pendiente_cuantificado
        & (
            entregadas_num.gt(0)
            | estatus_combinado.str.contains(
                r"ENTREGAD[OA]|COMPLETA",
                regex=True,
                na=False
            )
        )
    )

    # Si el origen aporta ambas fechas, no se permite que una entrega previa
    # a la captura atienda retroactivamente la urgencia.
    try:
        fecha_captura_dt = pd.to_datetime(
            df["FECHA_URGENCIA"],
            errors="coerce"
        )
        fecha_atencion_dt = pd.to_datetime(
            df["FECHA_ATENCION"],
            errors="coerce"
        )
        fecha_atencion_anterior = (
            fecha_captura_dt.notna()
            & fecha_atencion_dt.notna()
            & (
                fecha_atencion_dt.dt.normalize()
                < fecha_captura_dt.dt.normalize()
            )
        )
        entrega_completa = entrega_completa & ~fecha_atencion_anterior
    except Exception:
        # Los formatos de fecha históricos son heterogéneos; si no se pueden
        # comparar, se conserva la regla segura de estatus y piezas.
        pass

    df["_PENDIENTE"] = (
        tiene_orden
        & ~cancelada
        & ~entrega_completa
        & (
            df["PENDIENTE_ENTREGA"].gt(0)
            | texto_pendiente_efectivo
        )
    )
    df["_ATENDIDA"] = tiene_orden & entrega_completa
    df["_CANCELADA"] = cancelada

    # Una misma urgencia puede conservarse en Supabase con más de un registro
    # al actualizar una emisión. Mantén la última fila por ID cuando exista.
    if "ID" in df.columns:
        df["_ID_TEXTO"] = df["ID"].apply(limpiar_valor_visual)
        con_id = df["_ID_TEXTO"].ne("")
        if con_id.any():
            df_con_id = df.loc[con_id].drop_duplicates(
                subset=["_ID_TEXTO"],
                keep="last"
            )
            df_sin_id = df.loc[~con_id]
            df = pd.concat(
                [df_con_id, df_sin_id],
                ignore_index=True
            )

    # Recalcula la serie después de cualquier deduplicación para conservar el
    # mismo índice que el DataFrame al separar las solicitudes sin orden.
    tiene_orden = df["ORDEN_SUMINISTRO"].apply(
        limpiar_valor_visual
    ).ne("")

    df["_CLAVE_NORMALIZADA"] = df["CLAVE_CNIS"].apply(
        compactar_clave_urgencia
    )
    df["_ENTIDAD_NORMALIZADA"] = df["ENTIDAD"].apply(
        normalizar_entidad_urgencia
    )
    df["_LLAVE_URGENCIA"] = (
        df["_CLAVE_NORMALIZADA"]
        + "|"
        + df["_ENTIDAD_NORMALIZADA"]
    )
    df["_CLAVE_VALIDA"] = (
        df["_CLAVE_NORMALIZADA"].ne("")
        & df["_ENTIDAD_NORMALIZADA"].ne("")
    )
    df_activas = df.loc[~df["_CANCELADA"]].copy()

    # Una misma clave/entidad puede tener varias órdenes. La vista de
    # urgencias es una fila por clave/entidad, mientras que pendientes sigue
    # siendo una fila por orden para poder enviarlas a surtimiento.
    filas_claves = []
    for _, grupo in df_activas.loc[
        df_activas["_CLAVE_VALIDA"]
    ].groupby(
        "_LLAVE_URGENCIA",
        sort=False
    ):
        representante = grupo.iloc[0].copy()
        ordenes = []
        for orden in grupo["ORDEN_SUMINISTRO"].apply(limpiar_valor_visual):
            if orden and orden not in ordenes:
                ordenes.append(orden)

        pendientes_grupo = int(grupo["_PENDIENTE"].sum())
        atendidas_grupo = int(grupo["_ATENDIDA"].sum())
        tiene_orden_grupo = bool(
            grupo["ORDEN_SUMINISTRO"].apply(
                limpiar_valor_visual
            ).ne("").any()
        )

        if pendientes_grupo:
            estatus_urgencia = "Pendiente"
        elif atendidas_grupo and tiene_orden_grupo:
            estatus_urgencia = "Atendida"
        else:
            estatus_urgencia = "Revisar con emisión"

        representante["ORDEN_SUMINISTRO"] = "; ".join(ordenes)
        representante["NUM_ORDENES_URGENCIA"] = len(ordenes)
        representante["ORDENES_PENDIENTES"] = pendientes_grupo
        representante["ESTATUS_URGENCIA"] = estatus_urgencia

        for columna_cantidad in [
            "PIEZAS_EMITIDAS",
            "PIEZAS_RECIBIDAS_OL",
            "PIEZAS_ENTREGADAS_CLUES",
            "PENDIENTE_ENTREGA"
        ]:
            representante[columna_cantidad] = grupo[columna_cantidad].apply(
                convertir_numero
            ).sum()

        fechas_atencion = grupo["FECHA_ATENCION"].apply(
            limpiar_valor_visual
        )
        fechas_atencion = fechas_atencion[fechas_atencion.ne("")]
        if not fechas_atencion.empty:
            representante["FECHA_ATENCION"] = fechas_atencion.iloc[-1]

        filas_claves.append(representante)

    claves_trabajo = pd.DataFrame(filas_claves)

    def seleccionar_columnas(trabajo):
        if trabajo is None or trabajo.empty:
            return vacio.copy()
        resultado = trabajo.copy()
        for columna in columnas_salida:
            if columna not in resultado.columns:
                resultado[columna] = ""
        return resultado[columnas_salida].copy()

    capturadas = seleccionar_columnas(df_activas)
    claves = seleccionar_columnas(claves_trabajo)
    pendientes = seleccionar_columnas(
        df_activas[df_activas["_PENDIENTE"]]
    )
    sin_orden = seleccionar_columnas(
        df_activas[
            ~df_activas["ORDEN_SUMINISTRO"].apply(
                limpiar_valor_visual
            ).ne("")
        ]
    )
    atendidas = seleccionar_columnas(
        claves_trabajo[
            claves_trabajo["ESTATUS_URGENCIA"].eq("Atendida")
        ]
        if not claves_trabajo.empty
        else claves_trabajo
    )

    if not claves.empty:
        claves = claves.sort_values(
            ["ENTIDAD", "FECHA_URGENCIA", "CLAVE_CNIS"],
            kind="stable"
        )

    if not atendidas.empty:
        atendidas = atendidas.sort_values(
            ["ENTIDAD", "FECHA_ATENCION", "CLAVE_CNIS"],
            kind="stable"
        )

    if not pendientes.empty:
        pendientes = pendientes.sort_values(
            ["ENTIDAD", "FECHA_URGENCIA", "ORDEN_SUMINISTRO"],
            kind="stable"
        )

    return {
        "capturadas": capturadas,
        "claves": claves,
        "pendientes": pendientes,
        "atendidas": atendidas,
        "sin_orden": sin_orden,
        "columnas": columnas_salida
    }


def convertir_excel_urgencias(incidencias):
    """Genera el concentrado descargable de urgencias y pendientes."""

    if (
        isinstance(incidencias, dict)
        and all(
            clave in incidencias
            for clave in [
                "capturadas",
                "pendientes",
                "sin_orden"
            ]
        )
    ):
        datos = incidencias
    else:
        datos = preparar_datos_exportacion_urgencias(incidencias)
    capturadas = datos["capturadas"]
    claves = datos.get("claves", capturadas)
    pendientes = datos["pendientes"]
    atendidas = datos.get(
        "atendidas",
        pd.DataFrame(columns=claves.columns)
    )
    sin_orden = datos["sin_orden"]

    # Todas las fechas visibles/exportadas salen en formato corto DD/MM/AAAA.
    capturadas = formatear_fechas_cortas_df(capturadas)
    claves = formatear_fechas_cortas_df(claves)
    pendientes = formatear_fechas_cortas_df(pendientes)
    atendidas = formatear_fechas_cortas_df(atendidas)
    sin_orden = formatear_fechas_cortas_df(sin_orden)

    piezas_pendientes = 0
    if not pendientes.empty and "PENDIENTE_ENTREGA" in pendientes.columns:
        piezas_pendientes = int(
            pendientes["PENDIENTE_ENTREGA"]
            .apply(convertir_numero)
            .sum()
        )

    resumen = pd.DataFrame(
        [
            ["Fecha de generación", fecha_a_texto(datetime.now())],
            ["Urgencias capturadas (claves por entidad)", len(claves)],
            ["Órdenes pendientes de entrega", len(pendientes)],
            ["Claves atendidas", len(atendidas)],
            ["Piezas pendientes (cuando están cuantificadas)", piezas_pendientes],
            ["Solicitudes sin orden / revisar emisión", len(sin_orden)]
        ],
        columns=["Indicador", "Valor"]
    )

    salida = BytesIO()

    with pd.ExcelWriter(
        salida,
        engine="openpyxl"
    ) as writer:
        resumen.to_excel(
            writer,
            index=False,
            sheet_name="Resumen"
        )
        pendientes.to_excel(
            writer,
            index=False,
            sheet_name="Pendientes_Entrega"
        )
        claves.to_excel(
            writer,
            index=False,
            sheet_name="Urgencias_Capturadas"
        )
        atendidas.to_excel(
            writer,
            index=False,
            sheet_name="Atendidas"
        )
        sin_orden.to_excel(
            writer,
            index=False,
            sheet_name="Sin_Orden"
        )

        from openpyxl.styles import Alignment, Font, PatternFill

        relleno_header = PatternFill(
            "solid",
            fgColor="1F4E78"
        )
        relleno_pendiente = PatternFill(
            "solid",
            fgColor="FFF2CC"
        )

        for nombre_hoja in [
            "Resumen",
            "Pendientes_Entrega",
            "Urgencias_Capturadas",
            "Atendidas",
            "Sin_Orden"
        ]:
            ws = writer.sheets[nombre_hoja]
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions

            for celda in ws[1]:
                celda.fill = relleno_header
                celda.font = Font(
                    color="FFFFFF",
                    bold=True
                )
                celda.alignment = Alignment(
                    horizontal="center",
                    vertical="center",
                    wrap_text=True
                )

            for columna in ws.columns:
                letra = columna[0].column_letter
                ancho = min(
                    max(
                        len(str(celda.value or ""))
                        for celda in columna
                    ) + 2,
                    45
                )
                ws.column_dimensions[letra].width = max(ancho, 12)

        ws_pendientes = writer.sheets["Pendientes_Entrega"]
        for fila in range(2, ws_pendientes.max_row + 1):
            for celda in ws_pendientes[fila]:
                celda.fill = relleno_pendiente

        # Convierte los vínculos de evidencia en enlaces clicables.
        for nombre_hoja in [
            "Pendientes_Entrega",
            "Urgencias_Capturadas",
            "Atendidas",
            "Sin_Orden"
        ]:
            ws = writer.sheets[nombre_hoja]
            encabezados = {
                celda.value: celda.column
                for celda in ws[1]
            }
            col_pdf = encabezados.get("PDF_CORREO_SEGUIMIENTO")

            if col_pdf:
                for fila in range(2, ws.max_row + 1):
                    celda = ws.cell(
                        row=fila,
                        column=col_pdf
                    )
                    url = str(celda.value or "").strip()
                    if url and url.lower() not in {"nan", "none", "null"}:
                        celda.value = "Ver PDF"
                        celda.hyperlink = url
                        celda.style = "Hyperlink"

    salida.seek(0)
    return salida



def convertir_excel_cedulas_rechazo(df):
    """Genera un Excel independiente con hipervínculos clicables a las cédulas."""
    salida = BytesIO()
    exportar = formatear_fechas_cortas_df(df.copy())

    if "PDF_CEDULA_RECHAZO" in exportar.columns:
        exportar = exportar.rename(columns={"PDF_CEDULA_RECHAZO": "ABRIR PDF"})

    with pd.ExcelWriter(salida, engine="openpyxl") as writer:
        exportar.to_excel(writer, index=False, sheet_name="Cedulas de rechazo")
        ws = writer.sheets["Cedulas de rechazo"]

        encabezados = {celda.value: celda.column for celda in ws[1]}
        col_pdf = encabezados.get("ABRIR PDF")

        if col_pdf:
            for fila in range(2, ws.max_row + 1):
                celda = ws.cell(row=fila, column=col_pdf)
                url = str(celda.value or "").strip()
                if url and url.lower() not in {"nan", "none", "null"}:
                    celda.value = "Ver PDF"
                    celda.hyperlink = url
                    celda.style = "Hyperlink"

        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions

        for columna in ws.columns:
            letra = columna[0].column_letter
            ancho = min(max(len(str(c.value or "")) for c in columna) + 2, 45)
            ws.column_dimensions[letra].width = max(ancho, 12)

    salida.seek(0)
    return salida




def buscar_google_sheet_maestro():

    query = (
        f"name = '{NOMBRE_GOOGLE_SHEET_INCIDENCIAS}' "
        f"and mimeType = 'application/vnd.google-apps.spreadsheet' "
        f"and '{FOLDER_ID_INCIDENCIAS_DRIVE}' in parents "
        f"and trashed = false"
    )

    resultado = (
        drive_service.files()
        .list(
            q=query,
            fields="files(id, name, webViewLink)",
            corpora="allDrives",
            supportsAllDrives=True,
            includeItemsFromAllDrives=True
        )
        .execute()
    )

    archivos = resultado.get(
        "files",
        []
    )

    if archivos:

        return archivos[0]

    return None


def crear_google_sheet_maestro():

    metadata = {
        "name": NOMBRE_GOOGLE_SHEET_INCIDENCIAS,
        "mimeType": "application/vnd.google-apps.spreadsheet",
        "parents": [
            FOLDER_ID_INCIDENCIAS_DRIVE
        ]
    }

    archivo = (
        drive_service.files()
        .create(
            body=metadata,
            fields="id, name, webViewLink",
            supportsAllDrives=True
        )
        .execute()
    )

    return archivo


def obtener_o_crear_google_sheet_maestro():

    archivo = buscar_google_sheet_maestro()

    if archivo:

        return archivo

    return crear_google_sheet_maestro()


def preparar_dataframe_para_sheets(df):

    if df is None or df.empty:

        return [
            ["SIN_DATOS"]
        ]

    limpio = df.copy()

    limpio = limpiar_df_visual(
        limpio
    )

    for columna in limpio.columns:

        limpio[columna] = limpio[columna].apply(
            lambda x: fecha_a_texto(x)
            if isinstance(x, (pd.Timestamp, datetime))
            else limpiar_valor_visual(x)
        )

    valores = [
        limpio.columns.astype(str).tolist()
    ]

    valores.extend(
        limpio.astype(str).values.tolist()
    )

    return valores


def actualizar_google_sheets_maestro():

    archivo = obtener_o_crear_google_sheet_maestro()

    spreadsheet_id = archivo["id"]

    incidencias_sheet = cargar_incidencias()

    valores = preparar_dataframe_para_sheets(
        incidencias_sheet
    )

    sheets_service.spreadsheets().values().clear(
        spreadsheetId=spreadsheet_id,
        range="Incidencias!A:ZZ"
    ).execute()

    try:

        sheets_service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body={
                "requests": [
                    {
                        "updateSheetProperties": {
                            "properties": {
                                "sheetId": 0,
                                "title": "Incidencias"
                            },
                            "fields": "title"
                        }
                    }
                ]
            }
        ).execute()

    except Exception:

        pass

    sheets_service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range="Incidencias!A1",
        valueInputOption="RAW",
        body={
            "values": valores
        }
    ).execute()

    return archivo.get(
        "webViewLink",
        f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}"
    )


def generar_respaldo_drive():

    try:

        return actualizar_google_sheets_maestro()

    except Exception as e:

        st.warning(
            f"La incidencia se guardó, pero no se pudo actualizar el Google Sheets maestro: {e}"
        )

        return ""


# =========================
# URGENCIAS: CRUCE DE CLAVES Y ÓRDENES
# =========================

COLUMNAS_BASE_URGENCIAS_ORDENES = [
    "FECHA_URGENCIA",
    "SOLICITUD",
    "CLAVE_CNIS",
    "ENTIDAD",
    "ENTIDAD_COMPENDIO",
    "ORDEN_SUMINISTRO",
    "ESTATUS_BASE",
    "TIPO_ENTREGA",
    "CLUES_DESTINO",
    "UNIDAD_DESTINO",
    "ALMACEN",
    "PROVEEDOR",
    "DESCRIPCION",
    "PIEZAS_EMITIDAS",
    "PIEZAS_ENTREGADAS_CLUES",
    "PENDIENTE_ENTREGA",
    "ESTATUS_ENTREGA_ENTIDAD",
    "TIPO_RED",
    "GRUPO_TERAPEUTICO",
    "OPERADOR_LOGISTICO",
    "ORIGEN_COMPENDIO",
    "RESPONSABLE",
    "FECHA_REGISTRO",
    "ID_URGENCIA",
    "ID_INCIDENCIA",
    "OBSERVACIONES",
    "PDF_CORREO_SEGUIMIENTO",
    "FECHA_ACTUALIZACION_BASE"
]

COLUMNAS_BASE_URGENCIAS_SIN_ORDEN = [
    "FECHA_URGENCIA",
    "SOLICITUD",
    "CLAVE_CNIS",
    "ENTIDAD",
    "ESTATUS",
    "RESPONSABLE",
    "FECHA_REGISTRO",
    "ID_URGENCIA",
    "ID_INCIDENCIA",
    "MOTIVO",
    "OBSERVACIONES",
    "PDF_CORREO_SEGUIMIENTO",
    "FECHA_ACTUALIZACION_BASE"
]


def normalizar_clave_urgencia(valor):

    texto = limpiar_valor_visual(valor).upper().strip()
    texto = texto.replace("\ufeff", "").replace("\u200b", "")

    if re.fullmatch(r"\d+\.0+", texto):

        texto = texto.split(".", 1)[0]

    return re.sub(r"\s+", "", texto)


def compactar_clave_urgencia(valor):

    return re.sub(
        r"[^A-Z0-9]+",
        "",
        normalizar_clave_urgencia(valor)
    )


def variantes_clave_urgencia(valor):

    original = limpiar_valor_visual(valor).strip()

    if not original:

        return []

    normalizada = normalizar_clave_urgencia(original)
    compacta = compactar_clave_urgencia(original)
    variantes = [
        original,
        original.replace(" ", ""),
        normalizada,
        compacta
    ]

    if re.fullmatch(r"\d+", normalizada):

        variantes.extend(
            [
                f"{normalizada}.0",
                str(int(normalizada))
            ]
        )

    elif re.fullmatch(r"\d+", compacta):

        variantes.append(
            f"{compacta}.0"
        )

    resultado = []

    for variante in variantes:

        variante = str(variante).strip()

        if variante and variante not in resultado:

            resultado.append(variante)

    return resultado


def normalizar_entidad_urgencia(valor):

    texto = limpiar_valor_visual(valor).upper().strip()

    texto = unicodedata.normalize(
        "NFKD",
        texto
    ).encode(
        "ascii",
        "ignore"
    ).decode(
        "ascii"
    )

    texto = re.sub(
        r"[^A-Z0-9]+",
        " ",
        texto
    )

    texto = re.sub(
        r"\s+",
        " ",
        texto
    ).strip()

    equivalencias = {
        "MEXICO": "ESTADO DE MEXICO",
        "ESTADO DE MEXICO": "ESTADO DE MEXICO",
        "EDOMEX": "ESTADO DE MEXICO",
        "CDMX": "CIUDAD DE MEXICO",
        "CIUDAD DE MEXICO": "CIUDAD DE MEXICO",
        "DISTRITO FEDERAL": "CIUDAD DE MEXICO",
        "VERACRUZ": "VERACRUZ",
        "VERACRUZ DE IGNACIO DE LA LLAVE": "VERACRUZ",
        "MICHOACAN DE OCAMPO": "MICHOACAN",
        "MICHOACAN": "MICHOACAN",
        "QUERETARO DE ARTEAGA": "QUERETARO",
        "QUERETARO": "QUERETARO",
        "COAHUILA DE ZARAGOZA": "COAHUILA",
        "COAHUILA": "COAHUILA"
    }

    return equivalencias.get(
        texto,
        texto
    )


def entidades_urgencia_coinciden(valor_a, valor_b):

    entidad_a = normalizar_entidad_urgencia(valor_a)
    entidad_b = normalizar_entidad_urgencia(valor_b)

    if not entidad_a or not entidad_b:

        return False

    if entidad_a == entidad_b:

        return True

    return (
        entidad_a in entidad_b
        or entidad_b in entidad_a
    )


def homologar_entidad_urgencia_desde_compendio(
    compendio,
    entidad_solicitada
):

    if compendio is None or compendio.empty:

        return ""

    columnas = columnas_orden_urgencia()["entidad"]

    for _, fila in compendio.iterrows():

        entidad_compendio = limpiar_valor_visual(
            valor_fila_urgencia(
                fila,
                columnas
            )
        )

        if entidad_compendio and entidades_urgencia_coinciden(
            entidad_compendio,
            entidad_solicitada
        ):

            return entidad_compendio

    return ""


def resumir_error_consulta_compendio_urgencia(error):

    texto = str(
        error
    ).strip()

    if not texto:

        texto = type(
            error
        ).__name__

    for secreto in [
        os.environ.get(
            "SUPABASE_KEY",
            ""
        ),
        os.environ.get(
            "SUPABASE_URL",
            ""
        )
    ]:

        if secreto:

            texto = texto.replace(
                secreto,
                "[oculto]"
            )

    return texto[:600]


def consultar_compendio_por_clave_urgencia(clave):

    variantes = variantes_clave_urgencia(clave)

    if not variantes:

        return pd.DataFrame()

    datos = []
    errores = []

    # La tabla actual sólo tiene ``clave_cnis``.  Buscamos coincidencia exacta
    # y detenemos las variantes al encontrar filas, evitando recorridos
    # innecesarios sobre el compendio grande.
    for columna in [
        "clave_cnis"
    ]:

        for variante in variantes:

            try:

                respuesta = (
                    supabase
                    .table(
                        "compendio"
                    )
                    .select(
                        "*"
                    )
                    .eq(
                        columna,
                        variante
                    )
                    .limit(
                        1000
                    )
                    .execute()
                )

                filas = respuesta.data or []

                if filas:

                    datos.extend(
                        filas
                    )

                    break

            except Exception as error:

                errores.append(
                    f"{columna} exacta: "
                    f"{resumir_error_consulta_compendio_urgencia(error)}"
                )

        if datos:

            break

    if not datos:

        textos_busqueda = [
            normalizar_clave_urgencia(clave),
            compactar_clave_urgencia(clave)
        ]

        for columna in [
            "clave_cnis"
        ]:

            for texto_busqueda in textos_busqueda:

                if len(texto_busqueda) < 3:

                    continue

                try:

                    respuesta = (
                        supabase
                        .table(
                            "compendio"
                        )
                        .select(
                            "*"
                        )
                        .ilike(
                            columna,
                            f"%{texto_busqueda}%"
                        )
                        .limit(
                            1000
                        )
                        .execute()
                    )

                    filas = respuesta.data or []

                    if filas:

                        datos.extend(
                            filas
                        )

                        break

                except Exception as error:

                    errores.append(
                        f"{columna} parcial: "
                        f"{resumir_error_consulta_compendio_urgencia(error)}"
                    )

            if datos:

                break

    if not datos:

        salida = pd.DataFrame()

        salida.attrs[
            "error_consulta_urgencia"
        ] = " | ".join(
            dict.fromkeys(
                errores
            )
        )[:1200]

        return salida

    salida = pd.DataFrame(
        datos
    ).drop_duplicates(
        ignore_index=True
    )

    salida.attrs[
        "error_consulta_urgencia"
    ] = ""

    return salida


ENTIDADES_URGENCIA = [
    "AGUASCALIENTES",
    "BAJA CALIFORNIA",
    "BAJA CALIFORNIA SUR",
    "CAMPECHE",
    "CHIAPAS",
    "CHIHUAHUA",
    "CIUDAD DE MEXICO",
    "COAHUILA DE ZARAGOZA",
    "COLIMA",
    "DURANGO",
    "GUANAJUATO",
    "GUERRERO",
    "HIDALGO",
    "JALISCO",
    "MEXICO",
    "MICHOACAN DE OCAMPO",
    "MORELOS",
    "NAYARIT",
    "NUEVO LEON",
    "OAXACA",
    "PUEBLA",
    "QUERETARO DE ARTEAGA",
    "QUINTANA ROO",
    "SAN LUIS POTOSI",
    "SINALOA",
    "SONORA",
    "TABASCO",
    "TAMAULIPAS",
    "TLAXCALA",
    "VERACRUZ",
    "YUCATAN",
    "ZACATECAS"
]


def obtener_entidades_compendio_urgencia():

    # El catálogo de captura debe ser estable y no depender de los primeros
    # registros devueltos por la tabla compendio de Supabase.
    return ENTIDADES_URGENCIA.copy()


def valor_fila_urgencia(fila, opciones):

    return obtener_valor(
        fila,
        opciones
    )


def columnas_orden_urgencia():

    return {
        "clave": [
            "clave_cnis",
            "CLAVE_CNIS",
            "clave",
            "CLAVE"
        ],
        "entidad": [
            "entidad",
            "ENTIDAD",
            "estado",
            "ESTADO",
            "entidad_destino",
            "Entidad de destino"
        ],
        "orden": [
            "orden_suministro",
            "ORDEN_SUMINISTRO",
            "orden",
            "ORDEN",
            "no_orden",
            "NO_ORDEN"
        ],
        "estatus_base": [
            "estatus_base",
            "ESTATUS_BASE",
            "estatus",
            "ESTATUS"
        ],
        "tipo_entrega": [
            "tipo_entrega",
            "TIPO_ENTREGA"
        ],
        "clues": [
            "clues_destino",
            "CLUES_DESTINO",
            "clues de destino",
            "CLUES de destino",
            "clues destino",
            "CLUES destino",
            "clues",
            "CLUES"
        ],
        "unidad": [
            "unidad_destino",
            "UNIDAD_DESTINO",
            "unidad",
            "UNIDAD"
        ],
        "almacen": [
            "almacen",
            "ALMACEN",
            "almacén",
            "ALMACÉN"
        ],
        "proveedor": [
            "proveedor",
            "PROVEEDOR"
        ],
        "descripcion": [
            "descripcion",
            "DESCRIPCION",
            "descripción",
            "DESCRIPCIÓN"
        ],
        "piezas_emitidas": [
            "piezas_emitidas",
            "PIEZAS_EMITIDAS",
            "piezas emitidas",
            "Piezas emitidas"
        ],
        "piezas_recibidas_ol": [
            "piezas_recibidas_ol",
            "PIEZAS_RECIBIDAS_OL"
        ],
        "piezas_entregadas": [
            "piezas_entregadas_clues",
            "PIEZAS_ENTREGADAS_CLUES",
            "piezas entregadas clues",
            "Piezas entregadas CLUES"
        ],
        "tipo_red": [
            "tipo_red",
            "TIPO_RED"
        ],
        "grupo": [
            "grupo_terapeutico",
            "GRUPO_TERAPEUTICO"
        ],
        "operador": [
            "operador_logistico",
            "OPERADOR_LOGISTICO"
        ],
        "origen_compendio": [
            "origen_compendio",
            "ORIGEN_COMPENDIO"
        ],
        "estatus_entrega": [
            "estatus_entrega_estado",
            "estatus_entrega_clues",
            "estatus_entrega",
            "ESTATUS_ENTREGA_ESTADO"
        ]
    }


def fila_urgencia_tiene_entrega(fila):

    columnas = columnas_orden_urgencia()

    piezas_entregadas = convertir_numero(
        valor_fila_urgencia(
            fila,
            columnas["piezas_entregadas"]
        )
    )

    if piezas_entregadas > 0:

        return True

    estatus = limpiar_valor_visual(
        valor_fila_urgencia(
            fila,
            columnas["estatus_entrega"]
        )
    ).upper()

    if not estatus:

        return False

    if any(
        palabra in estatus
        for palabra in [
            "NO ENTREG",
            "SIN ENTREG",
            "PENDIENTE",
            "FALTANTE",
            "NO RECIB"
        ]
    ):

        return False

    return "ENTREG" in estatus


def construir_resultado_orden_urgencia(
    fila,
    clave_solicitada,
    entidad_solicitada
):

    columnas = columnas_orden_urgencia()

    clave = limpiar_valor_visual(
        valor_fila_urgencia(
            fila,
            columnas["clave"]
        )
    ) or limpiar_valor_visual(
        clave_solicitada
    )

    entidad_compendio = limpiar_valor_visual(
        valor_fila_urgencia(
            fila,
            columnas["entidad"]
        )
    )

    orden = limpiar_valor_visual(
        valor_fila_urgencia(
            fila,
            columnas["orden"]
        )
    )

    piezas_emitidas = valor_fila_urgencia(
        fila,
        columnas["piezas_emitidas"]
    )

    piezas_entregadas = valor_fila_urgencia(
        fila,
        columnas["piezas_entregadas"]
    )

    emitidas_num = convertir_numero(
        piezas_emitidas
    )

    entregadas_num = convertir_numero(
        piezas_entregadas
    )

    return {
        "CLAVE_CNIS": clave,
        "ENTIDAD": entidad_compendio or limpiar_valor_visual(
            entidad_solicitada
        ),
        "ENTIDAD_COMPENDIO": entidad_compendio,
        "ORDEN_SUMINISTRO": orden,
        "ESTATUS_BASE": limpiar_valor_visual(
            valor_fila_urgencia(
                fila,
                columnas["estatus_base"]
            )
        ),
        "TIPO_ENTREGA": limpiar_valor_visual(
            valor_fila_urgencia(
                fila,
                columnas["tipo_entrega"]
            )
        ),
        "CLUES_DESTINO": limpiar_valor_visual(
            valor_fila_urgencia(
                fila,
                columnas["clues"]
            )
        ),
        "UNIDAD_DESTINO": limpiar_valor_visual(
            valor_fila_urgencia(
                fila,
                columnas["unidad"]
            )
        ),
        "ALMACEN": construir_almacen(
            valor_fila_urgencia(
                fila,
                columnas["clues"]
            ),
            valor_fila_urgencia(
                fila,
                columnas["unidad"]
            ),
            valor_fila_urgencia(
                fila,
                columnas["almacen"]
            )
        ),
        "PROVEEDOR": limpiar_valor_visual(
            valor_fila_urgencia(
                fila,
                columnas["proveedor"]
            )
        ),
        "DESCRIPCION": limpiar_valor_visual(
            valor_fila_urgencia(
                fila,
                columnas["descripcion"]
            )
        ),
        "PIEZAS_EMITIDAS": piezas_emitidas,
        "PIEZAS_RECIBIDAS_OL": valor_fila_urgencia(
            fila,
            columnas["piezas_recibidas_ol"]
        ),
        "PIEZAS_ENTREGADAS_CLUES": piezas_entregadas,
        "PENDIENTE_ENTREGA": max(
            emitidas_num - entregadas_num,
            0
        ),
        "ESTATUS_ENTREGA_ENTIDAD": calcular_estatus_piezas(
            piezas_emitidas,
            piezas_entregadas
        ),
        "TIPO_RED": limpiar_valor_visual(
            valor_fila_urgencia(
                fila,
                columnas["tipo_red"]
            )
        ),
        "GRUPO_TERAPEUTICO": limpiar_valor_visual(
            valor_fila_urgencia(
                fila,
                columnas["grupo"]
            )
        ),
        "OPERADOR_LOGISTICO": limpiar_valor_visual(
            valor_fila_urgencia(
                fila,
                columnas["operador"]
            )
        ),
        "ORIGEN_COMPENDIO": limpiar_valor_visual(
            valor_fila_urgencia(
                fila,
                columnas["origen_compendio"]
            )
        )
    }


def buscar_ordenes_urgencia(clave, entidad):

    compendio = consultar_compendio_por_clave_urgencia(
        clave
    )

    error_consulta = str(
        compendio.attrs.get(
            "error_consulta_urgencia",
            ""
        )
        or ""
    )

    vacio = pd.DataFrame(
        columns=COLUMNAS_BASE_URGENCIAS_ORDENES
    )

    if compendio.empty:

        return {
            "ordenes": vacio,
            "ordenes_entregadas": pd.DataFrame(),
            "ordenes_emitidas": vacio,
            "total_clave": 0,
            "total_entidad": 0,
            "total_con_orden": 0,
            "sin_orden": False,
            "requiere_revision_emision": False,
            "clave_valida": False,
            "entidad_homologada": "",
            "error_consulta": error_consulta
        }

    columnas = columnas_orden_urgencia()
    entidad_homologada = homologar_entidad_urgencia_desde_compendio(
        compendio,
        entidad
    )

    if not entidad_homologada:

        return {
            "ordenes": vacio,
            "ordenes_entregadas": pd.DataFrame(),
            "ordenes_emitidas": vacio,
            "total_clave": len(compendio),
            "total_entidad": 0,
            "total_con_orden": 0,
            "sin_orden": True,
            "requiere_revision_emision": True,
            "clave_valida": True,
            "entidad_homologada": ""
        }

    filas_entidad = []

    for _, fila in compendio.iterrows():

        entidad_fila = valor_fila_urgencia(
            fila,
            columnas["entidad"]
        )

        if entidades_urgencia_coinciden(
            entidad_fila,
            entidad
        ):

            filas_entidad.append(
                fila
            )

    registros = []

    for fila in filas_entidad:

        orden = limpiar_valor_visual(
            valor_fila_urgencia(
                fila,
                columnas["orden"]
            )
        )

        if not orden:

            continue

        registro = construir_resultado_orden_urgencia(
            fila,
            clave,
            entidad_homologada
        )

        registro["_ORDEN_NORMALIZADA"] = normalizar_orden(
            orden
        )

        registro["_TIENE_ENTREGA"] = fila_urgencia_tiene_entrega(
            fila
        )

        registros.append(
            registro
        )

    if not registros:

        return {
            "ordenes": vacio,
            "ordenes_entregadas": pd.DataFrame(),
            "ordenes_emitidas": vacio,
            "total_clave": len(compendio),
            "total_entidad": len(filas_entidad),
            "total_con_orden": 0,
            "sin_orden": True,
            "requiere_revision_emision": True,
            "clave_valida": True,
            "entidad_homologada": entidad_homologada
        }

    elegibles = []
    entregadas = []
    registros_df = pd.DataFrame(
        registros
    )

    ordenes_emitidas = registros_df.drop(
        columns=[
            "_ORDEN_NORMALIZADA",
            "_TIENE_ENTREGA"
        ],
        errors="ignore"
    ).copy()

    if not ordenes_emitidas.empty:

        ordenes_emitidas = ordenes_emitidas.reindex(
            columns=[
                columna for columna in COLUMNAS_BASE_URGENCIAS_ORDENES
                if columna in ordenes_emitidas.columns
            ]
        )

    for _, grupo in registros_df.groupby(
        "_ORDEN_NORMALIZADA",
        dropna=False
    ):

        grupo = grupo.copy()

        if grupo["_TIENE_ENTREGA"].astype(bool).any():

            fila_entregada = grupo.sort_values(
                "PIEZAS_ENTREGADAS_CLUES",
                key=lambda serie: serie.apply(convertir_numero),
                ascending=False
            ).iloc[0].to_dict()

            fila_entregada["MOTIVO_EXCLUSION"] = (
                "Orden con entrega registrada en la entidad"
            )

            entregadas.append(
                fila_entregada
            )

        else:

            elegibles.append(
                grupo.iloc[0].to_dict()
            )

    ordenes = pd.DataFrame(
        elegibles
    ).drop(
        columns=[
            "_ORDEN_NORMALIZADA",
            "_TIENE_ENTREGA"
        ],
        errors="ignore"
    )

    if not ordenes.empty:

        ordenes = ordenes.reindex(
            columns=[
                columna for columna in COLUMNAS_BASE_URGENCIAS_ORDENES
                if columna in ordenes.columns
            ]
        )

    ordenes_entregadas = pd.DataFrame(
        entregadas
    ).drop(
        columns=[
            "_ORDEN_NORMALIZADA",
            "_TIENE_ENTREGA"
        ],
        errors="ignore"
    )

    return {
        "ordenes": ordenes,
        "ordenes_entregadas": ordenes_entregadas,
        "ordenes_emitidas": ordenes_emitidas,
        "total_clave": len(compendio),
        "total_entidad": len(filas_entidad),
        "total_con_orden": len(elegibles) + len(entregadas),
        "sin_orden": len(registros) == 0,
        "requiere_revision_emision": False,
        "clave_valida": True,
        "entidad_homologada": entidad_homologada
    }


def fecha_urgencia_iso(valor):

    fecha = pd.to_datetime(
        valor,
        errors="coerce"
    )

    if pd.isna(fecha):

        return limpiar_valor_visual(
            valor
        )

    return fecha.strftime(
        "%Y-%m-%d"
    )


def generar_id_urgencia():
    """Genera un identificador propio para la base independiente de urgencias."""

    return (
        "URG-"
        f"{datetime.now().strftime('%Y%m%d%H%M%S%f')}-"
        f"{uuid.uuid4().hex[:8].upper()}"
    )


def obtener_id_urgencia(registro):
    """Obtiene el ID de urgencia, conservando IDs históricos como respaldo."""

    return limpiar_valor_visual(
        obtener_valor(
            registro,
            [
                "ID_URGENCIA",
                "ID_INCIDENCIA",
                "ID",
                "id"
            ]
        )
    )


def observaciones_urgencia(fecha, observaciones=""):

    fecha_iso = fecha_urgencia_iso(
        fecha
    )

    texto = limpiar_valor_visual(
        observaciones
    )

    marcador = f"FECHA_URGENCIA={fecha_iso}"

    if texto:

        return f"{marcador}\n{texto}"

    return marcador


def obtener_fecha_urgencia_registro(fila):

    observaciones = limpiar_valor_visual(
        obtener_valor(
            fila,
            [
                "OBSERVACIONES",
                "observaciones"
            ]
        )
    )

    coincidencia = re.search(
        r"FECHA_URGENCIA\s*=\s*(\d{4}-\d{2}-\d{2})",
        observaciones,
        flags=re.IGNORECASE
    )

    if coincidencia:

        return coincidencia.group(1)

    return fecha_urgencia_iso(
        obtener_valor(
            fila,
            [
                "FECHA_REGISTRO",
                "fecha_registro"
            ]
        )
    )


def urgencia_ya_registrada(
    incidencias_actuales,
    clave,
    entidad,
    fecha,
    orden=""
):

    if incidencias_actuales is None or incidencias_actuales.empty:

        return False

    fecha_iso = fecha_urgencia_iso(
        fecha
    )

    clave_norm = compactar_clave_urgencia(
        clave
    )

    orden_norm = normalizar_orden(
        orden
    ) if limpiar_valor_visual(orden) else ""

    for _, fila in incidencias_actuales.iterrows():

        if limpiar_valor_visual(
            obtener_valor(
                fila,
                [
                    "TIPO_INCIDENCIA",
                    "tipo_incidencia"
                ]
            )
        ).upper() != "URGENCIA":

            continue

        if compactar_clave_urgencia(
            obtener_valor(
                fila,
                [
                    "CLAVE_CNIS",
                    "clave_cnis",
                    "CLAVE",
                    "clave"
                ]
            )
        ) != clave_norm:

            continue

        if not entidades_urgencia_coinciden(
            obtener_valor(
                fila,
                [
                    "ENTIDAD",
                    "entidad"
                ]
            ),
            entidad
        ):

            continue

        orden_fila = normalizar_orden(
            obtener_valor(
                fila,
                [
                    "ORDEN",
                    "orden",
                    "orden_suministro",
                    "ORDEN_BUSCADA"
                ]
            )
        )

        if orden_fila != orden_norm:

            continue

        observaciones = limpiar_valor_visual(
            obtener_valor(
                fila,
                [
                    "OBSERVACIONES",
                    "observaciones"
                ]
            )
        )

        if f"FECHA_URGENCIA={fecha_iso}" in observaciones:

            return True

        fecha_registro = pd.to_datetime(
            obtener_valor(
                fila,
                [
                    "FECHA_REGISTRO",
                    "fecha_registro"
                ]
            ),
            errors="coerce"
        )

        if not pd.isna(fecha_registro) and fecha_registro.strftime(
            "%Y-%m-%d"
        ) == fecha_iso:

            return True

    return False


def obtener_urgencia_revision_pendiente(
    incidencias_actuales,
    clave,
    entidad,
    ids_excluidos=None
):

    if incidencias_actuales is None or incidencias_actuales.empty:

        return None

    ids_excluidos = {
        str(valor).strip()
        for valor in (ids_excluidos or set())
        if str(valor).strip()
    }

    clave_compacta = compactar_clave_urgencia(clave)

    for _, fila in incidencias_actuales.iterrows():

        incidencia_id = limpiar_valor_visual(
            obtener_valor(
                fila,
                [
                    "ID_URGENCIA",
                    "ID_INCIDENCIA",
                    "ID",
                    "id"
                ]
            )
        )

        if incidencia_id in ids_excluidos:

            continue

        if not incidencia_id:

            continue

        if limpiar_valor_visual(
            obtener_valor(
                fila,
                [
                    "TIPO_INCIDENCIA",
                    "tipo_incidencia"
                ]
            )
        ).upper() != "URGENCIA":

            continue

        clave_fila = compactar_clave_urgencia(
            obtener_valor(
                fila,
                [
                    "CLAVE_CNIS",
                    "clave_cnis",
                    "CLAVE",
                    "clave"
                ]
            )
        )

        if not clave_compacta or clave_fila != clave_compacta:

            continue

        if not entidades_urgencia_coinciden(
            obtener_valor(
                fila,
                [
                    "ENTIDAD",
                    "entidad"
                ]
            ),
            entidad
        ):

            continue

        orden = limpiar_valor_visual(
            obtener_valor(
                fila,
                [
                    "ORDEN",
                    "orden",
                    "orden_suministro",
                    "ORDEN_BUSCADA"
                ]
            )
        )

        if orden:

            continue

        estatus = " ".join(
            limpiar_valor_visual(
                obtener_valor(
                    fila,
                    [
                        columna
                    ]
                )
            )
            for columna in [
                "ESTATUS_OPERATIVO",
                "ESTATUS_BASE",
                "ESTATUS_SEGUIMIENTO",
                "OBSERVACIONES"
            ]
        )

        estatus_normalizado = normalizar_entidad_urgencia(
            estatus
        )

        if (
            "REVISAR CON EMISION" not in estatus_normalizado
            and "SIN ORDEN" not in estatus_normalizado
        ):

            continue

        return fila.to_dict()

    return None


def buscar_google_sheet_urgencias():

    query = (
        f"name = '{NOMBRE_GOOGLE_SHEET_URGENCIAS}' "
        "and mimeType = 'application/vnd.google-apps.spreadsheet' "
        f"and '{FOLDER_ID_INCIDENCIAS_DRIVE}' in parents "
        "and trashed = false"
    )

    resultado = (
        drive_service
        .files()
        .list(
            q=query,
            fields="files(id, name, webViewLink)",
            corpora="allDrives",
            supportsAllDrives=True,
            includeItemsFromAllDrives=True
        )
        .execute()
    )

    archivos = resultado.get(
        "files",
        []
    )

    return archivos[0] if archivos else None


def crear_google_sheet_urgencias():

    metadata = {
        "name": NOMBRE_GOOGLE_SHEET_URGENCIAS,
        "mimeType": "application/vnd.google-apps.spreadsheet",
        "parents": [
            FOLDER_ID_INCIDENCIAS_DRIVE
        ]
    }

    return (
        drive_service
        .files()
        .create(
            body=metadata,
            fields="id, name, webViewLink",
            supportsAllDrives=True
        )
        .execute()
    )


def obtener_o_crear_google_sheet_urgencias():

    archivo = buscar_google_sheet_urgencias()

    if archivo:

        return archivo

    return crear_google_sheet_urgencias()


def asegurar_hojas_google_sheet_urgencias(spreadsheet_id):

    def obtener_hojas():

        respuesta = (
            sheets_service
            .spreadsheets()
            .get(
                spreadsheetId=spreadsheet_id,
                fields="sheets(properties(sheetId,title,index))"
            )
            .execute()
        )

        return [
            hoja.get(
                "properties",
                {}
            )
            for hoja in respuesta.get(
                "sheets",
                []
            )
        ]

    hojas = obtener_hojas()
    titulos = {
        hoja.get(
            "title",
            ""
        )
        for hoja in hojas
    }

    if NOMBRE_HOJA_URGENCIAS_ORDENES not in titulos:

        solicitudes = []

        if len(hojas) == 1 and hojas[0].get(
            "title",
            ""
        ) in [
            "Sheet1",
            "Hoja 1",
            "Página1",
            "Pagina1"
        ]:

            solicitudes.append(
                {
                    "updateSheetProperties": {
                        "properties": {
                            "sheetId": hojas[0].get(
                                "sheetId"
                            ),
                            "title": NOMBRE_HOJA_URGENCIAS_ORDENES
                        },
                        "fields": "title"
                    }
                }
            )

        else:

            solicitudes.append(
                {
                    "addSheet": {
                        "properties": {
                            "title": NOMBRE_HOJA_URGENCIAS_ORDENES
                        }
                    }
                }
            )

        sheets_service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body={
                "requests": solicitudes
            }
        ).execute()

    hojas = obtener_hojas()
    titulos = {
        hoja.get(
            "title",
            ""
        )
        for hoja in hojas
    }

    if NOMBRE_HOJA_URGENCIAS_SIN_ORDEN not in titulos:

        sheets_service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body={
                "requests": [
                    {
                        "addSheet": {
                            "properties": {
                                "title": NOMBRE_HOJA_URGENCIAS_SIN_ORDEN
                            }
                        }
                    }
                ]
            }
        ).execute()


def leer_hoja_urgencias(spreadsheet_id, nombre_hoja):

    try:

        respuesta = (
            sheets_service
            .spreadsheets()
            .values()
            .get(
                spreadsheetId=spreadsheet_id,
                range=f"{nombre_hoja}!A:ZZ"
            )
            .execute()
        )

    except Exception:

        return pd.DataFrame()

    valores = respuesta.get(
        "values",
        []
    )

    if len(valores) < 2:

        return pd.DataFrame()

    encabezados = [
        limpiar_valor_visual(valor)
        for valor in valores[0]
    ]

    if not encabezados or encabezados == [
        "SIN_DATOS"
    ]:

        return pd.DataFrame()

    filas = []

    for valores_fila in valores[1:]:

        fila = list(valores_fila)

        if len(fila) < len(encabezados):

            fila.extend(
                [""] * (len(encabezados) - len(fila))
            )

        filas.append(
            fila[:len(encabezados)]
        )

    return pd.DataFrame(
        filas,
        columns=encabezados
    )


def normalizar_urgencias_drive_para_exportacion(
    df,
    sin_orden=False
):
    """Adapta BASE_URGENCIAS a un formato común sin insertarla en incidencias."""

    if df is None or df.empty:
        return pd.DataFrame()

    trabajo = df.copy()

    def asegurar_columna(nombre, valor=""):
        if nombre not in trabajo.columns:
            trabajo[nombre] = valor

    def copiar_si_falta(destino, fuentes):
        asegurar_columna(destino, "")
        # Supabase puede entregar columnas como ``string[pyarrow]``. Se
        # convierten a object antes de completar valores para evitar el
        # TypeError de Pandas al asignar una serie con otro dtype.
        trabajo[destino] = trabajo[destino].astype(object)
        for fuente in fuentes:
            if fuente not in trabajo.columns:
                continue
            vacias = trabajo[destino].map(
                limpiar_valor_visual
            ).eq("")
            if vacias.any():
                valores_fuente = (
                    trabajo.loc[vacias, fuente]
                    .astype(object)
                    .map(limpiar_valor_visual)
                    .to_numpy(dtype=object)
                )
                trabajo.loc[vacias, destino] = valores_fuente

    # Encabezados de la hoja Ordenes_Urgencias.
    copiar_si_falta(
        "ORDEN",
        [
            "ORDEN_SUMINISTRO",
            "orden_suministro"
        ]
    )
    copiar_si_falta(
        "ORDEN_BUSCADA",
        [
            "ORDEN",
            "ORDEN_SUMINISTRO"
        ]
    )
    copiar_si_falta(
        "ALMACEN_CLUES_DESTINO",
        [
            "ALMACEN"
        ]
    )
    copiar_si_falta(
        "ESTATUS_ENTREGA_ESTADO",
        [
            "ESTATUS_ENTREGA_ENTIDAD",
            "ESTATUS_ENTREGA"
        ]
    )
    copiar_si_falta(
        "ID",
        [
            "ID_INCIDENCIA",
            "ID_URGENCIA",
            "id_incidencia"
        ]
    )
    copiar_si_falta(
        "ID_URGENCIA",
        [
            "ID_INCIDENCIA",
            "ID",
            "id",
            "id_incidencia"
        ]
    )
    copiar_si_falta(
        "FECHA_REGISTRO",
        [
            "FECHA_URGENCIA",
            "SOLICITUD"
        ]
    )

    asegurar_columna(
        "TIPO_INCIDENCIA",
        "URGENCIA"
    )
    trabajo["TIPO_INCIDENCIA"] = "URGENCIA"

    asegurar_columna(
        "ESTATUS_INCIDENCIA",
        "En proceso"
    )
    trabajo["ESTATUS_INCIDENCIA"] = trabajo[
        "ESTATUS_INCIDENCIA"
    ].apply(
        lambda valor: limpiar_valor_visual(valor) or "En proceso"
    )

    if sin_orden:
        trabajo["ORDEN"] = ""
        trabajo["ORDEN_BUSCADA"] = ""
        trabajo["ESTATUS_BASE"] = trabajo.get(
            "ESTATUS",
            "REVISAR CON EMISIÓN"
        )
        trabajo["ESTATUS_SEGUIMIENTO"] = "REVISAR CON EMISIÓN"
        trabajo["ESTATUS_ENTREGA_ESTADO"] = "NO ENTREGADA"

    return trabajo


@st.cache_data(
    ttl=300,
    show_spinner=False
)
def cargar_urgencias_desde_base():
    """Lee exclusivamente la base independiente BASE_URGENCIAS."""

    fuentes = []

    try:
        archivo = buscar_google_sheet_urgencias()

        if archivo and archivo.get("id"):
            ordenes = normalizar_urgencias_drive_para_exportacion(
                leer_hoja_urgencias(
                    archivo["id"],
                    NOMBRE_HOJA_URGENCIAS_ORDENES
                ),
                sin_orden=False
            )
            sin_orden = normalizar_urgencias_drive_para_exportacion(
                leer_hoja_urgencias(
                    archivo["id"],
                    NOMBRE_HOJA_URGENCIAS_SIN_ORDEN
                ),
                sin_orden=True
            )

            if not ordenes.empty:
                fuentes.append(ordenes)
            if not sin_orden.empty:
                fuentes.append(sin_orden)

    except Exception:
        pass

    if not fuentes:
        return pd.DataFrame()

    return pd.concat(
        fuentes,
        ignore_index=True,
        sort=False
    )


@st.cache_data(
    ttl=300,
    show_spinner=False
)
def cargar_urgencias_para_registro():
    """Obtiene urgencias de su base independiente y conserva legado para deduplicar."""

    fuentes = []

    base_urgencias = cargar_urgencias_desde_base()
    if base_urgencias is not None and not base_urgencias.empty:
        fuentes.append(base_urgencias)

    # Sólo se consulta el legado para evitar duplicar capturas antiguas.
    # Las nuevas urgencias nunca se insertan en la tabla incidencias.
    try:
        historicas = cargar_incidencias()
        if (
            historicas is not None
            and not historicas.empty
            and "TIPO_INCIDENCIA" in historicas.columns
        ):
            historicas = historicas[
                historicas["TIPO_INCIDENCIA"]
                .fillna("")
                .astype(str)
                .str.strip()
                .str.upper()
                .eq("URGENCIA")
            ].copy()
            if not historicas.empty:
                fuentes.insert(
                    0,
                    normalizar_urgencias_drive_para_exportacion(
                        historicas,
                        sin_orden=False
                    )
                )
    except Exception:
        pass

    if not fuentes:
        return pd.DataFrame()

    return pd.concat(
        fuentes,
        ignore_index=True,
        sort=False
    )


@st.cache_data(
    ttl=300,
    show_spinner=False
)
def cargar_urgencias_para_exportacion():
    """Consolida urgencias de Supabase y de la base histórica de Drive."""

    fuentes = []
    incidencias_supabase = pd.DataFrame()

    try:
        incidencias_supabase = cargar_incidencias()
    except Exception:
        incidencias_supabase = pd.DataFrame()

    # Se conserva el legado sólo mientras se migra. La base independiente se
    # agrega al final para que sus registros separados tengan prioridad.
    if incidencias_supabase is not None and not incidencias_supabase.empty:
        fuentes.append(incidencias_supabase)

    base_urgencias = cargar_urgencias_desde_base()
    if base_urgencias is not None and not base_urgencias.empty:
        fuentes.append(base_urgencias)

    if not fuentes:
        return preparar_datos_exportacion_urgencias(
            pd.DataFrame(
                columns=["TIPO_INCIDENCIA"]
            )
        )

    combinado = pd.concat(
        fuentes,
        ignore_index=True,
        sort=False
    )

    return preparar_datos_exportacion_urgencias(
        combinado
    )


def deduplicar_base_urgencias(df, claves):

    if df is None or df.empty:

        return pd.DataFrame()

    trabajo = df.copy()

    for columna in claves:

        if columna not in trabajo.columns:

            trabajo[columna] = ""

    trabajo["_LLAVE_UNICA"] = trabajo[claves].apply(
        lambda fila: "|".join(
            normalizar_clave_urgencia(valor)
            if columna == "CLAVE_CNIS"
            else normalizar_orden(valor)
            if columna == "ORDEN_SUMINISTRO"
            else limpiar_valor_visual(valor).upper().strip()
            for columna, valor in fila.items()
        ),
        axis=1
    )

    trabajo = trabajo.drop_duplicates(
        subset=[
            "_LLAVE_UNICA"
        ],
        keep="last"
    )

    return trabajo.drop(
        columns=[
            "_LLAVE_UNICA"
        ],
        errors="ignore"
    )


def escribir_hoja_urgencias(
    spreadsheet_id,
    nombre_hoja,
    df
):

    valores = preparar_dataframe_para_sheets(
        df
    )

    sheets_service.spreadsheets().values().clear(
        spreadsheetId=spreadsheet_id,
        range=f"{nombre_hoja}!A:ZZ"
    ).execute()

    sheets_service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range=f"{nombre_hoja}!A1",
        valueInputOption="RAW",
        body={
            "values": valores
        }
    ).execute()


def actualizar_base_urgencias_drive(
    ordenes_nuevas,
    sin_orden_nuevas,
    sin_orden_a_eliminar=None
):

    archivo = obtener_o_crear_google_sheet_urgencias()
    spreadsheet_id = archivo["id"]

    asegurar_hojas_google_sheet_urgencias(
        spreadsheet_id
    )

    ordenes_previas = leer_hoja_urgencias(
        spreadsheet_id,
        NOMBRE_HOJA_URGENCIAS_ORDENES
    )

    sin_orden_previas = leer_hoja_urgencias(
        spreadsheet_id,
        NOMBRE_HOJA_URGENCIAS_SIN_ORDEN
    )

    ordenes_nuevas = ordenes_nuevas.copy() if isinstance(
        ordenes_nuevas,
        pd.DataFrame
    ) else pd.DataFrame()

    sin_orden_nuevas = sin_orden_nuevas.copy() if isinstance(
        sin_orden_nuevas,
        pd.DataFrame
    ) else pd.DataFrame()

    sin_orden_a_eliminar = sin_orden_a_eliminar or []

    if not ordenes_nuevas.empty:

        ordenes_nuevas = ordenes_nuevas.reindex(
            columns=COLUMNAS_BASE_URGENCIAS_ORDENES
        )

    if not sin_orden_nuevas.empty:

        sin_orden_nuevas = sin_orden_nuevas.reindex(
            columns=COLUMNAS_BASE_URGENCIAS_SIN_ORDEN
        )

    ordenes = pd.concat(
        [
            ordenes_previas,
            ordenes_nuevas
        ],
        ignore_index=True
    )

    sin_orden = pd.concat(
        [
            sin_orden_previas,
            sin_orden_nuevas
        ],
        ignore_index=True
    )

    ids_a_eliminar = {
        obtener_id_urgencia(registro)
        for registro in sin_orden_a_eliminar
        if isinstance(registro, dict)
        and obtener_id_urgencia(registro)
    }

    llaves_a_eliminar = {
        (
            compactar_clave_urgencia(registro.get("CLAVE_CNIS", "")),
            normalizar_entidad_urgencia(registro.get("ENTIDAD", ""))
        )
        for registro in sin_orden_a_eliminar
        if isinstance(registro, dict)
    }

    if not sin_orden.empty and (ids_a_eliminar or llaves_a_eliminar):

        def es_revision_actualizada(fila):

            identificador = limpiar_valor_visual(
                obtener_valor(
                    fila,
                    [
                        "ID_URGENCIA",
                        "ID_INCIDENCIA",
                        "ID",
                        "id"
                    ]
                )
            )

            if identificador and identificador in ids_a_eliminar:

                return True

            llave = (
                compactar_clave_urgencia(
                    obtener_valor(
                        fila,
                        [
                            "CLAVE_CNIS",
                            "clave_cnis",
                            "CLAVE",
                            "clave"
                        ]
                    )
                ),
                normalizar_entidad_urgencia(
                    obtener_valor(
                        fila,
                        [
                            "ENTIDAD",
                            "entidad"
                        ]
                    )
                )
            )

            return (
                not identificador
                and llave in llaves_a_eliminar
            )

        sin_orden = sin_orden.loc[
            ~sin_orden.apply(
                es_revision_actualizada,
                axis=1
            )
        ].copy()

    ordenes = deduplicar_base_urgencias(
        ordenes,
        [
            "FECHA_URGENCIA",
            "CLAVE_CNIS",
            "ENTIDAD",
            "ORDEN_SUMINISTRO"
        ]
    )

    sin_orden = deduplicar_base_urgencias(
        sin_orden,
        [
            "FECHA_URGENCIA",
            "CLAVE_CNIS",
            "ENTIDAD"
        ]
    )

    if not ordenes.empty:

        ordenes = ordenes.reindex(
            columns=COLUMNAS_BASE_URGENCIAS_ORDENES
        )

    if not sin_orden.empty:

        sin_orden = sin_orden.reindex(
            columns=COLUMNAS_BASE_URGENCIAS_SIN_ORDEN
        )

    escribir_hoja_urgencias(
        spreadsheet_id,
        NOMBRE_HOJA_URGENCIAS_ORDENES,
        ordenes
    )

    escribir_hoja_urgencias(
        spreadsheet_id,
        NOMBRE_HOJA_URGENCIAS_SIN_ORDEN,
        sin_orden
    )

    return {
        "link": archivo.get(
            "webViewLink",
            f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}"
        ),
        "ordenes": len(ordenes),
        "sin_orden": len(sin_orden)
    }


def obtener_datos_orden_para_registro(
    valor_busqueda
):

    resultado = buscar_orden_fuerte(
        valor_busqueda
    )

    if len(resultado) == 0:

        return None, resultado

    resultado = resultado.copy()

    fila = resultado.iloc[
        0
    ]

    estatus_base = obtener_valor(
        fila,
        [
            "estatus_base"
        ]
    )

    origen_compendio = obtener_valor(
        fila,
        [
            "origen_compendio"
        ]
    )

    orden = obtener_valor(
        fila,
        [
            "orden_suministro",
            "orden",
            "no_orden"
        ]
    )

    tipo_entrega = obtener_valor(
        fila,
        [
            "tipo_entrega"
        ]
    )

    entidad = obtener_valor(
        fila,
        [
            "entidad",
            "estado"
        ]
    )

    clues_destino = obtener_valor(
        fila,
        [
            "clues_destino"
        ]
    )

    unidad_destino = obtener_valor(
        fila,
        [
            "unidad_destino"
        ]
    )

    almacen_original = obtener_valor(
        fila,
        [
            "almacen"
        ]
    )

    almacen = construir_almacen(
        clues_destino,
        unidad_destino,
        almacen_original
    )

    proveedor = obtener_valor(
        fila,
        [
            "proveedor"
        ]
    )

    clave = obtener_valor(
        fila,
        [
            "clave_cnis"
        ]
    )

    descripcion = obtener_valor(
        fila,
        [
            "descripcion"
        ]
    )

    piezas_emitidas = obtener_valor(
        fila,
        [
            "piezas_emitidas"
        ]
    )

    piezas_recibidas_ol = obtener_valor(
        fila,
        [
            "piezas_recibidas_ol"
        ]
    )

    piezas_entregadas = obtener_valor(
        fila,
        [
            "piezas_entregadas_clues"
        ]
    )

    operador = obtener_valor(
        fila,
        [
            "operador_logistico"
        ]
    )

    tipo_red = obtener_valor(
        fila,
        [
            "tipo_red"
        ]
    )

    grupo_terapeutico = obtener_valor(
        fila,
        [
            "grupo_terapeutico"
        ]
    )

    estatus_orden = obtener_valor(
        fila,
        [
            "estatus"
        ]
    )

    tiene_ol = es_operador_logistico(
        tipo_entrega
    )

    estatus_recepcion_ol = ""

    if tiene_ol:

        estatus_recepcion_ol = calcular_estatus_piezas(
            piezas_emitidas,
            piezas_recibidas_ol
        )

    estatus_entrega_estado = calcular_estatus_piezas(
        piezas_emitidas,
        piezas_entregadas
    )

    estatus_completa = calcular_estatus_incidencia_completa(
        estatus_entrega_estado
    )

    datos = {
        "resultado": resultado,
        "fila": fila,
        "estatus_base": estatus_base,
        "origen_compendio": origen_compendio,
        "orden": orden,
        "tipo_entrega": tipo_entrega,
        "entidad": entidad,
        "clues_destino": clues_destino,
        "unidad_destino": unidad_destino,
        "almacen": almacen,
        "proveedor": proveedor,
        "clave": clave,
        "descripcion": descripcion,
        "piezas_emitidas": piezas_emitidas,
        "piezas_recibidas_ol": piezas_recibidas_ol,
        "piezas_entregadas": piezas_entregadas,
        "operador": operador,
        "tipo_red": tipo_red,
        "grupo_terapeutico": grupo_terapeutico,
        "estatus_orden": estatus_orden,
        "estatus_recepcion_ol": estatus_recepcion_ol,
        "estatus_entrega_estado": estatus_entrega_estado,
        "estatus_completa": estatus_completa
    }

    return datos, resultado


def calcular_estatus_seguimiento_desde_datos(datos):

    estatus_base = str(
        datos.get(
            "estatus_base",
            ""
        )
    ).upper().strip()

    piezas_entregadas = convertir_numero(
        datos.get(
            "piezas_entregadas",
            0
        )
    )

    if piezas_entregadas > 0:

        return "Completo-Entregado"

    if estatus_base in [
        "INACTIVA",
        "CANCELADA",
        "CANCELADO",
        "INACTIVO"
    ]:

        return "Incompleta-Cancelada"

    return "Incompleta-sin Entregar"


def calcular_estatus_seguimiento_desde_compendio(fila_compendio):

    estatus_base = str(
        fila_compendio.get(
            "estatus_base",
            ""
        )
    ).upper().strip()

    piezas_entregadas = convertir_numero(
        fila_compendio.get(
            "piezas_entregadas_clues",
            0
        )
    )

    if piezas_entregadas > 0:

        return "Completo-Entregado"

    if estatus_base in [
        "INACTIVA",
        "CANCELADA",
        "CANCELADO",
        "INACTIVO"
    ]:

        return "Incompleta-Cancelada"

    return "Incompleta-sin Entregar"


def actualizar_urgencias_revision_con_emision(limite=10000):

    incidencias_actuales = cargar_urgencias_para_registro()

    if incidencias_actuales.empty:

        return {
            "actualizadas": 0,
            "sin_emision": 0,
            "errores": [],
            "drive": None
        }

    actualizadas = 0
    sin_emision = 0
    errores = []
    filas_base_ordenes = []
    revisiones_actualizadas = []

    for _, pendiente in incidencias_actuales.iterrows():

        if actualizadas >= limite:

            break

        tipo = limpiar_valor_visual(
            obtener_valor(
                pendiente,
                [
                    "TIPO_INCIDENCIA",
                    "tipo_incidencia"
                ]
            )
        ).upper()

        if tipo != "URGENCIA":

            continue

        orden_actual = limpiar_valor_visual(
            obtener_valor(
                pendiente,
                [
                    "ORDEN",
                    "orden",
                    "orden_suministro",
                    "ORDEN_BUSCADA"
                ]
            )
        )

        if orden_actual:

            continue

        estatus = " ".join(
            limpiar_valor_visual(
                obtener_valor(
                    pendiente,
                    [
                        columna
                    ]
                )
            )
            for columna in [
                "ESTATUS_OPERATIVO",
                "ESTATUS_BASE",
                "ESTATUS_SEGUIMIENTO",
                "OBSERVACIONES"
            ]
        )

        estatus_normalizado = normalizar_entidad_urgencia(
            estatus
        )

        if (
            "REVISAR CON EMISION" not in estatus_normalizado
            and "SIN ORDEN" not in estatus_normalizado
        ):

            continue

        clave = limpiar_valor_visual(
            obtener_valor(
                pendiente,
                [
                    "CLAVE_CNIS",
                    "clave_cnis",
                    "CLAVE",
                    "clave"
                ]
            )
        )

        entidad = limpiar_valor_visual(
            obtener_valor(
                pendiente,
                [
                    "ENTIDAD",
                    "entidad"
                ]
            )
        )

        if not clave or not entidad:

            continue

        try:

            resultado = buscar_ordenes_urgencia(
                clave,
                entidad
            )

            ordenes_emitidas = resultado.get(
                "ordenes_emitidas",
                pd.DataFrame()
            )

            if ordenes_emitidas.empty:

                sin_emision += 1
                continue

            ordenes_no_entregadas = resultado.get(
                "ordenes",
                pd.DataFrame()
            )

            candidatos = (
                ordenes_no_entregadas
                if not ordenes_no_entregadas.empty
                else ordenes_emitidas
            )

            registro = candidatos.iloc[0].to_dict()
            fecha = obtener_fecha_urgencia_registro(
                pendiente
            )
            responsable = limpiar_valor_visual(
                obtener_valor(
                    pendiente,
                    [
                        "RESPONSABLE",
                        "responsable"
                    ]
                )
            )
            pdf_link = limpiar_valor_visual(
                obtener_valor(
                    pendiente,
                    [
                        "PDF_CORREO_SEGUIMIENTO",
                        "pdf_correo_seguimiento"
                    ]
                )
            )

            actualizacion = actualizar_revision_urgencia_con_orden(
                pendiente,
                registro,
                fecha,
                responsable,
                "",
                pdf_link
            )

            actualizadas += 1
            revisiones_actualizadas.append(
                pendiente.to_dict()
            )
            filas_base_ordenes.append(
                actualizacion["fila_base"]
            )

        except Exception as e:

            errores.append(
                f"{clave} / {entidad}: {e}"
            )

    resultado_drive = None

    if filas_base_ordenes:

        try:

            resultado_drive = actualizar_base_urgencias_drive(
                pd.DataFrame(
                    filas_base_ordenes,
                    columns=COLUMNAS_BASE_URGENCIAS_ORDENES
                ),
                pd.DataFrame(
                    columns=COLUMNAS_BASE_URGENCIAS_SIN_ORDEN
                ),
                revisiones_actualizadas
            )

        except Exception as e:

            errores.append(
                f"Base de Urgencias en Drive: {e}"
            )

    if actualizadas:

        st.cache_data.clear()

    return {
        "actualizadas": actualizadas,
        "sin_emision": sin_emision,
        "errores": errores,
        "drive": resultado_drive
    }


def actualizar_estatus_seguimiento_con_compendio(limite=10000):

    incidencias_actuales = cargar_incidencias()

    if incidencias_actuales.empty:

        return {
            "actualizadas": 0,
            "sin_compendio": 0,
            "errores": []
        }

    actualizadas = 0
    sin_compendio = 0
    errores = []

    for _, fila in incidencias_actuales.iterrows():

        if actualizadas >= limite:

            break

        if limpiar_valor_visual(
            obtener_valor(
                fila,
                ["TIPO_INCIDENCIA", "tipo_incidencia"]
            )
        ).upper() == "URGENCIA":
            continue

        incidencia_id = fila.get(
            "ID",
            ""
        )

        orden = obtener_valor(
            fila,
            [
                "orden_suministro",
                "ORDEN",
                "ORDEN_BUSCADA"
            ]
        )

        if not incidencia_id or not orden:

            continue

        try:

            compendio = buscar_orden_compendio_rapida(
                orden
            )

            if compendio.empty:

                sin_compendio += 1
                continue

            fila_compendio = compendio.iloc[0]

            estatus_nuevo = calcular_estatus_seguimiento_desde_compendio(
                fila_compendio
            )

            supabase.table(
                "incidencias"
            ).update(
                {
                    "estatus_seguimiento": estatus_nuevo,
                    "estatus_base": preparar_valor_supabase(
                        fila_compendio.get(
                            "estatus_base",
                            ""
                        )
                    ),
                    "piezas_entregadas_clues": preparar_valor_supabase(
                        fila_compendio.get(
                            "piezas_entregadas_clues",
                            ""
                        )
                    )
                }
            ).eq(
                "id",
                incidencia_id
            ).execute()

            actualizadas += 1

        except Exception as e:

            errores.append(
                f"{orden}: {e}"
            )

    st.cache_data.clear()

    resultado_revisiones = actualizar_urgencias_revision_con_emision(
        limite=limite
    )

    actualizadas_revisiones = resultado_revisiones.get(
        "actualizadas",
        0
    )

    actualizadas += actualizadas_revisiones
    errores.extend(
        resultado_revisiones.get(
            "errores",
            []
        )
    )

    st.cache_data.clear()

    try:

        actualizar_google_sheets_maestro()

    except Exception as e:

        errores.append(
            f"Google Sheets: {e}"
        )

    return {
        "actualizadas": actualizadas,
        "urgencias_actualizadas": actualizadas_revisiones,
        "sin_compendio": sin_compendio,
        "errores": errores
    }



def fecha_hoy_sistema():

    try:

        return datetime.now(
            ZoneInfo(
                "America/Mexico_City"
            )
        ).strftime(
            "%Y-%m-%d"
        )

    except Exception:

        return datetime.now().strftime(
            "%Y-%m-%d"
        )


def obtener_control_sistema(
    clave
):

    try:

        respuesta = (
            supabase
            .table(
                "control_sistema"
            )
            .select(
                "clave, valor, actualizado_en"
            )
            .eq(
                "clave",
                clave
            )
            .limit(
                1
            )
            .execute()
        )

        if respuesta.data:

            return respuesta.data[0]

    except Exception as e:

        st.warning(
            f"No se pudo leer control_sistema: {e}"
        )

    return None


def guardar_control_sistema(
    clave,
    valor
):

    try:

        supabase.table(
            "control_sistema"
        ).upsert(
            {
                "clave": clave,
                "valor": valor,
                "actualizado_en": datetime.now().isoformat()
            }
        ).execute()

        return True

    except Exception as e:

        st.warning(
            f"No se pudo guardar control_sistema: {e}"
        )

        return False


def seguimiento_actualizado_hoy():

    control = obtener_control_sistema(
        "ultima_actualizacion_seguimiento"
    )

    if not control:

        return False

    return str(
        control.get(
            "valor",
            ""
        )
    ).strip() == fecha_hoy_sistema()


def actualizar_seguimiento_diario_si_corresponde(
    limite=10000,
    forzar=False
):

    hoy = fecha_hoy_sistema()

    if (
        not forzar
        and seguimiento_actualizado_hoy()
    ):

        resultado_revisiones = actualizar_urgencias_revision_con_emision(
            limite=limite
        )

        if resultado_revisiones.get(
            "actualizadas",
            0
        ):

            return {
                "ejecutado": True,
                "actualizadas": resultado_revisiones.get(
                    "actualizadas",
                    0
                ),
                "urgencias_actualizadas": resultado_revisiones.get(
                    "actualizadas",
                    0
                ),
                "sin_compendio": 0,
                "errores": resultado_revisiones.get(
                    "errores",
                    []
                ),
                "mensaje": (
                    "Se actualizaron solicitudes REVISAR CON EMISIÓN "
                    "con la emisión localizada en el compendio."
                )
            }

        return {
            "ejecutado": False,
            "actualizadas": 0,
            "urgencias_actualizadas": 0,
            "sin_compendio": 0,
            "errores": [],
            "mensaje": "El seguimiento completo ya fue actualizado hoy."
        }

    resultado = actualizar_estatus_seguimiento_con_compendio(
        limite=limite
    )

    guardar_control_sistema(
        "ultima_actualizacion_seguimiento",
        hoy
    )

    resultado[
        "ejecutado"
    ] = True

    resultado[
        "mensaje"
    ] = "Seguimiento completo actualizado."

    return resultado

def semaforo_seguimiento(valor):

    texto = str(
        valor
    ).upper().strip()

    if texto == "COMPLETO-ENTREGADO":

        return "🟢 Completo-Entregado"

    if texto == "INCOMPLETA-SIN ENTREGAR":

        return "🟡 Incompleta-sin Entregar"

    if normalizar_entidad_urgencia(texto) == "REVISAR CON EMISION":

        return "🟡 Revisar con emisión"

    if texto == "INCOMPLETA-CANCELADA":

        return "🔴 Incompleta-Cancelada"

    return "⚪ Sin estatus"


def construir_registro_incidencia(
    valor_busqueda,
    datos,
    atribuible,
    tipo,
    estatus,
    responsable,
    observaciones,
    fecha_incidencia=None,
    ruta_cedula="",
    ruta_correo=""
):

    return {
        "FECHA_REGISTRO": datetime.now(),
        "ORIGEN_REGISTRO": "SISTEMA",
        "ORDEN_BUSCADA": valor_busqueda,
        "orden_suministro": datos["orden"],
        "ORDEN": datos["orden"],
        "TIPO_ENTREGA": datos["tipo_entrega"],
        "ENTIDAD": datos["entidad"],
        "ALMACEN_CLUES_DESTINO": datos["almacen"],
        "CLUES_DESTINO": datos["clues_destino"],
        "UNIDAD_DESTINO": datos["unidad_destino"],
        "PROVEEDOR": datos["proveedor"],
        "CLAVE_CNIS": datos["clave"],
        "DESCRIPCION": datos["descripcion"],
        "PIEZAS_EMITIDAS": datos["piezas_emitidas"],
        "PIEZAS_RECIBIDAS_OL": datos["piezas_recibidas_ol"],
        "PIEZAS_ENTREGADAS_CLUES": datos["piezas_entregadas"],
        "TIPO_RED": datos["tipo_red"],
        "GRUPO_TERAPEUTICO": datos["grupo_terapeutico"],
        "ESTATUS_OPERATIVO": datos["estatus_base"],
        "ESTATUS_BASE": datos["estatus_base"],
        "ORIGEN_COMPENDIO": datos["origen_compendio"],
        "OPERADOR_LOGISTICO": datos["operador"],
        "ESTATUS_RECEPCION_OL": datos["estatus_recepcion_ol"],
        "ESTATUS_ENTREGA_ESTADO": datos["estatus_entrega_estado"],
        "ESTATUS_INCIDENCIA_COMPLETA": datos["estatus_completa"],
        "ESTATUS_SEGUIMIENTO": calcular_estatus_seguimiento_desde_datos(datos),
        "ATRIBUIBLE A": atribuible,
        "TIPO_INCIDENCIA": tipo,
        "ESTATUS_INCIDENCIA": estatus,
        "RESPONSABLE": responsable,
        "OBSERVACIONES": agregar_fecha_incidencia_observaciones(
            observaciones, fecha_incidencia or datetime.now().date()
        ),
        "PDF_CEDULA_RECHAZO": ruta_cedula,
        "PDF_CORREO_SEGUIMIENTO": ruta_correo
    }


def preparar_fila_base_urgencia_orden(
    registro,
    fecha,
    responsable,
    incidencia_id,
    observaciones,
    pdf_link=""
):

    fecha_visual = fecha_a_texto(
        fecha
    )

    return {
        "FECHA_URGENCIA": fecha_visual,
        "SOLICITUD": fecha_visual,
        "CLAVE_CNIS": registro.get("CLAVE_CNIS", ""),
        "ENTIDAD": registro.get("ENTIDAD", ""),
        "ENTIDAD_COMPENDIO": registro.get("ENTIDAD_COMPENDIO", ""),
        "ORDEN_SUMINISTRO": registro.get("ORDEN_SUMINISTRO", ""),
        "ESTATUS_BASE": registro.get("ESTATUS_BASE", ""),
        "TIPO_ENTREGA": registro.get("TIPO_ENTREGA", ""),
        "CLUES_DESTINO": registro.get("CLUES_DESTINO", ""),
        "UNIDAD_DESTINO": registro.get("UNIDAD_DESTINO", ""),
        "ALMACEN": registro.get("ALMACEN", ""),
        "PROVEEDOR": registro.get("PROVEEDOR", ""),
        "DESCRIPCION": registro.get("DESCRIPCION", ""),
        "PIEZAS_EMITIDAS": registro.get("PIEZAS_EMITIDAS", ""),
        "PIEZAS_ENTREGADAS_CLUES": registro.get(
            "PIEZAS_ENTREGADAS_CLUES",
            ""
        ),
        "PENDIENTE_ENTREGA": registro.get(
            "PENDIENTE_ENTREGA",
            ""
        ),
        "ESTATUS_ENTREGA_ENTIDAD": registro.get(
            "ESTATUS_ENTREGA_ENTIDAD",
            "NO ENTREGADA"
        ),
        "TIPO_RED": registro.get("TIPO_RED", ""),
        "GRUPO_TERAPEUTICO": registro.get("GRUPO_TERAPEUTICO", ""),
        "OPERADOR_LOGISTICO": registro.get("OPERADOR_LOGISTICO", ""),
        "ORIGEN_COMPENDIO": registro.get("ORIGEN_COMPENDIO", ""),
        "RESPONSABLE": responsable,
        "FECHA_REGISTRO": fecha_a_texto(
            datetime.now()
        ),
        "ID_URGENCIA": incidencia_id,
        "ID_INCIDENCIA": "",
        "OBSERVACIONES": observaciones,
        "PDF_CORREO_SEGUIMIENTO": pdf_link,
        "FECHA_ACTUALIZACION_BASE": fecha_a_texto(
            datetime.now()
        )
    }


def preparar_fila_base_urgencia_sin_orden(
    clave,
    entidad,
    fecha,
    responsable,
    incidencia_id,
    observaciones,
    pdf_link=""
):

    fecha_visual = fecha_a_texto(
        fecha
    )

    return {
        "FECHA_URGENCIA": fecha_visual,
        "SOLICITUD": fecha_visual,
        "CLAVE_CNIS": clave,
        "ENTIDAD": entidad,
        "ESTATUS": "REVISAR CON EMISIÓN",
        "RESPONSABLE": responsable,
        "FECHA_REGISTRO": fecha_a_texto(
            datetime.now()
        ),
        "ID_URGENCIA": incidencia_id,
        "ID_INCIDENCIA": "",
        "MOTIVO": "REVISAR CON EMISIÓN",
        "OBSERVACIONES": observaciones,
        "PDF_CORREO_SEGUIMIENTO": pdf_link,
        "FECHA_ACTUALIZACION_BASE": fecha_a_texto(
            datetime.now()
        )
    }


def actualizar_revision_urgencia_con_orden(
    pendiente,
    registro,
    fecha,
    responsable,
    observaciones="",
    pdf_link=""
):

    pdf_previo = limpiar_valor_visual(
        obtener_valor(
            pendiente,
            [
                "PDF_CORREO_SEGUIMIENTO",
                "pdf_correo_seguimiento"
            ]
        )
    )

    pdf_final = pdf_link or pdf_previo
    observaciones_finales = observaciones_urgencia(
        fecha,
        observaciones
    )
    observaciones_previas = limpiar_valor_visual(
        obtener_valor(
            pendiente,
            [
                "OBSERVACIONES",
                "observaciones"
            ]
        )
    )

    if observaciones_previas:
        observaciones_finales = (
            f"{observaciones_previas}\n"
            "Actualizada con emisión localizada en el compendio.\n"
            f"{observaciones_finales}"
        ).strip()

    urgencia_id = obtener_id_urgencia(pendiente) or generar_id_urgencia()
    fila_base = preparar_fila_base_urgencia_orden(
        registro,
        fecha,
        responsable,
        urgencia_id,
        observaciones_finales,
        pdf_final
    )

    fecha_registro_previa = obtener_valor(
        pendiente,
        [
            "FECHA_REGISTRO",
            "fecha_registro"
        ]
    )
    if limpiar_valor_visual(fecha_registro_previa):
        fila_base["FECHA_REGISTRO"] = fecha_registro_previa

    # Las urgencias se actualizan únicamente en BASE_URGENCIAS. La clave
    # ``incidencia_id`` se conserva en la respuesta por compatibilidad con
    # el flujo anterior, pero ya no representa un registro en incidencias.
    return {
        "incidencia_id": urgencia_id,
        "urgencia_id": urgencia_id,
        "nueva": fila_base.copy(),
        "fila_base": fila_base
    }


def registrar_urgencia_completa(
    clave,
    entidad,
    fecha,
    responsable,
    observaciones,
    resultado_busqueda,
    incidencias_actuales,
    archivo_pdf=None
):

    # Las urgencias se consultan y deduplican desde BASE_URGENCIAS. El
    # parámetro histórico se conserva sólo para compatibilidad con las
    # llamadas existentes; no se utiliza para insertar en incidencias.
    incidencias_actuales = cargar_urgencias_para_registro()

    ordenes = resultado_busqueda.get(
        "ordenes",
        pd.DataFrame()
    )

    ordenes_emitidas = resultado_busqueda.get(
        "ordenes_emitidas",
        pd.DataFrame()
    )

    sin_orden = bool(
        resultado_busqueda.get(
            "sin_orden",
            False
        )
    )

    if not resultado_busqueda.get(
        "clave_valida",
        True
    ):

        return {
            "guardadas": 0,
            "duplicadas": 0,
            "errores": [
                "La clave no fue localizada en la base del sistema. "
                "No se puede registrar la urgencia."
            ],
            "drive": None,
            "drive_incidencias": "",
            "ordenes_no_entregadas": 0,
            "sin_orden": 0
        }

    entidad_homologada = resultado_busqueda.get(
        "entidad_homologada",
        ""
    ) or entidad

    guardadas = 0
    actualizadas = 0
    duplicadas = 0
    errores = []
    filas_base_ordenes = []
    filas_base_sin_orden = []
    revisiones_actualizadas = []
    ids_revision_utilizados = set()
    pdf_link = ""
    pdf_intentado = False

    if not ordenes.empty:

        for _, fila in ordenes.iterrows():

            registro = fila.to_dict()
            orden = registro.get(
                "ORDEN_SUMINISTRO",
                ""
            )

            if urgencia_ya_registrada(
                incidencias_actuales,
                clave,
                entidad,
                fecha,
                orden
            ):

                duplicadas += 1

                filas_base_ordenes.append(
                    preparar_fila_base_urgencia_orden(
                        registro,
                        fecha,
                        responsable,
                        "",
                        observaciones_urgencia(
                            fecha,
                            observaciones
                        ),
                        pdf_link
                    )
                )

                continue

            pendiente = obtener_urgencia_revision_pendiente(
                incidencias_actuales,
                clave,
                entidad,
                ids_revision_utilizados
            )

            if pendiente:

                try:

                    if archivo_pdf is not None and not pdf_intentado:

                        pdf_intentado = True
                        pdf_link = subir_pdf_evidencia_drive(
                            archivo_pdf,
                            orden or clave,
                            "urgencia",
                            entidad_homologada,
                            registro.get("CLUES_DESTINO", "") or "SIN_CLUES"
                        )

                    actualizacion = actualizar_revision_urgencia_con_orden(
                        pendiente,
                        registro,
                        fecha,
                        responsable,
                        observaciones,
                        pdf_link
                    )

                    urgencia_id = actualizacion["urgencia_id"]

                    actualizadas += 1
                    guardadas += 1
                    ids_revision_utilizados.add(
                        urgencia_id
                    )
                    revisiones_actualizadas.append(
                        pendiente.to_dict()
                    )

                    filas_base_ordenes.append(
                        actualizacion["fila_base"]
                    )

                    continue

                except Exception as e:

                    errores.append(
                        f"Actualización {orden}: {e}"
                    )

            try:

                if archivo_pdf is not None and not pdf_intentado:

                    pdf_intentado = True
                    pdf_link = subir_pdf_evidencia_drive(
                        archivo_pdf,
                        orden or clave,
                        "urgencia",
                        entidad_homologada,
                        registro.get("CLUES_DESTINO", "") or "SIN_CLUES"
                    )

                urgencia_id = generar_id_urgencia()

                guardadas += 1

                filas_base_ordenes.append(
                    preparar_fila_base_urgencia_orden(
                        registro,
                        fecha,
                        responsable,
                        urgencia_id,
                        observaciones_urgencia(
                            fecha,
                            observaciones
                        ),
                        pdf_link
                    )
                )

            except Exception as e:

                errores.append(
                    f"{orden}: {e}"
                )

    elif not ordenes_emitidas.empty:

        pendiente = obtener_urgencia_revision_pendiente(
            incidencias_actuales,
            clave,
            entidad,
            ids_revision_utilizados
        )

        if pendiente:

            orden = clave

            try:

                registro = ordenes_emitidas.iloc[0].to_dict()
                orden = registro.get(
                    "ORDEN_SUMINISTRO",
                    ""
                )

                if archivo_pdf is not None and not pdf_intentado:

                    pdf_intentado = True
                    pdf_link = subir_pdf_evidencia_drive(
                        archivo_pdf,
                        orden or clave,
                        "urgencia",
                        entidad_homologada,
                        registro.get("CLUES_DESTINO", "") or "SIN_CLUES"
                    )

                actualizacion = actualizar_revision_urgencia_con_orden(
                    pendiente,
                    registro,
                    fecha,
                    responsable,
                    observaciones,
                    pdf_link
                )

                incidencia_id = actualizacion["incidencia_id"]
                actualizadas += 1
                guardadas += 1
                ids_revision_utilizados.add(
                    incidencia_id
                )
                revisiones_actualizadas.append(
                    pendiente
                )
                filas_base_ordenes.append(
                    actualizacion["fila_base"]
                )

            except Exception as e:

                errores.append(
                    f"Actualización {orden or clave}: {e}"
                )

    elif sin_orden:

        if urgencia_ya_registrada(
            incidencias_actuales,
            clave,
            entidad,
            fecha,
            ""
        ):

            duplicadas += 1

            filas_base_sin_orden.append(
                preparar_fila_base_urgencia_sin_orden(
                    clave,
                    entidad_homologada,
                    fecha,
                    responsable,
                    "",
                    observaciones_urgencia(
                        fecha,
                        observaciones
                    ),
                    pdf_link
                )
            )

        else:

            try:

                if archivo_pdf is not None and not pdf_intentado:

                    pdf_intentado = True
                    pdf_link = subir_pdf_evidencia_drive(
                        archivo_pdf,
                        clave,
                        "urgencia",
                        entidad_homologada,
                        "SIN_CLUES"
                    )

                urgencia_id = generar_id_urgencia()

                guardadas += 1

                filas_base_sin_orden.append(
                    preparar_fila_base_urgencia_sin_orden(
                        clave,
                        entidad_homologada,
                        fecha,
                        responsable,
                        urgencia_id,
                        observaciones_urgencia(
                            fecha,
                            observaciones
                        ),
                        pdf_link
                    )
                )

            except Exception as e:

                errores.append(
                    f"{clave} / {entidad}: {e}"
                )

    resultado_drive = None

    if filas_base_ordenes or filas_base_sin_orden:

        try:

            resultado_drive = actualizar_base_urgencias_drive(
                pd.DataFrame(
                    filas_base_ordenes,
                    columns=COLUMNAS_BASE_URGENCIAS_ORDENES
                ),
                pd.DataFrame(
                    filas_base_sin_orden,
                    columns=COLUMNAS_BASE_URGENCIAS_SIN_ORDEN
                ),
                revisiones_actualizadas
            )

        except Exception as e:

            errores.append(
                f"Base de Urgencias en Drive: {e}"
            )

    return {
        "guardadas": guardadas,
        "actualizadas": actualizadas,
        "duplicadas": duplicadas,
        "errores": errores,
        "drive": resultado_drive,
        "drive_incidencias": "",
        "ordenes_no_entregadas": len(filas_base_ordenes),
        "sin_orden": len(filas_base_sin_orden)
    }


def migrar_urgencias_fuera_de_incidencias():
    """Inhabilitada: las fuentes son independientes y no se migran registros."""
    return {"detectadas": 0, "migradas": 0, "eliminadas": 0,
            "errores": ["Migración entre módulos deshabilitada."], "drive": None}


def extraer_ordenes_masivas(
    texto
):

    ordenes = []

    for linea in str(texto).splitlines():

        linea = linea.strip()

        if not linea:

            continue

        partes = re.split(
            r"[,;\t]+",
            linea
        )

        for parte in partes:

            orden = normalizar_orden(
                parte
            )

            if orden and orden not in ordenes:

                ordenes.append(
                    orden
                )

    return ordenes


def extraer_claves_urgencia_masivas(
    texto
):

    claves = []

    for linea in str(texto).splitlines():

        linea = linea.strip()

        if not linea:

            continue

        partes = re.split(
            r"[,;\t]+",
            linea
        )

        for parte in partes:

            clave = normalizar_clave_urgencia(
                parte
            )

            if clave and clave not in claves:

                claves.append(
                    clave
                )

    return claves




# =========================
# REPORTE EJECUTIVO PDF
# =========================

def extraer_monitora(responsable):

    texto = limpiar_valor_visual(responsable)

    if "|" in texto:

        return texto.split("|")[-1].strip()

    return texto


def preparar_df_reporte_pdf(df):

    if df is None or df.empty:

        return pd.DataFrame()

    limpio = limpiar_df_visual(df).copy()

    for columna in limpio.columns:

        limpio[columna] = limpio[columna].apply(
            limpiar_valor_visual
        )

    return limpio


def tabla_pdf(titulo, df, max_filas=12, anchos=None):

    elementos = []

    styles = getSampleStyleSheet()

    estilo_titulo = ParagraphStyle(
        "TituloTablaReporte",
        parent=styles["Heading3"],
        textColor=colors.HexColor(COLOR_VERDE),
        fontSize=10,
        spaceAfter=5,
        leading=12
    )

    estilo_celda = ParagraphStyle(
        "CeldaReporte",
        parent=styles["Normal"],
        fontSize=6.2,
        leading=7,
        alignment=1
    )

    estilo_header = ParagraphStyle(
        "HeaderReporte",
        parent=styles["Normal"],
        fontSize=6.5,
        leading=7,
        alignment=1,
        textColor=colors.white
    )

    elementos.append(
        Paragraph(titulo, estilo_titulo)
    )

    if df is None or df.empty:

        elementos.append(
            Paragraph("Sin registros.", styles["Normal"])
        )

        elementos.append(
            Spacer(1, 0.25 * cm)
        )

        return elementos

    df_tabla = preparar_df_reporte_pdf(
        df
    ).head(
        max_filas
    ).copy()

    encabezados = [
        Paragraph(str(col), estilo_header)
        for col in df_tabla.columns
    ]

    datos = [
        encabezados
    ]

    for _, fila in df_tabla.iterrows():

        datos.append(
            [
                Paragraph(
                    limpiar_valor_visual(valor),
                    estilo_celda
                )
                for valor in fila.tolist()
            ]
        )

    if anchos is None:

        ancho_total = 25 * cm
        columnas = max(
            len(df_tabla.columns),
            1
        )
        anchos = [
            ancho_total / columnas
        ] * columnas

    tabla = Table(
        datos,
        repeatRows=1,
        colWidths=anchos
    )

    tabla.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor(COLOR_VERDE)),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("GRID", (0, 0), (-1, -1), 0.45, colors.black),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (-1, -1), 6.2),
                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("LEFTPADDING", (0, 0), (-1, -1), 2),
                ("RIGHTPADDING", (0, 0), (-1, -1), 2),
                ("TOPPADDING", (0, 0), (-1, -1), 2),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
            ]
        )
    )

    elementos.append(tabla)
    elementos.append(Spacer(1, 0.35 * cm))

    return elementos


def agregar_encabezado_reporte(canvas, doc):

    canvas.saveState()

    ancho, alto = landscape(letter)

    canvas.setStrokeColor(
        colors.HexColor(COLOR_VERDE)
    )

    canvas.setLineWidth(2)

    canvas.line(
        1 * cm,
        alto - 1.05 * cm,
        ancho - 1 * cm,
        alto - 1.05 * cm
    )

    canvas.setFont(
        "Helvetica",
        7
    )

    canvas.drawRightString(
        ancho - 1 * cm,
        0.55 * cm,
        f"Página {doc.page}"
    )

    canvas.restoreState()


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
        errors="coerce"
    )

    df["MONITORA"] = df["RESPONSABLE"].apply(
        extraer_monitora
    )

    df_hoy = df[
        df["FECHA_REGISTRO_DT"].dt.strftime("%Y-%m-%d") == hoy
    ].copy()

    total_hoy = len(df_hoy)

    monitoras_hoy = df_hoy["MONITORA"].replace(
        "",
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
            ["Monitoras que capturaron", monitoras_hoy],
            ["Entidades atendidas", entidades_hoy],
            ["Órdenes registradas", ordenes_hoy],
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

    if df_hoy.empty:

        elementos.append(
            Paragraph(
                "No se registraron capturas durante el día del reporte.",
                normal
            )
        )

        elementos.append(
            Spacer(1, 0.3 * cm)
        )

    else:

        por_monitora_hoy = (
            df_hoy
            .groupby(
                "MONITORA",
                dropna=False
            )
            .size()
            .reset_index(
                name="Capturas"
            )
            .sort_values(
                "Capturas",
                ascending=False
            )
        )

        elementos += tabla_pdf(
            "Capturas del día por monitora",
            por_monitora_hoy,
            max_filas=10,
            anchos=[
                12 * cm,
                4 * cm
            ]
        )

        por_entidad_hoy = (
            df_hoy
            .groupby(
                "ENTIDAD",
                dropna=False
            )
            .size()
            .reset_index(
                name="Incidencias"
            )
            .sort_values(
                "Incidencias",
                ascending=False
            )
        )

        elementos += tabla_pdf(
            "Entidades trabajadas hoy",
            por_entidad_hoy,
            max_filas=10,
            anchos=[
                10 * cm,
                4 * cm
            ]
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
            "Detalle breve del día",
            detalle_hoy,
            max_filas=15,
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

    elementos.append(
        Paragraph(
            "2. Resumen general acumulado",
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
        "",
        pd.NA
    ).dropna().nunique()

    indicadores_general = pd.DataFrame(
        [
            ["Total acumulado", total_general],
            ["Resueltas", resueltas],
            ["En proceso", en_proceso],
            ["Rechazadas", rechazadas],
            ["Monitoras con capturas", monitoras_general],
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

    por_monitora_general = (
        df
        .groupby(
            "MONITORA",
            dropna=False
        )
        .size()
        .reset_index(
            name="Capturas"
        )
        .sort_values(
            "Capturas",
            ascending=False
        )
    )

    elementos += tabla_pdf(
        "Capturas acumuladas por monitora",
        por_monitora_general,
        max_filas=12,
        anchos=[
            12 * cm,
            4 * cm
        ]
    )

    por_tipo = (
        df
        .groupby(
            "TIPO_INCIDENCIA",
            dropna=False
        )
        .size()
        .reset_index(
            name="Total"
        )
        .sort_values(
            "Total",
            ascending=False
        )
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
        .groupby(
            "ENTIDAD",
            dropna=False
        )
        .size()
        .reset_index(
            name="Total"
        )
        .sort_values(
            "Total",
            ascending=False
        )
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

    doc.build(
        elementos,
        onFirstPage=agregar_encabezado_reporte,
        onLaterPages=agregar_encabezado_reporte
    )

    buffer.seek(0)

    return buffer


