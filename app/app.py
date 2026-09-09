import streamlit as st
import pandas as pd
import numpy as np
import altair as alt
import re
import os
import json
import tempfile
import unicodedata

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
                f"GOOGLE_CREDENTIALS existe, pero no es un JSON válido: {e}"
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

    st.sidebar.write(
        f"Buscando archivo: {nombre_archivo}"
    )

    st.sidebar.write(
        f"Folder ID: {folder_id}"
    )

    archivo = buscar_archivo_drive(
        nombre_archivo,
        folder_id
    )

    st.sidebar.write(
        "Resultado Drive:",
        archivo
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
        "JESUS RICANO MURRIETA"
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

    for col in opciones:

        if col in fila.index:

            valor = fila.get(
                col,
                ""
            )

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

    limpio = df.copy()

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
    nueva
):

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
            ""
        )

    return ""


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


def subir_pdf_evidencia_drive(
    archivo,
    orden,
    tipo_pdf,
    estado,
    clues
):

    if archivo is None:

        return ""

    try:

        carpeta_destino_id = obtener_carpeta_evidencia(
            estado,
            clues
        )

        fecha = datetime.now().strftime(
            "%Y%m%d_%H%M%S"
        )

        nombre_orden = limpiar_nombre_archivo(
            orden
        )

        nombre_archivo = (
            f"{fecha}_{nombre_orden}_{tipo_pdf}.pdf"
        )

        ruta_local = os.path.join(
            TEMP_DIR,
            nombre_archivo
        )

        with open(
            ruta_local,
            "wb"
        ) as f:

            f.write(
                archivo.getbuffer()
            )

        metadata = {
            "name": nombre_archivo,
            "parents": [
                carpeta_destino_id
            ]
        }

        media = MediaFileUpload(
            ruta_local,
            mimetype="application/pdf",
            resumable=True
        )

        nuevo = (
            drive_service.files()
            .create(
                body=metadata,
                media_body=media,
                fields="id, webViewLink",
                supportsAllDrives=True
            )
            .execute()
        )

        link = nuevo.get(
            "webViewLink",
            ""
        )

        st.success(
            f"Evidencia subida correctamente: {tipo_pdf}"
        )

        st.write(
            "Link Drive:",
            link
        )

        return link

    except Exception as e:

        st.error(
            f"No se pudo subir la evidencia a Drive ({tipo_pdf}): {e}"
        )

        st.exception(
            e
        )

        return ""

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



def convertir_excel_cedulas_rechazo(df):
    """Genera un Excel independiente con hipervínculos clicables a las cédulas."""
    salida = BytesIO()
    exportar = df.copy()

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
    "RESPONSABLE",
    "FECHA_REGISTRO",
    "ID_INCIDENCIA",
    "MOTIVO",
    "OBSERVACIONES",
    "PDF_CORREO_SEGUIMIENTO",
    "FECHA_ACTUALIZACION_BASE"
]


def normalizar_clave_urgencia(valor):

    texto = limpiar_valor_visual(valor).upper().strip()

    if re.fullmatch(r"\d+\.0+", texto):

        texto = texto.split(".", 1)[0]

    return re.sub(r"\s+", "", texto)


def variantes_clave_urgencia(valor):

    original = limpiar_valor_visual(valor).strip()

    if not original:

        return []

    normalizada = normalizar_clave_urgencia(original)
    variantes = [
        original,
        original.replace(" ", ""),
        normalizada
    ]

    if re.fullmatch(r"\d+", normalizada):

        variantes.extend(
            [
                f"{normalizada}.0",
                str(int(normalizada))
            ]
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


def consultar_compendio_por_clave_urgencia(clave):

    variantes = variantes_clave_urgencia(clave)

    if not variantes:

        return pd.DataFrame()

    datos = []

    # La columna oficial es clave_cnis; clave sólo se usa como respaldo
    # para instalaciones que conservaron el encabezado anterior.
    for columna in [
        "clave_cnis",
        "clave"
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

                datos.extend(
                    respuesta.data or []
                )

            except Exception:

                # Una columna alternativa puede no existir en el proyecto.
                continue

    if not datos:

        texto_busqueda = normalizar_clave_urgencia(
            clave
        )

        if len(texto_busqueda) >= 3:

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
                        "clave_cnis",
                        f"%{texto_busqueda}%"
                    )
                    .limit(
                        1000
                    )
                    .execute()
                )

                datos.extend(
                    respuesta.data or []
                )

            except Exception:

                pass

    if not datos:

        return pd.DataFrame()

    return pd.DataFrame(
        datos
    ).drop_duplicates(
        ignore_index=True
    )


@st.cache_data(
    ttl=600,
    show_spinner=False
)
def obtener_entidades_compendio_urgencia():

    try:

        respuesta = (
            supabase
            .table(
                "compendio"
            )
            .select(
                "*"
            )
            .limit(
                10000
            )
            .execute()
        )

    except Exception:

        return []

    datos = respuesta.data or []

    if not datos:

        return []

    entidades = {}
    columnas = [
        "entidad",
        "ENTIDAD",
        "estado",
        "ESTADO",
        "entidad_destino",
        "Entidad de destino"
    ]

    for fila in datos:

        entidad = limpiar_valor_visual(
            obtener_valor(
                fila,
                columnas
            )
        )

        clave_entidad = normalizar_entidad_urgencia(
            entidad
        )

        if entidad and clave_entidad and clave_entidad not in entidades:

            entidades[clave_entidad] = entidad

    return sorted(
        entidades.values(),
        key=lambda valor: normalizar_entidad_urgencia(valor)
    )


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

    vacio = pd.DataFrame(
        columns=COLUMNAS_BASE_URGENCIAS_ORDENES
    )

    if compendio.empty:

        return {
            "ordenes": vacio,
            "ordenes_entregadas": pd.DataFrame(),
            "total_clave": 0,
            "total_entidad": 0,
            "total_con_orden": 0,
            "sin_orden": False,
            "clave_valida": False,
            "entidad_homologada": ""
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
            "total_clave": len(compendio),
            "total_entidad": 0,
            "total_con_orden": 0,
            "sin_orden": False,
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
            "total_clave": len(compendio),
            "total_entidad": len(filas_entidad),
            "total_con_orden": 0,
            "sin_orden": True,
            "clave_valida": True,
            "entidad_homologada": entidad_homologada
        }

    elegibles = []
    entregadas = []
    registros_df = pd.DataFrame(
        registros
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
        "total_clave": len(compendio),
        "total_entidad": len(filas_entidad),
        "total_con_orden": len(elegibles) + len(entregadas),
        "sin_orden": len(registros) == 0,
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


def datos_orden_urgencia_para_incidencia(registro):

    return {
        "orden": registro.get("ORDEN_SUMINISTRO", ""),
        "tipo_entrega": registro.get("TIPO_ENTREGA", ""),
        "entidad": registro.get("ENTIDAD", ""),
        "clues_destino": registro.get("CLUES_DESTINO", ""),
        "unidad_destino": registro.get("UNIDAD_DESTINO", ""),
        "almacen": registro.get("ALMACEN", ""),
        "proveedor": registro.get("PROVEEDOR", ""),
        "clave": registro.get("CLAVE_CNIS", ""),
        "descripcion": registro.get("DESCRIPCION", ""),
        "piezas_emitidas": registro.get("PIEZAS_EMITIDAS", ""),
        "piezas_recibidas_ol": registro.get("PIEZAS_RECIBIDAS_OL", ""),
        "piezas_entregadas": registro.get("PIEZAS_ENTREGADAS_CLUES", ""),
        "tipo_red": registro.get("TIPO_RED", ""),
        "grupo_terapeutico": registro.get("GRUPO_TERAPEUTICO", ""),
        "operador": registro.get("OPERADOR_LOGISTICO", ""),
        "estatus_base": registro.get("ESTATUS_BASE", ""),
        "origen_compendio": registro.get("ORIGEN_COMPENDIO", ""),
        "estatus_recepcion_ol": "",
        "estatus_entrega_estado": registro.get(
            "ESTATUS_ENTREGA_ENTIDAD",
            "NO ENTREGADA"
        ),
        "estatus_completa": "INCOMPLETA",
        "estatus_orden": registro.get("ESTATUS_BASE", "")
    }


def construir_registro_urgencia(
    registro_orden,
    fecha,
    responsable,
    observaciones="",
    ruta_pdf=""
):

    registro = construir_registro_incidencia(
        registro_orden.get("ORDEN_SUMINISTRO", ""),
        datos_orden_urgencia_para_incidencia(registro_orden),
        "Por determinar",
        "URGENCIA",
        "En proceso",
        responsable,
        observaciones_urgencia(fecha, observaciones),
        "",
        ruta_pdf
    )

    registro["ORIGEN_REGISTRO"] = "URGENCIAS"

    return registro


def construir_registro_urgencia_sin_orden(
    clave,
    entidad,
    fecha,
    responsable,
    observaciones="",
    ruta_pdf=""
):

    texto_observaciones = (
        "Sin orden de suministro encontrada para la clave y entidad.\n"
        f"{observaciones_urgencia(fecha, observaciones)}"
    )

    return {
        "FECHA_REGISTRO": datetime.now(),
        "ORIGEN_REGISTRO": "URGENCIAS",
        "ORDEN_BUSCADA": "",
        "orden_suministro": "",
        "ORDEN": "",
        "TIPO_ENTREGA": "",
        "ENTIDAD": entidad,
        "ALMACEN_CLUES_DESTINO": "",
        "CLUES_DESTINO": "",
        "UNIDAD_DESTINO": "",
        "PROVEEDOR": "",
        "CLAVE_CNIS": clave,
        "DESCRIPCION": "",
        "PIEZAS_EMITIDAS": "",
        "PIEZAS_RECIBIDAS_OL": "",
        "PIEZAS_ENTREGADAS_CLUES": "0",
        "TIPO_RED": "",
        "GRUPO_TERAPEUTICO": "",
        "ESTATUS_OPERATIVO": "SIN ORDEN",
        "ESTATUS_BASE": "SIN ORDEN",
        "ORIGEN_COMPENDIO": "",
        "OPERADOR_LOGISTICO": "",
        "ESTATUS_RECEPCION_OL": "",
        "ESTATUS_ENTREGA_ESTADO": "NO ENTREGADA",
        "ESTATUS_INCIDENCIA_COMPLETA": "INCOMPLETA",
        "ESTATUS_SEGUIMIENTO": "Incompleta-sin Entregar",
        "ATRIBUIBLE A": "Por determinar",
        "TIPO_INCIDENCIA": "URGENCIA",
        "ESTATUS_INCIDENCIA": "En proceso",
        "RESPONSABLE": responsable,
        "OBSERVACIONES": texto_observaciones,
        "PDF_CEDULA_RECHAZO": "",
        "PDF_CORREO_SEGUIMIENTO": ruta_pdf
    }


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

    clave_norm = normalizar_clave_urgencia(
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

        if normalizar_clave_urgencia(
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
    sin_orden_nuevas
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

    try:

        actualizar_google_sheets_maestro()

    except Exception as e:

        errores.append(
            f"Google Sheets: {e}"
        )

    return {
        "actualizadas": actualizadas,
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

        return {
            "ejecutado": False,
            "actualizadas": 0,
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
        "OBSERVACIONES": observaciones,
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
        "ID_INCIDENCIA": incidencia_id,
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
        "RESPONSABLE": responsable,
        "FECHA_REGISTRO": fecha_a_texto(
            datetime.now()
        ),
        "ID_INCIDENCIA": incidencia_id,
        "MOTIVO": "No se encontró orden para la clave y entidad",
        "OBSERVACIONES": observaciones,
        "PDF_CORREO_SEGUIMIENTO": pdf_link,
        "FECHA_ACTUALIZACION_BASE": fecha_a_texto(
            datetime.now()
        )
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

    ordenes = resultado_busqueda.get(
        "ordenes",
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
                "La clave no existe en el compendio. No se puede registrar la urgencia."
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
    duplicadas = 0
    errores = []
    filas_base_ordenes = []
    filas_base_sin_orden = []
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

                nueva = construir_registro_urgencia(
                    registro,
                    fecha,
                    responsable,
                    observaciones,
                    pdf_link
                )

                incidencia_id = guardar_incidencia(
                    nueva
                )

                guardadas += 1

                filas_base_ordenes.append(
                    preparar_fila_base_urgencia_orden(
                        registro,
                        fecha,
                        responsable,
                        incidencia_id,
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

                nueva = construir_registro_urgencia_sin_orden(
                    clave,
                    entidad_homologada,
                    fecha,
                    responsable,
                    observaciones,
                    pdf_link
                )

                incidencia_id = guardar_incidencia(
                    nueva
                )

                guardadas += 1

                filas_base_sin_orden.append(
                    preparar_fila_base_urgencia_sin_orden(
                        clave,
                        entidad_homologada,
                        fecha,
                        responsable,
                        incidencia_id,
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
    link_incidencias = ""

    if guardadas > 0:

        try:

            link_incidencias = generar_respaldo_drive()

        except Exception as e:

            errores.append(
                f"Respaldo general de incidencias: {e}"
            )

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
                )
            )

        except Exception as e:

            errores.append(
                f"Base de Urgencias en Drive: {e}"
            )

    return {
        "guardadas": guardadas,
        "duplicadas": duplicadas,
        "errores": errores,
        "drive": resultado_drive,
        "drive_incidencias": link_incidencias,
        "ordenes_no_entregadas": len(filas_base_ordenes),
        "sin_orden": len(filas_base_sin_orden)
    }


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


# =========================
# APP
# =========================

st.title(
    "📌 Reporteador de Incidencias 2026"
)

st.caption(
    "Sistema operativo IMSS-BIENESTAR para seguimiento de incidencias."
)

st.sidebar.title(
    "⚙️ Panel de control"
)

st.sidebar.success(
    "Base principal conectada a Supabase."
)
st.sidebar.divider()


# =========================
# MENÚ
# =========================

menu = st.sidebar.radio(
    "Menú",
    [
        "Registrar incidencia",
        "Urgencias",
        "Resumen Ejecutivo",
        "Seguimiento",
        "Cédulas de rechazo",
        "Base Supabase"
    ]
)


# =========================
# CARGA INTELIGENTE SUPABASE
# =========================

incidencias = cargar_incidencias()

if menu == "Seguimiento":

    with st.spinner(
        "Revisando actualización diaria de seguimiento..."
    ):

        resultado_auto = actualizar_seguimiento_diario_si_corresponde(
            limite=10000,
            forzar=False
        )

    if resultado_auto.get(
        "ejecutado",
        False
    ):

        st.success(
            f"Seguimiento actualizado automáticamente: {resultado_auto.get('actualizadas', 0)} registros."
        )

        if resultado_auto.get(
            "errores",
            []
        ):

            with st.expander(
                "Ver errores de actualización"
            ):

                st.write(
                    resultado_auto.get(
                        "errores",
                        []
                    )
                )

        st.cache_data.clear()

        incidencias = cargar_incidencias()

    else:

        st.info(
            "El seguimiento completo ya fue actualizado hoy. Las nuevas incidencias se cruzan al guardarse."
        )

st.sidebar.markdown("---")

total_real_sidebar = contar_incidencias_reales_supabase()

st.sidebar.write(
    "Incidencias Supabase filas:",
    total_real_sidebar or len(incidencias)
)

st.sidebar.caption(
    "Modo rápido: sin CPM, inventario ni agenda de citas"
)


st.sidebar.divider()

st.sidebar.caption(
    f"Incidencias registradas: {(total_real_sidebar or len(incidencias)):,}"
)

st.sidebar.caption(
    "Fuente: Supabase + Google Drive"
)


# =========================
# RESUMEN EJECUTIVO
# =========================

if menu == "Resumen Ejecutivo":

    st.subheader(
        "📊 Resumen Ejecutivo"
    )


    st.markdown(
        "### 📄 Reporte Ejecutivo PDF"
    )

    if st.button(
        "📄 Generar reporte ejecutivo PDF",
        use_container_width=True
    ):

        pdf = generar_reporte_ejecutivo_pdf(
            incidencias
        )

        st.download_button(
            label="⬇️ Descargar reporte ejecutivo",
            data=pdf,
            file_name=f"REPORTE_EJECUTIVO_INCIDENCIAS_{fecha_hoy_sistema()}.pdf",
            mime="application/pdf",
            use_container_width=True
        )

    st.divider()
    df_resumen = incidencias.copy()

    if df_resumen.empty:

        st.info(
            "Aún no hay incidencias registradas."
        )

    else:

        if "ESTATUS_SEGUIMIENTO" not in df_resumen.columns:

            df_resumen["ESTATUS_SEGUIMIENTO"] = ""

        df_resumen["ESTATUS_SEGUIMIENTO"] = (
            df_resumen["ESTATUS_SEGUIMIENTO"]
            .astype(str)
            .str.strip()
            .replace(
                {
                    "": "Sin estatus",
                    "nan": "Sin estatus",
                    "None": "Sin estatus"
                }
            )
        )

        if "SEMAFORO" not in df_resumen.columns:

            df_resumen["SEMAFORO"] = df_resumen[
                "ESTATUS_SEGUIMIENTO"
            ].apply(
                semaforo_seguimiento
            )

        total_real_supabase = contar_incidencias_reales_supabase()

        total = total_real_supabase or len(
            df_resumen
        )

        completo_entregado = df_resumen[
            df_resumen["ESTATUS_SEGUIMIENTO"]
            .astype(str)
            .str.upper()
            .eq("COMPLETO-ENTREGADO")
        ].shape[0]

        incompleta_sin_entregar = df_resumen[
            df_resumen["ESTATUS_SEGUIMIENTO"]
            .astype(str)
            .str.upper()
            .eq("INCOMPLETA-SIN ENTREGAR")
        ].shape[0]

        incompleta_cancelada = df_resumen[
            df_resumen["ESTATUS_SEGUIMIENTO"]
            .astype(str)
            .str.upper()
            .eq("INCOMPLETA-CANCELADA")
        ].shape[0]

        sin_estatus = df_resumen[
            df_resumen["ESTATUS_SEGUIMIENTO"]
            .astype(str)
            .str.upper()
            .isin(
                [
                    "",
                    "SIN ESTATUS",
                    "NAN",
                    "NONE"
                ]
            )
        ].shape[0]

        resueltas = 0

        if "ESTATUS_INCIDENCIA" in df_resumen.columns:

            resueltas = df_resumen[
                df_resumen["ESTATUS_INCIDENCIA"]
                .astype(str)
                .str.upper()
                .isin(
                    [
                        "RESUELTA",
                        "RESUELTO"
                    ]
                )
            ].shape[0]

        porcentaje_entregado = 0

        if total > 0:

            porcentaje_entregado = round(
                (completo_entregado / total) * 100,
                2
            )

        c1, c2, c3, c4, c5 = st.columns(
            5
        )

        c1.metric(
            "Total incidencias",
            f"{total:,}"
        )

        c2.metric(
            "🟢 Completo-Entregado",
            f"{completo_entregado:,}"
        )

        c3.metric(
            "🟡 Incompleta-sin Entregar",
            f"{incompleta_sin_entregar:,}"
        )

        c4.metric(
            "🔴 Incompleta-Cancelada",
            f"{incompleta_cancelada:,}"
        )

        c5.metric(
            "% entregado",
            f"{porcentaje_entregado}%"
        )

        st.divider()

        st.subheader(
            "🚦 Semáforo de seguimiento"
        )

        resumen_seguimiento = (
            df_resumen
            .groupby(
                "ESTATUS_SEGUIMIENTO",
                dropna=False
            )
            .size()
            .reset_index(
                name="TOTAL"
            )
            .sort_values(
                "TOTAL",
                ascending=False
            )
        )

        resumen_seguimiento["SEMAFORO"] = resumen_seguimiento[
            "ESTATUS_SEGUIMIENTO"
        ].apply(
            semaforo_seguimiento
        )

        chart_seguimiento = (
            alt.Chart(
                resumen_seguimiento
            )
            .mark_bar()
            .encode(
                x=alt.X(
                    "TOTAL:Q",
                    title="Total"
                ),
                y=alt.Y(
                    "SEMAFORO:N",
                    sort="-x",
                    title="Estatus seguimiento"
                ),
                tooltip=[
                    "SEMAFORO",
                    "TOTAL"
                ]
            )
        )

        st.altair_chart(
            chart_seguimiento,
            use_container_width=True
        )

        dataframe_limpio(
            resumen_seguimiento[
                [
                    "SEMAFORO",
                    "ESTATUS_SEGUIMIENTO",
                    "TOTAL"
                ]
            ]
        )

        st.divider()

        st.subheader(
            "🔎 Filtros ejecutivos"
        )

        f1, f2, f3, f4 = st.columns(
            4
        )

        entidad_filtro = f1.selectbox(
            "Entidad",
            ["Todas"] + sorted(
                df_resumen["ENTIDAD"]
                .dropna()
                .astype(str)
                .unique()
                .tolist()
            )
            if "ENTIDAD" in df_resumen.columns
            else ["Todas"]
        )

        estatus_filtro = f2.selectbox(
            "Estatus seguimiento",
            ["Todos"] + sorted(
                df_resumen["ESTATUS_SEGUIMIENTO"]
                .dropna()
                .astype(str)
                .unique()
                .tolist()
            )
        )

        atribuible_filtro = f3.selectbox(
            "Atribuible a",
            ["Todos"] + sorted(
                df_resumen["ATRIBUIBLE A"]
                .dropna()
                .astype(str)
                .unique()
                .tolist()
            )
            if "ATRIBUIBLE A" in df_resumen.columns
            else ["Todos"]
        )

        responsable_filtro = f4.selectbox(
            "Responsable",
            ["Todos"] + sorted(
                df_resumen["RESPONSABLE"]
                .dropna()
                .astype(str)
                .unique()
                .tolist()
            )
            if "RESPONSABLE" in df_resumen.columns
            else ["Todos"]
        )

        df_filtrado = df_resumen.copy()

        if entidad_filtro != "Todas" and "ENTIDAD" in df_filtrado.columns:

            df_filtrado = df_filtrado[
                df_filtrado["ENTIDAD"].astype(str) == entidad_filtro
            ]

        if estatus_filtro != "Todos":

            df_filtrado = df_filtrado[
                df_filtrado["ESTATUS_SEGUIMIENTO"].astype(str) == estatus_filtro
            ]

        if atribuible_filtro != "Todos" and "ATRIBUIBLE A" in df_filtrado.columns:

            df_filtrado = df_filtrado[
                df_filtrado["ATRIBUIBLE A"].astype(str) == atribuible_filtro
            ]

        if responsable_filtro != "Todos" and "RESPONSABLE" in df_filtrado.columns:

            df_filtrado = df_filtrado[
                df_filtrado["RESPONSABLE"].astype(str) == responsable_filtro
            ]

        st.caption(
            f"Registros filtrados: {len(df_filtrado):,}"
        )

        st.divider()

        col_izq, col_der = st.columns(
            2
        )

        with col_izq:

            st.subheader(
                "🏥 Estados con más incidencias"
            )

            if "ENTIDAD" in df_filtrado.columns:

                top_entidades = (
                    df_filtrado
                    .groupby(
                        "ENTIDAD",
                        dropna=False
                    )
                    .size()
                    .reset_index(
                        name="TOTAL"
                    )
                    .sort_values(
                        "TOTAL",
                        ascending=False
                    )
                    .head(
                        10
                    )
                )

                if not top_entidades.empty:

                    chart_entidades = (
                        alt.Chart(
                            top_entidades
                        )
                        .mark_bar()
                        .encode(
                            x=alt.X(
                                "TOTAL:Q",
                                title="Total"
                            ),
                            y=alt.Y(
                                "ENTIDAD:N",
                                sort="-x",
                                title="Entidad"
                            ),
                            tooltip=[
                                "ENTIDAD",
                                "TOTAL"
                            ]
                        )
                    )

                    st.altair_chart(
                        chart_entidades,
                        use_container_width=True
                    )

                    dataframe_limpio(
                        top_entidades
                    )

        with col_der:

            st.subheader(
                "🏢 Proveedores con más incidencias"
            )

            if "PROVEEDOR" in df_filtrado.columns:

                top_proveedores = (
                    df_filtrado
                    .groupby(
                        "PROVEEDOR",
                        dropna=False
                    )
                    .size()
                    .reset_index(
                        name="TOTAL"
                    )
                    .sort_values(
                        "TOTAL",
                        ascending=False
                    )
                    .head(
                        10
                    )
                )

                if not top_proveedores.empty:

                    chart_proveedores = (
                        alt.Chart(
                            top_proveedores
                        )
                        .mark_bar()
                        .encode(
                            x=alt.X(
                                "TOTAL:Q",
                                title="Total"
                            ),
                            y=alt.Y(
                                "PROVEEDOR:N",
                                sort="-x",
                                title="Proveedor"
                            ),
                            tooltip=[
                                "PROVEEDOR",
                                "TOTAL"
                            ]
                        )
                    )

                    st.altair_chart(
                        chart_proveedores,
                        use_container_width=True
                    )

                    dataframe_limpio(
                        top_proveedores
                    )

        st.divider()

        col_a, col_b = st.columns(
            2
        )

        with col_a:

            st.subheader(
                "📌 Tipos de incidencia"
            )

            if "TIPO_INCIDENCIA" in df_filtrado.columns:

                top_tipos = (
                    df_filtrado
                    .groupby(
                        "TIPO_INCIDENCIA",
                        dropna=False
                    )
                    .size()
                    .reset_index(
                        name="TOTAL"
                    )
                    .sort_values(
                        "TOTAL",
                        ascending=False
                    )
                    .head(
                        10
                    )
                )

                if not top_tipos.empty:

                    st.altair_chart(
                        alt.Chart(
                            top_tipos
                        )
                        .mark_bar()
                        .encode(
                            x=alt.X(
                                "TOTAL:Q",
                                title="Total"
                            ),
                            y=alt.Y(
                                "TIPO_INCIDENCIA:N",
                                sort="-x",
                                title="Tipo incidencia"
                            ),
                            tooltip=[
                                "TIPO_INCIDENCIA",
                                "TOTAL"
                            ]
                        ),
                        use_container_width=True
                    )

                    dataframe_limpio(
                        top_tipos
                    )

        with col_b:

            st.subheader(
                "👤 Atribuible a"
            )

            if "ATRIBUIBLE A" in df_filtrado.columns:

                top_atribuible = (
                    df_filtrado
                    .groupby(
                        "ATRIBUIBLE A",
                        dropna=False
                    )
                    .size()
                    .reset_index(
                        name="TOTAL"
                    )
                    .sort_values(
                        "TOTAL",
                        ascending=False
                    )
                    .head(
                        10
                    )
                )

                if not top_atribuible.empty:

                    st.altair_chart(
                        alt.Chart(
                            top_atribuible
                        )
                        .mark_bar()
                        .encode(
                            x=alt.X(
                                "TOTAL:Q",
                                title="Total"
                            ),
                            y=alt.Y(
                                "ATRIBUIBLE A:N",
                                sort="-x",
                                title="Atribuible a"
                            ),
                            tooltip=[
                                "ATRIBUIBLE A",
                                "TOTAL"
                            ]
                        ),
                        use_container_width=True
                    )

                    dataframe_limpio(
                        top_atribuible
                    )

        st.divider()

        st.subheader(
            "⚠️ Riesgo operativo"
        )

        riesgo = df_filtrado[
            df_filtrado["ESTATUS_SEGUIMIENTO"]
            .astype(str)
            .str.upper()
            .isin(
                [
                    "INCOMPLETA-SIN ENTREGAR",
                    "INCOMPLETA-CANCELADA"
                ]
            )
        ].copy()

        if riesgo.empty:

            st.success(
                "No hay registros de riesgo operativo con los filtros seleccionados."
            )

        else:

            riesgo["SEMAFORO"] = riesgo[
                "ESTATUS_SEGUIMIENTO"
            ].apply(
                semaforo_seguimiento
            )

            columnas_riesgo = [
                "SEMAFORO",
                "ENTIDAD",
                "ORDEN",
                "CLAVE_CNIS",
                "PROVEEDOR",
                "TIPO_INCIDENCIA",
                "ATRIBUIBLE A",
                "ESTATUS_SEGUIMIENTO",
                "ESTATUS_INCIDENCIA",
                "RESPONSABLE"
            ]

            columnas_riesgo = [
                c for c in columnas_riesgo
                if c in riesgo.columns
            ]

            dataframe_limpio(
                riesgo[columnas_riesgo].head(
                    300
                )
            )

        st.divider()

        with st.expander(
            "⬇️ Exportar resumen filtrado"
        ):

            excel = convertir_excel(
                df_filtrado
            )

            st.download_button(
                label="⬇️ Descargar resumen ejecutivo en Excel",
                data=excel,
                file_name="resumen_ejecutivo_incidencias.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )


# =========================
# REGISTRAR INCIDENCIA
# =========================

if menu == "Registrar incidencia":

    st.subheader(
        "📝 Registrar incidencia"
    )

    st.caption(
        "Modo rápido: compendio e incidencias se consultan directo en Supabase."
    )

    modo_registro = st.radio(
        "Tipo de captura",
        [
            "Individual",
            "Masiva"
        ],
        horizontal=True
    )

    if modo_registro == "Individual":

        valor_busqueda = st.text_input(
            "Orden de suministro",
            placeholder="Ejemplo: CIMB-28-01-2025-28030776-U013"
        )

        if valor_busqueda:

            with st.spinner(
                "Buscando orden en Supabase..."
            ):

                datos_orden, resultado = obtener_datos_orden_para_registro(
                    valor_busqueda
                )

            if datos_orden is None:

                st.warning(
                    "No encontré esa orden exacta."
                )

                sugerencias = sugerir_ordenes(
                    valor_busqueda
                )

                if len(sugerencias) > 0:

                    st.info(
                        "Posibles coincidencias:"
                    )

                    dataframe_limpio(
                        sugerencias
                    )

                else:

                    st.error(
                        "No encontré coincidencias en Supabase."
                    )

            else:

                estatus_base = datos_orden["estatus_base"]
                origen_compendio = datos_orden["origen_compendio"]
                orden = datos_orden["orden"]
                tipo_entrega = datos_orden["tipo_entrega"]
                entidad = datos_orden["entidad"]
                clues_destino = datos_orden["clues_destino"]
                unidad_destino = datos_orden["unidad_destino"]
                almacen = datos_orden["almacen"]
                proveedor = datos_orden["proveedor"]
                clave = datos_orden["clave"]
                descripcion = datos_orden["descripcion"]
                piezas_emitidas = datos_orden["piezas_emitidas"]
                piezas_recibidas_ol = datos_orden["piezas_recibidas_ol"]
                piezas_entregadas = datos_orden["piezas_entregadas"]
                operador = datos_orden["operador"]
                tipo_red = datos_orden["tipo_red"]
                grupo_terapeutico = datos_orden["grupo_terapeutico"]
                estatus_orden = datos_orden["estatus_orden"]
                incidencias_previas = obtener_incidencias_previas_supabase(
                    orden
                )

                if incidencias_previas.empty:

                    incidencias_previas = obtener_incidencias_previas(
                        incidencias,
                        orden
                    )

                st.divider()

                c_estado, c_previas = st.columns(
                    2
                )

                with c_estado:

                    if str(
                        estatus_base
                    ).upper().strip() == "INACTIVA":

                        st.error(
                            "🚫 Orden CANCELADA / INACTIVA"
                        )

                    else:

                        st.success(
                            "✅ Orden ACTIVA"
                        )

                with c_previas:

                    if len(
                        incidencias_previas
                    ) > 0:

                        st.warning(
                            f"⚠️ Ya tiene {len(incidencias_previas)} incidencia(s)"
                        )

                    else:

                        st.success(
                            "🟢 Sin incidencias previas"
                        )

                if len(
                    incidencias_previas
                ) > 0:

                    with st.expander(
                        "Ver incidencias previas"
                    ):

                        dataframe_limpio(
                            incidencias_previas
                        )

                st.subheader(
                    "📋 Datos encontrados en Supabase"
                )

                dataframe_limpio(
                    resultado.head(
                        50
                    )
                )

                st.divider()

                st.subheader(
                    "🔒 Información de la orden"
                )

                c1, c2, c3 = st.columns(
                    3
                )

                c1.text_input(
                    "Entidad",
                    entidad,
                    disabled=True
                )

                c2.text_input(
                    "Almacén / CLUES destino",
                    almacen,
                    disabled=True
                )

                c3.text_input(
                    "Proveedor",
                    proveedor,
                    disabled=True
                )

                c4, c5, c6 = st.columns(
                    3
                )

                c4.text_input(
                    "Orden",
                    orden,
                    disabled=True
                )

                c5.text_input(
                    "Clave CNIS",
                    clave,
                    disabled=True
                )

                c6.text_input(
                    "Tipo de entrega",
                    tipo_entrega,
                    disabled=True
                )

                c7, c8, c9 = st.columns(
                    3
                )

                c7.text_input(
                    "Estatus base",
                    estatus_base,
                    disabled=True
                )

                c8.text_input(
                    "Origen compendio",
                    origen_compendio,
                    disabled=True
                )

                c9.text_input(
                    "Estatus orden",
                    estatus_orden,
                    disabled=True
                )

                c10, c11 = st.columns(
                    2
                )

                c10.text_input(
                    "Tipo de red",
                    tipo_red,
                    disabled=True
                )

                c11.text_input(
                    "Grupo terapéutico",
                    grupo_terapeutico,
                    disabled=True
                )

                st.text_area(
                    "Descripción",
                    descripcion,
                    disabled=True
                )

                st.divider()

                st.subheader(
                    "🚚 Validación logística"
                )

                c12, c13, c14 = st.columns(
                    3
                )

                c12.metric(
                    "Piezas emitidas",
                    piezas_emitidas
                )

                c13.metric(
                    "Piezas recibidas OL",
                    piezas_recibidas_ol
                )

                c14.metric(
                    "Piezas entregadas CLUES",
                    piezas_entregadas
                )

                c15, c16 = st.columns(
                    2
                )

                c15.text_input(
                    "Estatus recepción OL",
                    estatus_recepcion_ol,
                    disabled=True
                )

                c16.text_input(
                    "Estatus entrega Estado / CLUES",
                    estatus_entrega_estado,
                    disabled=True
                )

                st.text_input(
                    "Incidencia automática",
                    estatus_completa,
                    disabled=True
                )

                st.divider()

                st.subheader(
                    "✍️ Captura de incidencia"
                )

                c17, c18, c19 = st.columns(
                    3
                )

                with c17:

                    atribuible = st.selectbox(
                        "Atribuible a",
                        ATRIBUIBLES
                    )

                with c18:

                    tipo = st.selectbox(
                        "Tipo de incidencia",
                        TIPOS_INCIDENCIA_GENERAL
                    )

                with c19:

                    estatus = st.selectbox(
                        "Estatus incidencia",
                        [
                            "En proceso",
                            "Rechazado",
                            "Resuelto"
                        ]
                    )

                responsable = st.selectbox(
                    "Responsable",
                    MONITORES
                )

                observaciones = st.text_area(
                    "Observaciones"
                )

                st.subheader(
                    "📎 Evidencias"
                )

                c20, c21 = st.columns(
                    2
                )

                with c20:

                    cedula_rechazo = st.file_uploader(
                        "Cédula rechazo PDF",
                        type=[
                            "pdf"
                        ]
                    )

                with c21:

                    correo_seguimiento = st.file_uploader(
                        "Correo seguimiento PDF",
                        type=[
                            "pdf"
                        ]
                    )

                guardar = st.button(
                    "💾 Guardar incidencia",
                    use_container_width=True
                )

                if guardar:

                    ruta_cedula = subir_pdf_evidencia_drive(
                        cedula_rechazo,
                        orden,
                        "cedula",
                        entidad,
                        clues_destino
                    )

                    ruta_correo = subir_pdf_evidencia_drive(
                        correo_seguimiento,
                        orden,
                        "correo",
                        entidad,
                        clues_destino
                    )

                    nueva = construir_registro_incidencia(
                        valor_busqueda,
                        datos_orden,
                        atribuible,
                        tipo,
                        estatus,
                        responsable,
                        observaciones,
                        ruta_cedula,
                        ruta_correo
                    )

                    duplicada = existe_incidencia_duplicada(
                        orden,
                        tipo,
                        atribuible
                    )

                    if duplicada:

                        st.warning(
                            "Esta orden ya tiene una incidencia igual registrada."
                        )

                    else:

                        guardar_incidencia(
                            nueva
                        )

                        st.cache_data.clear()

                        link_respaldo = generar_respaldo_drive()

                        if link_respaldo:

                            st.success(
                                "Incidencia guardada correctamente. Evidencias en Drive y Google Sheets actualizado."
                            )

                        else:

                            st.success(
                                "Incidencia guardada correctamente. Las evidencias se enviaron a Drive."
                            )

                        st.cache_data.clear()

                        st.rerun()

    else:

        st.info(
            "Pega una orden por línea. Se aplicará la misma incidencia a todas."
        )

        texto_ordenes = st.text_area(
            "Órdenes de suministro",
            height=220,
            placeholder="IMBB-16-02-2025-16244042-U013\\nIMBB-16-02-2025-16244043-U013"
        )

        ordenes = extraer_ordenes_masivas(
            texto_ordenes
        )

        st.caption(
            f"Órdenes detectadas: {len(ordenes)}"
        )

        c1, c2, c3 = st.columns(
            3
        )

        with c1:

            atribuible_m = st.selectbox(
                "Atribuible a",
                ATRIBUIBLES,
                key="masivo_atribuible"
            )

        with c2:

            tipo_m = st.selectbox(
                "Tipo de incidencia",
                TIPOS_INCIDENCIA_GENERAL,
                key="masivo_tipo"
            )

        with c3:

            estatus_m = st.selectbox(
                "Estatus incidencia",
                [
                    "En proceso",
                    "Rechazado",
                    "Resuelto"
                ],
                key="masivo_estatus"
            )

        responsable_m = st.selectbox(
            "Responsable",
            MONITORES,
            key="masivo_responsable"
        )

        observaciones_m = st.text_area(
            "Observaciones",
            key="masivo_observaciones"
        )

        validar_masivo = st.button(
            "🔎 Validar órdenes",
            use_container_width=True
        )

        if validar_masivo and ordenes:

            registros_preview = []

            for orden_m in ordenes:

                datos_m, resultado_m = obtener_datos_orden_para_registro(
                    orden_m
                )

                if datos_m is None:

                    registros_preview.append(
                        {
                            "ORDEN": orden_m,
                            "ESTATUS": "NO ENCONTRADA",
                            "ENTIDAD": "",
                            "CLUES": "",
                            "PROVEEDOR": "",
                            "INCIDENCIAS_PREVIAS": 0
                        }
                    )

                    continue

                previas_m = obtener_incidencias_previas_supabase(
                    datos_m["orden"]
                )

                registros_preview.append(
                    {
                        "ORDEN": datos_m["orden"],
                        "ESTATUS": datos_m["estatus_base"],
                        "ENTIDAD": datos_m["entidad"],
                        "CLUES": datos_m["clues_destino"],
                        "PROVEEDOR": datos_m["proveedor"],
                        "INCIDENCIAS_PREVIAS": len(previas_m)
                    }
                )

            st.session_state["preview_masivo"] = registros_preview

        if "preview_masivo" in st.session_state:

            preview_df = pd.DataFrame(
                st.session_state["preview_masivo"]
            )

            dataframe_limpio(
                preview_df
            )

            encontradas = preview_df[
                preview_df["ESTATUS"] != "NO ENCONTRADA"
            ].shape[0]

            st.success(
                f"Listas para guardar: {encontradas} de {len(preview_df)}"
            )

        guardar_masivo = st.button(
            "💾 Guardar incidencia para todas las órdenes encontradas",
            use_container_width=True
        )

        if guardar_masivo:

            if not ordenes:

                st.warning(
                    "Primero pega al menos una orden."
                )

            else:

                guardadas = 0

                no_encontradas = []

                errores = []

                with st.spinner(
                    "Guardando incidencias masivas..."
                ):

                    for orden_m in ordenes:

                        try:

                            datos_m, resultado_m = obtener_datos_orden_para_registro(
                                orden_m
                            )

                            if datos_m is None:

                                no_encontradas.append(
                                    orden_m
                                )

                                continue

                            duplicada_m = existe_incidencia_duplicada(
                                datos_m["orden"],
                                tipo_m,
                                atribuible_m
                            )

                            if duplicada_m:

                                continue

                            nueva_m = construir_registro_incidencia(
                                orden_m,
                                datos_m,
                                atribuible_m,
                                tipo_m,
                                estatus_m,
                                responsable_m,
                                observaciones_m,
                                "",
                                ""
                            )

                            guardar_incidencia(
                                nueva_m
                            )

                            guardadas += 1

                        except Exception as e:

                            errores.append(
                                f"{orden_m}: {e}"
                            )

                link_respaldo = ""

                if guardadas > 0:

                    st.cache_data.clear()

                    link_respaldo = generar_respaldo_drive()

                if link_respaldo:

                    st.success(
                        f"Guardadas correctamente: {guardadas}. Google Sheets actualizado."
                    )

                else:

                    st.success(
                        f"Guardadas correctamente: {guardadas}"
                    )

                if no_encontradas:

                    st.warning(
                        f"No encontradas: {len(no_encontradas)}"
                    )

                    st.write(
                        no_encontradas
                    )

                if errores:

                    st.error(
                        f"Errores: {len(errores)}"
                    )

                    st.write(
                        errores
                    )


# =========================
# URGENCIAS
# =========================

elif menu == "Urgencias":

    st.subheader(
        "🚨 Registro de urgencias"
    )

    st.caption(
        "Captura la clave, la entidad y la fecha de solicitud. "
        "El sistema localizará las órdenes existentes y conservará únicamente "
        "las que no tengan entrega registrada en la entidad."
    )

    c1, c2, c3 = st.columns(
        [
            1,
            2,
            1
        ]
    )

    with c1:

        clave_urgencia = st.text_input(
            "Clave CNIS",
            placeholder="Ejemplo: 010.000.0012.00",
            key="urgencias_clave"
        ).strip()

    with c2:

        entidades_urgencia = obtener_entidades_compendio_urgencia()

        opciones_entidades_urgencia = [
            "Selecciona una entidad"
        ] + entidades_urgencia

        entidad_urgencia = st.selectbox(
            "Estado / entidad",
            opciones_entidades_urgencia,
            key="urgencias_entidad"
        )

        if entidad_urgencia == "Selecciona una entidad":

            entidad_urgencia = ""

    with c3:

        fecha_urgencia = st.date_input(
            "Fecha de solicitud",
            value=datetime.now().date(),
            key="urgencias_fecha"
        )

    responsable_urgencia = st.selectbox(
        "Monitora / responsable",
        MONITORES,
        key="urgencias_responsable"
    )

    observaciones_urgencia_captura = st.text_area(
        "Observaciones de la urgencia",
        placeholder="Describe brevemente la necesidad o el seguimiento requerido.",
        key="urgencias_observaciones"
    )

    pdf_urgencia = st.file_uploader(
        "📎 Adjuntar PDF de correo o evidencia",
        type=[
            "pdf"
        ],
        key="urgencias_pdf",
        help="Selecciona el correo o la evidencia que respalda la urgencia."
    )

    buscar_urgencia = st.button(
        "🔎 Consultar órdenes de la clave",
        use_container_width=True,
        key="urgencias_buscar"
    )

    if buscar_urgencia:

        if not clave_urgencia or not entidad_urgencia:

            st.warning(
                "Captura la clave CNIS y el estado / entidad antes de consultar."
            )

        else:

            with st.spinner(
                "Consultando órdenes y validando entregas en la entidad..."
            ):

                resultado_urgencia = buscar_ordenes_urgencia(
                    clave_urgencia,
                    entidad_urgencia
                )

            st.session_state["resultado_urgencias"] = {
                "clave": normalizar_clave_urgencia(
                    clave_urgencia
                ),
                "entidad": normalizar_entidad_urgencia(
                    entidad_urgencia
                ),
                "fecha": fecha_urgencia_iso(
                    fecha_urgencia
                ),
                "resultado": resultado_urgencia
            }

    resultado_guardado = st.session_state.get(
        "resultado_urgencias"
    )

    consulta_actual = (
        resultado_guardado
        and resultado_guardado.get("clave") == normalizar_clave_urgencia(
            clave_urgencia
        )
        and resultado_guardado.get("entidad") == normalizar_entidad_urgencia(
            entidad_urgencia
        )
        and resultado_guardado.get("fecha") == fecha_urgencia_iso(
            fecha_urgencia
        )
    )

    if consulta_actual:

        resultado_urgencia = resultado_guardado["resultado"]
        ordenes_urgencia = resultado_urgencia.get(
            "ordenes",
            pd.DataFrame()
        )
        ordenes_entregadas_urgencia = resultado_urgencia.get(
            "ordenes_entregadas",
            pd.DataFrame()
        )

        st.divider()
        st.subheader(
            "Resultado de la consulta"
        )

        m1, m2, m3, m4 = st.columns(
            4
        )

        m1.metric(
            "Registros de la clave",
            resultado_urgencia.get("total_clave", 0)
        )

        m2.metric(
            "Órdenes en la entidad",
            resultado_urgencia.get("total_con_orden", 0)
        )

        m3.metric(
            "No entregadas elegibles",
            len(ordenes_urgencia)
        )

        m4.metric(
            "Excluidas por entrega",
            len(ordenes_entregadas_urgencia)
        )

        if not resultado_urgencia.get(
            "clave_valida",
            False
        ):

            st.error(
                "La clave capturada no existe en el compendio. "
                "No se puede registrar la urgencia."
            )

        elif not resultado_urgencia.get(
            "entidad_homologada",
            ""
        ):

            st.error(
                "La clave existe en el compendio, pero el estado capturado "
                "no coincide con una entidad homologada."
            )

        else:

            st.success(
                "Estado homologado desde el compendio: "
                f"{resultado_urgencia['entidad_homologada']}"
            )

        if not ordenes_urgencia.empty:

            st.success(
                "Estas son las órdenes que sí se pueden registrar como urgencia "
                "porque no tienen entrega registrada en la entidad."
            )

            columnas_preview = [
                "CLAVE_CNIS",
                "ENTIDAD",
                "ORDEN_SUMINISTRO",
                "ESTATUS_BASE",
                "TIPO_ENTREGA",
                "CLUES_DESTINO",
                "UNIDAD_DESTINO",
                "PROVEEDOR",
                "PIEZAS_EMITIDAS",
                "PIEZAS_ENTREGADAS_CLUES",
                "ESTATUS_ENTREGA_ENTIDAD"
            ]

            dataframe_limpio(
                ordenes_urgencia[
                    [
                        columna for columna in columnas_preview
                        if columna in ordenes_urgencia.columns
                    ]
                ]
            )

        elif resultado_urgencia.get("sin_orden", False):

            st.warning(
                "No se encontró una orden de suministro para la clave y entidad. "
                "La solicitud se guardará en la hoja Sin_Orden_Urgencias."
            )

        elif not ordenes_entregadas_urgencia.empty:

            st.warning(
                "La clave sí tiene órdenes en la entidad, pero todas tienen "
                "entrega registrada. No se guardará ninguna orden para la urgencia."
            )

        if not ordenes_entregadas_urgencia.empty:

            with st.expander(
                "Ver órdenes excluidas por tener entrega en la entidad"
            ):

                columnas_excluidas = [
                    "ORDEN_SUMINISTRO",
                    "PIEZAS_ENTREGADAS_CLUES",
                    "ESTATUS_ENTREGA_ENTIDAD",
                    "MOTIVO_EXCLUSION"
                ]

                dataframe_limpio(
                    ordenes_entregadas_urgencia[
                        [
                            columna for columna in columnas_excluidas
                            if columna in ordenes_entregadas_urgencia.columns
                        ]
                    ]
                )

        puede_guardar_urgencia = (
            resultado_urgencia.get("clave_valida", False)
            and bool(resultado_urgencia.get("entidad_homologada", ""))
            and (
                not ordenes_urgencia.empty
                or resultado_urgencia.get("sin_orden", False)
            )
        )

        if puede_guardar_urgencia:

            st.divider()

            guardar_urgencia = st.button(
                "💾 Registrar urgencia y actualizar base de Drive",
                use_container_width=True,
                key="urgencias_guardar"
            )

            if guardar_urgencia:

                with st.spinner(
                    "Registrando incidencia y actualizando la base de Urgencias..."
                ):

                    resultado_guardado_urgencia = registrar_urgencia_completa(
                        clave_urgencia,
                        resultado_urgencia.get(
                            "entidad_homologada",
                            entidad_urgencia
                        ),
                        fecha_urgencia,
                        responsable_urgencia,
                        observaciones_urgencia_captura,
                        resultado_urgencia,
                        incidencias,
                        pdf_urgencia
                    )

                if resultado_guardado_urgencia.get("guardadas", 0) > 0:

                    st.success(
                        "Urgencia registrada como incidencia."
                    )

                if resultado_guardado_urgencia.get("duplicadas", 0) > 0:

                    st.info(
                        "Se omitieron registros que ya estaban capturados "
                        "para la misma clave, entidad, orden y fecha."
                    )

                drive_urgencia = resultado_guardado_urgencia.get(
                    "drive"
                )

                if drive_urgencia:

                    st.success(
                        "Base de Urgencias actualizada en Drive: "
                        f"{drive_urgencia.get('ordenes', 0)} órdenes y "
                        f"{drive_urgencia.get('sin_orden', 0)} solicitudes sin orden."
                    )

                    st.link_button(
                        "🔗 Abrir base de Urgencias en Drive",
                        drive_urgencia.get("link", ""),
                        use_container_width=True
                    )

                if resultado_guardado_urgencia.get("errores"):

                    st.error(
                        "Se presentaron errores durante el registro."
                    )

                    st.write(
                        resultado_guardado_urgencia.get("errores")
                    )

                st.cache_data.clear()
                st.session_state.pop(
                    "resultado_urgencias",
                    None
                )

    st.divider()

    st.subheader(
        "📋 Carga masiva de claves"
    )

    st.caption(
        "Pega una clave por línea o varias separadas por coma, punto y coma "
        "o tabulador. Todas se validarán contra el compendio."
    )

    texto_claves_urgencias = st.text_area(
        "Claves CNIS para carga masiva",
        height=180,
        placeholder="010.000.0012.00\n010.000.0045.00\n010.000.0099.00",
        key="urgencias_claves_masivas"
    )

    claves_urgencias_masivas = extraer_claves_urgencia_masivas(
        texto_claves_urgencias
    )

    st.caption(
        f"Claves detectadas: {len(claves_urgencias_masivas)}"
    )

    consultar_masivo_urgencias = st.button(
        "🔎 Validar claves masivas",
        use_container_width=True,
        key="urgencias_validar_masivo"
    )

    if consultar_masivo_urgencias:

        if not claves_urgencias_masivas:

            st.warning(
                "Pega al menos una clave para realizar la validación."
            )

        elif not entidad_urgencia:

            st.warning(
                "Selecciona una entidad antes de validar las claves."
            )

        else:

            resultados_masivos = []

            with st.spinner(
                "Validando claves contra el compendio y revisando entregas..."
            ):

                for clave_masiva in claves_urgencias_masivas:

                    resultado_masivo = buscar_ordenes_urgencia(
                        clave_masiva,
                        entidad_urgencia
                    )

                    resultados_masivos.append(
                        {
                            "clave": clave_masiva,
                            "resultado": resultado_masivo
                        }
                    )

            st.session_state[
                "resultados_urgencias_masivas"
            ] = {
                "entidad": normalizar_entidad_urgencia(
                    entidad_urgencia
                ),
                "fecha": fecha_urgencia_iso(
                    fecha_urgencia
                ),
                "resultados": resultados_masivos
            }

    resultados_masivos_guardados = st.session_state.get(
        "resultados_urgencias_masivas"
    )

    consulta_masiva_actual = (
        resultados_masivos_guardados
        and resultados_masivos_guardados.get("entidad") == normalizar_entidad_urgencia(
            entidad_urgencia
        )
        and resultados_masivos_guardados.get("fecha") == fecha_urgencia_iso(
            fecha_urgencia
        )
        and resultados_masivos_guardados.get("resultados")
    )

    if consulta_masiva_actual:

        resumen_masivo = []

        for item_masivo in resultados_masivos_guardados["resultados"]:

            resultado_masivo = item_masivo["resultado"]
            ordenes_masivas = resultado_masivo.get(
                "ordenes",
                pd.DataFrame()
            )
            ordenes_entregadas_masivas = resultado_masivo.get(
                "ordenes_entregadas",
                pd.DataFrame()
            )

            if not resultado_masivo.get("clave_valida", False):

                estatus_masivo = "CLAVE NO EXISTE EN COMPENDIO"

            elif not resultado_masivo.get("entidad_homologada", ""):

                estatus_masivo = "ENTIDAD NO HOMOLOGADA"

            elif not ordenes_masivas.empty:

                estatus_masivo = "LISTA PARA REGISTRAR"

            elif resultado_masivo.get("sin_orden", False):

                estatus_masivo = "SIN ORDEN"

            elif not ordenes_entregadas_masivas.empty:

                estatus_masivo = "TODAS ENTREGADAS"

            else:

                estatus_masivo = "SIN RESULTADO"

            resumen_masivo.append(
                {
                    "CLAVE_CNIS": item_masivo["clave"],
                    "ESTATUS": estatus_masivo,
                    "ORDENES_NO_ENTREGADAS": len(ordenes_masivas),
                    "ORDENES_EXCLUIDAS_ENTREGADAS": len(
                        ordenes_entregadas_masivas
                    ),
                    "ENTIDAD_HOMOLOGADA": resultado_masivo.get(
                        "entidad_homologada",
                        ""
                    )
                }
            )

        st.subheader(
            "Resultado de validación masiva"
        )

        dataframe_limpio(
            pd.DataFrame(
                resumen_masivo
            )
        )

        guardar_masivo_urgencias = st.button(
            "💾 Registrar todas las urgencias válidas",
            use_container_width=True,
            key="urgencias_guardar_masivo"
        )

        if guardar_masivo_urgencias:

            total_guardadas_masivo = 0
            total_duplicadas_masivo = 0
            errores_masivos = []

            with st.spinner(
                "Registrando las urgencias válidas..."
            ):

                for item_masivo in resultados_masivos_guardados["resultados"]:

                    clave_masiva = item_masivo["clave"]
                    resultado_masivo = item_masivo["resultado"]

                    if not resultado_masivo.get("clave_valida", False):

                        errores_masivos.append(
                            f"{clave_masiva}: no existe en el compendio"
                        )
                        continue

                    entidad_masiva = resultado_masivo.get(
                        "entidad_homologada",
                        ""
                    )

                    if not entidad_masiva:

                        errores_masivos.append(
                            f"{clave_masiva}: entidad no homologada"
                        )
                        continue

                    resultado_registro_masivo = registrar_urgencia_completa(
                        clave_masiva,
                        entidad_masiva,
                        fecha_urgencia,
                        responsable_urgencia,
                        observaciones_urgencia_captura,
                        resultado_masivo,
                        incidencias,
                        pdf_urgencia
                    )

                    total_guardadas_masivo += resultado_registro_masivo.get(
                        "guardadas",
                        0
                    )
                    total_duplicadas_masivo += resultado_registro_masivo.get(
                        "duplicadas",
                        0
                    )
                    errores_masivos.extend(
                        resultado_registro_masivo.get(
                            "errores",
                            []
                        )
                    )

            if total_guardadas_masivo:

                st.success(
                    f"Se registraron {total_guardadas_masivo} urgencias correctamente."
                )

            if total_duplicadas_masivo:

                st.info(
                    f"Se omitieron {total_duplicadas_masivo} registros duplicados."
                )

            if errores_masivos:

                st.warning(
                    "Algunas claves no se registraron:"
                )
                st.write(
                    errores_masivos
                )

            st.cache_data.clear()
            st.session_state.pop(
                "resultados_urgencias_masivas",
                None
            )


# =========================
# SEGUIMIENTO
# =========================

elif menu == "Seguimiento":

    st.subheader(
        "📋 Seguimiento de incidencias"
    )

    st.caption(
        "El cruce con compendio se actualiza automáticamente solo una vez al día. "
        "Las incidencias nuevas se cruzan al momento de guardarse."
    )

    df_seg = incidencias.copy()

    if df_seg.empty:

        st.info(
            "Aún no hay incidencias registradas."
        )

    else:

        if "ESTATUS_SEGUIMIENTO" not in df_seg.columns:

            df_seg["ESTATUS_SEGUIMIENTO"] = ""

        if "ATRIBUIBLE A" in df_seg.columns:

            df_seg["ATRIBUIBLE A"] = df_seg["ATRIBUIBLE A"].apply(
                homologar_atribuible
            )

        df_seg["SEMAFORO"] = df_seg["ESTATUS_SEGUIMIENTO"].apply(
            semaforo_seguimiento
        )

        c1, c2, c3, c4 = st.columns(
            4
        )

        entidad_filtro = c1.selectbox(
            "Entidad",
            ["Todas"] + sorted(
                df_seg["ENTIDAD"]
                .dropna()
                .astype(str)
                .unique()
                .tolist()
            )
        )

        estatus_filtro = c2.selectbox(
            "Estatus seguimiento",
            ["Todos"] + sorted(
                df_seg["ESTATUS_SEGUIMIENTO"]
                .dropna()
                .astype(str)
                .unique()
                .tolist()
            )
        )

        atribuible_filtro = c3.selectbox(
            "Atribuible a",
            ["Todos"] + sorted(
                df_seg["ATRIBUIBLE A"]
                .dropna()
                .astype(str)
                .unique()
                .tolist()
            )
        )

        buscar_orden = c4.text_input(
            "Buscar orden"
        )

        if entidad_filtro != "Todas":

            df_seg = df_seg[
                df_seg["ENTIDAD"].astype(str) == entidad_filtro
            ]

        if estatus_filtro != "Todos":

            df_seg = df_seg[
                df_seg["ESTATUS_SEGUIMIENTO"].astype(str) == estatus_filtro
            ]

        if atribuible_filtro != "Todos":

            df_seg = df_seg[
                df_seg["ATRIBUIBLE A"].astype(str) == atribuible_filtro
            ]

        if buscar_orden.strip():

            orden_norm = normalizar_orden(
                buscar_orden
            )

            df_seg = df_seg[
                df_seg["ORDEN"]
                .astype(str)
                .apply(normalizar_orden)
                .str.contains(
                    orden_norm,
                    na=False,
                    regex=False
                )
            ]

        st.caption(
            f"Registros mostrados: {len(df_seg):,}"
        )

        columnas_mostrar = [
            "SEMAFORO",
            "FECHA_REGISTRO",
            "ENTIDAD",
            "ORDEN",
            "CLAVE_CNIS",
            "PROVEEDOR",
            "TIPO_INCIDENCIA",
            "ATRIBUIBLE A",
            "ESTATUS_SEGUIMIENTO",
            "ESTATUS_INCIDENCIA",
            "RESPONSABLE",
            "OBSERVACIONES",
            "PDF_CEDULA_RECHAZO",
            "PDF_CORREO_SEGUIMIENTO"
        ]

        columnas_mostrar = [
            c for c in columnas_mostrar
            if c in df_seg.columns
        ]

        dataframe_limpio(
            df_seg[columnas_mostrar]
        )

        st.divider()

        with st.expander(
            "📎 Agregar archivo a una incidencia previa"
        ):

            if "ID" not in df_seg.columns:

                st.warning(
                    "No encontré la columna ID para actualizar evidencias."
                )

            else:

                df_archivos = df_seg.copy()

                df_archivos["_OPCION_ARCHIVO"] = df_archivos.apply(
                    lambda x: f"{x.get('ID', '')} | {x.get('ORDEN', '')} | {x.get('CLAVE_CNIS', '')} | {x.get('TIPO_INCIDENCIA', '')}",
                    axis=1
                )

                opcion_archivo = st.selectbox(
                    "Incidencia a actualizar",
                    df_archivos["_OPCION_ARCHIVO"].astype(str).tolist()
                )

                fila_archivo = df_archivos[
                    df_archivos["_OPCION_ARCHIVO"].astype(str) == str(opcion_archivo)
                ].iloc[0]

                tipo_archivo = st.selectbox(
                    "Tipo de archivo",
                    [
                        "Cédula rechazo",
                        "Correo seguimiento"
                    ]
                )

                archivo_nuevo = st.file_uploader(
                    "Archivo PDF",
                    type=[
                        "pdf"
                    ],
                    key="archivo_incidencia_previa"
                )

                if st.button(
                    "📎 Subir archivo a incidencia previa",
                    use_container_width=True
                ):

                    tipo_pdf = (
                        "cedula"
                        if tipo_archivo == "Cédula rechazo"
                        else "correo"
                    )

                    ok = actualizar_evidencia_incidencia(
                        fila_archivo.get("ID", ""),
                        archivo_nuevo,
                        tipo_pdf,
                        fila_archivo.get("ORDEN", fila_archivo.get("ORDEN_BUSCADA", "")),
                        fila_archivo.get("ENTIDAD", "SIN_ESTADO"),
                        fila_archivo.get("CLUES_DESTINO", "SIN_CLUES")
                    )

                    if ok:

                        st.success(
                            "Archivo agregado correctamente a la incidencia previa."
                        )

                        st.rerun()

                    else:

                        st.error(
                            "No se pudo subir el archivo. Revisa que seleccionaste un PDF."
                        )

        excel = convertir_excel(
            df_seg
        )

        st.download_button(
            label="⬇️ Descargar incidencias filtradas en Excel",
            data=excel,
            file_name="reporte_incidencias_filtrado.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )


# =========================
# CÉDULAS DE RECHAZO
# =========================

elif menu == "Cédulas de rechazo":

    st.subheader("📄 Cédulas de rechazo")
    st.caption("Consulta las cédulas registradas, abre cada PDF con un clic o descarga el concentrado en Excel.")

    df_cedulas = incidencias.copy()

    if "PDF_CEDULA_RECHAZO" not in df_cedulas.columns:
        st.info("Aún no existe información de cédulas de rechazo en la base.")
    else:
        enlaces = (
            df_cedulas["PDF_CEDULA_RECHAZO"]
            .fillna("")
            .astype(str)
            .str.strip()
        )
        df_cedulas = df_cedulas[
            ~enlaces.str.lower().isin(["", "nan", "none", "null"])
        ].copy()

        if df_cedulas.empty:
            st.info("No hay cédulas de rechazo registradas para mostrar.")
        else:
            c1, c2, c3 = st.columns(3)

            with c1:
                if "ENTIDAD" in df_cedulas.columns:
                    entidades = sorted(df_cedulas["ENTIDAD"].fillna("").astype(str).unique().tolist())
                    entidad = st.selectbox("Entidad", ["Todas"] + [x for x in entidades if x.strip()], key="cedulas_entidad")
                    if entidad != "Todas":
                        df_cedulas = df_cedulas[df_cedulas["ENTIDAD"].astype(str) == entidad]

            with c2:
                if "TIPO_INCIDENCIA" in df_cedulas.columns:
                    tipos = sorted(df_cedulas["TIPO_INCIDENCIA"].fillna("").astype(str).unique().tolist())
                    tipo = st.selectbox("Tipo de incidencia", ["Todos"] + [x for x in tipos if x.strip()], key="cedulas_tipo")
                    if tipo != "Todos":
                        df_cedulas = df_cedulas[df_cedulas["TIPO_INCIDENCIA"].astype(str) == tipo]

            with c3:
                orden_buscar = st.text_input("Buscar orden", key="cedulas_orden").strip()
                if orden_buscar:
                    col_orden = "ORDEN" if "ORDEN" in df_cedulas.columns else "ORDEN_BUSCADA"
                    if col_orden in df_cedulas.columns:
                        df_cedulas = df_cedulas[df_cedulas[col_orden].fillna("").astype(str).str.contains(orden_buscar, case=False, na=False)]

            st.metric("Cédulas encontradas", len(df_cedulas))

            columnas = [
                "FECHA_REGISTRO", "ENTIDAD", "ORDEN", "ORDEN_BUSCADA",
                "CLAVE_CNIS", "PROVEEDOR", "TIPO_INCIDENCIA",
                "ESTATUS_INCIDENCIA", "RESPONSABLE", "PDF_CEDULA_RECHAZO"
            ]
            columnas = [c for c in columnas if c in df_cedulas.columns]
            vista = df_cedulas[columnas].copy()

            st.dataframe(
                vista,
                use_container_width=True,
                hide_index=True,
                column_config={
                    "PDF_CEDULA_RECHAZO": st.column_config.LinkColumn(
                        "📄 Abrir PDF",
                        display_text="Ver PDF"
                    )
                }
            )

            excel_cedulas = convertir_excel_cedulas_rechazo(vista)
            st.download_button(
                "⬇️ Descargar cédulas de rechazo en Excel",
                data=excel_cedulas,
                file_name="cedulas_de_rechazo.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )


# =========================
# BASE SUPABASE
# =========================

elif menu == "Base Supabase":

    st.subheader(
        "⚡ Base Supabase"
    )

    st.info(
        "Vista rápida de los primeros 100 registros desde Supabase."
    )

    respuesta = (
        supabase
        .table(
            "compendio"
        )
        .select(
            "*"
        )
        .limit(
            100
        )
        .execute()
    )

    muestra = pd.DataFrame(
        respuesta.data
    )

    dataframe_limpio(
        muestra
    )

# version seguimiento 2026	
