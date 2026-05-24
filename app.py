import streamlit as st
import pandas as pd
import numpy as np
import altair as alt
import re
import os
import json
import tempfile

from datetime import datetime

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import (
    MediaIoBaseDownload,
    MediaFileUpload
)

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

FOLDER_ID_BASES = (
    "1J1hHDZDTt8CMVBJ6TW8uPd_EU6laYbhG"
)

FOLDER_ID_EVIDENCIAS = (
    "1Fbxzc1SC-c5yaLh7z1h4qIG8sY5W78B0"
)

FOLDER_ID_INCIDENCIAS = (
    "1bMw-Un3KZHQds0zsRZAKSXuAggSuL4PG"
)

COMPENDIO_FILE_ID = (
    "1JNV2LZrDy-SPAly0bueay7DAz5hnczYf"
)

# =========================
# ARCHIVOS
# =========================

ARCHIVO_COMPENDIO = (
    "CompendioAbasto25-26_22.05.2026.xlsb"
)

ARCHIVO_PARQUET = (
    "base_compendio_ligera.parquet"
)

ARCHIVO_INCIDENCIAS = (
    "incidencias.xlsx"
)

ARCHIVO_AGENDA = (
    "agenda_citas.xlsx"
)

# =========================
# RUTAS TEMPORALES
# =========================

TEMP_DIR = tempfile.gettempdir()

RUTA_COMPENDIO = os.path.join(
    TEMP_DIR,
    ARCHIVO_COMPENDIO
)

RUTA_PARQUET = os.path.join(
    TEMP_DIR,
    ARCHIVO_PARQUET
)

RUTA_INCIDENCIAS = os.path.join(
    TEMP_DIR,
    ARCHIVO_INCIDENCIAS
)

RUTA_AGENDA = os.path.join(
    TEMP_DIR,
    ARCHIVO_AGENDA
)

# =========================
# GOOGLE DRIVE
# =========================

SCOPES = [
    "https://www.googleapis.com/auth/drive"
]


@st.cache_resource
def obtener_drive_service():

    credenciales = json.loads(
        os.environ[
            "GOOGLE_CREDENTIALS"
        ]
    )

    credentials = (
        service_account.Credentials
        .from_service_account_info(
            credenciales,
            scopes=SCOPES
        )
    )

    service = build(
        "drive",
        "v3",
        credentials=credentials
    )

    return service


drive_service = obtener_drive_service()


def buscar_archivo_drive(
    nombre_archivo,
    folder_id
):

    query = (
        f"'{folder_id}' in parents "
        f"and trashed = false "
        f"and name = '{nombre_archivo}'"
    )

    resultado = (
        drive_service.files()
        .list(
            q=query,
            fields="files(id,name)"
        )
        .execute()
    )

    archivos = resultado.get(
        "files",
        []
    )

    if len(archivos) == 0:

        return None

    return archivos[0]

def descargar_archivo_drive(
    file_id,
    ruta_destino
):

    request = (
        drive_service.files()
        .get_media(
            fileId=file_id
        )
    )

    with open(
        ruta_destino,
        "wb"
    ) as archivo:

        downloader = MediaIoBaseDownload(
            archivo,
            request
        )

        done = False

        while not done:

            status, done = downloader.next_chunk()


def descargar_por_nombre(
    nombre_archivo,
    folder_id,
    ruta_destino,
    obligatorio=True
):

    archivo = buscar_archivo_drive(
        nombre_archivo,
        folder_id
    )

    if archivo is None:

        if obligatorio:

            st.error(
                f"No se encontró en Google Drive: {nombre_archivo}"
            )

            st.stop()

        return False

    descargar_archivo_drive(
        archivo["id"],
        ruta_destino
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


@st.cache_data(
    ttl=600,
    show_spinner="Sincronizando archivos desde Google Drive..."
)
def sincronizar_archivos_drive():

    descargar_archivo_drive(
        "1WW0_0k-1UGX-7qnI4TJNLMwzIWFA6ARs",
        RUTA_PARQUET
    )

    descargar_por_nombre(
        ARCHIVO_INCIDENCIAS,
        FOLDER_ID_INCIDENCIAS,
        RUTA_INCIDENCIAS,
        obligatorio=False
    )

    descargar_por_nombre(
        ARCHIVO_AGENDA,
        FOLDER_ID_BASES,
        RUTA_AGENDA,
        obligatorio=False
    )

    return True


parquet_existe = sincronizar_archivos_drive()

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
    "Operador logístico",
    "IMSS-BIENESTAR",
    "Otro"
]

TIPOS_INCIDENCIA_GENERAL = [
    "Falta de cita",
    "Documentación errónea",
    "Rechazo parcial",
    "Rechazo total",
    "Diferencia de piezas",
    "Producto dañado",
    "Corta caducidad",
    "Entrega fuera de horario",
    "Incumplimiento proveedor",
    "Incumplimiento operador logístico",
    "Otro"
]

# =========================
# GRÁFICAS
# =========================

def grafica_barras(
    df,
    columna,
    titulo,
    limite=None
):

    if columna not in df.columns:

        st.info(
            f"No existe la columna {columna}."
        )

        return

    datos = (
        df[columna]
        .fillna("Sin dato")
        .replace("", "Sin dato")
        .value_counts()
        .reset_index()
    )

    datos.columns = [
        "Categoria",
        "Total"
    ]

    if limite:

        datos = datos.head(
            limite
        )

    chart = (
        alt.Chart(datos)
        .mark_bar(
            cornerRadiusTopLeft=6,
            cornerRadiusTopRight=6
        )
        .encode(
            x=alt.X(
                "Categoria:N",
                sort="-y",
                title=""
            ),
            y=alt.Y(
                "Total:Q",
                title="Total"
            ),
            color=alt.Color(
                "Categoria:N",
                scale=alt.Scale(
                    range=PALETA_INSTITUCIONAL
                ),
                legend=None
            ),
            tooltip=[
                "Categoria",
                "Total"
            ]
        )
        .properties(
            title=titulo,
            height=320
        )
    )

    st.altair_chart(
        chart,
        use_container_width=True
    )


def grafica_resumen(
    total,
    resueltas,
    incompletas
):

    datos = pd.DataFrame(
        {
            "Categoria": [
                "Capturadas",
                "Resueltas",
                "Incompletas"
            ],
            "Total": [
                total,
                resueltas,
                incompletas
            ]
        }
    )

    chart = (
        alt.Chart(datos)
        .mark_bar(
            cornerRadiusTopLeft=6,
            cornerRadiusTopRight=6
        )
        .encode(
            x=alt.X(
                "Categoria:N",
                sort=None,
                title=""
            ),
            y=alt.Y(
                "Total:Q",
                title="Total"
            ),
            color=alt.Color(
                "Categoria:N",
                scale=alt.Scale(
                    range=[
                        COLOR_DORADO,
                        COLOR_VERDE,
                        COLOR_ROJO
                    ]
                ),
                legend=None
            ),
            tooltip=[
                "Categoria",
                "Total"
            ]
        )
        .properties(
            title="Incidencias capturadas, resueltas e incompletas",
            height=320
        )
    )

    st.altair_chart(
        chart,
        use_container_width=True
    )

# =========================
# FUNCIONES BASE
# =========================

def preparar_columnas(
    df
):

    df = df.copy()

    df.columns = (
        df.columns
        .astype(str)
        .str.strip()
        .str.upper()
    )

    return df


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

    tipo = str(tipo_entrega).upper().strip()

    return (
        "OPERADOR" in tipo
        or "LOGISTICO" in tipo
        or "LOGÍSTICO" in tipo
    )

# =========================
# GENERAR PARQUET DESDE XLSB
# =========================

def generar_base_ligera_desde_xlsb():

    descargar_archivo_drive(
        COMPENDIO_FILE_ID,
        RUTA_COMPENDIO
    )

    data = pd.read_excel(
        RUTA_COMPENDIO,
        sheet_name="DATA",
        engine="pyxlsb"
    )

    canceladas = pd.read_excel(
        RUTA_COMPENDIO,
        sheet_name="CANCELADAS",
        engine="pyxlsb"
    )

    data = preparar_columnas(
        data
    )

    canceladas = preparar_columnas(
        canceladas
    )

    canceladas = canceladas.rename(
        columns={
            "NÚMERO DE ORDEN DE SUMINISTRO": "ORDEN DE SUMINISTRO",
            "NUMERO DE ORDEN DE SUMINISTRO": "ORDEN DE SUMINISTRO",
            "RAZÓN SOCIAL": "PROVEEDOR",
            "RAZON SOCIAL": "PROVEEDOR",
            "CLAVE DEL MEDICAMENTO": "CLAVE CNIS",
            "MEDICAMENTO": "DESCRIPCIÓN",
            "CLUES DE DESTINO": "CLUES DESTINO",
            "ENTIDAD DE DESTINO": "ENTIDAD",
            "NOMBRE DE LA UNIDAD": "UNIDAD DESTINO",
            "ALMACÉN DE ENTREGA": "LUGAR DE ENTREGA",
            "ALMACEN DE ENTREGA": "LUGAR DE ENTREGA",
            "CANTIDAD SOLICITADA": "NO. DE PZAS. EMITIDAS",
            "DESCRIPCIÓN DEL ESTATUS DE LA ORDEN DE SUMINISTRO": "ESTATUS DE LA ORDEN DE SUMINISTRO",
            "DESCRIPCION DEL ESTATUS DE LA ORDEN DE SUMINISTRO": "ESTATUS DE LA ORDEN DE SUMINISTRO"
        }
    )

    data["ESTATUS_BASE"] = "ACTIVA"
    canceladas["ESTATUS_BASE"] = "INACTIVA"

    data["ORIGEN_COMPENDIO"] = "DATA"
    canceladas["ORIGEN_COMPENDIO"] = "CANCELADAS"

    columnas_utiles = [
        "ORIGEN_COMPENDIO",
        "ESTATUS_BASE",
        "TIPO DE ENTREGA",
        "ENTIDAD",
        "ESTADO",
        "CLUES DESTINO",
        "CVE CLUES DESTINO",
        "UNIDAD DESTINO",
        "ALMACÉN",
        "ALMACEN",
        "LUGAR DE ENTREGA",
        "PROVEEDOR",
        "ORDEN",
        "ORDEN DE SUMINISTRO",
        "NO. ORDEN",
        "CLAVE",
        "CLAVE CNIS",
        "CLAVE INSUMO",
        "CVE INSUMO",
        "CLAVE DEL INSUMO",
        "DESCRIPCIÓN",
        "DESCRIPCION",
        "PIEZAS",
        "CANTIDAD",
        "TIPO DE RED",
        "GRUPO TERAPEUTICO",
        "GRUPO TERAPÉUTICO",
        "ESTATUS",
        "ESTATUS DE LA ORDEN DE SUMINISTRO",
        "OPERADOR LOGÍSTICO",
        "OPERADOR LOGISTICO",
        "NO. DE PZAS. EMITIDAS",
        "PZAS. RECIBIDAS POR O.L.",
        "PIEZAS REPORTADAS COMO ENTREGADAS CLUES DESTINO"
    ]

    base = pd.concat(
        [
            data,
            canceladas
        ],
        ignore_index=True
    )

    for col in columnas_utiles:

        if col not in base.columns:

            base[col] = ""

    base = base[
        columnas_utiles
    ].copy()

    base.to_parquet(
        RUTA_PARQUET,
        index=False
    )

    subir_archivo_drive(
        RUTA_PARQUET,
        ARCHIVO_PARQUET,
        FOLDER_ID_BASES
    )

    return base
# =========================
# CARGAR DATOS
# =========================

def cargar_compendio_ligero():

    if not os.path.exists(
        RUTA_PARQUET
    ):

        descargar_por_nombre(
            ARCHIVO_PARQUET,
            FOLDER_ID_BASES,
            RUTA_PARQUET,
            obligatorio=True
        )

    return pd.read_parquet(
        RUTA_PARQUET
    )


def cargar_incidencias():

    columnas_necesarias = {
        "FECHA_REGISTRO": "",
        "ORIGEN_REGISTRO": "BASE HISTÓRICA",
        "ORDEN_BUSCADA": "",
        "TIPO_ENTREGA": "",
        "ENTIDAD": "",
        "ALMACEN_CLUES_DESTINO": "",
        "CLUES_DESTINO": "",
        "UNIDAD_DESTINO": "",
        "PROVEEDOR": "",
        "ORDEN": "",
        "CLAVE_CNIS": "",
        "DESCRIPCION": "",
        "PIEZAS": "",
        "TIPO_RED": "",
        "GRUPO_TERAPEUTICO": "",
        "ESTATUS_OPERATIVO": "",
        "PIEZAS_EMITIDAS": "",
        "PIEZAS_RECIBIDAS_OL": "",
        "PIEZAS_ENTREGADAS_CLUES": "",
        "ES_OPERADOR_LOGISTICO": "",
        "ESTATUS_RECEPCION_OL": "",
        "ESTATUS_ENTREGA_ESTADO": "",
        "ESTATUS_INCIDENCIA_COMPLETA": "INCOMPLETA",
        "OPERADOR_LOGISTICO": "",
        "ESTATUS_BASE": "",
        "ORIGEN_COMPENDIO": "",
        "TIPO_INCIDENCIA": "",
        "ATRIBUIBLE A": "",
        "ESTATUS_INCIDENCIA": "Pendiente",
        "RESPONSABLE": "",
        "OBSERVACIONES": "",
        "PDF_CEDULA_RECHAZO": "",
        "PDF_CORREO_SEGUIMIENTO": ""
    }

    if os.path.exists(
        RUTA_INCIDENCIAS
    ):

        incidencias = pd.read_excel(
            RUTA_INCIDENCIAS
        )

        incidencias.columns = (
            incidencias.columns
            .astype(str)
            .str.strip()
            .str.upper()
        )

        for columna, valor_default in columnas_necesarias.items():

            if columna not in incidencias.columns:

                incidencias[columna] = valor_default

        return incidencias

    return pd.DataFrame(
        columns=list(
            columnas_necesarias.keys()
        )
    )


def cargar_agenda_citas():

    if not os.path.exists(
        RUTA_AGENDA
    ):

        return pd.DataFrame()

    agenda = pd.read_excel(
        RUTA_AGENDA
    )

    agenda.columns = (
        agenda.columns
        .astype(str)
        .str.strip()
        .str.upper()
    )

    if "ORDEN DE SUMINISTRO" not in agenda.columns:

        return pd.DataFrame()

    if "FECHA  DE CITA AGENDA" not in agenda.columns:

        return pd.DataFrame()

    agenda["_ORDEN_BUSQUEDA"] = agenda[
        "ORDEN DE SUMINISTRO"
    ].apply(
        normalizar_orden
    )

    agenda["FECHA  DE CITA AGENDA"] = pd.to_datetime(
        agenda["FECHA  DE CITA AGENDA"],
        errors="coerce"
    )

    agenda = agenda.sort_values(
        "FECHA  DE CITA AGENDA"
    )

    agenda = agenda.drop_duplicates(
        "_ORDEN_BUSQUEDA",
        keep="last"
    )

    return agenda


def guardar_incidencia(
    nueva
):

    incidencias = cargar_incidencias()

    incidencias = pd.concat(
        [
            incidencias,
            pd.DataFrame([nueva])
        ],
        ignore_index=True
    )

    incidencias.to_excel(
        RUTA_INCIDENCIAS,
        index=False
    )

    subir_archivo_drive(
        RUTA_INCIDENCIAS,
        ARCHIVO_INCIDENCIAS,
        FOLDER_ID_INCIDENCIAS
    )


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
            fields="files(id,name)"
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
            fields="id"
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


def obtener_carpeta_evidencia(
    estado,
    clues
):

    carpeta_estado_id = obtener_o_crear_carpeta_drive(
        estado,
        FOLDER_ID_EVIDENCIAS
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
            fields="id, webViewLink"
        )
        .execute()
    )

    return nuevo.get(
        "webViewLink",
        ""
    )


# =========================
# BÚSQUEDA Y CRUCE
# =========================

def columnas_orden_disponibles(
    base
):

    columnas_orden = [
        "ORDEN DE SUMINISTRO",
        "ORDEN",
        "NO. ORDEN"
    ]

    return [
        col for col in columnas_orden
        if col in base.columns
    ]


def preparar_base_para_busqueda(
    base
):

    base_temp = base.copy()

    columnas_orden = columnas_orden_disponibles(
        base_temp
    )

    base_temp["_ORDEN_BUSQUEDA"] = ""

    for col in columnas_orden:

        valores = base_temp[col].apply(
            normalizar_orden
        )

        base_temp["_ORDEN_BUSQUEDA"] = (
            base_temp["_ORDEN_BUSQUEDA"]
            .mask(
                base_temp["_ORDEN_BUSQUEDA"].eq("")
                | base_temp["_ORDEN_BUSQUEDA"].eq("NAN"),
                valores
            )
        )

    if "ESTATUS_BASE" in base_temp.columns:

        base_temp["_PRIORIDAD"] = (
            base_temp["ESTATUS_BASE"]
            .apply(
                lambda x:
                0 if str(x).upper().strip() == "INACTIVA"
                else 1
            )
        )

    else:

        base_temp["_PRIORIDAD"] = 1

    return base_temp


def buscar_orden_fuerte(
    base,
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

    base_temp = preparar_base_para_busqueda(
        base
    )

    exacto = base_temp[
        base_temp["_ORDEN_BUSQUEDA"] == valor
    ]

    contiene = base_temp[
        base_temp["_ORDEN_BUSQUEDA"]
        .str.contains(
            valor,
            case=False,
            na=False,
            regex=False
        )
    ]

    resultado = pd.concat(
        [
            exacto,
            contiene
        ],
        ignore_index=True
    )

    if resultado.empty:

        return pd.DataFrame()

    resultado = (
        resultado
        .sort_values(
            "_PRIORIDAD"
        )
        .drop_duplicates()
        .drop(
            columns=[
                "_ORDEN_BUSQUEDA",
                "_PRIORIDAD"
            ],
            errors="ignore"
        )
    )

    return resultado


def sugerir_ordenes(
    base,
    valor_busqueda,
    limite=10
):

    valor = normalizar_orden(
        valor_busqueda
    )

    if (
        not valor
        or valor == "NAN"
    ):

        return pd.DataFrame()

    base_temp = preparar_base_para_busqueda(
        base
    )

    partes = [
        p for p in valor.split("-")
        if len(p) >= 4
    ]

    sugerencias = []

    for parte in partes:

        encontrados = base_temp[
            base_temp["_ORDEN_BUSQUEDA"]
            .str.contains(
                parte,
                case=False,
                na=False,
                regex=False
            )
        ]

        if not encontrados.empty:

            sugerencias.append(
                encontrados
            )

    if not sugerencias:

        return pd.DataFrame()

    salida = pd.concat(
        sugerencias,
        ignore_index=True
    )

    salida = (
        salida
        .sort_values(
            "_PRIORIDAD"
        )
        .drop_duplicates()
        .head(limite)
        .drop(
            columns=[
                "_ORDEN_BUSQUEDA",
                "_PRIORIDAD"
            ],
            errors="ignore"
        )
    )

    return salida


def construir_almacen(
    clues_destino,
    unidad_destino,
    almacen_original
):

    clues = str(
        clues_destino
    ).strip()

    unidad = str(
        unidad_destino
    ).strip()

    if (
        clues
        and clues.lower() != "nan"
        and unidad
        and unidad.lower() != "nan"
    ):

        return f"{clues} - {unidad}"

    if (
        unidad
        and unidad.lower() != "nan"
    ):

        return unidad

    if (
        clues
        and clues.lower() != "nan"
    ):

        return clues

    return almacen_original


def obtener_orden_para_cruce(
    row
):

    opciones = [
        "ORDEN",
        "ORDEN_BUSCADA",
        "ORDEN DE SUMINISTRO",
        "NO. ORDEN"
    ]

    for col in opciones:

        if col in row.index:

            valor = row.get(
                col,
                ""
            )

            if (
                pd.notna(valor)
                and str(valor).strip()
                and str(valor).strip().lower() != "nan"
            ):

                return normalizar_orden(
                    valor
                )

    return ""


def obtener_cita_agenda(
    agenda,
    orden
):

    if agenda.empty:

        return None

    orden_norm = normalizar_orden(
        orden
    )

    encontrado = agenda[
        agenda["_ORDEN_BUSQUEDA"] == orden_norm
    ]

    if encontrado.empty:

        return None

    return encontrado.iloc[0]
def actualizar_incidencias_con_compendio(
    incidencias,
    base
):

    if incidencias.empty:

        return incidencias

    base_temp = preparar_base_para_busqueda(
        base
    )

    base_temp = (
        base_temp
        .sort_values(
            "_PRIORIDAD"
        )
        .drop_duplicates(
            "_ORDEN_BUSQUEDA"
        )
    )

    mapa_base = base_temp.set_index(
        "_ORDEN_BUSQUEDA"
    )

    incidencias = incidencias.copy()

    incidencias = incidencias.astype(
        "object"
    )

    incidencias["_ORDEN_CRUCE"] = incidencias.apply(
        obtener_orden_para_cruce,
        axis=1
    )

    for idx, row in incidencias.iterrows():

        orden_cruce = row[
            "_ORDEN_CRUCE"
        ]

        if (
            not orden_cruce
            or orden_cruce == "NAN"
        ):

            continue

        if orden_cruce not in mapa_base.index:

            continue

        fila = mapa_base.loc[
            orden_cruce
        ]

        tipo_entrega = obtener_valor(
            fila,
            ["TIPO DE ENTREGA"]
        )

        entidad = obtener_valor(
            fila,
            [
                "ENTIDAD",
                "ESTADO"
            ]
        )

        clues_destino = obtener_valor(
            fila,
            [
                "CLUES DESTINO",
                "CVE CLUES DESTINO"
            ]
        )

        unidad_destino = obtener_valor(
            fila,
            ["UNIDAD DESTINO"]
        )

        almacen_original = obtener_valor(
            fila,
            [
                "ALMACÉN",
                "ALMACEN",
                "LUGAR DE ENTREGA"
            ]
        )

        almacen = construir_almacen(
            clues_destino,
            unidad_destino,
            almacen_original
        )

        proveedor = obtener_valor(
            fila,
            ["PROVEEDOR"]
        )

        orden = obtener_valor(
            fila,
            [
                "ORDEN DE SUMINISTRO",
                "ORDEN",
                "NO. ORDEN"
            ]
        )

        clave = obtener_valor(
            fila,
            [
                "CLAVE CNIS",
                "CLAVE",
                "CLAVE INSUMO",
                "CVE INSUMO"
            ]
        )

        descripcion = obtener_valor(
            fila,
            [
                "DESCRIPCIÓN",
                "DESCRIPCION"
            ]
        )

        piezas_emitidas = obtener_valor(
            fila,
            ["NO. DE PZAS. EMITIDAS"]
        )

        piezas_recibidas_ol = obtener_valor(
            fila,
            ["PZAS. RECIBIDAS POR O.L."]
        )

        piezas_entregadas_clues = obtener_valor(
            fila,
            ["PIEZAS REPORTADAS COMO ENTREGADAS CLUES DESTINO"]
        )

        tipo_red = obtener_valor(
            fila,
            ["TIPO DE RED"]
        )

        grupo_terapeutico = obtener_valor(
            fila,
            [
                "GRUPO TERAPEUTICO",
                "GRUPO TERAPÉUTICO"
            ]
        )

        estatus_base = obtener_valor(
            fila,
            ["ESTATUS_BASE"]
        )

        origen_compendio = obtener_valor(
            fila,
            ["ORIGEN_COMPENDIO"]
        )

        operador = obtener_valor(
            fila,
            [
                "OPERADOR LOGÍSTICO",
                "OPERADOR LOGISTICO"
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
            piezas_entregadas_clues
        )

        estatus_completa = calcular_estatus_incidencia_completa(
            estatus_entrega_estado
        )

        estatus_incidencia_auto = (
            "Resuelta"
            if estatus_completa == "COMPLETA"
            else "Pendiente"
        )

        incidencias.at[idx, "TIPO_ENTREGA"] = tipo_entrega
        incidencias.at[idx, "ENTIDAD"] = entidad
        incidencias.at[idx, "ALMACEN_CLUES_DESTINO"] = almacen
        incidencias.at[idx, "CLUES_DESTINO"] = clues_destino
        incidencias.at[idx, "UNIDAD_DESTINO"] = unidad_destino
        incidencias.at[idx, "PROVEEDOR"] = proveedor
        incidencias.at[idx, "ORDEN"] = orden
        incidencias.at[idx, "CLAVE_CNIS"] = clave
        incidencias.at[idx, "DESCRIPCION"] = descripcion
        incidencias.at[idx, "TIPO_RED"] = tipo_red
        incidencias.at[idx, "GRUPO_TERAPEUTICO"] = grupo_terapeutico
        incidencias.at[idx, "ESTATUS_OPERATIVO"] = estatus_base
        incidencias.at[idx, "PIEZAS_EMITIDAS"] = piezas_emitidas
        incidencias.at[idx, "PIEZAS_RECIBIDAS_OL"] = piezas_recibidas_ol
        incidencias.at[idx, "PIEZAS_ENTREGADAS_CLUES"] = piezas_entregadas_clues
        incidencias.at[idx, "ESTATUS_RECEPCION_OL"] = estatus_recepcion_ol
        incidencias.at[idx, "ESTATUS_ENTREGA_ESTADO"] = estatus_entrega_estado
        incidencias.at[idx, "ESTATUS_INCIDENCIA_COMPLETA"] = estatus_completa
        incidencias.at[idx, "ESTATUS_INCIDENCIA"] = estatus_incidencia_auto
        incidencias.at[idx, "OPERADOR_LOGISTICO"] = operador
        incidencias.at[idx, "ESTATUS_BASE"] = estatus_base
        incidencias.at[idx, "ORIGEN_COMPENDIO"] = origen_compendio

    incidencias = incidencias.drop(
        columns=[
            "_ORDEN_CRUCE"
        ],
        errors="ignore"
    )

    return incidencias


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

actualizar_parquet = st.sidebar.button(
    "🔄 Actualizar base ligera desde compendio"
)

if actualizar_parquet:

    with st.spinner(
        "Generando base ligera desde el compendio..."
    ):

        generar_base_ligera_desde_xlsb()

        st.cache_data.clear()

        sincronizar_archivos_drive()

    st.success(
        "Base ligera actualizada correctamente."
    )

    st.rerun()

base = cargar_compendio_ligero()

agenda_citas = cargar_agenda_citas()

incidencias = cargar_incidencias()

incidencias = actualizar_incidencias_con_compendio(
    incidencias,
    base
)

menu = st.sidebar.radio(
    "Menú",
    [
        "Dashboard",
        "Registrar incidencia",
        "Seguimiento",
        "Base ligera"
    ]
)

st.sidebar.divider()

st.sidebar.caption(
    f"Registros base: {len(base):,}"
)

st.sidebar.caption(
    f"Incidencias: {len(incidencias):,}"
)

st.sidebar.caption(
    "Fuente: Google Drive"
)


# =========================
# DASHBOARD
# =========================

if menu == "Dashboard":

    st.subheader(
        "📊 Dashboard ejecutivo"
    )

    total = len(
        incidencias
    )

    completas = incidencias[
        incidencias["ESTATUS_INCIDENCIA_COMPLETA"]
        == "COMPLETA"
    ].shape[0]

    incompletas = incidencias[
        incidencias["ESTATUS_INCIDENCIA_COMPLETA"]
        == "INCOMPLETA"
    ].shape[0]

    resueltas = completas
    activas = incompletas

    porcentaje = 0

    if total > 0:

        porcentaje = round(
            (resueltas / total) * 100,
            2
        )

    c1, c2, c3, c4, c5 = st.columns(
        5
    )

    c1.metric(
        "Incidencias",
        total
    )

    c2.metric(
        "Resueltas",
        resueltas
    )

    c3.metric(
        "Incompletas",
        incompletas
    )

    c4.metric(
        "Activas",
        activas
    )

    c5.metric(
        "% resolución",
        f"{porcentaje}%"
    )

    st.divider()

    if total > 0:

        grafica_barras(
            incidencias,
            "ATRIBUIBLE A",
            "🎯 Atribuible a"
        )

        c1, c2 = st.columns(
            2
        )

        with c1:

            grafica_barras(
                incidencias,
                "PROVEEDOR",
                "🏢 Proveedores con mayor incidencia",
                limite=15
            )

            grafica_resumen(
                total,
                resueltas,
                incompletas
            )

        with c2:

            grafica_barras(
                incidencias,
                "ENTIDAD",
                "🗺️ Estados con mayor incidencia",
                limite=15
            )

            grafica_barras(
                incidencias,
                "ESTATUS_ENTREGA_ESTADO",
                "🚚 Estatus de entrega"
            )

        st.divider()

        st.subheader(
            "📋 Base de incidencias"
        )

        st.dataframe(
            incidencias,
            use_container_width=True
        )

    else:

        st.info(
            "Aún no hay incidencias registradas."
        )


# =========================
# REGISTRO
# =========================

elif menu == "Registrar incidencia":

    st.subheader(
        "📝 Registrar incidencia"
    )

    valor_busqueda = st.text_input(
        "Orden de suministro",
        placeholder="Ejemplo: CIMB-28-01-2025-28030776-U013"
    )

    if valor_busqueda:

        resultado = buscar_orden_fuerte(
            base,
            valor_busqueda
        )

        if len(resultado) == 0:

            st.warning(
                "No encontré esa orden exacta."
            )

            sugerencias = sugerir_ordenes(
                base,
                valor_busqueda
            )

            if len(sugerencias) > 0:

                st.info(
                    "Encontré posibles coincidencias. Revisa si alguna corresponde:"
                )

                st.dataframe(
                    sugerencias,
                    use_container_width=True
                )

            else:

                st.error(
                    "No encontré coincidencias ni en activas ni en canceladas. "
                    "Actualiza la base ligera y vuelve a intentar."
                )

        else:

            resultado = resultado.copy()

            if "ESTATUS_BASE" in resultado.columns:

                resultado["_PRIORIDAD_VISUAL"] = resultado[
                    "ESTATUS_BASE"
                ].apply(
                    lambda x:
                    0 if str(x).upper().strip() == "INACTIVA"
                    else 1
                )

                resultado = resultado.sort_values(
                    "_PRIORIDAD_VISUAL"
                )

                resultado = resultado.drop(
                    columns=[
                        "_PRIORIDAD_VISUAL"
                    ],
                    errors="ignore"
                )

            fila = resultado.iloc[0]

            estatus_base = obtener_valor(
                fila,
                [
                    "ESTATUS_BASE"
                ]
            )

            origen_compendio = obtener_valor(
                fila,
                [
                    "ORIGEN_COMPENDIO"
                ]
            )

            if str(
                estatus_base
            ).upper().strip() == "INACTIVA":

                st.error(
                    "🚫 Esta orden está CANCELADA / INACTIVA en el compendio."
                )

            else:

                st.success(
                    "✅ Esta orden está ACTIVA en el compendio."
                )

            st.success(
                f"Resultados encontrados: {len(resultado)}"
            )

            st.subheader(
                "📋 Datos encontrados en compendio"
            )

            st.dataframe(
                resultado.head(
                    50
                ),
                use_container_width=True
            )

            with st.expander(
                "🔎 Ver todos los datos completos del compendio"
            ):

                st.dataframe(
                    resultado,
                    use_container_width=True
                )

            orden = obtener_valor(
                fila,
                [
                    "ORDEN DE SUMINISTRO",
                    "ORDEN",
                    "NO. ORDEN"
                ]
            )

            cita = obtener_cita_agenda(
                agenda_citas,
                orden
            )

            incidencias_previas = incidencias[
                incidencias["ORDEN"]
                .astype(str)
                .apply(normalizar_orden)
                ==
                normalizar_orden(orden)
            ]

            if len(incidencias_previas) > 0:

                st.warning(
                    f"⚠️ Esta orden ya tiene {len(incidencias_previas)} incidencia(s) registrada(s)."
                )

                st.dataframe(
                    incidencias_previas,
                    use_container_width=True
                )

            tipo_entrega = obtener_valor(
                fila,
                [
                    "TIPO DE ENTREGA"
                ]
            )

            entidad = obtener_valor(
                fila,
                [
                    "ENTIDAD",
                    "ESTADO"
                ]
            )

            clues_destino = obtener_valor(
                fila,
                [
                    "CLUES DESTINO",
                    "CVE CLUES DESTINO"
                ]
            )

            unidad_destino = obtener_valor(
                fila,
                [
                    "UNIDAD DESTINO"
                ]
            )

            almacen_original = obtener_valor(
                fila,
                [
                    "ALMACÉN",
                    "ALMACEN",
                    "LUGAR DE ENTREGA"
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
                    "PROVEEDOR"
                ]
            )

            clave = obtener_valor(
                fila,
                [
                    "CLAVE CNIS",
                    "CLAVE",
                    "CLAVE INSUMO",
                    "CVE INSUMO"
                ]
            )

            descripcion = obtener_valor(
                fila,
                [
                    "DESCRIPCIÓN",
                    "DESCRIPCION"
                ]
            )

            piezas_emitidas = obtener_valor(
                fila,
                [
                    "NO. DE PZAS. EMITIDAS"
                ]
            )

            piezas_recibidas_ol = obtener_valor(
                fila,
                [
                    "PZAS. RECIBIDAS POR O.L."
                ]
            )

            piezas_entregadas = obtener_valor(
                fila,
                [
                    "PIEZAS REPORTADAS COMO ENTREGADAS CLUES DESTINO"
                ]
            )

            operador = obtener_valor(
                fila,
                [
                    "OPERADOR LOGÍSTICO",
                    "OPERADOR LOGISTICO"
                ]
            )

            tipo_red = obtener_valor(
                fila,
                [
                    "TIPO DE RED"
                ]
            )

            grupo_terapeutico = obtener_valor(
                fila,
                [
                    "GRUPO TERAPEUTICO",
                    "GRUPO TERAPÉUTICO"
                ]
            )

            estatus_orden_compendio = obtener_valor(
                fila,
                [
                    "ESTATUS DE LA ORDEN DE SUMINISTRO",
                    "ESTATUS"
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

            st.divider()

            st.subheader(
                "📅 Agenda de cita"
            )

            if cita is None:

                st.warning(
                    "🚫 Esta orden no tiene cita agendada."
                )

            else:

                fecha_cita = cita.get(
                    "FECHA  DE CITA AGENDA",
                    ""
                )

                estatus_agenda = cita.get(
                    "ESTATUS AGENDA",
                    ""
                )

                numero_reagendas = cita.get(
                    "NÚMERO DE REAGENDAS",
                    cita.get(
                        "NUMERO DE REAGENDAS",
                        0
                    )
                )

                ultima_reagenda = cita.get(
                    "ÚLTIMA REAGENDA",
                    cita.get(
                        "ULTIMA REAGENDA",
                        ""
                    )
                )

                c_ag1, c_ag2, c_ag3, c_ag4 = st.columns(
                    4
                )

                c_ag1.text_input(
                    "Última cita",
                    "" if pd.isna(fecha_cita)
                    else fecha_cita.strftime("%d/%m/%Y"),
                    disabled=True
                )

                c_ag2.text_input(
                    "Estatus agenda",
                    str(estatus_agenda),
                    disabled=True
                )

                c_ag3.text_input(
                    "# Reagendas",
                    str(numero_reagendas),
                    disabled=True
                )

                c_ag4.text_input(
                    "Última reagenda",
                    "" if pd.isna(ultima_reagenda)
                    else str(ultima_reagenda),
                    disabled=True
                )

            st.divider()

            st.subheader(
                "🔒 Información del compendio"
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
                estatus_orden_compendio,
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

            st.subheader(
                "🚚 Validación logística"
            )

            if tiene_ol:

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

            else:

                c12, c13 = st.columns(
                    2
                )

                c12.metric(
                    "Piezas emitidas",
                    piezas_emitidas
                )

                c13.metric(
                    "Piezas entregadas CLUES",
                    piezas_entregadas
                )

                st.text_input(
                    "Estatus entrega",
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
                        "Pendiente",
                        "En proceso",
                        "Escalado",
                        "Resuelta",
                        "Cancelada"
                    ]
                )

            responsable = st.text_input(
                "Responsable"
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

                nueva = {
                    "FECHA_REGISTRO": datetime.now(),
                    "ORIGEN_REGISTRO": "SISTEMA",
                    "ORDEN_BUSCADA": valor_busqueda,
                    "ORDEN": orden,
                    "TIPO_ENTREGA": tipo_entrega,
                    "ENTIDAD": entidad,
                    "ALMACEN_CLUES_DESTINO": almacen,
                    "CLUES_DESTINO": clues_destino,
                    "UNIDAD_DESTINO": unidad_destino,
                    "PROVEEDOR": proveedor,
                    "CLAVE_CNIS": clave,
                    "DESCRIPCION": descripcion,
                    "PIEZAS_EMITIDAS": piezas_emitidas,
                    "PIEZAS_RECIBIDAS_OL": piezas_recibidas_ol,
                    "PIEZAS_ENTREGADAS_CLUES": piezas_entregadas,
                    "TIPO_RED": tipo_red,
                    "GRUPO_TERAPEUTICO": grupo_terapeutico,
                    "ESTATUS_OPERATIVO": estatus_base,
                    "ESTATUS_BASE": estatus_base,
                    "ORIGEN_COMPENDIO": origen_compendio,
                    "OPERADOR_LOGISTICO": operador,
                    "ESTATUS_RECEPCION_OL": estatus_recepcion_ol,
                    "ESTATUS_ENTREGA_ESTADO": estatus_entrega_estado,
                    "ESTATUS_INCIDENCIA_COMPLETA": estatus_completa,
                    "ATRIBUIBLE A": atribuible,
                    "TIPO_INCIDENCIA": tipo,
                    "ESTATUS_INCIDENCIA": estatus,
                    "RESPONSABLE": responsable,
                    "OBSERVACIONES": observaciones,
                    "PDF_CEDULA_RECHAZO": ruta_cedula,
                    "PDF_CORREO_SEGUIMIENTO": ruta_correo
                }

                guardar_incidencia(
                    nueva
                )

                st.success(
                    "Incidencia guardada correctamente en Google Drive."
                )

                st.rerun()


# =========================
# SEGUIMIENTO
# =========================

elif menu == "Seguimiento":

    st.subheader(
        "📋 Seguimiento"
    )

    st.dataframe(
        incidencias,
        use_container_width=True
    )


# =========================
# BASE LIGERA
# =========================

elif menu == "Base ligera":

    st.subheader(
        "⚡ Base ligera"
    )

    st.dataframe(
        base.head(100),
        use_container_width=True
    )