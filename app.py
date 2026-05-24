import streamlit as st
import pandas as pd
import numpy as np
import altair as alt
import re
import os
import json
import tempfile

from datetime import datetime
from supabase import create_client
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload


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


# =========================
# ARCHIVOS DRIVE
# =========================

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
    show_spinner="Sincronizando agenda e incidencias..."
)
def sincronizar_archivos_drive():

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


sincronizar_archivos_drive()
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

    respuesta = (
        supabase
        .table(
            "compendio"
        )
        .select(
            "orden_suministro, orden, no_orden, estatus_base, entidad, proveedor, clave_cnis, descripcion"
        )
        .ilike(
            "orden_suministro",
            f"{valor}%"
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
# =========================
# CARGAR DATOS
# =========================

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

    agenda["_ORDEN_BUSQUEDA"] = agenda[
        "ORDEN DE SUMINISTRO"
    ].apply(
        normalizar_orden
    )

    if "FECHA  DE CITA AGENDA" in agenda.columns:

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


def obtener_incidencias_previas(
    incidencias,
    orden
):

    if incidencias.empty:

        return pd.DataFrame()

    if "ORDEN" not in incidencias.columns:

        return pd.DataFrame()

    orden_norm = normalizar_orden(
        orden
    )

    previas = incidencias[
        incidencias["ORDEN"]
        .astype(str)
        .apply(normalizar_orden)
        ==
        orden_norm
    ]

    return previas


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

incidencias = cargar_incidencias()

agenda_citas = cargar_agenda_citas()

menu = st.sidebar.radio(
    "Menú",
    [
        "Dashboard",
        "Registrar incidencia",
        "Seguimiento",
        "Base Supabase"
    ]
)

st.sidebar.divider()

st.sidebar.caption(
    f"Incidencias registradas: {len(incidencias):,}"
)

st.sidebar.caption(
    "Fuente: Supabase + Google Drive"
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

    if total == 0:

        st.info(
            "Aún no hay incidencias registradas."
        )

    else:

        completas = incidencias[
            incidencias["ESTATUS_INCIDENCIA_COMPLETA"]
            == "COMPLETA"
        ].shape[0]

        incompletas = incidencias[
            incidencias["ESTATUS_INCIDENCIA_COMPLETA"]
            == "INCOMPLETA"
        ].shape[0]

        resueltas = incidencias[
            incidencias["ESTATUS_INCIDENCIA"]
            .astype(str)
            .str.upper()
            .eq("RESUELTA")
        ].shape[0]

        porcentaje = 0

        if total > 0:

            porcentaje = round(
                (resueltas / total) * 100,
                2
            )

        c1, c2, c3, c4 = st.columns(
            4
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
            "% resolución",
            f"{porcentaje}%"
        )

        st.divider()

        st.dataframe(
            incidencias,
            use_container_width=True
        )


# =========================
# REGISTRAR INCIDENCIA
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

        with st.spinner(
            "Buscando orden en Supabase..."
        ):

            resultado = buscar_orden_fuerte(
                valor_busqueda
            )

        if len(resultado) == 0:

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

                st.dataframe(
                    sugerencias,
                    use_container_width=True
                )

            else:

                st.error(
                    "No encontré coincidencias en Supabase."
                )

        else:

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

            cita = obtener_cita_agenda(
                agenda_citas,
                orden
            )

            incidencias_previas = obtener_incidencias_previas(
                incidencias,
                orden
            )

            st.divider()

            c_estado, c_cita, c_previas = st.columns(
                3
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

            with c_cita:

                if cita is None:

                    st.warning(
                        "📅 Sin cita localizada"
                    )

                else:

                    fecha_cita = obtener_valor(
                        cita,
                        [
                            "FECHA  DE CITA AGENDA"
                        ]
                    )

                    st.success(
                        f"📅 Cita: {fecha_a_texto(fecha_cita)}"
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

                    st.dataframe(
                        incidencias_previas,
                        use_container_width=True
                    )

            st.subheader(
                "📋 Datos encontrados en Supabase"
            )

            st.dataframe(
                resultado.head(
                    50
                ),
                use_container_width=True
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
                    "Incidencia guardada correctamente."
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

    st.dataframe(
        muestra,
        use_container_width=True
    )