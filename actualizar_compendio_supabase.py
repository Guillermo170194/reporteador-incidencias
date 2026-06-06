import argparse
import os
import re
from pathlib import Path

import pandas as pd

ARCHIVO_XLSB = Path("CompendioAbasto25-26_22.05.2026.xlsb")
SALIDA_CSV = Path("compendio_supabase.csv")
TABLA_SUPABASE = "compendio_abasto"

COLUMNAS_SALIDA = [
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


def normalizar_columna(valor):

    texto = str(valor).strip().upper()

    reemplazos = {
        "Á": "A",
        "É": "E",
        "Í": "I",
        "Ó": "O",
        "Ú": "U",
        "Ü": "U"
    }

    for origen, destino in reemplazos.items():

        texto = texto.replace(
            origen,
            destino
        )

    texto = re.sub(
        r"\s+",
        " ",
        texto
    )

    return texto


def obtener_serie(df, nombres):

    mapa = {
        normalizar_columna(c): c
        for c in df.columns
    }

    for nombre in nombres:

        nombre_norm = normalizar_columna(
            nombre
        )

        if nombre_norm in mapa:

            return df[
                mapa[nombre_norm]
            ]

    return pd.Series(
        [""] * len(df),
        index=df.index
    )


def limpiar_texto(valor):

    if pd.isna(valor):
        return ""

    texto = str(valor).strip()

    if texto.lower() in [
        "nan",
        "none",
        "null",
        "nat"
    ]:
        return ""

    return texto


def limpiar_orden(valor):

    texto = limpiar_texto(
        valor
    ).upper()

    texto = (
        texto
        .replace("_", "-")
        .replace(" ", "")
        .replace(".PDF", "")
        .replace("–", "-")
        .replace("—", "-")
    )

    return texto


def numero(valor):

    if pd.isna(valor):
        return 0

    texto = str(valor).strip()

    if texto in [
        "",
        "-",
        "nan",
        "None"
    ]:
        return 0

    texto = texto.replace(",", "")

    try:

        return float(texto)

    except Exception:

        return 0


def texto(df, nombres):

    return (
        obtener_serie(
            df,
            nombres
        )
        .fillna("")
        .astype(str)
        .str.strip()
    )


def leer_hoja(nombre_hoja, estatus_base):

    print(f"Leyendo hoja: {nombre_hoja}")

    df = pd.read_excel(
        ARCHIVO_XLSB,
        sheet_name=nombre_hoja,
        engine="pyxlsb"
    )

    df.columns = [
        str(c).strip()
        for c in df.columns
    ]

    salida = pd.DataFrame()

    salida["orden"] = obtener_serie(
        df,
        [
            "Orden de Suministro",
            "Número de orden de suministro",
            "ORDEN DE SUMINISTRO",
            "orden_suministro",
            "ORDEN"
        ]
    ).apply(limpiar_orden)

    salida["orden_suministro"] = salida["orden"]
    salida["no_orden"] = salida["orden"]

    salida["estatus_base"] = estatus_base
    salida["origen_compendio"] = nombre_hoja

    salida["entidad"] = texto(
        df,
        [
            "Entidad",
            "Entidad de destino",
            "Estado",
            "ENTIDAD"
        ]
    )

    salida["estado"] = salida["entidad"]

    salida["clues_destino"] = texto(
        df,
        [
            "CLUES Destino",
            "CLUES de destino",
            "CLUES destino",
            "CLUES"
        ]
    )

    salida["unidad_destino"] = texto(
        df,
        [
            "Unidad Destino",
            "Nombre de la unidad",
            "Unidad"
        ]
    )

    salida["almacen"] = (
        salida["clues_destino"].fillna("").astype(str)
        + " - "
        + salida["unidad_destino"].fillna("").astype(str)
    ).str.strip(" -")

    salida["proveedor"] = texto(
        df,
        [
            "Proveedor ",
            "Proveedor",
            "Razón social",
            "Razon social"
        ]
    )

    salida["clave_cnis"] = texto(
        df,
        [
            "Clave CNIS",
            "Clave del medicamento",
            "CLAVE CNIS"
        ]
    )

    salida["descripcion"] = texto(
        df,
        [
            "Descripción",
            "Descripcion",
            "Medicamento"
        ]
    )

    salida["tipo_entrega"] = texto(
        df,
        [
            "Tipo de entrega",
            "TIPO DE ENTREGA"
        ]
    )

    salida["piezas_emitidas"] = obtener_serie(
        df,
        [
            "No. de pzas. Emitidas",
            "Cantidad solicitada",
            "Piezas emitidas"
        ]
    ).apply(numero)

    salida["piezas_recibidas_ol"] = obtener_serie(
        df,
        [
            "Pzas. Recibidas por O.L.",
            "Piezas Recibidas OL"
        ]
    ).apply(numero)

    salida["piezas_entregadas_clues"] = obtener_serie(
        df,
        [
            "Piezas Reportadas como entregadas CLUES Destino",
            "Piezas entregadas CLUES"
        ]
    ).apply(numero)

    salida["tipo_red"] = texto(
        df,
        [
            "Tipo de Red",
            "Tipo red"
        ]
    )

    salida["grupo_terapeutico"] = texto(
        df,
        [
            "Grupo Terapéutico",
            "Grupo Terapeutico"
        ]
    )

    salida["estatus"] = texto(
        df,
        [
            "Estatus de la orden de suministro",
            "Descripción del estatus de la orden de suministro",
            "Estatus"
        ]
    )

    salida = salida[
        salida["orden"] != ""
    ].copy()

    for columna in COLUMNAS_SALIDA:

        if columna not in salida.columns:

            salida[columna] = ""

    return salida[
        COLUMNAS_SALIDA
    ]


def generar_compendio():

    if not ARCHIVO_XLSB.exists():

        raise FileNotFoundError(
            f"No existe: {ARCHIVO_XLSB.resolve()}"
        )

    data = leer_hoja(
        "DATA",
        "ACTIVA"
    )

    canceladas = leer_hoja(
        "CANCELADAS",
        "INACTIVA"
    )

    compendio = pd.concat(
        [
            data,
            canceladas
        ],
        ignore_index=True
    )

    compendio = compendio.drop_duplicates(
        subset=[
            "orden",
            "estatus_base",
            "origen_compendio"
        ],
        keep="first"
    )

    compendio.to_csv(
        SALIDA_CSV,
        index=False,
        encoding="utf-8-sig"
    )

    print("")
    print("CSV generado:")
    print(SALIDA_CSV.resolve())
    print("")
    print(compendio["estatus_base"].value_counts())

    return compendio


def subir_supabase(compendio, tabla, limpiar=False):

    try:

        from supabase import create_client

    except Exception:

        raise Exception(
            "Instala supabase con: py -3.11 -m pip install supabase==1.2.0"
        )

    url = os.getenv(
        "SUPABASE_URL",
        ""
    ).strip()

    key = os.getenv(
        "SUPABASE_KEY",
        ""
    ).strip()

    if not url or not key:

        raise Exception(
            "Faltan SUPABASE_URL o SUPABASE_KEY"
        )

    client = create_client(
        url,
        key
    )

    if limpiar:

        print("Limpiando tabla...")

        client.table(
            tabla
        ).delete().neq(
            "orden",
            "__NO_EXISTE__"
        ).execute()

    registros = compendio.where(
        pd.notnull(compendio),
        None
    ).to_dict(
        orient="records"
    )

    total = len(registros)
    chunk = 500

    for inicio in range(
        0,
        total,
        chunk
    ):

        fin = min(
            inicio + chunk,
            total
        )

        bloque = registros[
            inicio:fin
        ]

        client.table(
            tabla
        ).insert(
            bloque
        ).execute()

        print(
            f"Subidas {fin:,} de {total:,}"
        )

    print("Carga terminada.")


def main():

    parser = argparse.ArgumentParser()

    parser.add_argument(
        "--upload",
        action="store_true"
    )

    parser.add_argument(
        "--limpiar-tabla",
        action="store_true"
    )

    parser.add_argument(
        "--archivo",
        default=str(ARCHIVO_XLSB)
    )

    parser.add_argument(
        "--tabla",
        default=TABLA_SUPABASE
    )

    args = parser.parse_args()

    archivo_xlsb = Path(
        args.archivo
    )

    tabla_supabase = args.tabla

    globals()["ARCHIVO_XLSB"] = archivo_xlsb

    compendio = generar_compendio()

    if args.upload:

        subir_supabase(
            compendio,
            tabla_supabase,
            limpiar=args.limpiar_tabla
        )


if __name__ == "__main__":

    main()
