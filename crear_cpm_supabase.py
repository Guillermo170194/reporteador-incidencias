import pandas as pd

RUTA = r"C:\Users\guillermo.ortega\Downloads\CPM CLUES Demanda IB Nov 2025 20251127 ACTUAL (4) (1).xlsb"

df = pd.read_excel(
    RUTA,
    sheet_name="Data",
    engine="pyxlsb",
    header=1
)

df.columns = (
    df.columns
    .astype(str)
    .str.strip()
    .str.replace("\n", " ")
)

renombres = {
    "Entidad": "entidad",
    "Clues SSA": "clues_ssa",
    "Clues IMSS-B": "clues_imss_b",
    "Unidad": "unidad",
    "Tipo": "tipo",
    "Clave CNIS": "clave_cnis",
    "Descripción": "descripcion",
    "CPM": "cpm",
    "Gpo Ter": "grupo_terapeutico",
    "P.U": "precio_unitario",
    "IVA": "iva",
    "Importe": "importe"
}

df = df.rename(
    columns=renombres
)

columnas = [
    "entidad",
    "clues_ssa",
    "clues_imss_b",
    "unidad",
    "tipo",
    "clave_cnis",
    "descripcion",
    "cpm",
    "grupo_terapeutico",
    "precio_unitario",
    "iva",
    "importe"
]

df = df[
    [c for c in columnas if c in df.columns]
].copy()

df = df.fillna("")

df["clues_busqueda"] = df["clues_imss_b"]

df.loc[
    (df["clues_busqueda"].astype(str).str.strip() == "")
    | (df["clues_busqueda"].astype(str).str.strip() == "-"),
    "clues_busqueda"
] = df["clues_ssa"]

df.to_csv(
    "cpm_clues_supabase.csv",
    index=False,
    encoding="utf-8-sig"
)

print("CSV CPM generado correctamente")
print("Filas:", len(df))
print(df.columns.tolist())
print(df.head())	