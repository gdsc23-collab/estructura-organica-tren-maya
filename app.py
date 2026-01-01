import streamlit as st
import pandas as pd
import plotly.express as px
from PIL import Image

# =========================
# CONFIGURACIÓN GENERAL
# =========================
st.set_page_config(
    page_title="Tablero EO – Tren Maya",
    layout="wide"
)

# Paleta institucional (pedida)
TM_WINE = "#691C32"
TM_GOLD = "#D4C19C"
TM_GREEN = "#00524C"
TM_COLORS = [TM_WINE, TM_GOLD, TM_GREEN]

# Orden jerárquico Tipo de Puesto (pedido)
ORDEN_PUESTO = [
    "Dirección General",
    "Titular de Unidad",
    "Coordinador General",
    "Director de área",
    "Gerente",
    "Subgerente",
    "Enlace",
    "Operativo"
]

# =========================
# UTILIDADES
# =========================
def load_logo():
    """Carga logo local si existe."""
    try:
        return Image.open("logo_tm.png")
    except Exception:
        return None


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Normaliza encabezados para evitar KeyError por saltos de línea/espacios."""
    df.columns = (
        df.columns.astype(str)
        .str.replace("\n", " ", regex=False)
        .str.replace("\r", " ", regex=False)
        .str.replace("\t", " ", regex=False)
        .str.replace("  ", " ", regex=False)
        .str.strip()
    )
    return df


def apply_aliases(df: pd.DataFrame) -> pd.DataFrame:
    """
    Renombra columnas a nombres canónicos aunque vengan con variantes
    (acentos, mayúsculas, guiones, etc.).
    """
    # Limpieza extra para matching
    cols_clean = {c: c for c in df.columns}

    def norm(s: str) -> str:
        s = str(s).strip().lower()
        s = s.replace("\n", " ").replace("\r", " ").replace("\t", " ")
        while "  " in s:
            s = s.replace("  ", " ")
        return s

    col_norm_map = {norm(c): c for c in df.columns}

    aliases = {
        "Unidad": ["unidad"],
        "Coordinación": ["coordinación", "coordinacion"],
        "Dirección": ["dirección", "direccion"],
        "Tipo de Puesto": ["tipo de puesto", "tipo puesto", "tipo_puesto"],
        "Nivel Salarial": ["nivel salarial", "nivel_salarial"],
        "Tipo de Plaza": ["tipo de plaza", "tipo plaza", "tipo_plaza"],
        "Justificación": ["justificación", "justificacion"],
        "Plazas": ["plazas", "no. plazas", "número de plazas", "numero de plazas", "num plazas"],
        # (Opcional) Si existe en tu BD y luego quieres usarlo:
        "PLAZAS OCUPADAS/VACANTES": [
            "plazas ocupadas/vacantes",
            "plazas ocupadas vacantes",
            "ocupadas/vacantes",
            "plazas ocupadas",
            "vacantes"
        ],
    }

    rename_map = {}

    for canonical, opts in aliases.items():
        found_original = None
        for opt in opts:
            key = norm(opt)
            if key in col_norm_map:
                found_original = col_norm_map[key]
                break
        if found_original and found_original in df.columns:
            rename_map[found_original] = canonical

    df = df.rename(columns=rename_map)
    return df


def ensure_required(df: pd.DataFrame):
    required = [
        "Unidad",
        "Coordinación",
        "Dirección",
        "Tipo de Puesto",
        "Nivel Salarial",
        "Tipo de Plaza",
        "Justificación",
        "Plazas",
    ]
    missing = [c for c in required if c not in df.columns]
    if missing:
        st.error(f"Faltan columnas requeridas en BD-EO: {missing}")
        st.write("Columnas detectadas en tu archivo:")
        st.write(list(df.columns))
        st.stop()


def safe_text_series(s: pd.Series) -> pd.Series:
    """Convierte a texto y limpia."""
    return s.astype(str).str.strip()


def coerce_plazas_numeric(df: pd.DataFrame) -> pd.DataFrame:
    """Asegura que Plazas sea numérico."""
    df["Plazas"] = pd.to_numeric(df["Plazas"], errors="coerce")
    return df


def drop_total_like_rows(df: pd.DataFrame) -> pd.DataFrame:
    """
    Intenta eliminar filas de totales/encabezados repetidos:
    - Filas donde 'Dirección' o 'Tipo de Plaza' parezcan "TOTAL"
    - Filas donde columnas clave estén vacías
    """
    # Quitar filas totalmente vacías en columnas clave
    df = df[df["Tipo de Plaza"].notna()]

    # Quitar filas "TOTAL"
    for col in ["Dirección", "Tipo de Plaza", "Unidad", "Coordinación"]:
        if col in df.columns:
            mask_total = safe_text_series(df[col]).str.upper().isin(["TOTAL", "TOTALES"])
            df = df[~mask_total]

    return df


# =========================
# UI - LOGO Y TÍTULO
# =========================
logo = load_logo()
if logo:
    st.image(logo, width=360)

st.title("Análisis de la Estructura Orgánica – Tren Maya")

uploaded_file = st.file_uploader(
    "Cargar archivo Excel BD-EO",
    type=["xlsx", "xlsm"]
)

if not uploaded_file:
    st.info("Por favor carga el archivo Excel BD-EO para visualizar el tablero.")
    st.stop()

# =========================
# LECTURA EXCEL (header fila 7 = header=6)
# =========================
try:
    df = pd.read_excel(uploaded_file, sheet_name="BD-EO", header=6)
except Exception as e:
    st.error("No pude leer la hoja 'BD-EO'. Verifica que exista en el archivo.")
    st.exception(e)
    st.stop()

# Normalizar + aliases para evitar KeyError
df = normalize_columns(df)
df = apply_aliases(df)
ensure_required(df)

# Limpieza datos
df = coerce_plazas_numeric(df)
df = drop_total_like_rows(df)

# =========================
# FILTROS
# =========================
st.sidebar.header("Filtros")

unidad_opts = sorted(df["Unidad"].dropna().unique())
coord_opts = sorted(df["Coordinación"].dropna().unique())
dir_opts = sorted(df["Dirección"].dropna().unique())

unidad_sel = st.sidebar.multiselect("Unidad", options=unidad_opts)
coord_sel = st.sidebar.multiselect("Coordinación", options=coord_opts)
dir_sel = st.sidebar.multiselect("Dirección", options=dir_opts)

df_f = df.copy()
if unidad_sel:
    df_f = df_f[df_f["Unidad"].isin(unidad_sel)]
if coord_sel:
    df_f = df_f[df_f["Coordinación"].isin(coord_sel)]
if dir_sel:
    df_f = df_f[df_f["Dirección"].isin(dir_sel)]

# =========================
# TABS
# =========================
tab1, tab2 = st.tabs(["📊 Tablero", "📝 Justificación"])

# =========================
# TAB 1: TABLERO
# =========================
with tab1:
    if logo:
        st.image(logo, width=300)

    # Primera fila (Dirección horizontal + Tipo de Puesto ordenado)
    c1, c2 = st.columns([1.2, 1])

    # 1) Colaboradores por Dirección (horizontal, legible)
    dir_count = (
        df_f["Dirección"]
        .fillna("Sin Dirección")
        .value_counts()
        .reset_index()
    )
    dir_count.columns = ["Dirección", "Colaboradores"]

    # Para lectura, limitamos a top 25 (si hay más, igual queda legible)
    top_n = 25
    dir_count_plot = dir_count.head(top_n)

    fig_dir = px.bar(
        dir_count_plot,
        x="Colaboradores",
        y="Dirección",
        orientation="h",
        text="Colaboradores",
        color_discrete_sequence=[TM_WINE],
        title=f"Colaboradores por Dirección (Top {top_n})" if len(dir_count) > top_n else "Colaboradores por Dirección"
    )
    fig_dir.update_traces(textposition="outside", cliponaxis=False)
    fig_dir.update_layout(
        yaxis={"categoryorder": "total ascending"},
        margin=dict(l=10, r=10, t=50, b=10),
        height=520
    )

    c1.plotly_chart(fig_dir, use_container_width=True)

    # 2) Colaboradores por Tipo de Puesto (orden jerárquico)
    puesto_series = df_f["Tipo de Puesto"].fillna("Sin Tipo de Puesto")

    puesto_count = (
        puesto_series.value_counts()
        .reindex(ORDEN_PUESTO)
        .dropna()
        .reset_index()
    )
    puesto_count.columns = ["Tipo de Puesto", "Colaboradores"]

    # Si existen tipos fuera del orden (poco probable), los agregamos al final:
    otros = (
        puesto_series.value_counts()
        .drop(index=ORDEN_PUESTO, errors="ignore")
        .reset_index()
    )
    if not otros.empty:
        otros.columns = ["Tipo de Puesto", "Colaboradores"]
        puesto_count = pd.concat([puesto_count, otros], ignore_index=True)

    fig_puesto = px.bar(
        puesto_count,
        x="Tipo de Puesto",
        y="Colaboradores",
        text="Colaboradores",
        color_discrete_sequence=[TM_GOLD],
        title="Colaboradores por Tipo de Puesto"
    )
    fig_puesto.update_traces(textposition="outside", cliponaxis=False)
    fig_puesto.update_layout(margin=dict(l=10, r=10, t=50, b=10), height=520)

    c2.plotly_chart(fig_puesto, use_container_width=True)

    # Segunda fila (Nivel Salarial + Tipo de Plaza donut)
    c3, c4 = st.columns(2)

    # 3) Nivel Salarial (barras)
    nivel_count = (
        df_f["Nivel Salarial"]
        .fillna("Sin Nivel Salarial")
        .value_counts()
        .reset_index()
    )
    nivel_count.columns = ["Nivel Salarial", "Colaboradores"]

    fig_nivel = px.bar(
        nivel_count,
        x="Nivel Salarial",
        y="Colaboradores",
        text="Colaboradores",
        color_discrete_sequence=[TM_GREEN],
        title="Colaboradores por Nivel Salarial"
    )
    fig_nivel.update_traces(textposition="outside", cliponaxis=False)
    fig_nivel.update_layout(margin=dict(l=10, r=10, t=50, b=10), height=520)

    c3.plotly_chart(fig_nivel, use_container_width=True)

    # 4) Tipo de Plaza (donut, 3 colores + valores)
    plaza_count = (
        df_f["Tipo de Plaza"]
        .fillna("Sin Tipo de Plaza")
        .value_counts()
        .reset_index()
    )
    plaza_count.columns = ["Tipo de Plaza", "Total"]

    fig_plaza = px.pie(
        plaza_count,
        names="Tipo de Plaza",
        values="Total",
        hole=0.55,
        color_discrete_sequence=TM_COLORS,
        title="Colaboradores por Tipo de Plaza"
    )
    fig_plaza.update_traces(
        textinfo="value+percent",
        textposition="outside"
    )
    fig_plaza.update_layout(margin=dict(l=10, r=10, t=50, b=10), height=520)

    c4.plotly_chart(fig_plaza, use_container_width=True)

# =========================
# TAB 2: JUSTIFICACIÓN (NO CAMBIAR LÓGICA)
# =========================
with tab2:
    if logo:
        st.image(logo, width=300)

    # Solo registros con justificación (filtrado)
    df_j = df_f.copy()
    df_j["Justificación"] = df_j["Justificación"].astype(str)

    df_j = df_j[df_j["Justificación"].notna()]
    df_j = df_j[df_j["Justificación"].str.strip() != ""]
    df_j = df_j[~df_j["Justificación"].str.strip().str.upper().isin(["NAN", "NONE"])]

    # KPIs (como lo pediste)
    colA, colB, colC = st.columns(3)
    colA.metric("Plazas (suma)", int(pd.to_numeric(df_j["Plazas"], errors="coerce").fillna(0).sum()))
    colB.metric("Registros con justificación", int(df_j.shape[0]))
    colC.metric("Registros (detalle)", int(df_f.shape[0]))

    # Tabla SOLO justificadas
    st.dataframe(
        df_j[[
            "Unidad",
            "Coordinación",
            "Dirección",
            "Tipo de Puesto",
            "Tipo de Plaza",
            "Plazas",
            "Justificación"
        ]],
        use_container_width=True,
        hide_index=True
    )
