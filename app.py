"""
Limpieza Nutrición Inferiores
Club Atlético Unión (Santa Fe)

App Streamlit gratuita para que los nutricionistas del club carguen la
planilla de Excel de cada categoría, filtren por hoja/fecha de evaluación,
revisen y confirmen los jugadores, y exporten los datos ya limpios y
normalizados en el formato de la base de datos NUTRI_LONG.
"""

import io
import re
import unicodedata
from datetime import date
from pathlib import Path

import pandas as pd
import streamlit as st

# --------------------------------------------------------------------------
# Configuración general / estilo del club (rojo y negro)
# --------------------------------------------------------------------------

CLUB_RED = "#D2001F"
CLUB_BLACK = "#111111"

st.set_page_config(
    page_title="Limpieza Nutrición Inferiores",
    page_icon="🔴",
    layout="wide",
)

ASSETS_DIR = Path(__file__).parent / "assets"
LOGO_PATH = ASSETS_DIR / "escudo.png"

st.markdown(
    f"""
    <style>
        .stApp {{ background-color: #fafafa; }}
        h1, h2, h3 {{ color: {CLUB_BLACK}; }}
        div.stButton > button, div.stDownloadButton > button {{
            background-color: {CLUB_RED};
            color: white;
            border: none;
            font-weight: 600;
        }}
        div.stButton > button:hover, div.stDownloadButton > button:hover {{
            background-color: {CLUB_BLACK};
            color: white;
        }}
        [data-testid="stMetricValue"] {{ color: {CLUB_RED}; }}
    </style>
    """,
    unsafe_allow_html=True,
)

col_logo, col_title = st.columns([1, 6])
with col_logo:
    if LOGO_PATH.exists():
        st.image(str(LOGO_PATH), width=90)
    else:
        st.markdown(
            f"<div style='font-size:52px; line-height:1;'>🔴⚫</div>",
            unsafe_allow_html=True,
        )
with col_title:
    st.title("Limpieza Nutrición Inferiores")
    st.caption("Club Atlético Unión (Santa Fe) · Divisiones Inferiores")

if not LOGO_PATH.exists():
    st.info(
        "No hay un escudo cargado todavía. Colocá el archivo del escudo del "
        "club como **assets/escudo.png** (junto a este app.py) y va a "
        "aparecer automáticamente acá arriba.",
        icon="🛡️",
    )

st.divider()

# --------------------------------------------------------------------------
# Configuración de columnas de la base de datos (NUTRI_LONG)
# --------------------------------------------------------------------------

DB_COLUMNS = [
    "JUGADOR",
    "CAT",
    "FECHA NAC",
    "EDAD",
    "POS",
    "PESO",
    "TALLA",
    "S6PLIEGUES",
    "IMO",
]

# Nombre de columna en la planilla del club -> campo de la base de datos
SOURCE_COLUMN_MAP = {
    "JUGADOR": "APELLIDO",
    "FECHA NAC": "FECHNAC",
    "EDAD": "EDAD",
    "PESO": "PESO",
    "TALLA": "TALLA",
    "S6PLIEGUES": "S6PLIEG",
    "IMO": "I M/O",
}

STAT_LABELS = {"promedio", "de", "max", "min"}

HEADER_ROW_INDEX = 1  # header en la fila 2 de excel (0-indexed = 1)

# --------------------------------------------------------------------------
# Normalización de nombres: "Apellido N" (sin comas, tildes ni caracteres raros)
# --------------------------------------------------------------------------

COMPOUND_GIVEN_STARTERS = {
    "juan", "jose", "maria", "gian", "ana", "luis", "carlos", "diego",
    "pedro", "victor",
}
CONNECTOR_WORDS = {
    "de", "del", "san", "santa", "los", "las", "von", "van", "mac", "mc",
}


def _strip_accents(text: str) -> str:
    nfkd = unicodedata.normalize("NFKD", text)
    return "".join(c for c in nfkd if not unicodedata.combining(c))


def _clean_token_text(text: str) -> str:
    text = _strip_accents(text)
    text = re.sub(r"[^A-Za-z\s]", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def format_player_name(raw) -> str:
    """Convierte 'Apellido, Nombre' / 'Apellido Nombre' (con posibles dobles
    apellidos o nombres compuestos) al formato 'Apellido N', sin comas,
    tildes ni caracteres especiales."""
    if raw is None or (isinstance(raw, float) and pd.isna(raw)):
        return ""
    s = str(raw).strip()
    if not s:
        return ""

    s = re.sub(r"\([^)]*\)", "", s).strip()  # saca notas tipo "(rva)"

    if "," in s:
        left, right = s.split(",", 1)
        surname_part = left.strip()
        given_part = right.strip()
    else:
        tokens = [t for t in s.split() if t]
        if len(tokens) == 0:
            return ""
        if len(tokens) == 1:
            surname_part, given_part = tokens[0], ""
        elif len(tokens) == 2 and len(_clean_token_text(tokens[1])) <= 2:
            # ya viene abreviado, ej "Acevedo T"
            surname_part, given_part = tokens[0], tokens[1]
        elif tokens[0].lower() in CONNECTOR_WORDS and len(tokens) >= 3:
            surname_part = " ".join(tokens[0:2])
            given_part = " ".join(tokens[2:])
        elif len(tokens) >= 3 and tokens[-2].lower() in COMPOUND_GIVEN_STARTERS:
            surname_part = " ".join(tokens[:-2])
            given_part = " ".join(tokens[-2:])
        else:
            surname_part = " ".join(tokens[:-1])
            given_part = tokens[-1]

    surname_clean = _clean_token_text(surname_part)
    surname_clean = " ".join(w.capitalize() for w in surname_clean.split())

    given_clean = _clean_token_text(given_part)
    given_words = given_clean.split()
    initial = given_words[0][0].upper() if given_words else ""

    return f"{surname_clean} {initial}".strip() if initial else surname_clean


def sheet_category_label(sheet_name: str) -> str:
    """'4ta (2006)' -> '4ta'. Se usa como CAT mientras esa columna no exista
    en la planilla del club."""
    return re.sub(r"\s*\(.*?\)\s*", " ", sheet_name).strip()


# --------------------------------------------------------------------------
# Carga de la planilla del club
# --------------------------------------------------------------------------


@st.cache_data(show_spinner=False)
def list_sheet_names(file_bytes: bytes):
    xls = pd.ExcelFile(io.BytesIO(file_bytes), engine="openpyxl")
    return xls.sheet_names


@st.cache_data(show_spinner=False)
def read_sheet(file_bytes: bytes, sheet_name: str) -> pd.DataFrame:
    df = pd.read_excel(
        io.BytesIO(file_bytes), sheet_name=sheet_name, header=HEADER_ROW_INDEX,
        engine="openpyxl",
    )
    df.columns = [str(c).strip() for c in df.columns]
    return df


def get_column(df: pd.DataFrame, name: str) -> pd.Series:
    if name in df.columns:
        return df[name]
    return pd.Series([None] * len(df), index=df.index)


def build_long_rows(file_bytes: bytes, sheet_name: str) -> pd.DataFrame:
    """Extrae y da formato a las filas de jugadores válidas de una hoja,
    conservando FECHEVAL para poder filtrar por fecha después."""
    df = read_sheet(file_bytes, sheet_name)

    apellido_col = get_column(df, "APELLIDO")
    valid = apellido_col.notna() & (
        ~apellido_col.astype(str).str.strip().str.lower().isin(STAT_LABELS)
    )
    df = df[valid].copy()
    if df.empty:
        return pd.DataFrame(columns=["_HOJA", "_FECHEVAL"] + DB_COLUMNS)

    out = pd.DataFrame(index=df.index)
    out["_HOJA"] = sheet_name
    out["_FECHEVAL"] = pd.to_datetime(get_column(df, "FECHEVAL"), errors="coerce")

    out["JUGADOR"] = get_column(df, "APELLIDO").apply(format_player_name)

    if "CAT" in df.columns:
        cat_fallback = sheet_category_label(sheet_name)
        cat_series = df["CAT"].astype(object)
        out["CAT"] = cat_series.where(cat_series.notna() & (cat_series.astype(str).str.strip() != ""), cat_fallback)
    else:
        out["CAT"] = sheet_category_label(sheet_name)

    fechnac = pd.to_datetime(get_column(df, "FECHNAC"), errors="coerce")
    out["FECHA NAC"] = fechnac.dt.strftime("%d/%m/%Y")
    out.loc[fechnac.isna(), "FECHA NAC"] = ""

    edad = pd.to_numeric(get_column(df, "EDAD"), errors="coerce")
    out["EDAD"] = edad.apply(lambda v: int(v) if pd.notna(v) else "")

    if "POS" in df.columns:
        pos_series = df["POS"].astype(object)
        out["POS"] = pos_series.where(pos_series.notna(), "")
    else:
        out["POS"] = ""

    peso = pd.to_numeric(get_column(df, "PESO"), errors="coerce").round(1)
    out["PESO"] = peso.apply(lambda v: v if pd.notna(v) else "")

    talla = pd.to_numeric(get_column(df, "TALLA"), errors="coerce").round(1)
    out["TALLA"] = talla.apply(lambda v: v if pd.notna(v) else "")

    s6pliegues = pd.to_numeric(get_column(df, "S6PLIEG"), errors="coerce").round(1)
    out["S6PLIEGUES"] = s6pliegues.apply(lambda v: v if pd.notna(v) else "")

    imo = pd.to_numeric(get_column(df, "I M/O"), errors="coerce").round(2)
    out["IMO"] = imo.apply(lambda v: v if pd.notna(v) else "")

    return out.reset_index(drop=True)


# --------------------------------------------------------------------------
# Carga de archivo
# --------------------------------------------------------------------------

st.subheader("1. Cargar planilla del club")
uploaded_file = st.file_uploader(
    "Excel de nutrición (con una hoja por categoría)", type=["xlsx", "xlsm"]
)

with st.expander("Base de datos NUTRI_LONG existente (opcional)"):
    st.caption(
        "Si subís la base de datos actual, la app va a fusionar los "
        "jugadores nuevos que confirmes con lo que ya tenés cargado."
    )
    existing_db_file = st.file_uploader(
        "NUTRI_LONG.csv o .xlsx actual", type=["csv", "xlsx"], key="existing_db"
    )

if not uploaded_file:
    st.stop()

file_bytes = uploaded_file.getvalue()

try:
    sheet_names = list_sheet_names(file_bytes)
except Exception as exc:
    st.error(f"No pude leer el Excel: {exc}")
    st.stop()

# Solo ofrecemos hojas que realmente tengan jugadores válidos
sheets_with_data = []
for sh in sheet_names:
    tmp = build_long_rows(file_bytes, sh)
    if not tmp.empty:
        sheets_with_data.append(sh)

if not sheets_with_data:
    st.warning("No encontré jugadores válidos en ninguna hoja de este Excel.")
    st.stop()

# --------------------------------------------------------------------------
# 2. Filtros: hoja (categoría) y fecha
# --------------------------------------------------------------------------

st.subheader("2. Elegí categoría(s) y fecha(s) de evaluación")

if "sheets_selected" not in st.session_state:
    st.session_state.sheets_selected = sheets_with_data.copy()

bcol1, bcol2, _ = st.columns([1, 1, 4])
with bcol1:
    if st.button("Seleccionar todas", key="btn_all_sheets"):
        st.session_state.sheets_selected = sheets_with_data.copy()
with bcol2:
    if st.button("Ninguna", key="btn_none_sheets"):
        st.session_state.sheets_selected = []

selected_sheets = st.multiselect(
    "Categoría (hoja)",
    options=sheets_with_data,
    key="sheets_selected",
)

if not selected_sheets:
    st.info("Elegí al menos una categoría para continuar.")
    st.stop()

combined = pd.concat(
    [build_long_rows(file_bytes, sh) for sh in selected_sheets], ignore_index=True
)

available_dates = sorted(combined["_FECHEVAL"].dropna().unique())
date_options = [pd.Timestamp(d).strftime("%d/%m/%Y") for d in available_dates]
date_lookup = {pd.Timestamp(d).strftime("%d/%m/%Y"): d for d in available_dates}

if not date_options:
    st.warning(
        "Ninguna fila de las hojas elegidas tiene fecha de evaluación (FECHEVAL) cargada."
    )
    st.stop()

if "dates_selected" not in st.session_state or not set(
    st.session_state.get("dates_selected", [])
).issubset(set(date_options)):
    st.session_state.dates_selected = date_options.copy()

dcol1, dcol2, _ = st.columns([1, 1, 4])
with dcol1:
    if st.button("Seleccionar todas", key="btn_all_dates"):
        st.session_state.dates_selected = date_options.copy()
with dcol2:
    if st.button("Ninguna", key="btn_none_dates"):
        st.session_state.dates_selected = []

selected_dates_labels = st.multiselect(
    "Fecha de evaluación (FECHEVAL)",
    options=date_options,
    key="dates_selected",
)

if not selected_dates_labels:
    st.info("Elegí al menos una fecha para continuar.")
    st.stop()

selected_dates = {date_lookup[d] for d in selected_dates_labels}
filtered = combined[combined["_FECHEVAL"].isin(selected_dates)].copy()
filtered = filtered.drop(columns=["_FECHEVAL"]).reset_index(drop=True)
filtered.insert(0, "Incluir", True)

# --------------------------------------------------------------------------
# 3. Vista previa editable
# --------------------------------------------------------------------------

st.subheader("3. Confirmá los jugadores a exportar")
st.caption(
    "Desmarcá los jugadores que no querés exportar. El nombre ya viene "
    "ajustado al formato 'Apellido N'; si algún caso quedó mal (por ejemplo "
    "doble apellido o error de carga en el Excel original) podés editarlo "
    "directamente en la tabla."
)

pcol1, pcol2, _ = st.columns([1, 1, 4])
with pcol1:
    if st.button("Marcar todos"):
        filtered["Incluir"] = True
with pcol2:
    if st.button("Desmarcar todos"):
        filtered["Incluir"] = False

edited = st.data_editor(
    filtered,
    key="preview_editor",
    use_container_width=True,
    hide_index=True,
    disabled=["_HOJA", "CAT", "FECHA NAC", "EDAD", "POS", "PESO", "TALLA", "S6PLIEGUES", "IMO"],
    column_config={
        "Incluir": st.column_config.CheckboxColumn("Incluir", width="small"),
        "_HOJA": st.column_config.TextColumn("Hoja"),
        "JUGADOR": st.column_config.TextColumn("JUGADOR"),
    },
)

n_total = len(edited)
n_incluidos = int(edited["Incluir"].sum())
m1, m2 = st.columns(2)
m1.metric("Jugadores en la vista previa", n_total)
m2.metric("Seleccionados para exportar", n_incluidos)

export_df = edited[edited["Incluir"]].drop(columns=["Incluir", "_HOJA"]).reset_index(drop=True)
export_df = export_df[DB_COLUMNS]

# --------------------------------------------------------------------------
# 4. Exportar
# --------------------------------------------------------------------------

st.subheader("4. Exportar")

if export_df.empty:
    st.info("No hay jugadores seleccionados todavía.")
    st.stop()

today_str = date.today().strftime("%Y%m%d")

new_batch_csv = export_df.to_csv(index=False).encode("utf-8-sig")
st.download_button(
    "⬇️ Descargar jugadores seleccionados (CSV, formato NUTRI_LONG)",
    data=new_batch_csv,
    file_name=f"nutri_long_nuevo_{today_str}.csv",
    mime="text/csv",
)

if existing_db_file is not None:
    try:
        if existing_db_file.name.lower().endswith(".csv"):
            existing_db = pd.read_csv(existing_db_file)
        else:
            existing_db = pd.read_excel(existing_db_file)
        existing_db.columns = [str(c).strip() for c in existing_db.columns]
        missing_cols = [c for c in DB_COLUMNS if c not in existing_db.columns]
        for c in missing_cols:
            existing_db[c] = ""
        existing_db = existing_db[DB_COLUMNS]

        merged = pd.concat([existing_db, export_df], ignore_index=True)
        merged = merged.drop_duplicates()

        st.success(
            f"Base fusionada: {len(existing_db)} filas existentes + "
            f"{len(export_df)} nuevas → {len(merged)} filas finales (sin duplicados exactos)."
        )

        merged_csv = merged.to_csv(index=False).encode("utf-8-sig")
        st.download_button(
            "⬇️ Descargar base NUTRI_LONG actualizada (CSV)",
            data=merged_csv,
            file_name=f"NUTRI_LONG_actualizado_{today_str}.csv",
            mime="text/csv",
        )

        xlsx_buffer = io.BytesIO()
        with pd.ExcelWriter(xlsx_buffer, engine="openpyxl") as writer:
            merged.to_excel(writer, index=False, sheet_name="NUTRI_LONG")
        st.download_button(
            "⬇️ Descargar base NUTRI_LONG actualizada (Excel)",
            data=xlsx_buffer.getvalue(),
            file_name=f"NUTRI_LONG_actualizado_{today_str}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    except Exception as exc:
        st.error(f"No pude leer la base de datos existente: {exc}")

st.divider()
st.dataframe(export_df, use_container_width=True, hide_index=True)
