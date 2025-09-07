# Streamlit app: Movilidad Académica FICT (2023–2025) con carga por defecto del Excel
import streamlit as st
import pandas as pd
import altair as alt
from PIL import Image
from pathlib import Path

st.set_page_config(
    page_title="Movilidad Académica FICT",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Add the logo of FICT
logo_fict = Image.open("Resources/LogoFICTverde.png")
st.image(logo_fict)

st.title("🎓 Movilidad Académica FICT — 2025")
st.caption("Fuente: Coordinación de Movilidad Académica FICT.")


# ------------------------- Helpers -------------------------
def load_excel(src):
    return pd.ExcelFile(src)


def parse_comparativa(xls: pd.ExcelFile):
    df = pd.read_excel(xls, sheet_name="Comparativa 2023 - 2025")
    blocks = [(3, 4, 5, "2023"), (8, 9, 10, "2024"), (13, 14, 15, "2025")]
    out = {lab: {} for *_, lab in blocks}
    for c_label, c_name, c_val, lab in blocks:
        current_block = None
        for _, row in df.iterrows():
            label = row.iloc[c_label]
            name = row.iloc[c_name]
            val = row.iloc[c_val]
            if isinstance(label, str) and label.strip():
                current_block = label.strip()
                out[lab].setdefault(current_block, {})
            if current_block and isinstance(name, str) and pd.notna(val):
                try:
                    out[lab][current_block][name.strip()] = float(val)
                except Exception:
                    pass
    totals = {}
    for y in out:
        if "Carreras y Programas" in out[y]:
            totals[y] = sum(out[y]["Carreras y Programas"].values())
        elif "Tipo de movilidad" in out[y]:
            totals[y] = sum(out[y]["Tipo de movilidad"].values())
        else:
            totals[y] = 0
    return out, totals


def tidy_from_block(dct, year, block):
    data = dct[year].get(block, {})
    return pd.DataFrame(
        {"Categoría": list(data.keys()), "Valor": list(data.values())}
    ).assign(Año=year)


def parse_countries(xls):
    """
    Lee las hojas 'Países 2023/2024/2025' con dos bloques:
    [Country, Modality, Count] (Entrante) y [Country, Modality, Count] (Saliente).
    Tolera filas vacías y encabezados en la primera/siguientes filas.
    """
    result = {}
    for sheet, year in [
        ("Países 2023", "2023"),
        ("Países 2024", "2024"),
        ("Países 2025", "2025"),
    ]:

        try:
            raw = pd.read_excel(xls, sheet_name=sheet, header=None)
        except Exception:
            result[year] = pd.DataFrame(
                columns=["Año", "Tipo", "País", "Modalidad", "Casos"]
            )
            continue

        # 1) Buscar la fila de encabezados (donde aparezca 'Country' y 'Modality')
        hdr_idx = None
        for i in range(min(10, len(raw))):  # busca en las primeras filas
            row_vals = raw.iloc[i].astype(str).str.strip().str.lower()
            if "country" in row_vals.values and "modality" in row_vals.values:
                hdr_idx = i
                break
        if hdr_idx is None:
            # No se encontró un header claro
            result[year] = pd.DataFrame(
                columns=["Año", "Tipo", "País", "Modalidad", "Casos"]
            )
            continue

        header = raw.iloc[hdr_idx].astype(str).str.strip().str.lower().tolist()
        body = raw.iloc[hdr_idx + 1 :].reset_index(drop=True)

        # 2) Localizar los bloques por la posición de 'country' en las columnas
        country_idx = [j for j, name in enumerate(header) if name == "country"]
        tidy_parts = []
        for k, j in enumerate(country_idx):
            # asumimos que las 2 columnas siguientes son 'modality' y 'count'
            cols = [j, j + 1, j + 2]
            sub = body.loc[:, cols].copy()
            sub.columns = ["País", "Modalidad", "Casos"]

            # limpiar y convertir
            sub["País"] = sub["País"].astype(str).str.strip()
            sub["Modalidad"] = sub["Modalidad"].astype(str).str.strip()
            sub["Casos"] = pd.to_numeric(
                sub["Casos"].astype(str).str.replace(",", ".", regex=False),
                errors="coerce",
            )

            sub = sub.dropna(subset=["País", "Casos"])  # elimina vacíos
            sub["Casos"] = sub["Casos"].astype(int)  # normalmente son enteros
            sub["Tipo"] = "Entrante" if k == 0 else "Saliente"
            sub["Año"] = year
            tidy_parts.append(sub[["Año", "Tipo", "País", "Modalidad", "Casos"]])

        result[year] = (
            pd.concat(tidy_parts, ignore_index=True)
            if tidy_parts
            else pd.DataFrame(columns=["Año", "Tipo", "País", "Modalidad", "Casos"])
        )

    return result


def bar(df, x, y, title, color=None, sort='-y'):
    enc = {
        "x": alt.X(f"{x}:N", sort=sort, title=""),
        "y": alt.Y(f"{y}:Q", title="Total"),
        "tooltip": [f"{x}:N", f"{y}:Q"]
    }
    if color:
        enc["color"] = alt.Color(f"{color}:N",
                                 legend=alt.Legend(title=""),
                                 scale=alt.Scale(scheme="tableau10"))
    return alt.Chart(df).mark_bar().encode(**enc).properties(height=330, title=title)


# ------------------------- Carga del Excel -------------------------
# Carpeta donde está el script
BASE_DIR = Path(__file__).resolve().parent

# Archivo por defecto dentro de la subcarpeta Data
DEFAULT_FILE = (
    BASE_DIR / "Data" / "Movilidad_FICT.xlsx"
)  # <- aquí sí permite subcarpeta

# App ESPOL Logo in Sidebar
logo_espol = Image.open("Resources/ESPOL_Negro.png")
st.sidebar.image(logo_espol)

# Carga con fallback al cargador de archivos
uploaded = st.sidebar.file_uploader("Cargar Excel (xlsx)", type=["xlsx"])

if uploaded is not None:
    xls = pd.ExcelFile(uploaded)
    st.sidebar.success(f"Archivo cargado: {uploaded.name}")
elif DEFAULT_FILE.exists():
    xls = pd.ExcelFile(DEFAULT_FILE)
    st.sidebar.info(f"Usando archivo por defecto: {DEFAULT_FILE}")
else:
    st.error("No se encontró el Excel por defecto. Suba el archivo para continuar.")
    st.stop()

comp_dict, year_totals = parse_comparativa(xls)
countries_dict = parse_countries(xls)

year = st.sidebar.selectbox("Año", ["2023", "2024", "2025"], index=2)

tabs = st.tabs(
    [
        "Comparativa 2023–2025",
        "Tipo de movilidad",
        "Categoría: Movilidades por carrera",
        "Modalidad",
        "Tipo de Actividad",
        "Países",
    ]
)

with tabs[0]:
    st.subheader("Comparativa global 2023–2025")
    c1, c2, c3 = st.columns(3)
    c1.metric("Total 2023", int(year_totals.get("2023", 0)))
    c2.metric("Total 2024", int(year_totals.get("2024", 0)))
    c3.metric("Total 2025", int(year_totals.get("2025", 0)))
    for block in ["Tipo de movilidad", "Nivel", "Categoría", "Modalidad"]:
        if not any(block in comp_dict[y] for y in ["2023", "2024", "2025"]):
            continue
        df_blk = pd.concat(
            [tidy_from_block(comp_dict, y, block) for y in ["2023", "2024", "2025"] if
             block in comp_dict[y]], ignore_index=True)
        st.altair_chart(
            bar(df_blk, "Categoría", "Valor", f"{block} — Comparativa 2023–2025",
                color="Año"),
            use_container_width=True)

with tabs[1]:
    st.subheader(f"Tipo de movilidad — {year}")
    if "Tipo de movilidad" in comp_dict[year]:
        df_m = tidy_from_block(comp_dict, year, "Tipo de movilidad")
        col1, col2 = st.columns(2)
        col1.metric(
            "Entrante",
            int(df_m.loc[df_m["Categoría"] == "Movilidad Entrante", "Valor"].sum()),
        )
        col2.metric(
            "Saliente",
            int(df_m.loc[df_m["Categoría"] == "Movilidad Saliente", "Valor"].sum()),
        )
        st.altair_chart(
            bar(df_m, "Categoría", "Valor", f"Tipo de movilidad ({year})",
                color="Categoría"),
            use_container_width=True
        )

    else:
        st.info("No hay datos de Tipo de movilidad para este año.")

with tabs[2]:
    st.subheader(f"Movilidades por carrera — {year}")
    if "Carreras y Programas" in comp_dict[year]:
        df_carr = tidy_from_block(comp_dict, year, "Carreras y Programas").sort_values(
            "Valor", ascending=False
        )
        topn = st.slider(
            "Mostrar top N carreras", 5, len(df_carr), min(10, len(df_carr))
        )
        col1, col2 = st.columns(2)
        col1.metric("Carreras con >0", int((df_carr["Valor"] > 0).sum()))
        col2.metric("Total", int(df_carr["Valor"].sum()))
        st.altair_chart(
            bar(df_carr.head(topn), "Categoría", "Valor",
                f"Carreras y Programas ({year})", color="Categoría"),
            use_container_width=True
        )

    else:
        st.info("No hay datos de carreras para este año.")

with tabs[3]:
    st.subheader(f"Modalidad — {year}")
    if "Modalidad" in comp_dict[year]:
        df_mod = tidy_from_block(comp_dict, year, "Modalidad")
        col1, col2 = st.columns(2)
        col1.metric(
            "Virtual",
            int(
                df_mod.loc[df_mod["Categoría"].str.lower() == "virtual", "Valor"].sum()
            ),
        )
        col2.metric(
            "Presencial",
            int(
                df_mod.loc[
                    df_mod["Categoría"].str.lower() == "presencial", "Valor"
                ].sum()
            ),
        )
        st.altair_chart(
            bar(df_mod, "Categoría", "Valor", f"Modalidad ({year})", color="Categoría"),
            use_container_width=True
        )

    else:
        st.info("No hay datos de modalidad para este año.")

with tabs[4]:
    st.subheader(f"Tipo de Actividad — {year}")
    if "Tipo de Actividad" in comp_dict[year]:
        df_act = tidy_from_block(comp_dict, year, "Tipo de Actividad").sort_values(
            "Valor", ascending=False
        )
        col1, col2, col3 = st.columns(3)
        col1.metric(
            "Intercambio Académico",
            int(
                df_act.loc[
                    df_act["Categoría"].str.startswith("Intercambio"), "Valor"
                ].sum()
            ),
        )
        col2.metric(
            "Cursos de Formación",
            int(
                df_act.loc[df_act["Categoría"].str.startswith("Cursos"), "Valor"].sum()
            ),
        )
        col3.metric(
            "Otros (Eventos/Estancia/Presentación)",
            int(
                df_act.loc[
                    ~df_act["Categoría"].str.startswith(("Intercambio", "Cursos")),
                    "Valor",
                ].sum()
            ),
        )
        st.altair_chart(
            bar(df_act, "Categoría", "Valor", f"Tipo de Actividad ({year})",
                color="Categoría"),
            use_container_width=True
        )

    else:
        st.info("No hay datos de tipo de actividad para este año.")

with tabs[5]:
    st.subheader(f"Países — {year}")
    df_pais = countries_dict.get(
        year, pd.DataFrame(columns=["País", "Tipo", "Modalidad", "Casos"])
    )
    if df_pais.empty:
        st.info("No se encontraron datos de países en el Excel para este año.")
    else:
        tipo = st.radio("Tipo", ["Entrante", "Saliente"], horizontal=True)
        df_t = df_pais[df_pais["Tipo"] == tipo]
        col1, col2 = st.columns(2)
        col1.metric(f"Países ({tipo})", df_t["País"].nunique())
        col2.metric("Total casos", int(df_t["Casos"].sum()))
        st.altair_chart(
            bar(
                df_t.groupby("País", as_index=False)["Casos"]
                .sum()
                .sort_values("Casos", ascending=False),
                "País",
                "Casos",
                f"Países — {tipo} ({year})",
            ),
            use_container_width=True,
        )
        st.altair_chart(
            bar(
                df_t.groupby("Modalidad", as_index=False)["Casos"].sum(),
                "Modalidad",
                "Casos",
                f"Modalidad — {tipo} ({year})",
            ),
            use_container_width=True,
        )

st.divider()
st.caption("© FICT — ESPOL | Septiembre 2025")
