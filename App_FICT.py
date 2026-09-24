# Streamlit app: Movilidad Académica FICT (2023–2025) con carga por defecto del Excel
import streamlit as st
import numpy as np
import pandas as pd
import altair as alt
import plotly.express as px
from PIL import Image
from pathlib import Path
from funciones_2026 import agregar_2026_desde_excel, obtener_resumen_petroleos_2026
from st_aggrid import AgGrid, GridOptionsBuilder

st.set_page_config(
    page_title="Movilidad Académica FICT",
    page_icon="🌍",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
        .main-header {
            font-size:32px !important;
            font-weight:600;
            padding-bottom:10px;
            border-bottom: 1px solid #eee;
        }
    </style>
    """,
    unsafe_allow_html=True,
)


@st.cache_resource
def load_image_cached(path: str):
    try:
        return Image.open(path)
    except Exception:
        return None


# ------------------------- Helpers -------------------------
def load_excel(src):
    return pd.ExcelFile(src)


def parse_comparativa(xls: pd.ExcelFile):
    df = pd.read_excel(xls, sheet_name="Comparativa 2022 - 2025")
    blocks = [
        (1, 2, 3, "2022"),
        (6, 7, 8, "2023"),
        (11, 12, 13, "2024"),
        (16, 17, 18, "2025"),
    ]
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
        ("Países 2022", "2022"),
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


@st.cache_data
def load_funcionarios(path="Data/REGISTROS_RELEX2025.xlsx", sheet="SALIENTE"):
    p = Path(path)
    if not p.exists():
        st.error(f"No se encontró el archivo: {p}. Colócalo en la carpeta Data/")
        return pd.DataFrame()

    # Lee la hoja SALIENTE
    df = pd.read_excel(p, sheet_name=sheet)

    # Limpia detalles de encabezados (espacio no-rompible / saltos de línea)
    df.columns = [
        str(c).replace("\xa0", " ").replace("\n", " ").strip() for c in df.columns
    ]

    # Combina nombres + apellidos (o usa "Nombre" si ya viene armado)
    first = "Nombre(s) de la persona saliente que realiza la movilidad"
    last = "Apellido(s) de la persona saliente que realiza la movilidad"
    if first in df.columns and last in df.columns:
        nombres = (
            df[first].astype(str).str.strip() + " " + df[last].astype(str).str.strip()
        ).str.strip()
    elif "Nombre" in df.columns:
        nombres = df["Nombre"].astype(str).str.strip()
    else:
        nombres = pd.Series("", index=df.index)

    # Construye el dataframe final con los nombres de columnas requeridos
    out = pd.DataFrame(
        {
            "Nombres": nombres,
            "Fecha de inicio": pd.to_datetime(
                df.get("Fecha de inicio de la movilidad saliente"), errors="coerce"
            ).dt.strftime("%Y-%m-%d"),
            "Fecha de fin": pd.to_datetime(
                df.get("Fecha de finalización de la movilidad saliente"),
                errors="coerce",
            ).dt.strftime("%Y-%m-%d"),
            "Duración (Horas)": df.get(
                "Duración en horas dedicadas a la movilidad saliente"
            ),
            "Actividad Realizada": df.get("Actividad realizada"),
            "Institución externa": df.get(
                "Nombre de la Institución externa que aplica la persona saliente"
            ),
            "País": df.get(
                "País donde se encuentra la Institución externa que aplica la persona saliente"
            ),
            "Rol en ESPOL": df.get(
                "En la ESPOL, indique el rol que desempeña la persona saliente"
            ),
            "Modalidad": df.get("Modalidad de la movilidad saliente"),
        }
    )

    return out


def bar(df, x, y, title, color=None, sort="-y"):
    enc = {
        "x": alt.X(f"{x}:N", sort=sort, title=""),
        "y": alt.Y(f"{y}:Q", title="Total"),
        "tooltip": [f"{x}:N", f"{y}:Q"],
    }
    if color:
        enc["color"] = alt.Color(
            f"{color}:N",
            legend=alt.Legend(title=""),
            scale=alt.Scale(scheme="tableau10"),
        )
    return alt.Chart(df).mark_bar().encode(**enc).properties(height=330, title=title)


logo_fict = load_image_cached("Resources/LogoFICTverde.png")
if logo_fict:
    st.image(logo_fict)
else:
    st.write("🌍 Movilidad Académica FICT")

st.markdown(
    "<h1 style='text-align:center;'>🌍 Movilidad Académica FICT — 2026</h1>",
    unsafe_allow_html=True,
)
st.caption("**Fuente:** Coordinación de Movilidad Académica FICT.")
st.markdown(
    "**Coordinador:** [M.Sc. Freddy Carrión Maldonado](https://www.linkedin.com/in/freddy-carri%C3%B3n-maldonado-b3579b125/)"
)

# Sidebar logo
logo_espol = load_image_cached("Resources/ESPOL_Negro.png")
if logo_espol:
    st.sidebar.image(logo_espol)

# ------------------------- Carga del Excel -------------------------
# Carpeta donde está el script
BASE_DIR = Path(__file__).resolve().parent

# Archivo por defecto dentro de la subcarpeta Data
DEFAULT_FILE = (
    BASE_DIR / "Data" / "Movilidad_FICT.xlsx"
)  # <- aquí sí permite subcarpeta

st.markdown(
    """
<style>
    [data-testid="stSidebar"] [data-testid="stFileUploader"] { display: none; }
</style>
""",
    unsafe_allow_html=True,
)

uploaded = st.sidebar.file_uploader("Cargar Excel (xlsx)", type=["xlsx"])

if uploaded is not None:
    xls = pd.ExcelFile(uploaded)
    st.sidebar.success(f"Archivo cargado: {uploaded.name}")
elif DEFAULT_FILE.exists():
    xls = pd.ExcelFile(DEFAULT_FILE)
    # st.sidebar.info(f"Usando archivo por defecto: {DEFAULT_FILE}")
else:
    st.error("No se encontró el Excel por defecto. Suba el archivo para continuar.")
    st.stop()

comp_dict, year_totals = parse_comparativa(xls)
countries_dict = parse_countries(xls)

# Incorporar las estadísticas de movilidad 2026
comp_dict, year_totals, countries_dict = agregar_2026_desde_excel(
    comp_dict,
    year_totals,
    countries_dict
)


year = st.sidebar.selectbox(
    "Año",
    ["2022", "2023", "2024", "2025", "2026"],
    index=4
)

tab_titles = [
    ("📊", "Comparativa 2022–2026"),
    ("🔁", "Tipo de movilidad"),
    ("🎓", "Movilidades por carrera"),
    ("🖥️", "Modalidad"),
    ("🧭", "Actividad"),
    ("🗺️", "Países"),
    ("📋", "Comunidad FICT"),
]
tabs = st.tabs([f"{ico} {title}" for ico, title in tab_titles])


with tabs[0]:

    st.subheader("Comparativa global 2022–2026")

    # Años disponibles para la comparativa
    anios = ["2022", "2023", "2024", "2025", "2026"]

    # Indicadores de movilidad por año
    columnas = st.columns(5)

    for columna, anio in zip(columnas, anios):
        columna.metric(
            f"Total {anio}",
            int(year_totals.get(anio, 0))
        )

    # Gráficos comparativos
    for block in [
        "Tipo de movilidad",
        "Nivel",
        "Categoría",
        "Modalidad"
    ]:

        if not any(
            block in comp_dict[anio]
            for anio in anios
        ):
            continue

        df_blk = pd.concat(
            [
                tidy_from_block(comp_dict, anio, block)
                for anio in anios
                if block in comp_dict[anio]
            ],
            ignore_index=True,
        )

        st.altair_chart(
            bar(
                df_blk,
                "Categoría",
                "Valor",
                f"{block} — Comparativa 2022–2026",
                color="Año",
            ),
            use_container_width=True,
        )

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
            bar(
                df_m,
                "Categoría",
                "Valor",
                f"Tipo de movilidad ({year})",
                color="Categoría",
            ),
            use_container_width=True,
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
            "Mostrar top N carreras",
            5,
            len(df_carr),
            min(10, len(df_carr))
        )

        col1, col2 = st.columns(2)
        col1.metric("Carreras con >0", int((df_carr["Valor"] > 0).sum()))
        col2.metric("Total", int(df_carr["Valor"].sum()))

        st.altair_chart(
            bar(
                df_carr.head(topn),
                "Categoría",
                "Valor",
                f"Carreras y Programas ({year})",
                color="Categoría",
            ),
            use_container_width=True,
        )

        # ----- Gráficos adicionales para Petróleos en 2026 -----
        if year == "2026":
            st.markdown("---")
            st.subheader("Detalle de la carrera de Petróleos — 2026")

            df_pet_tipo, df_pet_rol = obtener_resumen_petroleos_2026()

            g1, g2 = st.columns(2)

            with g1:
                st.altair_chart(
                    bar(
                        df_pet_tipo,
                        "Categoría",
                        "Valor",
                        "Petróleos: movilidades entrantes y salientes",
                        color="Categoría",
                    ),
                    use_container_width=True,
                )

            with g2:
                st.altair_chart(
                    bar(
                        df_pet_rol,
                        "Categoría",
                        "Valor",
                        "Petróleos: distribución entre estudiantes y profesores",
                        color="Categoría",
                    ),
                    use_container_width=True,
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
            use_container_width=True,
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
            bar(
                df_act,
                "Categoría",
                "Valor",
                f"Tipo de Actividad ({year})",
                color="Categoría",
            ),
            use_container_width=True,
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
        # --- Filtros principales
        categoria = st.radio("Categoría", ["Todas", "Entrante", "Saliente"], horizontal=True)
        map_type = st.radio("Tipo de mapa", ["Coroplético", "Burbujas"], horizontal=True)

        # --- Prepara datos base
        df_t = df_pais.copy()

        # Quita filas de totales si existieran
        df_t = df_t[
            ~df_t["País"].astype(str).str.strip().str.lower().isin(["total", "totales", "subtotal"])
        ]

        # Asegura numérico
        df_t["Casos"] = pd.to_numeric(df_t["Casos"], errors="coerce").fillna(0)

        # Normaliza modalidad
        df_t["Modalidad"] = df_t["Modalidad"].astype(str).str.strip()
        df_t["Modalidad_norm"] = df_t["Modalidad"].str.lower()

        # Filtra por categoría (si aplica)
        if categoria != "Todas":
            df_t = df_t[df_t["Tipo"] == categoria].copy()

        # Map de nombres en ES -> EN (ajusta si te faltan países)
        name_map = {
            "España": "Spain",
            "Italia": "Italy",
            "Colombia": "Colombia",
            "Rusia": "Russia",
            "Ecuador": "Ecuador",
            "Perú": "Peru",
            "Chile": "Chile",
            "Argentina": "Argentina",
            "México": "Mexico",
            "Brasil": "Brazil",
            "Estados Unidos": "United States",
            "Reino Unido": "United Kingdom",
            "Países Bajos": "Netherlands",
            "Corea del Sur": "South Korea",
            "Alemania": "Germany",
            "Francia": "France",
            "Suiza": "Switzerland",
            "Austria": "Austria",
            "Suecia": "Sweden",
            "Noruega": "Norway",
            "Finlandia": "Finland",
            "Dinamarca": "Denmark",
            "Polonia": "Poland",
            "Portugal": "Portugal",
            "Irlanda": "Ireland",
            "República Checa": "Czechia",
            "Hungría": "Hungary",
            "Grecia": "Greece",
            "Turquía": "Turkey",
            "Japón": "Japan",
            "China": "China",
            "India": "India",
            "Australia": "Australia",
            "Nueva Zelanda": "New Zealand",
            "Sudáfrica": "South Africa",
            "Marruecos": "Morocco",
        }

        df_t["country_en"] = df_t["País"].replace(name_map)
        df_t["country_en"] = np.where(
            df_t["country_en"].isna() | (df_t["country_en"].astype(str).str.strip() == ""),
            df_t["País"],
            df_t["country_en"],
        )

        # -----------------------------
        # CASO 1: "Todas" (Entrantes+Salientes) e independiente de modalidad
        # -----------------------------
        if categoria == "Todas":
            df_geo = df_t.groupby("country_en", as_index=False)["Casos"].sum()

            c1, c2 = st.columns(2)
            c1.metric("Países (Todas)", int(df_geo["country_en"].nunique()))
            c2.metric("Total casos", int(df_geo["Casos"].sum()))

            st.divider()
            st.markdown(f"### Todas las movilidades (Entrantes + Salientes) — {year}")

            if df_geo.empty or df_geo["Casos"].sum() == 0:
                st.info("Sin datos para 'Todas' las movilidades en este año.")
            else:
                if map_type == "Coroplético":
                    fig = px.choropleth(
                        df_geo,
                        locations="country_en",
                        locationmode="country names",
                        color="Casos",
                        color_continuous_scale="Blues",
                        projection="natural earth",
                        title=f"Todas las movilidades — {year}",
                        hover_name="country_en",
                        labels={"Casos": "Casos"},
                    )
                    fig.update_geos(
                        showcountries=True,
                        showcoastlines=True,
                        showframe=False,
                        countrycolor="black",
                        countrywidth=0.5,
                        coastlinecolor="gray",
                    )
                else:
                    fig = px.scatter_geo(
                        df_geo,
                        locations="country_en",
                        locationmode="country names",
                        size="Casos",
                        color="Casos",
                        color_continuous_scale="Blues",
                        projection="natural earth",
                        title=f"Todas las movilidades — {year}",
                        hover_name="country_en",
                        labels={"Casos": "Casos"},
                    )
                    fig.update_traces(marker_line_color="black", marker_line_width=0.3, opacity=0.85)
                    fig.update_geos(
                        showcountries=True,
                        showcoastlines=True,
                        showframe=False,
                        countrycolor="black",
                        countrywidth=0.5,
                        coastlinecolor="gray",
                    )

                st.plotly_chart(fig, use_container_width=True)

                st.altair_chart(
                    bar(
                        df_geo.sort_values("Casos", ascending=False),
                        "country_en",
                        "Casos",
                        f"Ranking de países — Todas las movilidades ({year})",
                        color="country_en",
                    ),
                    use_container_width=True,
                )

        # -----------------------------
        # CASO 2: "Entrante" o "Saliente" (con selección de modalidad y mapa por modalidad)
        # -----------------------------
        else:
            modalidad_candidates = df_t["Modalidad_norm"].dropna().unique().tolist()
            orden_preferido = ["virtual", "presencial"]
            modalidad_opts = [m for m in orden_preferido if m in modalidad_candidates] + [
                m for m in sorted(modalidad_candidates) if m not in orden_preferido
            ]

            modalidad_sel = st.multiselect(
                "Modalidad (elige una o varias)",
                options=modalidad_opts,
                default=modalidad_opts,
            )

            if not modalidad_sel:
                st.info("Selecciona al menos una modalidad para generar los mapas.")
            else:
                df_base = df_t[df_t["Modalidad_norm"].isin(modalidad_sel)].copy()

                c1, c2 = st.columns(2)
                c1.metric(f"Países ({categoria})", int(df_base["country_en"].nunique()))
                c2.metric("Total casos", int(df_base["Casos"].sum()))

                st.divider()

                for mod in modalidad_sel:
                    df_mod = df_t[df_t["Modalidad_norm"] == mod].copy()

                    if df_mod.empty or df_mod["Casos"].sum() == 0:
                        st.info(f"Sin datos para {categoria} — {mod.capitalize()} ({year}).")
                        continue

                    st.markdown(f"### {categoria} — {mod.capitalize()} ({year})")

                    df_geo = df_mod.groupby("country_en", as_index=False)["Casos"].sum()

                    if map_type == "Coroplético":
                        fig = px.choropleth(
                            df_geo,
                            locations="country_en",
                            locationmode="country names",
                            color="Casos",
                            color_continuous_scale="Blues",
                            projection="natural earth",
                            title=f"{categoria} — {mod.capitalize()} — {year}",
                            hover_name="country_en",
                            labels={"Casos": "Casos"},
                        )
                        fig.update_geos(
                            showcountries=True,
                            showcoastlines=True,
                            showframe=False,
                            countrycolor="black",
                            countrywidth=0.5,
                            coastlinecolor="gray",
                        )
                    else:
                        fig = px.scatter_geo(
                            df_geo,
                            locations="country_en",
                            locationmode="country names",
                            size="Casos",
                            color="Casos",
                            color_continuous_scale="Blues",
                            projection="natural earth",
                            title=f"{categoria} — {mod.capitalize()} — {year}",
                            hover_name="country_en",
                            labels={"Casos": "Casos"},
                        )
                        fig.update_traces(marker_line_color="black", marker_line_width=0.3, opacity=0.85)
                        fig.update_geos(
                            showcountries=True,
                            showcoastlines=True,
                            showframe=False,
                            countrycolor="black",
                            countrywidth=0.5,
                            coastlinecolor="gray",
                        )

                    st.plotly_chart(fig, use_container_width=True)

                    st.altair_chart(
                        bar(
                            df_geo.sort_values("Casos", ascending=False),
                            "country_en",
                            "Casos",
                            f"Ranking de países — {categoria} — {mod.capitalize()} ({year})",
                            color="country_en",
                        ),
                        use_container_width=True,
                    )

                    st.divider()


st.divider()
st.caption("© FICT — ESPOL | Septiembre 2026")
