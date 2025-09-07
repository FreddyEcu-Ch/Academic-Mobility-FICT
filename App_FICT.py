# Streamlit app: Movilidad Académica FICT (2023–2025)
# Reads the provided Excel and builds 6 sections with metrics and Altair bar charts.
import streamlit as st
import pandas as pd
import numpy as np
import altair as alt
from pathlib import Path

st.set_page_config(page_title="Movilidad Académica FICT 2023–2025",
                   page_icon="🎓", layout="wide", initial_sidebar_state="expanded")

st.title("🎓 Movilidad Académica FICT — 2023·2024·2025")
st.caption("Fuente: Coordinación de Movilidad Académica FICT. "
           "Los datos 2025 corresponden al corte indicado en la hoja de países (04 de septiembre).")


# ------------------------- Helpers to parse the Excel -------------------------
def load_excel(path: Path):
    return pd.ExcelFile(path)


def parse_comparativa(xls: pd.ExcelFile):
    """Parse 'Comparativa 2023 - 2025' into a nested dict {year: {block: {item: value}}}"""
    df = pd.read_excel(xls, sheet_name="Comparativa 2023 - 2025")
    # Column blocks for the three years (start label col, name col, value col)
    blocks = [
        (3, 4, 5, "2023"),
        (8, 9, 10, "2024"),
        (13, 14, 15, "2025"),
    ]
    out = {lab: {} for *_ , lab in blocks}
    for (c_label, c_name, c_val, lab) in blocks:
        current_block = None
        for _, row in df.iterrows():
            label = str(row.iloc[c_label]) if not pd.isna(row.iloc[c_label]) else None
            name  = row.iloc[c_name]
            val   = row.iloc[c_val]
            if isinstance(label, str) and label.strip() and label.lower() not in ("nan",):
                current_block = label.strip()
                if current_block not in out[lab]:
                    out[lab][current_block] = {}
            if current_block and isinstance(name, str) and not pd.isna(val):
                try:
                    out[lab][current_block][name.strip()] = float(val)
                except:
                    pass
    # compute totals (sum of each block)
    totals = {}
    for y in out:
        tot = 0
        if "Carreras y Programas" in out[y]:
            tot = sum(out[y]["Carreras y Programas"].values())
        elif "Tipo de movilidad" in out[y]:
            tot = sum(out[y]["Tipo de movilidad"].values())
        totals[y] = tot
    return out, totals


def tidy_from_block(dct, year, block):
    """Return a tidy dataframe (Categoria, Valor, Año)."""
    data = dct[year].get(block, {})
    return pd.DataFrame({"Categoría": list(data.keys()), "Valor": list(data.values())}).assign(Año=year)


def parse_countries(xls: pd.ExcelFile):
    result = {}
    for sheet, year in [("Países 2023","2023"), ("Países 2024","2024"), ("Países 2025","2025")]:
        try:
            df = pd.read_excel(xls, sheet_name=sheet)
        except Exception:
            continue
        # Find columns by header names, which appear split in two blocks (Entrante/Saliente)
        # For 2025 headers are: "Movilidad Entrante" and "Movilidad Saliente"
        # For 2023/24 they are "Entrante" and "Saliente"
        # We'll locate column indexes by searching for the "Country"/"Modality"/"Count" labels in first row.
        # Build tidy
        tidy = []
        # Find left block
        cols = df.columns
        # Find any column whose first row equals "Country"
        left_country_col = [i for i,c in enumerate(cols) if str(df.iloc[0,i]).strip().lower()=="country"]
        # In some sheets first row has NaN and second row has data: we'll still use neighbor info.
        # We'll just assume blocks at positions [1..4] and [5..8] if available.
        # Fallback:
        if not left_country_col and len(cols)>=4:
            left_block = (1,2,3)
        else:
            i = left_country_col[0]
            left_block = (i, i+1, i+2)
        # Right block: search for second "Country"
        right_candidates = [i for i,c in enumerate(cols) if str(df.iloc[0,i]).strip().lower()=="country" and i>left_block[0]]
        if right_candidates:
            j = right_candidates[0]
            right_block = (j, j+1, j+2)
        elif len(cols)>=8:
            right_block = (6,7,8)
        else:
            right_block = None
        # Read left
        for _, r in df.iterrows():
            if pd.isna(r.iloc[left_block[0]]) or pd.isna(r.iloc[left_block[2]]):
                continue
            tidy.append({"Año": year, "Tipo": "Entrante", "País": str(r.iloc[left_block[0]]),
                         "Modalidad": str(r.iloc[left_block[1]]), "Casos": float(r.iloc[left_block[2]])})
        # Read right
        if right_block:
            for _, r in df.iterrows():
                if pd.isna(r.iloc[right_block[0]]) or pd.isna(r.iloc[right_block[2]]):
                    continue
                tidy.append({"Año": year, "Tipo": "Saliente", "País": str(r.iloc[right_block[0]]),
                             "Modalidad": str(r.iloc[right_block[1]]), "Casos": float(r.iloc[right_block[2]])})
        result[year] = pd.DataFrame(tidy)
    return result


def bar(df, x, y, title, color=None, sort='-y'):
    enc = {
        "x": alt.X(x, sort=sort, title=""),
        "y": alt.Y(y, title="Total"),
        "tooltip": [x, y]
    }
    if color:
        enc["color"] = alt.Color(color, legend=alt.Legend(title=""))
    chart = alt.Chart(df).mark_bar().encode(**enc).properties(height=330, title=title)
    return chart


# ------------------------- Load Excel from working dir -------------------------
default_path = Path("Movilidad Académica - FICT 2025.xlsx")
uploaded = st.sidebar.file_uploader("Cargar Excel de Movilidad (xlsx)", type=["xlsx"])

if uploaded:
    xls = load_excel(uploaded)
else:
    if default_path.exists():
        xls = load_excel(default_path)
        st.sidebar.info(f"Usando archivo local: {default_path.name}")
    else:
        st.error("Suba el archivo 'Movilidad Académica - FICT 2025.xlsx' para continuar.")
        st.stop()

comp_dict, year_totals = parse_comparativa(xls)
countries_dict = parse_countries(xls)

# Selector de año (para secciones por año)
year = st.sidebar.selectbox("Año", ["2023","2024","2025"], index=2)

# ------------------------- Tabs (6 secciones) -------------------------
tabs = st.tabs([
    "Comparativa 2023–2025",
    "Tipo de movilidad",
    "Categoría: Movilidades por carrera",
    "Modalidad",
    "Tipo de Actividad",
    "Países",
])

# 1) Comparativa 2023–2025
with tabs[0]:
    st.subheader("Comparativa global 2023–2025")
    # KPIs
    c1,c2,c3 = st.columns(3)
    c1.metric("Total 2023", int(year_totals.get("2023",0)))
    c2.metric("Total 2024", int(year_totals.get("2024",0)))
    c3.metric("Total 2025", int(year_totals.get("2025",0)))
    # Comparativas por bloque
    for block in ["Tipo de movilidad","Nivel","Categoría","Modalidad"]:
        df_blk = pd.concat([tidy_from_block(comp_dict, y, block) for y in ["2023","2024","2025"] if block in comp_dict[y]], ignore_index=True)
        st.altair_chart(bar(df_blk, "Categoría", "Valor", f"{block} — Comparativa 2023–2025", color="Año"), use_container_width=True)

# 2) Tipo de movilidad (por año)
with tabs[1]:
    st.subheader(f"Tipo de movilidad — {year}")
    df_m = tidy_from_block(comp_dict, year, "Tipo de movilidad")
    col1, col2 = st.columns(2)
    col1.metric("Entrante", int(df_m.loc[df_m["Categoría"]=="Movilidad Entrante","Valor"].sum()))
    col2.metric("Saliente", int(df_m.loc[df_m["Categoría"]=="Movilidad Saliente","Valor"].sum()))
    st.altair_chart(bar(df_m, "Categoría", "Valor", f"Tipo de movilidad ({year})"), use_container_width=True)

# 3) Movilidades por carrera (por año)
with tabs[2]:
    st.subheader(f"Movilidades por carrera — {year}")
    df_carr = tidy_from_block(comp_dict, year, "Carreras y Programas").sort_values("Valor", ascending=False)
    topn = st.slider("Mostrar top N carreras", 5, len(df_carr), min(10, len(df_carr)))
    col1, col2 = st.columns(2)
    col1.metric("Carreras con >0", int((df_carr["Valor"]>0).sum()))
    col2.metric("Total", int(df_carr["Valor"].sum()))
    st.altair_chart(bar(df_carr.head(topn), "Categoría", "Valor", f"Carreras y Programas ({year})"), use_container_width=True)

# 4) Modalidad (por año)
with tabs[3]:
    st.subheader(f"Modalidad — {year}")
    df_mod = tidy_from_block(comp_dict, year, "Modalidad")
    col1, col2 = st.columns(2)
    col1.metric("Virtual", int(df_mod.loc[df_mod["Categoría"].str.lower()=="virtual","Valor"].sum()))
    col2.metric("Presencial", int(df_mod.loc[df_mod["Categoría"].str.lower()=="presencial","Valor"].sum()))
    st.altair_chart(bar(df_mod, "Categoría", "Valor", f"Modalidad ({year})"), use_container_width=True)

# 5) Tipo de Actividad (por año)
with tabs[4]:
    st.subheader(f"Tipo de Actividad — {year}")
    df_act = tidy_from_block(comp_dict, year, "Tipo de Actividad").sort_values("Valor", ascending=False)
    col1, col2, col3 = st.columns(3)
    col1.metric("Intercambio Académico", int(df_act.loc[df_act["Categoría"].str.startswith("Intercambio"),"Valor"].sum()))
    col2.metric("Cursos de Formación", int(df_act.loc[df_act["Categoría"].str.startswith("Cursos"),"Valor"].sum()))
    col3.metric("Eventos/Estancia/Presentación", int(df_act.loc[~df_act["Categoría"].str.startswith(("Intercambio","Cursos")),"Valor"].sum()))
    st.altair_chart(bar(df_act, "Categoría", "Valor", f"Tipo de Actividad ({year})"), use_container_width=True)

# 6) Países (por año y por tipo)
with tabs[5]:
    st.subheader(f"Países — {year}")
    df_pais = countries_dict.get(year, pd.DataFrame(columns=["País","Tipo","Modalidad","Casos"]))
    if df_pais.empty:
        st.info("No se encontraron datos de países en el Excel para este año.")
    else:
        tipo = st.radio("Tipo", ["Entrante","Saliente"], horizontal=True)
        df_t = df_pais[df_pais["Tipo"]==tipo]
        col1, col2 = st.columns(2)
        col1.metric(f"Países ({tipo})", df_t["País"].nunique())
        col2.metric("Total casos", int(df_t["Casos"].sum()))
        st.altair_chart(bar(df_t.groupby("País", as_index=False)["Casos"].sum().sort_values("Casos", ascending=False),
                            "País", "Casos", f"Países — {tipo} ({year})"), use_container_width=True)
        st.altair_chart(bar(df_t.groupby("Modalidad", as_index=False)["Casos"].sum(),
                            "Modalidad", "Casos", f"Modalidad — {tipo} ({year})"), use_container_width=True)

st.divider()
st.caption("© FICT — ESPOL | Dashboard construido con Streamlit y Altair")


