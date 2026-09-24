"""Integración del Excel 2026 con las estadísticas ya existentes de App_FICT.py."""
from pathlib import Path
import pandas as pd


def agregar_2026_desde_excel(comp_dict, year_totals, countries_dict):
    """Carga 2026 del Excel público y actualiza los diccionarios de la aplicación."""
    ruta = (Path(__file__).resolve().parent / "Data" /
            "Movilidad_FICT_2026_Carreras_Completas.xlsx")
    columnas = ["Rol en ESPOL", "Carrera/Programa en ESPOL", "Modalidad",
                "Actividad desarrollada", "País"]
    partes = []

    for hoja, direccion in (("Entrantes", "Entrante"),
                            ("Salientes", "Saliente")):
        tabla = pd.read_excel(ruta, sheet_name=hoja)
        faltantes = sorted(set(columnas) - set(tabla.columns))
        if faltantes:
            raise ValueError(f"Hoja '{hoja}': faltan las columnas {faltantes}")

        tabla = tabla[columnas].dropna(how="all").copy()
        for nombre in columnas:
            tabla[nombre] = tabla[nombre].fillna("").astype(str).str.strip()

        sin_carrera = tabla["Carrera/Programa en ESPOL"].eq("")
        if sin_carrera.any():
            raise ValueError(
                f"Hoja '{hoja}': {sin_carrera.sum()} registros sin carrera/programa."
            )
        tabla["Tipo"] = direccion
        partes.append(tabla)

    datos = pd.concat(partes, ignore_index=True)
    datos["Carrera/Programa en ESPOL"] = datos[
        "Carrera/Programa en ESPOL"
    ].replace({"Ingeniería civil": "Ingeniería Civil"})
    datos["Modalidad"] = datos["Modalidad"].str.capitalize()

    # Los mapas de App_FICT.py ya traducen algunos países; estos otros
    # se escriben en inglés para que Plotly pueda reconocerlos.
    datos["País"] = datos["País"].str.title().replace({
        "Peru": "Perú", "Belgica": "Belgium", "Panamá": "Panama",
        "Canada": "Canada", "Bolivia": "Bolivia",
    })

    roles = {
        "Estudiante de grado": ("Grado", "Estudiantes"),
        "Estudiante de Posgrado": ("Postgrado", "Estudiantes"),
        "Personal Académico (Docente o Investigador)":
            ("No especificado", "Académicos"),
        "Administrativo": ("No especificado", "Administrativos"),
    }
    actividades = {
        "Clases espejo": "Intercambio Académico",
        "Dictado de clases (profesor invitado)": "Intercambio Académico",
        "Capacitación / Entrenamiento": "Cursos de Formación",
        "Cursos cortos internacionales (escuelas de verano/invierno, semanas internacionales)":
            "Cursos de Formación",
        "Taller": "Cursos de Formación",
        "Charla / Conferencia": "Asistencia a Eventos",
        "Asistencia general a eventos": "Asistencia a Eventos",
        "Presentación de trabajos en eventos": "Presentación de trabajos",
        "Estancia de investigación (doctoral, postdoctoral, investigador visitante)":
            "Estancia",
        "Pasantía de investigación": "Estancia",
        "Reunión": "Reuniones",
        "Visita técnica": "Visita técnica",
        "Representación oficial de la institución": "Representación institucional",
    }

    bloques = {
        "Tipo de movilidad": datos["Tipo"].map({
            "Entrante": "Movilidad Entrante", "Saliente": "Movilidad Saliente"
        }),
        "Nivel": datos["Rol en ESPOL"].map(
            lambda rol: roles.get(rol, ("No especificado", "Sin clasificar"))[0]
        ),
        "Categoría": datos["Rol en ESPOL"].map(
            lambda rol: roles.get(rol, ("No especificado", "Sin clasificar"))[1]
        ),
        "Carreras y Programas": datos["Carrera/Programa en ESPOL"],
        "Modalidad": datos["Modalidad"],
        "Tipo de Actividad": datos["Actividad desarrollada"].replace(actividades),
    }
    # Asignar, no sumar: evita duplicar los datos de 2026 del libro histórico.
    comp_dict["2026"] = {
        nombre: serie.value_counts().to_dict()
        for nombre, serie in bloques.items()
    }
    year_totals["2026"] = len(datos)
    countries_dict["2026"] = (
        datos.assign(Año="2026")
        .groupby(["Año", "Tipo", "País", "Modalidad"], as_index=False)
        .size().rename(columns={"size": "Casos"})
    )
    return comp_dict, year_totals, countries_dict

def obtener_resumen_petroleos_2026():
    """
    Devuelve dos dataframes para graficar:
    1) Movilidades entrantes vs salientes de Petróleos
    2) Distribución entre estudiantes y profesores de Petróleos
    """
    ruta = (
        Path(__file__).resolve().parent
        / "Data"
        / "Movilidad_FICT_2026_Carreras_Completas.xlsx"
    )

    partes = []

    for hoja, tipo_movilidad in (
        ("Entrantes", "Movilidad Entrante"),
        ("Salientes", "Movilidad Saliente"),
    ):
        df = pd.read_excel(ruta, sheet_name=hoja)

        columnas_necesarias = ["Rol en ESPOL", "Carrera/Programa en ESPOL"]
        faltantes = [c for c in columnas_necesarias if c not in df.columns]
        if faltantes:
            raise ValueError(
                f"En la hoja '{hoja}' faltan las columnas: {faltantes}"
            )

        df = df.copy()
        df["Rol en ESPOL"] = df["Rol en ESPOL"].fillna("").astype(str).str.strip()
        df["Carrera/Programa en ESPOL"] = (
            df["Carrera/Programa en ESPOL"]
            .fillna("")
            .astype(str)
            .str.strip()
        )

        df = df[
            df["Carrera/Programa en ESPOL"].str.casefold()
            == "Petróleos".casefold()
        ].copy()

        df["Tipo"] = tipo_movilidad
        partes.append(df)

    if not partes:
        return (
            pd.DataFrame(columns=["Categoría", "Valor"]),
            pd.DataFrame(columns=["Categoría", "Valor"]),
        )

    petro = pd.concat(partes, ignore_index=True)

    # -------- Gráfico 1: Entrantes vs Salientes --------
    df_tipo = (
        petro["Tipo"]
        .value_counts()
        .rename_axis("Categoría")
        .reset_index(name="Valor")
    )

    orden_tipo = ["Movilidad Entrante", "Movilidad Saliente"]
    df_tipo["Categoría"] = pd.Categorical(
        df_tipo["Categoría"],
        categories=orden_tipo,
        ordered=True,
    )
    df_tipo = df_tipo.sort_values("Categoría").reset_index(drop=True)

    # -------- Gráfico 2: Estudiantes vs Profesores --------
    def clasificar_rol(rol):
        rol = str(rol).strip().lower()

        if "estudiante" in rol:
            return "Estudiantes"

        if (
            "académico" in rol
            or "academico" in rol
            or "docente" in rol
            or "investigador" in rol
            or "profesor" in rol
        ):
            return "Profesores"

        return "Otros"

    petro["Grupo"] = petro["Rol en ESPOL"].apply(clasificar_rol)

    df_rol = (
        petro[petro["Grupo"].isin(["Estudiantes", "Profesores"])]["Grupo"]
        .value_counts()
        .rename_axis("Categoría")
        .reset_index(name="Valor")
    )

    orden_rol = ["Estudiantes", "Profesores"]
    df_rol["Categoría"] = pd.Categorical(
        df_rol["Categoría"],
        categories=orden_rol,
        ordered=True,
    )
    df_rol = df_rol.sort_values("Categoría").reset_index(drop=True)

    return df_tipo, df_rol
