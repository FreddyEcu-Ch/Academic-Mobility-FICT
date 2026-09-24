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


def cargar_detalle_carreras_2026():
    """Devuelve los registros públicos de 2026 para graficar cualquier carrera.

    Cada fila es una participación, no necesariamente una persona diferente.
    Utiliza el MISMO Excel que agregar_2026_desde_excel().
    """
    ruta = (
        Path(__file__).resolve().parent
        / "Data"
        / "Movilidad_FICT_2026_Carreras_Completas.xlsx"
    )
    requeridas = [
        "Carrera/Programa en ESPOL",
        "Rol en ESPOL",
        "Modalidad",
        "Actividad desarrollada",
        "País",
    ]
    partes = []

    for hoja, tipo in (("Entrantes", "Entrante"), ("Salientes", "Saliente")):
        df = pd.read_excel(ruta, sheet_name=hoja)
        faltantes = [col for col in requeridas if col not in df.columns]
        if faltantes:
            raise ValueError(f"Hoja '{hoja}': faltan columnas {faltantes}")

        df = df[requeridas].dropna(how="all").copy()
        for col in requeridas:
            df[col] = df[col].fillna("").astype(str).str.strip()

        if df["Carrera/Programa en ESPOL"].eq("").any():
            raise ValueError(f"Hoja '{hoja}': existen registros sin carrera.")

        df["Carrera"] = df["Carrera/Programa en ESPOL"].replace(
            {"Ingeniería civil": "Ingeniería Civil"}
        )
        df["Tipo"] = tipo

        # Conservar todas las participaciones: administrativas aparecen en Otros roles.
        rol = df["Rol en ESPOL"].str.casefold()
        df["Grupo"] = "Otros roles"
        df.loc[rol.str.contains("estudiante", regex=False), "Grupo"] = "Estudiantes"
        df.loc[
            rol.str.contains("académico|academico|docente|investigador|profesor", regex=True),
            "Grupo",
        ] = "Profesores"

        df["Modalidad"] = df["Modalidad"].str.capitalize().replace(
            {"": "Sin información"}
        )
        df["Actividad"] = df["Actividad desarrollada"].replace(
            {"": "Sin información"}
        )
        df["País"] = df["País"].str.title().replace(
            {"Peru": "Perú", "Mexico": "México", "Panama": "Panamá", "": "Sin información"}
        )

        partes.append(df[["Carrera", "Tipo", "Grupo", "Modalidad", "Actividad", "País"]])

    return pd.concat(partes, ignore_index=True)
