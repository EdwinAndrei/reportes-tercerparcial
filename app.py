import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO

# ====================================
# CONFIGURACIÓN DE LA PÁGINA
# ====================================

st.set_page_config(
    page_title="Plataforma de Reportes",
    page_icon="📊",
    layout="wide"
)

st.title(" Plataforma de Reportes y Análisis de Datos")

st.markdown("""
Esta plataforma permite cargar archivos Excel o CSV,
analizar información y generar reportes interactivos.
""")

# ====================================
# CARGA DE ARCHIVOS
# ====================================

archivo = st.file_uploader(
    "Seleccione un archivo Excel o CSV",
    type=["xlsx", "csv"]
)

if archivo is not None:

    # ====================================
    # LECTURA DEL ARCHIVO
    # ====================================

    try:

        if archivo.name.endswith(".csv"):
            datos = pd.read_csv(archivo)

        else:
            datos = pd.read_excel(archivo)

        st.success("Archivo cargado correctamente")

    except Exception as e:

        st.error(f"Error al leer el archivo: {e}")
        st.stop()

    # ====================================
    # TABLA ORIGINAL
    # ====================================

    st.subheader("Vista previa de los datos")

    st.dataframe(
        datos,
        use_container_width=True
    )

    # ====================================
    # FILTROS
    # ====================================

    st.sidebar.header("Filtros")

    datos_filtrados = datos.copy()

    if "Categoria" in datos.columns:

        categorias = sorted(
            datos["Categoria"]
            .dropna()
            .unique()
        )

        categorias_seleccionadas = st.sidebar.multiselect(
            "Categorías",
            categorias,
            default=categorias
        )

        datos_filtrados = datos_filtrados[
            datos_filtrados["Categoria"]
            .isin(categorias_seleccionadas)
        ]

    # ====================================
    # KPIs
    # ====================================

    st.subheader("Indicadores")

    columnas_numericas = (
        datos_filtrados
        .select_dtypes(include="number")
        .columns
        .tolist()
    )

    if len(columnas_numericas) > 0:

        columna_indicador = st.selectbox(
            "Columna numérica para indicadores",
            columnas_numericas
        )

        total = datos_filtrados[columna_indicador].sum()

        promedio = datos_filtrados[columna_indicador].mean()

        registros = len(datos_filtrados)

        col1, col2, col3 = st.columns(3)

        col1.metric(
            "Total",
            f"{total:,.2f}"
        )

        col2.metric(
            "Promedio",
            f"{promedio:,.2f}"
        )

        col3.metric(
            "Registros",
            registros
        )

    # ====================================
    # GENERADOR DE GRÁFICOS
    # ====================================

    st.subheader("Generador de Gráficos")

    columnas = datos_filtrados.columns.tolist()

    eje_x = st.selectbox(
        "Seleccione columna para eje X",
        columnas
    )

    eje_y = st.selectbox(
        "Seleccione columna para eje Y",
        columnas_numericas
    )

    tipo_grafico = st.selectbox(
        "Tipo de gráfico",
        [
            "Barras",
            "Pastel",
            "Líneas",
            "Dispersión"
        ]
    )

    try:

        resumen = (
            datos_filtrados
            .groupby(eje_x)[eje_y]
            .sum()
            .reset_index()
        )

        if tipo_grafico == "Barras":

            fig = px.bar(
                resumen,
                x=eje_x,
                y=eje_y,
                title=f"{eje_y} por {eje_x}"
            )

        elif tipo_grafico == "Pastel":

            fig = px.pie(
                resumen,
                names=eje_x,
                values=eje_y,
                title=f"Distribución de {eje_y}"
            )

        elif tipo_grafico == "Líneas":

            fig = px.line(
                resumen,
                x=eje_x,
                y=eje_y,
                title=f"{eje_y} por {eje_x}"
            )

        else:

            fig = px.scatter(
                datos_filtrados,
                x=eje_x,
                y=eje_y,
                title=f"{eje_y} vs {eje_x}"
            )

        st.plotly_chart(
            fig,
            use_container_width=True
        )

    except Exception as e:

        st.error(
            f"No fue posible generar el gráfico: {e}"
        )

    # ====================================
    # DATOS FILTRADOS
    # ====================================

    st.subheader("Datos Filtrados")

    st.dataframe(
        datos_filtrados,
        use_container_width=True
    )

    # ====================================
    # DESCARGAR REPORTE EXCEL
    # ====================================

    st.subheader("Exportar Reporte")

    buffer = BytesIO()

    with pd.ExcelWriter(
        buffer,
        engine="openpyxl"
    ) as writer:

        datos_filtrados.to_excel(
            writer,
            sheet_name="Reporte",
            index=False
        )

    excel_data = buffer.getvalue()

    st.download_button(
        label=" Descargar Reporte Excel",
        data=excel_data,
        file_name="Reporte.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )