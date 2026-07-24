# Plataforma de Reportes y Análisis de Datos

GRUPO 15

Aplicación desarrollada con Python y Streamlit para procesar datos empresariales y generar reportes visuales.
Aplicación desarrollada con Python y Streamlit para procesar datos empresariales y generar reportes visuales interactivos a partir de archivos Excel o CSV.

![Python](https://img.shields.io/badge/Python-3.10+-blue?logo=python)
![Streamlit](https://img.shields.io/badge/Streamlit-1.57-red?logo=streamlit)
![Pandas](https://img.shields.io/badge/Pandas-2.x-150458?logo=pandas)
![Plotly](https://img.shields.io/badge/Plotly-6.x-3F4F75?logo=plotly)

---

## Funcionalidades

- Carga de archivos `.xlsx` y `.csv`
- Filtros por categoría y rango de fechas
- Indicadores automáticos: Total, Promedio, Mínimo, Máximo y Registros
- Generador de gráficos: Barras, Pastel, Líneas y Dispersión
- Exportación de datos filtrados en CSV

---

## Tecnologías

- Python 3.10+
- Streamlit 1.57
- Pandas 2.2
- Plotly 6.7
- OpenPyXL 3.1

---

## Instalación

### Ubuntu / Linux

```bash
# 1. Verificar Python
python3 --version

# 2. Clonar el repositorio
git clone https://github.com/tu-usuario/plataforma-reportes.git
cd plataforma-reportes

# 3. Crear entorno virtual (recomendado)
python3 -m venv venv
source venv/bin/activate

# 4. Instalar dependencias
pip install -r requirements.txt
```

### Windows

```powershell
# 1. Verificar Python (debe estar en PATH)
python --version

# 2. Clonar el repositorio
git clone https://github.com/tu-usuario/plataforma-reportes.git
cd plataforma-reportes

# 3. Crear entorno virtual (recomendado)
python -m venv venv
venv\Scripts\activate

# 4. Instalar dependencias
pip install -r requirements.txt
```

---

## Ejecución

**Ubuntu / Linux:**
```bash
python3 -m streamlit run app.py
```

**Windows:**
```powershell
python -m streamlit run app.py
```

Luego abrir en el navegador: [http://localhost:8501](http://localhost:8501)

---

## Formato de datos recomendado

El archivo debe tener encabezados en la primera fila. Ejemplo:

```
Fecha,Producto,Categoria,Total
2026-01-01,Laptop,Tecnologia,1200
2026-01-02,Mouse,Accesorios,25
2026-01-03,Teclado,Accesorios,50
2026-01-04,Monitor,Tecnologia,300
```

> Las columnas numéricas deben contener únicamente números.

---

## Autores
•	Esteban Josue Matute Rodriguez
0506200400432
•	Duan Kalek Antúnez Hernández
1503200401515
•	Rony Fabricio Moncada Gonzalez
0704199500278
•	Josué Raúl Navarrete Estrada
0801200311881
•	Erick Roberto Mencia Flores
0801200200160
•	Abiel Isaí Ordóñez Ordóñez
0801200300492

Proyecto académico desarrollado para la asignatura de Análisis y Visualización de Datos.
