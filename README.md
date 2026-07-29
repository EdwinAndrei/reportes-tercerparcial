# Plataforma de Reportes y Análisis de Datos

GRUPO 6

Descripción del Sistema

La Plataforma de Reportes y Análisis de Datos es una aplicación web desarrollada con Streamlit que permite cargar archivos en múltiples formatos (Excel, CSV, Word, PDF y bases de datos SQLite), analizarlos mediante filtros e indicadores, generar gráficos personalizables y exportar reportes combinados en Excel, CSV o PDF, incluyendo los gráficos generados. El sistema conserva el trabajo del usuario mediante un autoguardado y un historial de puntos de guardado manuales.

Requisitos Mínimos

•	Python 3.12 (recomendado; versiones muy recientes como 3.14 pueden presentar incompatibilidades con librerías de terceros).
•	Conexión a internet para la instalación inicial de dependencias.
•	Navegador web moderno (Chrome, Edge o Firefox) para acceder a la interfaz de la aplicación.
•	Las siguientes librerías de Python: streamlit, pandas, numpy, plotly, openpyxl, fpdf2, python-docx, pdfplumber y kaleido.

Guía de Uso Paso a Paso

Pasos de instalación y ejecución
1.	Instalar Python 3.10+ desde https://www.python.org/downloads/ (marcar 'Add Python to PATH').
2.	Descargar o clonar el repositorio del proyecto en una carpeta local.
3.	Abrir una terminal (CMD, PowerShell o Bash) en la carpeta del proyecto.
4.	Ejecutar: pip install -r requirements.txt

Cómo Iniciar la Aplicación

A diferencia de una aplicación de escritorio tradicional, este sistema no requiere un inicio de sesión con usuario y contraseña; el acceso se realiza ejecutando el servidor local de Streamlit. Los pasos son los siguientes:
1.	Abrir una terminal en la carpeta del proyecto.
2.	Ejecutar el comando: py -3.12 -m streamlit run appv3_penultima.py
3.	Esperar a que la terminal muestre la dirección local (por ejemplo, http://localhost:8501).
4.	La aplicación se abrirá automáticamente en el navegador; si no ocurre, copiar y pegar la dirección manualmente.

Cómo Usar las Funciones Principales


1. Cargar archivos.
En la sección “1. Cargar archivos”, seleccione uno o varios archivos en los formatos admitidos. Puede arrastrarlos directamente o usar el explorador de archivos.

2. Seleccionar hojas o tablas.
Si el archivo cargado contiene varias hojas (Excel) o tablas (Word, PDF, bases de datos), el sistema le permitirá elegir cuáles convertir en datasets independientes.

3. Elegir datasets para el análisis.
En la sección “3. Datasets a incluir en el análisis y el reporte”, seleccione uno o varios datasets. Cada uno tendrá su propia pestaña con filtros e indicadores independientes.

4. Crear un gráfico personalizado.
En la sección “4. Generador de Gráficos”, seleccione el dataset, el eje X, el eje Y y el tipo de gráfico. De forma opcional, escriba un título propio, elija el orden de los datos (mayor a menor o menor a mayor) y seleccione una paleta de colores. Presione “Agregar gráfico” para generarlo.

5. Duplicar o eliminar un gráfico.
Debajo de cada gráfico generado encontrará dos botones: “Duplicar este gráfico”, que crea una copia editable con la misma configuración, y “Eliminar este gráfico”, que lo remueve permanentemente de la sesión.

6. Agregar notas y exportar el reporte.
En la sección “5. Exportar Reporte”, revise el resumen de exportación (cantidad de datasets, filas y gráficos), escriba notas u observaciones adicionales si lo desea, y descargue el reporte completo en Excel o PDF, o bien datasets individuales en CSV o Excel.

7. Guardar y recuperar el trabajo.
En la barra lateral, presione “Guardar ahora” para crear una nueva entrada en el historial (identificada por fecha y hora). Puede expandir cualquier entrada del historial y presionar “Cargar” para restaurar ese estado, o “Eliminar” para borrarla individualmente.


