MAPEO POR VENTANA GEOMECÁNICA — Versión Python/Streamlit
=========================================================

REQUISITOS
----------
Python 3.9 o superior

Instalar dependencias:
    pip install streamlit pandas openpyxl

CÓMO EJECUTAR
-------------
    streamlit run app.py

Esto abrirá la aplicación en tu navegador en http://localhost:8501

FUNCIONALIDADES
---------------
- Registro de cabecera: TD2, coordenadas, largo/altura, orientación, litología, mapeador
- Tabla de discontinuidades por familias (1-9):
    * Distancia, Dip/DipDir, Tipo de estructura, N° estructuras
    * Abertura, Espesor, Continuidad, Espaciamiento, N° extremos visibles
    * Terminación, Relleno 1, Relleno 2 (valor mínimo = peor caso)
    * JRC, Rugosidad (1-9), Forma, Alteración
- Cálculo automático de ratings RMR'89 y RMR'76 por fila
- Catálogos de referencia en la barra lateral
- Resumen: PROM 1/2/3 y JV
- Exportar a JSON, CSV y Excel (.xlsx)

NOTAS
-----
- El valor de relleno final usa el MÍNIMO entre Relleno 1 y Relleno 2 (peor caso)
- La distancia se valida contra el Largo definido en el encabezado
- Catálogos basados en Bieniawski RMR'89 y RMR'76 (MMG / Las Bambas 2021)
