import streamlit as st
import pandas as pd
from io import BytesIO
from difflib import SequenceMatcher

st.title('ENCONTRAR COINCIDENCIAS')
# Aquí cargo el archivo de entrada
archivo_entrada = st.file_uploader(label='Insertar Archivo', type=["xlsx"], accept_multiple_files=False)

if archivo_entrada is not None:
    # Cargo los DataFrames
    df_ofima = pd.read_excel(archivo_entrada, sheet_name='Ofima')
    df_transfiriendo = pd.read_excel(archivo_entrada, sheet_name='Transfiriendo')

    # Función para encontrar coincidencia aproximada
    def encontrar_coincidencia_aproximada(factura_ofima, facturas_transfiriendo):
        mejor_coincidencia = ""
        mejor_similitud = 0

        # Iterar sobre las facturas en la hoja "Transfiriendo"
        for factura_transfiriendo in facturas_transfiriendo:
            # Calcular la similitud
            similitud = SequenceMatcher(None, str(factura_ofima), str(factura_transfiriendo)).ratio()

            # Actualizar la mejor coincidencia si es necesario
            if similitud > mejor_similitud:
                mejor_similitud = similitud
                mejor_coincidencia = factura_transfiriendo

        return mejor_coincidencia, mejor_similitud * 100  

    # Aplicar la función a cada fila de la hoja "Ofima"
    df_ofima[["Coincidencia Aproximada", "Porcentaje Similitud"]] = df_ofima["FACTURA"].apply(
        lambda x: pd.Series(encontrar_coincidencia_aproximada(x, df_transfiriendo["FACTURA"]))
    )

    # Guardar el DataFrame actualizado en un nuevo archivo Excel en memoria
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_ofima.to_excel(writer, index=False, sheet_name='Ofima_Actualizado')

    # Botón para descargar el archivo
    st.download_button(
        label="Descargar archivo procesado",
        data=output.getvalue(),
        file_name="archivo_actualizado.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
