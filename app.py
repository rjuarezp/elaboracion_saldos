import streamlit as st
import calclibrary
import os
import pandas as pd

st.set_page_config(page_title='Elaboración saldos', page_icon=None, layout='centered', initial_sidebar_state='auto')

st.title('Elaboración de saldos')

file = st.file_uploader("Seleccione el archivo a procesar", accept_multiple_files=False, type='xlsx')

targetFolder = st.text_input('Guardar los archivos procesados en...', value=os.getcwd())

process = st.button('Procesar archivo')

if (process) & (file is not None):
    status = calclibrary.process_file(file, targetFolder, streamlit=True)
    if status == 0:
         st.write('Archivo procesado correctamente')
    elif status == -1:
         st.write('Error')