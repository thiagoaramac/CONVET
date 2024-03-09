import streamlit as st
import pandas as pd
import os
import numpy as np

input_file = os.getcwd() + f'\\output-files\\CSV_Ranking.csv'
concurso = 'CNU'  # Opções: CNU, EMATER
simulado_num = 32

st.write(f"""
# Ranking Simulado CONVET
""")

genre = st.radio(
        "Qual o concurso atual?",
        ["MAPA", "EMATER"],
        index = None,
)

number = st.number_input("Qual o número do simulado?", value = 1, step = 1)

st.write('Concurso selecionado: ', genre, ' - nº', str(number))

uploaded_files = st.file_uploader("Arquivos do concurso atual",
                        accept_multiple_files = True,
                        #type = '.xlsx, .CSV',
                        label_visibility = "visible"
                        )

df = ['']

uploaded_file = st.radio(
        "Qual o concurso atual?",
        [uploaded_files],
        index = None,
)


st.table(df)
