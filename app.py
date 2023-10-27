import dash
from dash import dcc, html, Input, Output, State
from dash.exceptions import PreventUpdate
import polars
import rotinasAuxiliares
import os

# Limpa a pasta input-files --------------------------------------------------------------------------------------------
# DESCOMENTAR QUANDO ACABAR O CÓDIGO
# input_files_folder = os.getcwd() + '\\input-files\\'
# rotinasAuxiliares.limpar_diretorio(input_files_folder)

# Inicializa as leituras de arquivos -----------------------------------------------------------------------------------
# Lê o arquivo CSV com as disciplinas deste simulado -------------------------------------------------------------------
disciplinas = polars.scan_csv('disciplinas.csv').collect()
checklist_options, checklist_values = rotinasAuxiliares.criar_checkboxes(disciplinas)

# ----------------------------------------------------------------------------------------------------------------------
app = dash.Dash(__name__, title="Ranking Simulado CONVET")

app.layout = html.Div([

    html.H2("Ranking Simulado CONVET", style={'font-weight': 'bold', 'text-decoration': 'underline'}),

    html.Label("Matérias deste simulado:", style={'font-size': '18px', 'font-weight': 'bold'}),

    # Cria o Checklist com todas as matérias selecionadas
    dcc.Checklist(
        id='checklist',
        style={'font-size': '16px'},
        options=checklist_options,  # Pass the options variable
        value=checklist_values  # Pass the values variable
    ),

    html.Br(),

    html.Label("Selecionar 3 arquivos CSV com as notas:", style={'font-size': '18px', 'font-weight': 'bold'}),
    html.Ul([
        html.Li("Prova Objetiva Geral", style={'margin': '1px'}),
        html.Li("Prova Objetiva Específica", style={'margin': '1px'}),
        html.Li("Prova Discursiva", style={'margin': '1px'}),
    ]),

    dcc.Upload(
        id='btn_csv',
        children = html.Button("Selecione os arquivos .CSV"),
        multiple = True
    ),

    html.Div(id='file-list'),

])


if __name__ == '__main__':
    app.run(debug = True)
