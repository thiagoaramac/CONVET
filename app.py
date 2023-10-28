import dash
from dash import dcc, html, Input, Output, State, callback, dash_table
import polars
import pandas as pd
import base64
import io
import os
import rotinasAuxiliares
import ranking

# App ------------------------------------------------------------------------------------------------------------------
app = dash.Dash(__name__, title="Ranking Simulado CONVET")
server = app.server

# ----------------------------------------------------------------------------------------------------------------------
# Lê o arquivo CSV com as disciplinas deste simulado -------------------------------------------------------------------
initial_data = pd.DataFrame({
    'C1': ['Língua Portuguesa',
           'Noções de Direito Administrativo e Constitucional',
           'Noções de Raciocínio Lógico e Matemático',
           'Noções de Informática',
           'Disciplinas do Eixo Transversal'],
    'C3': [10, 10, 10, 5, 15],
    'C4': [1, 11, 21, 31, 36],
    'C5': [10, 20, 30, 35, 50],
})

# ----------------------------------------------------------------------------------------------------------------------
# Layout do Site -------------------------------------------------------------------------------------------------------
# ----------------------------------------------------------------------------------------------------------------------
app.layout = html.Div([
    # Título -----------------------------------------------------------------------------------------------------------
    html.H2(
            "Ranking Simulado CONVET",
            style={'font-weight': 'bold', 'text-decoration': 'underline'}
    ),

    # Cria o Checklist com todas as matérias selecionadas---------------------------------------------------------------
    dash_table.DataTable(
        id='editable-table',
        columns=[
            {'name': 'Disciplina', 'id': 'C1'},
            {'name': 'Quantidade de Questões', 'id': 'C3'},
            {'name': 'Questão Inicial', 'id': 'C4'},
            {'name': 'Questão Final', 'id': 'C5'},
        ],
        data=initial_data.to_dict('records'),
        editable=True,  # Enable editing
        row_deletable=True,  # Enable row deletion
        style_header={
                'backgroundColor': 'cadetblue',
                'fontWeight': 'bold',
                'color': 'black'
        },
        style_table={
            'overflowX': 'auto',
            'minWidth': '850px',
            'width': '850px',
            'maxWidth': '850px',
            'whiteSpace': 'normal'
        },
        style_data={
                'backgroundColor': 'rgb(245, 245, 245)',
                'color': 'black'
        },
        style_data_conditional=[{
                'if': {'state': 'selected'},
                'backgroundColor': 'powderblue',  # Change the background color for selected cells
                'border': '1px solid cadetblue'  # Add a border to selected cells
        }],
        style_cell_conditional=[
            {'if': {'column_id': 'C1'},
             'width': '540px',
             'textAlign': 'left'
             },
            {'if': {'column_id': 'C3'},
             'width': '100px',
             'textAlign': 'center'
             },
            {'if': {'column_id': 'C4'},
             'width': '80px',
             'textAlign': 'center'
             },
            {'if': {'column_id': 'C5'},
             'width': '80px',
             'textAlign': 'center'
             },
        ]
    ),
    html.Button('Adicionar Disciplina', id='add-row-button'),
    html.Button("Salvar", id="btn_salvar_materias"),
    dcc.Download(id="salvar-materias"),

    # Salta uma linha --------------------------------------------------------------------------------------------------
    html.Br(),
    html.Br(),

    # Seleciona os arquivos CSV ----------------------------------------------------------------------------------------
    html.Label(
            "Selecionar 3 arquivos CSV com as notas:",
            style={'font-size': '18px', 'font-weight': 'bold'}
    ),

    html.Ul([
        html.Li("Prova Objetiva Geral", style={'margin': '1px'}),
        html.Li("Prova Objetiva Específica", style={'margin': '1px'}),
        html.Li("Prova Discursiva", style={'margin': '1px'}),
    ]),

    # Botão do dcc.upload ----------------------------------------------------------------------------------------------
    dcc.Upload(
            id = 'upload-data',
            children = html.Div([
                'Arraste os arquivos para cá ou ',
                html.A('Selecione')
            ]),
            style = {
                # 'width': '100%',
                'width': '330px',
                'height': '60px',
                'lineHeight': '60px',
                'borderWidth': '1px',
                'borderStyle': 'dashed',
                'borderRadius': '5px',
                'textAlign': 'center',
                'margin': '10px'
            },
            accept = '.csv',
            multiple = True
    ),

    # Div onde vão aparecer os nomes dos arquivos ----------------------------------------------------------------------
    html.Div(id = 'output-data-upload'),

    # Download do arquivo final ----------------------------------------------------------------------------------------
    # dcc.Download(id="download-files"),

])


# ----------------------------------------------------------------------------------------------------------------------
# Funções antes dos Callbacks ------------------------------------------------------------------------------------------
def copiar_arquivos(contents, filename, input_files_folder):

    # Inicializa o DataFrame -------------------------------------------------------------------------------------------
    df = []

    # Decodifica o arquivo CSV -----------------------------------------------------------------------------------------
    content_type, content_string = contents.split(',')
    decoded = base64.b64decode(content_string)
    try:
        if 'csv' in filename:
            # Assume that the user uploaded a CSV file
            df = polars.read_csv(io.StringIO(decoded.decode('utf-8')))
        elif 'xls' in filename:
            # Assume that the user uploaded an Excel file
            df = polars.read_excel(io.BytesIO(decoded))
    except Exception as e:
        print(e)
        return html.Div([
            'Erro no processamento de um arquivo selecionado.'
        ])

    # Salva o arquivo CSV na pasta 'input-files' -----------------------------------------------------------------------
    df.write_csv(input_files_folder + filename)

    # Atualiza o HTML com os nomes dos arquivos importados -------------------------------------------------------------
    return html.Div([
        html.Label(
                'Arquivo: ' + filename + ' - Importado com sucesso.',
                style = {'font-size': '18px'}
        )
    ])


# ----------------------------------------------------------------------------------------------------------------------
# Callbacks ------------------------------------------------------------------------------------------------------------
# ----------------------------------------------------------------------------------------------------------------------

# Adicionar Linhas na tabela -------------------------------------------------------------------------------------------
@app.callback(
    Output('editable-table', 'data'),
    Input('add-row-button', 'n_clicks'),
    Input('editable-table', 'data_previous'),
    State('editable-table', 'data'),
    prevent_initial_call=True
)
def update_table(n_clicks, data_previous, current_data):
    ctx = dash.callback_context
    if not ctx.triggered:
        raise dash.exceptions.PreventUpdate

    trigger_id = ctx.triggered[0]['prop_id'].split('.')[0]

    if trigger_id == 'add-row-button':
        if current_data:
            current_data.append({key: '' for key in current_data[0]})
        else:
            current_data = [{'Column 1': '', 'Column 2': ''}]
    elif trigger_id == 'editable-table':
        if data_previous and len(data_previous) > len(current_data):
            current_data = data_previous

    return current_data

# Salvar Matérias-------------------------------------------------------------------------------------------------------
@app.callback(
        Output('salvar-materias', 'data'),
        Input('btn_salvar_materias', 'n_clicks'),
        Input('editable-table', 'data'),
)
def save_table_data(n_clicks, table_data):
    if n_clicks is not None:
        df = pd.DataFrame(table_data)
        df.to_csv('materias.csv', index = False, encoding = 'utf-8')
    pass


# Atualizar a página após ler CSV --------------------------------------------------------------------------------------
@callback(
          Output('output-data-upload', 'children'),
          Input('upload-data', 'contents'),
          State('upload-data', 'filename'),
          )
def update_output(contents, filename):
    # Limpa a pasta input-files ----------------------------------------------------------------------------------------
    input_files_folder = os.getcwd() + '\\input-files\\'
    output_files_folder = os.getcwd() + '\\output-files\\'
    rotinasAuxiliares.limpar_diretorio(input_files_folder)

    if contents is not None:
        children = [copiar_arquivos(c, n, input_files_folder) for c, n in zip(contents, filename)]
        try:
            ranking.compilar_notas()
            ranking.rankear_alunos()
            arquivo = os.getcwd() + '\\input-files\\CSV_Ranking.csv'
        except:
            pass
        return children


# ----------------------------------------------------------------------------------------------------------------------
# app.run --------------------------------------------------------------------------------------------------------------
if __name__ == '__main__':
    app.run_server(debug = True)
