import dash
from dash import dcc, html, Input, Output, State, callback, dash_table
import polars
import pandas as pd
import base64
import io
import os
import rankingCIDASC
import rankingCNU
import rankingEMATER
import rotinasAuxiliares


# App ------------------------------------------------------------------------------------------------------------------
concurso = 'CNU'  # Opções: CNU, EMATER, CIDASC
simulado_num = 32
app = dash.Dash(__name__, title="Ranking Simulado CONVET")
server = app.server

# ----------------------------------------------------------------------------------------------------------------------
# Lê o arquivo CSV com as disciplinas deste simulado -------------------------------------------------------------------
initial_data = pd.DataFrame({
    'C1': ['Eixo temático 1',
           'Eixo temático 2',
           'Eixo temático 3',
           'Eixo temático 4',
           'Eixo temático 5'],
    'C3': [10, 10, 10, 10, 10],
    'C4': [1, 11, 21, 31, 41],
    'C5': [10, 20, 30, 40, 50],
})

# ----------------------------------------------------------------------------------------------------------------------
# Layout do Site -------------------------------------------------------------------------------------------------------
# ----------------------------------------------------------------------------------------------------------------------
app.layout = html.Div([
    # Título -----------------------------------------------------------------------------------------------------------
    html.H2(
            f"Ranking Simulado nº{simulado_num} CONVET - {concurso}",
            style = {'font-weight': 'bold', 'text-decoration': 'underline'}
    ),

    # Seleciona os arquivos CSV ----------------------------------------------------------------------------------------
    html.Label(
            "Selecionar 3 arquivos CSV com as notas:",
            style = {'font-size': '18px', 'font-weight': 'bold'}
    ),

    html.Ul([
        html.Li("Prova Objetiva Geral", style = {'margin': '1px'}),
        html.Li("Prova Objetiva Específica", style = {'margin': '1px'}),
        html.Li("Prova Discursiva", style = {'margin': '1px'}),
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
            print('')
            print('Início da Execução')
            print('------------------------------------------------')
            print("Compilando notas....")
            if concurso == 'CNU':
                rankingCNU.compilar_notas()
            if concurso == 'EMATER':
                rankingEMATER.compilar_notas()
            if concurso == 'CIDASC':
                rankingCIDASC.compilar_notas()
            print('------------------------------------------------')
            print("Rankeando alunos....")
            if concurso == 'CNU':
                rankingCNU.rankear_alunos(simulado_num)
            if concurso == 'EMATER':
                rankingEMATER.rankear_alunos(simulado_num)
            if concurso == 'CIDASC':
                rankingCIDASC.rankear_alunos(simulado_num)
            print('------------------------------------------------')
            print("Operação Finalizada com Sucesso!")
            print('------------------------------------------------')
            print("Formatando planilha Excel....")
            if concurso == 'CNU':
                rankingCNU.arrumar_excel(simulado_num)
            if concurso == 'EMATER':
                rankingEMATER.arrumar_excel(simulado_num)
            if concurso == 'CIDASC':
                rankingCIDASC.arrumar_excel(simulado_num)
            # print("Colocando a macro de copia na planilha Excel....")
            # if concurso == 'CNU':
            #   rankingCNU.colocar_macro(simulado_num)
            print("Ajustando colunas da planilha Excel....")
            if concurso == 'CNU':
                rankingCNU.formatar_excel(simulado_num)
            if concurso == 'EMATER':
                rankingEMATER.formatar_excel(simulado_num)
            if concurso == 'CIDASC':
                rankingCIDASC.formatar_excel(simulado_num)
            print('------------------------------------------------')
            print('')
            print('Operação finalizada com sucesso!')
        except:
            pass
        return children


# ----------------------------------------------------------------------------------------------------------------------
# app.run --------------------------------------------------------------------------------------------------------------
if __name__ == '__main__':
    app.run_server(debug = True)
