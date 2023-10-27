import os
import dash
from dash import dcc, html, Input, Output, State
import dash_bootstrap_components as dbc
from dash.exceptions import PreventUpdate
import base64
import io
import pandas as pd

app = dash.Dash(__name__, external_stylesheets=[dbc.themes.MATERIAL])
server = app.server

app.layout = html.Div([
    html.H1("Ranking Simulado CONVET"),

    html.Div([
        html.H3("Matérias do Simulado:"),
        dcc.Checklist(
                id = 'checkboxes',
                options = [
                    {'label': 'Língua Portuguesa', 'value': 'portuguesa'},
                    {'label': 'Noções de direito administrativo e constitucional', 'value': 'direito'},
                    {'label': 'Noções de Raciocínio Lógico e Matemático', 'value': 'raciocinio'},
                    {'label': 'Noções de Informática', 'value': 'informatica'},
                    {'label': 'Disciplinas do Eixo Transversal', 'value': 'eixo_transversal'}
                ],
                value = ['portuguesa', 'direito', 'raciocinio', 'informatica', 'eixo_transversal']
        ),
    ]),

    html.Div([
        html.H3("Arquivos CSV dos resultados das Provas Objetiva Básica, Específica e Discursiva:"),
        dcc.Upload(
                id = 'upload-data',
                children = html.Button('Escolher Arquivos'),
                multiple = True
        ),
        html.Div(id = 'file-list')
    ])
])


def convert_and_save_files(contents, filenames):
    if contents is not None:
        # Limpar a pasta de uploads
        upload_folder = 'uploads'
        for filename in os.listdir(upload_folder):
            file_path = os.path.join(upload_folder, filename)
            try:
                if os.path.isfile(file_path):
                    os.unlink(file_path)
            except Exception as e:
                print(e)

        # Converter e salvar arquivos carregados
        for content, filename in zip(contents, filenames):
            file_content = base64.b64decode(content)
            base64_content = base64.b64encode(file_content).decode('utf-8')
            with open(os.path.join(upload_folder, filename), 'wb') as file:
                file.write(base64_content)


@app.callback(Output('file-list', 'children'),
              Input('upload-data', 'contents'),
              State('upload-data', 'filename'))
def update_output(contents, filenames):
    if contents is None:
        raise PreventUpdate

    convert_and_save_files(contents, filenames)

    return f'Arquivos carregados: {", ".join(filenames)}'


if __name__ == '__main__':
    app.run_server(debug = True)
