import dash
from dash import dcc, html, Input, Output, State, dash_table
import pandas as pd
import urllib.parse
import polars as pl

app = dash.Dash(__name__)
app.title = 'Editable Table Example'

initial_data = pd.DataFrame({
    'C1': ['Língua Portuguesa',
           'Noções de Direito Administrativo e Constitucional',
           'Noções de Raciocínio Lógico e Matemático',
           'Noções de Informática',
           'Disciplinas do Eixo Transversal'],
    'C2': ['M1', 'M2', 'M3', 'M4', 'M5'],
    'C3': [10, 10, 10, 5, 15],
    'C4': [1, 11, 21, 31, 36],
    'C5': [10, 20, 30, 35, 50],
})

app.layout = html.Div([
    
    # Tabela das matérias ----------------------------------------------------------------------------------------------
    dash_table.DataTable(
        id='editable-table',
        columns=[
            {'name': 'Disciplina', 'id': 'C1'},            
            {'name': 'Quantidade de Questões', 'id': 'C3'},
            {'name': 'Questão Inicial', 'id': 'C4'},
            {'name': 'Questão Final', 'id': 'C5'},
            {'name': 'Código', 'id': 'C2'},
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
            {'if': {'column_id': 'C2'},
             'width': '50px',
             'textAlign': 'center'
             },
            {'if': {'column_id': 'C3'},
             'width': '100px',
             'textAlign': 'center'
             },
            {'if': {'column_id': 'C4'},
             'width': '80px',
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
    dcc.Download(id="salvar-materias")
    # Fim da tabela das matérias ---------------------------------------------------------------------------------------
])


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
        #return dcc.send_data_frame(df.to_csv, "materias.csv")


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


if __name__ == '__main__':
    app.run_server(debug=True)