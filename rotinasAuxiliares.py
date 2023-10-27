import os
import polars

def limpar_diretorio(caminho):
    files = os.listdir(caminho)
    for f in files:
        os.remove(caminho + f)


def criar_checkboxes(disciplinas):
    checklist_values = []
    checklist_options1 = []
    for row in disciplinas.rows():
        checklist_options1.append((row[0], row[1]))
        checklist_values.append(row[1])

    checklist_options = [{'label': label, 'value': value} for label, value in checklist_options1]
    return checklist_options, checklist_values
