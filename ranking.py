import pandas as pd
import os


input_files_path = os.getcwd() + '\\input-files\\'
files = os.listdir(input_files_path)

input_basico = [file for file in files if "SIC" in file][0]
input_especifico = [file for file in files if "FIC" in file][0]
input_discursiva = [file for file in files if "RSIV" in file][0]

# Gera o CSV_Basico ----------------------------------------------------------------------------------------------------
df_basico = pd.read_csv(os.getcwd() + '\\input-files\\' + input_basico)
df_basico.pop('Estado')
df_basico.pop('Iniciado em')
df_basico.pop('Completo')
df_basico.pop('Tempo utilizado')
df_basico.pop('Avaliar/10,00')
df_basico = df_basico[~df_basico['Sobrenome'].str.startswith('Média geral')]
df_basico['Aluno'] = (df_basico['Nome'] + ' ' + df_basico['Sobrenome']).str.title()
df_basico = df_basico.drop(['Nome', 'Sobrenome'], axis=1)
df_basico = df_basico.replace('-', '0,00', regex=True)
df_basico = df_basico.replace(',', '.', regex=True)
df_basico = df_basico.apply(pd.to_numeric, errors='ignore')
df_basico.reset_index(drop=True, inplace=True)
df_basico.to_csv(os.getcwd() + '\\input-files\\CSV_Basico.csv')

# Gera o CSV_Especifico ------------------------------------------------------------------------------------------------
df_especifico = pd.read_csv(os.getcwd() + '\\input-files\\' + input_especifico)
df_especifico.pop('Estado')
df_especifico.pop('Iniciado em')
df_especifico.pop('Completo')
df_especifico.pop('Tempo utilizado')
df_especifico.pop('Avaliar/10,00')
df_especifico = df_especifico[~df_especifico['Sobrenome'].str.startswith('Média geral')]
df_especifico['Aluno'] = (df_especifico['Nome'] + ' ' + df_especifico['Sobrenome']).str.title()
df_especifico = df_especifico.drop(['Nome', 'Sobrenome'], axis=1)
df_especifico = df_especifico.replace('-', '0,00', regex=True)
df_especifico = df_especifico.replace(',', '.', regex=True)
df_especifico = df_especifico.apply(pd.to_numeric, errors='ignore')
df_especifico.reset_index(drop=True, inplace=True)
df_especifico.to_csv(os.getcwd() + '\\input-files\\CSV_Especifico.csv')

# Gera o CSV_Discursiva ------------------------------------------------------------------------------------------------
df_discursiva = pd.read_csv(os.getcwd() + '\\input-files\\' + input_discursiva)
df_discursiva.pop('Identificador')
df_discursiva.pop('Status')
df_discursiva.pop('Nota máxima')
df_discursiva.pop('Nota pode ser alterada')
df_discursiva.pop('Última modificação (envio)')
df_discursiva.pop('Última modificação (nota)')
df_discursiva.pop('Comentários de feedback')
df_discursiva = df_discursiva.dropna(subset=['Nota'])
df_discursiva.rename(columns={'Nome completo': 'Aluno'}, inplace=True)
df_discursiva.reset_index(drop=True, inplace=True)
df_discursiva.to_csv(os.getcwd() + '\\input-files\\CSV_Discursiva.csv')

