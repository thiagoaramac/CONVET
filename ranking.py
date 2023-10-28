import pandas as pd
import os


input_files_path = os.getcwd() + '\\input-files\\'
files = os.listdir(input_files_path)

input_basico = [file for file in files if "SIC" in file][0]
input_especifico = [file for file in files if "FIC" in file][0]
input_discursiva = [file for file in files if "RSI" in file][0]

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
df_basico = df_basico[['Aluno'] + [col for col in df_basico.columns if col != 'Aluno']]
df_basico.reset_index(drop=True, inplace=True)

# Lê as matérias e cria colunas com as somas de cada matéria ---
df_materias = pd.read_csv(os.getcwd() + '\\materias.csv')
for index, row in df_materias.iterrows():
    materia_atual = row['C1']
    if not pd.isna(materia_atual):
        coluna_inicial = int(row['C4'])
        coluna_final = int(row['C5']) + 1
        df_basico[materia_atual] = df_basico.iloc[:, coluna_inicial:coluna_final].sum(axis=1).round(2)
df_basico = df_basico.drop(df_basico.columns[1:51], axis=1)
# --------------------------------------------------------------
# Cria a soma total da prova objetiva --------------------------
df_basico['Prova Objetiva Geral'] = df_basico.iloc[:, 1:].sum(axis=1).round(2)
# --------------------------------------------------------------
# df_basico.to_csv(os.getcwd() + '\\input-files\\CSV_Basico.csv')


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
df_especifico = df_especifico[['Aluno'] + [col for col in df_especifico.columns if col != 'Aluno']]
# Cria a coluna com a soma de todas as questões ----------
df_especifico['Prova Objetiva Específica'] = df_especifico.iloc[:, 1:].sum(axis=1)
df_especifico = df_especifico.iloc[:, [0, -1]]
# --------------------------------------------------------
df_especifico.reset_index(drop=True, inplace=True)
# df_especifico.to_csv(os.getcwd() + '\\input-files\\CSV_Especifico.csv')


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
df_discursiva['Aluno'] = (df_discursiva['Aluno']).str.title()
df_discursiva = df_discursiva.replace(',', '.', regex=True)
df_discursiva = df_discursiva.apply(pd.to_numeric, errors='ignore')
df_discursiva = df_discursiva.loc[df_discursiva.groupby('Aluno')['Nota'].idxmax()]
df_discursiva.reset_index(drop=True, inplace=True)
df_discursiva.rename(columns={'Nota': 'Prova Discursiva'}, inplace=True)
# df_discursiva.to_csv(os.getcwd() + '\\input-files\\CSV_Discursiva.csv')


# Cria o CSV_NotasFinais
df_nota_final = pd.merge(df_basico, df_especifico, on='Aluno', how='outer')
df_nota_final = pd.merge(df_nota_final, df_discursiva, on='Aluno', how='outer')
df_nota_final['Prova Objetiva (Geral + Específica)'] = df_nota_final.iloc[:, 6:8].sum(axis=1).round(2)
df_nota_final['Nota Final'] = df_nota_final.iloc[:, 8:10].sum(axis=1).round(2)
df_nota_final = df_nota_final.sort_values(by='Prova Objetiva Específica', ascending=False)
df_nota_final = df_nota_final.drop_duplicates(subset='Aluno', keep='first')
df_nota_final = df_nota_final.fillna(0)
df_nota_final = df_nota_final.sort_values(by='Aluno', ascending=True)
df_nota_final.reset_index(drop=True, inplace=True)
df_nota_final.to_csv(os.getcwd() + '\\input-files\\CSV_NotasFinais.csv')

# Limpa a memória RAM e mantém só o df_nota_final
del df_basico, df_especifico, df_discursiva
