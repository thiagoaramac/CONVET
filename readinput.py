import polars


def lerarquivoCSV():
    df = polars.scan_csv('materias.csv').collect()
    return df


df = lermaterias()

print(df)
print(df[2, 0])
