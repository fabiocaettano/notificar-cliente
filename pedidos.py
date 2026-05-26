import pandas as pd
import glob


caminho_arquivo = r'/mnt/c/dados/pedidos/atendidos-por-item/*.csv'

lista_arquivos = glob.glob(caminho_arquivo)

dfs = []

for arquivo in lista_arquivos:
    df = pd.read_csv(arquivo, sep=";",low_memory=False)
    dfs.append(df)

df_final = pd.concat(dfs, ignore_index=True)
print(df_final.head())
print(df_final.info())
print(f"Total de linhas: {len(df_final)}")

# Manter apenas as três colunas desejadas
df_filtrado = df_final[['PED', 'TP_PED', 'CLIENTE']]

# Remover linhas duplicadas
df_filtrado = df_filtrado.drop_duplicates()

print(df_filtrado.head())
print(df_filtrado.info())
print(f"Total de linhas após filtragem: {len(df_filtrado)}")

caminho_arquivo = r'/mnt/c/dados/mcu/R55001A_ECT0001.txt'


df_mcu = pd.read_csv(
    caminho_arquivo,
    sep="\t",
    skiprows=2
)
#df_mcu['Unidades de Negócios'] = df_mcu['Unidades de Negócios'].astype(str).str.zfill(8)
#df_mcu['Subordinação administrativa'] = df_mcu['Subordinação administrativa'].astype(str).str.zfill(8)
print(df_mcu.head())
print(df_mcu.info())
print(f"Total de linhas no DataFrame MCU: {len(df_mcu)}")
# Manter apenas as três colunas desejadas
df_filtrado_mcu = df_mcu[['Unidades de Negócios', 'Nº Cad Geral', 'Email da Unidade','Subordinação administrativa']]

# Remover linhas duplicadas
df_filtrado_mcu = df_filtrado_mcu.drop_duplicates()

df_filtrado_mcu['Unidades de Negócios'] = df_filtrado_mcu['Unidades de Negócios'].astype(str).str.zfill(8)


df_filtrado_mcu['Subordinação administrativa'] = (
    df_filtrado_mcu['Subordinação administrativa']
    .astype(str)
    .str.replace(r'\.0$', '', regex=True)
    .str.zfill(8)
)

df_filtrado_mcu['Email Subordinação Administrativa'] = (
    df_filtrado_mcu['Subordinação administrativa']
    .map(
        df_filtrado_mcu
        .set_index('Unidades de Negócios')['Email da Unidade']
    )
)

df_filtrado_mcu['Email Subordinação Administrativa'] = (
    df_filtrado_mcu['Email Subordinação Administrativa']
    .fillna('')
)


print(df_filtrado_mcu.head())
print(df_filtrado_mcu.info())
print(f"Total de linhas após filtragem: {len(df_filtrado_mcu)}")


