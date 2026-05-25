import pandas as pd

caminho_arquivo = r'/mnt/c/dados/pedidos/atendidos-por-item/00072_20260102_20260131_Atendidos_por_item1.csv'
df = pd.read_csv(caminho_arquivo, sep=";",low_memory=False)
print(df.head())