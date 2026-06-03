import glob
import pandas as pd
import random
import requests
import json
import os
from dotenv import load_dotenv
from typing import List, Dict
from datetime import datetime, time
import math
import time
import numpy as np
import sys

#from envio_email_v01 import enviar_email_via_webhook

# Configurações
# Carregar variáveis do arquivo .env
load_dotenv()
# resumo da operação para monitoramento e análise posterior
resumo = {}

# Dicionário para mapear as SEs aos seus respectivos códigos
#se_dict = {}
se_dict = {
    "CS - CORREIOS SEDE": "CS",
    "SE - ACRE": "SE/ACR",
    "SE - ALAGOAS": "SE/AL",
    "SE - AMAPÁ": "SE/AP",
    "SE - AMAZONAS": "SE/AM",
    "SE - BAHIA": "SE/BA",
    "SE - BRASÍLIA": "SE/BSB",
    "SE - CEARÁ": "SE/CE",
    "SE - ESPIRITO SANTO": "SE/ES",
    "SE - GOIÁS": "SE/GO",
    "SE - MARANHÃO": "SE/MA",
    "SE - MINAS GERAIS": "SE/MG",
    "SE - MATO GROSSO DO SUL": "SE/MS",
    "SE - MATO GROSSO": "SE/MT",
    "SE - RONDONIA": "SE/RO",
    "SE - PARÁ": "SE/PA",
    "SE - PARAÍBA": "SE/PR",
    "SE - PERNAMBUCO": "SE/PE",
    "SE - PIAUÍ": "SE/PI",
    "SE - PARANÁ": "SE/PR",
    "SE - RIO DE JANEIRO": "SE/RJ",
    "SE - RIO GRANDE DO NORTE": "SE/RN",
    "SE - RIO GRANDE DO SUL": "SE/RS",
    "SE - RORAIMA": "SE/RR",
    "SE - SANTA CATARINA": "SE/SC",
    "SE - SERGIPE": "SE/SE",
    "SE - SÃO PAULO METROPOLITANA": "SE/SPM",
    "SE - SÃO PAULO INTERIOR": "SE/SPI",
    "SE - TOCANTINS": "SE/TO"
}

def importarObjetosEmDistribuicao(
    diretorio: str,
    numeroLinhasPular: int,
    intervalo_colunas: str,
    extensao: str = ".xlsx"
) -> pd.DataFrame:
    """
    Importa e consolida dados de múltiplos arquivos Excel de um diretório.

    Args:
        diretorio (str): Caminho do diretório contendo os arquivos.
        numeroLinhasPular (int): Número de linhas a serem puladas.
        intervalo_colunas (str): Intervalo de colunas (ex: "A:F").
        extensao (str): Extensão dos arquivos (default: ".xlsx").

    Returns:
        pd.DataFrame: DataFrame consolidado e tratado.
    """

    lista_dfs = []

    # Listar arquivos do diretório com a extensão desejada
    arquivos = [
        os.path.join(diretorio, f)
        for f in os.listdir(diretorio)
        if f.lower().endswith(extensao)
    ]

    for caminho_arquivo in arquivos:
        try:
            df = pd.read_excel(
                caminho_arquivo,
                header=numeroLinhasPular,
                usecols=intervalo_colunas,
                engine="openpyxl"
            )

            # Limpeza
            df = df.dropna(how='all')
            df.columns = df.columns.str.strip()
            df = df.applymap(lambda x: x.strip() if isinstance(x, str) else x)

            # Opcional: adicionar coluna com nome do arquivo (ótimo para rastreabilidade)
            df["arquivo_origem"] = os.path.basename(caminho_arquivo)

            lista_dfs.append(df)

        except Exception as e:
            print(f"Erro ao processar {caminho_arquivo}: {e}")

    # Concatenar todos os DataFrames
    if lista_dfs:
        df_final = pd.concat(lista_dfs, ignore_index=True)
    else:
        df_final = pd.DataFrame()

    return df_final

def importarPedidosCSV(caminho_arquivo, separador) -> pd.DataFrame:    
    """
    Importa dados de pedidos a partir de arquivos CSV.

    Args:
        caminho_arquivo (str): Caminho do arquivo CSV.
        separador (str): Separador dos dados no arquivo.

    Returns:
        pd.DataFrame: DataFrame com os dados importados.
    """
    lista_arquivos = glob.glob(caminho_arquivo)
    dfs = []
    for arquivo in lista_arquivos:
        df = pd.read_csv(arquivo, sep=separador,low_memory=False,index_col=False,dtype=str)
        dfs.append(df)
    df_final = pd.concat(dfs, ignore_index=True)
    return df_final

def importarMCU(caminho_arquivo, separador, numeroLinhasPular) -> pd.DataFrame:    
    """
    Importa dados de pedidos a partir de arquivos CSV.

    Args:
        caminho_arquivo (str): Caminho do arquivo CSV.
        separador (str): Separador dos dados no arquivo.
        numeroLinhasPular (int): Número de linhas a serem puladas.

    Returns:
        pd.DataFrame: DataFrame com os dados importados.
    """
    lista_arquivos = glob.glob(caminho_arquivo)
    dfs = []
    for arquivo in lista_arquivos:
        df = pd.read_csv(arquivo, sep=separador,skiprows=numeroLinhasPular,low_memory=False,index_col=False,dtype=str)
        dfs.append(df)
    df_final = pd.concat(dfs, ignore_index=True)
    return df_final

def importarTipoPedido(caminho_arquivo, numeroLinhasPular, intervalo_colunas) -> pd.DataFrame:    
    """
    Importa dados de tipo de pedido a partir de um arquivo Excel.

    Args:
        caminho_arquivo (str): Caminho do arquivo Excel.
        numeroLinhasPular (int): Número de linhas a serem puladas.
        intervalo_colunas (str): Intervalo de colunas (ex: "A:F").

    Returns:
        pd.DataFrame: DataFrame com os dados importados.
    """
    df = pd.read_excel(caminho_arquivo, header=numeroLinhasPular, usecols=intervalo_colunas)
    # Remover possíveis linhas vazias que costumam vir em relatórios exportados
    df = df.dropna(how='all')    
    # Remove espaços extras no início e fim dos nomes das colunas
    df.columns = df.columns.str.strip()
    # Se houver espaços extras dentro das células de texto
    df = df.applymap(lambda x: x.strip() if isinstance(x, str) else x)
    return df

def importarDescricaoSE(caminho_arquivo, numeroLinhasPular, intervalo_colunas) -> dict:    
    """
    Importa dados da descrição da SE a partir de um arquivo Excel.

    Args:
        caminho_arquivo (str): Caminho do arquivo Excel.
        numeroLinhasPular (int): Número de linhas a serem puladas.
        intervalo_colunas (str): Intervalo de colunas (ex: "A:F").

    Returns:
        dict: Dicionário com os dados importados.
    """
    df = pd.read_excel(caminho_arquivo, header=numeroLinhasPular, usecols=intervalo_colunas)
    # Remover possíveis linhas vazias que costumam vir em relatórios exportados
    df = df.dropna(how='all')    
    # Remove espaços extras no início e fim dos nomes das colunas
    df.columns = df.columns.str.strip()
    # Se houver espaços extras dentro das células de texto
    df = df.applymap(lambda x: x.strip() if isinstance(x, str) else x)

    se_dict = df.set_index("SE DESCRICAO COMPLETA")["SE SIGLA"].to_dict()
    
    return json.dumps(se_dict, ensure_ascii=False)

def unificarPedidos(df_pedidos_ajustar: pd.DataFrame) -> pd.DataFrame:
   """
   Unifica os dados de pedidos, mantendo apenas as colunas essenciais e eliminando duplicados.
        Args:
            df_pedidos_ajustar (pd.DataFrame): DataFrame contendo os dados dos pedidos a serem ajustados.
        Returns:
            pd.DataFrame: DataFrame unificado com as colunas "PED", "TP_PED" e "CLIENTE", sem duplicados.
   """ 
   df_filtrado = df_pedidos_ajustar[['PED', 'TP_PED', 'CLIENTE']]    
   df_filtrado = df_filtrado.drop_duplicates()
   return df_filtrado

def incluirEmailParaSubordinacaoAdministrativa(df_email_ajustar: pd.DataFrame) -> pd.DataFrame:
    """
    Inclui a coluna de email da subordinação administrativa no dataset de email das unidades.
    Args:
        df_email_ajustar (pd.DataFrame): DataFrame contendo os dados dos emails das unidades.
    Returns:
        pd.DataFrame: DataFrame com a coluna de email da subordinação administrativa incluída.
    """
    
    df_email_ajustar['Subordinação administrativa'] = (
    df_email_ajustar['Subordinação administrativa']
    .astype(str)
    .str.replace(r'\.0$','', regex=True)
    .str.strip()
    .str.zfill(8))

    df_email_ajustar['Email da Subordinação'] = (
    df_email_ajustar['Subordinação administrativa']
    .map(
        df_email_ajustar
        .set_index('Unidades de Negócios')['Email da Unidade']
        )
    )

    df_email_ajustar['Email da Subordinação'] = (
    df_email_ajustar['Email da Subordinação']
    .fillna('')
    )
    
    return df_email_ajustar

def ajustarColunasDatasetEmail(df_email_ajustar: pd.DataFrame) -> pd.DataFrame:  
    """
        Ajusta as colunas do dataset de email, mantendo apenas as colunas essenciais para o processo de notificação.    
        Args:
            df_email_ajustar (pd.DataFrame): DataFrame contendo os dados dos emails das unidades a serem ajustados.
        Returns:
            pd.DataFrame: DataFrame ajustado com as colunas essenciais para o processo de notificação.
    """  
    # Utilizar copy() para evitar o SettingWithCopyWarning
    # Importante gerar uma cópia indendente e não uma view, para evitar o SettingWithCopyWarning e garantir que as alterações sejam aplicadas corretamente.     
    
    df_filtrado_email = df_email_ajustar[['Unidades de Negócios', 'Nº Cad Geral', 'Email da Unidade','Subordinação administrativa','Email da Subordinação','Descrição DR','DR','Tipo do Órgão', 'Descrição Tp. órgão', 'Status do Órgão']].copy()    
    return df_filtrado_email

def incluirColunaEmail(df_email: pd.DataFrame, df_objetos: pd.DataFrame) -> pd.DataFrame:
    """
    Inclui a coluna de email no dataset de objetos de distribuição.
    Args:
        df_email (pd.DataFrame): DataFrame contendo os dados dos emails das unidades.
        df_objetos (pd.DataFrame): DataFrame contendo os dados dos objetos de distribuição.
    Returns:
        pd.DataFrame: DataFrame com a coluna de email incluída.
    """
    # Realizar o Merge (PROCV)
    # Selecionamos apenas as colunas necessárias do df_email para não poluir o dataset
    colunas_interesse_email = ['MCU Unidade', 'Email da Unidade', 'Email da Subordinação','Tipo do Órgão', 'Descrição Tp. órgão', 'Status do Órgão']
   
    df_resultado = pd.merge(
        df_objetos,
        df_email[colunas_interesse_email],
        left_on='MCU (Unidade Distribuição)',
        right_on='MCU Unidade',
        how='left'
    )        
    return df_resultado

def incluirColunaEmailNosPedidos(df_email: pd.DataFrame, df_pedidos: pd.DataFrame) -> pd.DataFrame:
    """
    Inclui a coluna de email no dataset de pedidos.
    Args:
        df_email (pd.DataFrame): DataFrame contendo os dados dos emails das unidades.
        df_pedidos (pd.DataFrame): DataFrame contendo os dados dos pedidos.
    Returns:
        pd.DataFrame: DataFrame com a coluna de email incluída.
    """
    df_pedidos = df_pedidos.merge(
    df_email[['CLIENTE', 'Email da Unidade', 'Email da Subordinação','SE Cliente','MCU Unidade','Tipo do Órgão']],
    on='CLIENTE',
    how='left'
    )
    return df_pedidos   

def incluirColunaEmailNosObjetosDistribuicao(df_pedidos: pd.DataFrame, df_objetos_distribuicao: pd.DataFrame) -> pd.DataFrame:
    """
    Inclui a coluna de email no dataset de objetos de distribuição, utilizando os dados dos pedidos para realizar o merge.
    Args:
        df_pedidos (pd.DataFrame): DataFrame contendo os dados dos pedidos, incluindo as colunas de email.
        df_objetos_distribuicao (pd.DataFrame): DataFrame contendo os dados dos objetos de distribuição.
    Returns:

    """
    df_objetos_distribuicao = df_objetos_distribuicao.merge(
    df_pedidos[['id_pedido', 'CLIENTE','Email da Unidade Cliente', 'Email da Subordinação Cliente', 'SE Cliente','MCU Unidade Cliente','Tipo do Órgão Cliente']],
    on='id_pedido',
    how='left'
    )
    return df_objetos_distribuicao

def ajustarColunasObjetosEmDistribuicao(df_objetos_em_distribuicao: pd.DataFrame) -> pd.DataFrame:
    """
    Ajusta as colunas do DataFrame de objetos em distribuição.
    Args:
        df_objetos_em_distribuicao (pd.DataFrame): DataFrame contendo os dados dos objetos de distribuição.
    Returns:
        pd.DataFrame: DataFrame com as colunas ajustadas.
    """
    # Extrair os primeiros dois caracteres para nova coluna "Zona de Separação"
    df_objetos_em_distribuicao['Zona de Separação'] = df_objetos_em_distribuicao['Tipo do Pedido'].str[2:]  # Remover os primeiros dois caracteres
    df_objetos_em_distribuicao['Tipo do Pedido'] = df_objetos_em_distribuicao['Tipo do Pedido'].str[:2]  # Manter apenas os dois primeiros caracteres
    return df_objetos_em_distribuicao

def excluirColunas(df: pd.DataFrame,lista_colunas_excluir: list) -> pd.DataFrame:
    """
    Exclui colunas específicas do DataFrame.
    Args:
        df (pd.DataFrame): DataFrame contendo os dados.
        lista_colunas_excluir (list): Lista com os nomes das colunas a serem excluídas.
    Returns:
        pd.DataFrame: DataFrame com as colunas excluídas.
    """
    # Excluir as colunas indesejadas 
    df_atualizado = df.drop(columns=lista_colunas_excluir, errors='ignore')
    return df_atualizado

def incluirCategoria(df_objetos_em_distribuicao: pd.DataFrame, df_tipo_pedido: pd.DataFrame) -> pd.DataFrame:
    """
    Inclui a coluna de categoria no dataset de objetos de distribuição, utilizando os dados do tipo de pedido para realizar o merge.
    Args:
        df_objetos_em_distribuicao (pd.DataFrame): DataFrame contendo os dados dos objetos de distribuição.
        df_tipo_pedido (pd.DataFrame): DataFrame contendo os dados do tipo de pedido.
    Returns:
        pd.DataFrame: DataFrame com a coluna de categoria incluída.
    """
    # Garantir que as colunas de junção sejam strings
    df_objetos_em_distribuicao['Tipo do Pedido'] = df_objetos_em_distribuicao['Tipo do Pedido'].astype(str)
    df_tipo_pedido['Tipo de Pedido'] = df_tipo_pedido['Tipo do Pedido'].astype(str)

    # Merge (similar a um PROCV) com base na coluna "Tipo de Pedido"
    df_objetos_em_distribuicao = df_objetos_em_distribuicao.merge(
        df_tipo_pedido[['Tipo do Pedido', 'Categoria']],
        on='Tipo do Pedido',
        how='left'
    )
   
    return df_objetos_em_distribuicao

def eliminarRegistrosSemEmail(df_objetos_em_distribuicao: pd.DataFrame) -> pd.DataFrame:
    """
    Elimina os registros do DataFrame de objetos em distribuição que não possuem email válido na coluna "Email da Unidade".
    Args:
        df_objetos_em_distribuicao (pd.DataFrame): DataFrame contendo os dados dos objetos de distribuição.
    Returns:
        pd.DataFrame: DataFrame com os registros sem email eliminados.
    """
    # Filtrar o DataFrame para manter apenas os registros onde "Email da Unidade" não é nulo ou vazio
    df_atualizado = df_objetos_em_distribuicao[df_objetos_em_distribuicao['Email da Unidade'].notna() & (df_objetos_em_distribuicao['Email da Unidade'].str.strip() != '')]
    return df_atualizado

def manterRegistrosConformeDataVigente(df_completo: pd.DataFrame) -> pd.DataFrame:
    """
    Mantém apenas os registros do DataFrame completo onde a "Data Prevista" é igual ou posterior à data atual.
    Args:
        df_completo (pd.DataFrame): DataFrame contendo os dados completos dos objetos de distribuição.
    Returns:

    """
    # Converter a coluna "Data Prevista" para datetime, se ainda não estiver nesse formato
    #df_completo['Data Prevista'] = pd.to_datetime(df_completo['Data Prevista'], dayfirst=True, errors='coerce')
    df_completo['Data Prevista'] = pd.to_datetime(df_completo['Data Prevista'],format='%d/%m/%Y', errors='coerce')

    # Obter a data atual
    data_atual = pd.to_datetime(datetime.now().date())    

    # Filtrar o DataFrame para manter apenas os registros onde "Data Prevista" é igual à data atual
    #df_completo_atualizado = df_completo[df_completo['Data Prevista'] >= data_atual]    
    df_completo_atualizado = df_completo

    return df_completo_atualizado

def excluirEmailsInvalidos(df: pd.DataFrame, nomeDaColunaEmail: str) -> pd.DataFrame:
    """
    Exclui os registros do DataFrame que possuem emails inválidos na coluna especificada.
    Args:
        df (pd.DataFrame): DataFrame contendo os dados.
        nomeDaColunaEmail (str): Nome da coluna que contém os emails.
    Returns:
        pd.DataFrame: DataFrame com os registros de emails inválidos eliminados.
    """
    # Expressão regular para validar emails
    regex_email = r'^[\w\.-]+@[\w\.-]+\.\w+$'
   
    # Filtrar o DataFrame para manter apenas os registros onde o email é válido
    df_filtrado = df[df[nomeDaColunaEmail].str.match(regex_email, na=False)]
   
    return df_filtrado

def corrigirFormatoDoEmail(df: pd.DataFrame, nomeDaColunaEmail: str) -> pd.DataFrame:
    """
    Corrige o formato dos emails na coluna especificada, padronizando-os e eliminando os que são inválidos.
    Args:
        df (pd.DataFrame): DataFrame contendo os dados.
        nomeDaColunaEmail (str): Nome da coluna que contém os emails.
    Returns:
        pd.DataFrame: DataFrame com os emails corrigidos.
    """
    # Expressão regular para validar emails
    regex_email = r'^[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$'

    # Padroniza antes (boa prática)
    df[nomeDaColunaEmail] = df[nomeDaColunaEmail].astype(str).str.strip().str.lower()

    # Substitui emails inválidos por vazio
    df.loc[
        ~df[nomeDaColunaEmail].str.match(regex_email, na=False),
        nomeDaColunaEmail
    ] = ''
  
    return df

def concatenarColunasDePedidoComTipo(tipo: str,df: pd.DataFrame) -> pd.DataFrame:
    """
    Concatena as colunas de pedido com o tipo de pedido para criar um identificador único de pedido.
    Args:
        tipo (str): Tipo de concatenação ("objeto" ou "pedido").
        df (pd.DataFrame): DataFrame contendo os dados dos objetos de distribuição ou dos pedidos.
    Returns:
        pd.DataFrame: DataFrame com a coluna "id_pedido" criada.
    """
    if tipo == "objeto":
        # Concatenar as colunas "Número do Pedido" e "Tipo do Pedido" em uma nova coluna "Pedido Completo"
        df['id_pedido'] = df['Número do Pedido'].astype(str) + '_' + df['Tipo do Pedido'].astype(str)
        return df 
    elif tipo == "pedido":
        # Concatenar as colunas "PED" e "TP_PED" em uma nova coluna "Pedido Completo"
        df['id_pedido'] = df['PED'].astype(str) + '_' + df['TP_PED'].astype(str)
        return df
 
def alterarTipoDaColuna(df: pd.DataFrame, tipo: str, nomeDaColuna: str) -> pd.DataFrame:
    """
    Altera o tipo da coluna especificada.
    Args:
        df (pd.DataFrame): DataFrame contendo os dados.
        tipo (str): Tipo de dado para o qual a coluna será convertida.
        nomeDaColuna (str): Nome da coluna a ser convertida.
    Returns:
        pd.DataFrame: DataFrame com a coluna do tipo especificado.
    """
    df[nomeDaColuna] = df[nomeDaColuna].astype(tipo)
    return df

def padronizarColunaMcu(df: pd.DataFrame, nomeDaColuna: str) -> pd.DataFrame:
    """
    Padroniza a coluna de MCU, garantindo que os valores sejam strings, sem espaços e com zeros à esquerda para completar 8 caracteres.
    Args:
        df (pd.DataFrame): DataFrame contendo os dados.
        nomeDaColuna (str): Nome da coluna a ser padronizada.
    Returns:
        pd.DataFrame: DataFrame com a coluna padronizada.
    """
    df[nomeDaColuna] = df[nomeDaColuna].astype(str).str.strip().str.zfill(8)
    return df

def renomearColuna(df: pd.DataFrame, colunaAtual: str, colunaNova: str) -> pd.DataFrame:
    """
    Renomeia uma coluna do DataFrame.
    Args:
        df (pd.DataFrame): DataFrame contendo os dados.
        colunaAtual (str): Nome da coluna atual.
        colunaNova (str): Novo nome da coluna.
    Returns:
        pd.DataFrame: DataFrame com a coluna renomeada.
    """
    df.rename(columns={colunaAtual: colunaNova}, inplace=True)
    return df

def agruparDadosObjetosDistribuicao(df_completo: pd.DataFrame) -> pd.DataFrame:
    """
    Agrupa os dados dos objetos de distribuição por unidade de distribuição.
    Args:
        df_completo (pd.DataFrame): DataFrame contendo os dados dos objetos de distribuição.
    Returns:
        pd.DataFrame: DataFrame com os dados agrupados.
    """
    # Primeiro, faça o groupby nas colunas que você quer manter e as que deseja agregar
    df_grouped = df_completo.groupby(
    ['SE','MCU (Unidade Distribuição)', 'Unidade Distribuição','Email da Unidade', 'Email da Subordinação', 'Data Prevista' ]
    ).agg({
        'Unidade Destino': lambda x: list(x),
        'CEP Destino': lambda x: list(x),
        'Objeto': lambda x: list(x),          
        'Data Postagem': lambda x: list(x),
        'Número do Pedido': lambda x: list(x) ,
        'Tipo do Pedido': lambda x: list(x),
        'Zona de Separação': lambda x: list(x),
        'Categoria': lambda x: list(x)

    }).reset_index()

    return df_grouped

def agruparDadosObjetosClientes(df_completo: pd.DataFrame) -> pd.DataFrame:
    """
    Agrupa os dados dos objetos de clientes por unidade de distribuição.
    Args:
        df_completo (pd.DataFrame): DataFrame contendo os dados dos objetos de clientes.
    Returns:
        pd.DataFrame: DataFrame com os dados agrupados.
    """
    # Primeiro, faça o groupby nas colunas que você quer manter e as que deseja agregar
    df_grouped = df_completo.groupby(
    ['SE Cliente','MCU Unidade Cliente', 'Unidade Destino','Email da Unidade Cliente', 'Email da Subordinação Cliente', 'Data Prevista' ]
    ).agg({
        'Unidade Distribuição': lambda x: list(x),
        'CEP Destino': lambda x: list(x),
        'Objeto': lambda x: list(x),          
        'Data Postagem': lambda x: list(x),
        'Número do Pedido': lambda x: list(x) ,
        'Tipo do Pedido': lambda x: list(x),
        'Zona de Separação': lambda x: list(x),
        'Categoria': lambda x: list(x)

    }).reset_index()

    return df_grouped

def criarEmailHtmlObjetoDistribuicao(row: pd.Series) -> tuple:
    """
    Cria o conteúdo HTML do email para uma unidade de distribuição
   
    Args:
        row: Linha do DataFrame agrupado contendo todos os dados
       
    Returns:
        tuple: (assunto, corpo_html, emails_destinatarios)
    """
   
    # Extrair dados da linha
    unidade_distribuicao = row['Unidade Distribuição']
    data_prevista = row['Data Prevista']
    email_unidade = row['Email da Unidade']
    email_subordinacao = row['Email da Subordinação']
    se = row['SE']
   
    # Criar lista de destinatários
    emails_destinatarios = [email_unidade]
    if pd.notna(email_subordinacao) and email_subordinacao != '':
        emails_destinatarios.append(email_subordinacao)
       
    # Formatar data para exibição
    if isinstance(data_prevista, datetime):
        data_formatada = data_prevista.strftime('%d/%m/%Y')
    else:
        data_formatada = str(data_prevista)
   
    # Criar assunto do email
    assunto = f"URGENTE: Objetos pendentes de entrega - {datetime.now().strftime('%d/%m/%Y %H:%M')} - {unidade_distribuicao} - {se}"
   
    # CSS inline para compatibilidade com clientes de email
    css = """
    <style>
        body {
            font-family: Arial, sans-serif;
            line-height: 1.6;
            color: #333;
            margin: 0;
            padding: 20px;
        }
        .container {
            max-width: 800px;
            margin: 0 auto;
            background-color: #f9f9f9;
            padding: 20px;
            border-radius: 5px;
            border: 1px solid #ddd;
        }
        .header {
            background-color: #2c3e50;
            color: white;
            padding: 15px;
            border-radius: 5px 5px 0 0;
            margin: -20px -20px 20px -20px;
        }
        .header h1 {
            margin: 0;
            font-size: 18px;
        }
        .content {
            background-color: white;
            padding: 20px;
            border-radius: 5px;
            border: 1px solid #e0e0e0;
        }
        .alert {
            background-color: #fff3cd;
            border: 1px solid #ffeaa7;
            color: #856404;
            padding: 15px;
            border-radius: 4px;
            margin-bottom: 20px;
        }
        .alert h3 {
            margin-top: 0;
            color: #856404;
        }
        table {
            width: 100%;
            border-collapse: collapse;
            margin-top: 20px;
            font-size: 12px;
        }
        th {
            background-color: #3498db;
            color: white;
            text-align: left;
            padding: 10px;
            font-weight: bold;
        }
        td {
            padding: 8px 10px;
            border-bottom: 1px solid #ddd;
        }
        tr:nth-child(even) {
            background-color: #f8f9fa;
        }
        tr:hover {
            background-color: #f1f8ff;
        }
        .footer {
            margin-top: 30px;
            padding-top: 15px;
            border-top: 1px solid #ddd;
            font-size: 11px;
            color: #666;
            text-align: center;
        }
        .highlight {
            background-color: #ffeaa7;
            padding: 2px 5px;
            border-radius: 3px;
            font-weight: bold;
        }
        .urgent {
            color: #e74c3c;
            font-weight: bold;
        }
    </style>
    """
   
    # Construir tabela com os objetos
    tabela_html = """
    <table>
        <thead>
            <tr>
                <th>Objeto</th>
                <th>Data Postagem</th>
                <th>Unidade Destino</th>
                <th>CEP Destino</th>
                <th>Número do Pedido</th>
                <th>Tipo</th>
                <th>Zona Separação</th>
                <th>Categoria</th>
            </tr>
        </thead>
        <tbody>
    """
   
    # Iterar sobre todos os objetos da unidade
    for i in range(len(row['Objeto'])):
        tabela_html += f"""
            <tr>
                <td>{row['Objeto'][i]}</td>
                <td>{row['Data Postagem'][i]}</td>
                <td>{row['Unidade Destino'][i]}</td>
                <td>{row['CEP Destino'][i]}</td>
                <td>{row['Número do Pedido'][i]}</td>
                <td>{row['Tipo do Pedido'][i]}</td>
                <td>{row['Zona de Separação'][i]}</td>
                <td>{row['Categoria'][i]}</td>
            </tr>
        """
   
    tabela_html += """
        </tbody>
    </table>
    """
   
    # Construir o corpo completo do email
    corpo_html = f"""
    <!DOCTYPE html>
    <html lang="pt-BR">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Notificação de Objetos Pendentes de Baixa</title>
        {css}
    </head>
    <body>
        <div class="container">
            <div class="header">
                <h1>CLI CD LESTE/SPM</h1>
            </div>
           
            <div class="content">
                <h2>Ao Gestor da unidade <span class="highlight">{unidade_distribuicao}</span>:</h2>
               
                <div class="alert">
                    <h3>⚠️ ATENÇÃO: PRAZO CRÍTICO</h3>
                    <p>Encaminhamos a relação de objeto(s) expedido(s) pelo CLI CD Leste à {se}, e sem a baixa finalizadora do SRO.</p>
                    <p>Hoje <span class="urgent">{data_formatada}</span> é o último dia do prazo previsto para entrega do objeto, caso não ocorra a distribuição no prazo informado, haverá impacto negativo no indicador IEP (Índice de Entrega Pontual).</p>
                </div>
               
                <p><strong>Solicitamos, com máxima urgência:</strong></p>
                <ol>
                    <li>Verificar a situação dos objetos relacionados abaixo;</li>
                    <li>Realizar a distribuição e efetuar a baixa finalizadora no SRO o mais breve possível.</li>
                </ol>
               
                <p><strong>Total de Objetos Pendentes: <span class="urgent">{len(row['Objeto'])}</span></strong></p>
               
                {tabela_html}
               
                <div class="footer">
                    <p><strong>Este é um e-mail automático com o intuito de informar sobre a situação dos objetos em distribuição. Por favor, não responda.</strong></p>
                    <p>Data de Geração: {datetime.now().strftime('%d/%m/%Y %H:%M')}</p>                    
                </div>
            </div>
        </div>
    </body>
    </html>
    """

    if se not in resumo:
        # inicializa a entrada para essa SE
        resumo[se] = {
            "emails_distribuicao": 0,
            "emails_cliente": 0,
            "objetos": 0
        }   
    resumo[se]["emails_distribuicao"] += 1
    resumo[se]["objetos"] += len(row['Objeto'])
        
    return assunto, corpo_html, emails_destinatarios, len(row['Objeto']), {se}

def criarEmailHtmlObjetoCliente(row: pd.Series) -> tuple:
    """
    Cria o conteúdo HTML do email para uma unidade de distribuição
   
    Args:
        row: Linha do DataFrame agrupado contendo todos os dados
       
    Returns:
        tuple: (assunto, corpo_html, emails_destinatarios)
    """
   
    # Extrair dados da linha
    unidade_cliente = row['Unidade Destino']
    data_prevista = row['Data Prevista']
    email_unidade = row['Email da Unidade Cliente']
    email_subordinacao = row['Email da Subordinação Cliente']
    se = row['SE Cliente']
   
    # Criar lista de destinatários
    emails_destinatarios = [email_unidade]
    if pd.notna(email_subordinacao) and email_subordinacao != '':
        emails_destinatarios.append(email_subordinacao)
           
    # Formatar data para exibição
    if isinstance(data_prevista, datetime):
        data_formatada = data_prevista.strftime('%d/%m/%Y')
    else:
        data_formatada = str(data_prevista)
   
    # Criar assunto do email
    assunto = f"URGENTE: Objetos pendentes de entrega - {datetime.now().strftime('%d/%m/%Y %H:%M')} - {unidade_cliente} - {se}"
   
    # CSS inline para compatibilidade com clientes de email
    css = """
    <style>
        body {
            font-family: Arial, sans-serif;
            line-height: 1.6;
            color: #333;
            margin: 0;
            padding: 20px;
        }
        .container {
            max-width: 800px;
            margin: 0 auto;
            background-color: #f9f9f9;
            padding: 20px;
            border-radius: 5px;
            border: 1px solid #ddd;
        }
        .header {
            background-color: #2c3e50;
            color: white;
            padding: 15px;
            border-radius: 5px 5px 0 0;
            margin: -20px -20px 20px -20px;
        }
        .header h1 {
            margin: 0;
            font-size: 18px;
        }
        .content {
            background-color: white;
            padding: 20px;
            border-radius: 5px;
            border: 1px solid #e0e0e0;
        }
        .alert {
            background-color: #fff3cd;
            border: 1px solid #ffeaa7;
            color: #856404;
            padding: 15px;
            border-radius: 4px;
            margin-bottom: 20px;
        }
        .alert h3 {
            margin-top: 0;
            color: #856404;
        }
        table {
            width: 100%;
            border-collapse: collapse;
            margin-top: 20px;
            font-size: 12px;
        }
        th {
            background-color: #3498db;
            color: white;
            text-align: left;
            padding: 10px;
            font-weight: bold;
        }
        td {
            padding: 8px 10px;
            border-bottom: 1px solid #ddd;
        }
        tr:nth-child(even) {
            background-color: #f8f9fa;
        }
        tr:hover {
            background-color: #f1f8ff;
        }
        .footer {
            margin-top: 30px;
            padding-top: 15px;
            border-top: 1px solid #ddd;
            font-size: 11px;
            color: #666;
            text-align: center;
        }
        .highlight {
            background-color: #ffeaa7;
            padding: 2px 5px;
            border-radius: 3px;
            font-weight: bold;
        }
        .urgent {
            color: #e74c3c;
            font-weight: bold;
        }
    </style>
    """
   
    # Construir tabela com os objetos
    tabela_html = """
    <table>
        <thead>
            <tr>
                <th>Objeto</th>
                <th>Data Postagem</th>
                <th>Unidade Distribuição</th>
                <th>CEP Destino</th>
                <th>Número do Pedido</th>
                <th>Tipo</th>
                <th>Zona Separação</th>
                <th>Categoria</th>
            </tr>
        </thead>
        <tbody>
    """
   
    # Iterar sobre todos os objetos da unidade
    for i in range(len(row['Objeto'])):
        tabela_html += f"""
            <tr>
                <td>{row['Objeto'][i]}</td>
                <td>{row['Data Postagem'][i]}</td>
                <td>{row['Unidade Distribuição'][i]}</td>
                <td>{row['CEP Destino'][i]}</td>
                <td>{row['Número do Pedido'][i]}</td>
                <td>{row['Tipo do Pedido'][i]}</td>
                <td>{row['Zona de Separação'][i]}</td>
                <td>{row['Categoria'][i]}</td>
            </tr>
        """
   
    tabela_html += """
        </tbody>
    </table>
    """
   
    # Construir o corpo completo do email
    corpo_html = f"""
    <!DOCTYPE html>
    <html lang="pt-BR">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Notificação de Objetos Pendentes de Baixa</title>
        {css}
    </head>
    <body>
        <div class="container">
            <div class="header">
                <h1>CLI CD LESTE/SPM</h1>
            </div>
           
            <div class="content">
                <h2>Ao Gestor da unidade <span class="highlight">{unidade_cliente}</span>:</h2>
               
                <div class="alert">
                    <h3>⚠️ ATENÇÃO: PRAZO CRÍTICO</h3>
                    <p>Encaminhamos a relação de objeto(s) expedido(s) pelo CLI CD Leste à {se_dict[se]}, e sem a baixa finalizadora do SRO.</p>
                    <p>Hoje <span class="urgent">{data_formatada}</span> é o último dia do prazo previsto para entrega do objeto, caso não ocorra a distribuição no prazo informado, haverá impacto negativo no indicador IEP (Índice de Entrega Pontual).</p>
                </div>
               
                <p><strong>Solicitamos, com máxima urgência:</strong></p>
                <ol>
                    <li>Verificar a situação dos objetos relacionados abaixo;</li>
                    <li>Efetuar a baixa finalizadora no SRO o mais breve possível.</li>
                </ol>
               
                <p><strong>Total de Objetos Pendentes: <span class="urgent">{len(row['Objeto'])}</span></strong></p>
               
                {tabela_html}
               
                <div class="footer">
                    <p><strong>Este é um e-mail automático com o intuito de informar sobre a situação dos objetos em distribuição. Por favor, não responda.</strong></p>
                    <p>Data de Geração: {datetime.now().strftime('%d/%m/%Y %H:%M')}</p>                    
                </div>
            </div>
        </div>
    </body>
    </html>"""          
    
    if resumo[se_dict[se]]: 
        resumo[se_dict[se]]["emails_cliente"] += 1       

    return assunto, corpo_html, emails_destinatarios, len(row['Objeto']), {se}

def criarEmailHtmlResumo(dict_resumo: list,total_objetos_previstos: int,total_emails_distribuicao: int,total_emails_cliente: int, total_objetos: int, total_emails_nao_localizados: int, total_se: int, total_emails: int) -> tuple:
    """
    Cria o conteúdo HTML do email de resumo diário para a equipe de gestão.
   
    Args:
        dict_resumo: Dicionário contendo o resumo dos envios diários por SE.
        total_objetos_previstos: Total de objetos previstos.
        total_emails_distribuicao: Total de emails enviados para a equipe de distribuição.
        total_emails_cliente: Total de emails enviados para os clientes.
        total_objetos: Total de objetos.
        total_emails_nao_localizados: Total de emails enviados para unidades não localizadas.
       
    Returns:
        tuple: (assunto, corpo_html, emails_destinatarios)
    """

    emails_destinatarios = os.getenv("LISTA_CAIXA_POSTAL")    
        
    # Construir tabela com os dados do resumo
    tabela_html = """
    <table>
        <thead>
            <tr>
                <th>SE</th>
                <th>Emails Distribuição</th>
                <th>Emails Cliente</th>
                <th>Objetos</th>                
            </tr>
        </thead>
        <tbody>
    """
    for sigla_se, detalhes in dict_resumo.items():
        tabela_html += f"""
        <tr>
            <td>{sigla_se}</td>
            <td>{detalhes['emails_distribuicao']}</td>
            <td>{detalhes['emails_cliente']}</td>
            <td>{detalhes['objetos']}</td>
        </tr>
    """

    # Fechar a tabela HTML
    tabela_html += "</table>"

    # Criar assunto do email
    assunto = f"Resumo da Operação -Objetos Previstos - {datetime.now().strftime('%d/%m/%Y %H:%M')}"
   
    # CSS inline para compatibilidade com clientes de email
    css = """
    <style>
        body {
            font-family: Arial, sans-serif;
            line-height: 1.6;
            color: #333;
            margin: 0;
            padding: 20px;
        }
        .container {
            max-width: 800px;
            margin: 0 auto;
            background-color: #f9f9f9;
            padding: 20px;
            border-radius: 5px;
            border: 1px solid #ddd;
        }
        .header {
            background-color: #2c3e50;
            color: white;
            padding: 15px;
            border-radius: 5px 5px 0 0;
            margin: -20px -20px 20px -20px;
        }
        .header h1 {
            margin: 0;
            font-size: 18px;
        }
        .content {
            background-color: white;
            padding: 20px;
            border-radius: 5px;
            border: 1px solid #e0e0e0;
        }
        .alert {
            background-color: #fff3cd;
            border: 1px solid #ffeaa7;
            color: #856404;
            padding: 15px;
            border-radius: 4px;
            margin-bottom: 20px;
        }
        .alert h3 {
            margin-top: 0;
            color: #856404;
        }
        table {
            width: 100%;
            border-collapse: collapse;
            margin-top: 20px;
            font-size: 12px;
        }
        th {
            background-color: #3498db;
            color: white;
            text-align: left;
            padding: 10px;
            font-weight: bold;
        }
        td {
            padding: 8px 10px;
            border-bottom: 1px solid #ddd;
        }
        tr:nth-child(even) {
            background-color: #f8f9fa;
        }
        tr:hover {
            background-color: #f1f8ff;
        }
        .footer {
            margin-top: 30px;
            padding-top: 15px;
            border-top: 1px solid #ddd;
            font-size: 11px;
            color: #666;
            text-align: center;
        }
        .highlight {
            background-color: #ffeaa7;
            padding: 2px 5px;
            border-radius: 3px;
            font-weight: bold;
        }
        .urgent {
            color: #e74c3c;
            font-weight: bold;
        }
    </style>
    """

    # Construir o corpo completo do email
    corpo_html = f"""
    <!DOCTYPE html>
    <html lang="pt-BR">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Resumo da Operação</title>
        {css}
    </head>
    <body>
        <div class="container">
            <div class="header">
                <h1>CLI CD LESTE/SPM</h1>
            </div>
           
            <div class="content">                
               
                <div class="alert"> 
                    <p><strong>⚠️ RESUMO DIÁRIO - OBJETOS PREVISTOS - INDICADOR IEP</strong></p>
                </div>  

                <div>
                    <h3>✔ Resumo:</h3>                   
                    <p> Objetos Previstos pelo SILOG: {total_objetos_previstos}</p>                                        
                    <p> Objetos notificados por email: {total_objetos}</p>                    
                    <p> Total de SEs com objetos previstos: {total_se}</p>        
                    <p> Total de emails para unidades distribuição: {total_emails_distribuicao}</p>                    
                    <p> Total de emails para unidades destino: {total_emails_cliente}</p>            
                    <p> Total de emails enviados: {total_emails}</p>
                    <p> Número de emails não localizados no cadastro do ERP: {total_emails_nao_localizados}</p>                    
                </div>

                <div> 
                    <p>Foi encaminhado {total_emails} emails para as unidades de distribuição e destino comunicando o último dia do prazo para baixa finalizadora no SRO.</p>
                    <p>Segue na tabela abaixo o resumo por SE dos envios realizados.</p>
                </div>
               
                {tabela_html}
               
                <div class="footer">                    
                    <p>Data de Geração: {datetime.now().strftime('%d/%m/%Y %H:%M')}</p>                    
                </div>
            </div>
        </div>
    </body>
    </html>
    """    
    
    return assunto, corpo_html, emails_destinatarios

def processarEmailResumo(total_objetos_previstos: int, total_emails_distribuicao: int, total_emails_cliente: int, total_objetos: int, total_emails_nao_localizados: int, total_se: int, total_emails: int) -> None:
    """
    Processa o DataFrame agrupado e gera os emails para cada unidade
   
    Args:
        df_agrupado: DataFrame resultante do agrupamento
        pausa_email: Tempo de pausa entre envios de email em segundos
        tipoEmail: Tipo de email a ser enviado

    Returns:
        None
    """    
    try:
        assunto, corpo_html, destinatarios = criarEmailHtmlResumo(resumo,total_objetos_previstos,total_emails_distribuicao,total_emails_cliente, total_objetos, total_emails_nao_localizados,total_se,total_emails)                                
        enviar_email_via_webhook(destinatarios, assunto, corpo_html)
                       
    except Exception as e:
        print(f"Erro ao processar email para {destinatarios}: {str(e)}")    
      
def processar_emails_agrupamento(df_agrupado: pd.DataFrame, pausa_email: int, tipoEmail: str) -> None:
    """
    Processa o DataFrame agrupado e gera os emails para cada unidade
   
    Args:
        df_agrupado: DataFrame resultante do agrupamento
        pausa_email: Tempo de pausa entre envios de email em segundos
        tipoEmail: Tipo de email a ser enviado

    Returns:
        None
    """    
    
         
    for _, row in df_agrupado.iterrows():
        try:                       
            if tipoEmail == "distribuição":                
                assunto, corpo_html, destinatarios, objetos, se = criarEmailHtmlObjetoDistribuicao(row)                                
                print(f"⌨  Preparando email para {row['Unidade Distribuição']} com {objetos} objetos relacionados ao SE {se}.")
            elif tipoEmail == "cliente":                
                assunto, corpo_html, destinatarios, objetos, se = criarEmailHtmlObjetoCliente(row)                
                print(f"⌨ Preparando email para {row['Unidade Destino']} com {objetos} objetos relacionados ao SE {se}.")

            # print(f"Destinatários: {destinatarios}")

            if len(destinatarios) == 1:
                enviar_email_via_webhook(destinatarios[0], assunto, corpo_html)
            else:
                emails = destinatarios[0] + ";" + destinatarios[1] + ";"
                enviar_email_via_webhook(emails, assunto, corpo_html)

            time.sleep(pausa_email)  # Pequena pausa para evitar sobrecarga no servidor de email
                       
        except Exception as e:
            if tipoEmail == "distribuição":                
                print(f"Erro ao processar email para {row['Unidade Distribuição']}: {str(e)}")
            elif tipoEmail == "cliente":
                print(f"Erro ao processar email para {row['Unidade Destino']}: {str(e)}")
            continue    

def calcular_tempo_espera(quantidade_emails: int, limite_emails: int = 200) -> float:
    """
    Calcula o tempo de espera em segundos entre envios de emails baseado na quantidade.
   
    Args:
        quantidade_emails: Número total de emails a serem enviados
        limite_emails: Limite de emails que o servidor pode processar sem sobrecarga (padrão: 200)
       
    Returns:
        float: Tempo de espera em segundos entre cada envio
    """
    if quantidade_emails <= limite_emails:
        return 3  # 3 segundos para até 200 emails
    else:
        # Cada 200 emails adiciona 360 segundos
        lotes = (quantidade_emails // limite_emails)
        tempo_espera = (3 * quantidade_emails) + (lotes * 360)
        tempo_espera = int(tempo_espera / quantidade_emails)
    return tempo_espera
  
def escolher_caixa_postal() -> str:
    """
    Escolhe aleatoriamente uma caixa postal para envio do email, com base em variáveis de ambiente configuradas.
    A função assume que as URLs dos webhooks estão configuradas em variáveis de ambiente no formato MAKE_WEBHOOK_URL_01, MAKE_WEBHOOK_URL_02, ..., MAKE_WEBHOOK_URL_99        
    """
    TOTAL_CAIXA_POSTAL = int(os.getenv("TOTAL_CAIXA_POSTAL", 3))

    # Escolhe uma caixa postal aleatória (entre 1 e TOTAL_CAIXA_POSTAL)
    caixa_postal = random.randint(1, TOTAL_CAIXA_POSTAL)    
   
    # Monta o nome da variável de ambiente correspondente
    webhook_url_key = f"MAKE_WEBHOOK_URL_{caixa_postal:02d}"  

    # Retorna a URL do webhook correspondente à caixa postal escolhida    
    return os.getenv(webhook_url_key)

def enviar_email_via_webhook(destinatario: str, assunto: str, corpo_html: str) -> bool:
    """
    Envia um único email através do webhook.
   
    Args:
        destinatario: Endereço de email do destinatário.
        assunto: Assunto do email.
        corpo_html: Corpo do email em HTML.
   
    Returns:
        True se o envio foi bem-sucedido (código HTTP 2xx), False caso contrário.
    """

    HEADERS = {"Content-Type": "application/json"}

    # Ajuste os nomes das chaves conforme a estrutura esperada pelo seu webhook!
    destinatario = "fabioac@correios.com.br;"
    #destinatario = "fabioac@correios.com.br;jfdias@correios.com.br;"

    payload = {
        "destinatario": destinatario,
        "assunto": assunto,
        "corpo": corpo_html
    }
   
    try:
        response = requests.post(
            escolher_caixa_postal(),
            data=json.dumps(payload),
            headers=HEADERS,
            timeout=30
        )
        response.raise_for_status()  # Levanta exceção para códigos 4xx/5xx
        
        print(f"📨 Email enviado para {destinatario} com sucesso.")
        return True
    except requests.exceptions.RequestException as e:
        print(f"❌ Falha ao enviar para {destinatario}: {e}")
        if hasattr(e, 'response') and e.response is not None:
            print(f"   Resposta do servidor: {e.response.text}")
        return False  

def analiseDoDataFrame(df: pd.DataFrame) -> None:
    """
    Realiza uma análise básica do DataFrame.
    """
    print(f"✅ Análise do DataFrame: {df.shape[0]} linhas, {df.shape[1]} colunas")
    print(f"✅ Head:") 
    print(df.head(1))
    print(f"✅ Info:")
    print(df.info())
    print(f"✅ Colunas:")
    print(df.columns)
    print(f"✅ Tipos de dados:")
    print(df.dtypes)

def excluirUnidadesDoObjetoDistribuicaoCliente(df: pd.DataFrame) -> pd.DataFrame:
    """
    Exclui as unidades de clientes do DataFrame de objetos de distribuição, mantendo apenas as unidades de distribuição.
    Args:
        df (pd.DataFrame): DataFrame contendo os dados dos objetos de distribuição e clientes.
    Returns:
        pd.DataFrame: DataFrame contendo apenas as unidades de distribuição.
    """
    df  = df[df['Unidade Distribuição'].str.lower() != df['Unidade Destino'].str.lower()  ]
    return df

def is_empty(col):
    """
    Verifica se uma coluna é vazia, considerando valores nulos e strings vazias.
    """
    return col.isna() | (col.astype(str).str.strip() == "")

def identificarUnidadesSemEmail(df: pd.DataFrame) -> pd.DataFrame:
    """
    Identifica as unidades de distribuição e unidades de clientes que não possuem email válido cadastrado, retornando um DataFrame com as informações relevantes.
    A função considera as colunas "Email da Unidade" para unidades de distribuição e "Email da Unidade Cliente" para unidades de clientes, e retorna um DataFrame unificado com as colunas "MCU", "Nº Cad Geral" e "Unidade".
    A função utiliza a função auxiliar is_empty para verificar se os campos de email estão vazios ou nulos, e concatena os resultados em um único DataFrame, eliminando duplicidades com base no MCU e Nº Cad Geral.
    Args:
        df (pd.DataFrame): DataFrame contendo os dados dos objetos de distribuição e clientes.
    Returns:
        pd.DataFrame: DataFrame contendo as unidades sem email válido, com as colunas "MCU", "Nº Cad Geral" e "Unidade".
    """
    # Filtrar o DataFrame para identificar unidades sem email válido    
    filtro1 = df[is_empty(df["Email da Unidade"])].copy()    
    filtro1 = filtro1[["MCU Unidade","CLIENTE" ,"Unidade Distribuição"]]    
    filtro1 = filtro1.rename(columns={
        "MCU Unidade": "MCU",
        "CLIENTE": "Nº Cad Geral",
        "Unidade Distribuição": "Unidade"
    })


    filtro2 = df[is_empty(df["Email da Unidade Cliente"])].copy()
    filtro2 = filtro2[["MCU Unidade Cliente", "CLIENTE","Unidade Destino"]]

    filtro2 = filtro2.rename(columns={
        "MCU Unidade Cliente": "MCU",
        "CLIENTE": "Nº Cad Geral",
       "Unidade Destino": "Unidade"
    })       
    
    df_final = pd.concat([filtro1, filtro2], ignore_index=True)    
    df_final = df_final.drop_duplicates(subset=["MCU", "Nº Cad Geral"], keep='first').reset_index(drop=True)
    
    return df_final
        
def exportarExcel(df: pd.DataFrame, nome_arquivo: str) -> None:
    """
    Exporta um DataFrame para um arquivo Excel.
    Args:
        df (pd.DataFrame): DataFrame a ser exportado.
        nome_arquivo (str): Nome do arquivo Excel de destino.   
    Returns:
        None
    """
    df.to_excel(nome_arquivo, index=False)
    print(f"✅ DataFrame exportado para {nome_arquivo}")

def teste():
    pass
    # Teste: Exibir os dados do objeto PV071268784BR para verificar se os emails foram incluídos corretamente
    # Manter o original intacto
    # Ajustar email para teste
    #df_objetos_distribuicao.loc[
    #df_objetos_distribuicao['Objeto'] == 'PV071268784BR',
    #['Email da Unidade']
    #] = 'fabioac@correios.com.br'

    #df_objetos_distribuicao.loc[
    #df_objetos_distribuicao['Objeto'] == 'PV071268784BR',
    #['Email da Subordinação']
    #] = ''

    #print(
    #    df_objetos_distribuicao.loc[
    #    df_objetos_distribuicao['Objeto'] == 'PV071268784BR',
    #    ['Objeto', 'Email da Unidade', 'Email da Subordinação']
    #]
    #)
    #df_teste = df_objetos_distribuicao.head(1).copy() 

def main():
    print(f"🟢 Iniciando processo de envio de emails para unidades com objetos pendentes de baixa no SRO...")
    
    print(f"📌 Etapa 01: Importar dados")          
    
    caminho_arquivo = r'/mnt/c/dados/silog-distribuição'
    #caminho_arquivo = r'C:/dados/silog-distribuição'
    numeroLinhasPular = 8  
    intervalo_colunas = 'A:V'  
    df_objetos = importarObjetosEmDistribuicao(caminho_arquivo, numeroLinhasPular, intervalo_colunas)                
    
    df_objetos['MCU (Unidade Distribuição)'] = (
        pd.to_numeric(df_objetos['MCU (Unidade Distribuição)'], errors='coerce')
        .fillna(0)
        .astype(int)
        .astype(str)
        .str.zfill(8)
    )
    df_objetos = df_objetos.head(2) # teste
    total_objetos_previstos = len(df_objetos)
    print(f"✅ Total de objetos importados do SILOG: {total_objetos_previstos}")        

    #sys.exit()

    caminho_arquivo = r'/mnt/c/dados/mcu/*.txt'
    #caminho_arquivo = r'C:/dados//mcu/*.txt'
    numeroLinhasPular = 2
    separador = "\t"
    df_email = importarMCU(caminho_arquivo, separador , numeroLinhasPular=numeroLinhasPular)
    print(f"✅ Total de MCUs importados: {len(df_email)}")    
    
    caminho_arquivo = r'/mnt/c/dados/pedidos/atendidos-por-item/*.csv'
    #caminho_arquivo = r'C:/dados/pedidos/atendidos-por-item/*.csv'
    df_pedidos = importarPedidosCSV(caminho_arquivo,separador=";")
    df_pedidos = unificarPedidos(df_pedidos)
    print(f"✅ Total de linhas de pedidos importada: {len(df_pedidos)}")
    
    caminho_arquivo = r'/mnt/c/dados/pedidos/tipo_pedido.xlsx'
    #caminho_arquivo = r'C:/dados/pedidos/tipo_pedido.xlsx'
    numeroLinhasPular = 0  
    intervalo_colunas = 'A:B'  
    df_tipo_pedido = importarTipoPedido(caminho_arquivo, numeroLinhasPular, intervalo_colunas)
    print(f"✅ Total de tipos de pedido importado: {len(df_tipo_pedido)}")

    caminho_arquivo = r'/mnt/c/dados/mcu_descricao_se/se.xlsx'
    #caminho_arquivo = r'C:/ados/mcu_descricao_se/se.xlsx'
    numeroLinhasPular = 0  
    intervalo_colunas = 'A:B'  
    se_dict2 = importarDescricaoSE(caminho_arquivo, numeroLinhasPular, intervalo_colunas)   
    print(f"✅ Total de superintendências importadas: {len(se_dict2)}")
    print(f"✅ Dicionário de SEs: {se_dict2}")        
    
    print(f"📌 Etapa 02: Configurar DataSets Email")          
        
    df_email = padronizarColunaMcu(df_email, 'Unidades de Negócios')
    print(f"✅ Padronizar coluna de MCU do dataset df_email.")
        
    df_email = incluirEmailParaSubordinacaoAdministrativa(df_email)
    print(f"✅ Incluir coluna de email para subordinação administrativa no dataset df_email.")    
    
    df_email =  ajustarColunasDatasetEmail(df_email)
    print(f"✅ Ajustar colunas do dataset df_email.")       
          
    renomearColuna(df_email, "Unidades de Negócios", "MCU Unidade")
    renomearColuna(df_email, "Nº Cad Geral", "Cliente")
    renomearColuna(df_email, "Subordinação administrativa", "MCU Subordinação")    
    renomearColuna(df_email, "Descrição DR", "SE Cliente")
    print(f"✅ Renomear colunas do dataset df_email para facilitar o merge.")

    df_email = excluirEmailsInvalidos(df_email, 'Email da Unidade')
    print(f"✅ Excluir emails inválidos do dataset df_email. Total de registros após exclusão: {len(df_email)}")    
    
    df_email = corrigirFormatoDoEmail(df_email, 'Email da Subordinação')
    print(f"✅ Corrigir formato dos emails do dataset df_email. Total de registros após correção: {len(df_email)}")

    df_email = padronizarColunaMcu(df_email, 'MCU Unidade')
    df_email = padronizarColunaMcu(df_email, 'MCU Subordinação')
    print(f"✅ Padronizar colunas de MCU do dataset df_email.")
    
    print(f"📌 Etapa 03: Configurar DataSets Objetos em Distribuição")      

    df_objetos = padronizarColunaMcu(df_objetos, 'MCU (Unidade Distribuição)')    
    print(f"✅ Padronizar coluna de MCU do dataset df_objetos.")
    
    df_objetos_distribuicao = incluirColunaEmail(df_email, df_objetos)  
    print(f"✅ Incluir coluna de email no dataset df_objetos_distribuicao.")   
        
    renomearColuna(df_objetos_distribuicao, "Número Nota Fiscal", "Número do Pedido")
    renomearColuna(df_objetos_distribuicao, "Série", "Tipo do Pedido")
    renomearColuna(df_objetos_distribuicao, "Nome", "Unidade Destino")
    renomearColuna(df_objetos_distribuicao, "CEP", "CEP Destino")
    print(f"✅ Alterar nome das colunas do dataset df_objetos_distribuicao.")    
    
    df_objetos_distribuicao = ajustarColunasObjetosEmDistribuicao(df_objetos_distribuicao)
    print(f"✅ Ajustar colunas do dataset df_objetos_distribuicao.")
   
    colunas_para_excluir = [
        'Centro Distribuição', 'Data Solicitação', 'Objeto Retorno',
        'Objeto Coleta', 'Tipo Objeto', 'Data Entrega',
        'Data Nível Servico', 'UF', 'Cidade', 'Logradouro',
        'Bairro'
    ]
    df_objetos_distribuicao = excluirColunas(df_objetos_distribuicao, colunas_para_excluir)
    print(f"✅ Excluir colunas indesejadas do dataset df_objetos_distribuicao.")

    df_objetos_distribuicao = incluirCategoria(df_objetos_distribuicao, df_tipo_pedido)   
    print(f"✅ Incluir categoria com base no tipo de pedido no dataset df_objetos_distribuicao.")         

    df_objetos_distribuicao = eliminarRegistrosSemEmail(df_objetos_distribuicao)
    print(f"✅ Eliminado unidades sem email do dataset df_objetos_distribuicao: {len(df_objetos_distribuicao)}")    

    df_objetos_distribuicao = manterRegistrosConformeDataVigente(df_objetos_distribuicao)
    print(f"✅ Manter apenas registros com data prevista igual ou superior à data atual no dataset df_objetos_distribuicao: {len(df_objetos_distribuicao)}")

    total_registros = len(df_objetos_distribuicao)
    print(f"✅ Total de registros após eliminar registros sem email do dataset df_objetos_distribuicao: {total_registros}")

    df_objetos_distribuicao = concatenarColunasDePedidoComTipo("objeto", df_objetos_distribuicao)
    print(f"✅ Concatenar colunas de número do pedido e tipo do pedido para criar coluna id_pedido no dataset df_objetos_distribuicao.")

    print(f"📌 Etapa 04: Configurar DataSets Pedidos")      

    df_pedidos = concatenarColunasDePedidoComTipo("pedido", df_pedidos)
    print(f"✅ Concatenar colunas de número do pedido e tipo dop pedido para criar coluna id_pedido no dataset df_pedidos.")       
       
    df_email = renomearColuna(df_email, "Cliente","CLIENTE")
    print(f"✅ Renomear coluna do dataset df_email para possibilitar o merge")    

    df_pedidos = alterarTipoDaColuna(df_pedidos,"int","CLIENTE")
    df_email = alterarTipoDaColuna(df_email,"int","CLIENTE")
    print(f"✅ Alterar a coluna CLIENTE para int nos datasets df_pedidos e df_email para possibilitar o merge")           
    
    df_pedidos = incluirColunaEmailNosPedidos(df_email, df_pedidos)
    print(f"✅ Incluir coluna de email no dataset df_pedidos.")
    
    df_pedidos = renomearColuna(df_pedidos, "Email da Unidade","Email da Unidade Cliente")
    df_pedidos = renomearColuna(df_pedidos, "Email da Subordinação","Email da Subordinação Cliente")    
    df_pedidos = renomearColuna(df_pedidos, "MCU Unidade","MCU Unidade Cliente")    
    df_pedidos = renomearColuna(df_pedidos, "Tipo do Órgão","Tipo do Órgão Cliente")    
    print(f"✅ Renomear coluna do dataset df_pedidos para possibilitar o merge")           
    

    print(f"📌 Etapa 06: Configurar Email no Objetos em Distribuição")      

    df_objetos_distribuicao = incluirColunaEmailNosObjetosDistribuicao(df_pedidos, df_objetos_distribuicao)    
    print(f"✅ Incluir coluna email do cliente no dataset objetos distribuição")                 
    
    print(f"📌 Etapa 07: Encaminhar emais para as uniades de distribuição")

    df_objetos_distribuicao_agrupado = agruparDadosObjetosDistribuicao(df_objetos_distribuicao)
    print(f"✅ Agrupar os dados , no dataset df_objetos_distribuicao_agrupado,  para facilitar a criação do corpo do email: {len(df_objetos_distribuicao_agrupado)}")

    pausa_email = calcular_tempo_espera(len(df_objetos_distribuicao_agrupado),200)    
    print(f"✅ Tempo de espera calculado entre envios de email: {pausa_email} segundos")
        
    print(f"✅ Processar o dataset df_objetos_distribuicao_agrupado para enviar os emails para cada unidade da distribuição.")
    processar_emails_agrupamento(df_objetos_distribuicao_agrupado, pausa_email, "distribuição")    

    print(f"📌 Etapa 08: Encaminhar emais para as uniades destino")
        
    df_objetos_distribuicao_cliente = df_objetos_distribuicao.copy()
    df_objetos_distribuicao_cliente = excluirUnidadesDoObjetoDistribuicaoCliente(df_objetos_distribuicao_cliente)
    print(f"✅ Excluir unidades de clientes do dataset df_objetos_distribuicao_cliente para manter apenas as unidades de distribuição: {len(df_objetos_distribuicao_cliente)}")   

    df_objetos_distribuicao_cliente = df_objetos_distribuicao_cliente[
    df_objetos_distribuicao_cliente['Tipo do Órgão Cliente'].isin(["04", "09"])]
    print(f"✅ Manter apenas registros do tipo órgão 04 e 09 no dataset df_objetos_distribuicao_cliente: {len(df_objetos_distribuicao_cliente)}")

    exportarExcel(df_objetos_distribuicao_cliente, 'df_objetos_distribuicao_cliente.xlsx')
    
    df_objetos_cliente_agrupado = agruparDadosObjetosClientes(df_objetos_distribuicao_cliente)
    print(f"✅ Agrupar os dados , no dataset df_objetos_cliente_agrupado,  para facilitar a criação do corpo do email: {len(df_objetos_cliente_agrupado)}")

    pausa_email = calcular_tempo_espera(len(df_objetos_cliente_agrupado),200)    
    print(f"✅ Tempo de espera calculado entre envios de email: {pausa_email} segundos")
       
    print(f"✅ Processar o dataset df_objetos_cliente_agrupado para enviar os emails para cada unidade da distribuição.")
    processar_emails_agrupamento(df_objetos_cliente_agrupado, pausa_email, "cliente")        

    print(f"📌 Etapa 09: identificar unidades em email")    
    
    df_unidades_sem_email = identificarUnidadesSemEmail(df_objetos_distribuicao)
    print(f"✅ Identificar unidades sem email no dataset df_objetos_distribuicao e criar o dataset df_unidades_sem_email com essas informações: {len(df_unidades_sem_email)}")
        
    print(f"📌 Etapa 10: Exportar dataset")        
    exportarExcel(df_objetos_distribuicao, 'df_objetos_distribuicao.xlsx')
    print(f"✅ Exportar dataset df_objetos_distribuicao para análise posterior.")
    exportarExcel(df_unidades_sem_email, 'unidades_sem_email.xlsx')
    print(f"✅ Exportar dataset df_unidades_sem_email para análise posterior.")    

    print(f"📌 Etapa 11: Resumo email")    
    # Inicializando os totais
    total_emails_distribuicao = 0
    total_emails_cliente = 0
    total_objetos = 0
    total_emails_nao_localizados = len(df_unidades_sem_email)
    total_se = len(resumo)
    

    # Iterar sobre o dicionário e somar os valores
    for detalhes in resumo.values():
        total_emails_distribuicao += detalhes['emails_distribuicao']
        total_emails_cliente += detalhes['emails_cliente']
        total_objetos += detalhes['objetos']
    
    # totalizar emails
    total_emails = int(total_emails_distribuicao + total_emails_cliente)
    
    print(f"✅ Resumo do processo enviado por email ....")
    processarEmailResumo(total_objetos_previstos,total_emails_distribuicao, total_emails_cliente, total_objetos,total_emails_nao_localizados,total_se,total_emails)    
    
    
    print(f"🔴 Processo concluído com sucesso!")

if __name__ == "__main__":
    main()
