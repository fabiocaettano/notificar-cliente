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

# Configurações
# Carregar variáveis do arquivo .env
load_dotenv()

# importar objetos em distribuição
def importar_objetos_em_distribuicao(caminho_arquivo, numeroLinhasPular, intervalo_colunas) -> pd.DataFrame:    
    df = pd.read_excel(caminho_arquivo, header=numeroLinhasPular, usecols=intervalo_colunas)
    # Remover possíveis linhas vazias que costumam vir em relatórios exportados
    df = df.dropna(how='all')    
    # Remove espaços extras no início e fim dos nomes das colunas
    df.columns = df.columns.str.strip()
    # Se houver espaços extras dentro das células de texto
    df = df.applymap(lambda x: x.strip() if isinstance(x, str) else x)
    return df

# Importar dados dos pedidos
def importar_pedidos_csv(caminho_arquivo, separador) -> pd.DataFrame:
    lista_arquivos = glob.glob(caminho_arquivo)
    dfs = []
    for arquivo in lista_arquivos:
        df = pd.read_csv(arquivo, sep=separador,low_memory=False)
        dfs.append(df)
    df_final = pd.concat(dfs, ignore_index=True)
    return df_final

# importar tabela MCU
def importar_mcu(caminho_arquivo, separador, numeroLinhasPular) -> pd.DataFrame:
    df_mcu = pd.read_csv(
        caminho_arquivo,
        sep=separador,
        skiprows=numeroLinhasPular
    )
    return df_mcu

# importar tipo de pedido
def importar_tipo_pedido(caminho_arquivo, numeroLinhasPular, intervalo_colunas) -> pd.DataFrame:    
    df = pd.read_excel(caminho_arquivo, header=numeroLinhasPular, usecols=intervalo_colunas)
    # Remover possíveis linhas vazias que costumam vir em relatórios exportados
    df = df.dropna(how='all')    
    # Remove espaços extras no início e fim dos nomes das colunas
    df.columns = df.columns.str.strip()
    # Se houver espaços extras dentro das células de texto
    df = df.applymap(lambda x: x.strip() if isinstance(x, str) else x)
    return df

def unificarPedidos(df_pedidos: pd.DataFrame) -> pd.DataFrame:
    # Manter apenas as três colunas desejadas
    df_filtrado = df_pedidos[['PED', 'TP_PED', 'CLIENTE']]

    # Remover linhas duplicadas
    df_filtrado = df_filtrado.drop_duplicates()
    return df_filtrado

def incluirColunaEmail(df_email: pd.DataFrame, df_objetos: pd.DataFrame) -> pd.DataFrame:
    # Garantir que as colunas de junção sejam strings para evitar erro de zeros à esquerda    
    df_objetos['MCU (Unidade Distribuição)'] = df_objetos['MCU (Unidade Distribuição)'].astype(str).str.zfill(8)
    df_email['MCU Unidade'] = df_email['MCU Unidade'].astype(str).str.zfill(8)
    df_email['MCU Subordinação'] = df_email['MCU Subordinação'].astype(str).str.zfill(8)

    # Realizar o Merge (PROCV)
    # Selecionamos apenas as colunas necessárias do df_email para não poluir o dataset
    colunas_interesse_email = ['MCU Unidade', 'Email da Unidade', 'Email da Subordinação']
   
    df_resultado = pd.merge(
        df_objetos,
        df_email[colunas_interesse_email],
        left_on='MCU (Unidade Distribuição)',
        right_on='MCU Unidade',
        how='left'
    )        
    return df_resultado

def alterarNomeDasColunasMCU(df_completo: pd.DataFrame) -> pd.DataFrame:
    # Renomear a coluna "Número Nota Fiscal" para "Número do Pedido"
    df_completo_ajustado = df_completo.rename(columns={'Número Nota Fiscal': 'Número do Pedido'})

    # Renomear a coluna "Série" para "Tipo Pedido"
    df_completo_ajustado = df_completo_ajustado.rename(columns={'Série': 'Tipo do Pedido'})

    # Renomear a coluna "Nome" para "Unidade Destino"
    df_completo_ajustado = df_completo_ajustado.rename(columns={'Nome': 'Unidade Destino'})

    # Renomear a coluna "CEP" para "CEP Destino"
    df_completo_ajustado = df_completo_ajustado.rename(columns={'CEP': 'CEP Destino'})

    return df_completo_ajustado


def alterarNomeDasColunasObjetosEmDistribuicao(df_completo: pd.DataFrame) -> pd.DataFrame:
    # Renomear a coluna "Número Nota Fiscal" para "Número do Pedido"
    df_completo_ajustado = df_completo.rename(columns={'Número Nota Fiscal': 'Número do Pedido'})

    # Renomear a coluna "Série" para "Tipo Pedido"
    df_completo_ajustado = df_completo_ajustado.rename(columns={'Série': 'Tipo do Pedido'})

    # Renomear a coluna "Nome" para "Unidade Destino"
    df_completo_ajustado = df_completo_ajustado.rename(columns={'Nome': 'Unidade Destino'})

    # Renomear a coluna "CEP" para "CEP Destino"
    df_completo_ajustado = df_completo_ajustado.rename(columns={'CEP': 'CEP Destino'})

    return df_completo_ajustado

def ajustarColunas(df_completo: pd.DataFrame) -> pd.DataFrame:
    # Extrair os primeiros dois caracteres para nova coluna "Zona de Separação"
    df_completo['Zona de Separação'] = df_completo['Tipo do Pedido'].str[2:]  # Remover os primeiros dois caracteres
    df_completo['Tipo do Pedido'] = df_completo['Tipo do Pedido'].str[:2]  # Manter apenas os dois primeiros caracteres
    return df_completo

def excluirColunas(df_completo: pd.DataFrame) -> pd.DataFrame:
    # Excluir as colunas indesejadas
    colunas_para_excluir = [
        'Centro Distribuição', 'Data Solicitação', 'Objeto Retorno',
        'Objeto Coleta', 'Tipo Objeto', 'Data Entrega',
        'Data Nível Servico', 'UF', 'Cidade', 'Logradouro',
        'Bairro'
    ]
   
    df_completo_atualizado = df_completo.drop(columns=colunas_para_excluir, errors='ignore')

    return df_completo_atualizado

def incluirCategoria(df_completo: pd.DataFrame, df_tipo_pedido: pd.DataFrame) -> pd.DataFrame:
    # Garantir que as colunas de junção sejam strings
    df_completo['Tipo do Pedido'] = df_completo['Tipo do Pedido'].astype(str)
    df_tipo_pedido['Tipo de Pedido'] = df_tipo_pedido['Tipo do Pedido'].astype(str)

    # Merge (similar a um PROCV) com base na coluna "Tipo de Pedido"
    df_completo = df_completo.merge(
        df_tipo_pedido[['Tipo do Pedido', 'Categoria']],
        on='Tipo do Pedido',
        how='left'
    )
   
    return df_completo

def eliminarRegistrosSemEmail(df_completo: pd.DataFrame) -> pd.DataFrame:
    # Filtrar o DataFrame para manter apenas os registros onde "Email da Unidade" não é nulo ou vazio
    df_completo_atualizado = df_completo[df_completo['Email da Unidade'].notna() & (df_completo['Email da Unidade'].str.strip() != '')]
    return df_completo_atualizado

def manterRegistrosConformeDataVigente(df_completo: pd.DataFrame) -> pd.DataFrame:
    # Converter a coluna "Data Prevista" para datetime, se ainda não estiver nesse formato
    #df_completo['Data Prevista'] = pd.to_datetime(df_completo['Data Prevista'], dayfirst=True, errors='coerce')
    df_completo['Data Prevista'] = pd.to_datetime(df_completo['Data Prevista'],format='%d/%m/%Y', errors='coerce')

    # Obter a data atual
    data_atual = pd.to_datetime(datetime.now().date())

    print(f"{data_atual}")

    # Filtrar o DataFrame para manter apenas os registros onde "Data Prevista" é igual à data atual
    df_completo_atualizado = df_completo[df_completo['Data Prevista'] >= data_atual]    

    return df_completo_atualizado

def agruparDados(df_completo: pd.DataFrame) -> pd.DataFrame:
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

def criar_email_html(row: pd.Series) -> tuple:
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
                <h1>CLI CD LESTE/SPM - Objetos Suprimento</h1>
            </div>
           
            <div class="content">
                <h2>Ao Gestor da <span class="highlight">{unidade_distribuicao}</span>:</h2>
               
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
   
    return assunto, corpo_html, emails_destinatarios

def processar_emails_agrupamento(df_agrupado: pd.DataFrame, pausa_email: int) -> None:
    """
    Processa o DataFrame agrupado e gera os emails para cada unidade
   
    Args:
        df_agrupado: DataFrame resultante do agrupamento
        pausa_email: Tempo de pausa entre envios de email em segundos
       
    Returns:
        list: Lista de dicionários com informações dos emails
    """
   
    emails_para_enviar = []
   
    for _, row in df_agrupado.iterrows():
        try:
            assunto, corpo_html, destinatarios = criar_email_html(row)
            enviar_email_via_webhook(destinatarios[0], assunto, corpo_html)  # Enviar para o email da unidade
            time.sleep(pausa_email)  # Pequena pausa para evitar sobrecarga no servidor de email
                       
        except Exception as e:
            print(f"Erro ao processar email para {row['Unidade Distribuição']}: {str(e)}")
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
   
    if quantidade_emails <= 0:
        return 0
   
    if quantidade_emails <= limite_emails:
        # Se quantidade <= limite: tempo = quantidade * 2 segundos
        tempo_espera_segundos = quantidade_emails * 1
    else:
        # Se quantidade > limite:
       
        # 1. Dividir total de emails pelo limite, resultado +1
        divisao = quantidade_emails / limite_emails
        resultado_divisao_mais_um = math.ceil(divisao) + 1
       
        # 2. Multiplicar por 60 = "total da pausa"
        total_pausa = resultado_divisao_mais_um * 60
       
        # 3. Multiplicar total de emails por 2 = "total de segundos acima do limite"
        total_segundos_acima_limite = quantidade_emails * 2
       
        # 4. Soma do "total de segundos acima do limite" + "total da pausa" = "tempo total de espera"
        tempo_total_espera = total_segundos_acima_limite + total_pausa
       
        # 5. Dividir "tempo total de espera" pelo número de emails = "time espera em segundos"
        tempo_espera_segundos = tempo_total_espera / quantidade_emails
   
    return math.ceil(tempo_espera_segundos)

def escolher_caixa_postal() -> str:
    """
    Escolhe uma URL de webhook aleatória com base no número total de caixas postais.
    Retorna a URL correspondente à caixa postal escolhida.
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

        numero_emails += 1

        print(f"✅ Email {numero_emails} enviado para {destinatario} com sucesso.")
        return True
    except requests.exceptions.RequestException as e:
        print(f"❌ Falha ao enviar para {destinatario}: {e}")
        if hasattr(e, 'response') and e.response is not None:
            print(f"   Resposta do servidor: {e.response.text}")
        return False  


def main():      
    # Importar dados dos objetos em distribuição
    caminho_arquivo = r'/mnt/c/dados/silog-previsto/objetos.xlsx'
    #caminho_arquivo = r'C:/dados/silog-distribuição/objetos.xlsx'
    numeroLinhasPular = 8  
    intervalo_colunas = 'A:V'  
    df_objetos = importar_objetos_em_distribuicao(caminho_arquivo, numeroLinhasPular, intervalo_colunas)    
    df_objetos['MCU (Unidade Distribuição)'] = df_objetos['MCU (Unidade Distribuição)'].apply(lambda x: f"{int(x):08d}")
    print(f"Total de Objetos importados: {len(df_objetos)}")    

    # Importar dados dos emails das unidades
    caminho_arquivo = r'/mnt/c/dados/mcu/R55001A_ECT0001.txt'
    #caminho_arquivo = r'C:/dados/informação-dos-orgãos/email.xlsx'
    numeroLinhasPular = 2  
    separador = "\t"
    df_email = importar_mcu(caminho_arquivo, separador , numeroLinhasPular=numeroLinhasPular)
    print(f"Total de registros de email importados: {len(df_email)}")
    
    # Importar dados dos pedidos
    caminho_arquivo = r'/mnt/c/dados/pedidos/atendidos-por-item/*.csv'
    df_pedidos = importar_pedidos_csv(caminho_arquivo, separador=";")
    df_pedidos_filtrado = unificarPedidos(df_pedidos)
    print(f"Total de pedidos importados: {len(df_pedidos_filtrado)}")

    # Importar dados dos objetos em distribuição
    caminho_arquivo = r'/mnt/c/dados/silog-previsto/tipo_pedido.xlsx'
    #caminho_arquivo = r'C:/dados/tipo-pedido/tipo_pedido.xlsx'
    numeroLinhasPular = 0  
    intervalo_colunas = 'A:B'  
    df_tipo_pedido = importar_tipo_pedido(caminho_arquivo, numeroLinhasPular, intervalo_colunas)
    print(f"Total de Objetos importados: {len(df_tipo_pedido)}")

    # Incluir coluna de email no DataFrame de objetos    
    df_completo = incluirColunaEmail(df_email, df_objetos)  
    print(f"Incluir coluna emaIL")

    # Alterar nome das colunas para melhor entendimento
    df_completo = alterarNomeDasColunasObjetosEmDistribuicao(df_completo)
    print(f"ALterar nome das colunas")

    # Ajustar colunas para separar tipo do pedido e zona de separação
    df_completo = ajustarColunas(df_completo)
    print(f"Ajustar colunas")

    # Excluir colunas indesejadas
    df_completo = excluirColunas(df_completo)
    print(f"exlcuir colunas")

    # Incluir categoria com base no tipo de pedido
    df_completo = incluirCategoria(df_completo, df_tipo_pedido)            


    # Eliminar registros sem email
    df_completo = eliminarRegistrosSemEmail(df_completo)
    print(f"Elminado unidades sem email: {len(df_completo)}")

    # Manter apenas registros com data prevista igual ou superior à data atual
    df_completo = manterRegistrosConformeDataVigente(df_completo)

    total_registros = len(df_completo)
    print(f"Total de registros após eliminar sem email: {total_registros}")

    # Manter o original intacto
    df_teste = df_completo.head(1).copy()    

    # Agrupar os dados para facilitar a criação do corpo do email
    df_agrupado = agruparDados(df_teste)

    # Calcular tempo de espera entre envios de email baseado na quantidade
    pausa_email = calcular_tempo_espera(len(df_agrupado),200)    
   
    # Processar o DataFrame agrupado para criar os emails
    #processar_emails_agrupamento(df_agrupado, pausa_email)    
     

if __name__ == "__main__":
    main()
