# -*- coding: utf-8 -*-
"""
Este script processa uma planilha de dados, filtra informações relevantes
com base em regras de negócio predefinidas e envia e-mails de notificação
personalizados para os técnicos responsáveis usando o Microsoft Outlook.

O script é projetado para ser executado em um ambiente Windows com o Outlook
instalado e configurado.
"""

import logging
from collections import defaultdict
from typing import List, Dict, Tuple, Any

import pandas as pd
import win32com.client as win32
from pandas import DataFrame, Series

# --- Configurações Globais ---

# Caminho para a planilha de origem que será processada.
FILE_PATH = r"C:\Users\diego.brito\Downloads\robov1\saida.xlsx"

# E-mail a ser usado quando o técnico for "A DISTRIBUIR - SUSPENSIVA".
SPECIAL_RECIPIENT_EMAIL = "barbara.salatiel@esporte.gov.br"

# Configuração do sistema de logging.
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("relatorio_log.txt"),
        logging.StreamHandler()
    ]
)


def carregar_e_preparar_dados(file_path: str) -> DataFrame:
    """
    Carrega a planilha, limpa os nomes das colunas e seleciona as colunas necessárias.

    Args:
        file_path (str): O caminho para o arquivo Excel.

    Returns:
        DataFrame: Um DataFrame do Pandas contendo os dados preparados.
    """
    logging.info(f"Carregando planilha de: {file_path}")
    df = pd.read_excel(file_path, engine='openpyxl')

    df.columns = df.columns.str.strip()
    colunas_necessarias = [
        'Instrumento', 'Número Ajustes', 'Situação P.Trabalho', 'Situação TA', 'Número TA',
        'Aba Anexos', 'Data Esclarecimento', 'Resposta Esclarecimento', 'Técnico', 'e-mail do Técnico'
    ]

    # Verifica se todas as colunas necessárias existem.
    colunas_faltando = [col for col in colunas_necessarias if col not in df.columns]
    if colunas_faltando:
        raise ValueError(f"As seguintes colunas obrigatórias não foram encontradas: {colunas_faltando}")

    return df[colunas_necessarias].fillna("")


def aplicar_regras_de_negocio(df: DataFrame) -> DataFrame:
    """
    Aplica as regras de filtragem para identificar linhas que necessitam de atenção.

    Args:
        df (DataFrame): O DataFrame com os dados preparados.

    Returns:
        DataFrame: Um DataFrame filtrado contendo apenas os itens de interesse.
    """
    logging.info("Aplicando regras de negócio para filtrar dados...")
    df['Data Esclarecimento'] = pd.to_datetime(df['Data Esclarecimento'], format='%d/%m/%Y', errors='coerce')

    # Regras de filtragem: mantém a linha se QUALQUER uma das condições for verdadeira.
    df_filtered = df[
        (df['Situação P.Trabalho'] == "Em Análise (aguardando parecer)") |
        (df['Situação TA'].isin(["Cadastrada", "Em Análise"])) |
        (df['Resposta Esclarecimento'].str.upper() == "SIM")
    ].copy()

    return df_filtered


def limpar_valores_irrelevantes(row: Series) -> Series:
    """
    Limpa os dados de uma linha que não correspondem à regra que a acionou.
    Isso ajuda a focar a atenção no relatório final.

    Args:
        row (Series): Uma linha do DataFrame.

    Returns:
        Series: A linha com os valores irrelevantes limpos.
    """
    return pd.Series({
        'Instrumento': row['Instrumento'],
        'Número Ajustes': row['Número Ajustes'],
        'Situação P.Trabalho': row['Situação P.Trabalho'] if row['Situação P.Trabalho'] == "Em Análise (aguardando parecer)" else "",
        'Situação TA': row['Situação TA'] if row['Situação TA'] in ["Cadastrada", "Em Análise"] else "",
        'Número TA': row['Número TA'],
        'Aba Anexos': row['Aba Anexos'],
        'Data Esclarecimento': row['Data Esclarecimento'],
        'Resposta Esclarecimento': row['Resposta Esclarecimento'] if str(row['Resposta Esclarecimento']).upper() == "SIM" else "",
        'Técnico': row['Técnico'],
        'e-mail do Técnico': row['e-mail do Técnico']
    })


def agrupar_dados_por_tecnico(df: DataFrame) -> Dict[Tuple[str, str], List[List[Any]]]:
    """
    Agrupa os dados filtrados por técnico para preparar o envio de e-mails.

    Args:
        df (DataFrame): O DataFrame final e limpo.

    Returns:
        Dict: Um dicionário onde as chaves são tuplas (técnico, email) e os valores
              são listas de linhas de dados para o relatório.
    """
    logging.info("Agrupando dados por técnico para envio de e-mails.")
    grouped_data = defaultdict(list)
    for _, row in df.iterrows():
        key = (row['Técnico'], row['e-mail do Técnico'])
        data_list = [
            row['Técnico'], row['Instrumento'], row['Situação P.Trabalho'],
            row['Situação TA'], row['Resposta Esclarecimento'], row['Aba Anexos']
        ]
        grouped_data[key].append(data_list)
    return grouped_data


def gerar_tabela_html(data: List[List[Any]]) -> str:
    """
    Gera uma tabela HTML a partir de uma lista de dados para o corpo do e-mail.

    Args:
        data (List[List[Any]]): Lista de listas, onde cada lista interna é uma linha da tabela.

    Returns:
        str: Uma string contendo a tabela em formato HTML.
    """
    table_html = """
    <head>
      <style>
        table { width: 100%; border-collapse: collapse; font-family: Arial, sans-serif; }
        th, td { border: 1px solid #dddddd; padding: 8px; text-align: left; }
        th { background-color: #f2f2f2; font-weight: bold; }
        tr:nth-child(even) { background-color: #f9f9f9; }
      </style>
    </head>
    <body>
      <h2>Relatório de Pendências e Alertas</h2>
      <table>
        <tr>
          <th>Técnico</th>
          <th>Instrumento</th>
          <th>Situação P.Trabalho</th>
          <th>Situação TA</th>
          <th>Resposta Esclarecimento</th>
          <th>Aba Anexos</th>
        </tr>
    """
    for row_data in data:
        table_html += "<tr>"
        for item in row_data:
            # Garante que valores nulos ou NaT sejam exibidos como vazios.
            cleaned_item = "" if pd.isna(item) else item
            table_html += f"<td>{cleaned_item}</td>"
        table_html += "</tr>"

    table_html += "</table></body>"
    return table_html


def enviar_email_outlook(subject: str, body_html: str, recipient: str) -> None:
    """
    Cria e envia um e-mail usando o Microsoft Outlook.

    Args:
        subject (str): O assunto do e-mail.
        body_html (str): O corpo do e-mail em formato HTML.
        recipient (str): O endereço de e-mail do destinatário.
    """
    try:
        outlook = win32.Dispatch('Outlook.Application')
        mail = outlook.CreateItem(0)
        mail.Subject = subject
        mail.HTMLBody = body_html
        mail.To = recipient
        mail.Send()
        logging.info(f"E-mail enviado com sucesso para {recipient}.")
    except Exception as e:
        logging.error(f"Falha ao enviar e-mail para {recipient}: {e}")


def main() -> None:
    """
    Função principal que orquestra todo o processo de
    carregamento, filtragem e envio de e-mails.
    """
    try:
        df_original = carregar_e_preparar_dados(FILE_PATH)
        df_filtrado = aplicar_regras_de_negocio(df_original)

        if df_filtrado.empty:
            logging.info("Nenhum dado correspondeu aos critérios de filtragem. Nenhum e-mail a ser enviado.")
            return

        df_limpo = df_filtrado.apply(limpar_valores_irrelevantes, axis=1)
        
        # Remove linhas que possam ter ficado vazias após a limpeza.
        df_final = df_limpo.loc[(df_limpo.drop(columns=['Técnico', 'e-mail do Técnico']) != "").any(axis=1)]

        if df_final.empty:
            logging.info("Após a limpeza, não restaram dados relevantes. Nenhum e-mail a ser enviado.")
            return
            
        dados_agrupados = agrupar_dados_por_tecnico(df_final)

        for (tecnico, email), dados_relatorio in dados_agrupados.items():
            destinatario = email
            
            if tecnico == "A DISTRIBUIR - SUSPENSIVA":
                destinatario = SPECIAL_RECIPIENT_EMAIL
                logging.info(f"Técnico '{tecnico}' encontrado. E-mail será redirecionado para {destinatario}.")
            
            if not destinatario or pd.isna(destinatario):
                logging.warning(f"E-mail do técnico '{tecnico}' está vazio ou inválido. Pulando envio.")
                continue

            corpo_email = gerar_tabela_html(dados_relatorio)
            
            # Montagem do corpo completo do e-mail.
            html_completo = (
                f"<p><strong>Prezado(a) {tecnico},</strong></p>"
                f"<p>Segue abaixo o relatório de instrumentos sob sua responsabilidade que requerem atenção.</p>"
                f"{corpo_email}"
                "<br><p>Atenciosamente,<br><strong>Equipe de Automação</strong></p>"
            )

            enviar_email_outlook(
                subject=f"Relatório de Alertas e Pendências - {tecnico}",
                body_html=html_completo,
                recipient=destinatario
            )

        logging.info("Processo concluído com sucesso.")

    except Exception as e:
        logging.critical(f"Ocorreu um erro fatal no processo: {e}")


if __name__ == "__main__":
    main()
