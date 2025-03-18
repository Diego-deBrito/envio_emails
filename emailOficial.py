import pandas as pd
import win32com.client as win32
from datetime import datetime, timedelta
from collections import defaultdict

# 📂 Caminho da planilha de origem
file_path = r"C:\Users\diego.brito\Downloads\robov1\saida.xlsx"

# 📌 Carregar a planilha
df = pd.read_excel(file_path, engine='openpyxl')

# 📌 Remover espaços extras nos nomes das colunas
df.columns = df.columns.str.strip()

# 📌 Selecionar colunas desejadas (incluindo "Aba Anexos")
colunas_necessarias = [
    'Instrumento', 'Número Ajustes', 'Situação P.Trabalho', 'Situação TA', 'Número TA',
    'Aba Anexos', 'Data Esclarecimento', 'Resposta Esclarecimento', 'Técnico', 'e-mail do Técnico'
]
df_filtered = df[colunas_necessarias].fillna("")

# 📌 Filtrar as colunas com base nas regras (sem intervalo da semana para "Aba Anexos")
df_filtered['Data Esclarecimento'] = pd.to_datetime(df_filtered['Data Esclarecimento'], format='%d/%m/%Y', errors='coerce')
df_filtered = df_filtered[
    (df_filtered['Situação P.Trabalho'] == "Em Análise (aguardando parecer)") |
    (df_filtered['Situação TA'].isin(["Cadastrada", "Em Análise"])) |
    (df_filtered['Resposta Esclarecimento'].str.upper() == "SIM")
]

# 📌 Função para limpar valores irrelevantes e deixar células vazias
def limpar_valores(row):
    return pd.Series({
        'Instrumento': row['Instrumento'],  # Sempre incluir o campo Instrumento
        'Número Ajustes': row['Número Ajustes'],
        'Situação P.Trabalho': row['Situação P.Trabalho'] if row['Situação P.Trabalho'] == "Em Análise (aguardando parecer)" else "",
        'Situação TA': row['Situação TA'] if row['Situação TA'] in ["Cadastrada", "Em Análise"] else "",
        'Número TA': row['Número TA'],
        'Aba Anexos': row['Aba Anexos'],  # Mantém a coluna "Aba Anexos"
        'Data Esclarecimento': row['Data Esclarecimento'],  # Mantém a data sem filtro de intervalo
        'Resposta Esclarecimento': row['Resposta Esclarecimento'] if row['Resposta Esclarecimento'].upper() == "SIM" else "",
        'Técnico': row['Técnico'],
        'e-mail do Técnico': row['e-mail do Técnico']
    })

df_filtered = df_filtered.apply(limpar_valores, axis=1)

# 📌 Remover linhas que ficaram completamente vazias (exceto Técnico e E-mail)
df_filtered = df_filtered[
    (df_filtered.drop(columns=['Técnico', 'e-mail do Técnico']) != "").any(axis=1)
]

# 📌 Verificar se há dados após a filtragem
if df_filtered.empty:
    print("⚠️ Nenhum dado encontrado após a filtragem. Processo interrompido.")
    exit()

# 📌 Função para enviar e-mails
def send_email(subject, body, recipient):
    try:
        outlook = win32.Dispatch('Outlook.Application')
        mail = outlook.CreateItem(0)
        mail.Subject = subject
        mail.HTMLBody = body
        mail.To = recipient
        mail.Send()
        print(f"📧 E-mail enviado para {recipient}")
    except Exception as e:
        print(f"⚠️ Erro ao enviar e-mail para {recipient}: {e}")

# 📌 Função para gerar tabela HTML no e-mail
def generate_email_table(data):
    if not data:
        return "<p>Não há dados para exibir.</p>"

    table_html = """
    <html>
      <head>
        <style>
          table { width: 100%; border-collapse: collapse; }
          th, td { border: 1px solid black; padding: 8px; text-align: left; }
          th { background-color: #f2f2f2; }
        </style>
      </head>
      <body>
        <h2>Relatório de Ajustes</h2>
        <table>
          <tr>
            <th>Técnico</th>
            <th>Instrumento</th>
            <th>Situação P.Trabalho</th>
            <th>Situação TA</th>
            <th>Resposta Esclarecimento</th>
            <th>Aba Anexos</th>  <!-- Adicionado "Aba Anexos" na tabela -->
          </tr>
    """
    for row in data:
        row = ["" if pd.isna(item) else item for item in row]  # Substitui NaN por ""
        table_html += "<tr>" + "".join(f"<td>{item}</td>" for item in row) + "</tr>"

    table_html += "</table></body></html>"
    return table_html

# 📌 Preparar dados para e-mail
grouped_data = defaultdict(list)
for _, row in df_filtered.iterrows():
    grouped_data[(row['Técnico'], row['e-mail do Técnico'])].append([
        row['Técnico'], row['Instrumento'], row['Situação P.Trabalho'], row['Situação TA'],
        row['Resposta Esclarecimento'], row['Aba Anexos']  # Incluído "Aba Anexos" nos dados do e-mail
    ])

# 📧 Enviar e-mails
for (técnico, email_do_tecnico), data in grouped_data.items():
    # Verifica se o técnico é "A DISTRIBUIR - SUSPENSIVA"
    if técnico == "A DISTRIBUIR - SUSPENSIVA":
        email_do_tecnico = "barbara.salatiel@esporte.gov.br"

    # Verifica se o e-mail do técnico está vazio
    if not email_do_tecnico or pd.isna(email_do_tecnico) or email_do_tecnico.strip() == "":
        print(f"⚠️ E-mail do técnico {técnico} está vazio ou inválido. Pulando...")
        continue

    # Gera o corpo do e-mail
    email_body = generate_email_table(data)

    # Envia o e-mail
    send_email(
        subject=f"Relatório de análises - {técnico}",
        body=(f"<p><strong>Prezado(a) {técnico},</strong></p>"
              f"<p>Segue abaixo o relatório de ajustes para os instrumentos sob sua responsabilidade.</p>"
              f"{email_body}"
              "<p>Atenciosamente,<br><strong>Equipe de Automação</strong></p>"
              "<p>🤖</p>"),
        recipient=email_do_tecnico  # Usa o e-mail do técnico diretamente
    )

print("Processo concluído.")