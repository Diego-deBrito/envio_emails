# 📌 Automação de Relatórios e Envio de E-mails

## 🛠 Sobre o Projeto
Este script automatiza o processo de **filtragem de dados** em um arquivo Excel, agrupando informações relevantes e enviando e-mails automaticamente para os responsáveis. Ele utiliza **Python**, **Pandas** e **Outlook** para criar relatórios estruturados e enviar notificações de acompanhamento.

---

## 🚀 Funcionalidades

### 🔹 1. Leitura e Filtragem de Dados
O código carrega um arquivo **Excel** e seleciona apenas as colunas necessárias, removendo espaços extras e preenchendo valores nulos:
```python
df = pd.read_excel(file_path, engine='openpyxl')
df.columns = df.columns.str.strip()
df_filtered = df[colunas_necessarias].fillna("")
```

### 🔹 2. Aplicação de Regras de Filtragem
Filtra os dados com base nas seguintes condições:
- Instrumentos **"Em Análise (aguardando parecer)"**
- Instrumentos com **Situação TA "Cadastrada" ou "Em Análise"**
- Instrumentos com **Resposta de Esclarecimento = "SIM"**

```python
df_filtered = df_filtered[
    (df_filtered['Situação P.Trabalho'] == "Em Análise (aguardando parecer)") |
    (df_filtered['Situação TA'].isin(["Cadastrada", "Em Análise"])) |
    (df_filtered['Resposta Esclarecimento'].str.upper() == "SIM")
]
```

### 🔹 3. Organização dos Dados e Limpeza
- Remove valores irrelevantes para evitar ruído na análise.
- Mantém sempre os campos **Instrumento, Técnico e E-mail**.

```python
def limpar_valores(row):
    return pd.Series({
        'Instrumento': row['Instrumento'],
        'Situação P.Trabalho': row['Situação P.Trabalho'] if row['Situação P.Trabalho'] == "Em Análise (aguardando parecer)" else "",
        'Situação TA': row['Situação TA'] if row['Situação TA'] in ["Cadastrada", "Em Análise"] else "",
        'Resposta Esclarecimento': row['Resposta Esclarecimento'] if row['Resposta Esclarecimento'].upper() == "SIM" else "",
        'Técnico': row['Técnico'],
        'e-mail do Técnico': row['e-mail do Técnico']
    })

df_filtered = df_filtered.apply(limpar_valores, axis=1)
```

### 🔹 4. Envio de E-mails Automático
Os e-mails são enviados usando a biblioteca **win32com.client**, permitindo a automação do **Outlook**:

```python
def send_email(subject, body, recipient):
    outlook = win32.Dispatch('Outlook.Application')
    mail = outlook.CreateItem(0)
    mail.Subject = subject
    mail.HTMLBody = body
    mail.To = recipient
    mail.Send()
```

Cada técnico recebe um relatório formatado em **HTML**, incluindo os instrumentos pendentes de análise.

---

## 🔧 Como Configurar e Rodar o Projeto

1️⃣ **Instale as dependências:**
```sh
pip install pandas openpyxl pywin32
```

2️⃣ **Configure seu Outlook** para estar aberto durante a execução.

3️⃣ **Execute o script:**
```sh
python script.py
```

---

## 📌 Autor
👤 **Diego Bruno Santos de Brito**

📧 Entre em contato: [seu_email@email.com](mailto:seu_email@email.com)

📝 _Projeto em constante evolução! Sugestões são bem-vindas!_ 🚀

