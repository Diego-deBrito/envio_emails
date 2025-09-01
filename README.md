# Gerador de Relatórios por E-mail via Outlook

![Python](https://img.shields.io/badge/Python-3.7+-blue.svg)
![Libraries](https://img.shields.io/badge/Libraries-Pandas%20%7C%20PyWin32-blue)
![Platform](https://img.shields.io/badge/Platform-Windows-blue)

##  Descrição do Projeto

Este projeto é um script de automação focado no pós-processamento de dados e na geração de notificações. Sua função é ler uma planilha de controle (gerada, por exemplo, por um robô de coleta de dados), aplicar um conjunto de regras de negócio para identificar itens que exigem atenção, e enviar relatórios personalizados por e-mail para os técnicos responsáveis, utilizando o **Microsoft Outlook**.

Ele serve como a "última milha" de um fluxo de automação, transformando dados brutos em ações concretas e comunicando as pendências de forma clara e direcionada.

##  Requisito Crítico: Windows e Outlook

> **Este script depende fundamentalmente da biblioteca `pywin32` para se comunicar com aplicações do Windows. Portanto, ele só funcionará em um sistema operacional Windows que tenha o Microsoft Outlook instalado, configurado e em execução.**

##  Funcionalidades Principais

- **Processamento de Planilhas:** Utiliza a biblioteca Pandas para carregar e manipular dados de arquivos Excel (`.xlsx`).
- **Motor de Regras de Negócio:** Filtra os dados com base em condições específicas (ex: status de um processo, resposta de um formulário) para isolar apenas os registros relevantes.
- **Agrupamento e Personalização:** Agrupa os itens filtrados por técnico responsável, garantindo que cada pessoa receba um relatório contendo apenas os seus próprios itens.
- **Geração de E-mails em HTML:** Cria e formata tabelas em HTML para apresentar os dados de forma clara e profissional no corpo do e-mail.
- **Integração Nativa com Outlook:** Conecta-se diretamente ao cliente de e-mail Outlook para enviar as notificações a partir da conta do usuário que executa o script.
- **Tratamento de Casos Especiais:** Lida com regras específicas, como redirecionar e-mails de um determinado "técnico" para um destinatário fixo.
- **Logging Detalhado:** Registra todas as etapas do processo, sucessos e falhas em um arquivo de log (`relatorio_log.txt`) para fácil auditoria e depuração.

##  Pré-requisitos

- **Sistema Operacional:** Windows
- **Software:** Microsoft Outlook (instalado e com uma conta de e-mail configurada)
- **Linguagem:** [Python 3.7](https://www.python.org/downloads/) ou superior

##  Instalação e Configuração

1.  **Clone o repositório:**
    ```bash
    git clone [https://github.com/seu-usuario/seu-repositorio.git](https://github.com/seu-usuario/seu-repositorio.git)
    cd seu-repositorio
    ```

2.  **Crie um ambiente virtual (recomendado):**
    ```bash
    python -m venv venv
    .\venv\Scripts\activate
    ```

3.  **Instale as dependências:**
    Crie um arquivo `requirements.txt` com o conteúdo abaixo:
    ```
    pandas
    openpyxl
    pywin32
    ```
    Em seguida, instale as bibliotecas:
    ```bash
    pip install -r requirements.txt
    ```

4.  **Configure o Script:**
    Abra o arquivo Python e ajuste as constantes no topo do código:
    ```python
    # Caminho para a planilha de origem que será processada.
    FILE_PATH = r"C:\caminho\completo\para\sua\planilha.xlsx"

    # E-mail a ser usado quando o técnico for "A DISTRIBUIR - SUSPENSIVA".
    SPECIAL_RECIPIENT_EMAIL = "email.especial@exemplo.com.br"
    ```

##  Como Executar

1.  **Prepare a Planilha de Entrada:**
    Certifique-se de que a planilha especificada em `FILE_PATH` exista e contenha as colunas necessárias pelo script (ex: `Instrumento`, `Situação P.Trabalho`, `Técnico`, `e-mail do Técnico`, etc.).

2.  **Garanta que o Outlook esteja aberto:**
    Para um funcionamento mais fluido, é recomendado que o Microsoft Outlook já esteja em execução no seu computador.

3.  **Execute o Script:**
    Abra o terminal na pasta do projeto e execute o script:
    ```bash
    python nome_do_script.py
    ```

O script irá processar a planilha e começar a enviar os e-mails através do seu Outlook. Acompanhe o progresso pelo console ou pelo arquivo `relatorio_log.txt`.

##  Lógica de Negócio Implementada

Este script envia uma notificação sobre um instrumento se **uma ou mais** das seguintes condições forem verdadeiras:
- A coluna `Situação P.Trabalho` tem o valor "Em Análise (aguardando parecer)".
- A coluna `Situação TA` tem o valor "Cadastrada" ou "Em Análise".
- A coluna `Resposta Esclarecimento` tem o valor "SIM".

O e-mail enviado destacará apenas as condições que foram atendidas para cada instrumento.
