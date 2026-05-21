from flask import Flask, send_file
import requests
from openpyxl import load_workbook
from datetime import datetime
from apscheduler.schedulers.background import BackgroundScheduler
import time
import os

app = Flask(__name__)

ARQUIVO_PRONTO = "planilha_pronta.xlsx"

# ====================================
# GERAR PLANILHA
# ====================================

def gerar_planilha():

    print("Iniciando atualização da planilha...")

    app_key = os.getenv("OMIE_APP_KEY")
    app_secret = os.getenv("OMIE_APP_SECRET")

    if not app_key or not app_secret:
        raise Exception("Credenciais OMIE não configuradas nas variáveis do Railway.")

    linhas = []
    pagina = 1
    total_paginas = 1

    while pagina <= total_paginas:

        url = "https://app.omie.com.br/api/v1/geral/clientes/"

        payload = {
            "call": "ListarClientes",
            "app_key": app_key,
            "app_secret": app_secret,
            "param": [
                {
                    "pagina": pagina,
                    "registros_por_pagina": 500,
                    "apenas_importado_api": "N"
                }
            ]
        }

        response = requests.post(url, json=payload, timeout=60)

        if response.status_code == 429:
            print("Limite da API OMIE atingido. Aguardando...")
            time.sleep(5)
            continue

        response.raise_for_status()

        dados = response.json()

        total_paginas = dados.get("total_de_paginas", 1)
        clientes = dados.get("clientes_cadastro", [])

        print(f"Página {pagina} carregada...")

        for cliente in clientes:

            linha = [""] * 65
            dados_bancarios = cliente.get("dadosBancarios", {})

            linha[1] = cliente.get("cnpj_cpf") or "N/D"
            linha[2] = cliente.get("razao_social") or "N/D"
            linha[17] = cliente.get("website") or "N/D"
            linha[18] = dados_bancarios.get("codigo_banco") or "N/D"
            linha[19] = dados_bancarios.get("agencia") or "N/D"
            linha[20] = dados_bancarios.get("conta_corrente") or "N/D"
            linha[22] = dados_bancarios.get("nome_titular") or "N/D"
            linha[64] = dados_bancarios.get("cChavePix") or "N/D"

            linhas.append(linha)

        pagina += 1
        time.sleep(0.2)

    wb = load_workbook("modelo.xlsx")
    ws = wb["BASE OMIE"]

    ws.sheet_state = "veryHidden"
    ws.delete_rows(1, ws.max_row)

    cabecalhos = [""] * 65

    cabecalhos[1] = "CNPJ/CPF"
    cabecalhos[2] = "Razão Social"
    cabecalhos[17] = "Website"
    cabecalhos[18] = "Banco"
    cabecalhos[19] = "Agência"
    cabecalhos[20] = "Conta Corrente"
    cabecalhos[22] = "Nome Titular Conta"
    cabecalhos[64] = "Chave PIX"

    ws.append(cabecalhos)

    for linha in linhas:
        ws.append(linha)

    wb.save(ARQUIVO_PRONTO)

    print("Planilha atualizada com sucesso!")


# ====================================
# AGENDADOR AUTOMÁTICO
# ====================================

scheduler = BackgroundScheduler()

scheduler.add_job(
    gerar_planilha,
    "interval",
    hours=1
)

scheduler.start()


# ====================================
# PÁGINA INICIAL
# ====================================

@app.route("/")
def home():

    existe_planilha = os.path.exists(ARQUIVO_PRONTO)

    status = "Planilha disponível para download." if existe_planilha else "Planilha ainda não foi gerada."

    return f"""
    <html>

    <head>
        <title>Portal Geniale</title>
    </head>

    <body style="
        font-family:Arial;
        background-color:#f4f4f4;
        text-align:center;
        padding-top:100px;
    ">

        <h1 style="font-size:40px;">
            Portal Geniale
        </h1>

        <p style="font-size:20px;">
            Relatórios Automatizados
        </p>

        <p style="font-size:16px;color:#555;">
            {status}
        </p>

        <br><br>

        <a href="/baixar">
            <button style="
                font-size:24px;
                padding:20px 40px;
                background-color:#0d6efd;
                color:white;
                border:none;
                border-radius:10px;
                cursor:pointer;
            ">
                Baixar Planilha
            </button>
        </a>

        <br><br>

        <a href="/atualizar">
            <button style="
                font-size:20px;
                padding:15px 30px;
                background-color:#198754;
                color:white;
                border:none;
                border-radius:10px;
                cursor:pointer;
            ">
                Atualizar Agora
            </button>
        </a>

    </body>

    </html>
    """


# ====================================
# DOWNLOAD
# ====================================

@app.route("/baixar")
def baixar():

    if not os.path.exists(ARQUIVO_PRONTO):
        return """
        <h1>A planilha ainda não foi gerada.</h1>
        <p>Clique em "Atualizar Agora" para gerar a primeira versão.</p>
        <a href="/">Voltar</a>
        """

    data_hoje = datetime.now().strftime("%d%m%y")

    return send_file(
        ARQUIVO_PRONTO,
        as_attachment=True,
        download_name=f"Pagamentos{data_hoje}.xlsx"
    )


# ====================================
# ATUALIZAR MANUALMENTE
# ====================================

@app.route("/atualizar")
def atualizar():

    gerar_planilha()

    return """
    <h1>Planilha atualizada com sucesso!</h1>

    <br><br>

    <a href="/">
        Voltar
    </a>
    """


# ====================================
# SERVIDOR LOCAL
# ====================================

if __name__ == "__main__":
    app.run(debug=True)