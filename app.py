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

    status = "Disponível para download" if existe_planilha else "Ainda não gerada"
    status_cor = "#198754" if existe_planilha else "#dc3545"

    return f"""
    <html>
    <head>
        <title>Portal Geniale</title>
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
    </head>

    <body style="
        margin:0;
        font-family:Arial, sans-serif;
        background:linear-gradient(135deg, #0d6efd, #0b2c66);
        min-height:100vh;
        display:flex;
        align-items:center;
        justify-content:center;
    ">

        <div style="
            background:white;
            width:90%;
            max-width:520px;
            border-radius:22px;
            padding:42px 36px;
            box-shadow:0 20px 45px rgba(0,0,0,0.25);
            text-align:center;
        ">

            <div style="
                font-size:14px;
                letter-spacing:2px;
                color:#6c757d;
                font-weight:bold;
                margin-bottom:12px;
            ">
                GENIALE
            </div>

            <h1 style="
                margin:0;
                font-size:34px;
                color:#1f2937;
            ">
                Portal de Relatórios
            </h1>

            <p style="
                font-size:17px;
                color:#6c757d;
                margin-top:12px;
                margin-bottom:28px;
            ">
                Planilhas automatizadas integradas ao OMIE
            </p>

            <div style="
                background:#f8f9fa;
                border:1px solid #e9ecef;
                border-radius:14px;
                padding:16px;
                margin-bottom:28px;
            ">
                <div style="
                    font-size:13px;
                    color:#6c757d;
                    margin-bottom:6px;
                ">
                    Status da planilha
                </div>

                <div style="
                    font-size:18px;
                    color:{status_cor};
                    font-weight:bold;
                ">
                    {status}
                </div>
            </div>

            <a href="/baixar" style="text-decoration:none;">
                <button style="
                    width:100%;
                    font-size:20px;
                    padding:18px 24px;
                    background-color:#0d6efd;
                    color:white;
                    border:none;
                    border-radius:14px;
                    cursor:pointer;
                    font-weight:bold;
                    margin-bottom:14px;
                ">
                    Baixar Planilha
                </button>
            </a>

            <a href="/atualizar" style="text-decoration:none;">
                <button style="
                    width:100%;
                    font-size:17px;
                    padding:15px 24px;
                    background-color:#198754;
                    color:white;
                    border:none;
                    border-radius:14px;
                    cursor:pointer;
                    font-weight:bold;
                ">
                    Atualizar Agora
                </button>
            </a>

            <p style="
                font-size:12px;
                color:#adb5bd;
                margin-top:26px;
                margin-bottom:0;
            ">
                Atualização automática a cada 1 hora
            </p>

        </div>

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
    <html>
    <head>
        <title>Planilha Atualizada</title>
    </head>

    <body style="
        margin:0;
        font-family:Arial, sans-serif;
        background:linear-gradient(135deg, #198754, #0f5132);
        min-height:100vh;
        display:flex;
        align-items:center;
        justify-content:center;
    ">

        <div style="
            background:white;
            width:90%;
            max-width:480px;
            border-radius:22px;
            padding:42px 36px;
            box-shadow:0 20px 45px rgba(0,0,0,0.25);
            text-align:center;
        ">

            <h1 style="
                color:#198754;
                font-size:32px;
                margin-bottom:12px;
            ">
                Planilha atualizada com sucesso!
            </h1>

            <p style="
                color:#6c757d;
                font-size:16px;
                margin-bottom:28px;
            ">
                A versão mais recente já está disponível para download.
            </p>

            <a href="/" style="text-decoration:none;">
                <button style="
                    width:100%;
                    font-size:18px;
                    padding:16px 24px;
                    background-color:#0d6efd;
                    color:white;
                    border:none;
                    border-radius:14px;
                    cursor:pointer;
                    font-weight:bold;
                ">
                    Voltar ao Portal
                </button>
            </a>

        </div>

    </body>
    </html>
    """


# ====================================
# SERVIDOR LOCAL
# ====================================

if __name__ == "__main__":
    app.run(debug=True)