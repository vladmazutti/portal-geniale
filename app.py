from flask import Flask, send_file
import requests
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from datetime import datetime

app = Flask(__name__)

# ====================================
# PÁGINA INICIAL
# ====================================

@app.route("/")
def home():

    return """
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
                Planilha de Pagamento
            </button>
        </a>

    </body>

    </html>
    """

# ====================================
# DOWNLOAD RELATÓRIO
# ====================================

@app.route("/baixar")
def baixar():

    # ====================================
    # CREDENCIAIS OMIE
    # ====================================

    app_key = "2543276123388"
    app_secret = "cd84271c41f00486e438191c09b49522"

    # ====================================
    # URL API OMIE
    # ====================================

    url = "https://app.omie.com.br/api/v1/geral/clientes/"

    linhas = []
    pagina = 1

    while True:

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

        response = requests.post(url, json=payload)
        dados = response.json()

        clientes = dados.get("clientes_cadastro", [])

        if not clientes:
            break

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

        print(f"Página {pagina} carregada...")
        pagina += 1

    df = pd.DataFrame(linhas)

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

    for row in dataframe_to_rows(df, index=False, header=False):
        ws.append(row)

    data_hoje = datetime.now().strftime("%d%m%y")
    arquivo_final = f"Pagamentos{data_hoje}.xlsx"

    wb.save(arquivo_final)

    return send_file(
        arquivo_final,
        as_attachment=True,
        download_name=arquivo_final
    )

# ====================================
# INICIAR SERVIDOR
# ====================================

app.run(debug=True)