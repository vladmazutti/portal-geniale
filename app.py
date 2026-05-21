from flask import Flask, send_file
import requests
from openpyxl import load_workbook
from datetime import datetime
from apscheduler.schedulers.background import BackgroundScheduler
import time
import os

app = Flask(__name__)

ARQUIVO_PRONTO = "planilha_pronta.xlsx"
ARQUIVO_ATUALIZACAO = "ultima_atualizacao.txt"


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

    agora = datetime.now().strftime("%d/%m/%Y %H:%M")
    with open(ARQUIVO_ATUALIZACAO, "w", encoding="utf-8") as f:
        f.write(agora)

    print("Planilha atualizada com sucesso!")


scheduler = BackgroundScheduler()
scheduler.add_job(gerar_planilha, "interval", hours=1)
scheduler.start()


@app.route("/")
def home():
    existe_planilha = os.path.exists(ARQUIVO_PRONTO)

    if os.path.exists(ARQUIVO_ATUALIZACAO):
        with open(ARQUIVO_ATUALIZACAO, "r", encoding="utf-8") as f:
            ultima_atualizacao = f.read()
    else:
        ultima_atualizacao = "Ainda não atualizada"

    status = "Disponível para download" if existe_planilha else "Ainda não gerada"
    status_cor = "#1f9d45" if existe_planilha else "#dc3545"

    return f"""
    <html>
    <head>
        <title>Portal Geniale</title>
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
    </head>

    <body style="
        margin:0;
        font-family:Arial, sans-serif;
        background:#f5f5f5;
        color:#1f1f1f;
    ">

        <div style="height:8px;background:#ff8a00;"></div>

        <main style="
            min-height:calc(100vh - 180px);
            padding:48px 20px;
            text-align:center;
        ">

            <img src="/static/logo.png" style="
                width:230px;
                max-width:80%;
                margin-bottom:30px;
            ">

            <h1 style="
                font-size:44px;
                margin:0;
                font-weight:800;
            ">
                Portal de <span style="color:#ff8a00;">Relatórios</span>
            </h1>

            <p style="
                font-size:19px;
                color:#555;
                margin-top:12px;
                margin-bottom:34px;
            ">
                Planilhas automatizadas integradas ao OMIE
            </p>

            <div style="
                width:90%;
                max-width:760px;
                margin:0 auto;
                background:white;
                border-radius:22px;
                padding:38px 34px;
                box-shadow:0 16px 40px rgba(0,0,0,0.12);
                border:1px solid #e8e8e8;
            ">

                <div style="
                    font-size:46px;
                    color:#ff8a00;
                    margin-bottom:14px;
                ">
                    📄
                </div>

                <div style="
                    font-size:13px;
                    color:#777;
                    letter-spacing:1.5px;
                    font-weight:bold;
                    text-transform:uppercase;
                ">
                    Status da planilha
                </div>

                <div style="
                    font-size:25px;
                    color:{status_cor};
                    font-weight:800;
                    margin-top:10px;
                ">
                    {status}
                </div>

                <p style="
                    color:#666;
                    font-size:15px;
                    margin-top:10px;
                    margin-bottom:32px;
                ">
                    Última atualização: <span style="color:#ff8a00;font-weight:bold;">{ultima_atualizacao}</span>
                </p>

                <a href="/baixar" style="text-decoration:none;">
                    <button style="
                        width:100%;
                        max-width:560px;
                        font-size:22px;
                        padding:22px 26px;
                        background:#1f1f1f;
                        color:white;
                        border:none;
                        border-radius:14px;
                        cursor:pointer;
                        font-weight:bold;
                        box-shadow:0 12px 24px rgba(0,0,0,0.22);
                        margin-bottom:18px;
                    ">
                        ⬇️ Baixar Planilha
                    </button>
                </a>

                <a href="/atualizar" style="text-decoration:none;">
                    <button style="
                        width:100%;
                        max-width:560px;
                        font-size:21px;
                        padding:20px 26px;
                        background:#ff8a00;
                        color:white;
                        border:none;
                        border-radius:14px;
                        cursor:pointer;
                        font-weight:bold;
                        box-shadow:0 12px 24px rgba(255,138,0,0.28);
                    ">
                        🔄 Atualizar Agora
                    </button>
                </a>

            </div>

            <div style="
                width:90%;
                max-width:760px;
                margin:26px auto 0;
                display:grid;
                grid-template-columns:repeat(auto-fit, minmax(180px, 1fr));
                background:white;
                border-radius:18px;
                border:1px solid #e8e8e8;
                overflow:hidden;
                box-shadow:0 10px 28px rgba(0,0,0,0.08);
            ">

                <div style="padding:24px;">
                    <div style="font-size:32px;color:#ff8a00;">⏱️</div>
                    <strong>Atualização Automática</strong>
                    <p style="color:#666;margin-bottom:0;">A cada 1 hora</p>
                </div>

                <div style="padding:24px;border-left:1px solid #eee;border-right:1px solid #eee;">
                    <div style="font-size:32px;color:#ff8a00;">🛡️</div>
                    <strong>Dados Seguros</strong>
                    <p style="color:#666;margin-bottom:0;">Integração direta com OMIE</p>
                </div>

                <div style="padding:24px;">
                    <div style="font-size:32px;color:#ff8a00;">⚡</div>
                    <strong>Rápido e Prático</strong>
                    <p style="color:#666;margin-bottom:0;">Download da última versão</p>
                </div>

            </div>

        </main>

        <footer style="
            background:#1f1f1f;
            color:white;
            padding:36px 20px;
            border-top:6px solid #ff8a00;
            text-align:center;
        ">
            <img src="/static/logo.png" style="
                width:150px;
                background:white;
                padding:8px;
                border-radius:12px;
                margin-bottom:16px;
            ">

            <p style="font-size:17px;font-weight:bold;margin:8px 0;">
                Geniale Promoção Eventos e Merchandising
            </p>

            <p style="color:#cfcfcf;margin:0;">
                Transformamos ideias em experiências que geram resultados.
            </p>
        </footer>

    </body>
    </html>
    """


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
        background:#f5f5f5;
        min-height:100vh;
        display:flex;
        align-items:center;
        justify-content:center;
        text-align:center;
    ">

        <div style="
            background:white;
            width:90%;
            max-width:500px;
            border-radius:22px;
            padding:42px 36px;
            box-shadow:0 18px 40px rgba(0,0,0,0.16);
            border:1px solid #e8e8e8;
        ">

            <img src="/static/logo.png" style="
                width:190px;
                max-width:80%;
                margin-bottom:22px;
            ">

            <h1 style="
                color:#ff8a00;
                font-size:32px;
                margin-bottom:12px;
            ">
                Planilha atualizada com sucesso!
            </h1>

            <p style="
                color:#666;
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
                    background:#1f1f1f;
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


if __name__ == "__main__":
    app.run(debug=True)