from flask import Flask, send_file, request, redirect, url_for
import requests
from openpyxl import load_workbook
from datetime import datetime
from zoneinfo import ZoneInfo
from apscheduler.schedulers.background import BackgroundScheduler
import time
import os

app = Flask(__name__)

FUSO_BRASIL = ZoneInfo("America/Sao_Paulo")

RELATORIOS = {
    "geniale": {
        "nome": "GENIALE",
        "titulo": "Planilha de Pagamento",
        "descricao": "Pagamentos operacionais Geniale",
        "modelo": "modelo.xlsx",
        "arquivo_pronto": "planilha_geniale.xlsx",
        "arquivo_atualizacao": "ultima_atualizacao_geniale.txt",
        "env_key": "OMIE_APP_KEY",
        "env_secret": "OMIE_APP_SECRET",
        "download_name": "Pagamentos_Geniale",
        "restrito": False,
    },
    "paniz": {
        "nome": "PANIZ",
        "titulo": "Planilha de Pagamento",
        "descricao": "Pagamentos operacionais Paniz",
        "modelo": "modelo.xlsx",
        "arquivo_pronto": "planilha_paniz.xlsx",
        "arquivo_atualizacao": "ultima_atualizacao_paniz.txt",
        "env_key": "OMIE_PANIZ_APP_KEY",
        "env_secret": "OMIE_PANIZ_APP_SECRET",
        "download_name": "Pagamentos_Paniz",
        "restrito": False,
    },
    "financeiro": {
        "nome": "FINANCEIRO",
        "titulo": "Consolidadora Financeiro",
        "descricao": "Planilha consolidada de acesso restrito",
        "modelo": "modelo_financeiro.xlsx",
        "arquivo_pronto": "planilha_financeiro.xlsx",
        "arquivo_atualizacao": "ultima_atualizacao_financeiro.txt",
        "env_key": "OMIE_APP_KEY",
        "env_secret": "OMIE_APP_SECRET",
        "download_name": "Consolidadora_Financeiro",
        "restrito": True,
    },
}


def agora_brasil():
    return datetime.now(FUSO_BRASIL)


def agora_formatado():
    return agora_brasil().strftime("%d/%m/%Y %H:%M")


def data_arquivo():
    return agora_brasil().strftime("%d%m%y")


def obter_ultima_atualizacao(config):
    arquivo = config["arquivo_atualizacao"]

    if os.path.exists(arquivo):
        with open(arquivo, "r", encoding="utf-8") as f:
            return f.read()

    return "Ainda não atualizada"


def status_relatorio(config):
    existe = os.path.exists(config["arquivo_pronto"])

    return {
        "texto": "Disponível para download" if existe else "Ainda não gerada",
        "cor": "#1f9d45" if existe else "#dc3545",
        "ultima": obter_ultima_atualizacao(config),
    }


def buscar_clientes_omie(app_key, app_secret):
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

    return linhas


def gerar_planilha(tipo):
    if tipo not in RELATORIOS:
        raise Exception("Relatório inválido.")

    config = RELATORIOS[tipo]

    print(f"Iniciando atualização da planilha {config['nome']}...")

    app_key = os.getenv(config["env_key"])
    app_secret = os.getenv(config["env_secret"])

    if not app_key or not app_secret:
        raise Exception(
            f"Credenciais OMIE não configuradas para {config['nome']}."
        )

    if not os.path.exists(config["modelo"]):
        raise Exception(
            f"Modelo não encontrado: {config['modelo']}"
        )

    linhas = buscar_clientes_omie(app_key, app_secret)

    wb = load_workbook(config["modelo"])

    if "BASE OMIE" not in wb.sheetnames:
        raise Exception(
            f"Aba 'BASE OMIE' não encontrada em {config['modelo']}"
        )

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

    wb.save(config["arquivo_pronto"])

    with open(config["arquivo_atualizacao"], "w", encoding="utf-8") as f:
        f.write(agora_formatado())

    print(f"Planilha {config['nome']} atualizada com sucesso!")


scheduler = BackgroundScheduler(
    timezone="America/Sao_Paulo"
)

# GENIALE — 08:00
scheduler.add_job(
    lambda: gerar_planilha("geniale"),
    trigger="cron",
    hour=8,
    minute=0
)

# PANIZ — 08:00
scheduler.add_job(
    lambda: gerar_planilha("paniz"),
    trigger="cron",
    hour=8,
    minute=0
)

# FINANCEIRO — 08:00 e 12:00
scheduler.add_job(
    lambda: gerar_planilha("financeiro"),
    trigger="cron",
    hour="8,12",
    minute=0
)

scheduler.start()


def card_relatorio(tipo, config):
    status = status_relatorio(config)

    if tipo == "paniz":
        logo = "/static/logo_paniz.png"
    else:
        logo = "/static/logo.png"

    if config["restrito"]:
        botao_principal = f"""
            <a href="/acessar/{tipo}" style="text-decoration:none;">
                <button class="btn btn-dark">
                    🔒 Acessar
                </button>
            </a>
        """
    else:
        botao_principal = f"""
            <a href="/baixar/{tipo}" style="text-decoration:none;">
                <button class="btn btn-dark">
                    ⬇️ Baixar
                </button>
            </a>
        """

    return f"""
        <div class="card">

            <img src="{logo}" class="card-logo">

            <div class="card-tag">
                {config['nome']}
            </div>

            <h2>
                {config['titulo']}
            </h2>

            <p class="desc">
                {config['descricao']}
            </p>

            <div class="status" style="color:{status['cor']};">
                {status['texto']}
            </div>

            <p class="ultima">
                Última atualização:<br>
                <span>{status['ultima']}</span>
            </p>

            <div class="actions">

                {botao_principal}

                <a href="/atualizar/{tipo}" style="text-decoration:none;">
                    <button class="btn btn-orange">
                        🔄 Atualizar
                    </button>
                </a>

            </div>

        </div>
    """


@app.route("/")
def home():

    cards = "".join(
        card_relatorio(tipo, config)
        for tipo, config in RELATORIOS.items()
    )

    return f"""
    <html>

    <head>

        <title>Portal Corporativo</title>

        <meta
            name="viewport"
            content="width=device-width, initial-scale=1.0"
        >

        <style>

            * {{
                box-sizing:border-box;
            }}

            body {{
                margin:0;
                font-family:Arial, sans-serif;
                background:#f5f5f5;
                color:#1f1f1f;
            }}

            .topbar {{
                height:8px;
                background:#ff8a00;
            }}

            main {{
                min-height:calc(100vh - 180px);
                padding:48px 20px;
                text-align:center;
            }}

            .logos {{
                display:flex;
                justify-content:center;
                align-items:center;
                gap:28px;
                flex-wrap:wrap;
                margin-bottom:26px;
            }}

            .logos img {{
                height:70px;
                max-width:240px;
                object-fit:contain;
            }}

            h1 {{
                font-size:44px;
                margin:0;
                font-weight:800;
            }}

            h1 span {{
                color:#ff8a00;
            }}

            .subtitle {{
                font-size:19px;
                color:#555;
                margin-top:12px;
                margin-bottom:38px;
            }}

            .grid {{
                width:94%;
                max-width:1180px;
                margin:0 auto;
                display:grid;
                grid-template-columns:repeat(auto-fit, minmax(300px, 1fr));
                gap:24px;
            }}

            .card {{
                background:white;
                border-radius:22px;
                padding:34px 28px;
                box-shadow:0 16px 40px rgba(0,0,0,0.12);
                border:1px solid #e8e8e8;
            }}

            .card-logo {{
                height:54px;
                object-fit:contain;
                margin-bottom:18px;
            }}

            .card-tag {{
                font-size:13px;
                color:#777;
                letter-spacing:1.5px;
                font-weight:bold;
                text-transform:uppercase;
                margin-bottom:10px;
            }}

            .card h2 {{
                font-size:25px;
                margin:0 0 12px;
                font-weight:800;
            }}

            .desc {{
                color:#666;
                font-size:15px;
                min-height:42px;
                margin-bottom:22px;
            }}

            .status {{
                font-size:20px;
                font-weight:800;
                margin-top:12px;
            }}

            .ultima {{
                color:#666;
                font-size:14px;
                margin-top:10px;
                margin-bottom:24px;
                line-height:1.5;
            }}

            .ultima span {{
                color:#ff8a00;
                font-weight:bold;
            }}

            .actions {{
                display:grid;
                gap:12px;
            }}

            .btn {{
                width:100%;
                font-size:18px;
                padding:17px 22px;
                color:white;
                border:none;
                border-radius:14px;
                cursor:pointer;
                font-weight:bold;
            }}

            .btn-dark {{
                background:#1f1f1f;
                box-shadow:0 12px 24px rgba(0,0,0,0.22);
            }}

            .btn-orange {{
                background:#ff8a00;
                box-shadow:0 12px 24px rgba(255,138,0,0.28);
            }}

            .info {{
                width:94%;
                max-width:1180px;
                margin:30px auto 0;
                display:grid;
                grid-template-columns:repeat(auto-fit, minmax(220px, 1fr));
                background:white;
                border-radius:18px;
                border:1px solid #e8e8e8;
                overflow:hidden;
                box-shadow:0 10px 28px rgba(0,0,0,0.08);
            }}

            .info-item {{
                padding:24px;
                border-right:1px solid #eee;
            }}

            .info-item:last-child {{
                border-right:none;
            }}

            .info-icon {{
                font-size:32px;
                color:#ff8a00;
            }}

            .info p {{
                color:#666;
                margin-bottom:0;
            }}

            footer {{
                background:#1f1f1f;
                color:white;
                padding:36px 20px;
                border-top:6px solid #ff8a00;
                text-align:center;
                margin-top:50px;
            }}

            .footer-logos {{
                display:flex;
                justify-content:center;
                align-items:center;
                gap:24px;
                flex-wrap:wrap;
                margin-bottom:18px;
            }}

            .footer-logos img {{
                height:52px;
                background:white;
                padding:8px;
                border-radius:12px;
            }}

            footer .name {{
                font-size:17px;
                font-weight:bold;
                margin:8px 0;
            }}

            footer .text {{
                color:#cfcfcf;
                margin:0;
            }}

        </style>

    </head>

    <body>

        <div class="topbar"></div>

        <main>

            <div class="logos">
                <img src="/static/logo.png">
                <img src="/static/logo_paniz.png">
            </div>

            <h1>
                Portal de <span>Relatórios</span>
            </h1>

            <p class="subtitle">
                Geniale | Paniz | Financeiro
            </p>

            <div class="grid">
                {cards}
            </div>

            <div class="info">

                <div class="info-item">
                    <div class="info-icon">⏱️</div>

                    <strong>
                        Atualização Inteligente
                    </strong>

                    <p>
                        Horários programados
                    </p>
                </div>

                <div class="info-item">
                    <div class="info-icon">🛡️</div>

                    <strong>
                        Dados Seguros
                    </strong>

                    <p>
                        Integração protegida no Railway
                    </p>
                </div>

                <div class="info-item">
                    <div class="info-icon">⚡</div>

                    <strong>
                        Download Rápido
                    </strong>

                    <p>
                        Arquivos prontos para baixar
                    </p>
                </div>

            </div>

        </main>

        <footer>

            <div class="footer-logos">
                <img src="/static/logo.png">
                <img src="/static/logo_paniz.png">
            </div>

            <p class="name">
                Portal Corporativo de Relatórios
            </p>

            <p class="text">
                Geniale | Paniz | Financeiro
            </p>

        </footer>

    </body>

    </html>
    """


@app.route("/baixar/<tipo>")
def baixar(tipo):

    if tipo not in RELATORIOS:
        return "Relatório inválido.", 404

    config = RELATORIOS[tipo]

    if config["restrito"]:
        return redirect(
            url_for("acessar", tipo=tipo)
        )

    if not os.path.exists(config["arquivo_pronto"]):
        return pagina_erro(
            "A planilha ainda não foi gerada.",
            "Clique em Atualizar para gerar a primeira versão."
        )

    return send_file(
        config["arquivo_pronto"],
        as_attachment=True,
        download_name=(
            f"{config['download_name']}_{data_arquivo()}.xlsx"
        )
    )


@app.route("/acessar/<tipo>", methods=["GET", "POST"])
def acessar(tipo):

    if tipo not in RELATORIOS:
        return "Relatório inválido.", 404

    config = RELATORIOS[tipo]

    erro = ""

    if request.method == "POST":

        senha_digitada = request.form.get("senha", "")
        senha_correta = os.getenv("FINANCEIRO_PASSWORD")

        if senha_digitada == senha_correta:

            if not os.path.exists(config["arquivo_pronto"]):
                return pagina_erro(
                    "A planilha ainda não foi gerada.",
                    "Clique em Atualizar para gerar a primeira versão."
                )

            return send_file(
                config["arquivo_pronto"],
                as_attachment=True,
                download_name=(
                    f"{config['download_name']}_{data_arquivo()}.xlsx"
                )
            )

        else:
            erro = "Senha incorreta."

    return f"""
    <html>

    <head>

        <title>Acesso Restrito</title>

        <meta
            name="viewport"
            content="width=device-width, initial-scale=1.0"
        >

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
            max-width:460px;
            border-radius:22px;
            padding:42px 36px;
            box-shadow:0 18px 40px rgba(0,0,0,0.16);
            border:1px solid #e8e8e8;
        ">

            <img
                src="/static/logo.png"
                style="
                    width:180px;
                    max-width:80%;
                    margin-bottom:22px;
                "
            >

            <h1 style="
                color:#ff8a00;
                font-size:32px;
                margin-bottom:10px;
            ">
                Acesso Restrito
            </h1>

            <p style="
                color:#666;
                font-size:16px;
                margin-bottom:24px;
            ">
                Informe a senha para baixar a consolidadora financeira.
            </p>

            <form method="POST">

                <input
                    type="password"
                    name="senha"
                    placeholder="Digite a senha"

                    style="
                        width:100%;
                        font-size:18px;
                        padding:16px;
                        border-radius:12px;
                        border:1px solid #ddd;
                        margin-bottom:16px;
                        text-align:center;
                    "
                >

                <button
                    type="submit"

                    style="
                        width:100%;
                        font-size:18px;
                        padding:16px 24px;
                        background:#1f1f1f;
                        color:white;
                        border:none;
                        border-radius:14px;
                        cursor:pointer;
                        font-weight:bold;
                    "
                >
                    Baixar Planilha
                </button>

            </form>

            <p style="
                color:#dc3545;
                font-weight:bold;
                margin-top:18px;
            ">
                {erro}
            </p>

            <a
                href="/"

                style="
                    color:#ff8a00;
                    font-weight:bold;
                    text-decoration:none;
                "
            >
                Voltar ao Portal
            </a>

        </div>

    </body>

    </html>
    """


@app.route("/atualizar/<tipo>")
def atualizar(tipo):

    if tipo not in RELATORIOS:
        return "Relatório inválido.", 404

    try:
        gerar_planilha(tipo)

    except Exception as erro:

        return pagina_erro(
            "Erro ao atualizar a planilha.",
            str(erro)
        )

    config = RELATORIOS[tipo]

    return f"""
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

            <img
                src="/static/logo.png"

                style="
                    width:190px;
                    max-width:80%;
                    margin-bottom:22px;
                "
            >

            <h1 style="
                color:#ff8a00;
                font-size:32px;
                margin-bottom:12px;
            ">
                {config['nome']} atualizada com sucesso!
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


def pagina_erro(titulo, mensagem):

    return f"""
    <html>

    <head>
        <title>Erro</title>
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
            max-width:520px;
            border-radius:22px;
            padding:42px 36px;
            box-shadow:0 18px 40px rgba(0,0,0,0.16);
            border:1px solid #e8e8e8;
        ">

            <h1 style="
                color:#dc3545;
                font-size:30px;
                margin-bottom:12px;
            ">
                {titulo}
            </h1>

            <p style="
                color:#666;
                font-size:16px;
                margin-bottom:28px;
            ">
                {mensagem}
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