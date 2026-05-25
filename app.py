from flask import Flask, send_file, request, redirect, url_for, render_template
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
        "logo": "logo.png",
        "classe": "geniale",
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
        "logo": "logo_paniz.png",
        "classe": "paniz",
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
        "logo": None,
        "classe": "financeiro",
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
        "cor": "ok" if existe else "erro",
        "existe": existe,
        "ultima": obter_ultima_atualizacao(config),
    }


def montar_relatorios_para_tela():
    lista = []

    for tipo, config in RELATORIOS.items():
        item = dict(config)
        item["tipo"] = tipo
        item["status"] = status_relatorio(config)
        lista.append(item)

    return lista


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
                    "apenas_importado_api": "N",
                }
            ],
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
            f"Credenciais OMIE não configuradas para {config['nome']}. "
            f"Verifique {config['env_key']} e {config['env_secret']} no Railway."
        )

    if not os.path.exists(config["modelo"]):
        raise Exception(f"Modelo não encontrado: {config['modelo']}")

    linhas = buscar_clientes_omie(app_key, app_secret)

    wb = load_workbook(config["modelo"])

    if "BASE OMIE" not in wb.sheetnames:
        raise Exception(f"Aba 'BASE OMIE' não encontrada em {config['modelo']}")

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


scheduler = BackgroundScheduler(timezone="America/Sao_Paulo")

scheduler.add_job(
    lambda: gerar_planilha("geniale"),
    trigger="cron",
    hour=8,
    minute=0,
)

scheduler.add_job(
    lambda: gerar_planilha("paniz"),
    trigger="cron",
    hour=8,
    minute=0,
)

scheduler.add_job(
    lambda: gerar_planilha("financeiro"),
    trigger="cron",
    hour="8,12",
    minute=0,
)

scheduler.start()


@app.route("/")
def home():
    return render_template(
        "index.html",
        relatorios=montar_relatorios_para_tela(),
        agora=agora_formatado(),
    )


@app.route("/baixar/<tipo>")
def baixar(tipo):
    if tipo not in RELATORIOS:
        return render_template(
            "erro.html",
            titulo="Relatório inválido",
            mensagem="O relatório solicitado não existe.",
        ), 404

    config = RELATORIOS[tipo]

    if config["restrito"]:
        return redirect(url_for("acessar", tipo=tipo))

    if not os.path.exists(config["arquivo_pronto"]):
        return render_template(
            "erro.html",
            titulo="A planilha ainda não foi gerada.",
            mensagem="Clique em Atualizar para gerar a primeira versão.",
        )

    return send_file(
        config["arquivo_pronto"],
        as_attachment=True,
        download_name=f"{config['download_name']}_{data_arquivo()}.xlsx",
    )


@app.route("/acessar/<tipo>", methods=["GET", "POST"])
def acessar(tipo):
    if tipo not in RELATORIOS:
        return render_template(
            "erro.html",
            titulo="Relatório inválido",
            mensagem="O relatório solicitado não existe.",
        ), 404

    config = RELATORIOS[tipo]

    if not config["restrito"]:
        return redirect(url_for("baixar", tipo=tipo))

    erro = ""

    if request.method == "POST":
        senha_digitada = request.form.get("senha", "")
        senha_correta = os.getenv("FINANCEIRO_PASSWORD")

        if not senha_correta:
            erro = "Senha do Financeiro não configurada no Railway."

        elif senha_digitada == senha_correta:
            if not os.path.exists(config["arquivo_pronto"]):
                return render_template(
                    "erro.html",
                    titulo="A planilha ainda não foi gerada.",
                    mensagem="Clique em Atualizar para gerar a primeira versão.",
                )

            return send_file(
                config["arquivo_pronto"],
                as_attachment=True,
                download_name=f"{config['download_name']}_{data_arquivo()}.xlsx",
            )

        else:
            erro = "Senha incorreta."

    return render_template(
        "acesso.html",
        tipo=tipo,
        relatorio=config,
        erro=erro,
    )


@app.route("/atualizar/<tipo>")
def atualizar(tipo):
    if tipo not in RELATORIOS:
        return render_template(
            "erro.html",
            titulo="Relatório inválido",
            mensagem="O relatório solicitado não existe.",
        ), 404

    try:
        gerar_planilha(tipo)

    except Exception as erro:
        return render_template(
            "erro.html",
            titulo="Erro ao atualizar a planilha.",
            mensagem=str(erro),
        )

    return render_template(
        "sucesso.html",
        relatorio=RELATORIOS[tipo],
    )


if __name__ == "__main__":
    app.run(debug=True)