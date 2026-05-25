from flask import Flask, send_file, request, redirect, url_for, render_template, jsonify
import requests
from requests.exceptions import Timeout, RequestException
from openpyxl import load_workbook
from datetime import datetime
from zoneinfo import ZoneInfo
from apscheduler.schedulers.background import BackgroundScheduler
from threading import Thread, Lock
import time
import os

app = Flask(__name__)

FUSO_BRASIL = ZoneInfo("America/Sao_Paulo")

jobs_em_execucao = {}
jobs_lock = Lock()

RELATORIOS = {
    "geniale": {
        "nome": "GENIALE",
        "titulo": "Planilha de Pagamento",
        "descricao": "Pagamentos operacionais Geniale",
        "modelo": "modelo.xlsx",
        "arquivo_pronto": "planilha_geniale.xlsx",
        "arquivo_atualizacao": "ultima_atualizacao_geniale.txt",
        "arquivo_status": "status_geniale.txt",
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
        "arquivo_status": "status_paniz.txt",
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
        "arquivo_status": "status_financeiro.txt",
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
    return agora_brasil().strftime("%d/%m/%Y %H:%M:%S")


def data_arquivo():
    return agora_brasil().strftime("%d%m%y")


def escrever_arquivo(caminho, conteudo):
    with open(caminho, "w", encoding="utf-8") as f:
        f.write(conteudo)


def ler_arquivo(caminho, padrao=""):
    if os.path.exists(caminho):
        with open(caminho, "r", encoding="utf-8") as f:
            return f.read().strip()

    return padrao


def definir_status(tipo, status):
    config = RELATORIOS[tipo]
    escrever_arquivo(config["arquivo_status"], status)


def obter_ultima_atualizacao(config):
    return ler_arquivo(
        config["arquivo_atualizacao"],
        "Ainda não atualizada"
    )


def obter_status(tipo, config):
    status_interno = ler_arquivo(config["arquivo_status"], "")

    if status_interno == "atualizando":
        return {
            "texto": "Atualizando...",
            "cor": "atualizando",
            "existe": os.path.exists(config["arquivo_pronto"]),
            "ultima": obter_ultima_atualizacao(config),
        }

    if status_interno.startswith("erro"):
        return {
            "texto": "Erro na atualização",
            "cor": "erro",
            "existe": os.path.exists(config["arquivo_pronto"]),
            "ultima": obter_ultima_atualizacao(config),
        }

    existe = os.path.exists(config["arquivo_pronto"])

    return {
        "texto": "Disponível para download" if existe else "Ainda não gerada",
        "cor": "ok" if existe else "pendente",
        "existe": existe,
        "ultima": obter_ultima_atualizacao(config),
    }


def montar_relatorios_para_tela():
    lista = []

    for tipo, config in RELATORIOS.items():
        item = dict(config)
        item["tipo"] = tipo
        item["status"] = obter_status(tipo, config)
        lista.append(item)

    return lista


def buscar_pagina_omie(app_key, app_secret, pagina, tentativa_maxima=3):
    url = "https://app.omie.com.br/api/v1/geral/clientes/"

    payload = {
        "call": "ListarClientes",
        "app_key": app_key,
        "app_secret": app_secret,
        "param": [
            {
                "pagina": pagina,
                "registros_por_pagina": 1000,
                "apenas_importado_api": "N",
            }
        ],
    }

    for tentativa in range(1, tentativa_maxima + 1):
        try:
            response = requests.post(
                url,
                json=payload,
                timeout=90
            )

            if response.status_code == 429:
                espera = tentativa * 5
                print(
                    f"Limite OMIE atingido na página {pagina}. "
                    f"Aguardando {espera}s..."
                )
                time.sleep(espera)
                continue

            response.raise_for_status()
            return response.json()

        except Timeout:
            espera = tentativa * 5
            print(
                f"Timeout na página {pagina}. "
                f"Tentativa {tentativa}/{tentativa_maxima}. "
                f"Aguardando {espera}s..."
            )
            time.sleep(espera)

        except RequestException as erro:
            espera = tentativa * 5
            print(
                f"Erro de conexão na página {pagina}: {erro}. "
                f"Tentativa {tentativa}/{tentativa_maxima}. "
                f"Aguardando {espera}s..."
            )
            time.sleep(espera)

    raise Exception(
        f"Falha ao consultar OMIE na página {pagina} "
        f"após {tentativa_maxima} tentativas."
    )


def buscar_clientes_omie(app_key, app_secret):
    linhas = []
    pagina = 1
    total_paginas = 1

    while pagina <= total_paginas:
        dados = buscar_pagina_omie(app_key, app_secret, pagina)

        total_paginas = dados.get("total_de_paginas", 1)
        clientes = dados.get("clientes_cadastro", [])

        print(f"Página {pagina}/{total_paginas} carregada...")

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

    return linhas


def salvar_workbook_com_seguranca(wb, arquivo_final):
    arquivo_temporario = f"{arquivo_final}.tmp.xlsx"

    if os.path.exists(arquivo_temporario):
        os.remove(arquivo_temporario)

    wb.save(arquivo_temporario)

    if not os.path.exists(arquivo_temporario):
        raise Exception("Arquivo temporário não foi gerado corretamente.")

    os.replace(arquivo_temporario, arquivo_final)


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

    salvar_workbook_com_seguranca(
        wb,
        config["arquivo_pronto"]
    )

    escrever_arquivo(
        config["arquivo_atualizacao"],
        agora_formatado()
    )

    print(f"Planilha {config['nome']} atualizada com sucesso!")


def executar_atualizacao_background(tipo):
    try:
        definir_status(tipo, "atualizando")
        gerar_planilha(tipo)
        definir_status(tipo, "ok")

    except Exception as erro:
        print(f"Erro ao atualizar {tipo}: {erro}")
        definir_status(tipo, f"erro: {erro}")

    finally:
        with jobs_lock:
            jobs_em_execucao[tipo] = False


def iniciar_atualizacao_background(tipo):
    if tipo not in RELATORIOS:
        raise Exception("Relatório inválido.")

    with jobs_lock:
        if jobs_em_execucao.get(tipo):
            return False

        jobs_em_execucao[tipo] = True

    thread = Thread(
        target=executar_atualizacao_background,
        args=(tipo,),
        daemon=True
    )
    thread.start()

    return True


def atualizar_agendado(tipo):
    try:
        iniciar_atualizacao_background(tipo)
    except Exception as erro:
        print(f"Erro ao iniciar atualização agendada de {tipo}: {erro}")


scheduler = BackgroundScheduler(timezone="America/Sao_Paulo")

scheduler.add_job(
    lambda: atualizar_agendado("geniale"),
    trigger="cron",
    hour=8,
    minute=0,
)

scheduler.add_job(
    lambda: atualizar_agendado("paniz"),
    trigger="cron",
    hour=8,
    minute=0,
)

scheduler.add_job(
    lambda: atualizar_agendado("financeiro"),
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


@app.route("/api/status")
def api_status():
    return jsonify(montar_relatorios_para_tela())


@app.route("/api/atualizar/<tipo>")
def api_atualizar(tipo):
    if tipo not in RELATORIOS:
        return jsonify({
            "sucesso": False,
            "mensagem": "Relatório inválido."
        }), 404

    try:
        iniciado = iniciar_atualizacao_background(tipo)

        if iniciado:
            return jsonify({
                "sucesso": True,
                "mensagem": (
                    f"A atualização da planilha {RELATORIOS[tipo]['nome']} "
                    f"foi iniciada em segundo plano."
                )
            })

        return jsonify({
            "sucesso": False,
            "mensagem": (
                f"A planilha {RELATORIOS[tipo]['nome']} "
                f"já está sendo atualizada."
            )
        })

    except Exception as erro:
        return jsonify({
            "sucesso": False,
            "mensagem": str(erro)
        }), 500


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
        iniciado = iniciar_atualizacao_background(tipo)

    except Exception as erro:
        return render_template(
            "erro.html",
            titulo="Erro ao iniciar atualização.",
            mensagem=str(erro),
        )

    if iniciado:
        mensagem = (
            f"A atualização da planilha {RELATORIOS[tipo]['nome']} "
            f"foi iniciada em segundo plano."
        )
    else:
        mensagem = (
            f"A planilha {RELATORIOS[tipo]['nome']} "
            f"já está sendo atualizada."
        )

    return render_template(
        "sucesso.html",
        relatorio=RELATORIOS[tipo],
        mensagem=mensagem,
    )


if __name__ == "__main__":
    app.run(debug=True)