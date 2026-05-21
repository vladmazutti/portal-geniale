from flask import Flask, send_file
import requests
from openpyxl import load_workbook
from datetime import datetime
import time

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

        <h1 style="font-size:40px;">Portal Geniale</h1>

        <p style="font-size:20px;">Relatórios Automatizados</p>

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
# BUSCAR PÁGINA OMIE
# ====================================

def buscar_pagina(pagina, app_key, app_secret):
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

    tentativas = 0

    while tentativas < 5:
        response = requests.post(url, json=payload, timeout=60)

        if response.status_code == 429:
            print(f"Página {pagina}: limite OMIE atingido. Aguardando...")
            time.sleep(5)
            tentativas += 1
            continue

        response.raise_for_status()

        dados = response.json()
        clientes = dados.get("clientes_cadastro", [])
        total_paginas = dados.get("total_de_paginas", 1)

        print(f"Página {pagina} carregada...")

        return clientes, total_paginas

    raise Exception(f"Não foi possível carregar a página {pagina} após várias tentativas.")

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

    linhas = []

    # ====================================
    # PRIMEIRA PÁGINA
    # ====================================

    clientes, total_paginas = buscar_pagina(1, app_key, app_secret)

    todas_paginas = {
        1: clientes
    }

    # ====================================
    # DEMAIS PÁGINAS - SEQUENCIAL SEGURO
    # ====================================

    for pagina in range(2, total_paginas + 1):
        clientes, _ = buscar_pagina(pagina, app_key, app_secret)
        todas_paginas[pagina] = clientes

        # pausa leve para evitar bloqueio da API
        time.sleep(0.2)

    # ====================================
    # MONTAR LINHAS
    # ====================================

    for pagina in sorted(todas_paginas.keys()):
        for cliente in todas_paginas[pagina]:

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

    # ====================================
    # ABRIR MODELO
    # ====================================

    wb = load_workbook("modelo.xlsx")
    ws = wb["BASE OMIE"]

    # ====================================
    # OCULTAR ABA BASE
    # ====================================

    ws.sheet_state = "veryHidden"

    # ====================================
    # LIMPAR DADOS ANTIGOS
    # ====================================

    ws.delete_rows(1, ws.max_row)

    # ====================================
    # CABEÇALHOS
    # ====================================

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

    # ====================================
    # COLAR DADOS
    # ====================================

    for linha in linhas:
        ws.append(linha)

    # ====================================
    # SALVAR ARQUIVO
    # ====================================

    data_hoje = datetime.now().strftime("%d%m%y")
    arquivo_final = f"Pagamentos{data_hoje}.xlsx"

    wb.save(arquivo_final)

    return send_file(
        arquivo_final,
        as_attachment=True,
        download_name=arquivo_final
    )

# ====================================
# INICIAR SERVIDOR LOCAL
# ====================================

if __name__ == "__main__":
    app.run(debug=True)