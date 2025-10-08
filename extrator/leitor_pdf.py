# extrator/leitor_pdf.py
# Módulo responsável por ler e interpretar relatórios PDF de boletins
# Caminho: relatorio_horas/extrator/leitor_pdf.py
#
# Autor: Valdinei Lankewicz
# Criado em: 22/07/2025
# Histórico:
# - 22/07/2025: Separação da lógica de extração e parsing de PDFs
# - 22/07/2025: Adicionado suporte a barras de progresso visuais com set_arquivo
# - 22/07/2025: Adicionado suporte à leitura de cartões ponto

import re

import fitz  # PyMuPDF
import pandas as pd


# =================================================================================
# FUNÇÃO: extrair_dados_pdf(caminho_arquivo, set_arquivo=None)
# Lê o PDF e extrai o texto completo e os dados de cabeçalho como BOLETIM, DATA etc.
# =================================================================================
def extrair_dados_pdf(caminho_arquivo, set_arquivo=None):
    try:
        doc = fitz.open(caminho_arquivo)
        texto_completo = ""
        total_paginas = len(doc)
        for i, pagina in enumerate(doc):
            texto_completo += pagina.get_text("text", sort=True) + "\n"
            if set_arquivo:
                set_arquivo(((i + 1) / total_paginas) * 100)
        doc.close()
    except Exception as e:
        print(f"[ERRO] Falha ao abrir PDF '{caminho_arquivo}': {e}")
        return None, {}

    cabecalho = {"BOLETIM": None, "Data de Medição": None, "Contrato": None}
    linhas = texto_completo.split("\n")
    for i, linha in enumerate(linhas):
        if i + 1 < len(linhas):
            linha_seguinte = linhas[i + 1].strip()
            if "BOLETIM" in linha.upper() and "MEDIÇÃO" not in linha.upper():
                numeros = re.findall(r"\d{5,}", linha_seguinte)
                if numeros:
                    cabecalho["BOLETIM"] = numeros[-1]
            if "DATA MEDIÇÃO" in linha and "CONTRATO" in linha:
                partes = linha_seguinte.split()
                if len(partes) > 0:
                    cabecalho["Data de Medição"] = partes[0]
                if len(partes) > 1:
                    cabecalho["Contrato"] = partes[-1]
    return texto_completo, cabecalho


# =================================================================================
# FUNÇÃO: parse_horas_funcionarios(texto)
# Analisa o texto extraído e monta uma tabela com as horas dos funcionários
# =================================================================================
def parse_horas_funcionarios(texto):
    dados_extraidos = []
    registro_atual, funcionario_atual = None, None
    regex_ct = re.compile(r"CENTRO DE TRABALHO\s*-\s*(T\d+)-(.+)")

    for linha in texto.split("\n"):
        match = regex_ct.search(linha)
        if match:
            registro_atual = match.group(1).strip()
            funcionario_atual = match.group(2).strip().split("PROD.")[0].strip()
            continue

        if registro_atual and re.match(r"^\d{2}\.\d{2}\.\d{4}", linha.strip()):
            partes = re.split(r"\s{2,}", linha.strip())
            if len(partes) >= 12:
                dados_extraidos.append(
                    {
                        "Registro": registro_atual,
                        "Funcionário": funcionario_atual,
                        "DATA": partes[0],
                        "SERV.": partes[1],
                        "KM": partes[2],
                        "HORA NORMAL": partes[3],
                        "H.E.": partes[4],
                        "H.E.D.": partes[5],
                        "H.E.N.": partes[6],
                        "H.E.N.D.": partes[7],
                        "S.A.": partes[8],
                        "H.N.": partes[9],
                        "DESLOC.": partes[10],
                        "PROD.": partes[11],
                    }
                )

    if not dados_extraidos:
        return pd.DataFrame()

    df = pd.DataFrame(dados_extraidos)
    colunas_numericas = [
        "SERV.",
        "KM",
        "HORA NORMAL",
        "H.E.",
        "H.E.D.",
        "H.E.N.",
        "H.E.N.D.",
        "S.A.",
        "H.N.",
        "DESLOC.",
        "PROD.",
    ]
    for col in colunas_numericas:
        if col in df.columns:
            df[col] = df[col].str.replace(",", ".", regex=False)
            df[col] = pd.to_numeric(df[col], errors="coerce")
    return df


# =================================================================================
# FUNÇÃO: extrair_cartao_ponto(texto)
# Extrai dados de cartões ponto (registro de entrada/saída)
# =================================================================================
def extrair_cartao_ponto(texto):
    linhas = texto.split("\n")
    registros = []
    funcionario = None

    for linha in linhas:
        if linha.startswith("Nome do funcionário:"):
            funcionario = linha.split(":", 1)[1].strip()
            continue

        if re.match(r"\d{2}/\d{2}/\d{4}", linha):
            partes = linha.strip().split()
            if len(partes) >= 7:
                registros.append(
                    {
                        "Funcionário": funcionario,
                        "DATA": partes[0],
                        "ENTRADA 1": partes[1],
                        "SAÍDA 1": partes[2],
                        "ENTRADA 2": partes[3],
                        "SAÍDA 2": partes[4],
                        "ENTRADA 3": partes[5],
                        "SAÍDA 3": partes[6],
                    }
                )

    return pd.DataFrame(registros)
