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
    linhas = [l.strip() for l in texto_completo.split("\n") if l.strip()]

    for i, linha in enumerate(linhas):
        linha_upper = linha.upper()

        # Parse CONTRATO
        if "CONTRATO" in linha_upper:
            m = re.search(r"CONTRATO\s*[:\s]*(\d{8,})", linha_upper)
            if m:
                cabecalho["Contrato"] = m.group(1)
            elif i + 1 < len(linhas):
                prox = linhas[i + 1]
                nums = re.findall(r"\d{8,}", prox)
                if nums:
                    if "DATA" in linha_upper or linha_upper.endswith("CONTRATO"):
                        cabecalho["Contrato"] = nums[-1]
                    else:
                        cabecalho["Contrato"] = nums[0]

        # Parse BOLETIM (ignora o título 'BOLETIM DE MEDIÇÃO')
        if "BOLETIM" in linha_upper and not re.search(r"BOLETIM\s+DE\s+MEDI[ÇC][ÃA]O", linha_upper):
            m = re.search(r"BOLETIM\s*[:\s]*(\d{5,})", linha_upper)
            if m:
                cabecalho["BOLETIM"] = m.group(1)
            elif i + 1 < len(linhas):
                prox = linhas[i + 1]
                nums = re.findall(r"\d{5,}", prox)
                if nums:
                    boletim_candidates = [n for n in nums if len(n) >= 5]
                    if boletim_candidates:
                        if "DATA" in linha_upper or linha_upper.startswith("BOLETIM"):
                            cabecalho["BOLETIM"] = boletim_candidates[0]
                        else:
                            cabecalho["BOLETIM"] = boletim_candidates[-1]

        # Parse DATA MEDIÇÃO
        if "DATA" in linha_upper and ("MEDIÇÃO" in linha_upper or "MEDICÃO" in linha_upper or "MEDICAO" in linha_upper):
            m = re.search(r"DATA\s+MEDI[ÇC][ÃA]O\s*[:\s]*(\d{2}[\./]\d{2}[\./]\d{4})", linha_upper)
            if m:
                cabecalho["Data de Medição"] = m.group(1)
            elif i + 1 < len(linhas):
                prox = linhas[i + 1]
                dates = re.findall(r"\b\d{2}[\./]\d{2}[\./]\d{4}\b", prox)
                if dates:
                    cabecalho["Data de Medição"] = dates[0]

    # Fallbacks globais via expressão regular caso alguma chave não tenha sido encontrada
    if not cabecalho["BOLETIM"]:
        m = re.search(r"(?:BOLETIM\s*[:\s]*|BOLETIM\s*\n\s*)(\d{5,})", texto_completo, re.IGNORECASE)
        if m and m.group(1) != "00000":
            cabecalho["BOLETIM"] = m.group(1)

    if not cabecalho["Contrato"]:
        m = re.search(r"(?:CONTRATO\s*[:\s]*|CONTRATO\s*\n\s*)(\d{8,})", texto_completo, re.IGNORECASE)
        if m:
            cabecalho["Contrato"] = m.group(1)
        else:
            m_file = re.search(r"CONTRATO\s*(\d{8,})", Path(caminho_arquivo).name if isinstance(caminho_arquivo, (str, Path)) else "", re.IGNORECASE)
            if m_file:
                cabecalho["Contrato"] = m_file.group(1)

    if not cabecalho["Data de Medição"]:
        m = re.search(r"(?:DATA\s+MEDI[ÇC][ÃA]O\s*[:\s]*|DATA\s+MEDI[ÇC][ÃA]O\s*\n\s*)(\d{2}[\./]\d{2}[\./]\d{4})", texto_completo, re.IGNORECASE)
        if m:
            cabecalho["Data de Medição"] = m.group(1)

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

# =================================================================================
# NOVO: utilitários para VAR (pt-BR numérico + mapeamentos)
# =================================================================================
import re
import pandas as pd

_MAP_TERMO_POR_VAR = {
    "VAR000": "HORA NORMAL",  # DSP001 HORA NORMAL
    "VAR001": "H.E.",         # Horas Extras 50%
    "VAR002": "H.E.N.",       # Horas Extras Noturnas 50%
    "VAR003": "H.E.D.",       # Horas Extras 100%
    "VAR004": "H.E.N.D.",     # Horas Extras Noturnas 100%
    "VAR005": "S.A.",         # Sobreaviso
    "VAR006": "H.N.",         # Horas Noturnas (adicional)
    "VAR007": "KM",           # Quilometragem rodada (base para VAR008 também)
    "VAR008": "DESLOC.",      # Deslocamento de pessoal p/KM
}

def _ptbr_float(s: str | None) -> float | None:
    """Converte '10.809,438' -> 10809.438; retorna None se vazio/inválido."""
    if s is None:
        return None
    s = str(s).strip()
    if not s:
        return None
    s = s.replace('.', '').replace(',', '.')
    try:
        return float(s)
    except Exception:
        return None

# =================================================================================
# NOVO: localizar o valor da US do boletim (ex.: 'R$ 42,01', 'R$ 45,00', ...)
# =================================================================================
def parse_valor_us(texto: str) -> float | None:
    """
    Procura o primeiro padrão 'R$ 9.999,99' no texto.
    Se houver mais de um, retorna o primeiro (geralmente o valor US do boletim).
    """
    m = re.search(r'R\$\s*([0-9\.\,]{3,})', texto)
    return _ptbr_float(m.group(1)) if m else None

# =================================================================================
# NOVO: extrair bloco de itens VAR da seção 'DESCRIÇÃO  US MÉDIA  QTDE  TOTAL'
# =================================================================================
def _slice_bloco_var(texto: str) -> str:
    """
    Corta o texto a partir do cabeçalho da tabela dos itens faturados.
    Aceita 'MÉDIA' com/sem acento.
    """
    cab = re.search(r'DESCRI[ÇC][ÃA]O\s+US\s+M[EÉ]DIA\s+QTDE\s+TOTAL', texto, flags=re.IGNORECASE)
    if not cab:
        return ""
    return texto[cab.end():]

def parse_var_itens(texto: str) -> pd.DataFrame:
    """
    Lê as linhas 'VARxxx - DESCRIÇÃO  US  QTDE  TOTAL' do bloco de itens do BMD.
    Retorna colunas: var_code, descrição, US, qtde, total (total é auxiliar para conferência).
    """
    bloco = _slice_bloco_var(texto)
    if not bloco:
        return pd.DataFrame(columns=['var_code', 'descrição', 'US', 'qtde', 'total'])

    padrao = re.compile(
        r'^(VAR\d{3})\s*-\s*(.*?)\s+([0-9\.,]+)\s+([0-9\.,]+)\s+([0-9\.,]+)$',
        flags=re.MULTILINE
    )
    rows = []
    for m in padrao.finditer(bloco):
        rows.append({
            'var_code': m.group(1).strip(),
            'descrição': m.group(2).strip(),
            'US': _ptbr_float(m.group(3)),
            'qtde': _ptbr_float(m.group(4)),
            'total': _ptbr_float(m.group(5)),  # não faz parte do layout final; só para QA
        })
    df = pd.DataFrame(rows, columns=['var_code', 'descrição', 'US', 'qtde', 'total'])
    return df

# =================================================================================
# NOVO: montagem de dataset padrão para os VAR (com seus termos canônicos)
# =================================================================================
def montar_dataset_var(cabecalho: dict, texto: str, incluir_termo: bool = True) -> pd.DataFrame:
    """
    Monta o dataset com as colunas finais:
    contrato, boletim, data_medicao, valor_us, var_code, [termo], descrição, US, qtde

    - 'termo' é opcional; quando True, usa seus rótulos: HORA NORMAL, H.E., H.E.D., ...
    - 'valor_us' é lido do cabeçalho (R$ ...), se presente.
    """
    df_var = parse_var_itens(texto)
    if df_var.empty:
        cols = ['contrato','boletim','data_medicao','valor_us','var_code','descrição','US','qtde']
        if incluir_termo:
            cols.insert(cols.index('descrição'), 'termo')
        return pd.DataFrame(columns=cols)

    valor_us = parse_valor_us(texto)
    df_var.insert(0, 'contrato', cabecalho.get('Contrato'))
    df_var.insert(1, 'boletim', cabecalho.get('BOLETIM'))
    df_var.insert(2, 'data_medicao', cabecalho.get('Data de Medição'))
    df_var.insert(3, 'valor_us', valor_us)

    if incluir_termo:
        df_var.insert(5, 'termo', df_var['var_code'].map(_MAP_TERMO_POR_VAR))

    # Ordena colunas conforme layout desejado
    base_cols = ['contrato','boletim','data_medicao','valor_us','var_code']
    if incluir_termo:
        out_cols = base_cols + ['termo','descrição','US','qtde']
    else:
        out_cols = base_cols + ['descrição','US','qtde']
    return df_var[out_cols]

# =================================================================================
# NOVO (opcional): fluxo completo para extrair VAR direto do PDF
# =================================================================================
def extrair_var_dataset(caminho_arquivo: str, set_arquivo=None, incluir_termo: bool = True) -> pd.DataFrame:
    """
    Abre o PDF (reusa extrair_dados_pdf), e retorna o DataFrame dos VAR no layout final.
    """
    texto, cab = extrair_dados_pdf(caminho_arquivo, set_arquivo=set_arquivo)
    if not texto:
        return pd.DataFrame(columns=['contrato','boletim','data_medicao','valor_us',
                                     'var_code'] + (['termo'] if incluir_termo else []) +
                                    ['descrição','US','qtde'])
    return montar_dataset_var(cab, texto, incluir_termo=incluir_termo)
