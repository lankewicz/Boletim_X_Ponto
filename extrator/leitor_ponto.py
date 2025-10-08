# extrator/leitor_ponto.py
# Versão final com suporte a datas como "01/06/2025 DOM" e normalização de colunas
# Versão com deduplicacao robusta usando CPF ou PIS

import os
import unicodedata
from io import StringIO

import pandas as pd

CAMINHO_PARQUET = "data/RegistroPonto.parquet"
CAMINHO_VERIFICACAO = "data/RegistroPonto_VERIFICACAO.txt"


def tempo_para_decimal(valor):
    try:
        if pd.isna(valor) or str(valor).strip() == "":
            return 0
        if isinstance(valor, (float, int)):
            return round(float(valor), 2)
        partes = str(valor).strip().replace(",", ".").split(":")
        if len(partes) == 2:
            horas = int(partes[0])
            minutos = int(partes[1])
            return round(horas + minutos / 60, 2)
        return round(float(partes[0]), 2)
    except:
        return 0


def detectar_delimitador_conteudo(texto):
    if texto.count("\t") > texto.count(";") and texto.count("\t") > texto.count(","):
        return "\t"
    elif texto.count(";") > texto.count(","):
        return ";"
    else:
        return ","


def normalizar_coluna(nome):
    return (
        unicodedata.normalize("NFKD", nome)
        .encode("ASCII", "ignore")
        .decode("utf-8")
        .strip()
        .lower()
    )


def encontrar_coluna(df, opcoes):
    normalizado = {normalizar_coluna(col): col for col in df.columns}
    for opcao in opcoes:
        chave = normalizar_coluna(opcao)
        if chave in normalizado:
            return normalizado[chave]
    return None


def ler_arquivo_compativel(path):
    try:
        if path.lower().endswith(".csv"):
            return pd.read_csv(path, encoding="utf-8", sep=None, engine="python")
        elif path.lower().endswith((".xls", ".xlsx")):
            try:
                return pd.read_excel(path, engine="openpyxl")
            except Exception:
                with open(path, "r", encoding="utf-8", errors="ignore") as f:
                    conteudo = f.read()
                delimitador = detectar_delimitador_conteudo(conteudo)
                return pd.read_csv(StringIO(conteudo), sep=delimitador)
    except Exception as e:
        print(f"[ERRO] Falha ao ler {path}: {e}")
        return None


def processar_todos_csvs(pasta_ponto, set_total=None, set_arquivo=None):
    arquivos = []
    for raiz, _, nomes in os.walk(pasta_ponto):
        for nome in nomes:
            if nome.lower().endswith((".csv", ".xls", ".xlsx")):
                arquivos.append(os.path.join(raiz, nome))

    total = len(arquivos)
    todos_registros = []

    for idx, caminho in enumerate(arquivos):
        if set_total:
            set_total((idx + 1) * 100 / total)
        if set_arquivo:
            set_arquivo(0, os.path.basename(caminho))

        df = ler_arquivo_compativel(caminho)
        if df is None or df.empty:
            print(f"[AVISO] Arquivo vazio ou ilegível: {caminho}")
            continue

        col_nome = encontrar_coluna(df, ["Nome do funcionário", "funcionário", "nome"])
        col_data = encontrar_coluna(df, ["Dia", "Data"])
        col_cpf = encontrar_coluna(df, ["CPF do funcionário", "cpf"])
        col_pis = encontrar_coluna(df, ["PIS do funcionário", "pis"])

        if not col_nome or not col_data:
            print(f"[IGNORADO] {os.path.basename(caminho)} - Colunas obrigatórias ausentes.")
            print(f" > Colunas detectadas: {list(df.columns)}")
            continue

        for _, linha in df.iterrows():
            try:
                nome = str(linha.get(col_nome)).strip()
                cpf = str(linha.get(col_cpf, "")).strip()
                pis = str(linha.get(col_pis, "")).strip()
                data = linha.get(col_data)

                if isinstance(data, str):
                    data = data.strip().split()[0]
                    data = pd.to_datetime(data, dayfirst=True, errors="coerce")
                elif isinstance(data, (float, int)):
                    data = pd.to_datetime("1899-12-30") + pd.to_timedelta(data, unit="D")

                if pd.isna(data):
                    continue

                chave = f"{cpf}__{data.date()}" if cpf else f"{pis}__{data.date()}"

                registro = {
                    "Nome": nome.upper(),
                    "CPF": cpf,
                    "PIS": pis,
                    "Data": data.date(),
                    "Total Normais": tempo_para_decimal(linha.get("Total Normais", "")),
                    "Total Noturno": tempo_para_decimal(linha.get("Total Noturno", "")),
                    "Extra 50%D": tempo_para_decimal(linha.get("Extra   50%D", "")),
                    "Extra 100%D": tempo_para_decimal(linha.get("Extra   100%D", "")),
                    "Extra 50%N": tempo_para_decimal(linha.get("Extra   50%N", "")),
                    "Extra 100%N": tempo_para_decimal(linha.get("Extra   100%N", "")),
                    "chave_unica": chave,
                }
                todos_registros.append(registro)
            except Exception as e:
                print(f"[ERRO] Falha na linha em {caminho}: {e}")

    df_final = pd.DataFrame(todos_registros)

    if df_final.empty:
        print("[AVISO] Nenhum dado de ponto processado.")
        return

    if os.path.exists(CAMINHO_PARQUET):
        df_antigo = pd.read_parquet(CAMINHO_PARQUET)
        df_final = pd.concat([df_antigo, df_final], ignore_index=True)

    df_final.drop_duplicates(subset=["chave_unica"], keep="last", inplace=True)

    df_final["Data"] = pd.to_datetime(df_final["Data"], errors="coerce")
    df_final.to_parquet(CAMINHO_PARQUET, index=False)

    try:
        with open(CAMINHO_VERIFICACAO, "w", encoding="utf-8") as f:
            f.write("Data;Nome;CPF;PIS;Total Normais;Extra 50%D;Extra 100%D;Total Noturno\n")
            for _, linha in df_final.head(20).iterrows():
                f.write(
                    f"{linha['Data']};{linha['Nome']};{linha['CPF']};{linha['PIS']};{linha.get('Total Normais', 0)};{linha.get('Extra 50%D', 0)};{linha.get('Extra 100%D', 0)};{linha.get('Total Noturno', 0)}\n"
                )
    except Exception as e:
        print(f"[AVISO] Falha ao gerar verificação: {e}")
