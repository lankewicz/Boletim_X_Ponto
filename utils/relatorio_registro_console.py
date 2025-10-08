# relatorio_registro_console.py
# ------------------------------------------------------------
# Relatório por REGISTRO, organizado por Funcionário.
# Mostra colunas: Funcionário | Registro | Boletim | Contrato | S.A.
# Regras:
# - Agrega S.A. por (Funcionário, Registro, BOLETIM, Contrato).
# - Ordena por Funcionário (depois BOLETIM para manter agrupamento estável).
# - Adiciona linha "Subtotal <NOME>" SOMENTE se o funcionário aparecer
#   em DOIS OU MAIS boletins distintos.
#
# Execução (Windows):
#   py relatorio_registro_console.py
#   py relatorio_registro_console.py --arquivo .\data\RegistroBoletim.parquet
# ------------------------------------------------------------

from __future__ import annotations
import argparse
import sys
from typing import List, Dict, Any

import pandas as pd


ARQUIVO_PADRAO = r".\data\RegistroBoletim.parquet"
COLS_NECESSARIAS = ["Funcionário", "Registro", "BOLETIM", "Contrato", "S.A."]


def _carregar_parquet(caminho: str) -> pd.DataFrame:
    """
    Carrega o parquet e retorna DataFrame.
    Dá uma mensagem amigável se faltar 'pyarrow' ou 'fastparquet'.
    """
    try:
        df = pd.read_parquet(caminho)
        return df
    except ImportError as e:
        msg = str(e).lower()
        print(
            "[ERRO] Suporte a Parquet não encontrado.\n"
            "Instale um motor: 'pip install pyarrow' (recomendado) ou 'pip install fastparquet'.",
            file=sys.stderr,
        )
        raise
    except FileNotFoundError:
        print(f"[ERRO] Arquivo não encontrado: {caminho}", file=sys.stderr)
        raise
    except Exception as e:
        print(f"[ERRO] Falha ao ler parquet '{caminho}': {e}", file=sys.stderr)
        raise



def _validar_colunas(df: pd.DataFrame):
    faltando = [c for c in COLS_NECESSARIAS if c not in df.columns]
    if faltando:
        raise ValueError(f"Colunas ausentes no arquivo parquet: {faltando}")

def _forcar_tipos(df: pd.DataFrame, cols_extras: List[str] | None = None) -> pd.DataFrame:
    """
    Garante que as colunas existam e estejam no tipo esperado para processamento.
    NOVA VERSÃO: Aceita uma lista de colunas extras para preservar.
    """
    df = df.copy()
    
    # Lista de colunas a serem mantidas
    cols_a_manter = COLS_NECESSARIAS[:] # Faz uma cópia da lista original
    if cols_extras:
        for col in cols_extras:
            if col in df.columns and col not in cols_a_manter:
                cols_a_manter.append(col)

    # Garante presença das colunas essenciais
    for c in COLS_NECESSARIAS:
        if c not in df.columns:
            df[c] = pd.NA

    # Normaliza 'S.A.' para numérico
    df["S.A."] = pd.to_numeric(df["S.A."], errors="coerce").fillna(0.0)

    # Converte identificadores para string (evita None/float estranhos)
    for c in ["Funcionário", "Registro", "BOLETIM", "Contrato"]:
        if c in df.columns: # Verifica se a coluna existe antes de converter
            df[c] = df[c].astype(str).fillna("")

    # Retorna apenas as colunas desejadas (essenciais + extras)
    return df[cols_a_manter]

def _formatar_e_imprimir(linhas: List[Dict[str, Any]]):
    """
    Imprime em formato de tabela no console, com larguras ajustadas.
    Numéricos (S.A.) alinhados à direita com 3 casas.
    """
    if not linhas:
        print("(Relatório vazio)")
        return

    # Define ordem e cabeçalho
    headers = ["Funcionário", "Registro", "BOLETIM", "Contrato", "S.A."]

    # Calcula larguras
    col_w = {h: len(h) for h in headers}
    for row in linhas:
        for h in headers:
            txt = row[h]
            if h == "S.A.":
                # comprimento considerando formatação numérica
                try:
                    width_val = len(f"{float(txt):.3f}")
                except Exception:
                    width_val = len(str(txt))
                col_w[h] = max(col_w[h], width_val)
            else:
                col_w[h] = max(col_w[h], len(str(txt)))

    # Funções de formatação
    def fmt_cell(h: str, v: Any) -> str:
        if h == "S.A.":
            if isinstance(v, (int, float)) or str(v).replace(".", "", 1).isdigit():
                try:
                    return f"{float(v):>{col_w[h]}.3f}"
                except Exception:
                    return f"{str(v):>{col_w[h]}}"
            else:
                return f"{str(v):>{col_w[h]}}"
        else:
            return f"{str(v):<{col_w[h]}}"

    # Linhas de separação
    sep = " | "
    line_sep = "-+-".join("-" * col_w[h] for h in headers)

    # Cabeçalho
    header_line = sep.join(f"{h:<{col_w[h]}}" for h in headers)
    print(header_line)
    print(line_sep)

    # Corpo
    for row in linhas:
        # Negrito não existe no console padrão; apenas destacamos "Subtotal" com asteriscos
        if str(row["Funcionário"]).startswith("Subtotal "):
            # Marca o rótulo do subtotal
            row_fmt = row.copy()
            row_fmt["Funcionário"] = f"**{row['Funcionário']}**"
        else:
            row_fmt = row

        print(sep.join(fmt_cell(h, row_fmt[h]) for h in headers))

def _parse_args():
    ap = argparse.ArgumentParser(
        description="Relatório por REGISTRO (console): Funcionário | Registro | Boletim | Contrato | S.A."
    )
    ap.add_argument(
        "--arquivo",
        "-f",
        default=ARQUIVO_PADRAO,
        help=f"Caminho do arquivo .parquet (padrão: {ARQUIVO_PADRAO})",
    )
    return ap.parse_args()




# NO ARQUIVO: relatorio_registro_console.py


if __name__ == "__main__":
    args = _parse_args()
    gerar_relatorio_console(args.arquivo)



