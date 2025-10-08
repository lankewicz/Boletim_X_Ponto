# extrator/processador.py
# Versão: 2025-07-28 - Armazenamento em formato Parquet
import logging
from pathlib import Path
from typing import Callable, Set, Tuple

import pandas as pd

from extrator.leitor_pdf import extrair_dados_pdf, parse_horas_funcionarios

# Logging
logging.basicConfig(level=logging.INFO, format="[%(asctime)s] %(message)s")


class ProcessadorPDF:
    def __init__(self, pasta_raiz: str, pasta_saida: str = "data"):
        self.pasta_raiz = Path(pasta_raiz)
        self.pasta_saida = Path(pasta_saida)
        self.pasta_saida.mkdir(parents=True, exist_ok=True)
        self.caminho_parquet = self.pasta_saida / "RegistroBoletim.parquet"

        self.tipos_texto = ["BOLETIM", "Contrato", "Registro", "Funcionário"]
        self.tipos_data = ["DATA", "Data de Medição"]
        self.colunas_chave = ["BOLETIM", "DATA", "Funcionário"]

    def _criar_chave_unica(self, df: pd.DataFrame) -> pd.Series:
        df_temp = df.copy()
        df_temp["DATA_STR"] = pd.to_datetime(df_temp["DATA"], errors="coerce").dt.strftime(
            "%Y-%m-%d"
        )
        return (
            df_temp["BOLETIM"].astype(str).str.strip()
            + "__"
            + df_temp["DATA_STR"]
            + "__"
            + df_temp["Funcionário"].astype(str).str.strip().str.upper()
        )

    def _padronizar_tipos(self, df: pd.DataFrame) -> pd.DataFrame:
        df = df.copy()
        for col in self.tipos_texto:
            if col in df.columns:
                df[col] = df[col].astype(str).str.strip()
        for col in self.tipos_data:
            if col in df.columns:
                df[col] = pd.to_datetime(df[col], errors="coerce", dayfirst=True)
        for col in df.columns:
            if col not in self.tipos_texto + self.tipos_data:
                df[col] = pd.to_numeric(df[col], errors="coerce")
        return df

    def _carregar_existente(self) -> Tuple[pd.DataFrame, Set[str]]:
        if not self.caminho_parquet.exists():
            return pd.DataFrame(), set()
        df = pd.read_parquet(self.caminho_parquet)
        df = self._padronizar_tipos(df)
        df["chave_unica"] = self._criar_chave_unica(df)
        df = df.drop_duplicates(subset=["chave_unica"])
        return df, set(df["chave_unica"])

    def _buscar_pdfs(self) -> list:
        return sorted(self.pasta_raiz.rglob("*.pdf"))

    def _processar_pdf(
        self,
        caminho_pdf: Path,
        chaves_existentes: Set[str],
        callback_arquivo: Callable | None = None,
    ) -> pd.DataFrame | None:
        texto, cabecalho = extrair_dados_pdf(str(caminho_pdf), callback_arquivo)
        boletim = str(cabecalho.get("BOLETIM", "")).strip()
        if not boletim:
            return None
        df = parse_horas_funcionarios(texto)
        if df is None or df.empty:
            return None
        df["BOLETIM"] = boletim
        df["Data de Medição"] = cabecalho.get("Data de Medição")
        df["Contrato"] = cabecalho.get("Contrato")
        df = self._padronizar_tipos(df)
        df = df.dropna(subset=["DATA", "Funcionário"])
        df["chave_unica"] = self._criar_chave_unica(df)
        df_novos = df[~df["chave_unica"].isin(chaves_existentes)].copy()
        return df_novos if not df_novos.empty else None

    def consolidar_pdfs_em_parquet(
        self, callback_total: Callable | None = None, callback_arquivo: Callable | None = None
    ) -> str:
        df_existente, chaves = self._carregar_existente()
        arquivos = self._buscar_pdfs()
        novos = []

        for i, pdf in enumerate(arquivos):
            if callback_total:
                callback_total((i + 1) / len(arquivos) * 100)
            df_novo = self._processar_pdf(pdf, chaves, callback_arquivo)
            if df_novo is not None:
                novos.append(df_novo)
                chaves.update(df_novo["chave_unica"])

        if not novos and df_existente.empty:
            raise ValueError("Nenhum dado novo e nada existente.")

        todos = pd.concat([df_existente] + novos, ignore_index=True) if novos else df_existente
        todos = todos.drop_duplicates(subset=["chave_unica"])
        todos = todos[todos["chave_unica"].str.len() > 6]
        todos = todos.drop(columns=["chave_unica"], errors="ignore")
        todos.to_parquet(self.caminho_parquet, index=False, compression="snappy")
        return str(self.caminho_parquet)


# Interface compatível
def consolidar_pdfs_em_excel(pasta_raiz, set_total=None, set_arquivo=None):
    return ProcessadorPDF(pasta_raiz).consolidar_pdfs_em_parquet(set_total, set_arquivo)
