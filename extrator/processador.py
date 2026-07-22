# extrator/processador.py
# Versão: 2025-07-28 - Armazenamento em formato Parquet
import logging

import re
from pathlib import Path
from typing import Callable, Set, Tuple

import pandas as pd

from extrator.leitor_pdf import extrair_dados_pdf, parse_horas_funcionarios, extrair_var_dataset

# Logging
logging.basicConfig(level=logging.INFO, format="[%(asctime)s] %(message)s")


class ProcessadorPDF:
    def __init__(self, pasta_raiz: str, pasta_saida: str = "data"):
        self.pasta_raiz = Path(pasta_raiz)
        self.pasta_saida = Path(pasta_saida)
        self.pasta_saida.mkdir(parents=True, exist_ok=True)
        self.caminho_parquet = self.pasta_saida / "RegistroBoletim.parquet"
        # NOVO: saídas extras
        self.caminho_var = self.pasta_saida / "boletim_var.parquet"
        self.caminho_relacao = self.pasta_saida / "relacao_contrato_boletim.csv"

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

    # NOVO: padronização dos rótulos de horas
    def _padronizar_rotulos_horas(self, df: pd.DataFrame) -> pd.DataFrame:
        if df is None or df.empty:
            return df
        mapa = {
            # sinônimos possíveis -> padrão
            "HORA NORMAL": "HORA NORMAL",
            "H.N.": "H.N.",
            "HN": "H.N.",
            "HE": "H.E.",
            "H.E": "H.E.",
            "H.E.": "H.E.",
            "HED": "H.E.D.",
            "H.E.D": "H.E.D.",
            "H.E.D.": "H.E.D.",
            "HEN": "H.E.N.",
            "H.E.N": "H.E.N.",
            "H.E.N.": "H.E.N.",
            "HEND": "H.E.N.D.",
            "H.E.N.D": "H.E.N.D.",
            "H.E.N.D.": "H.E.N.D.",
            "SA": "S.A.",
            "S.A": "S.A.",
            "S.A.": "S.A.",
        }
        renames = {}
        for col in list(df.columns):
            if col.upper() in mapa:
                renames[col] = mapa[col.upper()]
        if renames:
            df = df.rename(columns=renames)
        return df

    # Helpers para VAR ------------------------------------------------------
    @staticmethod
    def _to_float_br(s: str) -> float | None:
        if s is None:
            return None
        s = str(s).strip()
        if not s:
            return None
        # remove separadores de milhar (ponto) e troca vírgula por ponto
        s = s.replace(".", "").replace(",", ".")
        try:
            return float(s)
        except ValueError:
            return None

    @staticmethod
    def _mk_var_code(descricao: str) -> str:
        # código derivado da descrição (sem acentos/espacos -> hífen simples)
        base = re.sub(r"\s+", "-", str(descricao).strip().upper())
        base = base.replace("Á","A").replace("Â","A").replace("Ã","A").replace("À","A")\
                   .replace("É","E").replace("Ê","E").replace("Í","I")\
                   .replace("Ó","O").replace("Ô","O").replace("Õ","O")\
                   .replace("Ú","U").replace("Ç","C")
        return base

    def _processar_pdf_var(self, caminho_pdf: Path) -> pd.DataFrame | None:
        # extrai VAR no formato longo a partir do leitor novo
        dfv = extrair_var_dataset(str(caminho_pdf), set_arquivo=None, incluir_termo=True)
        if dfv is None or dfv.empty:
            return None

        dfv = dfv.copy()
        # datas / competencia
        dfv["data_medicao"] = pd.to_datetime(dfv["data_medicao"], errors="coerce", dayfirst=True)
        dfv["competencia"] = dfv["data_medicao"].dt.strftime("%Y-%m")

        # textos
        for c in ["contrato","boletim","var_code","descrição","termo","competencia"]:
            if c in dfv.columns:
                dfv[c] = dfv[c].astype("string").str.strip()

        # numéricos
        for c in ["US","qtde","valor_us"]:
            if c in dfv.columns:
                dfv[c] = pd.to_numeric(dfv[c], errors="coerce").astype("float64")

        # alias do valor US do cabeçalho
        dfv["us_global"] = dfv["valor_us"].astype("float64")

        # chave de dedupe para VAR (pode manter contrato aqui; não atrapalha)
        dfv["_chave"] = (
            dfv["boletim"].fillna("").astype(str) + "|" +
            dfv["competencia"].fillna("").astype(str) + "|" +
            dfv["var_code"].fillna("").astype(str) + "|" +
            dfv["US"].fillna(pd.NA).astype(str) + "|" +
            dfv["qtde"].fillna(pd.NA).astype(str)
        ).astype("string")

        dfv = dfv.drop_duplicates(subset=["_chave"]).reset_index(drop=True)

        # ordenação padrão de colunas
        col_order = [
            "contrato","boletim","data_medicao","valor_us","var_code","termo","descrição","US","qtde",
            "competencia","us_global","_chave"
        ]
        cols = [c for c in col_order if c in dfv.columns] + [c for c in dfv.columns if c not in col_order]
        return dfv[cols]

    def _persistir_boletim_var(self, df_var: pd.DataFrame):
        if df_var is None or df_var.empty:
            return
        # concatena com existente (se houver) e deduplica por _chave
        if self.caminho_var.exists():
            atual = pd.read_parquet(self.caminho_var)
            unidos = pd.concat([atual, df_var], ignore_index=True)
        else:
            unidos = df_var.copy()

        # Garante tipos mínimos
        for c in ["contrato","boletim","var_code","descrição","termo","competencia","_chave"]:
            if c in unidos.columns:
                unidos[c] = unidos[c].astype("string")
        for c in ["data_medicao"]:
            if c in unidos.columns:
                unidos[c] = pd.to_datetime(unidos[c], errors="coerce")
        for c in ["US","qtde","valor_us","us_global"]:
            if c in unidos.columns:
                unidos[c] = pd.to_numeric(unidos[c], errors="coerce").astype("float64")

        # Dedupe global
        if "_chave" in unidos.columns:
            unidos = unidos.drop_duplicates(subset=["_chave"])
        else:
            unidos = unidos.drop_duplicates()

        # Salva
        unidos.to_parquet(self.caminho_var, index=False, compression="snappy")


    def _persistir_relacao_contrato_boletim(self, contrato: str, boletim: str, data_medicao):
        import csv
        self.caminho_relacao.parent.mkdir(parents=True, exist_ok=True)
        cabe = ["contrato","boletim","data_medicao"]
        linhas = []
        if self.caminho_relacao.exists():
            with self.caminho_relacao.open("r", newline="", encoding="utf-8") as f:
                rdr = csv.DictReader(f)
                for r in rdr:
                    linhas.append(r)
        novo = {"contrato": str(contrato).strip(), "boletim": str(boletim).strip(),
                "data_medicao": pd.to_datetime(data_medicao, errors="coerce", dayfirst=True).date().isoformat() if data_medicao else ""}
        if not any((r["contrato"], r["boletim"]) == (novo["contrato"], novo["boletim"]) for r in linhas):
            linhas.append(novo)
        with self.caminho_relacao.open("w", newline="", encoding="utf-8") as f:
            w = csv.DictWriter(f, fieldnames=cabe)
            w.writeheader()
            w.writerows(linhas)


    def _carregar_existente(self) -> Tuple[pd.DataFrame, Set[str]]:
        if not self.caminho_parquet.exists():
            return pd.DataFrame(), set()
        df = pd.read_parquet(self.caminho_parquet)
        df = self._padronizar_tipos(df)
        df = self._padronizar_rotulos_horas(df)
        df["chave_unica"] = self._criar_chave_unica(df)
        df = df.drop_duplicates(subset=["chave_unica"])

        # NOVO (correto): atualizar relação contrato×boletim a partir do DF existente
        if {"Contrato", "BOLETIM", "Data de Medição"}.issubset(df.columns):
            rel = (
                df[["Contrato", "BOLETIM", "Data de Medição"]]
                .dropna(subset=["BOLETIM"])
                .drop_duplicates()
            )
            for _, r in rel.iterrows():
                self._persistir_relacao_contrato_boletim(
                    r["Contrato"], r["BOLETIM"], r["Data de Medição"]
                )
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
        # NOVO: já normaliza rótulos de horas nos dados recém-extraídos
        df = self._padronizar_rotulos_horas(df)
        df = df.dropna(subset=["DATA", "Funcionário"])
        df["chave_unica"] = self._criar_chave_unica(df)
        df_novos = df[~df["chave_unica"].isin(chaves_existentes)].copy()
        return df_novos if not df_novos.empty else None
    
    def consolidar_pdfs_em_parquet(self, callback_total=None, callback_arquivo=None) -> str:
        df_existente, chaves = self._carregar_existente()
        arquivos = self._buscar_pdfs()

        # --- barra 1: total de arquivos ---
        total_arquivos = len(arquivos)
        if total_arquivos == 0 and df_existente.empty:
            raise ValueError(f"Nenhum arquivo PDF encontrado na pasta '{self.pasta_raiz}'.")

        novos, novos_var = [], []

        for i, pdf in enumerate(arquivos, start=1):
            # --- barra 1: avança a porcentagem total ---
            if callback_total:
                try:
                    pct_total = (i / total_arquivos) * 100 if total_arquivos else 100
                    callback_total(pct_total)
                except Exception:
                    pass

            # --- barra 2: porcentagem por arquivo (0..100) + nome opcional ---
            if callback_arquivo:
                try:
                    callback_arquivo(0, pdf.name)  # inicia em 0% para este arquivo
                except TypeError:
                    try:
                        callback_arquivo(0)
                    except Exception:
                        pass
                except Exception:
                    pass

            # processa BOLETIM (horas) — mantém seu fluxo atual
            df_novo = self._processar_pdf(pdf, chaves, callback_arquivo)
            if df_novo is not None and not df_novo.empty:
                novos.append(df_novo)
                # atualiza conjunto de chaves para evitar duplicatas
                if "chave_unica" in df_novo.columns:
                    chaves.update(df_novo["chave_unica"].astype(str))

            # processa VAR (US por termo no formato longo com colunas novas)
            df_var = self._processar_pdf_var(pdf)
            if df_var is not None and not df_var.empty:
                novos_var.append(df_var)

            # --- barra 2: marca 100% ao concluir este arquivo ---
            if callback_arquivo:
                try:
                    callback_arquivo(100, pdf.name)
                except TypeError:
                    try:
                        callback_arquivo(100)
                    except Exception:
                        pass
                except Exception:
                    pass

        # se nada novo e nada existente → sinaliza vazio
        if not novos and df_existente.empty:
            raise ValueError("Nenhum dado novo e nada existente.")

        # consolida boletim
        todos = pd.concat([df_existente] + novos, ignore_index=True) if novos else df_existente
        if "chave_unica" in todos.columns:
            todos = todos.drop_duplicates(subset=["chave_unica"])
            # remove registros inválidos (ex.: chave muito curta)
            todos = todos[todos["chave_unica"].astype(str).str.len() > 6]
            todos = todos.drop(columns=["chave_unica"], errors="ignore")
        todos.to_parquet(self.caminho_parquet, index=False, compression="snappy")

        # persiste VAR
        if novos_var:
            df_var_all = pd.concat(novos_var, ignore_index=True)
            self._persistir_boletim_var(df_var_all)

        return str(self.caminho_parquet)

# Interface compatível
def consolidar_pdfs_em_excel(pasta_raiz, set_total=None, set_arquivo=None):
    return ProcessadorPDF(pasta_raiz).consolidar_pdfs_em_parquet(set_total, set_arquivo)
