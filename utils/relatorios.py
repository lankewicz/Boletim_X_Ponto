# utils/relatorios.py
from __future__ import annotations
import pandas as pd
import tkinter as tk
from tkinter import ttk
import ttkbootstrap as ttkb
from datetime import datetime
from typing import Iterable, List, Dict, Any
import sys
from pathlib import Path
import re


_ARQ_PAD = Path(__file__).resolve().parents[1] / "data" / "RegistroBoletim.parquet"
ARQUIVO_PADRAO = str(_ARQ_PAD)

# Colunas a ignorar quando somando "por Funcionário"
IGNORAR_COLS_DEFAULT = {
    "Data", "DATA",
    "Boletim", "BOLETIM",
    "Registro", "REGISTRO",
    "Contrato", "CONTRATO",
    "Arquivo", "ARQUIVO",
    "Fornecedor", "FORNECEDOR",
}

# ---------------------------------------------------------------------
# Conversão numérica robusta
# ---------------------------------------------------------------------
def _to_number_series(s: pd.Series) -> pd.Series:
    """
    Converte strings com milhar/decimal em número:
    - remove espaços e NBSP
    - remove pontos de milhar (ex.: 1.234,56 -> 1234,56)
    - troca vírgula decimal por ponto
    """
    s = s.astype(str).str.strip().str.replace("\u00A0", "", regex=False)
    s = s.str.replace(r"\.(?=\d{3}(?:\D|$))", "", regex=True)  # ponto de milhar
    s = s.str.replace(",", ".", regex=False)                   # vírgula decimal
    return pd.to_numeric(s, errors="coerce")

def _coerce_numeric(df: pd.DataFrame, numeric_only_cols: list[str] | None = None) -> pd.DataFrame:
    """Converte colunas para numéricas (quando possível) sem quebrar strings."""
    work = df.copy()
    cols = numeric_only_cols or [c for c in work.columns if c != "Funcionário"]
    for c in cols:
        if c in work.columns and c != "Funcionário":
            s = work[c]
            work[c] = _to_number_series(s) if s.dtype == object else pd.to_numeric(s, errors="coerce")
    return work

# ---------------------------------------------------------------------
# Totalização por Funcionário
# ---------------------------------------------------------------------



def carregar_boletim_parquet(caminho: str = ARQUIVO_PADRAO) -> pd.DataFrame:
    return pd.read_parquet(caminho)

def gerar_df_sa_por_registro_periodo(
    caminho: str = ARQUIVO_PADRAO, data_ini=None, data_fim=None
) -> pd.DataFrame:
    """
    Lê o parquet e retorna o DF de S.A. agregado por
    (Funcionário, Registro, Boletim, Contrato) no período.
    """
    df = carregar_boletim_parquet(caminho)
    return relatorio_sa_por_registro_periodo(df, data_ini, data_fim)


def totalizar_por_funcionario(
    df: pd.DataFrame,
    *,
    coluna_func: str = "Funcionário",
    ignorar_cols: set[str] | None = None,
    ordenar_por_nome: bool = True,
) -> pd.DataFrame:
    """
    Soma todas as colunas numéricas por Funcionário.
    Mantém apenas [Funcionário] + colunas numéricas úteis.
    """
    if df is None or df.empty:
        return pd.DataFrame(columns=[coluna_func])

    ignorar = set(ignorar_cols or IGNORAR_COLS_DEFAULT)
    if coluna_func not in df.columns:
        raise KeyError(f"Coluna '{coluna_func}' não encontrada no DataFrame.")

    # colunas candidatas a soma (evita chaves/ids)
    candidatas = [c for c in df.columns if c not in ignorar and c != coluna_func]
    work = _coerce_numeric(df[[coluna_func] + candidatas], numeric_only_cols=candidatas)

    num_cols = [c for c in candidatas if pd.api.types.is_numeric_dtype(work[c])]
    if not num_cols:
        return pd.DataFrame(columns=[coluna_func])

    agrupado = work.groupby(coluna_func, dropna=False, sort=ordenar_por_nome)[num_cols].sum().reset_index()

    if ordenar_por_nome:
        agrupado = agrupado.sort_values(coluna_func, kind="stable", ignore_index=True)

    # adiciona linha TOTAL GERAL (se não quiser, remova as 2 linhas abaixo)
    total = agrupado[num_cols].sum(numeric_only=True)
    total_row = pd.DataFrame([{coluna_func: "TOTAL GERAL", **total.to_dict()}])
    return pd.concat([agrupado, total_row], ignore_index=True)


# === RELATÓRIO S.A. (Funcionário | Registro | BOLETIM | Contrato) ============



_COLS_SA = ["Funcionário", "Registro", "BOLETIM", "Contrato", "S.A."]

def _forcar_tipos_sa(df: pd.DataFrame) -> pd.DataFrame:
    """
    Garante colunas e tipos corretos para somatório de S.A.
    - aceita 'Boletim' e normaliza para 'BOLETIM'
    - converte S.A. para numérico (suporta milhar com ponto e decimal com vírgula)
    """
    work = df.copy()

    # aliases -> nomes padrão
    if "Boletim" in work.columns and "BOLETIM" not in work.columns:
        work["BOLETIM"] = work["Boletim"]

    for c in ["Funcionário", "Registro", "BOLETIM", "Contrato"]:
        if c not in work.columns:
            work[c] = ""
        work[c] = work[c].astype(str).fillna("")

    # S.A. numérico robusto
    if "S.A." not in work.columns:
        work["S.A."] = 0.0
    s = work["S.A."].astype(str).str.strip().str.replace("\u00A0", "", regex=False)
    s = s.str.replace(r"\.(?=\d{3}(?:\D|$))", "", regex=True)  # 1.234,56 -> 1234,56
    s = s.str.replace(",", ".", regex=False)
    work["S.A."] = pd.to_numeric(s, errors="coerce").fillna(0.0)

    return work[_COLS_SA].copy()

def relatorio_sa_por_registro(df: pd.DataFrame) -> pd.DataFrame:
    """
    Retorna um DataFrame pronto para exibição contendo:
      Funcionário | Registro | BOLETIM | Contrato | S.A.

    Regras:
    - Soma S.A. por (Funcionário, Registro, BOLETIM, Contrato)
    - Ordena por Funcionário, BOLETIM, Registro, Contrato (estável)
    - Insere uma linha "Subtotal <NOME>" APENAS quando o funcionário
      aparece em 2+ boletins distintos.
    """
    if df is None or df.empty:
        return pd.DataFrame(columns=_COLS_SA)

    base = _forcar_tipos_sa(df)

    grouped = (
        base.groupby(["Funcionário", "Registro", "BOLETIM", "Contrato"], as_index=False)["S.A."]
        .sum()
        .sort_values(["Funcionário", "BOLETIM", "Registro", "Contrato"], kind="stable")
        .reset_index(drop=True)
    )

    # monta linhas + subtotais sob demanda
    linhas = []
    for nome, bloco in grouped.groupby("Funcionário", sort=False):
        # linhas do funcionário
        linhas.extend(bloco.to_dict(orient="records"))
        # subtotal só se houver 2+ boletins
        if bloco["BOLETIM"].nunique() > 1:
            linhas.append({
                "Funcionário": f"Subtotal {nome}",
                "Registro": "",
                "BOLETIM": "",
                "Contrato": "",
                "S.A.": float(bloco["S.A."].sum()),
            })

    return pd.DataFrame(linhas, columns=_COLS_SA)

# === UI opcional: diálogo Tk para visualizar o Relatório S.A. =================
# (colocada aqui a pedido, para evitar um arquivo em `interface/`)


_COLS_SA_UI = ["Funcionário", "Registro", "Boletim", "Contrato", "S.A."]

def _fmt_num_sa(v):
    try:
        return f"{float(v):.2f}".replace(".", ",")
    except Exception:
        return str(v)

class RelatorioSADialog(ttkb.Toplevel):
    def __init__(self, parent, df: pd.DataFrame, *, data_ini=None, data_fim=None):
        super().__init__(parent)
        self.title("Relatório S.A. — Funcionário / Registro / Boletim / Contrato")
        self.geometry("900x540")
        self.minsize(760, 420)
        self.transient(parent)
        self.grab_set()

        topo = ttkb.Frame(self, padding=(10, 8))
        topo.pack(fill="x")

        periodo_txt = ""
        if data_ini and data_fim:
            try:
                di = pd.to_datetime(data_ini).strftime("%d/%m/%Y")
                df_ = pd.to_datetime(data_fim).strftime("%d/%m/%Y")
                periodo_txt = f"Período: {di} a {df_}"
            except Exception:
                pass

        ttkb.Label(topo, text=periodo_txt, bootstyle="secondary").pack(side="left")
        ttkb.Label(
            topo,
            text="Agrupamento: Funcionário → Registro → Boletim → Contrato  |  Métrica: S.A.",
            bootstyle="secondary"
        ).pack(side="left", padx=12)
        ttkb.Label(topo, text=f"Gerado em: {datetime.now():%d/%m/%Y %H:%M}",
                   bootstyle="secondary").pack(side="right")

        frame_tbl = ttkb.Frame(self, padding=(10, 0))
        frame_tbl.pack(fill="both", expand=True)

        self.tree = ttk.Treeview(frame_tbl, columns=_COLS_SA_UI, show="headings")
        vsb = ttkb.Scrollbar(frame_tbl, orient="vertical", command=self.tree.yview)
        hsb = ttkb.Scrollbar(frame_tbl, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        frame_tbl.rowconfigure(0, weight=1)
        frame_tbl.columnconfigure(0, weight=1)

        for c in _COLS_SA_UI:
            self.tree.heading(c, text=c, command=lambda col=c: self._ordenar_por(col, False))

        self._popular(df)

        rodape = ttkb.Frame(self, padding=(10, 8))
        rodape.pack(fill="x")
        ttkb.Button(rodape, text="Fechar", command=self.destroy,
                    bootstyle="secondary").pack(side="right")

    def _popular(self, df: pd.DataFrame):
        df = df.rename(columns={"BOLETIM": "Boletim"}).copy()
        for c in _COLS_SA_UI:
            if c not in df.columns:
                df[c] = ""

        for iid in self.tree.get_children():
            self.tree.delete(iid)

        for _, row in df.iterrows():
            vals = [
                row["Funcionário"],
                row["Registro"],
                row["Boletim"],
                row["Contrato"],
                _fmt_num_sa(row["S.A."]),
            ]
            item = self.tree.insert("", "end", values=vals)
            if str(row["Funcionário"]).startswith("Subtotal "):
                self.tree.item(item, tags=("subtotal",))
        self.tree.tag_configure("subtotal", foreground="blue")

        self._ajustar_larguras()

    def _ajustar_larguras(self):
        minw = {"Funcionário": 260, "Registro": 90, "Boletim": 110, "Contrato": 110, "S.A.": 90}
        for c in _COLS_SA_UI:
            maxlen = len(c)
            for k, iid in enumerate(self.tree.get_children()):
                if k > 500:  # limite de amostragem
                    break
                txt = str(self.tree.set(iid, c))
                maxlen = max(maxlen, len(txt))
            width = max(minw.get(c, 80), min(500, int(maxlen * 7.2) + 16))
            self.tree.column(c, width=width, anchor="e" if c == "S.A." else "w")

    def _ordenar_por(self, col: str, reverso: bool):
        dados = [(self.tree.set(k, col), k) for k in self.tree.get_children("")]
        if col == "S.A.":
            def _key(x):
                try:
                    return float(str(x[0]).replace(".", "").replace(",", "."))
                except Exception:
                    return 0.0
        else:
            def _key(x):
                return str(x[0]).casefold()

        dados.sort(key=_key, reverse=reverso)
        for idx, (_, k) in enumerate(dados):
            self.tree.move(k, "", idx)

        self.tree.heading(col, command=lambda: self._ordenar_por(col, not reverso))
# ---------------------------------------------------------------------
# Relatório de Sobreaviso (S.A.) por Funcionário → Contrato → Boletim
# ---------------------------------------------------------------------

def _normalize_min_cols(df: pd.DataFrame) -> pd.DataFrame:
    """
    Normaliza nomes mínimos usados pelo relatório de S.A.
    - DATA -> Data (date puro)
    - BOLETIM -> Boletim
    - S.A -> S.A. (se vier sem ponto)
    """
    work = df.copy()
    rename_map = {}
    if "DATA" in work.columns and "Data" not in work.columns:
        rename_map["DATA"] = "Data"
    if "BOLETIM" in work.columns and "Boletim" not in work.columns:
        rename_map["BOLETIM"] = "Boletim"
    if "S.A" in work.columns and "S.A." not in work.columns:
        rename_map["S.A"] = "S.A."
    if rename_map:
        work = work.rename(columns=rename_map)
    return work

def relatorio_sa_por_func_contrato_boletim(
    df: pd.DataFrame,
    *,
    ordenar_por=("Funcionário", "Contrato", "Boletim"),
    incluir_zeros: bool = False,
) -> pd.DataFrame:
    """
    S.A. por Funcionário → Contrato → Boletim usando a mesma lógica da Aba 1:
      1) normaliza colunas
      2) Data como date puro + S.A. numérico robusto
      3) remove duplicatas exatas
      4) soma por DIA
      5) soma no nível (Funcionário, Contrato, Boletim)
      6) insere '*NOME*' só se houver mais de um boletim/contrato p/ o funcionário
    """
    if df is None or df.empty:
        return pd.DataFrame(columns=["Funcionário", "Contrato", "Boletim", "S.A."])

    work = _normalize_min_cols(df)

    # garante chaves
    for c in ("Funcionário", "Contrato", "Boletim"):
        if c not in work.columns:
            work[c] = ""

    # Data (date puro)
    has_data = "Data" in work.columns
    if has_data:
        work["Data"] = pd.to_datetime(work["Data"], dayfirst=True, errors="coerce").dt.date

    # S.A. numérico robusto
    if "S.A." not in work.columns:
        work["S.A."] = 0.0
    work["S.A."] = _to_number_series(work["S.A."])

    # remove duplicatas exatas
    subset_dup = ["Funcionário", "Contrato", "Boletim", "S.A."]
    if has_data:
        subset_dup.append("Data")
    work = work.drop_duplicates(subset=subset_dup, keep="last")

    # soma por DIA
    if has_data:
        daily = (
            work.groupby(["Funcionário", "Contrato", "Boletim", "Data"], dropna=False)["S.A."]
            .sum()
            .reset_index()
        )
    else:
        daily = work[["Funcionário", "Contrato", "Boletim", "S.A."]].copy()

    # soma no período (F,C,B)
    base = (
        daily.groupby(["Funcionário", "Contrato", "Boletim"], dropna=False)["S.A."]
        .sum()
        .reset_index()
    )

    if not incluir_zeros:
        base = base[base["S.A."].fillna(0) != 0]

    # ordenação
    if ordenar_por:
        base = base.sort_values(list(ordenar_por), kind="stable", ignore_index=True)

    # subtotal '*NOME*' apenas quando houver >1 linha para o funcionário
    partes = []
    for nome, bloco in base.groupby("Funcionário", sort=False):
        bloco = bloco.sort_values(["Contrato", "Boletim"], kind="stable")
        partes.append(bloco)
        if len(bloco) > 1:
            partes.append(pd.DataFrame([{
                "Funcionário": f"*{nome}*",
                "Contrato": "",
                "Boletim": "",
                "S.A.": bloco["S.A."].sum()
            }]))

    saida = pd.concat(partes, ignore_index=True) if partes else base
    return saida[["Funcionário", "Contrato", "Boletim", "S.A."]]

# Alias (se algum lugar ainda importar por este nome)
def totalizar_sa_func_contrato_boletim(df: pd.DataFrame, **kwargs) -> pd.DataFrame:
    return relatorio_sa_por_func_contrato_boletim(df, **kwargs)

# ---------------------------------------------------------------------
# utils/relatorios.py


_COLS_NECS = ["Funcionário", "Registro", "BOLETIM", "Contrato", "S.A.", "DATA"]

def _coerce_sa_one(x) -> float:
    """
    Converte S.A. para float de forma robusta.
    Aceita:
      - número (int/float)
      - '13,54' (vírgula decimal)
      - '13.54' (ponto decimal)
      - '1.234,56' / '1 234,56' (milhar + vírgula decimal)
      - '13:54' (HH:MM  => 13 + 54/60)
    Qualquer valor inválido vira 0.0
    """
    if x is None:
        return 0.0
    # Já numérico
    if isinstance(x, (int, float)) and not pd.isna(x):
        return float(x)

    s = str(x).strip()
    if not s:
      return 0.0
    s = s.replace("\xa0", " ").strip()   # NBSP

    # HH:MM
    m = re.fullmatch(r"\s*(\d{1,3})\s*:\s*(\d{1,2})\s*", s)
    if m:
        hh = int(m.group(1))
        mm = int(m.group(2))
        return float(hh) + (mm / 60.0)

    # remover separadores de milhar (ponto ou espaço) apenas quando fazem sentido
    # 12.345,67 -> 12345,67   |   12 345,67 -> 12345,67
    s = re.sub(r"(?<=\d)[\.\s](?=\d{3}(\D|$))", "", s)

    # vírgula decimal -> ponto
    s = s.replace(",", ".")

    # manter apenas dígitos, sinal e ponto decimal
    s = re.sub(r"[^0-9\.\-]", "", s)
    try:
        return float(s)
    except Exception:
        return 0.0


def _coerce_sa(col: pd.Series) -> pd.Series:
    """Aplica a conversão robusta elemento a elemento."""
    return col.apply(_coerce_sa_one)

def _ensure_cols(df: pd.DataFrame, cols: Iterable[str]) -> pd.DataFrame:
    w = df.copy()
    for c in cols:
        if c not in w.columns:
            w[c] = pd.NA
    return w

def relatorio_sa_por_registro_periodo(
    df: pd.DataFrame,
    data_ini,  # date/datetime/str dd/mm/aaaa ou None
    data_fim,  # date/datetime/str dd/mm/aaaa ou None
) -> pd.DataFrame:
    """
    Soma S.A. por (Funcionário, Registro, BOLETIM, Contrato) apenas no período.
    Adiciona 'Subtotal <NOME>' somente quando o funcionário aparece em 2+ boletins.
    Retorna colunas: Funcionário | Registro | Boletim | Contrato | S.A.
    """
    cols_saida = ["Funcionário", "Registro", "Boletim", "Contrato", "S.A."]
    if df is None or df.empty:
        return pd.DataFrame(columns=cols_saida)

    work = df.copy()

    # Normaliza aliases para nomes-padrão internos
    rename_in = {}
    if "Boletim" in work.columns and "BOLETIM" not in work.columns:
        rename_in["Boletim"] = "BOLETIM"
    if "Data" in work.columns and "DATA" not in work.columns:
        rename_in["Data"] = "DATA"
    if rename_in:
        work = work.rename(columns=rename_in)

    # Garante colunas mínimas
    for c in ["Funcionário", "Registro", "BOLETIM", "Contrato", "S.A.", "DATA"]:
        if c not in work.columns:
            work[c] = pd.NA

    # Datas e filtro de período (inclusive nas bordas)
    work["DATA"] = pd.to_datetime(work["DATA"], dayfirst=True, errors="coerce")

    if isinstance(data_ini, str):
        data_ini = pd.to_datetime(data_ini, dayfirst=True, errors="coerce")
    if isinstance(data_fim, str):
        data_fim = pd.to_datetime(data_fim, dayfirst=True, errors="coerce")

    if data_ini is None or pd.isna(data_ini):
        data_ini = work["DATA"].min()
    if data_fim is None or pd.isna(data_fim):
        data_fim = work["DATA"].max()

    work = work[(work["DATA"] >= data_ini) & (work["DATA"] <= data_fim)].copy()
    if work.empty:
        return pd.DataFrame(columns=cols_saida)

    # Tipos
    work["S.A."] = _coerce_sa(work["S.A."])
    for c in ["Funcionário", "Registro", "BOLETIM", "Contrato"]:
        work[c] = work[c].astype(str).fillna("")

    # (Opcional) Se quiser evitar linhas repetidas 100% iguais, descomente:
    # work = work.drop_duplicates(subset=["Funcionário","Registro","BOLETIM","Contrato","S.A.","DATA"], keep="last")

    # Soma por chave (mesma lógica do console)
    base = (
        work.groupby(["Funcionário", "Registro", "BOLETIM", "Contrato"], as_index=False)["S.A."]
            .sum()
            .sort_values(["Funcionário", "BOLETIM", "Registro", "Contrato"], kind="stable")
            .reset_index(drop=True)
    )

    # Renomeia para a coluna que a GUI usa
    base = base.rename(columns={"BOLETIM": "Boletim"})

    # Linhas + SUBTOTAL somente se houver 2+ boletins distintos
    partes = []
    for nome, g in base.groupby("Funcionário", sort=False):
        partes.append(g)
        if g["Boletim"].nunique() > 1:
            partes.append(pd.DataFrame([{
                "Funcionário": f"Subtotal {nome}",
                "Registro": "",
                "Boletim": "",
                "Contrato": "",
                "S.A.": float(g["S.A."].sum()),
            }]))

    saida = pd.concat(partes, ignore_index=True)
    cols_saida = ["Funcionário", "Registro", "Boletim", "Contrato", "S.A."]
    saida = saida[cols_saida]
    return saida

# mapeamento tolerante para normalizar cabeçalhos
ALIASES_SA = {
    "Funcionário": {"Funcionário", "FUNCIONÁRIO", "FUNCIONARIO"},
    "Registro": {"Registro", "REGISTRO", "Matrícula", "MATRÍCULA", "MATRICULA"},
    "Boletim": {"Boletim", "BOLETIM"},
    "Contrato": {"Contrato", "CONTRATO"},
    "S.A.": {"S.A.", "SA", "S A", "Sobreaviso", "SOBREAVISO"},
    "Data": {"Data", "DATA"},
}

def _normalize_cols_sa(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return df
    rename_map = {}
    cols = set(df.columns)
    for target, aliases in ALIASES_SA.items():
        for a in list(aliases):
            if a in cols:
                rename_map[a] = target
        for a in list(aliases):
            au = str(a).upper()
            if au in cols:
                rename_map[au] = target
    return df.rename(columns=rename_map).copy()

def _to_number_sa(s: pd.Series) -> pd.Series:
    s = s.astype(str).str.strip().str.replace("\u00A0", "", regex=False)
    s = s.str.replace(r"\.(?=\d{3}(?:\D|$))", "", regex=True)  # ponto milhar
    s = s.str.replace(",", ".", regex=False)                   # vírgula decimal
    return pd.to_numeric(s, errors="coerce").fillna(0.0)

def montar_sa_por_registro(
    df_boletim: pd.DataFrame,
    *,
    data_ini,  # datetime.date | datetime.datetime | str dd/mm/yyyy
    data_fim,  # idem
    incluir_subtotal: bool = True,
) -> pd.DataFrame:
    """
    Retorna um DataFrame com as colunas:
      Funcionário | Registro | Boletim | Contrato | S.A.
    Agrupado por (Funcionário, Registro, Boletim, Contrato) com soma do S.A.
    Filtra SOMENTE pelo período recebido.
    """
    if df_boletim is None or df_boletim.empty:
        return pd.DataFrame(columns=["Funcionário", "Registro", "Boletim", "Contrato", "S.A."])

    work = _normalize_cols_sa(df_boletim).copy()

    # garante colunas
    for c in ("Funcionário", "Registro", "Boletim", "Contrato", "S.A."):
        if c not in work.columns:
            work[c] = "" if c != "S.A." else 0.0

    # normaliza data (se existir)
    if "Data" in work.columns:
        work["Data"] = pd.to_datetime(work["Data"], dayfirst=True, errors="coerce")

    # filtro pelo período (somente Data)
    def _to_dt(s):
        import pandas as _pd
        if isinstance(s, str):
            return _pd.to_datetime(s, dayfirst=True, errors="coerce")
        return _pd.to_datetime(s, errors="coerce")
    di = _to_dt(data_ini)
    df = _to_dt(data_fim)
    if "Data" in work.columns and pd.notna(di) and pd.notna(df):
        work = work[(work["Data"] >= di) & (work["Data"] <= df)]

    # S.A. numérico robusto
    work["S.A."] = _to_number_sa(work["S.A."])

    # agrupa por registro (func, registro, boletim, contrato)
    keys = ["Funcionário", "Registro", "Boletim", "Contrato"]
    base = work.groupby(keys, dropna=False)["S.A."].sum().reset_index()

    # ordena por Funcionário, Registro, Boletim, Contrato
    base = base.sort_values(keys, kind="stable", ignore_index=True)

    if not incluir_subtotal:
        return base[["Funcionário", "Registro", "Boletim", "Contrato", "S.A."]]

    # injeta "Subtotal <NOME>" quando houver mais de uma linha para o funcionário
    partes = []
    for nome, bloco in base.groupby("Funcionário", sort=False):
        partes.append(bloco)
        if len(bloco) > 1:
            partes.append(pd.DataFrame([{
                "Funcionário": f"Subtotal {nome}",
                "Registro": "",
                "Boletim": "",
                "Contrato": "",
                "S.A.": bloco["S.A."].sum(),
            }]))
    saida = pd.concat(partes, ignore_index=True)
    return saida[["Funcionário", "Registro", "Boletim", "Contrato", "S.A."]]

#
# === Relatório de Console (script relatorio_registro_console.py) ============
#
# Mostra colunas: Funcionário | Registro | Boletim | Contrato | S.A.
# Regras:
# - Agrega S.A. por (Funcionário, Registro, BOLETIM, Contr
# - Ordena por Funcionário (depois BOLETIM para manter agrupamento estável).
# - Adiciona linha "Subtotal <NOME>" SOMENTE se o funcionário aparecer
#   em DOIS OU MAIS boletins distintos.
# - Usa arquivo padrão ".\data\RegistroBoletim.parquet" se não for passado. 
# - Dá mensagem amigável se faltar 'pyarrow' ou 'fastparquet'.
# - Valida colunas mínimas.


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


def _agrupar_por_registro(df: pd.DataFrame) -> pd.DataFrame:
    """
    Soma S.A. por (Funcionário, Registro, BOLETIM, Contrato) e ordena.
    """
    agrupado = (
        df.groupby(["Funcionário", "Registro", "BOLETIM", "Contrato"], as_index=False)["S.A."]
        .sum()
        .sort_values(["Funcionário", "BOLETIM", "Registro", "Contrato"], kind="stable")
        .reset_index(drop=True)
    )
    return agrupado


def gerar_relatorio_console(caminho_arquivo: str = ARQUIVO_PADRAO):
    df = _carregar_parquet(caminho_arquivo)
    _validar_colunas(df)
    df = _forcar_tipos(df)
    df_grouped = _agrupar_por_registro(df)
    linhas = _montar_linhas_com_subtotais(df_grouped)
    _formatar_e_imprimir(linhas)


def _montar_linhas_com_subtotais(df_grouped: pd.DataFrame) -> List[Dict[str, Any]]:
    """
    Gera a lista de linhas finais, inserindo 'Subtotal <nome>' apenas
    se o funcionário tiver registros em 2+ boletins distintos.
    """
    linhas: List[Dict[str, Any]] = []
    for nome, grupo in df_grouped.groupby("Funcionário", sort=False):
        # Adiciona as linhas do funcionário
        for _, row in grupo.iterrows():
            linhas.append(
                {
                    "Funcionário": row["Funcionário"],
                    "Registro": row["Registro"],
                    "BOLETIM": row["BOLETIM"],
                    "Contrato": row["Contrato"],
                    "S.A.": row["S.A."],
                }
            )

        # Regra do subtotal (somente se houver 2+ boletins)
        if grupo["BOLETIM"].nunique() > 1:
            linhas.append(
                {
                    "Funcionário": f"Subtotal {nome}",
                    "Registro": "",
                    "BOLETIM": "",
                    "Contrato": "",
                    "S.A.": grupo["S.A."].sum(),
                }
            )
    return linhas
