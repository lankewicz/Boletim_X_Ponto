# utils/comparacao.py
# Núcleo da Comparação Boletim x Ponto (sem dependências de UI)
from __future__ import annotations

import math
from typing import Dict, Iterable, List, Tuple
import pandas as pd

from utils.constantes import HEADERS_VIZ  # ordem canônica das colunas “visuais”
from utils.dataframe_utils import (
    preparar_df_boletim_para_comparacao,
    preparar_df_ponto_para_comparacao,
    resolver_base_ponto,
    groupby_sum_by_date,
)

__all__ = [
    "montar_triplet_comparacao",
    "dfs_sem_ponto",
    "_format_df",
    "ensure_boletim",
    "map_boletim_por_data",
    "montar_tres_grids",
]


# =============================
# Utilidades internas de formato
# =============================

def _decimal_to_hhmm(v: float | int | None) -> str:
    """Converte horas decimais para 'HH:MM'. Aceita negativos."""
    if v is None or (isinstance(v, float) and math.isnan(v)):
        return ""
    try:
        neg = v < 0
        v = abs(float(v))
        h = int(v)
        m = int(round((v - h) * 60))
        # normaliza 60 minutos -> +1 hora
        if m >= 60:
            h += 1
            m -= 60
        s = f"{h:d}:{m:02d}"
        return f"-{s}" if neg else s
    except Exception:
        return str(v)


def _format_df(df: pd.DataFrame, exibir_hhmm: bool) -> pd.DataFrame:
    """
    Formata um DF (cópia) para exibição:
    - Garante 'Data' como datetime (floor D).
    - Converte colunas de HEADERS_VIZ para texto HH:MM se exibir_hhmm=True.
    - Caso contrário mantém valor numérico (float) ou string original.
    """
    if df is None or df.empty:
        return pd.DataFrame(columns=["Data"] + (["Boletim"] if "Boletim" in (df.columns if df is not None else []) else []) + HEADERS_VIZ)

    out = df.copy()
    if "Data" in out.columns:
        out["Data"] = pd.to_datetime(out["Data"], errors="coerce").dt.floor("D")

    for col in HEADERS_VIZ:
        if col in out.columns:
            if exibir_hhmm:
                out[col] = out[col].apply(_decimal_to_hhmm)
            else:
                # numérico “limpo”
                out[col] = pd.to_numeric(out[col], errors="coerce").round(2)
    return out


def ensure_boletim(df: pd.DataFrame) -> pd.DataFrame:
    """Garante que exista a coluna 'Boletim' (preenche com strings vazias se não houver)."""
    if df is None or df.empty:
        return pd.DataFrame(columns=["Data", "Boletim"] + HEADERS_VIZ)
    if "Boletim" in df.columns:
        return df
    out = df.copy()
    out["Boletim"] = ""
    # Mantém ordem: Data | Boletim | HEADERS (presentes)
    cols = ["Data", "Boletim"] + [c for c in HEADERS_VIZ if c in out.columns]
    return out[cols]


# =====================================
# Extração/mapeamento do número Boletim
# =====================================

def _limpa_boletim(valor) -> str:
    """Limpa '123.0' -> '123' e normaliza para string."""
    if pd.isna(valor):
        return ""
    s = str(valor).strip()
    if s.endswith(".0"):
        try:
            s = str(int(float(s)))
        except Exception:
            pass
    return s


def map_boletim_por_data(
    dados_df: pd.DataFrame,
    funcionario: str,
    data_ini,  # date/datetime
    data_fim,  # date/datetime
) -> Dict[pd.Timestamp, str]:
    """
    Devolve dict {data_normalizada: 'NUM_BOLETIM'} para o período/funcionário.
    Data normalizada = Timestamp (00:00:00).
    """
    if dados_df is None or dados_df.empty:
        return {}

    try:
        di_date = pd.Timestamp(data_ini).date()
    except Exception:
        di_date = data_ini
    try:
        df_date = pd.Timestamp(data_fim).date()
    except Exception:
        df_date = data_fim

    mask = (
        (dados_df.get("Funcionário") == funcionario)
        & (pd.to_datetime(dados_df.get("DATA"), errors="coerce").dt.date >= di_date)
        & (pd.to_datetime(dados_df.get("DATA"), errors="coerce").dt.date <= df_date)
    )
    dfb = dados_df.loc[mask, ["DATA", "BOLETIM"]].copy()
    if dfb.empty:
        return {}

    dfb["DATA"] = pd.to_datetime(dfb["DATA"], errors="coerce").dt.floor("D")
    dfb["BOL_LIMPO"] = dfb["BOLETIM"].map(_limpa_boletim)

    # se houver múltiplos boletins no mesmo dia, pega o primeiro não-vazio
    out = {}
    for d, grupo in dfb.groupby("DATA", sort=True):
        vals = [v for v in grupo["BOL_LIMPO"].tolist() if v]
        out[d] = vals[0] if vals else ""
    return out


# ======================================
# Geração das 3 grades alinhadas por dia
# ======================================

def _row_as_list(row: pd.Series, headers: Iterable[str]) -> List[str]:
    lst = []
    for h in headers:
        v = row.get(h, None)
        if pd.isna(v):
            lst.append("")
        elif isinstance(v, (int, float)):
            # arredonda + corrige "-0,00"
            num = round(float(v), 2)
            if abs(num) < 1e-9:
                num = 0.0
            lst.append(f"{num:.2f}".replace(".", ","))
        else:
            s = str(v)
            # se parecer decimal com ponto, normaliza para vírgula
            if "." in s and "," not in s:
                try:
                    float(s)
                    s = s.replace(".", ",")
                except Exception:
                    pass
            lst.append(s)
    return lst



def montar_tres_grids(
    datas_mes: List[str],
    df_b: pd.DataFrame,
    df_p: pd.DataFrame,
    df_d: pd.DataFrame,
    bol_map: Dict[pd.Timestamp, str] | None = None,
) -> Tuple[List[List[str]], List[List[str]], List[List[str]], List[str]]:
    """
    Recebe 3 DFs 'amigáveis' e monta listas de linhas para as sheets:
    - dados_b: [Data, Boletim?, headers_vis...]
    - dados_p: [Data, headers_vis...]
    - dados_d: [Data, headers_vis...]

    Retorna também 'headers_vis' (ordem baseada em HEADERS_VIZ filtrando as que existem).
    """
    # Normaliza Data como date indexável
    def to_map(df: pd.DataFrame) -> Dict[pd.Timestamp, pd.Series]:
        if df is None or df.empty or "Data" not in df.columns:
            return {}
        dfx = df.copy()
        dfx["Data"] = pd.to_datetime(dfx["Data"], errors="coerce").dt.floor("D")
        return {d: row for d, row in dfx.set_index("Data").iterrows()}

    map_b = to_map(df_b)
    map_p = to_map(df_p)
    map_d = to_map(df_d)

    # Determina headers visíveis a partir da ordem canônica
    presentes = set()
    for df in (df_b, df_p, df_d):
        if df is not None and not df.empty:
            presentes |= set([c for c in df.columns if c in HEADERS_VIZ])
    headers_vis = [h for h in HEADERS_VIZ if h in presentes] or HEADERS_VIZ[:]

    dados_b: List[List[str]] = []
    dados_p: List[List[str]] = []
    dados_d: List[List[str]] = []

    # Converte "dd/mm/aaaa" -> Timestamp do dia para casar com index
    for dstr in datas_mes:
        try:
            dkey = pd.to_datetime(dstr, dayfirst=True, errors="coerce").floor("D")
        except Exception:
            dkey = None

        # --- Boletim
        row_b = map_b.get(dkey, None)
        if row_b is not None:
            linha_b = [dstr]
            # Boletim (se DF já possui)
            if "Boletim" in (df_b.columns if df_b is not None else []):
                bol_val = row_b.get("Boletim", "")
            else:
                # tenta via mapa
                bol_val = (bol_map or {}).get(dkey, "")
            linha_b.append("" if pd.isna(bol_val) else str(bol_val))
            linha_b += _row_as_list(row_b, headers_vis)
        else:
            # linha vazia
            linha_b = [dstr]
            bol_val = (bol_map or {}).get(dkey, "") if bol_map else ""
            linha_b.append(bol_val or "")
            linha_b += [""] * len(headers_vis)
        dados_b.append(linha_b)

        # --- Ponto
        row_p = map_p.get(dkey, None)
        if row_p is not None:
            linha_p = [dstr] + _row_as_list(row_p, headers_vis)
        else:
            linha_p = [dstr] + [""] * len(headers_vis)
        dados_p.append(linha_p)

        # --- Diferença
        row_d = map_d.get(dkey, None)
        if row_d is not None:
            linha_d = [dstr] + _row_as_list(row_d, headers_vis)
        else:
            linha_d = [dstr] + [""] * len(headers_vis)
        dados_d.append(linha_d)

    return dados_b, dados_p, dados_d, headers_vis


# ====================================
# Fallback quando NÃO há base de ponto
# ====================================

def dfs_sem_ponto(di_date, df_date, headers_viz: Iterable[str]) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """
    Gera 2 DataFrames alinhados por dia quando NÃO há base de ponto:
      - df_p_f: Data | <headers> (valores None -> exibem vazio)
      - df_d  : Data | <headers> (valores 0.0 -> diferença zerada)
    """
    base = pd.date_range(pd.Timestamp(di_date), pd.Timestamp(df_date), freq="D")
    df_p_f = pd.DataFrame({"Data": base})
    for h in headers_viz:
        df_p_f[h] = None

    df_d = pd.DataFrame({"Data": base})
    for h in headers_viz:
        df_d[h] = 0.0

    return df_p_f, df_d


# ======================================
# Núcleo lógico: montar 3 DFs “amigáveis”
# ======================================

def montar_triplet_comparacao(
    app,
    funcionario: str,
    data_ini,   # date/datetime
    data_fim,   # date/datetime
) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, str | None, set, bool]:
    """
    Constrói os 3 DFs lógicos para a aba/relatório de Comparação:

      - df_b_f (Boletim “amigável”: Data | [Boletim] | HEADERS_VIZ)
      - df_p_f (Ponto   “amigável”: Data | HEADERS_VIZ; vazio se “sem ponto”)
      - df_d   (Diferença:          Data | HEADERS_VIZ; zerado se “sem ponto”)

    Retorna ainda:
      registro (str | None), boletins_set (set[str]), sem_ponto (bool)
    """
    # --- normaliza limites para comparar com .dt.date
    try:
        di_date = pd.Timestamp(data_ini).date()
    except Exception:
        di_date = data_ini if hasattr(data_ini, "year") else None
    try:
        df_date = pd.Timestamp(data_fim).date()
    except Exception:
        df_date = data_fim if hasattr(data_fim, "year") else None

    # ----- BOLETIM: filtra período/funcionário
    df_b_base = app.dados_df[
        (app.dados_df["Funcionário"] == funcionario)
        & (app.dados_df["DATA"].dt.date >= di_date)
        & (app.dados_df["DATA"].dt.date <= df_date)
    ].copy()

    if df_b_base.empty:
        # Sem boletim no período — devolve tudo vazio
        return (
            pd.DataFrame(columns=["Data", "Boletim"] + HEADERS_VIZ),
            pd.DataFrame(columns=["Data"] + HEADERS_VIZ),
            pd.DataFrame(columns=["Data"] + HEADERS_VIZ),
            None,
            set(),
            True,
        )

    # Metadados (registro / boletins do período)
    registro = None
    if "Registro" in df_b_base.columns:
        regs = df_b_base["Registro"].dropna().astype(str).str.strip()
        regs = regs[regs != ""]
        if not regs.empty:
            registro = regs.iloc[0]

    boletins_set: set = set()
    if "BOLETIM" in df_b_base.columns:
        bol = (
            df_b_base["BOLETIM"]
            .dropna().astype(str).str.strip()
            .str.replace(r"\.0+$", "", regex=True)
        )
        boletins_set = {b for b in bol.tolist() if b}

    # Normaliza e agrega boletim por Data
    df_b_norm = preparar_df_boletim_para_comparacao(df_b_base)
    df_b = groupby_sum_by_date(
        df_b_norm.rename(columns={"DATA": "Data"}) if "DATA" in df_b_norm.columns else df_b_norm,
        "Data",
    )
    # Cabeçalhos amigáveis (MAP_BOL está em app; aqui convertemos só o que já veio normalizado)
    df_b_f = df_b.rename(
        columns={k: v for k, v in app.MAP_BOL.items() if k in df_b.columns}
    ).copy()
    if "Data" in df_b_f.columns:
        df_b_f["Data"] = pd.to_datetime(df_b_f["Data"], errors="coerce").dt.floor("D")

    # ----- PONTO: resolver relação; tratar “sem ponto”
    sem_ponto = False
    if app.df_ponto is None or app.df_ponto.empty:
        sem_ponto = True
        df_p_f, df_d = dfs_sem_ponto(di_date, df_date, HEADERS_VIZ)
    else:
        df_ponto_sel, _, _ = resolver_base_ponto(
            app.df_ponto, app.df_relacao, funcionario, data_ini, data_fim, corte_similaridade=0.75
        )
        if df_ponto_sel is None or df_ponto_sel.empty:
            sem_ponto = True
            df_p_f, df_d = dfs_sem_ponto(di_date, df_date, HEADERS_VIZ)
        else:
            # 1) prepara base do ponto por dia
            df_p = preparar_df_ponto_para_comparacao(df_ponto_sel, data_ini, data_fim)

            # 2) renomeia do nome real -> visual (HEADERS_VIZ)
            inv_map_pto = {v: k for k, v in app.MAP_PTO.items()}
            df_p_f = df_p.rename(
                columns={k: inv_map_pto[k] for k in df_p.columns if k in inv_map_pto}
            ).copy()
            if "Data" in df_p_f.columns:
                df_p_f["Data"] = pd.to_datetime(df_p_f["Data"], errors="coerce").dt.floor("D")

            # 3) Diferença (B - P)
            df_m = pd.merge(df_b_f, df_p_f, on="Data", how="outer", suffixes=("_B", "_P")).sort_values("Data")
            for h in HEADERS_VIZ:
                if f"{h}_B" not in df_m.columns: df_m[f"{h}_B"] = 0.0
                if f"{h}_P" not in df_m.columns: df_m[f"{h}_P"] = 0.0
                df_m[f"{h}_B"] = pd.to_numeric(df_m[f"{h}_B"], errors="coerce").fillna(0.0)
                df_m[f"{h}_P"] = pd.to_numeric(df_m[f"{h}_P"], errors="coerce").fillna(0.0)
                df_m[h] = df_m[f"{h}_B"] - df_m[f"{h}_P"]
            df_d = df_m[["Data"] + [h for h in HEADERS_VIZ if h in df_m.columns]].copy()

    # --- arredonda DIFERENÇA para 2 casas e evita "-0.00"
    for _h in HEADERS_VIZ:
        if _h in df_d.columns:
            s = pd.to_numeric(df_d[_h], errors="coerce").round(2)
            # zera valores muito pequenos para não aparecer "-0,00"
            s = s.where(s.abs() >= 5e-4, 0.0)
            df_d[_h] = s


    # Garante colunas/ordem finais
    keep_b = (["Data"] + (["Boletim"] if "Boletim" in df_b_f.columns else [])
              + [h for h in HEADERS_VIZ if h in df_b_f.columns])
    df_b_f = df_b_f[keep_b].copy() if "Data" in df_b_f.columns else pd.DataFrame(
        columns=["Data", "Boletim"] + HEADERS_VIZ
    )

    keep_p = ["Data"] + [h for h in HEADERS_VIZ if h in df_p_f.columns]
    df_p_f = df_p_f[keep_p].copy() if "Data" in df_p_f.columns else pd.DataFrame(
        columns=["Data"] + HEADERS_VIZ
    )

    # df_d já foi montado usando HEADERS_VIZ; garante ordem
    keep_d = ["Data"] + [h for h in HEADERS_VIZ if h in df_d.columns]
    df_d = df_d[keep_d].copy() if "Data" in df_d.columns else pd.DataFrame(
        columns=["Data"] + HEADERS_VIZ
    )

    return df_b_f, df_p_f, df_d, registro, boletins_set, sem_ponto
