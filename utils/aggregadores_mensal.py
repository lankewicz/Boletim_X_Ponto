# utils/aggregadores_mensal.py
from __future__ import annotations
import pandas as pd
from typing import List, Dict
from utils.constantes import HEADERS_VIZ, MAP_BOL, MAP_PTO
from utils.dataframe_utils import (
    preparar_df_boletim_para_comparacao,
    preparar_df_ponto_para_comparacao,
)

def _mes(col: pd.Series) -> pd.Series:
    s = pd.to_datetime(col, errors="coerce")
    return s.dt.to_period("M").astype(str)

def _renomear_boletim_para_visuais(df_b_sum: pd.DataFrame) -> pd.DataFrame:
    # exemplo: "HORA NORMAL" -> "Horas Normais"
    ren = {src: dst for src, dst in MAP_BOL.items() if src in df_b_sum.columns}
    return df_b_sum.rename(columns=ren)

def _renomear_ponto_para_visuais(df_p_sum: pd.DataFrame) -> pd.DataFrame:
    # exemplo: "Total Normais" -> "Horas Normais"
    inv = {v: k for k, v in MAP_PTO.items()}
    ren = {src: inv[src] for src in df_p_sum.columns if src in inv}
    return df_p_sum.rename(columns=ren)

def comparar_mensal_por_funcionario(
    df_boletim: pd.DataFrame,
    df_ponto: pd.DataFrame,
) -> pd.DataFrame:
    """
    Produz uma tabela mensal mantendo Funcionário, Registro e Contrato (do Boletim).
    Merge por Funcionário+Mês (Ponto não tem Registro/Contrato).
    Blocos: Boletim | Ponto | Diferença (B − P), nas colunas HEADERS_VIZ.
    """
    if df_boletim is None or df_boletim.empty:
        return pd.DataFrame()

    # --- Boletim: normaliza e agrega por mês/func/registro/contrato
    b = preparar_df_boletim_para_comparacao(df_boletim).copy()
    # garante 'Data'
    if "Data" not in b.columns:
        b = b.rename(columns={"DATA": "Data"})
    b["Mes"] = _mes(b["Data"])
    keys_b = ["Funcionário", "Registro", "Contrato", "Mes"]
    num_b: List[str] = [c for c in b.columns if c not in keys_b and pd.api.types.is_numeric_dtype(b[c])]
    b_sum = b.groupby(keys_b, dropna=False)[num_b].sum(numeric_only=True).reset_index()
    b_vis = _renomear_boletim_para_visuais(b_sum)

    # --- Ponto: normaliza e agrega por mês/nome
    p = pd.DataFrame()
    if df_ponto is not None and not df_ponto.empty:
        p_day = preparar_df_ponto_para_comparacao(df_ponto)
        if not p_day.empty:
            p_day["Mes"] = _mes(p_day["Data"])
            # soma mensal por Nome (quando existir)
            nome_col = "Nome" if "Nome" in df_ponto.columns else None
            if nome_col is not None and nome_col in df_ponto.columns:
                # Precisamos do Nome associado à data para somar por mês.
                base = df_ponto.copy()
                base["Data"] = pd.to_datetime(base["Data"], errors="coerce")
                base["Mes"] = base["Data"].dt.to_period("M").astype(str)
                cols_pontos = list(MAP_PTO.values())
                exist = [c for c in cols_pontos if c in base.columns]
                p_sum = base.groupby([nome_col, "Mes"], dropna=False)[exist].sum(numeric_only=True).reset_index()
                p_vis = _renomear_ponto_para_visuais(p_sum)
                p = p_vis.rename(columns={nome_col: "Funcionário"})  # alinhamos pelo nome
            else:
                p = pd.DataFrame()

    # --- Merge por Funcionário+Mes (mantendo Registro/Contrato do Boletim)
    if p.empty:
        out = b_vis.copy()
        for h in HEADERS_VIZ:
            if h not in out.columns:
                out[h] = 0.0
        out_p = out.copy()
        for h in HEADERS_VIZ:
            out_p[h] = 0.0
        merged = out.merge(out_p[["Funcionário", "Mes"] + HEADERS_VIZ], on=["Funcionário", "Mes"], how="left", suffixes=("_B", "_P"))
    else:
        merged = b_vis.merge(
            p[["Funcionário", "Mes"] + [c for c in p.columns if c in HEADERS_VIZ]],
            on=["Funcionário", "Mes"], how="left", suffixes=("_B", "_P")
        )

    # --- Gera blocos Boletim / Ponto / Diferença (B − P)
    for h in HEADERS_VIZ:
        if f"{h}_B" not in merged.columns:
            merged[f"{h}_B"] = 0.0
        if h not in merged.columns:
            merged[h] = 0.0
        # renomeado Ponto
        merged[f"{h}_P"] = pd.to_numeric(merged.get(h), errors="coerce").fillna(0.0)
        merged[f"{h}_B"] = pd.to_numeric(merged[f"{h}_B"], errors="coerce").fillna(0.0)

    # monta saída final
    cols_ident = ["Funcionário", "Registro", "Contrato", "Mes"]
    cols_b = [f"{h} (Boletim)" for h in HEADERS_VIZ]
    cols_p = [f"{h} (Ponto)" for h in HEADERS_VIZ]
    cols_d = [f"{h} (Dif. B−P)" for h in HEADERS_VIZ]

    saida = merged[cols_ident].copy()
    for h in HEADERS_VIZ:
        saida[f"{h} (Boletim)"] = merged[f"{h}_B"].round(2)
    for h in HEADERS_VIZ:
        saida[f"{h} (Ponto)"] = merged[f"{h}_P"].round(2)
    for h in HEADERS_VIZ:
        saida[f"{h} (Dif. B−P)"] = (merged[f"{h}_B"] - merged[f"{h}_P"]).round(2)

    # ordena por Funcionário (estável), depois Mes, Contrato, Registro
    saida = saida.sort_values(["Funcionário", "Mes", "Contrato", "Registro"], kind="stable", ignore_index=True)
    return saida
