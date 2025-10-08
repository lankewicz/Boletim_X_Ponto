# interface/relatorios.py
from __future__ import annotations
import pandas as pd

from utils.periodo import datas_formatadas_do_mes
from utils.comparacao import (
    montar_triplet_comparacao,  # já monta df_b_f, df_p_f, df_d (e trata “sem ponto”)
    _format_df,
    ensure_boletim,
    map_boletim_por_data,
    montar_tres_grids,
)
from utils.dataframe_utils import resolver_base_ponto
from datetime import datetime as dt
# +++ Relatório Geral (sobreavisos) +++
from utils.relatorios import (
    carregar_boletim_parquet,
    relatorio_sa_por_registro_periodo,
    RelatorioSADialog,
    ARQUIVO_PADRAO,
)
import tkinter as tk
import pandas as pd


# ---------- Relatório: COMPARAÇÃO (dados para as 3 grids) ----------
def build_comparacao_grids(app, funcionario: str, di, df):
    """
    Retorna:
      dados_b, dados_p, dados_d, headers_vis, registro, boletins_set, sem_ponto
    Pronto para alimentar sheet_comp_boletim / sheet_comp_ponto / sheet_comp_diferenca.
    """
    data_ini = pd.Timestamp(dt.combine(di, dt.min.time()))
    data_fim = pd.Timestamp(dt.combine(df, dt.max.time()))
    datas_mes = datas_formatadas_do_mes(di, df)

    # 1) Triplet lógico (já trata "sem ponto")
    df_b_f, df_p_f, df_d, registro, boletins_set, sem_ponto = montar_triplet_comparacao(
        app, funcionario, data_ini, data_fim
    )

    # 2) Formatação (decimal ↔ HH:MM) e grades por calendário
    exibir_hhmm = not app.modo_exibicao_decimal
    df_b_v = ensure_boletim(_format_df(df_b_f, exibir_hhmm))
    df_p_v = _format_df(df_p_f, exibir_hhmm)
    df_d_v = _format_df(df_d, exibir_hhmm)

    bol_map = map_boletim_por_data(app.dados_df, funcionario, di, df)
    dados_b, dados_p, dados_d, headers_vis = montar_tres_grids(
        datas_mes, df_b_v, df_p_v, df_d_v, bol_map
    )
    return dados_b, dados_p, dados_d, headers_vis, registro, boletins_set, sem_ponto


# ---------- Relatório: PONTO (uma grid) ----------
def build_ponto_grid(app, funcionario: str, di, df):
    """
    Resolve base do ponto e devolve:
      dados_lista, ident, origem, df_ponto_filtrado
    Caso não haja ponto: dados vazios e df_ponto_filtrado vazio (sem messagebox).
    """
    data_ini = pd.Timestamp(dt.combine(di, dt.min.time()))
    data_fim = pd.Timestamp(dt.combine(df, dt.max.time()))

    if app.df_ponto is None or app.df_ponto.empty or not funcionario:
        return [], None, None, pd.DataFrame()

    df_filtrado, origem, ident = resolver_base_ponto(
        app.df_ponto, app.df_relacao, funcionario, data_ini, data_fim, corte_similaridade=0.75
    )
    if df_filtrado is None or df_filtrado.empty or "Data" not in df_filtrado.columns:
        return [], ident, origem, pd.DataFrame()

    df_filtrado = df_filtrado.sort_values(by="Data", kind="mergesort")
    dados_lista = []
    for _, row in df_filtrado.iterrows():
        linha = [row["Data"].strftime("%d/%m/%Y") if pd.notna(row["Data"]) else ""]
        for col_df in app.colunas_ponto_df[1:]:
            val = row.get(col_df)
            linha.append(app.formatar_tempo(val) if pd.notna(val) else "")
        dados_lista.append(linha)

    # Totais
    totais = ["TOTAIS:"]
    for col_df in app.colunas_ponto_df[1:]:
        soma = df_filtrado[col_df].sum()
        totais.append(app.formatar_tempo(soma))
    dados_lista.append(totais)

    return dados_lista, ident, origem, df_filtrado.copy()


# ---------- Relatório: BOLETIM (uma grid) ----------
def build_boletim_grid(app, funcionario: str, di, df):
    """
    Monta a grid do boletim por dia + totais (sem UI).
    Retorna dados_lista (linhas já formatadas) e df_filtrado (para eventuais somas).
    """
    data_ini = pd.Timestamp(dt.combine(di, dt.min.time()))
    data_fim = pd.Timestamp(dt.combine(df, dt.max.time()))

    if app.dados_df is None or app.dados_df.empty or not funcionario:
        return [], pd.DataFrame()

    # Filtra base
    mask = (
        (app.dados_df["Funcionário"] == funcionario)
        & (app.dados_df["DATA"] >= data_ini)
        & (app.dados_df["DATA"] <= data_fim)
    )
    df_filtrado = app.dados_df.loc[mask].copy()
    if df_filtrado.empty:
        return [], df_filtrado

    # Converte colunas numéricas e garante lista de colunas da tela
    for col in app.colunas_para_somar:
        if col in df_filtrado:
            df_filtrado[col] = pd.to_numeric(df_filtrado[col], errors="coerce").fillna(0)

    if not getattr(app, "colunas_boletim_dados", None):
        app.colunas_boletim_dados = [
            "DATA", "BOLETIM", "HORA NORMAL", "H.E.", "H.E.D.", "H.E.N.", "H.E.N.D.", "H.N.", "S.A.",
        ]

    # Mapa: DATA(date) -> linha
    dados_mapeados = {d.date(): row for d, row in df_filtrado.set_index("DATA").iterrows()}

    dados_lista = []
    for dia in pd.date_range(data_ini, data_fim, freq="D"):
        if dia.date() in dados_mapeados:
            row = dados_mapeados[dia.date()]
            linha = [dia.strftime("%d/%m/%Y")]
            for col in app.colunas_boletim_dados[1:]:
                val = row.get(col)
                if col == "BOLETIM" and pd.notna(val):
                    try:
                        linha.append(str(int(float(val))))
                    except Exception:
                        linha.append(str(val))
                elif col in app.colunas_para_somar:
                    linha.append(app.formatar_tempo(val) if pd.notna(val) else "")
                else:
                    linha.append(str(val) if pd.notna(val) else "")
            dados_lista.append(linha)
        else:
            dados_lista.append([dia.strftime("%d/%m/%Y")] + [""] * (len(app.colunas_boletim_dados) - 1))

    # Totais
    totais = ["TOTAIS:"]
    for col in app.colunas_boletim_dados[1:]:
        if col in app.colunas_para_somar:
            soma = df_filtrado[col].sum()
            totais.append(app.formatar_tempo(soma))
        else:
            totais.append("")
    dados_lista.append(totais)

    return dados_lista, df_filtrado
# ---------- Relatório Geral (sobreavisos) ----------
def _coerce_dt(v):
    if v is None:
        return None
    if isinstance(v, (pd.Timestamp, )):
        return v
    return pd.to_datetime(v, dayfirst=True, errors="coerce")

def abrir_relatorio_geral(
    parent=None,
    *,
    data_ini=None,
    data_fim=None,
    caminho_parquet: str = ARQUIVO_PADRAO
):
    """
    Abre a janela 'Formulário Geral (sobreavisos)' usando RelatorioSADialog.
    Se data_ini/data_fim forem None, tenta derivar do próprio DF (min/max).
    """
    df = carregar_boletim_parquet(caminho_parquet)

    di = _coerce_dt(data_ini)
    df_ = _coerce_dt(data_fim)

    # fallback: usa faixa total do DF se não vierem datas
    if (di is None or pd.isna(di)) or (df_ is None or pd.isna(df_)):
        col_data = "DATA" if "DATA" in df.columns else ("Data" if "Data" in df.columns else None)
        if col_data:
            s = pd.to_datetime(df[col_data], dayfirst=True, errors="coerce").dropna()
            if not s.empty:
                di, df_ = s.min(), s.max()

    # Gera DataFrame de visualização
    df_view = relatorio_sa_por_registro_periodo(df, di, df_)

    # Abre o diálogo (Toplevel) – sem raiz “fantasma”
    owner = parent
    created_root = False
    if owner is None:
        owner = tk.Tk()
        owner.withdraw()
        created_root = True

    dlg = RelatorioSADialog(owner, df_view, data_ini=di, data_fim=df_)

    if created_root:
        try:
            dlg.wait_window()
        except Exception:
            pass
        owner.destroy()
