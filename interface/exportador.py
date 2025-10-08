# interface/exportador.py (versão enxuta)
# ------------------------------------------------------------
# Exportadores (Excel/PDF) centralizados e sem duplicações de
# formatação. Este módulo consome utilitários de utils/formatadores
# (cabeçalho, larguras, soma preservando formato e detecção de
# negativos) e mantém apenas o mínimo de helpers locais para
# leitura do tksheet e organização de abas/arquivos.
#
# Principais funções expostas (compatíveis com o app):
# - exportar_para_excel(app)
# - exportar_para_pdf(app)
# - exportar_ponto_para_excel(app, completo=False)
# - exportar_ponto_para_pdf(app, completo=False)
# - exportar_comparacao_individual(app)
# - exportar_comparacao_unificada(app)   ← inclui exportação "Diferenças apenas"
# - exportar_comparacao_geral(app)
# - exportar_totais_mensais(df_group, nome_base, pasta_destino)
# - exportar_totais_consolidados_excel(app, lista_contratos)  ← usa utils/formatadores
#
# Autor: Valdinei Lankewicz
# ------------------------------------------------------------
from __future__ import annotations

import re
import unicodedata
import datetime as dt
from datetime import datetime
from tkinter import filedialog, messagebox
from difflib import get_close_matches

import pandas as pd
_pd = pd 
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter as _gcl

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape, letter
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import cm, inch
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

# ============================================================
# Integração com utils/formatadores (com fallbacks seguros)
# ============================================================
try:  # preferir sempre usar os utilitários centrais
    from utils.formatadores import (
        aplicar_cabecalho_excel,
        formatar_planilha_excel,
        STARTROW_DADOS_XLSX as STARTROW_DADOS,
        eh_negativo_str,
        somar_preservando_formato,
    )
except Exception:
    # ---- fallbacks mínimos para não interromper a exportação ----
    STARTROW_DADOS = 7

    def aplicar_cabecalho_excel(ws, titulo, contrato, dt_ini, dt_fim, boletins, df_cols):
        max_col = max(1, len(df_cols))
        ws.append([titulo])
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max_col)
        ws["A1"].font = Font(bold=True, size=14)
        ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
        ws.append(["Contrato:", str(contrato)])
        try:
            periodo = f"{dt_ini:%d/%m/%Y} a {dt_fim:%d/%m/%Y}"
        except Exception:
            periodo = ""
        ws.append(["Período de Apuração:", periodo])
        ws.append(["Boletins Incluídos:", str(boletins)])
        ws.append([])
        ws.append([])  # dados começam na linha 8

    def formatar_planilha_excel(ws, startrow: int = STARTROW_DADOS):
        for c in range(1, ws.max_column + 1):
            ws.column_dimensions[_gcl(c)].width = 14
        ws.freeze_panes = f"A{startrow+1}"

    def eh_negativo_str(val) -> bool:
        if val is None:
            return False
        s = str(val).strip()
        if not s:
            return False
        if re.match(r"^-?\d{1,3}:\d{2}$", s):
            return s.startswith("-")
        s2 = s.replace(".", "").replace(",", ".")
        try:
            return float(s2) < 0
        except Exception:
            return s.startswith("-")

    def _is_hhmm(s: str) -> bool:
        return bool(re.match(r"^-?\d{1,3}:\d{2}$", str(s).strip()))

    def _to_minutes(s: str) -> int:
        m = re.match(r"^(-)?(\d{1,3}):(\d{2})$", str(s).strip())
        if not m:
            return 0
        sign = -1 if m.group(1) else 1
        return sign * (int(m.group(2)) * 60 + int(m.group(3)))

    def somar_preservando_formato(series) -> str:
        vals = [v for v in series if str(v).strip() not in ("", "-", "None", "nan")]
        if not vals:
            return "0:00"
        hhmm_count = sum(1 for v in vals if _is_hhmm(v))
        if hhmm_count >= max(1, len(vals) // 2):
            total_min = sum(_to_minutes(v) for v in vals if _is_hhmm(v))
            sign = "-" if total_min < 0 else ""
            total_min = abs(total_min)
            h, mi = divmod(total_min, 60)
            return f"{sign}{h}:{mi:02d}"
        acc = []
        for v in vals:
            s2 = str(v).replace(".", "").replace(",", ".")
            try:
                acc.append(float(s2))
            except Exception:
                return ""
        return f"{sum(acc):.2f}"

# ============================================================
# Constantes
# ============================================================
COLUNAS_PONTO = [
    "Data",
    "Nome",
    "CPF",
    "PIS",
    "Total Normais",
    "Total Noturno",
    "Extra 50%D",
    "Extra 100%D",
    "Extra 50%N",
    "Extra 100%N",
]

# ============================================================
# Helpers locais (pequenos e reaproveitáveis)
# ============================================================

from utils.dataframe_utils import (
    resolver_base_ponto,
    preparar_df_ponto_para_comparacao,
    preparar_df_boletim_para_comparacao,
    groupby_sum_by_date,
)
from utils.comparacao import dfs_sem_ponto
from utils.constantes import HEADERS_VIZ  # se já importar do app, ajuste

def _montar_tripla_para_export(app, funcionario, data_ini, data_fim):
    """
    Retorna (df_b_f, df_p_f, df_d_f, sem_ponto) prontos para exportação.
    - df_b_f: Boletim por dia, com cabeçalhos visuais
    - df_p_f: Ponto por dia, com cabeçalhos visuais (ou limpo se sem ponto)
    - df_d_f: Diferença por dia com cabeçalhos visuais (zerada se sem ponto)
    """
    # --- Boletim (B) ---
    df_b_base = app.dados_df[
        (app.dados_df["Funcionário"] == funcionario)
        & (app.dados_df["DATA"].dt.date >= data_ini)
        & (app.dados_df["DATA"].dt.date <= data_fim)
    ].copy()

    df_b_norm = preparar_df_boletim_para_comparacao(df_b_base)
    df_b = groupby_sum_by_date(
        df_b_norm.rename(columns={"DATA": "Data"}) if "DATA" in df_b_norm.columns else df_b_norm,
        "Data"
    )
    df_b_f = df_b.rename(columns={k: v for k, v in app.MAP_BOL.items() if k in df_b.columns}).copy()
    if "Data" in df_b_f.columns:
        df_b_f["Data"] = pd.to_datetime(df_b_f["Data"], errors="coerce").dt.floor("D")
    cols_b = ["Data"] + (["Boletim"] if "Boletim" in df_b_f.columns else []) + [h for h in app.HEADERS_VIZ if h in df_b_f.columns]
    df_b_f = df_b_f[cols_b] if "Data" in df_b_f.columns else pd.DataFrame(columns=["Data","Boletim"]+app.HEADERS_VIZ)

    # --- Ponto (P) ---
    df_ponto_sel, _, _ = resolver_base_ponto(
        app.df_ponto, app.df_relacao, funcionario, data_ini, data_fim, corte_similaridade=0.75
    )
    sem_ponto = (df_ponto_sel is None) or df_ponto_sel.empty

    if sem_ponto:
        df_p_f, df_d_f = dfs_sem_ponto(data_ini, data_fim, app.HEADERS_VIZ)
    else:
        df_p = preparar_df_ponto_para_comparacao(df_ponto_sel, data_ini, data_fim)
        inv_map_pto = {v: k for k, v in app.MAP_PTO.items()}
        df_p_f = df_p.rename(columns={k: inv_map_pto[k] for k in df_p.columns if k in inv_map_pto}).copy()
        if "Data" in df_p_f.columns:
            df_p_f["Data"] = pd.to_datetime(df_p_f["Data"], errors="coerce").dt.floor("D")
        cols_p = ["Data"] + [h for h in app.HEADERS_VIZ if h in df_p_f.columns]
        df_p_f = df_p_f[cols_p] if "Data" in df_p_f.columns else pd.DataFrame(columns=["Data"] + app.HEADERS_VIZ)

        # --- Diferença (D = B - P) ---
        df_m = pd.merge(df_b_f, df_p_f, on="Data", how="outer", suffixes=("_B", "_P")).sort_values("Data")
        for h in app.HEADERS_VIZ:
            if f"{h}_B" not in df_m.columns: df_m[f"{h}_B"] = 0.0
            if f"{h}_P" not in df_m.columns: df_m[f"{h}_P"] = 0.0
            df_m[f"{h}_B"] = pd.to_numeric(df_m[f"{h}_B"], errors="coerce").fillna(0.0)
            df_m[f"{h}_P"] = pd.to_numeric(df_m[f"{h}_P"], errors="coerce").fillna(0.0)
            df_m[h] = df_m[f"{h}_B"] - df_m[f"{h}_P"]
        df_d_f = df_m[["Data"] + app.HEADERS_VIZ].copy()

    return df_b_f, df_p_f, df_d_f, sem_ponto


def _get_headers_from_sheet(sheet):
    if sheet is None:
        return None
    h = getattr(sheet, "headers", None)
    return list(h() if callable(h) else h) if h else None


def _df_from_sheet_safe(sheet, expected_cols):
    data = [row[:] for row in sheet.get_sheet_data()] if sheet else []
    if not data:
        return pd.DataFrame(columns=expected_cols)
    row_len = max(len(r) for r in data)
    data = [(r + [""] * (row_len - len(r)))[:row_len] for r in data]
    cols = list(expected_cols[:row_len])
    if len(cols) < row_len:
        cols += [f"Col{idx}" for idx in range(len(cols) + 1, row_len + 1)]
    return pd.DataFrame(data, columns=cols)


def _sheet_name_unique(base, used: set) -> str:
    name = re.sub(r"[][:\\/?*]+", "", str(base)).strip() or "Aba"
    name = name[:31]
    orig = name
    i = 2
    while name in used:
        suf = f"_{i}"
        name = orig[: 31 - len(suf)] + suf
        i += 1
    used.add(name)
    return name


def _ult5_contratos(s):
    if s is None:
        return ""
    itens = [p.strip() for p in str(s).split(",") if p.strip()]
    return ", ".join([(i[-5:] if len(i) >= 5 else i) for i in itens])


def _canon_nome(nome: str) -> str:
    if not isinstance(nome, str):
        nome = "" if nome is None else str(nome)
    nome_norm = unicodedata.normalize("NFKC", nome)
    nome_norm = " ".join(nome_norm.strip().split())
    return nome_norm.upper()

# ============================================================
# Exportações simples (Aba 1 e 2)
# ============================================================

def exportar_para_excel(app):
    if app.df_filtrado is None or app.df_filtrado.empty:
        messagebox.showwarning("Aviso", "Não há dados para exportar.")
        return
    arquivo = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Arquivos Excel", "*.xlsx")],
        title="Salvar Relatório Excel Como",
        initialfile=f"Relatorio_{app.combo_funcionario.get()}.xlsx",
    )
    if not arquivo:
        return

    df = app.df_filtrado[app.colunas_visiveis].copy()
    if "DATA" in df.columns:
        df["DATA"] = pd.to_datetime(df["DATA"], errors="coerce").dt.strftime("%d/%m/%Y")

    somas = {}
    for col in getattr(app, "colunas_para_somar", []):
        try:
            somas[col] = pd.to_numeric(df[col], errors="coerce").fillna(0).sum()
        except Exception:
            pass
    if somas:
        total = pd.Series({**{c: "" for c in df.columns}, **somas}, name="TOTAL")
        if "DATA" in total.index:
            total["DATA"] = "TOTAL"
        df = pd.concat([df, total.to_frame().T], ignore_index=True)

    try:
        df.to_excel(arquivo, index=False)
        messagebox.showinfo("Sucesso", f"Relatório exportado para\n'{arquivo}'")
    except Exception as e:
        messagebox.showerror("Erro de Exportação", str(e))


def exportar_para_pdf(app):
    if app.df_filtrado is None or app.df_filtrado.empty:
        messagebox.showwarning("Aviso", "Não há dados para exportar.")
        return
    arquivo = filedialog.asksaveasfilename(
        defaultextension=".pdf",
        filetypes=[("Arquivos PDF", "*.pdf")],
        title="Salvar Relatório PDF Como",
        initialfile=f"Relatorio_{app.combo_funcionario.get()}.pdf",
    )
    if not arquivo:
        return

    df = app.df_filtrado[app.colunas_visiveis].copy()
    if "DATA" in df.columns:
        df["DATA"] = pd.to_datetime(df["DATA"], errors="coerce").dt.strftime("%d/%m/%Y")

    dados = [list(df.columns)] + df.values.tolist()

    somas = {}
    for col in getattr(app, "colunas_para_somar", []):
        try:
            somas[col] = pd.to_numeric(app.df_filtrado[col], errors="coerce").fillna(0).sum()
        except Exception:
            pass
    if somas:
        linha_total = ["TOTAL" if c == "DATA" else f"{somas.get(c, ''):.2f}" if c in somas else "" for c in df.columns]
        dados.append(linha_total)

    doc = SimpleDocTemplate(arquivo, pagesize=landscape(letter))
    styles = getSampleStyleSheet()
    story = [
        Paragraph("Relatório de Horas", styles["Title"]),
        Paragraph(f"Funcionário: {app.combo_funcionario.get()}", styles["Normal"]),
        Paragraph(
            f"Período: {app.cal_data_inicial.entry.get()} a {app.cal_data_final.entry.get()}",
            styles["Normal"],
        ),
        Spacer(1, 0.2 * inch),
    ]

    tabela = Table(dados)
    tabela.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#4472C4")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.whitesmoke),
                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("BOTTOMPADDING", (0, 0), (-1, 0), 12),
                ("BACKGROUND", (0, 1), (-1, -2), colors.HexColor("#B4C6E7")),
                ("BACKGROUND", (0, -1), (-1, -1), colors.HexColor("#8EA9DB")),
                ("FONTNAME", (0, -1), (-1, -1), "Helvetica-Bold"),
                ("GRID", (0, 0), (-1, -1), 1, colors.black),
            ]
        )
    )
    story.append(tabela)
    try:
        doc.build(story)
        messagebox.showinfo("Sucesso", f"Relatório exportado para\n'{arquivo}'")
    except Exception as e:
        messagebox.showerror("Erro de Exportação", str(e))

# ============================================================
# Exportações de Ponto (Aba 2)
# ============================================================

def exportar_ponto_para_excel(app, completo: bool = False):
    df = app.df_ponto if completo else app.df_ponto_filtrado
    if df is None or df.empty:
        messagebox.showwarning("Aviso", "Nenhum dado de ponto disponível para exportar.")
        return

    if completo:
        nome_arquivo = "RegistroPonto_Completo.xlsx"
    else:
        funcionario = app.combo_funcionario.get().strip().replace(" ", "_")
        data_ini = dt.datetime.strptime(app.cal_data_inicial.entry.get(), "%d/%m/%Y")
        nome_arquivo = f"RegistroPonto_{funcionario}_{data_ini.strftime('%Y%m')}.xlsx"

    arquivo = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Arquivo Excel", "*.xlsx")],
        initialfile=nome_arquivo,
        title="Salvar Relatório de Ponto",
    )
    if not arquivo:
        return

    df_export = df.copy()
    df_export["Data"] = pd.to_datetime(df_export["Data"]).dt.strftime("%d/%m/%Y")
    df_export = df_export[COLUNAS_PONTO]

    soma_linha = {
        col: pd.to_numeric(df_export[col], errors="coerce").fillna(0).sum()
        for col in df_export.columns
        if col not in ["Data", "Nome", "CPF", "PIS"]
    }
    total_row = pd.DataFrame(
        [
            {
                "Data": "TOTAIS:",
                "Nome": "",
                "CPF": "",
                "PIS": "",
                **{k: round(v, 2) for k, v in soma_linha.items()},
            }
        ]
    )
    df_export = pd.concat([df_export, total_row], ignore_index=True)

    try:
        df_export.to_excel(arquivo, index=False)
        messagebox.showinfo("Sucesso", f"Relatório exportado para\n'{arquivo}'")
    except Exception as e:
        messagebox.showerror("Erro de Exportação", str(e))


def exportar_ponto_para_pdf(app, completo: bool = False):
    df = app.df_ponto if completo else app.df_ponto_filtrado
    if df is None or df.empty:
        messagebox.showwarning("Aviso", "Nenhum dado de ponto disponível para exportar.")
        return

    if completo:
        nome_arquivo = "RegistroPonto_Completo.pdf"
    else:
        funcionario = app.combo_funcionario.get().strip().replace(" ", "_")
        data_ini = dt.datetime.strptime(app.cal_data_inicial.entry.get(), "%d/%m/%Y")
        nome_arquivo = f"RegistroPonto_{funcionario}_{data_ini.strftime('%Y%m')}.pdf"

    arquivo = filedialog.asksaveasfilename(
        defaultextension=".pdf",
        filetypes=[("Arquivo PDF", "*.pdf")],
        initialfile=nome_arquivo,
        title="Salvar Relatório de Ponto",
    )
    if not arquivo:
        return

    doc = SimpleDocTemplate(arquivo, pagesize=landscape(letter))
    styles = getSampleStyleSheet()
    story = []

    titulo = "Relatório de Ponto (Completo)" if completo else "Relatório de Ponto do Funcionário"
    story.append(Paragraph(titulo, styles["Title"]))
    if not completo:
        funcionario = app.combo_funcionario.get()
        periodo = f"{app.cal_data_inicial.entry.get()} a {app.cal_data_final.entry.get()}"
        story.append(Paragraph(f"Funcionário: {funcionario}", styles["Normal"]))
        story.append(Paragraph(f"Período: {periodo}", styles["Normal"]))
    story.append(Spacer(1, 12))

    df_export = df.copy()
    df_export["Data"] = pd.to_datetime(df_export["Data"]).dt.strftime("%d/%m/%Y")
    df_export = df_export[COLUNAS_PONTO]

    dados = [COLUNAS_PONTO] + df_export.values.tolist()
    soma_linha = ["TOTAIS:", "", "", ""] + [
        f"{pd.to_numeric(df_export[col], errors='coerce').fillna(0).sum():.2f}"
        for col in COLUNAS_PONTO[4:]
    ]
    dados.append(soma_linha)

    tabela = Table(dados)
    tabela.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#4472C4")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.whitesmoke),
                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("BOTTOMPADDING", (0, 0), (-1, 0), 12),
                ("BACKGROUND", (0, 1), (-1, -2), colors.HexColor("#D9E1F2")),
                ("BACKGROUND", (0, -1), (-1, -1), colors.HexColor("#B4C6E7")),
                ("FONTNAME", (0, -1), (-1, -1), "Helvetica-Bold"),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
            ]
        )
    )
    try:
        doc.build(story + [tabela])
        messagebox.showinfo("Sucesso", f"Relatório exportado para\n'{arquivo}'")
    except Exception as e:
        messagebox.showerror("Erro de Exportação", str(e))

# ============================================================
# Comparação por funcionário (Aba 4)
# ============================================================

def _gerar_dados_comparacao_funcionario(app, funcionario, data_ini_dt, data_fim_dt, contrato_id_especifico=None):
    """Retorna (df_boletim, df_ponto, df_diferenca, registro, boletins_set)."""
    df_boletim_orig = app.dados_df.copy()
    df_ponto_orig = app.df_ponto.copy()

    df_boletim_orig["DATA"] = pd.to_datetime(df_boletim_orig["DATA"], errors="coerce")
    df_ponto_orig["Data"] = pd.to_datetime(df_ponto_orig["Data"], errors="coerce")

    # Boletim base
    filtro_base = (
        (df_boletim_orig["Funcionário"].str.upper() == funcionario.upper())
        & (df_boletim_orig["DATA"].dt.date >= data_ini_dt.date())
        & (df_boletim_orig["DATA"].dt.date <= data_fim_dt.date())
    )
    df_boletim_filt = df_boletim_orig[filtro_base].copy()

    dias_do_contrato = None
    if contrato_id_especifico and not df_boletim_filt.empty:
        df_boletim_filt = df_boletim_filt[df_boletim_filt["Contrato"] == contrato_id_especifico].copy()
        dias_do_contrato = df_boletim_filt["DATA"].dt.date.unique()

    # Ponto: escolha de nome similar
    nomes_ponto = df_ponto_orig["Nome"].dropna().unique()
    nome_similar = get_close_matches(funcionario.upper(), [n.upper() for n in nomes_ponto], n=1, cutoff=0.75)
    df_ponto_filt = pd.DataFrame()
    if nome_similar:
        nome_ponto = next((n for n in nomes_ponto if n.upper() == nome_similar[0]), None)
        if nome_ponto:
            df_ponto_base = df_ponto_orig[
                (df_ponto_orig["Nome"].str.upper() == nome_ponto.upper())
                & (df_ponto_orig["Data"].dt.date >= data_ini_dt.date())
                & (df_ponto_orig["Data"].dt.date <= data_fim_dt.date())
            ].copy()
            df_ponto_filt = (
                df_ponto_base if dias_do_contrato is None else df_ponto_base[df_ponto_base["Data"].dt.date.isin(dias_do_contrato)].copy()
            )

    campos = [
        ("HORA NORMAL", "Total Normais"),
        ("H.N.", "Total Noturno"),
        ("H.E.", "Extra 50%D"),
        ("H.E.D.", "Extra 100%D"),
        ("H.E.N.", "Extra 50%N"),
        ("H.E.N.D.", "Extra 100%N"),
    ]
    colunas_resultado = app.headers_comparacao  # ["Data", ...]

    bgrp = (df_boletim_filt.groupby(df_boletim_filt["DATA"].dt.date).sum(numeric_only=True)) if not df_boletim_filt.empty else pd.DataFrame()
    pgrp = (df_ponto_filt.groupby(df_ponto_filt["Data"].dt.date).sum(numeric_only=True)) if not df_ponto_filt.empty else pd.DataFrame()

    datas = sorted(set(bgrp.index).union(pgrp.index))
    dados_boletim, dados_ponto, dados_dif = [], [], []

    for data in datas:
        lb = [data.strftime("%d/%m/%Y")]
        lp = [data.strftime("%d/%m/%Y")]
        ld = [data.strftime("%d/%m/%Y")]
        for campo_b, campo_p in campos:
            vb = bgrp.at[data, campo_b] if (data in bgrp.index and campo_b in bgrp.columns) else 0
            vp = pgrp.at[data, campo_p] if (data in pgrp.index and campo_p in pgrp.columns) else 0
            d = vb - vp
            lb.append(f"{vb:.2f}" if vb else "-")
            lp.append(f"{vp:.2f}" if vp else "-")
            ld.append(f"{d:.2f}" if d else "-")
        dados_boletim.append(lb)
        dados_ponto.append(lp)
        dados_dif.append(ld)

    df_res_b = pd.DataFrame(dados_boletim, columns=colunas_resultado)
    df_res_p = pd.DataFrame(dados_ponto, columns=colunas_resultado)
    df_res_d = pd.DataFrame(dados_dif, columns=colunas_resultado)

    registro = (
        df_boletim_filt.get("Registro", pd.Series(["-"])).iloc[0]
        if not df_boletim_filt.empty
        else "-"
    )

    boletins_set = set()
    if not df_boletim_filt.empty and "BOLETIM" in df_boletim_filt.columns:
        nums = pd.to_numeric(df_boletim_filt["BOLETIM"], errors="coerce").dropna().astype(int)
        boletins_set = set(nums.unique().tolist())

    return df_res_b, df_res_p, df_res_d, registro, boletins_set

# ============================================================
# Exportação comparativa (layout da tela + Diferenças apenas)
# ============================================================

def exportar_comparacao_individual(app):
    pasta = filedialog.askdirectory(title="Escolha a pasta para salvar os arquivos")
    if not pasta:
        return

    try:
        need = False
        if hasattr(app.sheet_comp_boletim, "get_total_rows"):
            need = app.sheet_comp_boletim.get_total_rows() == 0
        else:
            need = not app.sheet_comp_boletim.get_sheet_data()
        if need:
            app.gerar_relatorio_comparacao()
    except Exception:
        pass

    nome_display = (app.combo_funcionario.get() or "Funcionario").strip()
    nome_file = nome_display.replace(" ", "_")
    periodo = f"{app.cal_data_inicial.entry.get()} a {app.cal_data_final.entry.get()}"

    headers_vis = _get_headers_from_sheet(app.sheet_comp_boletim) or [
        "Data",
        "Horas\nNormais",
        "Horas\nNoturnas",
        "Extra\n50%D",
        "Extra\n100%D",
        "Extra\n50%N",
        "Extra\n100%N",
    ]
    headers_horas = headers_vis[1:]

    df_bol = _df_from_sheet_safe(app.sheet_comp_boletim, ["Data"] + headers_horas)
    df_pto_raw = _df_from_sheet_safe(app.sheet_comp_ponto, headers_horas)
    df_dif_raw = _df_from_sheet_safe(app.sheet_comp_diferenca, headers_horas)

    min_len = min(len(df_bol), len(df_pto_raw), len(df_dif_raw)) if len(df_bol) else 0
    df_bol = df_bol.iloc[:min_len].reset_index(drop=True)
    df_pto = pd.concat([df_bol["Data"], df_pto_raw.iloc[:min_len].reset_index(drop=True)], axis=1)
    df_dif = pd.concat([df_bol["Data"], df_dif_raw.iloc[:min_len].reset_index(drop=True)], axis=1)

    # remove possíveis linhas "TOTAIS:"
    for df in (df_bol, df_pto, df_dif):
        if not df.empty and isinstance(df.iloc[-1, 0], str) and "TOTAIS" in df.iloc[-1, 0].upper():
            df.drop(df.index[-1], inplace=True)

    # ---- XLSX layout tela ----
    caminho_xlsx = f"{pasta}/{nome_file}_comparacao.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "Comparação"

    # Título + metadados simples
    ws.append([f"Relatório Comparativo — Funcionário: {nome_display}"])
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=1 + (3*len(headers_horas)))
    ws["A1"].font = Font(bold=True, size=14)
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.append([f"Período: {periodo}"])
    ws.append([])
    ws.append([])

    ws.append([""] + ["Boletim"] * len(headers_horas) + ["Ponto"] * len(headers_horas) + ["Diferença"] * len(headers_horas))
    ws.append(["Data"] + headers_horas * 3)

    df_final = pd.DataFrame()
    df_final["Data"] = df_bol["Data"]
    for origem, df_src in (("Boletim", df_bol), ("Ponto", df_pto), ("Diferença", df_dif)):
        for h in headers_horas:
            df_final[f"{origem} {h.replace(chr(10), ' ')}"] = df_src[h]

    for _, row in df_final.iterrows():
        ws.append(row.tolist())

    # linha final sem somas
    ws.append(["TOTAIS:"] + [""] * (len(df_final.columns) - 1))

    azul = PatternFill(start_color="B7DFFB", end_color="B7DFFB", fill_type="solid")
    verde = PatternFill(start_color="C5EDC1", end_color="C5EDC1", fill_type="solid")
    laranja = PatternFill(start_color="FFDAB9", end_color="FFDAB9", fill_type="solid")

    for cell in ws[5]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")
    for cell in ws[6]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    ws.column_dimensions["A"].width = 14
    width_data = round(12 * 0.8, 1)
    for c in range(2, 2 + 3 * len(headers_horas)):
        ws.column_dimensions[_gcl(c)].width = width_data

    max_row = ws.max_row
    # Tonalizar grupos e marcar negativos
    # Boletim
    for c in range(2, 2 + len(headers_horas)):
        ws.cell(row=5, column=c).fill = azul
        ws.cell(row=6, column=c).fill = azul
        for r in range(7, max_row + 1):
            ws.cell(row=r, column=c).fill = azul
            ws.cell(row=r, column=c).alignment = Alignment(horizontal="right", vertical="center")
    # Ponto
    p_ini, p_fim = 2 + len(headers_horas), 1 + 2 * len(headers_horas)
    for c in range(p_ini, p_fim + 1):
        ws.cell(row=5, column=c).fill = verde
        ws.cell(row=6, column=c).fill = verde
        for r in range(7, max_row + 1):
            ws.cell(row=r, column=c).fill = verde
            ws.cell(row=r, column=c).alignment = Alignment(horizontal="right", vertical="center")
    # Diferença
    d_ini, d_fim = p_fim + 1, 1 + 3 * len(headers_horas)
    for c in range(d_ini, d_fim + 1):
        ws.cell(row=5, column=c).fill = laranja
        ws.cell(row=6, column=c).fill = laranja
        for r in range(7, max_row + 1):
            cell = ws.cell(row=r, column=c)
            cell.fill = laranja
            cell.alignment = Alignment(horizontal="right", vertical="center")
            if eh_negativo_str(cell.value):
                cell.font = Font(color="FF0000", bold=True)

    ws.freeze_panes = "A7"
    wb.save(caminho_xlsx)

    # ---- PDF enxuto (layout tela) ----
    caminho_pdf = f"{pasta}/{nome_file}_comparacao.pdf"
    doc = SimpleDocTemplate(caminho_pdf, pagesize=landscape(A4), rightMargin=20, leftMargin=20, topMargin=20, bottomMargin=12)
    styles = getSampleStyleSheet()

    header_cells = ["Data"] + headers_horas * 3
    dados = [header_cells] + df_final.values.tolist() + [["TOTAIS:"] + [""] * (len(header_cells) - 1)]

    col_widths = [1.8 * cm] + [1.36 * cm] * (len(header_cells) - 1)
    tabela = Table(dados, colWidths=col_widths, repeatRows=1)

    estilo = TableStyle(
        [
            ("GRID", (0, 0), (-1, -1), 0.3, colors.grey),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("ALIGN", (0, 0), (-1, 0), "CENTER"),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.whitesmoke),
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#183C5F")),
            ("ALIGN", (1, 1), (-1, -1), "RIGHT"),
            ("FONTSIZE", (0, 0), (-1, -1), 7),
        ]
    )
    b_ini, b_fim = 1, len(headers_horas)
    p_ini, p_fim = b_fim + 1, b_fim + len(headers_horas)
    d_ini, d_fim = p_fim + 1, p_fim + len(headers_horas)
    estilo.add("BACKGROUND", (b_ini, 1), (b_fim, -1), colors.HexColor("#B7DFFB"))
    estilo.add("BACKGROUND", (p_ini, 1), (p_fim, -1), colors.HexColor("#C5EDC1"))
    estilo.add("BACKGROUND", (d_ini, 1), (d_fim, -1), colors.HexColor("#FFDAB9"))

    # negativos em Diferença
    for r in range(1, len(dados)):
        for c in range(d_ini, d_fim + 1):
            if eh_negativo_str(dados[r][c]):
                estilo.add("TEXTCOLOR", (c, r), (c, r), colors.red)
                estilo.add("FONTNAME", (c, r), (c, r), "Helvetica-Bold")

    tabela.setStyle(estilo)
    story = [
        Paragraph("<b>Relatório Comparativo — Funcionário Atual</b>", styles["Title"]),
        Paragraph(f"<b>Funcionário:</b> {nome_display}", styles["Normal"]),
        Paragraph(f"<b>Período:</b> {periodo}", styles["Normal"]),
        Spacer(1, 8),
        tabela,
    ]
    doc.build(story)

    messagebox.showinfo("Sucesso", f"Arquivos salvos:\n- {caminho_xlsx}\n- {caminho_pdf}")


def exportar_comparacao_unificada(app):
    """Exporta layout da tela + cria um par XLSX/PDF apenas com as DIFERENÇAS.
    - Negativos em vermelho e negrito (xlsx + pdf)
    - Totais preservando o formato (HH:MM ou decimal) via utils/formatadores
    """
    nome_display = (app.combo_funcionario.get() or "Funcionario").strip()
    arquivo_base = filedialog.asksaveasfilename(
        defaultextension="",
        filetypes=[("Todos os Arquivos", "*.*")],
        title="Salvar Exportação Unificada",
        initialfile=f"Comparacao_{nome_display.replace(' ', '_')}",
    )
    if not arquivo_base:
        return

    caminho_xlsx = arquivo_base + ".xlsx"
    caminho_pdf = arquivo_base + ".pdf"

    headers_vis = _get_headers_from_sheet(app.sheet_comp_boletim) or [
        "Data",
        "Horas\nNormais",
        "Horas\nNoturnas",
        "Extra\n50%D",
        "Extra\n100%D",
        "Extra\n50%N",
        "Extra\n100%N",
    ]
    headers = headers_vis[1:]

    df_bol = _df_from_sheet_safe(app.sheet_comp_boletim, ["Data"] + headers)
    df_pto_raw = _df_from_sheet_safe(app.sheet_comp_ponto, headers)
    df_dif_raw = _df_from_sheet_safe(app.sheet_comp_diferenca, headers)

    min_len = min(len(df_bol), len(df_pto_raw), len(df_dif_raw)) if len(df_bol) else 0
    df_bol = df_bol.iloc[:min_len].reset_index(drop=True)
    df_pto = pd.concat([df_bol["Data"], df_pto_raw.iloc[:min_len].reset_index(drop=True)], axis=1)
    df_dif = pd.concat([df_bol["Data"], df_dif_raw.iloc[:min_len].reset_index(drop=True)], axis=1)

    for df in (df_bol, df_pto, df_dif):
        if not df.empty and isinstance(df.iloc[-1, 0], str) and "TOTAIS" in df.iloc[-1, 0].upper():
            df.drop(df.index[-1], inplace=True)

    # monta DF final para o layout tela
    df_final = pd.DataFrame()
    df_final["Data"] = df_bol["Data"]
    for origem, df_src in (("Boletim", df_bol), ("Ponto", df_pto), ("Diferença", df_dif)):
        for h in headers:
            df_final[f"{origem} {h.replace(chr(10), ' ')}"] = df_src[h]

    totais = [somar_preservando_formato(df_final[col]) for col in df_final.columns[1:]]

    # ========= XLSX layout tela =========
    wb = Workbook(); ws = wb.active; ws.title = "Comparação"
    periodo = f"{app.cal_data_inicial.entry.get()} a {app.cal_data_final.entry.get()}"
    ws.append([f"Relatório Comparativo Unificado — Funcionário: {nome_display}"])
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=1 + len(df_final.columns) - 1)
    ws["A1"].font = Font(bold=True, size=14)
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.append([f"Período: {periodo}"])
    ws.append([]); ws.append([]); ws.append([])
    ws.append([""] + ["Boletim"] * len(headers) + ["Ponto"] * len(headers) + ["Diferença"] * len(headers))
    ws.append(["Data"] + headers * 3)

    for _, row in df_final.iterrows():
        ws.append(row.tolist())
    ws.append(["TOTAIS:"] + totais)

    azul = PatternFill(start_color="B7DFFB", end_color="B7DFFB", fill_type="solid")
    verde = PatternFill(start_color="C5EDC1", end_color="C5EDC1", fill_type="solid")
    laranja = PatternFill(start_color="FFDAB9", end_color="FFDAB9", fill_type="solid")

    for cell in ws[6]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")
    for cell in ws[7]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    ws.column_dimensions["A"].width = 14
    width_data = round(12 * 0.8, 1)
    for col in range(2, 2 + 3 * len(headers)):
        ws.column_dimensions[_gcl(col)].width = width_data

    max_row = ws.max_row
    # Boletim
    for c in range(2, 2 + len(headers)):
        ws.cell(row=6, column=c).fill = azul
        ws.cell(row=7, column=c).fill = azul
        for r in range(8, max_row + 1):
            ws.cell(row=r, column=c).fill = azul
            ws.cell(row=r, column=c).alignment = Alignment(horizontal="right", vertical="center")
    # Ponto
    p_ini, p_fim = 2 + len(headers), 1 + 2 * len(headers)
    for c in range(p_ini, p_fim + 1):
        ws.cell(row=6, column=c).fill = verde
        ws.cell(row=7, column=c).fill = verde
        for r in range(8, max_row + 1):
            ws.cell(row=r, column=c).fill = verde
            ws.cell(row=r, column=c).alignment = Alignment(horizontal="right", vertical="center")
    # Diferença + negativos
    d_ini, d_fim = p_fim + 1, 1 + 3 * len(headers)
    for c in range(d_ini, d_fim + 1):
        ws.cell(row=6, column=c).fill = laranja
        ws.cell(row=7, column=c).fill = laranja
        for r in range(8, max_row + 1):
            cell = ws.cell(row=r, column=c)
            cell.fill = laranja
            cell.alignment = Alignment(horizontal="right", vertical="center")
            if eh_negativo_str(cell.value):
                cell.font = Font(color="FF0000", bold=True)

    ws.freeze_panes = "A8"
    wb.save(caminho_xlsx)

    # ========= DIFERENÇAS APENAS =========
    caminho_xlsx_dif = arquivo_base + "_Diferencas.xlsx"
    caminho_pdf_dif = arquivo_base + "_Diferencas.pdf"

    totais_dif = [somar_preservando_formato(df_dif[h]) for h in headers]

    wb_d = Workbook(); ws_d = wb_d.active; ws_d.title = "Diferenças"
    ws_d.append([f"Relatório — Diferenças (Boletim − Ponto) — Funcionário: {nome_display}"])
    ws_d.merge_cells(start_row=1, start_column=1, end_row=1, end_column=1 + len(headers))
    ws_d["A1"].font = Font(bold=True, size=14)
    ws_d["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws_d.append([f"Período: {periodo}"])
    ws_d.append([]); ws_d.append([])
    ws_d.append([""] + ["Diferença"] * len(headers))
    ws_d.append(["Data"] + headers)
    for _, row in df_dif.iterrows():
        ws_d.append(row.tolist())
    ws_d.append(["TOTAIS:"] + totais_dif)

    for cell in ws_d[6]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")
    for cell in ws_d[7]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    ws_d.column_dimensions["A"].width = 14
    for col in range(2, 2 + len(headers)):
        ws_d.column_dimensions[_gcl(col)].width = round(12 * 0.8, 1)

    laranja = PatternFill(start_color="FFDAB9", end_color="FFDAB9", fill_type="solid")
    max_row = ws_d.max_row
    for c in range(2, 2 + len(headers)):
        ws_d.cell(row=6, column=c).fill = laranja
        ws_d.cell(row=7, column=c).fill = laranja
        for r in range(8, max_row + 1):
            cell = ws_d.cell(row=r, column=c)
            cell.fill = laranja
            cell.alignment = Alignment(horizontal="right", vertical="center")
            if eh_negativo_str(cell.value):
                cell.font = Font(color="FF0000", bold=True)

    ws_d.freeze_panes = "A8"
    wb_d.save(caminho_xlsx_dif)

    # PDF diferenças
    doc_d = SimpleDocTemplate(caminho_pdf_dif, pagesize=landscape(A4), rightMargin=20, leftMargin=20, topMargin=20, bottomMargin=12)
    styles = getSampleStyleSheet()
    header_cells_d = ["Data"] + headers
    dados_d = [header_cells_d] + df_dif.values.tolist() + [["TOTAIS:"] + totais_dif]
    tabela_d = Table(dados_d, colWidths=[1.8*cm] + [1.36*cm]*(len(header_cells_d)-1), repeatRows=1)
    estilo_d = TableStyle([
        ("GRID", (0,0), (-1,-1), 0.3, colors.grey),
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ("ALIGN", (0,0), (-1,0), "CENTER"),
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("TEXTCOLOR", (0,0), (-1,0), colors.whitesmoke),
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#183C5F")),
        ("ALIGN", (1,1), (-1,-1), "RIGHT"),
        ("FONTSIZE", (0,0), (-1,-1), 7),
    ])
    # pintar bloco
    estilo_d.add("BACKGROUND", (1,1), (-1,-1), colors.HexColor("#FFDAB9"))
    for r in range(1, len(dados_d)):
        for c in range(1, len(header_cells_d)):
            if eh_negativo_str(dados_d[r][c]):
                estilo_d.add("TEXTCOLOR", (c, r), (c, r), colors.red)
                estilo_d.add("FONTNAME", (c, r), (c, r), "Helvetica-Bold")
    tabela_d.setStyle(estilo_d)

    story_d = [
        Paragraph("<b>Relatório — Diferenças (Boletim − Ponto)</b>", styles["Title"]),
        Paragraph(f"<b>Funcionário:</b> {nome_display}", styles["Normal"]),
        Paragraph(f"<b>Período:</b> {periodo}", styles["Normal"]),
        Spacer(1, 8),
        tabela_d,
    ]
    doc_d.build(story_d)

    # PDF layout tela
    doc = SimpleDocTemplate(caminho_pdf, pagesize=landscape(A4), rightMargin=20, leftMargin=20, topMargin=20, bottomMargin=12)
    styles = getSampleStyleSheet()
    header_cells = ["Data"] + headers * 3
    dados = [header_cells] + df_final.values.tolist() + [["TOTAIS:"] + totais]
    tabela = Table(dados, colWidths=[1.8*cm] + [1.36*cm]*(len(header_cells)-1), repeatRows=1)
    estilo = TableStyle([
        ("GRID", (0,0), (-1,-1), 0.3, colors.grey),
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ("ALIGN", (0,0), (-1,0), "CENTER"),
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("TEXTCOLOR", (0,0), (-1,0), colors.whitesmoke),
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#183C5F")),
        ("ALIGN", (1,1), (-1,-1), "RIGHT"),
        ("FONTSIZE", (0,0), (-1,-1), 7),
    ])
    b_ini, b_fim = 1, len(headers)
    p_ini, p_fim = b_fim + 1, b_fim + len(headers)
    d_ini, d_fim = p_fim + 1, p_fim + len(headers)
    estilo.add("BACKGROUND", (b_ini,1), (b_fim,-1), colors.HexColor("#B7DFFB"))
    estilo.add("BACKGROUND", (p_ini,1), (p_fim,-1), colors.HexColor("#C5EDC1"))
    estilo.add("BACKGROUND", (d_ini,1), (d_fim,-1), colors.HexColor("#FFDAB9"))
    for r in range(1, len(dados)):
        for c in range(d_ini, d_fim + 1):
            if eh_negativo_str(dados[r][c]):
                estilo.add("TEXTCOLOR", (c,r), (c,r), colors.red)
                estilo.add("FONTNAME", (c,r), (c,r), "Helvetica-Bold")
    tabela.setStyle(estilo)

    story = [
        Paragraph("<b>Relatório Comparativo Unificado</b>", styles["Title"]),
        Paragraph(f"<b>Funcionário:</b> {nome_display}", styles["Normal"]),
        Paragraph(f"<b>Período:</b> {periodo}", styles["Normal"]),
        Spacer(1, 8),
        tabela,
    ]
    doc.build(story)

    messagebox.showinfo(
        "Sucesso",
        "Exportações concluídas!\n\n"
        f"XLSX: {caminho_xlsx}\nPDF:  {caminho_pdf}\n\n"
        f"XLSX (Diferenças): {caminho_xlsx_dif}\nPDF  (Diferenças): {caminho_pdf_dif}"
    )

# ============================================================
# Comparação geral e Totais Mensais (mesmas regras visuais)
# ============================================================

def exportar_comparacao_geral(app):
    try:
        pasta = filedialog.askdirectory(title="Escolha uma pasta para salvar os arquivos")
        if not pasta:
            return
        if app.dados_df is None or app.df_ponto is None:
            messagebox.showwarning("Aviso", "Dados completos não carregados.")
            return
        hoje = dt.datetime.now().strftime("%Y-%m-%d")
        app.dados_df.to_excel(f"{pasta}/boletim_completo_{hoje}.xlsx", index=False)
        app.df_ponto.to_excel(f"{pasta}/ponto_completo_{hoje}.xlsx", index=False)
        doc = SimpleDocTemplate(f"{pasta}/comparacao_geral_{hoje}.pdf", pagesize=landscape(A4))
        styles = getSampleStyleSheet()
        story = [
            Paragraph("Comparação Geral de Dados", styles["Title"]),
            Paragraph(f"Exportado em: {hoje}", styles["Normal"]),
            Spacer(1, 0.3 * inch),
            Paragraph("⚠️ PDF não contém tabelas completas. Verifique os .xlsx gerados.", styles["Normal"]),
        ]
        doc.build(story)
        messagebox.showinfo("Sucesso", f"Dados completos exportados para: {pasta}")
    except Exception as e:
        messagebox.showerror("Erro", str(e))


def exportar_totais_mensais(df_group: pd.DataFrame, nome_base: str, pasta_destino: str):
    try:
        caminho_xlsx = f"{pasta_destino}/{nome_base}.xlsx"
        caminho_pdf = f"{pasta_destino}/{nome_base}.pdf"
        df_group.to_excel(caminho_xlsx, index=False)

        dados = [list(df_group.columns)] + df_group.values.tolist()
        for i in range(1, len(dados)):
            for j in range(3, len(dados[i])):
                try:
                    dados[i][j] = f"{float(dados[i][j]):.2f}"
                except Exception:
                    pass

        doc = SimpleDocTemplate(caminho_pdf, pagesize=A4)
        styles = getSampleStyleSheet()
        tabela = Table(dados, repeatRows=1)
        tabela.setStyle(
            TableStyle([
                ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#003366")),
                ("TEXTCOLOR", (0,0), (-1,0), colors.white),
                ("ALIGN", (0,0), (-1,-1), "CENTER"),
                ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
                ("FONTSIZE", (0,0), (-1,0), 6),
                ("FONTSIZE", (0,1), (-1,-1), 6),
                ("BACKGROUND", (0,1), (-1,-1), colors.whitesmoke),
                ("GRID", (0,0), (-1,-1), 0.5, colors.grey),
            ])
        )
        doc.build([
            Paragraph(f"Totais Mensais por Funcionário - {nome_base[-6:]}", styles["Title"]),
            Spacer(1, 12),
            tabela,
        ])
        messagebox.showinfo("Exportado", f"Totais mensais salvos:\n- {nome_base}.xlsx\n- {nome_base}.pdf")
    except Exception as e:
        messagebox.showerror("Erro", f"Falha na exportação de totais mensais:\n{e}")

# ============================================================
# Consolidação por Contrato (enxuta e com utils)
# ============================================================

def exportar_totais_consolidados_excel(app, lista_contratos):
    """
    Cria um workbook com:
      • Abas por contrato: funcionários pertencentes ao contrato e seus totais,
        com cabeçalho padronizado (utils.formatadores) e dados iniciando em
        STARTROW_DADOS.
      • Aba "Resumo Geral": Nome | Contrato(últimos 5) | Boletim(lista por nome) | totais…
      • Um segundo workbook "_Diferencas.xlsx" contendo apenas colunas "Diferença …"
        com negativos em vermelho e negrito.
    """
    arquivo_saida = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Arquivos Excel", "*.xlsx")],
        title="Salvar Totais Consolidados por Contrato",
        initialfile="Totais_Consolidados_por_Contrato.xlsx",
    )
    if not arquivo_saida:
        return

    # Período a partir da UI
    dt_ini = datetime.strptime(app.cal_data_inicial.entry.get(), "%d/%m/%Y")
    dt_fim = datetime.strptime(app.cal_data_final.entry.get(), "%d/%m/%Y")

    # 1) Calcula uma vez os totais por funcionário (função do App)
    df_boletim_filtrado = app.dados_df[app.dados_df["Contrato"].isin(lista_contratos)]
    todos_funcionarios = sorted(df_boletim_filtrado["Funcionário"].dropna().unique())

    resultados = []
    for nome in todos_funcionarios:
        totais = app._calcular_totais_funcionario(nome)
        if not totais:
            continue
        contratos_do_func = df_boletim_filtrado[df_boletim_filtrado["Funcionário"] == nome]["Contrato"].unique()
        totais["Contratos"] = ", ".join(map(str, contratos_do_func))
        resultados.append(totais)

    if not resultados:
        messagebox.showwarning("Aviso", "Nenhum dado encontrado para os critérios selecionados.")
        return

    df_all = pd.DataFrame(resultados)

    with pd.ExcelWriter(arquivo_saida, engine="openpyxl") as writer:
        used_names = set()
        for contrato_id in sorted(set(lista_contratos)):
            sheet_name = _sheet_name_unique(contrato_id, used_names)
            df_contrato = df_all[df_all["Contratos"].apply(lambda x: str(contrato_id) in str(x))].copy()
            if df_contrato.empty:
                continue

            # totalição e limpeza de colunas auxiliares
            df_contrato.drop(columns=["Contratos"], inplace=True, errors="ignore")
            somas = df_contrato.select_dtypes(include="number").sum()
            linha_total = pd.Series(somas, name="TOTAIS")
            linha_total["Funcionário"] = "TOTAIS"
            df_contrato = pd.concat([df_contrato, linha_total.to_frame().T], ignore_index=True)

            df_contrato.columns = [str(c).replace("_", " ") for c in df_contrato.columns]
            df_contrato.to_excel(writer, sheet_name=sheet_name, index=False, startrow=STARTROW_DADOS)

            # cabeçalho padronizado + boletins listados
            ws = writer.sheets[sheet_name]
            df_b = app.dados_df.copy()
            df_b["DATA"] = pd.to_datetime(df_b["DATA"], errors="coerce")
            mask = (
                (df_b["Contrato"] == contrato_id)
                & (df_b["DATA"].dt.date >= dt_ini.date())
                & (df_b["DATA"].dt.date <= dt_fim.date())
            )
            df_b = df_b.loc[mask]
            if "BOLETIM" in df_b.columns:
                boletins_txt = ", ".join(map(str, pd.to_numeric(df_b["BOLETIM"], errors="coerce").dropna().astype(int).sort_values().unique().tolist())) or "-"
            else:
                boletins_txt = "-"

            aplicar_cabecalho_excel(
                ws,
                titulo="Relatório Consolidado de Totais",
                contrato=contrato_id,
                dt_ini=dt_ini,
                dt_fim=dt_fim,
                boletins=boletins_txt,
                df_cols=list(df_contrato.columns),
            )
            formatar_planilha_excel(ws, startrow=STARTROW_DADOS)

        # ---- Resumo Geral ----
        df_resumo = df_all.copy()
        df_resumo["Contrato"] = df_resumo.get("Contratos", "").apply(_ult5_contratos)

        # Boletins por funcionário (no período + nos contratos selecionados)
        boletins_por_func = {}
        df_b_all = app.dados_df.copy()
        df_b_all["DATA"] = pd.to_datetime(df_b_all["DATA"], errors="coerce")
        mask_all = (
            df_b_all["DATA"].dt.date.between(dt_ini.date(), dt_fim.date())
            & df_b_all["Contrato"].isin(lista_contratos)
        )
        df_b_all = df_b_all.loc[mask_all]
        if "BOLETIM" in df_b_all.columns:
            for nome in df_resumo["Funcionário"].dropna().unique():
                nums = (
                    pd.to_numeric(df_b_all.loc[df_b_all["Funcionário"] == nome, "BOLETIM"], errors="coerce")
                    .dropna().astype(int).sort_values().unique().tolist()
                )
                boletins_por_func[nome] = ", ".join(map(str, nums)) if nums else ""
        df_resumo["Boletim"] = df_resumo["Funcionário"].map(boletins_por_func).fillna("")

        cols_rest = [c for c in df_resumo.columns if c not in ("Funcionário", "Contratos", "Contrato", "Boletim")]
        df_resumo = df_resumo[["Funcionário", "Contrato", "Boletim"] + cols_rest]

        num_cols = df_resumo.select_dtypes(include="number").columns
        totais = df_resumo[num_cols].sum(numeric_only=True)
        linha_total = pd.Series(index=df_resumo.columns, dtype=object)
        linha_total["Funcionário"] = "TOTAIS GERAIS"
        for c in num_cols:
            try:
                linha_total[c] = round(float(totais.get(c, 0) or 0), 2)
            except Exception:
                linha_total[c] = totais.get(c, 0)
        linha_total["Contrato"] = ""; linha_total["Boletim"] = ""
        df_resumo = pd.concat([df_resumo, linha_total.to_frame().T], ignore_index=True)

        df_resumo.columns = [str(c).replace("_", " ") for c in df_resumo.columns]
        df_resumo.to_excel(writer, sheet_name="Resumo Geral", index=False, startrow=STARTROW_DADOS)

        ws_resumo = writer.sheets["Resumo Geral"]
        # Boletins gerais do período
        if "BOLETIM" in df_b_all.columns:
            all_nums = pd.to_numeric(df_b_all["BOLETIM"], errors="coerce").dropna().astype(int).sort_values().unique().tolist()
            boletins_all_text = ", ".join(map(str, all_nums)) if all_nums else "-"
        else:
            boletins_all_text = "-"

        aplicar_cabecalho_excel(
            ws_resumo,
            titulo="Resumo Geral - Totais por Funcionário",
            contrato=", ".join(map(str, lista_contratos)),
            dt_ini=dt_ini,
            dt_fim=dt_fim,
            boletins=boletins_all_text,
            df_cols=list(df_resumo.columns),
        )
        formatar_planilha_excel(ws_resumo, startrow=STARTROW_DADOS)

    # ================= Workbook apenas de DIFERENÇAS =================
    diff_cols_raw = [c for c in df_all.columns if str(c).startswith("Diferença ")]
    if diff_cols_raw:
        arquivo_dif = (arquivo_saida[:-5] + "_Diferencas.xlsx") if arquivo_saida.lower().endswith(".xlsx") else (arquivo_saida + "_Diferencas.xlsx")
        with pd.ExcelWriter(arquivo_dif, engine="openpyxl") as writer:
            used_d = set()
            for contrato_id in sorted(set(lista_contratos)):
                sheet_name = _sheet_name_unique(contrato_id, used_d)
                df_contrato = df_all[df_all["Contratos"].apply(lambda x: str(contrato_id) in str(x))].copy()
                if df_contrato.empty:
                    continue
                df_dif = df_contrato[["Funcionário"] + diff_cols_raw].copy()
                df_dif.columns = [str(c).replace("_", " ") for c in df_dif.columns]
                somas = df_dif.select_dtypes(include="number").sum()
                linha_total = pd.Series(somas, name="TOTAIS"); linha_total["Funcionário"] = "TOTAIS"
                df_dif = pd.concat([df_dif, linha_total.to_frame().T], ignore_index=True)

                df_dif.to_excel(writer, sheet_name=sheet_name, index=False, startrow=STARTROW_DADOS)
                ws = writer.sheets[sheet_name]

                # cabeçalho padrão
                aplicar_cabecalho_excel(
                    ws,
                    titulo="Relatório — DIFERENÇAS (Boletim − Ponto)",
                    contrato=contrato_id,
                    dt_ini=dt_ini,
                    dt_fim=dt_fim,
                    boletins="-",
                    df_cols=list(df_dif.columns),
                )
                formatar_planilha_excel(ws, startrow=STARTROW_DADOS)

                # negativos em vermelho + negrito (todas as colunas de diferença)
                max_row, max_col = ws.max_row, ws.max_column
                for r in range(STARTROW_DADOS + 1, max_row + 1):
                    for c in range(2, max_col + 1):  # pula col A (Funcionário)
                        cell = ws.cell(row=r, column=c)
                        if eh_negativo_str(cell.value):
                            cell.font = Font(color="FF0000", bold=True)

            # Resumo Geral (diferenças)
            df_res = df_all.copy()
            df_res["Contrato"] = df_res.get("Contratos", "").apply(_ult5_contratos)
            keep = ["Funcionário", "Contrato"] + diff_cols_raw
            df_res = df_res[keep]
            df_res.columns = [str(c).replace("_", " ") for c in df_res.columns]

            num_cols = df_res.select_dtypes(include="number").columns
            tot = df_res[num_cols].sum(numeric_only=True)
            linha_total = pd.Series(index=df_res.columns, dtype=object)
            linha_total["Funcionário"] = "TOTAIS GERAIS"; linha_total["Contrato"] = ""
            for c in num_cols:
                try:
                    linha_total[c] = round(float(tot.get(c, 0) or 0), 2)
                except Exception:
                    linha_total[c] = tot.get(c, 0)
            df_res = pd.concat([df_res, linha_total.to_frame().T], ignore_index=True)

            df_res.to_excel(writer, sheet_name="Resumo Geral", index=False, startrow=STARTROW_DADOS)
            ws = writer.sheets["Resumo Geral"]
            aplicar_cabecalho_excel(
                ws,
                titulo="Resumo Geral — DIFERENÇAS (Boletim − Ponto)",
                contrato=", ".join(map(str, lista_contratos)),
                dt_ini=dt_ini,
                dt_fim=dt_fim,
                boletins="-",
                df_cols=list(df_res.columns),
            )
            formatar_planilha_excel(ws, startrow=STARTROW_DADOS)

            max_row, max_col = ws.max_row, ws.max_column
            for r in range(STARTROW_DADOS + 1, max_row + 1):
                for c in range(3, max_col + 1):  # pula Nome e Contrato
                    cell = ws.cell(row=r, column=c)
                    if eh_negativo_str(cell.value):
                        cell.font = Font(color="FF0000", bold=True)

    messagebox.showinfo("Sucesso", f"Relatório de totais consolidados salvo em:\n{arquivo_saida}")


def _safe_filename(name: str) -> str:
    name = str(name).strip()
    name = re.sub(r'[\\/:*?"<>|]+', "_", name)
    return name[:120] or "Contrato"

def exportar_totais_consolidados_por_contrato_em_arquivos(app, lista_contratos):
    """
    Gera UM PAR de arquivos por contrato:
    • Totais_Consolidados_<contrato>.xlsx
    • Totais_Consolidados_<contrato>_Diferencas.xlsx
    Cada arquivo tem uma única aba com os mesmos dados/estilo dos consolidados atuais.
    """
    try:
        from utils.formatadores import (
            aplicar_cabecalho_excel,
            formatar_planilha_excel,
            STARTROW_DADOS_XLSX as STARTROW_DADOS,
            eh_negativo_str,
        )
    except Exception as e:
        messagebox.showerror("Erro", f"Falha ao importar utils/formatadores:\n{e}")
        return

    # pasta de saída
    pasta = filedialog.askdirectory(title="Escolha a pasta para salvar os arquivos por contrato")
    if not pasta:
        return

    # período da UI
    try:
        dt_ini = datetime.strptime(app.cal_data_inicial.entry.get(), "%d/%m/%Y")
        dt_fim = datetime.strptime(app.cal_data_final.entry.get(), "%d/%m/%Y")
    except Exception:
        messagebox.showerror("Erro", "Período inválido na interface.")
        return

    # === prepara base de totais por funcionário (mesma lógica do consolidado atual) ===
    df_boletim_sel = app.dados_df[app.dados_df["Contrato"].isin(lista_contratos)].copy()
    funcionarios = sorted(df_boletim_sel["Funcionário"].dropna().unique())

    resultados = []
    for nome in funcionarios:
        totais = app._calcular_totais_funcionario(nome)
        if not totais:
            continue
        contratos_do_func = df_boletim_sel[df_boletim_sel["Funcionário"] == nome]["Contrato"].unique()
        totais["Contratos"] = ", ".join(map(str, contratos_do_func))
        resultados.append(totais)

    if not resultados:
        messagebox.showwarning("Aviso", "Nenhum dado encontrado para os contratos selecionados.")
        return

    df_all = pd.DataFrame(resultados)

    gerados = []

    # ===== loop por contrato → 2 arquivos =====
    for contrato_id in sorted(set(lista_contratos)):
        # ---- Totais (semelhante à aba por contrato do consolidado) ----
        df_contrato = df_all[df_all["Contratos"].apply(lambda x: str(contrato_id) in str(x))].copy()
        if not df_contrato.empty:
            df_contrato.drop(columns=["Contratos"], inplace=True, errors="ignore")
            # linha TOTAIS
            somas = df_contrato.select_dtypes(include="number").sum()
            linha_total = pd.Series(somas, name="TOTAIS")
            linha_total["Funcionário"] = "TOTAIS"
            df_contrato = pd.concat([df_contrato, linha_total.to_frame().T], ignore_index=True)

            df_contrato.columns = [str(c).replace("_", " ") for c in df_contrato.columns]

            # cria workbook e escreve
            wb = Workbook(); ws = wb.active; ws.title = str(contrato_id)

            # escrevendo a partir de STARTROW_DADOS
            for j, col in enumerate(df_contrato.columns, start=1):
                ws.cell(row=STARTROW_DADOS, column=j, value=col)
            for i, (_, row) in enumerate(df_contrato.iterrows(), start=STARTROW_DADOS + 1):
                for j, col in enumerate(df_contrato.columns, start=1):
                    ws.cell(row=i, column=j, value=row[col])

            # boletins do contrato no período
            df_b = app.dados_df.copy()
            df_b["DATA"] = pd.to_datetime(df_b["DATA"], errors="coerce")
            mask = (df_b["Contrato"] == contrato_id) & (df_b["DATA"].dt.date.between(dt_ini.date(), dt_fim.date()))
            df_b = df_b.loc[mask]
            if "BOLETIM" in df_b.columns:
                boletins_txt = ", ".join(
                    map(str, pd.to_numeric(df_b["BOLETIM"], errors="coerce").dropna().astype(int).sort_values().unique().tolist())
                ) or "-"
            else:
                boletins_txt = "-"

            aplicar_cabecalho_excel(
                ws,
                titulo="Relatório Consolidado de Totais",
                contrato=contrato_id,
                dt_ini=dt_ini,
                dt_fim=dt_fim,
                boletins=boletins_txt,
                df_cols=list(df_contrato.columns),
            )
            formatar_planilha_excel(ws, startrow=STARTROW_DADOS)

            # salvar
            nome_tot = _safe_filename(f"Totais_Consolidados_{contrato_id}.xlsx")
            caminho_tot = f"{pasta}/{nome_tot}"
            wb.save(caminho_tot)
            gerados.append(caminho_tot)

        # ---- Diferenças (semelhante ao arquivo _Diferencas.xlsx) ----
        diff_cols = [c for c in df_all.columns if str(c).startswith("Diferença ")]
        if diff_cols:
            df_dif = df_all[df_all["Contratos"].apply(lambda x: str(contrato_id) in str(x))].copy()
            if not df_dif.empty:
                df_dif = df_dif[["Funcionário"] + diff_cols].copy()
                df_dif.columns = [str(c).replace("_", " ") for c in df_dif.columns]
                # linha TOTAIS
                somas = df_dif.select_dtypes(include="number").sum()
                linha_total = pd.Series(somas, name="TOTAIS"); linha_total["Funcionário"] = "TOTAIS"
                df_dif = pd.concat([df_dif, linha_total.to_frame().T], ignore_index=True)

                wb_d = Workbook(); ws_d = wb_d.active; ws_d.title = str(contrato_id)
                for j, col in enumerate(df_dif.columns, start=1):
                    ws_d.cell(row=STARTROW_DADOS, column=j, value=col)
                for i, (_, row) in enumerate(df_dif.iterrows(), start=STARTROW_DADOS + 1):
                    for j, col in enumerate(df_dif.columns, start=1):
                        ws_d.cell(row=i, column=j, value=row[col])

                aplicar_cabecalho_excel(
                    ws_d,
                    titulo="Relatório — DIFERENÇAS (Boletim − Ponto)",
                    contrato=contrato_id,
                    dt_ini=dt_ini,
                    dt_fim=dt_fim,
                    boletins="-",
                    df_cols=list(df_dif.columns),
                )
                formatar_planilha_excel(ws_d, startrow=STARTROW_DADOS)

                # negativos em vermelho + negrito (todas as colunas de diferença)
                max_row, max_col = ws_d.max_row, ws_d.max_column
                for r in range(STARTROW_DADOS + 1, max_row + 1):
                    for c in range(2, max_col + 1):  # pula col A (Funcionário)
                        if eh_negativo_str(ws_d.cell(row=r, column=c).value):
                            ws_d.cell(row=r, column=c).font = Font(color="FF0000", bold=True)

                nome_dif = _safe_filename(f"Totais_Consolidados_{contrato_id}_Diferencas.xlsx")
                caminho_dif = f"{pasta}/{nome_dif}"
                wb_d.save(caminho_dif)
                gerados.append(caminho_dif)

    if gerados:
        msg = "Arquivos gerados:\n- " + "\n- ".join(gerados[:14])
        if len(gerados) > 14:
            msg += f"\n... (+{len(gerados)-14})"
        messagebox.showinfo("Sucesso", msg)
    else:
        messagebox.showwarning("Aviso", "Nada foi gerado (dados vazios ou contratos sem registros).")