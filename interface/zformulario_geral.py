# interface/relatorio_geral.py
from __future__ import annotations
import tkinter as tk
from tkinter import ttk
import pandas as pd
from utils.relatorios import montar_sa_por_registro

def _fmt_br(x) -> str:
    try:
        s = f"{float(x):,.2f}"
        return s.replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return str(x)

def abrir_relatorio_geral(parent, df_boletim: pd.DataFrame, *, data_ini, data_fim):
    """
    Abre janela com:
      Funcionário | Registro | Boletim | Contrato | S.A.
    filtrado pelo período recebido e com subtotais por funcionário.
    """
    # prepara dados
    df_out = montar_sa_por_registro(df_boletim, data_ini=data_ini, data_fim=data_fim, incluir_subtotal=True)

    win = tk.Toplevel(parent)
    win.title("Relatório S.A. — Funcionário / Registro / Boletim / Contrato")
    win.geometry("1100x600")
    win.resizable(True, True)

    cols = ("Funcionário", "Registro", "Boletim", "Contrato", "S.A.")
    tree = ttk.Treeview(win, columns=cols, show="headings")
    vsb = ttk.Scrollbar(win, orient="vertical", command=tree.yview)
    hsb = ttk.Scrollbar(win, orient="horizontal", command=tree.xview)
    tree.configure(yscroll=vsb.set, xscroll=hsb.set)

    # cabeçalhos
    for c in cols:
        tree.heading(c, text=c)
        anchor = "e" if c == "S.A." else "w"
        width = 130 if c in ("Registro", "Boletim", "Contrato") else 260 if c == "Funcionário" else 110
        tree.column(c, width=width, anchor=anchor, stretch=True)

    # insere linhas
    for _, r in df_out.iterrows():
        is_sub = str(r["Funcionário"]).startswith("Subtotal ")
        valores = [
            str(r["Funcionário"]),
            str(r["Registro"]) if pd.notna(r["Registro"]) else "",
            str(r["Boletim"]) if pd.notna(r["Boletim"]) else "",
            str(r["Contrato"]) if pd.notna(r["Contrato"]) else "",
            _fmt_br(r["S.A."]),
        ]
        tree.insert("", "end", values=valores, tags=("subtotal",) if is_sub else ())

    # destaque p/ subtotal
    tree.tag_configure("subtotal", font=("Segoe UI", 10, "bold"))

    # layout
    tree.grid(row=0, column=0, sticky="nsew")
    vsb.grid(row=0, column=1, sticky="ns")
    hsb.grid(row=1, column=0, sticky="ew")
    win.grid_rowconfigure(0, weight=1)
    win.grid_columnconfigure(0, weight=1)

    # botão fechar
    ttk.Button(win, text="Fechar", command=win.destroy).grid(row=2, column=0, sticky="e", padx=10, pady=8)
