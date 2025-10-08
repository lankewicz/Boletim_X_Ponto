# interface/dialogos/formulario_geral.py
from __future__ import annotations
import tkinter as tk
import ttkbootstrap as ttkb
from tkinter import ttk
import pandas as pd

from utils.relatorios import relatorio_sa_por_registro_periodo

def _fmt_ptbr(v: float) -> str:
    try:
        s = f"{float(v):.2f}"
        return s.replace(".", ",")
    except Exception:
        return str(v)

class FormularioGeralDialog:
    def __init__(self, parent, df_periodo: pd.DataFrame, *, data_ini=None, data_fim=None):
        self.parent = parent
        self.df_base = df_periodo if df_periodo is not None else pd.DataFrame()
        self.data_ini = data_ini
        self.data_fim = data_fim

        self.win = ttkb.Toplevel(parent)
        self.win.title("Formulário Geral")
        self.win.geometry("1100x600")
        self.win.resizable(True, True)
        try:
            self.win.state("zoomed")  # maximiza no Windows
        except Exception:
            pass

        topo = ttkb.Frame(self.win, padding=8)
        topo.pack(fill=tk.X)
        ttkb.Label(topo, text=f"Período: {data_ini} a {data_fim}", bootstyle="secondary").pack(side=tk.LEFT)

        # tabela
        cols = ("Funcionário", "Registro", "Boletim", "Contrato", "S.A.")
        self.tree = ttk.Treeview(self.win, columns=cols, show="headings")
        self.tree.pack(fill=tk.BOTH, expand=True)

        self.tree.heading("Funcionário", text="Funcionário", command=lambda: self._ordenar_por("Funcionário"))
        self.tree.heading("Registro",    text="Registro",    command=lambda: self._ordenar_por("Registro"))
        self.tree.heading("Boletim",     text="Boletim",     command=lambda: self._ordenar_por("Boletim"))
        self.tree.heading("Contrato",    text="Contrato",    command=lambda: self._ordenar_por("Contrato"))
        self.tree.heading("S.A.",        text="S.A.",        command=lambda: self._ordenar_por("S.A.", numeric=True))

        for c, w in zip(cols, (360,120,120,120,120)):
            self.tree.column(c, width=w, anchor=tk.E if c=="S.A." else tk.W)

        self._carregar()

        ttkb.Button(self.win, text="Fechar", command=self.win.destroy).pack(side=tk.RIGHT, padx=8, pady=6)

    def _carregar(self):
        # gera o DF exatamente como o script de console, usando só o PERÍODO
        df = relatorio_sa_por_registro_periodo(self.df_base, self.data_ini, self.data_fim)

        # popula
        for row in self.tree.get_children():
            self.tree.delete(row)

        for _, r in df.iterrows():
            val = (
                r["Funcionário"],
                r["Registro"],
                r["BOLETIM"],
                r["Contrato"],
                _fmt_ptbr(r["S.A."]),
            )
            # destaca subtotal
            tag = ()
            if str(r["Funcionário"]).startswith("Subtotal "):
                tag = ("subtotal",)
            self.tree.insert("", "end", values=val, tags=tag)

        self.tree.tag_configure("subtotal", font=("Segoe UI", 9, "bold"))

        # autoajuste simples de largura
        self.win.after(100, self._autoajustar)

    def _autoajustar(self):
        for col in ("Funcionário", "Registro", "Boletim", "Contrato", "S.A."):
            self.tree.column(col, width=tkFontMeasure(self.tree, col))

    def _ordenar_por(self, coluna, numeric: bool=False):
        items = [(self.tree.set(i, coluna), i) for i in self.tree.get_children("")]
        if numeric:
            def key(t):
                s = str(t[0]).replace(".", "").replace(",", ".")
                try: return float(s)
                except: return 0.0
        else:
            def key(t): return str(t[0])

        # alterna asc/desc
        desc = (getattr(self, "_ord_col", None) == coluna and not getattr(self, "_ord_desc", False))
        items.sort(key=key, reverse=desc)
        for idx, (_, iid) in enumerate(items):
            self.tree.move(iid, "", idx)
        self._ord_col, self._ord_desc = coluna, desc


# utilzinho para medir largura “aproximada”
def tkFontMeasure(tree: ttk.Treeview, col: str) -> int:
    import tkinter.font as tkfont
    f = tkfont.nametofont(tree.cget("font"))
    header = tree.heading(col,"text")
    w = f.measure(header) + 24
    for iid in tree.get_children(""):
        text = tree.set(iid, col)
        w = max(w, f.measure(text) + 24)
    return min(max(w, 80), 600)
