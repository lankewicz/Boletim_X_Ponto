# utils/threading_ui.py
# Módulo utilitário para exibir uma janela modal com progresso duplo
# enquanto uma função é executada em segundo plano.
# Caminho: relatorio_horas/utils/threading_ui.py
#
# Autor: Valdinei Lankewicz
# Criado em: 22/07/2025
# Histórico:
# - 22/07/2025: Versão inicial com barra de progresso principal e secundária
# - 28/07/2025: Adicionado suporte à exibição do nome do arquivo e boletim

import threading
from tkinter import Label, Toplevel

from ttkbootstrap import Progressbar


def mostrar_modal_progresso(janela, mensagem, funcao_execucao):
    modal = Toplevel(janela)
    modal.title("Processando")
    modal.geometry("400x180")
    modal.transient(janela)
    modal.grab_set()
    modal.resizable(False, False)

    label = Label(
        modal,
        text=f"🔄 {mensagem}\nIsso pode levar alguns instantes.",
        font=("Segoe UI", 11),
        justify="center",
    )
    label.pack(padx=10, pady=(10, 5))

    # NOVO: Label que mostra o nome do arquivo e o boletim
    label_detalhe = Label(modal, text="", font=("Segoe UI", 9), justify="center", wraplength=380)
    label_detalhe.pack(pady=(0, 5))

    progress_total = Progressbar(modal, orient="horizontal", mode="determinate", maximum=100)
    progress_total.pack(fill="x", padx=20, pady=(5, 5))

    progress_arquivo = Progressbar(modal, orient="horizontal", mode="determinate", maximum=100)
    progress_arquivo.pack(fill="x", padx=20, pady=(0, 10))

    def set_total(percent):
        progress_total["value"] = max(0, min(100, percent))
        modal.update_idletasks()

    def set_arquivo(percent, info=""):
        progress_arquivo["value"] = max(0, min(100, percent))
        if info:
            label_detalhe.config(text=info)
        modal.update_idletasks()

    def thread_target():
        try:
            funcao_execucao(set_total, set_arquivo)
        finally:
            if modal.winfo_exists():
                modal.destroy()

    threading.Thread(target=thread_target, daemon=True).start()
