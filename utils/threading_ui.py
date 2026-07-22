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


import threading
from tkinter import Label, Toplevel

from ttkbootstrap import Progressbar


def mostrar_modal_progresso(janela, mensagem, funcao_execucao, ao_concluir=None):
    modal = Toplevel(janela)
    modal.title("Processando")
    modal.geometry("420x190")
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

    label_detalhe = Label(modal, text="", font=("Segoe UI", 9), justify="center", wraplength=390)
    label_detalhe.pack(pady=(0, 5))

    progress_total = Progressbar(modal, orient="horizontal", mode="determinate", maximum=100)
    progress_total.pack(fill="x", padx=20, pady=(5, 5))

    progress_arquivo = Progressbar(modal, orient="horizontal", mode="determinate", maximum=100)
    progress_arquivo.pack(fill="x", padx=20, pady=(0, 10))

    def set_total(percent):
        def _update():
            if modal.winfo_exists():
                try:
                    progress_total["value"] = max(0, min(100, float(percent)))
                except Exception:
                    pass
        janela.after(0, _update)

    def set_arquivo(percent, info=""):
        def _update():
            if modal.winfo_exists():
                try:
                    progress_arquivo["value"] = max(0, min(100, float(percent)))
                    if info:
                        label_detalhe.config(text=info)
                except Exception:
                    pass
        janela.after(0, _update)

    def thread_target():
        erro = None
        resultado = None
        try:
            resultado = funcao_execucao(set_total, set_arquivo)
        except Exception as e:
            erro = e
            print(f"[ERRO] Erro no processamento em segundo plano: {e}")
        finally:
            def _finalizar():
                if modal.winfo_exists():
                    try:
                        modal.grab_release()
                    except Exception:
                        pass
                    modal.destroy()
                if ao_concluir:
                    ao_concluir(resultado, erro)
            janela.after(0, _finalizar)

    threading.Thread(target=thread_target, daemon=True).start()

