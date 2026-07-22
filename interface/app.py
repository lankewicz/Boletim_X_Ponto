# interface/app.py
# Contém a classe principal da interface gráfica: RelatorioHorasApp
# Caminho: interface/app.py
#
# Autor: Valdinei Lankewicz
# Criado em: 22/07/2025
# Histórico:
# - 28/07/2025: Adicionada lógica para carregar e exibir dados do ponto.
# - 10/08/2025: Ajustes na aba Comparação (barra preta, linhas extras Diferença, colunas fixas).
# - 10/08/2025: Loader de tarifas COPEL via JSON em ./data/tarifas_copel.json.

# --- imports padrão ---
import os
from datetime import datetime as dt
from difflib import get_close_matches
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import ttkbootstrap as ttkb

from extrator.leitor_ponto import processar_todos_csvs
from extrator.pasta_utils import carregar_ultima_pasta, salvar_ultima_pasta
from extrator.processador import consolidar_pdfs_em_excel

from interface.exportador import (
    exportar_para_excel,
    exportar_para_pdf,
    exportar_ponto_para_excel,
    exportar_ponto_para_pdf,
    exportar_comparacao_unificada,
    exportar_totais_mensais,
    exportar_totais_consolidados_excel,
    exportar_totais_consolidados_por_contrato_em_arquivos,   
)

from utils import verificador_dependencias as deps
from utils.threading_ui import mostrar_modal_progresso

from utils.formatadores import (
    finalizar_conjunto_comparacao,
    finalizar_sheet,
    formatar_tempo as fmt_tempo,
    ocultar_colunas,
    formatar_decimal,
)

from utils.constantes import HEADERS_VIZ, MAP_BOL, MAP_PTO
from utils.periodo import ler_datas_da_ui

from interface.relatorios import (
    build_boletim_grid,
    build_ponto_grid,
    build_comparacao_grids,
    abrir_relatorio_geral,
)


class RelatorioHorasApp:
    def __init__(self):
        # Arquivos/pastas usados pelo app
        self.arquivo_ultima_pasta = "ultima_pasta.txt"
        self.pasta_pdfs = carregar_ultima_pasta(self.arquivo_ultima_pasta)
        self.arquivo_ponto_parquet = "data/RegistroPonto.parquet"
        self.arquivo_relacao_nomes = "data/relacao_nomes.parquet"
        self.df_relacao = None

        # DataFrames e estado
        self.dados_df = None
        self.df_filtrado = None
        self.df_ponto = None
        self.data_minima = None
        self.data_maxima = None
        self.todos_funcionarios = []
        self.funcionarios_filtrados = []
        self._debounce_id = None

        # Widgets/controles
        self.janela = None
        self.tree = None
        self.tree_totais = None
        self.tree_ponto = None
        self.cal_data_inicial = None
        self.cal_data_final = None
        self.combo_funcionario = None
        self.status_bar = None
        self.progress_bar = None
        self.funcionario_boletim_label = None
        self.nome_funcionario_ponto = "N/D"
        self.origem_primaria = "boletim"
        self.is_syncing = False

        # Labels de status (evitar messagebox em fluxo normal)
        self.label_status_boletim = None
        self.label_status_ponto = None
        self.label_status_comparacao = None

        # Config de exibição de horas (True -> decimal; False -> HH:MM)
        self.modo_exibicao_decimal = True
        self.var_formato_hora = None  # será definida em widgets.criar_interface()

        # Colunas relevantes do Boletim (tela principal)
        self.colunas_visiveis = [
            "DATA",
            "HORA NORMAL",
            "H.E.",
            "H.E.D.",
            "H.E.N.",
            "H.E.N.D.",
            "H.N.",
            "BOLETIM",
            "S.A.",
        ]
        self.colunas_para_somar = [
            "HORA NORMAL",
            "H.E.",
            "H.E.D.",
            "H.E.N.",
            "H.E.N.D.",
            "H.N.",
            "S.A.",
        ]
        self.largura_colunas = {
            "DATA": 85,
            "SERV.": 80,
            "KM": 80,
            "HORA NORMAL": 100,
            "H.E.": 80,
            "H.E.D.": 80,
            "H.E.N.": 80,
            "H.E.N.D.": 80,
            "S.A.": 80,
            "H.N.": 80,
            "DESLOC.": 80,
            "PROD.": 80,
        }
        # ===== Cabeçalhos e mapeamentos centralizados =====
        self.HEADERS_VIZ = HEADERS_VIZ
        self.MAP_BOL = MAP_BOL
        self.MAP_PTO = MAP_PTO

        # Derivados
        self.headers_comparacao = ["Data"] + self.HEADERS_VIZ
        self.colunas_ponto      = ["Data"] + self.HEADERS_VIZ[:]                       # rótulos bonitos
        self.colunas_ponto_df   = ["Data"] + [self.MAP_PTO[h] for h in self.HEADERS_VIZ]  # nomes reais no DF

        self.df_ponto_filtrado = pd.DataFrame()  # usado pelos exportadores de ponto (filtrado)

        ## self para filtro contrato
        self.combo_contrato = None
        self.var_filtrar_contrato = None


        # Inicializa UI e carrega dados
        self.inicializar_aplicacao()

    # ---------------- Helpers de formatação ----------------
    def formatar_tempo(self, valor):
        """Wrapper fino para o formatador centralizado (vírgula para decimais)."""
        return fmt_tempo(valor, modo_decimal=self.modo_exibicao_decimal, separador_decimal=",")

    def _fmt_decimal(self, valor, casas=2, zero_vira_traco=False):
        """Formata número NÃO-hora usando o utilitário central (vírgula como separador)."""
        return formatar_decimal(
            valor,
            casas=casas,
            separador=",",
            zero_vira_traco=zero_vira_traco,
        )
    
        # ---------------- Ações de UI ----------------
    def trocar_formato_hora(self, *_args, **_kwargs):
        """
        Alterna o formato de exibição (decimal ↔ HH:MM) e atualiza as abas.
        Se existir self.var_formato_hora (BooleanVar), usa o valor dela.
        """
        try:
            # Se a checkbox/BooleanVar existir: True => decimal, False => HH:MM
            self.modo_exibicao_decimal = bool(self.var_formato_hora.get())
        except Exception:
            # Se a var ainda não existir, apenas alterna
            self.modo_exibicao_decimal = not getattr(self, "modo_exibicao_decimal", True)

        # Re-renderiza as abas (sem quebrar a UI se ainda não estiver pronta)
        try:
            self.atualizar_todas_abas()
        except Exception:
            pass

    # Alias opcional (se algum lugar chamar por outro nome)
    def alternar_formato_hora(self, *a, **k):
        return self.trocar_formato_hora(*a, **k)


    # ===============================================================
    # CÁLCULO DE TOTAIS POR FUNCIONÁRIO (usado em exportações/relatórios)
    # ===============================================================
    def _calcular_totais_funcionario(self, funcionario_nome):
        """
        Calcula os totais de um funcionário usando a base de relação (CPF/PIS)
        como prioridade e a similaridade de nomes como fallback.
        Retorna um dicionário com totais de Boletim, Ponto e Diferença.
        """
        try:
            data_ini = self.cal_data_inicial.get_date()
            data_fim = self.cal_data_final.get_date()

            df_boletim_func = self.dados_df[
                (self.dados_df["Funcionário"] == funcionario_nome)
                & (self.dados_df["DATA"] >= data_ini)
                & (self.dados_df["DATA"] <= data_fim)
            ]

            df_ponto_func = pd.DataFrame()
            if self.df_relacao is not None and not self.df_relacao.empty:
                mapeamento = self.df_relacao[self.df_relacao["Nome_Boletim"] == funcionario_nome]
                if not mapeamento.empty:
                    mapeamento_valido = mapeamento.iloc[0]
                    cpf = mapeamento_valido["CPF_Ponto"]
                    pis = mapeamento_valido["PIS_Ponto"]
                    if pd.notna(cpf) and not self.df_ponto[self.df_ponto["CPF"] == cpf].empty:
                        df_ponto_func = self.df_ponto[self.df_ponto["CPF"] == cpf]
                        if self.status_bar:
                            self.status_bar.config(
                                text=f"Relação usada: {funcionario_nome} → CPF {cpf}"
                            )
                    elif pd.notna(pis) and not self.df_ponto[self.df_ponto["PIS"] == pis].empty:
                        df_ponto_func = self.df_ponto[self.df_ponto["PIS"] == pis]
                        if self.status_bar:
                            self.status_bar.config(
                                text=f"Relação usada: {funcionario_nome} → PIS {pis}"
                            )

            if df_ponto_func.empty:
                nomes_ponto = self.df_ponto["Nome"].dropna().unique()
                nome_similar_list = get_close_matches(
                    funcionario_nome.upper(), [n.upper() for n in nomes_ponto], n=1, cutoff=0.75
                )
                if nome_similar_list:
                    nome_ponto_usado = next(
                        (n for n in nomes_ponto if n.upper() == nome_similar_list[0]), None
                    )
                    if nome_ponto_usado:
                        df_ponto_func = self.df_ponto[self.df_ponto["Nome"] == nome_ponto_usado]
                        if self.status_bar:
                            self.status_bar.config(
                                text=f"Relação por similaridade: {funcionario_nome} → {nome_ponto_usado}"
                            )

            if not df_ponto_func.empty:
                df_ponto_func = df_ponto_func[
                    (pd.to_datetime(df_ponto_func["Data"]) >= data_ini)
                    & (pd.to_datetime(df_ponto_func["Data"]) <= data_fim)
                ]

            if df_boletim_func.empty and df_ponto_func.empty:
                return None

            campos = {
                "HORA_NORMAL": ("HORA NORMAL", "Total Normais"),
                "HORA_NOTURNA": ("H.N.", "Total Noturno"),
                "EXTRA_50_D": ("H.E.", "Extra 50%D"),
                "EXTRA_100_D": ("H.E.D.", "Extra 100%D"),
                "EXTRA_50_N": ("H.E.N.", "Extra 50%N"),
                "EXTRA_100_N": ("H.E.N.D.", "Extra 100%N"),
            }
            totais = {"Funcionário": funcionario_nome}
            for nome_amigavel, (campo_b, campo_p) in campos.items():
                soma_b = df_boletim_func[campo_b].sum() if campo_b in df_boletim_func else 0
                soma_p = df_ponto_func[campo_p].sum() if campo_p in df_ponto_func else 0
                totais[f"Boletim {nome_amigavel}"] = round(soma_b, 2)
                totais[f"Ponto {nome_amigavel}"] = round(soma_p, 2)
                totais[f"Diferença {nome_amigavel}"] = round(soma_b - soma_p, 2)
            return totais
        except Exception as e:
            print(f"Erro ao calcular totais para '{funcionario_nome}': {e}")
            return None

    # ===============================================================
    # RELAÇÃO DE NOMES (CPF/PIS) — utilidades
    # ===============================================================
    def _schema_relacao_minimo(self):
        return ["Nome_Boletim", "Nome_Ponto_Mapeado", "CPF_Ponto", "PIS_Ponto"]

    def _sanear_relacao_df(self, df):
        import pandas as pd

        cols = self._schema_relacao_minimo()
        if df is None or df.empty:
            return pd.DataFrame({c: [] for c in cols})
        for c in cols:
            if c not in df.columns:
                df[c] = ""
        # remove duplicatas por Nome_Boletim, mantendo última edição
        df = df.drop_duplicates(subset=["Nome_Boletim"], keep="last").copy()
        # normaliza tipos básicos
        for c in ["CPF_Ponto", "PIS_Ponto", "Nome_Boletim", "Nome_Ponto_Mapeado"]:
            df[c] = df[c].astype(str).fillna("").str.strip()
        return df[cols + [c for c in df.columns if c not in cols]]

    def _carregar_relacao_nomes(self):
        import os

        import pandas as pd

        try:
            os.makedirs(os.path.dirname(self.arquivo_relacao_nomes), exist_ok=True)
            if os.path.exists(self.arquivo_relacao_nomes):
                df = pd.read_parquet(self.arquivo_relacao_nomes)
            else:
                # cria arquivo vazio com esquema mínimo
                df = self._sanear_relacao_df(None)
                df.to_parquet(self.arquivo_relacao_nomes, index=False)
            self.df_relacao = self._sanear_relacao_df(df)
        except Exception as e:
            # não quebra a app; mantém df_relacao vazio mas válido
            try:
                import pandas as pd

                self.df_relacao = self._sanear_relacao_df(None)
            except Exception:
                self.df_relacao = None
            from tkinter import messagebox

            messagebox.showwarning(
                "Relação de Nomes",
                f"Não foi possível ler/criar '{self.arquivo_relacao_nomes}'.\n"
                f"Continuando sem relação.\n\nDetalhes: {e}",
            )

    def _get_funcionarios_filtrados(self) -> list[str]:
        """Retorna a lista atual da combobox (filtrada pelo período)."""
        try:
            vals = self.combo_funcionario.cget("values")
        except Exception:
            return []
        if vals is None:
            return []
        if isinstance(vals, (tuple, list)):
            seq = list(vals)
        elif isinstance(vals, str):
            # fallback raro: pode vir como string; tenta separar por vírgula
            seq = [s.strip() for s in vals.split(",") if s.strip()]
        else:
            try:
                seq = list(vals)
            except Exception:
                seq = []
        return [v for v in seq if isinstance(v, str) and v.strip()]


    def _on_period_changed(self, *_):
        """Debounce: reagenda atualização da combo quando período muda."""
        try:
            if getattr(self, "_debounce_id", None):
                self.janela.after_cancel(self._debounce_id)
        except Exception:
            pass
        # chama em 200ms para evitar disparos múltiplos no mesmo gesto
        self._debounce_id = self.janela.after(200, lambda: self.atualizar_combo_funcionarios())


    # ===============================================================
    # INICIALIZAÇÃO DA INTERFACE E CARREGAMENTO INICIAL
    # ===============================================================

    def inicializar_aplicacao(self):
        """Cria a interface e carrega dados iniciais de Boletim e Ponto."""
        from interface.widgets import criar_interface

        # 1) Construir UI
        self.janela = criar_interface(self)

        # 2) Reporter do verificador -> status bar
        def _gui_report(msg: str):
            if self.status_bar:
                self.status_bar.config(text=msg)
                try:
                    self.janela.update_idletasks()
                except Exception:
                    pass

        deps.set_reporter(_gui_report)
        try:
            deps.verificar_dependencias_automaticamente(".")
        except Exception:
            if self.status_bar:
                self.status_bar.config(text="Aviso: verificação de dependências não pôde ser concluída.")

        # 3) Relação de nomes
        try:
            self._carregar_relacao_nomes()
            if self.status_bar:
                txt = self.status_bar.cget("text") or ""
                self.status_bar.config(text=(txt + " | Relação de nomes carregada.").strip(" |"))
        except Exception:
            pass

        # 4) Boletim (parquet)
        caminho_parquet_boletim = "data/RegistroBoletim.parquet"
        if os.path.exists(caminho_parquet_boletim):
            self.dados_df, self.data_minima, self.data_maxima = self.carregar_dados_boletim(
                caminho_parquet_boletim
            )
            if self.dados_df is not None:
                # Define datas padrão e prepara filtros iniciais
                self.atualizar_interface_filtros()
                if self.status_bar:
                    self.status_bar.config(text="Dados de boletim carregados.")
        else:
            if self.status_bar:
                self.status_bar.config(text="Dados de boletim não encontrados. Use 'Atualizar Dados'.")

        # 5) Ponto (parquet)
        self.recarregar_dados_ponto()

        # 6) Atualiza combo de Funcionário conforme período atual (primeira carga)
        try:
            self.atualizar_combo_funcionarios(silent=True)
        except Exception:
            pass

        # 7) Binds para reagir à mudança de período
        #    - Seleção via calendário
        for w in (getattr(self, "cal_data_inicial", None), getattr(self, "cal_data_final", None)):
            if w:
                try:
                    w.bind("<<DateEntrySelected>>", self._on_period_changed)
                    # quando o usuário DIGITA a data e sai do campo
                    w.bind("<FocusOut>", self._on_period_changed)
                except Exception:
                    pass

        #    - Se houver Combobox de mês/ano
        for w in (getattr(self, "combo_mes", None), getattr(self, "combo_ano", None)):
            if w:
                try:
                    w.bind("<<ComboboxSelected>>", self._on_period_changed)
                except Exception:
                    pass

        # 8) Loop principal
        self.janela.mainloop()


    # ===============================================================
    # UTILIDADES / WORKERS
    # ===============================================================

    def executar_extrator_boletim(self, forcar_dialogo=False):
        """Abre pasta com PDFs de boletim, processa e recarrega parquet consolidado."""

        def worker(set_total, set_arquivo):
            pasta = self.pasta_pdfs
            if forcar_dialogo or not pasta:
                pasta = filedialog.askdirectory(
                    title="Selecione a pasta com os relatórios PDF de BOLETIM"
                )
                if not pasta:
                    return
                salvar_ultima_pasta(self.arquivo_ultima_pasta, pasta)
            self.pasta_pdfs = pasta
            caminho_parquet_boletim = consolidar_pdfs_em_excel(pasta, set_total, set_arquivo)
            self.dados_df, self.data_minima, self.data_maxima = self.carregar_dados_boletim(
                caminho_parquet_boletim
            )
            if self.dados_df is not None:
                self.atualizar_interface_filtros()
                if self.status_bar:
                    self.status_bar.config(text="Dados de boletim atualizados com sucesso.")

        mostrar_modal_progresso(self.janela, "Processando relatórios de Boletim...", worker)

    def executar_extrator_ponto(self):
        """Processa CSVs de ponto na pasta ./ponto e recarrega parquet."""

        def worker(set_total, set_arquivo):
            pasta_ponto = "ponto"
            if not os.path.exists(pasta_ponto):
                messagebox.showerror(
                    "Pasta não encontrada",
                    f"A pasta '{pasta_ponto}' não foi encontrada. Crie-a e coloque os dados de origem de ponto dentro dela.",
                )
                return

            processar_todos_csvs(pasta_ponto, set_total, set_arquivo)
            self.recarregar_dados_ponto()
            if self.status_bar:
                self.status_bar.config(text="Dados de ponto atualizados com sucesso.")

        mostrar_modal_progresso(self.janela, "Processando cartões de ponto...", worker)

    def recarregar_dados_ponto(self):
        """Recarrega o parquet do Ponto em self.df_ponto."""
        if os.path.exists(self.arquivo_ponto_parquet):
            try:
                self.df_ponto = pd.read_parquet(self.arquivo_ponto_parquet)
                # robusto: sem format fixo; parquet pode vir como date/datetime
                self.df_ponto["Data"] = pd.to_datetime(self.df_ponto["Data"], errors="coerce")
                if self.status_bar:
                    txt = (self.status_bar.cget("text") or "")
                    self.status_bar.config(text=(txt + " | Dados de ponto carregados.").strip(" |"))
                    
            except Exception as e:
                messagebox.showerror(
                    "Erro ao carregar Ponto",
                    f"Não foi possível ler '{self.arquivo_ponto_parquet}':\n{e}",
                )
        else:
            if self.status_bar:
                txt = (self.status_bar.cget("text") or "")
                self.status_bar.config(text=(txt + " | Dados de ponto não encontrados.").strip(" |"))


    def carregar_dados_boletim(self, caminho):
        """Carrega parquet do boletim e normaliza a coluna DATA."""
        try:
            df = pd.read_parquet(caminho)
            df["DATA"] = pd.to_datetime(df["DATA"], dayfirst=True, errors="coerce")
            df.dropna(subset=["DATA"], inplace=True)
            return df, df["DATA"].min().date(), df["DATA"].max().date()
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar dados do Boletim: {e}")
            return None, None, None

    def atualizar_combo_funcionarios(self, silent: bool = False):
        """
        Atualiza a combo de funcionários para listar apenas quem aparece nos boletins
        do período selecionado na UI.
        """
        import pandas as pd

        try:
            # 1) Ler período da UI
            data_ini, data_fim = ler_datas_da_ui(self.cal_data_inicial, self.cal_data_final)

            # 2) Garantir que há boletins carregados e colunas esperadas
            if getattr(self, "dados_df", None) is None or self.dados_df.empty:
                try:
                    self.combo_funcionario.configure(values=[])
                    self.combo_funcionario.set("")
                except Exception:
                    pass
                if not silent and self.status_bar:
                    self.status_bar.config(text="Nenhum boletim carregado.")
                return

            df = self.dados_df.copy()
            if "DATA" not in df.columns or "Funcionário" not in df.columns:
                try:
                    self.combo_funcionario.configure(values=[])
                    self.combo_funcionario.set("")
                except Exception:
                    pass
                if not silent and self.status_bar:
                    self.status_bar.config(text="Boletim sem colunas 'DATA' ou 'Funcionário'.")
                return

            # 3) Normalizar DATA (se necessário) e filtrar por período
            if not pd.api.types.is_datetime64_any_dtype(df["DATA"]):
                df["DATA"] = pd.to_datetime(df["DATA"], errors="coerce")

            df_mes = df[(df["DATA"].dt.date >= data_ini) & (df["DATA"].dt.date <= data_fim)]
            nomes = sorted(df_mes["Funcionário"].dropna().astype(str).unique().tolist())

            # 3.1) Se o filtro por contrato estiver ativo, restringe o DF
            try:
                if self.var_filtrar_contrato and bool(self.var_filtrar_contrato.get()):
                    contrato_sel = (self.combo_contrato.get() or "").strip()
                    if contrato_sel and "Contrato" in df_mes.columns:
                        df_mes = df_mes[df_mes["Contrato"] == contrato_sel]
            except Exception:
                pass

            # (2) agora sim, gere a lista de nomes
            nomes = sorted(df_mes["Funcionário"].dropna().astype(str).unique().tolist())

            # 4) Atualizar combo (limpa seleção se não pertencer ao período)
            try:
                self.combo_funcionario.configure(values=nomes)
                atual = (self.combo_funcionario.get() or "").strip()
                if not nomes:
                    self.combo_funcionario.set("")
                elif atual not in nomes:
                    self.combo_funcionario.set(nomes[0])  # pré-seleciona o primeiro do período
            except Exception:
                pass

            if not silent and self.status_bar:
                self.status_bar.config(text=f"{len(nomes)} funcionário(s) no período selecionado.")
        except Exception as e:
            # Nunca deixar a UI cair
            try:
                if self.status_bar:
                    self.status_bar.config(text=f"Erro ao atualizar funcionários: {e}")
            except Exception:
                pass


    # ===============================================================
    # DIÁLOGOS DE EXPORTAÇÃO
    # ===============================================================
    def mostrar_dialogo_exportacao(self):
        """Dialog simples de exportação do relatório atual (Excel/PDF)."""
        if self.df_filtrado is None or self.df_filtrado.empty:
            messagebox.showwarning("Aviso", "Nenhum relatório gerado. Gere um relatório primeiro.")
            return
        dialog = ttkb.Toplevel(self.janela)
        dialog.title("Exportar Relatório")
        largura = 380
        altura = 230
        x = (dialog.winfo_screenwidth() // 2) - (largura // 2)
        y = (dialog.winfo_screenheight() // 2) - (altura // 2)
        dialog.geometry(f"{largura}x{altura}+{x}+{y}")

        dialog.resizable(False, False)
        dialog.transient(self.janela)
        dialog.grab_set()
        label = ttkb.Label(dialog, text="Escolha o formato para exportar:")
        label.pack(pady=10)
        frame_botoes = ttkb.Frame(dialog)
        frame_botoes.pack(pady=5)
        ttkb.Button(
            frame_botoes,
            text="Exportar para Excel",
            command=lambda: [exportar_para_excel(self), dialog.destroy()],
            bootstyle="success",
        ).pack(side=tk.LEFT, padx=5)
        ttkb.Button(
            frame_botoes,
            text="Exportar para PDF",
            command=lambda: [exportar_para_pdf(self), dialog.destroy()],
            bootstyle="danger",
        ).pack(side=tk.LEFT, padx=5)
        ttkb.Button(
            frame_botoes, text="Cancelar", command=dialog.destroy, bootstyle="secondary"
        ).pack(side=tk.LEFT, padx=5)

    def mostrar_dialogo_exportacao_unificada(self):
        """Dialog com três opções: funcionário atual, contrato atual, todos (por contrato)."""
        dialog = ttkb.Toplevel(self.janela)
        dialog.title("Opções de Exportação")
        largura = 380
        altura = 230
        x = (dialog.winfo_screenwidth() // 2) - (largura // 2)
        y = (dialog.winfo_screenheight() // 2) - (altura // 2)
        dialog.geometry(f"{largura}x{altura}+{x}+{y}")
        
        dialog.resizable(False, False)
        dialog.transient(self.janela)
        dialog.grab_set()

        label = ttkb.Label(dialog, text="Escolha o escopo da exportação:", font=("Segoe UI", 10))
        label.pack(pady=10)

        frame_botoes = ttkb.Frame(dialog)
        frame_botoes.pack(pady=5, fill=tk.X, padx=20)

        ttkb.Button(
            frame_botoes,
            text="Apenas funcionário atual",
            command=lambda: [dialog.destroy(), exportar_comparacao_unificada(self)],
            bootstyle="info",
        ).pack(fill=tk.X, pady=3)

        ttkb.Button(
            frame_botoes,
            text="Todos do contrato atual",
            command=lambda: [
                dialog.destroy(),
                self.executar_exportacao_consolidada(modo="contrato_atual"),
            ],
            bootstyle="secondary",
        ).pack(fill=tk.X, pady=3)

        ttkb.Button(
            frame_botoes,
            text="Todos os dados (agrupado por contrato)",
            command=lambda: [dialog.destroy(), self.executar_exportacao_consolidada(modo="todos")],
            bootstyle="success",
        ).pack(fill=tk.X, pady=3)

        ttkb.Button(
            frame_botoes,
            text="1 arquivo Excel por contrato",
            command=lambda: [
                dialog.destroy(),
                self.executar_exportacao_consolidada(modo="por_contrato"),
            ],
            bootstyle="warning",
        ).pack(fill=tk.X, pady=3)


    def executar_exportacao_consolidada(self, modo):
        """
        Prepara os dados e chama as exportações consolidadas.
        modos:
        - "contrato_atual": usa só o contrato do funcionário selecionado (1 consolidado)
        - "todos": todos os contratos (1 consolidado geral + diferenças)
        - "por_contrato": gera 1 arquivo por contrato (e 1 arquivo de diferenças por contrato)
        """
        if self.dados_df is None or self.df_ponto is None:
            messagebox.showerror("Erro", "Os dados de Boletim ou Ponto ainda não foram carregados.")
            return

        lista_contratos_exportar = []

        if modo == "contrato_atual":
            funcionario_atual = (self.combo_funcionario.get() or "").strip()
            if not funcionario_atual:
                messagebox.showwarning("Aviso", "Selecione um funcionário para identificar o contrato.")
                return
            contrato_df = self.dados_df[self.dados_df["Funcionário"] == funcionario_atual]
            if contrato_df.empty or pd.isna(contrato_df["Contrato"].iloc[0]):
                messagebox.showerror("Erro", f"Não foi possível encontrar um contrato para '{funcionario_atual}'.")
                return
            lista_contratos_exportar = [contrato_df["Contrato"].iloc[0]]

            # Consolidado (único)
            exportar_totais_consolidados_excel(self, lista_contratos_exportar)
            return

        elif modo == "todos":
            lista_contratos_exportar = sorted(self.dados_df["Contrato"].dropna().unique().tolist())
            if not lista_contratos_exportar:
                messagebox.showerror("Erro", "Não foram encontrados contratos para exportar.")
                return

            # Consolidado geral (com aba por contrato) + arquivo de diferenças geral
            exportar_totais_consolidados_excel(self, lista_contratos_exportar)
            return

        elif modo == "por_contrato":
            lista_contratos_exportar = sorted(self.dados_df["Contrato"].dropna().unique().tolist())
            if not lista_contratos_exportar:
                messagebox.showerror("Erro", "Não foram encontrados contratos para exportar.")
                return

            # 1 par de arquivos por contrato (Totais e Diferenças)
            exportar_totais_consolidados_por_contrato_em_arquivos(self, lista_contratos_exportar)
            return

        # fallback defensivo
        messagebox.showwarning("Aviso", f"Modo de exportação desconhecido: {modo}")


    # ===============================================================
    # TELAS DE BOLETIM / PONTO (exibição simples)
    # ===============================================================
    '''
    def limpar_relatorio_visualizacao(self):
        """Limpa widgets de árvores/sheets da tela principal."""
        if hasattr(self, "tree") and self.tree:
            for i in self.tree.get_children():
                self.tree.delete(i)
        if hasattr(self, "tree_totais") and self.tree_totais:
            for i in self.tree_totais.get_children():
                self.tree_totais.delete(i)
        if hasattr(self, "tree_ponto") and self.tree_ponto:
            for i in self.tree_ponto.get_children():
                self.tree_ponto.delete(i)

        if hasattr(self, "sheet_boletim") and self.sheet_boletim:
            self.sheet_boletim.set_sheet_data([])

        self.df_filtrado = None
        if self.combo_funcionario:
            self.combo_funcionario.set("")


        filtro_ponto = (
            (self.df_ponto["Nome"] == funcionario)
            & (self.df_ponto["Data"] >= data_ini)
            & (self.df_ponto["Data"] <= data_fim)
        )
        df_ponto_filtrado = self.df_ponto.loc[filtro_ponto].copy()

        if not df_ponto_filtrado.empty:
            df_ponto_filtrado = df_ponto_filtrado.sort_values(by="Data")
            for _, linha in df_ponto_filtrado.iterrows():
                valores = [linha["Data"].strftime("%d/%m/%Y")]
                for col_df in self.colunas_ponto_df[1:]:
                    valor = linha.get(col_df)
                    valores.append(self.formatar_tempo(valor) if pd.notna(valor) else "")
                self.tree_ponto.insert("", "end", values=valores)
    '''
    def limpar_relatorio_visualizacao(self):
        """Limpa widgets de árvores/sheets da tela principal."""
        if hasattr(self, "tree") and self.tree:
            for i in self.tree.get_children():
                self.tree.delete(i)
        if hasattr(self, "tree_totais") and self.tree_totais:
            for i in self.tree_totais.get_children():
                self.tree_totais.delete(i)
        if hasattr(self, "tree_ponto") and self.tree_ponto:
            for i in self.tree_ponto.get_children():
                self.tree_ponto.delete(i)

        if hasattr(self, "sheet_boletim") and self.sheet_boletim:
            self.sheet_boletim.set_sheet_data([])

        self.df_filtrado = None
        if self.combo_funcionario:
            self.combo_funcionario.set("")
            
    def atualizar_interface_filtros(self):
        """Ajusta datas padrão e atualiza a combo APENAS com quem tem boletim no período."""
        if self.dados_df is None or self.dados_df.empty:
            return

        # mantém a lista completa em memória, se precisar
        self.todos_funcionarios = sorted(
            self.dados_df["Funcionário"].dropna().astype(str).unique().tolist()
        )

        # datas padrão: 1º dia até o último disponível (usando helper compatível)
        if self.cal_data_inicial and self.cal_data_final and self.data_maxima:
            self._set_date_safe(self.cal_data_final, self.data_maxima)
            self._set_date_safe(self.cal_data_inicial, self.data_maxima.replace(day=1))

        # NÃO sobrescreva a combo com todos os nomes aqui — deixe filtrado pelo período:
        try:
            self.atualizar_combo_funcionarios(silent=True)
        except Exception:
            pass

        if hasattr(self, "funcionario_boletim_label") and self.funcionario_boletim_label:
            self.funcionario_boletim_label.set(f"Funcionário: {self.combo_funcionario.get()}")
        try:
            if self.dados_df is not None and "Contrato" in self.dados_df.columns and self.combo_contrato:
                contratos = sorted(self.dados_df["Contrato"].dropna().unique().tolist())
                self.combo_contrato.configure(values=contratos)
                if contratos and not self.combo_contrato.get():
                    self.combo_contrato.current(0)
        except Exception:
            pass

    def gerar_relatorio_ponto(self):
        funcionario = (self.combo_funcionario.get() or "").strip()
        if not funcionario or self.df_ponto is None or self.df_ponto.empty:
            # limpa e sai silenciosamente
            self.df_ponto_filtrado = pd.DataFrame()
            if hasattr(self, "sheet_ponto") and self.sheet_ponto:
                self.sheet_ponto.set_sheet_data([])
            return

        try:
            di, df = ler_datas_da_ui(self.cal_data_inicial, self.cal_data_final)
        except Exception:
            return

        # builder resolve base (CPF/PIS/nome mapeado/similaridade)
        dados_lista, ident, origem, df_filtrado = build_ponto_grid(self, funcionario, di, df)
        self.df_ponto_filtrado = df_filtrado  # guarda para exportações

        # render
        if hasattr(self, "sheet_ponto") and self.sheet_ponto:
            self.sheet_ponto.set_sheet_data(dados_lista or [])
            finalizar_sheet(self.sheet_ponto, start_col=1, total_row="auto", header_bg="#ABE08E")

        # status “friendly” (sem messagebox)
        if self.label_status_ponto:
            if df_filtrado is not None and not df_filtrado.empty:
                if origem in ("CPF", "PIS"):
                    self.label_status_ponto.config(text=f"Relação usada: {funcionario} → {origem} {ident}")
                elif origem == "NOME_MAPEADO":
                    self.label_status_ponto.config(text=f"Relação usada: {funcionario} → {ident}")
                else:
                    self.label_status_ponto.config(text=f"Similaridade usada: {funcionario} → {ident}")
            else:
                self.label_status_ponto.config(text=f"Nenhum registro de ponto para '{funcionario}' no período.")

        # título da aba (se existir notebook/aba)
        if hasattr(self, "notebook") and hasattr(self, "aba_ponto_id") and ident:
            try:
                self.notebook.tab(self.aba_ponto_id, text=f"Registro Ponto - {ident}")
            except Exception:
                pass

        # banner
        try:
            reg_bol, nome_ponto_usado = self._descobrir_info_relacao(funcionario, di, df)
            self._atualizar_banners_abas(funcionario, reg_bol, nome_ponto_usado)
        except Exception:
            pass

    # ===============================================================
    # EXPORTAÇÕES SIMPLES
    # ===============================================================
    def exportar_comparacao_atual(self):
        try:
            funcionario = self.combo_funcionario.get()
            if not funcionario:
                messagebox.showwarning("Aviso", "Selecione um funcionário.")
                return
            exportar_comparacao_unificada(self)
            messagebox.showinfo("Exportação Concluída", "Arquivos PDF e Excel exportados com sucesso.")
        except Exception as e:
            messagebox.showerror("Erro na Exportação", str(e))

    def exportar_comparacao_completa(self):
        """Exporta boletim e ponto completos para XLSX (sem filtros)."""
        try:
            pasta = filedialog.askdirectory(title="Escolha uma pasta para salvar os arquivos")
            if not pasta:
                return

            if self.dados_df is None or self.df_ponto is None:
                messagebox.showwarning("Aviso", "Dados completos não carregados.")
                return

            nome_data = dt.now().strftime("%Y%m%d")
            self.dados_df.to_excel(
                os.path.join(pasta, f"boletim_completo_{nome_data}.xlsx"), index=False
            )
            self.df_ponto.to_excel(
                os.path.join(pasta, f"ponto_completo_{nome_data}.xlsx"), index=False
            )

            messagebox.showinfo(
                "Exportação Concluída", "Arquivos completos exportados com sucesso."
            )
        except Exception as e:
            messagebox.showerror("Erro na Exportação", str(e))

    def gerar_relatorio_totais_mensais(self):
        """Gera/Exporta totais mensais por funcionário."""
        if self.dados_df is None:
            messagebox.showwarning("Aviso", "Nenhum dado de boletim carregado.")
            return

        try:
            data_ini = self.cal_data_inicial.get_date()
            data_fim = self.cal_data_final.get_date()
        except Exception as e:
            messagebox.showerror("Erro de Data", f"Erro ao obter datas: {e}")
            return

        df = self.dados_df.copy()
        df = df[(df["DATA"] >= data_ini) & (df["DATA"] <= data_fim)].copy()

        if df.empty:
            self.sheet_comp_boletim.set_sheet_data([])
            self.sheet_comp_ponto.set_sheet_data([])
            self.sheet_comp_diferenca.set_sheet_data([])
            self.label_status_comparacao.config(
                text="Nenhum registro de ponto encontrado para o período."
            )
            return

        # Garante que todas as colunas numéricas existam
        for col in self.colunas_para_somar:
            if col not in df.columns:
                df[col] = 0.0

        # Agrupamento por Funcionário, Registro e Contrato
        df_group = (
            df.groupby(["Funcionário", "Registro", "Contrato"])[self.colunas_para_somar]
            .sum()
            .reset_index()
        )
        df_group = df_group.sort_values(by="Funcionário")

        pasta = filedialog.askdirectory(title="Escolha onde salvar o relatório")
        if not pasta:
            return

        nome_base = f"TotaisMensais_{data_ini.strftime('%Y%m')}"
        from interface.exportador import exportar_totais_mensais

        exportar_totais_mensais(df_group, nome_base, pasta)

    # ===============================================================
    # ATUALIZAÇÃO GERAL DAS ABAS / FORMATAÇÕES
    # ===============================================================
    def atualizar_todas_abas(self):
        """Refaz as 3 abas principais de exibição."""
        self.gerar_relatorio_tela()
        self.gerar_relatorio_ponto()
        self.gerar_relatorio_comparacao()

    def mudar_funcionario(self, direcao: int):
        """Troca funcionário atual na combobox (+1 / -1) usando a lista FILTRADA do período."""
        # Garante que a lista está coerente com as datas atuais
        try:
            self.atualizar_combo_funcionarios(silent=True)
        except Exception:
            pass

        lista = self._get_funcionarios_filtrados()
        if not lista:
            if self.status_bar:
                self.status_bar.config(text="Nenhum funcionário no período selecionado.")
            return

        atual = (self.combo_funcionario.get() or "").strip()
        # Se o selecionado atual não está mais na lista filtrada, escolhe início/fim conforme direção
        if atual not in lista:
            self.combo_funcionario.set(lista[0 if direcao >= 0 else -1])
            self.atualizar_todas_abas()
            return

        try:
            idx = lista.index(atual)
        except ValueError:
            self.combo_funcionario.set(lista[0 if direcao >= 0 else -1])
            self.atualizar_todas_abas()
            return

        novo_idx = (idx + (1 if direcao >= 0 else -1)) % len(lista)

        # Comportamento "travar" nas pontas (mesmo do seu código original)
        if novo_idx < 0:
            if self.status_bar:
                self.status_bar.config(text="Você já está no primeiro funcionário do período.")
            return
        if novo_idx >= len(lista):
            if self.status_bar:
                self.status_bar.config(text="Você já está no último funcionário do período.")
            return

        self.combo_funcionario.set(lista[novo_idx])
        try:
            self._on_funcionario_trocado()
        except Exception:
            pass

        self.atualizar_todas_abas()


    # ===============================================================
    # *** A B A  C O M P A R A Ç Ã O ***
    # Gera Boletim | Ponto | Diferença com barra preta e linhas extras na Diferença
    # ===============================================================
    def gerar_relatorio_comparacao(self):
        funcionario = (self.combo_funcionario.get() or "").strip()
        if not funcionario:
            return
        if getattr(self, "dados_df", None) is None or self.dados_df.empty:
            return
        try:
            di, df = ler_datas_da_ui(self.cal_data_inicial, self.cal_data_final)
        except Exception:
            return

        # Builders (já tratam “sem ponto”: Ponto vazio e Diferença = Boletim)
        try:
            dados_b, dados_p, dados_d, headers_vis, registro, boletins_set, sem_ponto = \
                build_comparacao_grids(self, funcionario, di, df)
        except Exception:
            # falha defensiva: zera todas
            dados_b = dados_p = dados_d = []
            headers_vis = ["Horas Normais","Horas Noturnas","Extra 50%D","Extra 100%D","Extra 50%N","Extra 100%N"]
            registro, boletins_set, sem_ponto = None, set(), True

        # Render nas 3 sheets
        for (sheet, dados) in [
            (self.sheet_comp_boletim,   dados_b),
            (self.sheet_comp_ponto,     dados_p),
            (self.sheet_comp_diferenca, dados_d),
        ]:
            if sheet is None:
                continue
            try:
                sheet.set_sheet_data(dados or [], reset_col_positions=True, reset_row_positions=True)
            except Exception:
                try:
                    sheet.set_data(dados or [])
                except Exception:
                    pass
            for m in ("refresh", "redraw"):
                try:
                    getattr(sheet, m)()
                    break
                except Exception:
                    continue

        # cabeçalhos
        try: self.sheet_comp_boletim.headers(["Data", "Boletim"] + headers_vis)
        except Exception: pass
        try: self.sheet_comp_ponto.headers(["Data"] + headers_vis)
        except Exception: pass
        try: self.sheet_comp_diferenca.headers(["Data"] + headers_vis)
        except Exception: pass

        # acabamento visual
        finalizar_conjunto_comparacao(
            self.sheet_comp_boletim, self.sheet_comp_ponto, self.sheet_comp_diferenca,
            start_col_b=2, start_col_p=1, start_col_d=1, destacar_totais=True
        )
        ocultar_colunas(self.sheet_comp_ponto, [0])      # oculta Data
        ocultar_colunas(self.sheet_comp_diferenca, [0])  # oculta Data

        # banners + status
        try:
            reg_bol, nome_ponto_usado = self._descobrir_info_relacao(funcionario, di, df)
            if registro is None:
                registro = reg_bol
            self._atualizar_banners_abas(funcionario, registro, nome_ponto_usado if not sem_ponto else None)
        except Exception:
            pass

        if self.status_bar:
            boletins_info = ", ".join(map(str, sorted(boletins_set))) if boletins_set else "-"
            tag_sem_ponto = " • Sem Ponto" if sem_ponto else ""
            self.status_bar.config(
                text=f"Comparação atualizada • Registro: {registro or '—'} • Boletins: {boletins_info}{tag_sem_ponto}"
            )


    # ===============================================================
    # RELAÇÃO DE NOMES — TELA DE GERENCIAMENTO
    # ===============================================================
    def gerenciar_relacao_nomes(self):
        from difflib import SequenceMatcher
        from tkinter import Listbox, ttk

        import ttkbootstrap as ttkb

        try:
            from unidecode import unidecode
        except Exception:

            def unidecode(s):
                return s

        def _sort_key_nome(s: str) -> str:
            return unidecode(str(s)).casefold().strip()

        if self.dados_df is None or self.df_ponto is None:
            messagebox.showwarning(
                "Dados incompletos", "Carregue os dados do Boletim e do Ponto primeiro."
            )
            return

        # --- PARTE 1: ATUALIZAÇÃO DA BASE DE DADOS (Lógica existente) ---
        df_relacao_existente = pd.DataFrame(
            columns=["Nome_Boletim", "CPF_Ponto", "PIS_Ponto", "Nome_Ponto_Mapeado"]
        )
        if os.path.exists(self.arquivo_relacao_nomes):
            df_relacao_existente = pd.read_parquet(self.arquivo_relacao_nomes)
        nomes_boletim_todos = self.dados_df["Funcionário"].dropna().unique()
        nomes_ponto_df = (
            self.df_ponto[["Nome", "CPF", "PIS"]]
            .dropna(subset=["Nome"])
            .drop_duplicates(subset=["Nome"])
        )
        novas_relacoes = []
        nomes_ja_mapeados = (
            df_relacao_existente["Nome_Boletim"].tolist() if not df_relacao_existente.empty else []
        )
        for nome_b in nomes_boletim_todos:
            if nome_b in nomes_ja_mapeados:
                continue
            correspondencias = get_close_matches(
                nome_b.upper(), nomes_ponto_df["Nome"].str.upper(), n=1, cutoff=0.75
            )  #
            nova_linha = {
                "Nome_Boletim": nome_b,
                "CPF_Ponto": None,
                "PIS_Ponto": None,
                "Nome_Ponto_Mapeado": None,
            }
            if correspondencias:
                dados_ponto = nomes_ponto_df[
                    nomes_ponto_df["Nome"].str.upper() == correspondencias[0]
                ].iloc[0]
                nova_linha.update(
                    {
                        "CPF_Ponto": dados_ponto["CPF"],
                        "PIS_Ponto": dados_ponto["PIS"],
                        "Nome_Ponto_Mapeado": dados_ponto["Nome"],
                    }
                )
            novas_relacoes.append(nova_linha)
        if novas_relacoes:
            df_novas_relacoes = pd.DataFrame(novas_relacoes)
            df_relacao_existente = pd.concat(
                [df_relacao_existente, df_novas_relacoes], ignore_index=True
            )
            os.makedirs("data", exist_ok=True)
            df_relacao_existente.to_parquet(self.arquivo_relacao_nomes, index=False)

        # --- PARTE 2: CRIAÇÃO DA NOVA JANELA ---
        janela_relacao = tk.Toplevel(self.janela)
        janela_relacao.title("Gerenciar Relação de Nomes")
        janela_relacao.geometry("800x600")
        janela_relacao.transient(self.janela)
        janela_relacao.grab_set()

        notebook = ttk.Notebook(janela_relacao)
        notebook.pack(pady=10, padx=10, fill="both", expand=True)

        # --- PARTE 3: LÓGICA DE DADOS E FUNÇÕES AUXILIARES ---
        master_list = []

        def recarregar_e_filtrar_dados():
            nonlocal master_list
            df = pd.read_parquet(self.arquivo_relacao_nomes)

            master_list = []
            for _, row in df.iterrows():
                nome_b = row.get("Nome_Boletim")
                nome_p = row.get("Nome_Ponto_Mapeado")
                score = (
                    SequenceMatcher(None, str(nome_b).upper(), str(nome_p).upper()).ratio()
                    if pd.notna(nome_p)
                    else 0
                )
                master_list.append({"boletim": nome_b, "ponto": nome_p, "score": score})

            # 🔽 Ordena a lista mestre por Nome do Boletim (sem acento/case)
            master_list.sort(key=lambda x: _sort_key_nome(x["boletim"]))

            sem_relacao = [i for i in master_list if pd.isna(i["ponto"])]
            incertezas = [i for i in master_list if pd.notna(i["ponto"]) and i["score"] < 1.0]

            return {
                "Todos": master_list,
                "Sem Relação": sorted(sem_relacao, key=lambda x: _sort_key_nome(x["boletim"])),
                "Incertezas (<100%)": sorted(
                    incertezas, key=lambda x: _sort_key_nome(x["boletim"])
                ),
            }

        # --- PARTE 4: FUNÇÃO PARA CRIAR O CONTEÚDO DE CADA ABA ---
        def criar_conteudo_aba(parent_frame, nome_aba, data_list):
            # PAINEL ESQUERDO: LISTA DE NOMES DO BOLETIM
            frame_lista = ttk.Frame(parent_frame)
            frame_lista.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 10))
            ttk.Label(frame_lista, text=f"Nomes no Boletim ({nome_aba}):").pack(anchor="w")
            from tkinter import Listbox

            listbox_boletim = Listbox(frame_lista, height=25, exportselection=False, width=40)
            listbox_boletim.pack(fill=tk.Y)
            for item in data_list:
                listbox_boletim.insert(tk.END, item["boletim"])

            # PAINEL DIREITO: EDITOR
            frame_editor = ttk.Frame(parent_frame)
            frame_editor.pack(side=tk.LEFT, fill="both", expand=True)

            ttk.Label(frame_editor, text="Nome do Boletim Selecionado:").pack(
                anchor="w", pady=(5, 0)
            )
            nome_boletim_selecionado_var = tk.StringVar()
            ttkb.Label(
                frame_editor,
                textvariable=nome_boletim_selecionado_var,
                font=("Segoe UI", 14, "bold"),
                foreground="blue",
            ).pack(anchor="w", pady=(0, 10))

            # ### NOVOS WIDGETS DE FILTRO ###
            ttk.Label(frame_editor, text="Filtrar correspondência:").pack(anchor="w")
            entry_filtro = ttkb.Entry(frame_editor, font=("Segoe UI", 10))
            entry_filtro.pack(anchor="w", fill="x", pady=(0, 5))

            ttk.Label(frame_editor, text="Correspondência Sugerida/Salva no Ponto:").pack(
                anchor="w"
            )
            nomes_ponto_df_local = (
                self.df_ponto[["Nome", "CPF", "PIS"]]
                .dropna(subset=["Nome"])
                .drop_duplicates(subset=["Nome"])
            )
            nomes_ponto_lista = sorted(
                nomes_ponto_df_local["Nome"].unique().tolist(), key=lambda s: s.casefold().strip()
            )
            combobox_ponto = ttkb.Combobox(frame_editor, width=50, font=("Segoe UI", 10))
            combobox_ponto.pack(anchor="w", fill="x", pady=(0, 10))

            ttk.Label(frame_editor, text="Índice de Certeza da Sugestão:").pack(anchor="w")
            certeza_var = tk.StringVar()
            label_certeza = ttkb.Label(
                frame_editor, textvariable=certeza_var, font=("Segoe UI", 12, "bold")
            )
            label_certeza.pack(anchor="w", pady=(0, 20))

            btn_salvar = ttkb.Button(frame_editor, text="Salvar Relação", bootstyle="success")
            btn_limpar = ttkb.Button(
                frame_editor, text="Limpar Relação", bootstyle="danger-outline"
            )
            btn_salvar.pack(side=tk.LEFT, pady=10)
            btn_limpar.pack(side=tk.LEFT, padx=10, pady=10)

            def filtrar_correspondencias(_event):
                filtro_texto = entry_filtro.get().upper()
                selection_indices = listbox_boletim.curselection()
                if not selection_indices:
                    combobox_ponto["values"] = ["--- NENHUMA CORRESPONDÊNCIA ---"]
                    return

                nome_b_selecionado = data_list[selection_indices[0]]["boletim"]

                # Mantém sugestões principais no topo
                sugestoes_principais = get_close_matches(
                    nome_b_selecionado.upper(), nomes_ponto_lista, n=5, cutoff=0.6
                )
                sugestoes_filtradas = [s for s in sugestoes_principais if filtro_texto in s.upper()]
                outros_nomes = [
                    n
                    for n in nomes_ponto_lista
                    if n not in sugestoes_principais and filtro_texto in n.upper()
                ]

                nova_lista_valores = (
                    ["--- NENHUMA CORRESPONDÊNCIA ---"]
                    + sugestoes_filtradas
                    + ["-" * 20]
                    + outros_nomes
                )
                combobox_ponto["values"] = nova_lista_valores

            entry_filtro.bind("<KeyRelease>", filtrar_correspondencias)

            def on_boletim_name_selected(_event):
                selection_indices = listbox_boletim.curselection()
                if not selection_indices:
                    return
                selected_item = data_list[selection_indices[0]]
                nome_b, nome_p_atual, score_atual = (
                    selected_item["boletim"],
                    selected_item["ponto"],
                    selected_item["score"],
                )

                nome_boletim_selecionado_var.set(nome_b)
                entry_filtro.delete(0, tk.END)
                filtrar_correspondencias(None)

                combobox_ponto.set(
                    nome_p_atual if pd.notna(nome_p_atual) else "--- NENHUMA CORRESPONDÊNCIA ---"
                )

                certeza_var.set(f"{score_atual:.1%}" if score_atual > 0 else "---")
                if score_atual >= 0.9:
                    label_certeza.config(foreground="green")
                elif score_atual >= 0.75:
                    label_certeza.config(foreground="orange")
                else:
                    label_certeza.config(foreground="red")

            def _salvar_relacao_nomes(df_editado):
                # df_editado = DataFrame que saiu da sua tela de edição
                import os

                tmp = self.arquivo_relacao_nomes + ".tmp"
                df_ok = self._sanear_relacao_df(df_editado)
                df_ok.to_parquet(tmp, index=False)
                os.replace(tmp, self.arquivo_relacao_nomes)  # gravação atômica

                # 🔄 recarrega em memória imediatamente
                self._carregar_relacao_nomes()

                # ✅ opcional: atualizar telas que dependem da relação
                try:
                    self.atualizar_todas_abas()  # se existir
                except Exception:
                    pass

            def salvar(limpar=False):
                selection_indices = listbox_boletim.curselection()
                if not selection_indices:
                    messagebox.showwarning("Aviso", "Nenhum nome selecionado na lista da esquerda.")
                    return

                nome_b_selecionado = data_list[selection_indices[0]]["boletim"]
                nome_p_antigo = data_list[selection_indices[0]]["ponto"]
                nome_p_novo = combobox_ponto.get() if not limpar else "NENHUM"

                if nome_p_novo in ["--- NENHUMA CORRESPONDÊNCIA ---", "NENHUM", "-" * 20]:
                    nome_p_novo_display = "NENHUMA"
                else:
                    nome_p_novo_display = nome_p_novo

                msg = (
                    f"Você está alterando a relação para:\n\n"
                    f"Funcionário (Boletim): {nome_b_selecionado}\n\n"
                    f"Correspondência ANTERIOR: {nome_p_antigo if pd.notna(nome_p_antigo) else 'NENHUMA'}\n"
                    f"Correspondência NOVA: {nome_p_novo_display}\n\n"
                    f"Deseja confirmar a alteração?"
                )
                if not messagebox.askyesno("Confirmar Alteração", msg):
                    return

                df = pd.read_parquet(self.arquivo_relacao_nomes)
                idx_linha = df.index[df["Nome_Boletim"] == nome_b_selecionado].tolist()[0]
                if limpar or nome_p_novo in ["--- NENHUMA CORRESPONDÊNCIA ---", "-" * 20]:
                    df.loc[idx_linha, ["CPF_Ponto", "PIS_Ponto", "Nome_Ponto_Mapeado"]] = (
                        None,
                        None,
                        None,
                    )
                else:
                    dados_ponto = nomes_ponto_df[nomes_ponto_df["Nome"] == nome_p_novo].iloc[0]
                    df.loc[idx_linha, ["CPF_Ponto", "PIS_Ponto", "Nome_Ponto_Mapeado"]] = (
                        dados_ponto["CPF"],
                        dados_ponto["PIS"],
                        dados_ponto["Nome"],
                    )
                # ✔️ gravação + recarga GLOBAL via helper (caminho único de gravação)
                _salvar_relacao_nomes(df)

                # 🧭 tentar preservar seleção atual
                sel_indices = listbox_boletim.curselection()
                sel_nome_b = None
                if sel_indices:
                    sel_nome_b = data_list[sel_indices[0]]["boletim"]

                # 🔁 recarrega dados e atualiza a aba atual
                listas = recarregar_e_filtrar_dados()
                data_list[:] = listas[nome_aba]  # atualiza lista in-place

                # repopula listbox
                listbox_boletim.delete(0, tk.END)
                for item in data_list:
                    listbox_boletim.insert(tk.END, item["boletim"])

                # restaura seleção, se possível
                restaurado = False
                if sel_nome_b:
                    try:
                        idx = next(
                            i
                            for i, it in enumerate(data_list)
                            if str(it["boletim"]) == str(sel_nome_b)
                        )
                        listbox_boletim.selection_clear(0, tk.END)
                        listbox_boletim.selection_set(idx)
                        listbox_boletim.see(idx)
                        on_boletim_name_selected(None)
                        restaurado = True
                    except StopIteration:
                        pass

                if not restaurado and data_list:
                    cand = [it["boletim"] for it in data_list]
                    alvo = sel_nome_b or cand[0]
                    similar = get_close_matches(
                        str(alvo).upper(), [str(c).upper() for c in cand], n=1, cutoff=0.6
                    )
                    if similar:
                        alvo_sim = next(c for c in cand if str(c).upper() == similar[0])
                        idx_sim = cand.index(alvo_sim)
                        listbox_boletim.selection_clear(0, tk.END)
                        listbox_boletim.selection_set(idx_sim)
                        listbox_boletim.see(idx_sim)
                        on_boletim_name_selected(None)
                    else:
                        listbox_boletim.selection_set(0)
                        listbox_boletim.see(0)
                        on_boletim_name_selected(None)

                # atualiza fonte do combobox também
                nomes_ponto_df_local = (
                    self.df_ponto[["Nome", "CPF", "PIS"]]
                    .dropna(subset=["Nome"])
                    .drop_duplicates(subset=["Nome"])
                )
                nomes_ponto_lista[:] = sorted(
                    nomes_ponto_df_local["Nome"].unique().tolist(),
                    key=lambda s: s.casefold().strip(),
                )
                filtrar_correspondencias(None)

                messagebox.showinfo("Sucesso", "Relação salva e tela atualizada.")

            listbox_boletim.bind("<<ListboxSelect>>", on_boletim_name_selected)
            btn_salvar.config(command=lambda: salvar(limpar=False))
            btn_limpar.config(command=lambda: salvar(limpar=True))

            if data_list:
                listbox_boletim.selection_set(0)
                on_boletim_name_selected(None)

        # --- PARTE 5: POPULAR AS ABAS ---
        listas_filtradas = recarregar_e_filtrar_dados()
        for nome_aba, data in listas_filtradas.items():
            frame = ttk.Frame(notebook)
            notebook.add(frame, text=f"{nome_aba} ({len(data)})")
            criar_conteudo_aba(frame, nome_aba, data)

    def gerar_relatorio_tela(self):
        # 1) Guardas
        if self.dados_df is None:
            return
        funcionario = (self.combo_funcionario.get() or "").strip()
        if not funcionario:
            return
        try:
            di, df = ler_datas_da_ui(self.cal_data_inicial, self.cal_data_final)
        except Exception:
            return

        # 2) Dados prontos (builder)
        dados_lista, df_filtrado = build_boletim_grid(self, funcionario, di, df)
        self.df_filtrado = df_filtrado  # mantém para exportações

        # 3) Render
        if hasattr(self, "sheet_boletim") and self.sheet_boletim:
            self.sheet_boletim.set_sheet_data(dados_lista or [])
            finalizar_sheet(self.sheet_boletim, start_col=1, total_row="auto", header_bg="#BAD8F3")

        # 4) Banner (registro do boletim + nome do ponto)
        try:
            reg_bol, nome_ponto_usado = self._descobrir_info_relacao(funcionario, di, df)
            self._atualizar_banners_abas(funcionario, reg_bol, nome_ponto_usado)
        except Exception:
            pass


    # ===============================================================
    # DIAGNÓSTICO GERAL
    # ===============================================================
    def mostrar_dialogo_exportacao_ponto(self):
        #Dialog para exportar ponto (filtrado/completo) em Excel/PDF.
        if self.df_ponto is None or self.df_ponto.empty:
            messagebox.showwarning("Aviso", "Nenhum dado de ponto carregado para exportar.")
            return

        dialog = ttkb.Toplevel(self.janela)
        dialog.title("Exportar Registro de Ponto")
        dialog.geometry("400x140")
        dialog.resizable(False, False)
        dialog.transient(self.janela)
        dialog.grab_set()

        label = ttkb.Label(
            dialog, text="Escolha o formato e o tipo de relatório:", font=("Segoe UI", 10)
        )
        label.pack(pady=10)

        frame_botoes = ttkb.Frame(dialog)
        frame_botoes.pack(pady=5)

        ttkb.Button(
            frame_botoes,
            text="Excel (Filtrado)",
            command=lambda: [exportar_ponto_para_excel(self, completo=False), dialog.destroy()],
            bootstyle="success",
        ).pack(side=tk.LEFT, padx=5)
        ttkb.Button(
            frame_botoes,
            text="PDF (Filtrado)",
            command=lambda: [exportar_ponto_para_pdf(self, completo=False), dialog.destroy()],
            bootstyle="danger",
        ).pack(side=tk.LEFT, padx=5)
        ttkb.Button(
            frame_botoes,
            text="Excel (Completo)",
            command=lambda: [exportar_ponto_para_excel(self, completo=True), dialog.destroy()],
            bootstyle="success-outline",
        ).pack(side=tk.LEFT, padx=5)
        ttkb.Button(
            frame_botoes,
            text="PDF (Completo)",
            command=lambda: [exportar_ponto_para_pdf(self, completo=True), dialog.destroy()],
            bootstyle="danger-outline",
        ).pack(side=tk.LEFT, padx=5)

        ttkb.Button(dialog, text="Cancelar", command=dialog.destroy, bootstyle="secondary").pack(
            pady=(5, 0)
        )

    def executar_diagnostico(self):
        #Exibe um relatório rápido sobre estado de dados e widgets.
        mensagens = []
        ok = lambda x: f"✅ {x}"
        erro = lambda x: f"❌ {x}"

        # Verificações de DataFrames
        if self.dados_df is None or self.dados_df.empty:
            mensagens.append(erro("dados_df (Boletim) está vazio ou não carregado."))
        else:
            mensagens.append(ok("dados_df carregado."))

        if self.df_ponto is None or self.df_ponto.empty:
            mensagens.append(erro("df_ponto está vazio ou não carregado."))
        else:
            mensagens.append(ok("df_ponto carregado."))

        if self.df_filtrado is None or self.df_filtrado.empty:
            mensagens.append(erro("df_filtrado (filtrado por funcionário) não está disponível."))
        else:
            mensagens.append(ok("df_filtrado disponível."))

        # Verificações de widgets importantes
        componentes = {
            "sheet_comp_boletim": self.sheet_comp_boletim,
            "sheet_comp_ponto": self.sheet_comp_ponto,
            "sheet_comp_diferenca": self.sheet_comp_diferenca,
            "sheet_boletim": self.sheet_boletim,
            "sheet_ponto": self.sheet_ponto,
            "combo_funcionario": self.combo_funcionario,
        }
        for nome, comp in componentes.items():
            if comp is None:
                mensagens.append(erro(f"{nome} não foi inicializado."))
            else:
                mensagens.append(ok(f"{nome} OK."))

        # — checagens leves para avisar sobre configurações esperadas na UI —
        try:
            _ = self.colunas_para_somar
            mensagens.append(ok("colunas_para_somar OK."))
        except Exception:
            mensagens.append(erro("colunas_para_somar não definido."))

        try:
            _ = self.colunas_ponto
            mensagens.append(ok("colunas_ponto OK."))
        except Exception:
            mensagens.append(erro("colunas_ponto não definido."))

        # headers/colunas do boletim usados na tela
        if hasattr(self, "colunas_boletim_dados") and self.colunas_boletim_dados:
            mensagens.append(ok("colunas_boletim_dados OK."))
        else:
            mensagens.append(erro("colunas_boletim_dados não definido (usando fallback padrão)."))

        # formato de hora (decimal/HH:MM)
        try:
            _ = self.modo_exibicao_decimal
            mensagens.append(ok(f"modo_exibicao_decimal = {self.modo_exibicao_decimal}"))
        except Exception:
            mensagens.append(erro("modo_exibicao_decimal não definido."))

        # teste opcional do JSON de tarifas (não falha o diagnóstico)
        try:
            from pathlib import Path as _Path

            _json_path = _Path(__file__).resolve().parent.parent / "data" / "tarifas_copel.json"
            if _json_path.exists():
                mensagens.append(ok(f"tarifas_copel.json encontrado em '{_json_path.name}'."))
            else:
                mensagens.append(
                    "⚠️ Arquivo 'tarifas_copel.json' não encontrado (linhas de valores em R$ ficarão ocultas)."
                )
        except Exception:
            pass

        # Exibe o resultado final
        texto = "\n".join(mensagens)
        messagebox.showinfo("Diagnóstico do Sistema", texto)


    def abrir_formulario_geral(self):
        # Usa o utils/relatorios diretamente para montar e abrir a janela
        from utils.processador import (
            ARQUIVO_PADRAO,
            carregar_boletim_parquet,
            relatorio_sa_por_registro_periodo,
            RelatorioSADialog,
        )
        
        parent = getattr(self, "janela", None) or getattr(self, "window", None) or getattr(self, "root", None)

        # tenta pegar as datas atuais da UI (aceita dd/mm/aaaa)
        try:
            data_ini = self.cal_data_inicial.entry.get()
            data_fim = self.cal_data_final.entry.get()
        except Exception:
            data_ini = data_fim = None

        df_bol = carregar_boletim_parquet()  # data/RegistroBoletim.parquet
        df_view = relatorio_sa_por_registro_periodo(df_bol, data_ini, data_fim)
        RelatorioSADialog(parent, df_view, data_ini=data_ini, data_fim=data_fim)

    
    
    # --- Helpers de “faixa informativa” para todas as abas ---

    def _format_info_bol_pto(self, nome_boletim: str, registro: str | None, nome_ponto: str | None) -> str:
        reg = f", {registro}" if registro else ""
        ponto = nome_ponto or "—"
        # agora em duas linhas:
        return f"Boletim: {nome_boletim}{reg}\nPonto: {ponto}"

    def _atualizar_banners_abas(self, nome_boletim: str, registro: str | None, nome_ponto: str | None) -> None:
        txt = self._format_info_bol_pto(nome_boletim, registro, nome_ponto)
        labels = [
            getattr(self, "label_status_boletim", None),
            getattr(self, "label_status_ponto", None),
            getattr(self, "label_status_comparacao", None),
        ]
        updated = False
        for lbl in labels:
            try:
                if lbl:
                    # aceita quebra de linha e alinha à esquerda
                    lbl.configure(text=txt, justify="left", anchor="w")
                    # se quiser limitar quebra automática, defina wraplength (em pixels), ex.: 900
                    # lbl.configure(wraplength=900)
                    updated = True
            except Exception:
                pass
        if not updated and getattr(self, "status_bar", None):
            # status bar costuma ser 1 linha; use versão “achatada”
            self.status_bar.config(text=txt.replace("\n", " | "))


    def _descobrir_info_relacao(self, funcionario: str, data_ini, data_fim):
        """
        Retorna (registro_do_boletim, nome_ponto_usado) para o período.
        - registro: busca na base de boletim filtrada pelo período.
        - nome_ponto: usa resolver_base_ponto (CPF→PIS→Nome mapeado→similaridade).
        """
 
        from utils.dataframe_utils import resolver_base_ponto

        # 1) Registro do Boletim no período
        registro = None
        try:
            mask = (
                (self.dados_df["Funcionário"] == funcionario)
                & (self.dados_df["DATA"].dt.date >= data_ini)
                & (self.dados_df["DATA"].dt.date <= data_fim)
            )
            regs = (
                self.dados_df.loc[mask, "Registro"]
                .dropna()
                .astype(str)
                .unique()
                .tolist()
            )
            registro = regs[0] if regs else None
        except Exception:
            pass

        # 2) Nome do Ponto realmente usado (pela relação)
        nome_ponto = None
        try:
            df_sel, origem, ident = resolver_base_ponto(
                self.df_ponto, self.df_relacao, funcionario, data_ini, data_fim, corte_similaridade=0.75
            )
            if not df_sel.empty:
                if origem == "NOME_MAPEADO" and ident:
                    nome_ponto = str(ident)
                elif "Nome" in df_sel.columns and df_sel["Nome"].notna().any():
                    # pega o mais frequente (mode) ou o primeiro não-nulo
                    nomes = df_sel["Nome"].dropna().astype(str)
                    if not nomes.empty:
                        try:
                            nome_ponto = nomes.mode().iat[0]
                        except Exception:
                            nome_ponto = nomes.iloc[0]
        except Exception:
            pass

        return registro, nome_ponto
    
    # --- DateEntry helpers (compat) ---
    def _set_date_safe(self, widget, date_obj):
        """Tenta widget.set_date(); se não existir, escreve no .entry ou no próprio widget."""
        from datetime import date, datetime
        try:
            # tkcalendar geralmente tem
            widget.set_date(date_obj)
            return True
        except Exception:
            pass

        # formata dd/mm/aaaa
        try:
            if isinstance(date_obj, (datetime, date)):
                s = date_obj.strftime("%d/%m/%Y")
            else:
                s = str(date_obj)
        except Exception:
            s = str(date_obj)

        # tenta via .entry (tkcalendar.DateEntry possui)
        try:
            entry = getattr(widget, "entry", None) or widget  # fallback: o próprio widget
            entry.delete(0, tk.END)
            entry.insert(0, s)
            return True
        except Exception:
            return False

    def _on_toggle_filtrar_contrato(self):
        """Liga/desliga o filtro por contrato e atualiza a combo de funcionários."""
        try:
            filtrar = bool(self.var_filtrar_contrato.get())
        except Exception:
            filtrar = False
        try:
            if self.combo_contrato:
                self.combo_contrato.configure(state=("readonly" if filtrar else "disabled"))
        except Exception:
            pass
        # ao mudar o estado, refaz a lista de funcionários
        self.atualizar_combo_funcionarios()

    def _on_funcionario_trocado(self):
        """Atualiza a combobox de contrato para refletir o(s) contrato(s) do funcionário selecionado
        no período atual. Se o checkbox 'Filtrar contrato' estiver ligado, também refiltra a lista
        de funcionários pelo contrato escolhido."""
        if self.dados_df is None or self.dados_df.empty or not self.combo_contrato:
            return

        try:
            func = (self.combo_funcionario.get() or "").strip()
            if not func:
                return

            di, df = ler_datas_da_ui(self.cal_data_inicial, self.cal_data_final)

            dfp = self.dados_df.copy()
            if not pd.api.types.is_datetime64_any_dtype(dfp["DATA"]):
                dfp["DATA"] = pd.to_datetime(dfp["DATA"], errors="coerce")

            # contratos do funcionário no período
            mask = (
                (dfp["Funcionário"] == func)
                & (dfp["DATA"].dt.date >= di)
                & (dfp["DATA"].dt.date <= df)
                & (dfp["Contrato"].notna())
            )
            contratos = sorted(dfp.loc[mask, "Contrato"].astype(str).unique().tolist())

            # popula a combo de contrato (apenas valores do funcionário)
            self.combo_contrato.configure(values=contratos)
            if contratos:
                # mantém seleção se existir; senão escolhe a primeira
                atual = (self.combo_contrato.get() or "").strip()
                if atual not in contratos:
                    self.combo_contrato.set(contratos[0])
            else:
                self.combo_contrato.set("")

            # se o filtro estiver ativo, refaz a lista de funcionários conforme o contrato
            if self.var_filtrar_contrato and bool(self.var_filtrar_contrato.get()):
                self.atualizar_combo_funcionarios(silent=True)
        except Exception:
            pass


