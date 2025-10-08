# gui_relatorio.py
# ------------------------------------------------------------
# Interface gráfica avançada com carregamento automático,
# definição de período de data inteligente e filtros interativos.
# VERSÃO CORRIGIDA do filtro de data.
# ------------------------------------------------------------

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import pandas as pd
from typing import List, Dict, Any
import datetime

# --- Importações para Widgets e Exportação ---
from tkcalendar import DateEntry
from reportlab.lib.pagesizes import letter, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors
from reportlab.lib.units import inch

# Tenta importar a lógica do script original.
try:
    from relatorio_registro_console import (
        _carregar_parquet,
        _validar_colunas,
        _forcar_tipos,
        _agrupar_por_registro,
        _montar_linhas_com_subtotais,
        ARQUIVO_PADRAO,
    )
except ImportError:
    messagebox.showerror(
        "Erro de Importação",
        "Não foi possível encontrar o script 'relatorio_registro_console.py'.\n"
        "Certifique-se de que ele esteja na mesma pasta que este script.",
    )
    exit()


def _formatar_para_string(linhas: List[Dict[str, Any]]) -> str:
    # (Esta função permanece a mesma)
    if not linhas: return "(Nenhum resultado encontrado com os filtros aplicados)"
    headers = ["Funcionário", "Registro", "BOLETIM", "Contrato", "S.A."]; col_w = {h: len(h) for h in headers}
    for row in linhas:
        for h in headers:
            txt = row[h]
            if h == "S.A.":
                try: width_val = len(f"{float(txt):.3f}")
                except (ValueError, TypeError): width_val = len(str(txt))
                col_w[h] = max(col_w[h], width_val)
            else: col_w[h] = max(col_w[h], len(str(txt)))
    def fmt_cell(h: str, v: Any) -> str:
        if h == "S.A.":
            try: return f"{float(v):>{col_w[h]}.3f}"
            except (ValueError, TypeError): return f"{str(v):>{col_w[h]}}"
        else: return f"{str(v):<{col_w[h]}}"
    sep = " | "; line_sep_str = "-+-".join("-" * col_w[h] for h in headers)
    header_line = sep.join(f"{h:<{col_w[h]}}" for h in headers)
    output_str = f"{header_line}\n{line_sep_str}\n"
    for row in linhas: output_str += sep.join(fmt_cell(h, row[h]) for h in headers) + "\n"
    return output_str


class RelatorioApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Gerador de Relatório de Registro")
        self.root.geometry("1100x750")

        self.caminho_arquivo = tk.StringVar(value=ARQUIVO_PADRAO)
        self.dados_relatorio = None
        self.df_original = None
        self.lista_funcionarios_completa = []

        # --- Frames de UI e Widgets (sem alterações) ---
        main_frame = tk.Frame(root, padx=10, pady=10); main_frame.pack(fill=tk.BOTH, expand=True)
        frame_arquivo = tk.Frame(main_frame); frame_arquivo.pack(fill=tk.X, pady=(0, 10))
        frame_filtros = tk.LabelFrame(main_frame, text="Filtros", padx=10, pady=10); frame_filtros.pack(fill=tk.X, pady=(0, 10))
        frame_botoes = tk.Frame(main_frame); frame_botoes.pack(fill=tk.X)
        self.text_resultado = scrolledtext.ScrolledText(main_frame, wrap=tk.NONE, font=("Courier New", 10)); self.text_resultado.pack(padx=0, pady=(10, 0), expand=True, fill=tk.BOTH)
        tk.Label(frame_arquivo, text="Arquivo Parquet:").pack(side=tk.LEFT, padx=(0, 5))
        self.entry_arquivo = tk.Entry(frame_arquivo, textvariable=self.caminho_arquivo); self.entry_arquivo.pack(side=tk.LEFT, expand=True, fill=tk.X)
        self.btn_selecionar = tk.Button(frame_arquivo, text="Selecionar...", command=self.selecionar_arquivo); self.btn_selecionar.pack(side=tk.LEFT, padx=(5, 0))
        tk.Label(frame_filtros, text="Data Inicial:").grid(row=0, column=0, sticky="w", padx=5, pady=2)
        self.filtro_data_ini = DateEntry(frame_filtros, date_pattern='dd/mm/yyyy', width=12); self.filtro_data_ini.grid(row=0, column=1, sticky="w", pady=2)
        tk.Label(frame_filtros, text="Data Final:").grid(row=0, column=2, sticky="w", padx=(15, 5), pady=2)
        self.filtro_data_fim = DateEntry(frame_filtros, date_pattern='dd/mm/yyyy', width=12); self.filtro_data_fim.grid(row=0, column=3, sticky="w", pady=2)
        tk.Label(frame_filtros, text="Contrato:").grid(row=1, column=0, sticky="w", padx=5, pady=2)
        self.filtro_contrato = tk.StringVar(); self.combo_contrato = ttk.Combobox(frame_filtros, textvariable=self.filtro_contrato, width=30, state="readonly"); self.combo_contrato.grid(row=1, column=1, columnspan=3, sticky="w", pady=2)
        tk.Label(frame_filtros, text="Funcionário:").grid(row=2, column=0, sticky="w", padx=5, pady=2)
        self.filtro_funcionario = tk.StringVar(); self.combo_funcionario = ttk.Combobox(frame_filtros, textvariable=self.filtro_funcionario, width=50); self.combo_funcionario.grid(row=2, column=1, columnspan=3, sticky="w", pady=2)
        self.combo_funcionario.bind('<KeyRelease>', self.atualizar_filtro_funcionario)
        self.btn_gerar = tk.Button(frame_botoes, text="Gerar Relatório", command=self.gerar_relatorio); self.btn_gerar.pack(side=tk.LEFT, padx=(0, 5))
        self.btn_limpar = tk.Button(frame_botoes, text="Limpar Filtros", command=self.limpar_filtros); self.btn_limpar.pack(side=tk.LEFT, padx=5)
        self.btn_export_excel = tk.Button(frame_botoes, text="Exportar para Excel", state=tk.DISABLED, command=self.exportar_para_excel); self.btn_export_excel.pack(side=tk.LEFT, padx=5)
        self.btn_export_pdf = tk.Button(frame_botoes, text="Exportar para PDF", state=tk.DISABLED, command=self.exportar_para_pdf); self.btn_export_pdf.pack(side=tk.LEFT, padx=5)

        self.root.after(100, self._carregar_arquivo_inicial)

    def _processar_dataframe_carregado(self):
        """
        Processa o DataFrame, define o período de datas corretamente e popula os filtros.
        """
        if self.df_original is None: return
        
        coluna_data = 'DATA'
        if coluna_data in self.df_original.columns:
            # Converte a coluna para datetime
            self.df_original[coluna_data] = pd.to_datetime(self.df_original[coluna_data], errors='coerce')
            
            # CORREÇÃO: Lógica para definir o período de datas
            max_date = self.df_original[coluna_data].max()
            if pd.notna(max_date):
                start_of_month = max_date.replace(day=1)
                # A data final é a própria data máxima encontrada
                end_date = max_date 
                
                self.filtro_data_ini.set_date(start_of_month)
                self.filtro_data_fim.set_date(end_date)
        
        self._popular_filtros()
        self.text_resultado.delete("1.0", tk.END)
        self.text_resultado.insert(tk.END, "Arquivo carregado. Clique em 'Gerar Relatório' para ver os dados do período sugerido.")

    def _carregar_arquivo_inicial(self):
        caminho = self.caminho_arquivo.get()
        if not caminho: return
        try:
            self.text_resultado.insert(tk.END, f"Carregando arquivo padrão: {caminho}...")
            self.root.update_idletasks()
            self.df_original = _carregar_parquet(caminho)
            
            # CORREÇÃO: Chama _forcar_tipos passando a coluna extra a ser mantida
            self.df_original = _forcar_tipos(self.df_original, cols_extras=['DATA'])

            # Valida as colunas essenciais APÓS a preparação
            _validar_colunas(self.df_original)

            self._processar_dataframe_carregado()
        except FileNotFoundError:
            self.text_resultado.delete("1.0", tk.END)
            self.text_resultado.insert(tk.END, "Arquivo padrão não encontrado. Por favor, selecione um arquivo manualmente.")
        except Exception as e:
            messagebox.showerror("Erro ao Carregar", f"Não foi possível carregar ou processar o arquivo inicial:\n{e}")
            self.text_resultado.delete("1.0", tk.END)
            self.df_original = None
    
    def selecionar_arquivo(self):
        caminho = filedialog.askopenfilename(title="Selecione o arquivo Parquet", filetypes=[("Parquet files", "*.parquet")])
        if not caminho: return
        self.caminho_arquivo.set(caminho)
        self.text_resultado.delete("1.0", tk.END)
        self._limpar_estado_relatorio()
        self.df_original = None
        try:
            self.text_resultado.insert(tk.END, "Carregando arquivo e populando filtros...")
            self.root.update_idletasks()
            self.df_original = _carregar_parquet(caminho)

            # CORREÇÃO: Chama _forcar_tipos passando a coluna extra a ser mantida
            self.df_original = _forcar_tipos(self.df_original, cols_extras=['DATA'])

            # Valida as colunas essenciais APÓS a preparação
            _validar_colunas(self.df_original)

            self._processar_dataframe_carregado()
        except Exception as e:
            messagebox.showerror("Erro ao Carregar", f"Não foi possível carregar ou processar o arquivo:\n{e}")
            self.text_resultado.delete("1.0", tk.END)
            self.df_original = None

    def _popular_filtros(self):
        if self.df_original is not None:
            contratos = sorted(self.df_original['Contrato'].unique()); self.combo_contrato['values'] = [""] + contratos
            self.lista_funcionarios_completa = sorted(self.df_original['Funcionário'].unique()); self.combo_funcionario['values'] = [""] + self.lista_funcionarios_completa

    def _aplicar_filtros(self, df: pd.DataFrame) -> pd.DataFrame:
        df_filtrado = df.copy()
        coluna_data = 'DATA'
        
        # ##################################################################
        # ##                     INÍCIO DA CORREÇÃO                     ##
        # ##################################################################
        if coluna_data in df_filtrado.columns:
            try:
                # Pega a data dos widgets (são objetos datetime.date)
                start_date = self.filtro_data_ini.get_date()
                end_date = self.filtro_data_fim.get_date()
                
                # Garante que a coluna de data não tenha valores nulos de data
                df_filtrado_sem_na = df_filtrado.dropna(subset=[coluna_data])
                
                # COMPARA a parte da data da coluna (.dt.date) com a data do widget
                df_filtrado = df_filtrado_sem_na[
                    (df_filtrado_sem_na[coluna_data].dt.date >= start_date) & 
                    (df_filtrado_sem_na[coluna_data].dt.date <= end_date)
                ]
            except Exception as e:
                 messagebox.showwarning("Aviso de Filtro de Data", f"Não foi possível aplicar o filtro de data.\nErro: {e}")
        # ##################################################################
        # ##                       FIM DA CORREÇÃO                      ##
        # ##################################################################

        f_contrato = self.filtro_contrato.get()
        if f_contrato: df_filtrado = df_filtrado[df_filtrado['Contrato'] == f_contrato]
        f_funcionario = self.filtro_funcionario.get()
        if f_funcionario: df_filtrado = df_filtrado[df_filtrado['Funcionário'] == f_funcionario]
            
        return df_filtrado

    def limpar_filtros(self):
        self.filtro_contrato.set(""); self.filtro_funcionario.set("")
        self._processar_dataframe_carregado()
        self.text_resultado.delete("1.0", tk.END)

    def atualizar_filtro_funcionario(self, event=None):
        texto_digitado = self.filtro_funcionario.get().lower()
        if not texto_digitado: self.combo_funcionario['values'] = [""] + self.lista_funcionarios_completa
        else: self.combo_funcionario['values'] = [nome for nome in self.lista_funcionarios_completa if texto_digitado in nome.lower()]

    def _limpar_estado_relatorio(self):
        self.dados_relatorio = None; self.btn_export_excel.config(state=tk.DISABLED); self.btn_export_pdf.config(state=tk.DISABLED)

    def gerar_relatorio(self):
        if self.df_original is None: messagebox.showwarning("Aviso", "Por favor, selecione e carregue um arquivo primeiro."); return
        self._limpar_estado_relatorio()
        self.text_resultado.delete("1.0", tk.END); self.text_resultado.insert(tk.END, "Processando dados com filtros, por favor aguarde..."); self.root.update_idletasks()
        try:
            df_filtrado = self._aplicar_filtros(self.df_original)
            df_grouped = _agrupar_por_registro(df_filtrado)
            self.dados_relatorio = _montar_linhas_com_subtotais(df_grouped)
            resultado_str = _formatar_para_string(self.dados_relatorio)
            self.text_resultado.delete("1.0", tk.END); self.text_resultado.insert(tk.END, resultado_str)
            if self.dados_relatorio: self.btn_export_excel.config(state=tk.NORMAL); self.btn_export_pdf.config(state=tk.NORMAL)
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao gerar o relatório: {e}"); self.text_resultado.delete("1.0", tk.END); self._limpar_estado_relatorio()

    # As funções de exportação (exportar_para_excel e exportar_para_pdf) permanecem as mesmas
    def exportar_para_excel(self):
        if not self.dados_relatorio: messagebox.showwarning("Aviso", "Não há dados para exportar."); return
        caminho_salvar = filedialog.asksaveasfilename(title="Salvar como Excel", defaultextension=".xlsx", filetypes=[("Excel Workbook", "*.xlsx")])
        if not caminho_salvar: return
        try:
            df_export = pd.DataFrame(self.dados_relatorio); df_export['S.A.'] = pd.to_numeric(df_export['S.A.'], errors='coerce')
            with pd.ExcelWriter(caminho_salvar, engine='openpyxl') as writer:
                df_export.to_excel(writer, index=False, sheet_name="Relatorio")
                workbook, worksheet = writer.book, writer.sheets["Relatorio"]
                col_idx_sa = df_export.columns.get_loc('S.A.') + 1
                for cell in worksheet.iter_cols(min_col=col_idx_sa, max_col=col_idx_sa, min_row=2): cell[0].number_format = '0.000'
                for column_cells in worksheet.columns:
                    max_length = 0; column_letter = column_cells[0].column_letter
                    for cell in column_cells:
                        try:
                            if len(str(cell.value)) > max_length: max_length = len(str(cell.value))
                        except: pass
                    worksheet.column_dimensions[column_letter].width = (max_length + 2)
            messagebox.showinfo("Sucesso", f"Relatório exportado com sucesso para:\n{caminho_salvar}")
        except Exception as e: messagebox.showerror("Erro ao Exportar", f"Não foi possível salvar o arquivo Excel:\n{e}")

    def exportar_para_pdf(self):
        if not self.dados_relatorio: messagebox.showwarning("Aviso", "Não há dados para exportar."); return
        caminho_salvar = filedialog.asksaveasfilename(title="Salvar como PDF", defaultextension=".pdf", filetypes=[("PDF Document", "*.pdf")])
        if not caminho_salvar: return
        try:
            doc = SimpleDocTemplate(caminho_salvar, pagesize=landscape(letter)); story = []; styles = getSampleSheetStyleSheet()
            story.append(Paragraph("Relatório por Registro", styles['h1'])); story.append(Spacer(1, 0.2 * inch))
            headers = list(self.dados_relatorio[0].keys()); data_for_pdf = [headers]
            for row_dict in self.dados_relatorio:
                sa_val = row_dict.get('S.A.', ''); sa_formatted = str(sa_val)
                try: sa_formatted = f"{float(sa_val):.3f}"
                except (ValueError, TypeError): pass
                row_list = list(row_dict.values()); row_list[-1] = sa_formatted; data_for_pdf.append(row_list)
            table = Table(data_for_pdf, repeatRows=1)
            style = TableStyle([('BACKGROUND', (0, 0), (-1, 0), colors.grey), ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke), ('ALIGN', (0, 0), (-1, -1), 'CENTER'), ('ALIGN', (0, 1), (0, -1), 'LEFT'), ('ALIGN', (4, 1), (4, -1), 'RIGHT'), ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'), ('BOTTOMPADDING', (0, 0), (-1, 0), 12), ('BACKGROUND', (0, 1), (-1, -1), colors.beige), ('GRID', (0, 0), (-1, -1), 1, colors.black)])
            table.setStyle(style); story.append(table); doc.build(story)
            messagebox.showinfo("Sucesso", f"Relatório exportado com sucesso para:\n{caminho_salvar}")
        except Exception as e: messagebox.showerror("Erro ao Exportar", f"Não foi possível salvar o arquivo PDF:\n{e}")





if __name__ == "__main__":
    root = tk.Tk()
    app = RelatorioApp(root)
    root.mainloop()

# NO ARQUIVO: gui_relat.py