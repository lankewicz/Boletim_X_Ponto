#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
parquet_to_json_gui.py
----------------------
Interface gráfica para conversão de .parquet -> .json/.jsonl
Com visualização de estrutura e preview dos dados
"""
from __future__ import annotations

import gzip
import json
import threading
from base64 import b64encode
from decimal import Decimal
from pathlib import Path
from tkinter import *
from tkinter import filedialog, messagebox, ttk
from typing import Iterable, List

import pandas as pd
import pyarrow as pa
import pyarrow.parquet as pq


class ParquetConverterGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Parquet → JSON Converter")
        self.root.geometry("1200x700")
        
        # Variáveis
        self.input_path = StringVar()
        self.output_path = StringVar(value="json_out")
        self.format_var = StringVar(value="jsonl")
        self.gzip_var = BooleanVar(value=False)
        self.recursive_var = BooleanVar(value=False)
        self.batch_size = IntVar(value=100000)
        self.current_df = None
        self.current_file = None
        
        self.setup_ui()
        
    def setup_ui(self):
        # Container principal com 2 colunas - pesos fixos
        self.root.columnconfigure(0, weight=0)
        self.root.rowconfigure(0, weight=1)
        
        # PanedWindow para manter proporções fixas
        paned = ttk.PanedWindow(self.root, orient=HORIZONTAL)
        paned.grid(row=0, column=0, sticky=(N, W, E, S))
        
        main_container = ttk.Frame(paned)
        paned.add(main_container, weight=0)
        
        # Container com colunas fixas
        main_container.columnconfigure(0, weight=0, minsize=400)
        main_container.columnconfigure(1, weight=1, minsize=750)
        main_container.rowconfigure(0, weight=1)
        
        # === PAINEL ESQUERDO: Controles ===
        left_panel = ttk.Frame(main_container, padding="10", width=400)
        left_panel.grid(row=0, column=0, sticky=(N, W, E, S))
        left_panel.grid_propagate(False)  # Manter largura fixa
        
        # Título
        ttk.Label(left_panel, text="Conversor Parquet → JSON", 
                 font=("Segoe UI", 14, "bold")).grid(row=0, column=0, columnspan=3, pady=(0, 15))
        
        # Input
        ttk.Label(left_panel, text="Arquivo Parquet:", 
                 font=("Segoe UI", 9, "bold")).grid(row=1, column=0, sticky=W, pady=5)
        ttk.Entry(left_panel, textvariable=self.input_path, width=35).grid(
            row=2, column=0, columnspan=2, sticky=(W, E), pady=5)
        ttk.Button(left_panel, text="📂", width=3,
                  command=self.browse_input).grid(row=2, column=2, padx=(5, 0))
        
        ttk.Button(left_panel, text="🔍 Carregar e Visualizar", 
                  command=self.load_preview).grid(row=3, column=0, columnspan=3, 
                                                  pady=10, sticky=(W, E))
        
        ttk.Separator(left_panel, orient=HORIZONTAL).grid(row=4, column=0, 
                                                          columnspan=3, sticky=(W, E), pady=15)
        
        # Estrutura do arquivo
        struct_header = ttk.Frame(left_panel)
        struct_header.grid(row=5, column=0, columnspan=3, sticky=(W, E), pady=5)
        
        ttk.Label(struct_header, text="Estrutura do Arquivo:", 
                 font=("Segoe UI", 9, "bold")).pack(side=LEFT)
        ttk.Button(struct_header, text="📋 Copiar", width=10,
                  command=self.copy_structure).pack(side=RIGHT)
        
        struct_frame = ttk.Frame(left_panel)
        struct_frame.grid(row=6, column=0, columnspan=3, sticky=(W, E, N, S), pady=5)
        
        self.struct_tree = ttk.Treeview(struct_frame, height=10, 
                                        columns=("tipo", "exemplo"), show="tree headings")
        self.struct_tree.heading("#0", text="Coluna")
        self.struct_tree.heading("tipo", text="Tipo")
        self.struct_tree.heading("exemplo", text="Exemplo")
        self.struct_tree.column("#0", width=150)
        self.struct_tree.column("tipo", width=100)
        self.struct_tree.column("exemplo", width=150)
        
        struct_scroll = ttk.Scrollbar(struct_frame, orient=VERTICAL, 
                                     command=self.struct_tree.yview)
        self.struct_tree.configure(yscrollcommand=struct_scroll.set)
        
        self.struct_tree.grid(row=0, column=0, sticky=(W, E, N, S))
        struct_scroll.grid(row=0, column=1, sticky=(N, S))
        struct_frame.columnconfigure(0, weight=1)
        struct_frame.rowconfigure(0, weight=1)
        
        # Info do arquivo
        self.info_label = ttk.Label(left_panel, text="", foreground="gray")
        self.info_label.grid(row=7, column=0, columnspan=3, sticky=W, pady=5)
        
        ttk.Separator(left_panel, orient=HORIZONTAL).grid(row=8, column=0, 
                                                          columnspan=3, sticky=(W, E), pady=15)
        
        # Opções de conversão
        ttk.Label(left_panel, text="Opções de Exportação:", 
                 font=("Segoe UI", 9, "bold")).grid(row=9, column=0, sticky=W, pady=5)
        
        options_frame = ttk.Frame(left_panel)
        options_frame.grid(row=10, column=0, columnspan=3, sticky=(W, E), pady=5)
        
        ttk.Radiobutton(options_frame, text="JSONL", 
                       variable=self.format_var, value="jsonl").grid(row=0, column=0, sticky=W)
        ttk.Radiobutton(options_frame, text="JSON Array", 
                       variable=self.format_var, value="json").grid(row=0, column=1, sticky=W)
        
        ttk.Checkbutton(options_frame, text="Gzip", 
                       variable=self.gzip_var).grid(row=1, column=0, sticky=W, pady=3)
        
        # Output
        ttk.Label(left_panel, text="Pasta de saída:").grid(row=11, column=0, sticky=W, pady=(10, 5))
        ttk.Entry(left_panel, textvariable=self.output_path, width=35).grid(
            row=12, column=0, columnspan=2, sticky=(W, E), pady=5)
        ttk.Button(left_panel, text="📂", width=3,
                  command=self.browse_output).grid(row=12, column=2, padx=(5, 0))
        
        # Botão exportar
        self.export_btn = ttk.Button(left_panel, text="💾 Exportar para Arquivo", 
                                     command=self.export_file, state=DISABLED)
        self.export_btn.grid(row=13, column=0, columnspan=3, pady=15, sticky=(W, E))
        
        left_panel.rowconfigure(6, weight=1)
        
        # === PAINEL DIREITO: Preview dos Dados ===
        right_panel = ttk.Frame(main_container, padding="10")
        right_panel.grid(row=0, column=1, sticky=(N, W, E, S))
        right_panel.columnconfigure(0, weight=1)
        right_panel.rowconfigure(1, weight=1)
        
        # Título do preview
        preview_header = ttk.Frame(right_panel)
        preview_header.grid(row=0, column=0, sticky=(W, E), pady=(0, 10))
        
        ttk.Label(preview_header, text="Preview dos Dados", 
                 font=("Segoe UI", 12, "bold")).pack(side=LEFT)
        
        self.preview_info = ttk.Label(preview_header, text="(primeiras 100 linhas)", 
                                     foreground="gray")
        self.preview_info.pack(side=LEFT, padx=10)
        
        # Notebook para diferentes visualizações
        self.notebook = ttk.Notebook(right_panel)
        self.notebook.grid(row=1, column=0, sticky=(N, W, E, S))
        
        # Tab 1: Tabela
        table_frame = ttk.Frame(self.notebook)
        self.notebook.add(table_frame, text="📊 Tabela")
        
        # Treeview para dados
        self.data_tree = ttk.Treeview(table_frame, show="tree headings")
        
        data_scroll_y = ttk.Scrollbar(table_frame, orient=VERTICAL, 
                                     command=self.data_tree.yview)
        data_scroll_x = ttk.Scrollbar(table_frame, orient=HORIZONTAL, 
                                     command=self.data_tree.xview)
        self.data_tree.configure(yscrollcommand=data_scroll_y.set,
                                xscrollcommand=data_scroll_x.set)
        
        self.data_tree.grid(row=0, column=0, sticky=(N, W, E, S))
        data_scroll_y.grid(row=0, column=1, sticky=(N, S))
        data_scroll_x.grid(row=1, column=0, sticky=(W, E))
        
        table_frame.columnconfigure(0, weight=1)
        table_frame.rowconfigure(0, weight=1)
        
        # Tab 2: JSON Preview
        json_frame = ttk.Frame(self.notebook)
        self.notebook.add(json_frame, text="{ } JSON")
        
        self.json_text = Text(json_frame, wrap=NONE, font=("Consolas", 9))
        json_scroll_y = ttk.Scrollbar(json_frame, orient=VERTICAL, 
                                     command=self.json_text.yview)
        json_scroll_x = ttk.Scrollbar(json_frame, orient=HORIZONTAL, 
                                     command=self.json_text.xview)
        self.json_text.configure(yscrollcommand=json_scroll_y.set,
                                xscrollcommand=json_scroll_x.set)
        
        self.json_text.grid(row=0, column=0, sticky=(N, W, E, S))
        json_scroll_y.grid(row=0, column=1, sticky=(N, S))
        json_scroll_x.grid(row=1, column=0, sticky=(W, E))
        
        json_frame.columnconfigure(0, weight=1)
        json_frame.rowconfigure(0, weight=1)
        
        # Progress bar
        self.progress = ttk.Progressbar(right_panel, mode='indeterminate')
        self.progress.grid(row=2, column=0, sticky=(W, E), pady=10)
        
        # Status
        self.status_label = ttk.Label(right_panel, text="Carregue um arquivo para visualizar", 
                                     foreground="gray")
        self.status_label.grid(row=3, column=0, sticky=W)
        
    def browse_input(self):
        path = filedialog.askopenfilename(
            title="Selecione um arquivo Parquet",
            filetypes=[("Parquet files", "*.parquet"), ("All files", "*.*")]
        )
        if path:
            self.input_path.set(path)
            
    def browse_output(self):
        path = filedialog.askdirectory(title="Selecione a pasta de saída")
        if path:
            self.output_path.set(path)
    
    def copy_structure(self):
        """Copia a estrutura do arquivo para a área de transferência"""
        if self.current_file is None:
            messagebox.showwarning("Aviso", "Carregue um arquivo primeiro!")
            return
        
        try:
            # Montar texto da estrutura
            lines = [f"Arquivo: {self.current_file.name}", ""]
            lines.append(f"{'Coluna':<40} {'Tipo':<20} {'Exemplo'}")
            lines.append("-" * 100)
            
            for item in self.struct_tree.get_children():
                coluna = self.struct_tree.item(item)["text"]
                valores = self.struct_tree.item(item)["values"]
                tipo = valores[0] if len(valores) > 0 else ""
                exemplo = valores[1] if len(valores) > 1 else ""
                lines.append(f"{coluna:<40} {tipo:<20} {exemplo}")
            
            # Copiar para clipboard
            text = "\n".join(lines)
            self.root.clipboard_clear()
            self.root.clipboard_append(text)
            self.root.update()
            
            # Feedback visual
            self.status_label.config(text="✓ Estrutura copiada para área de transferência!")
            
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao copiar: {str(e)}")
            
    def load_preview(self):
        if not self.input_path.get():
            messagebox.showerror("Erro", "Selecione um arquivo Parquet!")
            return
            
        self.status_label.config(text="Carregando...")
        self.progress.start(10)
        
        thread = threading.Thread(target=self._load_preview_thread, daemon=True)
        thread.start()
        
    def _load_preview_thread(self):
        try:
            file_path = Path(self.input_path.get())
            self.current_file = file_path
            
            # Ler amostra
            pf = pq.ParquetFile(str(file_path))
            table = pf.read_row_groups([0], use_threads=True)
            df = table.to_pandas()
            
            # Limitar a 100 linhas para preview
            preview_df = df.head(100)
            self.current_df = preview_df
            
            # Atualizar UI na thread principal
            self.root.after(0, self._update_preview_ui, df, preview_df, pf)
            
        except Exception as e:
            self.root.after(0, self._show_error, f"Erro ao carregar: {str(e)}")
            
    def _update_preview_ui(self, df_full, df_preview, pf):
        try:
            # Limpar estrutura anterior
            for item in self.struct_tree.get_children():
                self.struct_tree.delete(item)
                
            # Preencher estrutura
            for col in df_full.columns:
                dtype = str(df_full[col].dtype)
                exemplo = str(df_preview[col].iloc[0]) if len(df_preview) > 0 else ""
                if len(exemplo) > 50:
                    exemplo = exemplo[:47] + "..."
                self.struct_tree.insert("", END, text=col, values=(dtype, exemplo))
                
            # Info do arquivo
            total_rows = pf.metadata.num_rows
            num_cols = len(df_full.columns)
            file_size = self.current_file.stat().st_size / (1024*1024)
            self.info_label.config(text=f"📊 {total_rows:,} linhas × {num_cols} colunas | 💾 {file_size:.2f} MB")
            
            # Atualizar preview da tabela
            self._update_table_view(df_preview)
            
            # Atualizar preview JSON
            self._update_json_view(df_preview)
            
            self.export_btn.config(state=NORMAL)
            self.status_label.config(text=f"✓ Arquivo carregado: {self.current_file.name}")
            
        except Exception as e:
            self._show_error(f"Erro ao exibir preview: {str(e)}")
        finally:
            self.progress.stop()
            
    def _update_table_view(self, df):
        # Limpar tabela
        self.data_tree.delete(*self.data_tree.get_children())
        
        # Configurar colunas
        cols = list(df.columns)
        self.data_tree["columns"] = cols
        self.data_tree.column("#0", width=50, anchor=W)
        self.data_tree.heading("#0", text="#")
        
        for col in cols:
            self.data_tree.heading(col, text=col)
            self.data_tree.column(col, width=120, anchor=W)
            
        # Preencher dados
        for idx, row in df.iterrows():
            values = [str(v)[:100] if v is not None else "" for v in row]
            self.data_tree.insert("", END, text=str(idx), values=values)
            
    def _update_json_view(self, df):
        self.json_text.delete(1.0, END)
        
        # Normalizar e converter para JSON
        df_norm = self._normalize_df(df.copy())
        records = df_norm.head(5).to_dict(orient="records")
        json_str = json.dumps(records, ensure_ascii=False, indent=2)
        
        self.json_text.insert(1.0, json_str)
        self.json_text.insert(END, f"\n\n... (mostrando apenas 5 de {len(df)} registros)")
        
    def _show_error(self, message):
        self.progress.stop()
        self.status_label.config(text=f"✗ {message}")
        messagebox.showerror("Erro", message)
        
    def export_file(self):
        if self.current_file is None:
            messagebox.showerror("Erro", "Carregue um arquivo primeiro!")
            return
            
        self.status_label.config(text="Exportando...")
        self.progress.start(10)
        self.export_btn.config(state=DISABLED)
        
        thread = threading.Thread(target=self._export_thread, daemon=True)
        thread.start()
        
    def _export_thread(self):
        try:
            output_dir = Path(self.output_path.get())
            out_path = self.convert_file(self.current_file, output_dir)
            
            self.root.after(0, self._export_success, out_path)
            
        except Exception as e:
            self.root.after(0, self._show_error, f"Erro ao exportar: {str(e)}")
            self.root.after(0, lambda: self.export_btn.config(state=NORMAL))
            
    def _export_success(self, out_path):
        self.progress.stop()
        self.export_btn.config(state=NORMAL)
        self.status_label.config(text=f"✓ Exportado: {out_path.name}")
        messagebox.showinfo("Sucesso", f"Arquivo exportado com sucesso!\n\n{out_path}")
        
    def convert_file(self, in_path: Path, out_dir: Path) -> Path:
        out_dir.mkdir(parents=True, exist_ok=True)
        suffix = ".jsonl" if self.format_var.get() == "jsonl" else ".json"
        out_path = out_dir / f"{in_path.stem}{suffix}"
        
        if self.gzip_var.get():
            out_path = out_path.with_suffix(out_path.suffix + ".gz")
            
        opener = (lambda p: gzip.open(p, "wt", encoding="utf-8")) if self.gzip_var.get() \
                 else (lambda p: open(p, "w", encoding="utf-8"))
                 
        with opener(out_path) as fh:
            if self.format_var.get() == "jsonl":
                for batch in self._iter_parquet_batches(in_path):
                    df = batch.to_pandas(types_mapper=None)
                    self._write_jsonl(df, fh)
            else:
                def _batches_df():
                    for batch in self._iter_parquet_batches(in_path):
                        yield batch.to_pandas(types_mapper=None)
                self._write_json_array_stream(_batches_df(), fh)
                
        return out_path
        
    def _iter_parquet_batches(self, path: Path) -> Iterable[pa.RecordBatch]:
        pf = pq.ParquetFile(str(path))
        for batch in pf.iter_batches(batch_size=self.batch_size.get()):
            yield batch
            
    def _normalize_df(self, df: pd.DataFrame) -> pd.DataFrame:
        if df.empty:
            return df
            
        for col in df.columns:
            s = df[col]
            if pd.api.types.is_object_dtype(s):
                sample = None
                try:
                    for v in s.array:
                        if pd.notna(v):
                            sample = v
                            break
                except Exception:
                    for v in s:
                        if pd.notna(v):
                            sample = v
                            break
                            
                if isinstance(sample, Decimal):
                    df[col] = s.map(lambda x: (str(x) if isinstance(x, Decimal) 
                                              else (None if pd.isna(x) else x)))
                elif isinstance(sample, (bytes, bytearray, memoryview)):
                    df[col] = s.map(lambda x: (b64encode(x).decode("ascii") 
                                              if isinstance(x, (bytes, bytearray, memoryview)) 
                                              else (None if pd.isna(x) else x)))
        return df
        
    def _write_jsonl(self, df: pd.DataFrame, fh) -> None:
        if df.empty:
            return
        df = self._normalize_df(df)
        df.to_json(fh, orient="records", lines=True, force_ascii=False,
                  date_format="iso", date_unit="ms")
                  
    def _write_json_array_stream(self, batches: Iterable[pd.DataFrame], fh) -> None:
        fh.write("[")
        first = True
        for df in batches:
            if df is None or df.empty:
                continue
            df = self._normalize_df(df)
            records = df.to_dict(orient="records")
            for rec in records:
                if not first:
                    fh.write(",")
                fh.write(json.dumps(rec, ensure_ascii=False, allow_nan=False))
                first = False
        fh.write("]")


def main():
    root = Tk()
    app = ParquetConverterGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()