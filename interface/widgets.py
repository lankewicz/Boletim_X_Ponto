# interface/widgets.py
# Responsável por montar a interface (janela, abas, botões e grids).
# Mantém apenas lógica de UI; toda a lógica de dados fica em app.py.
#
# Autor: Valdinei Lankewicz
# Criado em: 22/07/2025
# Histórico:
# - 10/08/2025: Revisão geral para remover chamadas inválidas a Sheet.headers().
# - 10/08/2025: Padronização de estilos e alinhamentos nas Sheets.
# - 10/08/2025: Inclusão de aba ROTÁLOG (vazia, por enquanto).
# - 10/08/2025: Abas agora expõem seus Frames em app (app.aba_*), para uso com notebook.tab().

import tkinter as tk
from datetime import datetime
from tkinter import ttk

import ttkbootstrap as ttkb

from utils.formatadores import ocultar_indices_sheet

try:
    # tksheet é o grid usado para exibir as tabelas
    from tksheet import Sheet
except Exception:  # fallback (para ferramentas de análise)
    Sheet = None


# =============================================================================
# Helpers de formatação/estilo de Sheet (tksheet)
# =============================================================================
def _configurar_sheet(
    sheet: "Sheet",
    headers: list[str],
    *,
    header_bg: str = "",
    first_col_width: bool = False,
    first_col_width_px: int = 120,
    default_col_width_px: int = 90,
    right_align_from: int = 0,
    header_wrap: bool = True,
    header_align_center: bool = True,
    header_height_lines: int | None = 2,
    enable_bindings: bool = True,
) -> None:
    """
    Aplica cabeçalhos e opções visuais em uma Sheet.

    Parâmetros principais:
    - headers: lista de cabeçalhos.
    - header_bg: cor de fundo do cabeçalho (string HEX).
    - first_col_width: se True, aplica 'first_col_width_px' à coluna 0.
    - right_align_from: alinha à direita todas as colunas a partir desse índice.
    - header_height_lines: tenta aumentar a altura do cabeçalho para N linhas.
    """
    if sheet is None:
        return

    # 1) Define os cabeçalhos (API correta)
    try:
        sheet.headers(headers)
    except Exception:
        try:
            sheet.set_options(headers=headers)
        except Exception:
            pass

    # 2) Estilo do cabeçalho
    if header_bg:
        try:
            sheet.headers_bg(header_bg)
        except Exception:
            try:
                sheet.set_options(header_bg=header_bg)
            except Exception:
                pass

    if header_wrap:
        try:
            sheet.headers_wrap = True
        except Exception:
            pass

    if header_align_center:
        try:
            sheet.headers_align = "center"
        except Exception:
            pass

    if header_height_lines is not None:
        try:
            sheet.set_header_height_lines(header_height_lines)
        except Exception:
            pass

    # 3) Larguras de coluna
    try:
        total_cols = sheet.get_total_columns()
    except Exception:
        data = sheet.get_sheet_data(return_copy=True)
        total_cols = len(data[0]) if data else len(headers)

    for c in range(total_cols):
        width = default_col_width_px
        if first_col_width and c == 0:
            width = first_col_width_px
        try:
            sheet.column_width(column=c, width=width)
        except Exception:
            try:
                sheet.set_column_width(c, width)
            except Exception:
                pass

    # 4) Alinhamento à direita das colunas numéricas
    for c in range(max(0, right_align_from), total_cols):
        try:
            sheet.align_column(c, align="e")  # east/right
        except Exception:
            pass

    # 5) Liga/desliga *bindings* úteis
    if enable_bindings:
        try:
            sheet.enable_bindings(
                (
                    "single_select",
                    "row_select",
                    "arrowkeys",
                    "right_click_popup_menu",
                    "rc_select",
                    "rc_insert_row",
                    "rc_delete_row",
                    "copy",
                    "cut",
                    "paste",
                    "edit_cell",
                    "drag_select",
                    "select_all",
                )
            )
        except Exception:
            pass


# =============================================================================
# Sincronização de seleção entre as 3 sheets da aba Comparação
# =============================================================================
def bind_comparison_sync(app):
    """
    Liga bindings de seleção entre as três sheets:
      - app.sheet_comp_boletim (tem 'Data' na col 0 e 6 métricas nas cols 1..6)
      - app.sheet_comp_ponto    (6 métricas nas cols 0..5)
      - app.sheet_comp_diferenca(6 métricas nas cols 0..5)

    Comportamentos:
      * Clicar na 'Data' (col 0 do Boletim) seleciona a mesma LINHA nas 3 (todas as métricas).
      * Clicar em qualquer métrica seleciona a mesma métrica e linha nas 3.
    À prova de versão do tksheet: tenta várias APIs (get_currently_selected, selected_cells, select_cells, etc).
    """

    # --- Helpers à prova de versão -------------------------------------------
    def _get_sel_rc(sheet):
        """Retorna (row, col) da célula atual, se possível."""
        # tksheet >= 6.x
        try:
            rc = sheet.get_currently_selected()
            if isinstance(rc, (list, tuple)) and len(rc) == 2:
                return rc[0], rc[1]
        except Exception:
            pass
        # fallback: selected_cells (set de tuples)
        try:
            sel = getattr(sheet, "selected_cells", None)
            if sel:
                r, c = list(sel)[0]
                return r, c
        except Exception:
            pass
        # fallback: selected (string como "row,column")
        try:
            s = sheet.get_selected()
            if isinstance(s, str) and "," in s:
                rr, cc = s.split(",", 1)
                return int(rr), int(cc)
        except Exception:
            pass
        return None, None

    def _deselect_all(sheet):
        for m in ("deselect", "de_select", "clear_selected"):
            try:
                getattr(sheet, m)("all")
                return
            except Exception:
                pass
        # último recurso: nada

    def _see(sheet, r, c):
        for m in ("see", "see_cell"):
            try:
                getattr(sheet, m)(r, c)
                return
            except Exception:
                pass

    def _select_cell(sheet, r, c, keep=False):
        # tenta seleção em bloco
        for m in ("select_cell",):
            try:
                getattr(sheet, m)(r, c, redraw=True, keep_other_selections=keep)
                return
            except Exception:
                pass
        # fallback: select_cells(r1,c1,r2,c2)
        try:
            sheet.select_cells(r, c, r, c, redraw=True)
        except Exception:
            pass

    def _select_block(sheet, r, c1, c2):
        # seleciona de (r,c1) até (r,c2)
        for m in ("select_cells",):
            try:
                getattr(sheet, m)(r, c1, r, c2, redraw=True)
                return
            except Exception:
                pass
        # fallback: iterar célula a célula
        _deselect_all(sheet)
        for c in range(c1, c2 + 1):
            _select_cell(sheet, r, c, keep=(c > c1))

    # --- Seleção por linha inteira (todas as métricas) ------------------------
    def _select_row_metrics(row):
        # Boletim: métricas nas cols 1..6
        try:
            _select_block(app.sheet_comp_boletim, row, 1, 6)
            _see(app.sheet_comp_boletim, row, 1)
        except Exception:
            pass
        # Ponto: 0..5
        try:
            _select_block(app.sheet_comp_ponto, row, 0, 5)
            _see(app.sheet_comp_ponto, row, 0)
        except Exception:
            pass
        # Diferença: 0..5
        try:
            _select_block(app.sheet_comp_diferenca, row, 0, 5)
            _see(app.sheet_comp_diferenca, row, 0)
        except Exception:
            pass

    # --- Seleção por métrica específica --------------------------------------
    def _sync_select_from(kind, row, col):
        """
        kind: 'bol' | 'pto' | 'dif'
        Converte (row,col) para o índice de métrica 0..5 e replica nas 3 sheets.
        Clique na Data (Boletim, col=0) => seleciona linha toda.
        """
        if row is None or col is None:
            return

        # clique na Data do Boletim => linha inteira
        if kind == "bol" and col == 0:
            _select_row_metrics(row)
            return

        # métrica 0..5
        metric = (col - 1) if kind == "bol" else col
        if metric < 0 or metric > 5:
            return

        # Boletim (offset +1)
        try:
            _deselect_all(app.sheet_comp_boletim)
            _select_cell(app.sheet_comp_boletim, row, metric + 1)
            _see(app.sheet_comp_boletim, row, metric + 1)
        except Exception:
            pass

        # Ponto
        try:
            _deselect_all(app.sheet_comp_ponto)
            _select_cell(app.sheet_comp_ponto, row, metric)
            _see(app.sheet_comp_ponto, row, metric)
        except Exception:
            pass

        # Diferença
        try:
            _deselect_all(app.sheet_comp_diferenca)
            _select_cell(app.sheet_comp_diferenca, row, metric)
            _see(app.sheet_comp_diferenca, row, metric)
        except Exception:
            pass

    # --- Handlers de evento (compatíveis com várias versões) -----------------
    def _mk_handler(sheet, kind):
        def _on_select(event=None):
            r, c = _get_sel_rc(sheet)
            if r is None or c is None:
                return
            _sync_select_from(kind, r, c)

        return _on_select

    # evita re-binds repetidos
    if getattr(app, "_cmp_sync_bound", False):
        return

    # liga nos três (se existirem)
    bindings = ("cell_select", "row_select", "table_select", "<<SheetSelect>>")
    try:
        if app.sheet_comp_boletim:
            for b in bindings:
                try:
                    app.sheet_comp_boletim.extra_bindings(
                        [(b, _mk_handler(app.sheet_comp_boletim, "bol"))]
                    )
                except Exception:
                    pass
        if app.sheet_comp_ponto:
            for b in bindings:
                try:
                    app.sheet_comp_ponto.extra_bindings(
                        [(b, _mk_handler(app.sheet_comp_ponto, "pto"))]
                    )
                except Exception:
                    pass
        if app.sheet_comp_diferenca:
            for b in bindings:
                try:
                    app.sheet_comp_diferenca.extra_bindings(
                        [(b, _mk_handler(app.sheet_comp_diferenca, "dif"))]
                    )
                except Exception:
                    pass
        app._cmp_sync_bound = True
    except Exception:
        # não falha a UI se algo der errado
        app._cmp_sync_bound = False


# =============================================================================
# Seção de construção da janela/estrutura básica
# =============================================================================
def criar_interface(app) -> ttkb.Window:
    """
    Cria e retorna a janela principal.
    Este método amarra os componentes (widgets) no 'app' para que o app.py
    possa manipulá-los (gerar relatórios, formatar sheets etc).
    """

    # ---------------- Janela Principal ----------------
    janela = ttkb.Window(themename="flatly")
    janela.title("Relatório de Horas - Comparação Boletim x Ponto")
    janela.geometry("1200x720")

    # Variável de formato de horas (compartilhada com o app)
    # False -> decimal   |   True -> HH:MM
    if not hasattr(app, "var_formato_hora"):
        app.var_formato_hora = tk.BooleanVar(value=False)

    # ---------------- Topbar (Filtros) ----------------
    frame_top = ttkb.Frame(janela, padding=(10, 6))
    frame_top.pack(side=tk.TOP, fill=tk.X)

    # Botões de navegação de funcionário


    # Funcionário (combobox)
    ttkb.Label(frame_top, text="Funcionário:").pack(side=tk.LEFT, padx=(0, 6))
    ttkb.Button(frame_top, text="◀", width=3, command=lambda: app.mudar_funcionario(-1)).pack(side=tk.LEFT, padx=(0, 4))

    app.combo_funcionario = ttkb.Combobox(frame_top, width=40)
    app.combo_funcionario.pack(side=tk.LEFT, padx=(0, 10))
    app.combo_funcionario.bind("<<ComboboxSelected>>", lambda e: app._on_funcionario_trocado())

    app._on_funcionario_trocado()



    # Botões de navegação de funcionário
    ttkb.Button(frame_top, text="▶", width=3, command=lambda: app.mudar_funcionario(+1)).pack(side=tk.LEFT, padx=(0, 10))

    # ---- Contrato + Filtrar contrato ----
    # expose a toolbar frame pro app, se ainda não tiver
    app.toolbar = frame_top

    # variável de filtro (se o app ainda não criou)
    if not hasattr(app, "var_filtrar_contrato") or app.var_filtrar_contrato is None:
        app.var_filtrar_contrato = tk.BooleanVar(value=False)

    ttkb.Checkbutton(
        frame_top,
        text="Filtrar contrato:",
        variable=app.var_filtrar_contrato,
        command=getattr(app, "_on_toggle_filtrar_contrato", lambda: None),
        bootstyle="secondary",
    ).pack(side=tk.LEFT, padx=(0, 10))

#    ttkb.Label(frame_top, text="Contrato:").pack(side=tk.LEFT, padx=(10, 6))
    app.combo_contrato = ttkb.Combobox(frame_top, width=24, state="disabled")
    app.combo_contrato.pack(side=tk.LEFT, padx=(0, 8))
    # quando escolher um contrato, recarrega a lista de funcionários
    app.combo_contrato.bind("<<ComboboxSelected>>", lambda e: app.atualizar_combo_funcionarios())



    # Data Inicial
    app.cal_data_inicial = ttkb.DateEntry(
        frame_top,
        width=12,
        bootstyle="info",
        dateformat="%d/%m/%Y",  # <- formato fixo (evita warning do %x)
    )
    app.cal_data_inicial.set_date(datetime.today())
    app.cal_data_inicial.pack(side=tk.LEFT, padx=5)

    # Data Final
    app.cal_data_final = ttkb.DateEntry(
        frame_top,
        width=12,
        bootstyle="info",
        dateformat="%d/%m/%Y",  # <- formato fixo
    )
    app.cal_data_final.set_date(datetime.today())
    app.cal_data_final.pack(side=tk.LEFT, padx=5)

    # Formato da hora (alternar)
    chk = ttkb.Checkbutton(
        frame_top,
        text="Exibir em HH:MM",
        variable=app.var_formato_hora,
        command=getattr(app, "trocar_formato_hora", lambda: None),
        bootstyle="secondary-round-toggle",
    )
    chk.pack(side=tk.LEFT, padx=(10, 10))

    # Botões rápidos (atualizar/diagnóstico)
    ttkb.Button(
        frame_top, text="Atualizar Tudo", command=app.atualizar_todas_abas, bootstyle="success"
    ).pack(side=tk.LEFT, padx=4)
    #ttkb.Button(
    #    frame_top, text="Diagnóstico", command=app.executar_diagnostico, bootstyle="warning"
    #).pack(side=tk.LEFT, padx=4)

    # ---------------- Notebook (Abas) ----------------
    app.notebook = ttkb.Notebook(janela, bootstyle="secondary")
    app.notebook.pack(fill=tk.BOTH, expand=True, padx=8, pady=6)

    # Abas: Boletim, Registro de Ponto, Comparação e ROTÁLOG (vazia)
    _build_tab_boletim(app, app.notebook)
    _build_tab_ponto(app, app.notebook)
    _build_tab_comparacao(app, app.notebook)
    _build_tab_rotalog(app, app.notebook)  # vazia por enquanto

    # ---------------- Status Bar ----------------
    status_frame = ttkb.Frame(janela, padding=(8, 4))
    status_frame.pack(side=tk.BOTTOM, fill=tk.X)
    app.status_bar = ttkb.Label(status_frame, text="Pronto.", anchor="w")
    app.status_bar.pack(side=tk.LEFT, fill=tk.X, expand=True)

    # ---------------- Menubar (opcional) ----------------
    _criar_menubar(janela, app)

    return janela


# =============================================================================
# Abas
# =============================================================================


def _build_tab_boletim(app, notebook: ttk.Notebook) -> None:
    """
    Aba "Boletim": uma grid com os dias do período, incluindo US / Serviços / Emergências.
    """
    frame = ttkb.Frame(notebook)
    notebook.add(frame, text="Boletim")

    # --- Barra de ações da aba ---
    actions = ttkb.Frame(frame, padding=(8, 6))
    actions.pack(side=tk.TOP, fill=tk.X)

    ttkb.Button(
        actions, text="Atualizar Tela", command=app.gerar_relatorio_tela, bootstyle="success"
    ).pack(side=tk.LEFT, padx=4)
    
    ttkb.Button(
        actions,
        text="Extrair Boletim (PDFs)",
        command=lambda: app.executar_extrator_boletim(forcar_dialogo=True),
        bootstyle="info",
    ).pack(side=tk.LEFT, padx=4)

    ttkb.Button(
        actions,
        text="Exportar (Excel/PDF)",
        command=app.mostrar_dialogo_exportacao,
        bootstyle="secondary",
    ).pack(side=tk.LEFT, padx=4)

    ttkb.Button(
        actions,              
        text="Formulário geral (SobreAvisos)",
        command=app.abrir_formulario_geral,   # <-- era app.abrir_relatorio_sa
        bootstyle="info",
    ).pack(side=tk.LEFT, padx=4)



    app.label_status_boletim = ttkb.Label(actions, text="", bootstyle="secondary")
    app.label_status_boletim.pack(side=tk.RIGHT, padx=8)

    # --- Grid (Sheet) ---
    grid_frame = ttkb.Frame(frame)
    grid_frame.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)

    # Cabeçalhos esperados pelo app (mantém em app.colunas_boletim_dados)
    app.colunas_boletim_dados = [
        "DATA",
        "BOLETIM",
        "HORA NORMAL",
        "H.E.",
        "H.E.D.",
        "H.E.N.",
        "H.E.N.D.",
        "H.N.",
        "S.A.",
    ]
    headers_visuais = [
        "Data",
        "Boletim",
        "Horas\n Normais",
        "Extra\n 50%D",
        "Extra\n 100%D",
        "Extra\n 50%N",
        "Extra\n 100%N",
        "Horas\n Noturnas",
        "S.A.",
    ]

    if Sheet is None:
        app.sheet_boletim = None
        ttkb.Label(grid_frame, text="tksheet não disponível", bootstyle="danger").pack()
    else:
        app.sheet_boletim = Sheet(
            grid_frame,
            headers=headers_visuais,
            show_row_index=False,  # ⬅️ NOVO
            show_x_scrollbar=True,
            show_y_scrollbar=True,
        )
        ocultar_indices_sheet(app.sheet_boletim)  # ⬅️ NOVO
        app.sheet_boletim.pack(fill=tk.BOTH, expand=True)

        _configurar_sheet(
            app.sheet_boletim,
            headers_visuais,
            header_bg="#BAD8F3",
            first_col_width=True,
            first_col_width_px=120,
            default_col_width_px=90,
            right_align_from=2,
            header_wrap=True,
            header_align_center=True,
            header_height_lines=2,
        )


def _build_tab_ponto(app, notebook: ttk.Notebook) -> None:
    """
    Aba "Registro Ponto": grid por dia para o funcionário selecionado.
    """
    # >>> GUARDA O FRAME DA ABA NO app <<<
    app.aba_ponto = ttkb.Frame(notebook)
    notebook.add(app.aba_ponto, text="Registro Ponto")

    # --- Barra de ações da aba ---
    actions = ttkb.Frame(app.aba_ponto, padding=(8, 6))
    actions.pack(side=tk.TOP, fill=tk.X)

    ttkb.Button(
        actions, text="Atualizar Ponto", command=app.gerar_relatorio_ponto, bootstyle="success"
    ).pack(side=tk.LEFT, padx=4)
    ttkb.Button(
        actions, text="Extrair Ponto (CSVs)", command=app.executar_extrator_ponto, bootstyle="info"
    ).pack(side=tk.LEFT, padx=4)
    ttkb.Button(
        actions,
        text="Exportar Registro Ponto",
        command=app.mostrar_dialogo_exportacao_ponto,
        bootstyle="secondary",
    ).pack(side=tk.LEFT, padx=4)

    # Label de status da aba
    app.label_status_ponto = ttkb.Label(actions, text="", bootstyle="secondary")
    app.label_status_ponto.pack(side=tk.RIGHT, padx=8)

    # --- Grid (Sheet) ---
    grid_frame = ttkb.Frame(app.aba_ponto)  # <-- CORREÇÃO: usar app.aba_ponto
    grid_frame.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)

    # Cabeçalhos visuais alinhados ao layout central (HEADERS_VIZ)
    headers_ponto = ["Data"] + (
        app.HEADERS_VIZ
        if hasattr(app, "HEADERS_VIZ")
        else [
            "Horas\n Normais",
            "Horas\n Noturnas",
            "Extra\n 50%D",
            "Extra\n 100%D",
            "Extra\n 50%N",
            "Extra\n 100%N",
            "Inter\njornada",
        ]
    )

    if Sheet is None:
        app.sheet_ponto = None
        ttkb.Label(grid_frame, text="tksheet não disponível", bootstyle="danger").pack()
    else:
        app.sheet_ponto = Sheet(
            grid_frame,
            headers=headers_ponto,
            show_row_index=False,  # <- oculta no construtor (quando suportado)
            show_x_scrollbar=True,
            show_y_scrollbar=True,
        )
        ocultar_indices_sheet(app.sheet_ponto)  # <- oculta de forma robusta
        app.sheet_ponto.pack(fill=tk.BOTH, expand=True)

        _configurar_sheet(
            app.sheet_ponto,
            headers_ponto,
            header_bg="#ABE08E",  # verde claro (consistente)
            first_col_width=True,  # "Data" mais larga
            first_col_width_px=120,
            default_col_width_px=90,
            right_align_from=1,  # números à direita
            header_wrap=True,
            header_align_center=True,
            header_height_lines=2,
        )


# =============================================================================
# Aba: COMPARAÇÃO — 50% (Boletim) | 25% (Ponto) | 25% (Diferença)
# Usa grid() para ocupar 100% da área disponível em largura e altura.
# =============================================================================
def _build_tab_comparacao(app, notebook: ttk.Notebook) -> None:
    """
    Aba "Comparação": 3 grids lado a lado (Boletim | Ponto | Diferença).
    Proporção: 50% (Boletim) / 25% (Ponto) / 25% (Diferença)
    - Seleção sincronizada entre as 3 (clicar em Data seleciona a linha toda).
    """
    frame = ttkb.Frame(notebook)
    notebook.add(frame, text="Comparação")

    # ---------------- Barra de ações ----------------
    actions = ttkb.Frame(frame, padding=(8, 6))
    actions.pack(side=tk.TOP, fill=tk.X)

    ttkb.Button(
        actions,
        text="Gerar Comparação",
        command=app.gerar_relatorio_comparacao,
        bootstyle="primary",
    ).pack(side=tk.LEFT, padx=4)
    ttkb.Button(
        actions,
        text="Exportar (Opções…)",
        command=app.mostrar_dialogo_exportacao_unificada,
        bootstyle="secondary",
    ).pack(side=tk.LEFT, padx=4)
    ttkb.Button(
        actions,
        text="Relação de Nomes (CPF/PIS)",
        command=app.gerenciar_relacao_nomes,
        bootstyle="info",
    ).pack(side=tk.LEFT, padx=4)

    app.label_status_comparacao = ttkb.Label(actions, text="", bootstyle="secondary")
    app.label_status_comparacao.pack(side=tk.RIGHT, padx=8)

    # ---------------- Container de grids ----------------
    grids = ttkb.Frame(frame)
    grids.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)

    # Configura grid do container pra ocupar 100% da altura,
    # e proporções de largura 50 / 25 / 25
    grids.rowconfigure(0, weight=1)
    grids.columnconfigure(0, weight=2)  # 50%
    grids.columnconfigure(1, weight=1)  # 25%
    grids.columnconfigure(2, weight=1)  # 25%

    headers_visuais = [
        " Hora\nNormais ",
        "Horas\nNoturnas",
        "  Extra\n50%D  ",
        "  Extra\n100%D ",
        "  Extra\n50%N  ",
        "   Extra\n100%N  ",
    ]

    # --- Boletim (tem Data) ---
    bloco_bol = ttkb.Labelframe(grids, text="Boletim", padding=(8, 6))
    bloco_bol.grid(row=0, column=0, sticky="nsew", padx=(0, 6))

    headers_bol = ["Data", "Boletim"] + headers_visuais
    if Sheet is None:
        app.sheet_comp_boletim = None
        ttkb.Label(bloco_bol, text="tksheet não disponível", bootstyle="danger").pack()
    else:
        app.sheet_comp_boletim = Sheet(
            bloco_bol, headers=headers_bol, show_x_scrollbar=True, show_y_scrollbar=True
        )
        ocultar_indices_sheet(app.sheet_comp_boletim)
        app.sheet_comp_boletim.pack(fill=tk.BOTH, expand=True)

        _configurar_sheet(
            app.sheet_comp_boletim,
            headers_bol,
            header_bg="#BAD8F3",
            first_col_width=True,
            first_col_width_px=130,
            default_col_width_px=90,
            right_align_from=1,
            header_wrap=True,
            header_align_center=True,
            header_height_lines=2,
        )

    # --- Ponto ---
    bloco_pto = ttkb.Labelframe(grids, text="Ponto", padding=(8, 6))
    bloco_pto.grid(row=0, column=1, sticky="nsew", padx=6)

    headers_pto = headers_visuais[:]
    if Sheet is None:
        app.sheet_comp_ponto = None
        ttkb.Label(bloco_pto, text="tksheet não disponível", bootstyle="danger").pack()
    else:
        app.sheet_comp_ponto = Sheet(
            bloco_pto,
            headers=headers_pto,
            show_row_index=False,  # ⬅️ adiciona: esconde índices no construtor
            show_x_scrollbar=True,
            show_y_scrollbar=True,  # ⬅️ recomendo TRUE para rolagem vertical
        )
        ocultar_indices_sheet(app.sheet_comp_ponto)
        app.sheet_comp_ponto.pack(fill=tk.BOTH, expand=True)

        _configurar_sheet(
            app.sheet_comp_ponto,
            headers_pto,
            header_bg="#ABE08E",
            first_col_width=False,
            default_col_width_px=90,
            right_align_from=0,
            header_wrap=True,
            header_align_center=True,
            header_height_lines=2,
        )

    # --- Diferença ---
    bloco_dif = ttkb.Labelframe(grids, text="Diferença", padding=(8, 6))
    bloco_dif.grid(row=0, column=2, sticky="nsew", padx=(6, 0))

    headers_dif = headers_visuais[:]
    if Sheet is None:
        app.sheet_comp_diferenca = None
        ttkb.Label(bloco_dif, text="tksheet não disponível", bootstyle="danger").pack()
    else:
        app.sheet_comp_diferenca = Sheet(
            bloco_dif,
            headers=headers_dif,
            show_row_index=False,  # ⬅️ oculta índice no construtor (quando suportado)
            show_x_scrollbar=True,
            show_y_scrollbar=True,
        )
        ocultar_indices_sheet(app.sheet_comp_diferenca)
        #ocultar_indices_sheet(app.sheet_comp_boletim)
        app.sheet_comp_diferenca.pack(fill=tk.BOTH, expand=True)

        _configurar_sheet(
            app.sheet_comp_diferenca,
            headers_dif,
            header_bg="#F3C1B9",
            first_col_width=False,
            default_col_width_px=90,
            right_align_from=0,
            header_wrap=True,
            header_align_center=True,
            header_height_lines=2,
        )

    # === SINCRONIZAÇÃO DE SELEÇÃO ENTRE AS 3 GRIDS ===========================
    def _select_row_metrics(row):
        """Seleciona TODAS as 6 métricas (colunas) do row nas 3 sheets."""
        # Boletim: colunas 1..6 (pula a 'Data')
        try:
            if hasattr(app.sheet_comp_boletim, "select_cells"):
                app.sheet_comp_boletim.select_cells(row, 1, row, 6, redraw=True)
            else:
                for c in range(1, 7):
                    app.sheet_comp_boletim.select_cell(
                        row, c, redraw=True, keep_other_selections=(c > 1)
                    )
            if hasattr(app.sheet_comp_boletim, "see"):
                app.sheet_comp_boletim.see(row, 1)
        except Exception:
            pass

        # Ponto: 0..5
        try:
            if hasattr(app.sheet_comp_ponto, "select_cells"):
                app.sheet_comp_ponto.select_cells(row, 0, row, 5, redraw=True)
            else:
                for c in range(0, 6):
                    app.sheet_comp_ponto.select_cell(
                        row, c, redraw=True, keep_other_selections=(c > 0)
                    )
            if hasattr(app.sheet_comp_ponto, "see"):
                app.sheet_comp_ponto.see(row, 0)
        except Exception:
            pass

        # Diferença: 0..5
        try:
            if hasattr(app.sheet_comp_diferenca, "select_cells"):
                app.sheet_comp_diferenca.select_cells(row, 0, row, 5, redraw=True)
            else:
                for c in range(0, 6):
                    app.sheet_comp_diferenca.select_cell(
                        row, c, redraw=True, keep_other_selections=(c > 0)
                    )
            if hasattr(app.sheet_comp_diferenca, "see"):
                app.sheet_comp_diferenca.see(row, 0)
        except Exception:
            pass

            # === liga a sincronização de seleção entre as 3 sheets ===
        try:
            bind_comparison_sync(app)
        except Exception:
            pass

    def _sync_select_from(kind, row, col):
        """
        kind: 'bol' | 'pto' | 'dif'
        row/col: coordenadas da folha origem (0-based)
        Converte para o índice de 'métrica' (0..5) e seleciona a mesma nas 3.
        Caso clique na coluna 'Data' do Boletim (col == 0), seleciona a LINHA inteira.
        """
        try:
            if kind == "bol" and col == 0:
                _select_row_metrics(row)
                return

            metric = (col - 1) if kind == "bol" else col
            if metric is None or metric < 0 or metric > 5:
                return

            # Boletim (offset +1)
            try:
                app.sheet_comp_boletim.select_cell(
                    row, metric + 1, redraw=True, keep_other_selections=False
                )
                if hasattr(app.sheet_comp_boletim, "see"):
                    app.sheet_comp_boletim.see(row, metric + 1)
            except Exception:
                pass

            # Ponto
            try:
                app.sheet_comp_ponto.select_cell(
                    row, metric, redraw=True, keep_other_selections=False
                )
                if hasattr(app.sheet_comp_ponto, "see"):
                    app.sheet_comp_ponto.see(row, metric)
            except Exception:
                pass

            # Diferença
            try:
                app.sheet_comp_diferenca.select_cell(
                    row, metric, redraw=True, keep_other_selections=False
                )
                if hasattr(app.sheet_comp_diferenca, "see"):
                    app.sheet_comp_diferenca.see(row, metric)
            except Exception:
                pass
        except Exception:
            pass

    def _bind_sync(sheet, kind):
        def _on_cell_select(event=None):
            try:
                r, c = sheet.get_currently_selected()
            except Exception:
                sel = getattr(sheet, "selected_cells", None)
                if sel:
                    r, c = list(sel)[0]
                else:
                    return
            if r is None or c is None:
                return
            _sync_select_from(kind, r, c)

        try:
            sheet.extra_bindings(
                [("cell_select", _on_cell_select), ("row_select", _on_cell_select)]
            )
        except Exception:
            pass

    if Sheet is not None:
        _bind_sync(app.sheet_comp_boletim, "bol")
        _bind_sync(app.sheet_comp_ponto, "pto")
        _bind_sync(app.sheet_comp_diferenca, "dif")


def _build_tab_rotalog(app, notebook: ttk.Notebook) -> None:
    """
    Aba "ROTALOG": vazia por enquanto (conforme solicitado).
    """
    # >>> GUARDA O FRAME DA ABA NO app <<<
    app.aba_rotalog = ttkb.Frame(notebook)
    notebook.add(app.aba_rotalog, text="ROTALOG")

    vazio = ttkb.Label(
        app.aba_rotalog,
        text="Aba ROTALOG ainda não implementada.",
        bootstyle="secondary",
        anchor="center",
        justify="center",
    )
    vazio.pack(fill=tk.BOTH, expand=True, padx=12, pady=12)



# =============================================================================
# Menubar
# =============================================================================
def _criar_menubar(janela: ttkb.Window, app) -> None:
    """
    Cria a barra de menus principal com algumas ações rápidas.
    """
    menubar = tk.Menu(janela)

    # Menu Arquivo
    menu_arq = tk.Menu(menubar, tearoff=0)
    menu_arq.add_command(
        label="Exportar Relatório Atual...", command=app.mostrar_dialogo_exportacao
    )
    menu_arq.add_separator()
    menu_arq.add_command(label="Sair", command=janela.destroy)
    menubar.add_cascade(label="Arquivo", menu=menu_arq)

    # Menu Dados
    menu_dados = tk.Menu(menubar, tearoff=0)
    menu_dados.add_command(
        label="Extrair Boletim (PDFs)...",
        command=lambda: app.executar_extrator_boletim(forcar_dialogo=True),
    )
    menu_dados.add_command(label="Extrair Ponto (CSVs)...", command=app.executar_extrator_ponto)
    menu_dados.add_separator()
    menu_dados.add_command(
        label="Relação de Nomes (CPF/PIS)...", command=app.gerenciar_relacao_nomes
    )
    menubar.add_cascade(label="Dados", menu=menu_dados)

    # Menu Relatórios
    menu_rel = tk.Menu(menubar, tearoff=0)
    menu_rel.add_command(label="Atualizar Boletim", command=app.gerar_relatorio_tela)
    menu_rel.add_command(label="Atualizar Ponto", command=app.gerar_relatorio_ponto)
    menu_rel.add_command(label="Gerar Comparação", command=app.gerar_relatorio_comparacao)
    menubar.add_cascade(label="Relatórios", menu=menu_rel)

    # Menu Ferramentas
    menu_tools = tk.Menu(menubar, tearoff=0)
    menu_tools.add_command(label="Diagnóstico", command=app.executar_diagnostico)
    menubar.add_cascade(label="Ferramentas", menu=menu_tools)

    janela.config(menu=menubar)
