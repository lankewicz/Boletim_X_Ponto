import flet as ft
import pandas as pd
from core_adapter import (
    load_all, period_defaults, Periodo, list_contracts, list_employees,
    resolve_base_ponto, boletim_por_funcionario, comparacao_diaria, HEADERS_VIZ
)

def app(page: ft.Page):
    page.title = "Relatório — Nova Interface (Flet)"
    page.padding = 16
    page.scroll = ft.ScrollMode.AUTO

    df_b, df_p, df_r = load_all()
    per = period_defaults(df_b)
    contratos = list_contracts(df_b)

    dd_contrato = ft.Dropdown(label="Contrato (opcional)",
                              options=[ft.dropdown.Option("")]+[ft.dropdown.Option(c) for c in contratos],
                              value="")
    dp_ini = ft.TextField(label="Data inicial (AAAA-MM-DD)", value=str(per.inicio.date()), width=160)
    dp_fim = ft.TextField(label="Data final (AAAA-MM-DD)", value=str(per.fim.date()), width=160)
    dd_func = ft.Dropdown(label="Funcionário", options=[], width=360)

    cols = [ft.DataColumn(ft.Text("Data")), ft.DataColumn(ft.Text("Boletim"))]
    for h in HEADERS_VIZ:
        cols += [ft.DataColumn(ft.Text(f"B | {h}")),
                 ft.DataColumn(ft.Text(f"P | {h}")),
                 ft.DataColumn(ft.Text(f"D | {h}"))]
    tbl_comp = ft.DataTable(columns=cols, rows=[])

    kpi_total = ft.Text("—", weight=ft.FontWeight.BOLD)

    def _refresh_funcs():
        try:
            per_loc = Periodo(pd.to_datetime(dp_ini.value), pd.to_datetime(dp_fim.value))
        except Exception:
            return
        funcs = list_employees(df_b, per_loc, dd_contrato.value or None)
        dd_func.options = [ft.dropdown.Option(f) for f in funcs]
        if funcs:
            dd_func.value = funcs[0]
        page.update()

    def _apply(_=None):
        _refresh_funcs()
        _update_tables()

    def _update_tables():
        tbl_comp.rows = []
        if not dd_func.value:
            page.update()
            return
        per_loc = Periodo(pd.to_datetime(dp_ini.value), pd.to_datetime(dp_fim.value))
        bol = boletim_por_funcionario(df_b, dd_func.value, per_loc)
        pto, origem, ident = resolve_base_ponto(df_p, df_r, dd_func.value, per_loc)
        comp = comparacao_diaria(bol, pto)

        for _, r in comp.iterrows():
            cells = [ft.DataCell(ft.Text(str(r["Data"]))), ft.DataCell(ft.Text(str(r.get("Boletim", ""))))]
            for h in HEADERS_VIZ:
                cells += [
                    ft.DataCell(ft.Text(str(r.get(f"B | {h}", "")))),
                    ft.DataCell(ft.Text(str(r.get(f"P | {h}", "")))),
                    ft.DataCell(ft.Text(str(r.get(f"D | {h}", "")))),
                ]
            tbl_comp.rows.append(ft.DataRow(cells=cells))

        total_dif = 0.0
        if not comp.empty:
            for h in HEADERS_VIZ:
                total_dif += comp[f"D | {h}"].sum()
        kpi_total.value = f"Soma das diferenças (todas métricas): {total_dif:.2f} — Origem Ponto: {origem} {ident or ''}"
        page.update()

    page.add(
        ft.Column(
            [
                ft.Row([dd_contrato, dp_ini, dp_fim, ft.ElevatedButton("Aplicar filtros", on_click=_apply)], spacing=12),
                ft.Row([dd_func, ft.TextButton("Atualizar funcionário", on_click=lambda _: (_refresh_funcs(), _update_tables()))], spacing=12),
                ft.Divider(),
                ft.Text("Comparação diária", weight=ft.FontWeight.BOLD),
                tbl_comp,
                ft.Divider(),
                kpi_total,
            ],
            spacing=12,
        )
    )

    _apply()

if __name__ == "__main__":
    ft.app(target=app)
