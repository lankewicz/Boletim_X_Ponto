import io
import pandas as pd
import streamlit as st
from core_adapter import (
    load_all, period_defaults, Periodo, list_contracts, list_employees,
    resolve_base_ponto, boletim_por_funcionario, comparacao_diaria, HEADERS_VIZ
)

st.set_page_config(page_title="Relatório — Nova Interface (Streamlit)", layout="wide")

df_b, df_p, df_r = load_all()
per = period_defaults(df_b)

st.sidebar.header("Filtros")
contratos = [""] + list_contracts(df_b)
contrato = st.sidebar.selectbox("Contrato (opcional)", contratos, index=0)
ini = st.sidebar.date_input("Data inicial", per.inicio.date())
fim = st.sidebar.date_input("Data final", per.fim.date())
per_loc = Periodo(pd.to_datetime(ini), pd.to_datetime(fim))

funcs = list_employees(df_b, per_loc, contrato or None)
func = st.sidebar.selectbox("Funcionário", funcs, index=0 if funcs else None)
st.sidebar.caption("Lendo de ./data/*.parquet")

st.title("Comparação diária — Boletim × Ponto")
if not func:
    st.info("Nenhum funcionário encontrado no período/contrato selecionado.")
    st.stop()

bol = boletim_por_funcionario(df_b, func, per_loc)
pto, origem, ident = resolve_base_ponto(df_p, df_r, func, per_loc)
comp = comparacao_diaria(bol, pto)

c1, c2, c3 = st.columns(3)
c1.metric("Dias no período", len(comp))
soma_b = sum(comp.get(f"B | {h}", 0).sum() if f"B | {h}" in comp else 0 for h in HEADERS_VIZ) if not comp.empty else 0
soma_p = sum(comp.get(f"P | {h}", 0).sum() if f"P | {h}" in comp else 0 for h in HEADERS_VIZ) if not comp.empty else 0
c2.metric("Total Boletim (todas)", f"{soma_b:.2f}")
c3.metric("Total Ponto (todas)", f"{soma_p:.2f}")
st.caption(f"Origem Ponto: **{origem}** {ident or ''}")

st.dataframe(comp, use_container_width=True)

st.subheader("Exportar")
def _to_xlsx_bytes(df: pd.DataFrame) -> bytes:
    bio = io.BytesIO()
    with pd.ExcelWriter(bio, engine="xlsxwriter") as xw:
        df.to_excel(xw, index=False, sheet_name="Comparacao")
    return bio.getvalue()

col_a, col_b = st.columns(2)
with col_a:
    st.download_button("Baixar comparação (CSV)", data=comp.to_csv(index=False).encode("utf-8"),
                       file_name=f"comparacao_{func}.csv", mime="text/csv")
with col_b:
    st.download_button("Baixar comparação (Excel)", data=_to_xlsx_bytes(comp),
                       file_name=f"comparacao_{func}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
