# utils/aggregadores.py
import pandas as pd

COL_CHAVE = ['Funcionario', 'Registro', 'Contrato']
COL_NUMS  = ['H.N.', 'H.E.', 'H.E.D.', 'H.E.N.', 'H.E.N.D.', 'S.A.']  # ajuste aos seus nomes finais

def comparar_boletim_ponto(df_boletim: pd.DataFrame, df_ponto: pd.DataFrame) -> pd.DataFrame:
    # normaliza nomes chave (se necessário)
    for c in COL_CHAVE + ['Data']:
        if c not in df_boletim.columns: raise KeyError(f"Boletim sem coluna {c}")
        if c not in df_ponto.columns:   raise KeyError(f"Ponto sem coluna {c}")

    # agrega por mês/funcionário/registro/contrato
    def _sumarizar(df):
        g = (df
             .assign(Mes=lambda d: pd.to_datetime(d['Data']).dt.to_period('M').astype(str))
             .groupby(COL_CHAVE + ['Mes'], as_index=False)[COL_NUMS].sum())
        return g

    b = _sumarizar(df_boletim).rename(columns={c: f'Boletim_{c}' for c in COL_NUMS})
    p = _sumarizar(df_ponto).rename(columns={c: f'Ponto_{c}'   for c in COL_NUMS})

    base = (b.merge(p, on=COL_CHAVE + ['Mes'], how='outer')
              .fillna(0.0))

    # difs
    for c in COL_NUMS:
        base[f'Dif_{c}'] = base[f'Ponto_{c}'] - base[f'Boletim_{c}']

    # ordenação por nome do funcionário
    base = base.sort_values(['Funcionario','Mes','Contrato','Registro'], kind='stable')
    return base
