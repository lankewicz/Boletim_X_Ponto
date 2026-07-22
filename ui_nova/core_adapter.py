from __future__ import annotations
from dataclasses import dataclass
from pathlib import Path
from typing import Optional
import pandas as pd
from difflib import get_close_matches

DATA_DIR = Path(__file__).resolve().parent.parent / "data"
ARQ_BOL = DATA_DIR / "RegistroBoletim.parquet"
ARQ_PTO = DATA_DIR / "RegistroPonto.parquet"
ARQ_REL = DATA_DIR / "relacao_nomes.parquet"

HEADERS_VIZ = [
    "Horas Normais",
    "Horas Noturnas",
    "Extra 50%D",
    "Extra 100%D",
    "Extra 50%N",
    "Extra 100%N",
]
# mapeia colunas do ponto para nomes “viz”
MAP_PTO = {
    "Horas Normais": "Total Normais",
    "Horas Noturnas": "Total Noturno",
    "Extra 50%D": "Extra 50%D",
    "Extra 100%D": "Extra 100%D",
    "Extra 50%N": "Extra 50%N",
    "Extra 100%N": "Extra 100%N",
}
import pandas as pd
import re
from datetime import datetime

def _coerce_header_row(df: pd.DataFrame) -> pd.DataFrame:
    """
    Se as colunas forem RangeIndex (0..N), tenta usar a primeira linha como header.
    Ex.: DataFrame lido sem header=0.
    """
    if isinstance(df.columns, pd.RangeIndex):
        # Heurística: primeira linha parece conter strings (nomes)?
        first_row = df.iloc[0]
        if first_row.map(lambda x: isinstance(x, str)).mean() > 0.5:
            df = df.rename(columns=first_row).iloc[1:].reset_index(drop=True)
    return df

def _normalize_colname(name: str) -> str:
    """
    Normaliza nome de coluna: minúsculas, sem acentos, sem espaços duplos,
    substitui separadores por '_'.
    """
    import unicodedata
    s = unicodedata.normalize("NFKD", name)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    s = s.strip().lower()
    s = re.sub(r"[^\w]+", "_", s)  # troca espaços, barras, etc. por _
    s = re.sub(r"_+", "_", s).strip("_")
    return s

def _normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [_normalize_colname(str(c)) for c in df.columns]
    return df

def _pick_first(df: pd.DataFrame, candidates: list[str]) -> str | None:
    """
    Retorna o primeiro nome de coluna existente dentre os candidatos (já normalizados).
    """
    cols = set(df.columns)
    for c in candidates:
        if c in cols:
            return c
    return None

def _ensure_date_column(df: pd.DataFrame) -> pd.DataFrame:
    """
    Garante a existência de uma coluna 'Data' (exata, com D maiúsculo) a partir de
    variações comuns: data, data_hora, dia, date, competencia, etc.
    Normaliza para datetime (somente a data, sem hora).
    """
    df = _coerce_header_row(df)
    df = _normalize_columns(df)

    # Possíveis candidatos (normalizados):
    candidates = [
        "data", "dia", "date", "data_hora", "datahora", "dt", "competencia"
    ]
    col = _pick_first(df, candidates)

    if col is None:
        # Tenta detectar uma coluna com muitos valores que “parecem” datas
        for c in df.columns:
            try:
                parsed = pd.to_datetime(df[c], errors="coerce", dayfirst=True)
                if parsed.notna().mean() > 0.6:  # heurística: 60% parecem datas
                    col = c
                    break
            except Exception:
                pass

    if col is None:
        # Nada encontrado → erro claro e guiado
        raise KeyError(
            "Não encontrei coluna de data. Esperava algo como "
            "'Data', 'Dia', 'Data/Hora', 'Date' ou semelhante. "
            "Verifique se a primeira linha é cabeçalho e se o arquivo não perdeu o header."
        )

    # Cria a coluna 'Data' (title case) a partir do candidato
    date_series = pd.to_datetime(df[col], errors="coerce", dayfirst=True).dt.normalize()
    if date_series.isna().all():
        raise ValueError(
            f"A coluna '{col}' foi identificada como data, mas não consegui converter valores. "
            "Confira o formato (ex.: dia/mês/ano) ou se a coluna tem apenas datas."
        )

    df = df.assign(Data=date_series)  # adiciona 'Data' com D maiúsculo para o restante do código
    return df

def _ensure_data_column_only(df: pd.DataFrame) -> pd.DataFrame:
    """
    Garante a existência de uma coluna 'Data' (D maiúsculo) SEM normalizar outros nomes de colunas.
    Prioriza colunas candidatas comuns; se não houver, tenta detectar por cobertura de conversão datetime.
    """
    if df.empty:
        return df

    df = df.copy()

    # Caso já exista 'Data', apenas normaliza dtype
    if "Data" in df.columns:
        df["Data"] = pd.to_datetime(df["Data"], errors="coerce", dayfirst=True).dt.normalize()
        return df

    # Candidatos comuns (mantendo caixa/acentos mais usuais)
    candidates = [
        "DATA", "data", "Data/Hora", "DATA/HORA", "data_hora", "DataHora",
        "Dia", "DIA", "date", "dt", "competencia", "Competência",
    ]
    for c in candidates:
        if c in df.columns:
            s = pd.to_datetime(df[c], errors="coerce", dayfirst=True).dt.normalize()
            if s.notna().any():
                df["Data"] = s
                return df

    # Detecção heurística por cobertura de datetime
    best, best_score = None, 0.0
    for c in df.columns:
        try:
            s = pd.to_datetime(df[c], errors="coerce", dayfirst=True)
            score = float(s.notna().mean())
            if score > best_score:
                best, best_score = c, score
        except Exception:
            pass

    if best is not None and best_score >= 0.6:
        df["Data"] = pd.to_datetime(df[best], errors="coerce", dayfirst=True).dt.normalize()
        return df

    raise KeyError(
        "Não encontrei coluna de data em df_ponto (esperava 'Data', 'DATA', 'Data/Hora', etc.)."
    )



@dataclass
class Periodo:
    inicio: pd.Timestamp
    fim: pd.Timestamp

def _read_parquet(path: Path) -> pd.DataFrame:
    if not path.exists():
        return pd.DataFrame()
    return pd.read_parquet(path)

def load_all() -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    df_b = _read_parquet(ARQ_BOL)
    df_p = _read_parquet(ARQ_PTO)
    df_r = _read_parquet(ARQ_REL)
    if not df_b.empty:
        df_b["DATA"] = pd.to_datetime(df_b["DATA"], errors="coerce")
        df_b = df_b.dropna(subset=["DATA"]).copy()
    if not df_p.empty:
        try:
            if "Data" not in df_p.columns:
                df_p = _ensure_date_column(df_p)  # cria 'Data' a partir de variações
            else:
                df_p["Data"] = pd.to_datetime(df_p["Data"], errors="coerce")
            df_p = df_p.dropna(subset=["Data"]).copy()
        except Exception:
            # fallback suave: tenta forçar criação de 'Data'
            df_p = _ensure_date_column(df_p)

    return df_b, df_p, df_r

def period_defaults(df_b: pd.DataFrame) -> Periodo:
    if df_b.empty:
        hoje = pd.Timestamp.today().normalize()
        return Periodo(hoje.replace(day=1), hoje)
    mi, ma = df_b["DATA"].min().normalize(), df_b["DATA"].max().normalize()
    return Periodo(ma.replace(day=1), ma)

def list_contracts(df_b: pd.DataFrame) -> list[str]:
    if df_b.empty or "Contrato" not in df_b.columns:
        return []
    return sorted(df_b["Contrato"].dropna().astype(str).unique().tolist())

def list_employees(df_b: pd.DataFrame, per: Periodo, contrato: Optional[str] = None) -> list[str]:
    if df_b.empty:
        return []
    sel = df_b[(df_b["DATA"].dt.date >= per.inicio.date()) & (df_b["DATA"].dt.date <= per.fim.date())]
    if contrato and "Contrato" in sel.columns:
        sel = sel[sel["Contrato"].astype(str) == str(contrato)]
    return sorted(sel["Funcionário"].dropna().astype(str).unique().tolist())

def resolve_base_ponto(
    df_p: pd.DataFrame,
    df_r: pd.DataFrame,
    nome_boletim: str,
    per: Periodo,
    corte: float = 0.75,
) -> tuple[pd.DataFrame, str, Optional[str]]:
    """CPF → PIS → Nome mapeado → similaridade. Retorna (df_ponto, origem, identificador)."""
    if df_p.empty:
        return pd.DataFrame(), "SEM_PONTO", None
    if "Data" not in df_p.columns:
        df_p = _ensure_date_column(df_p)

    base = df_p[(df_p["Data"].dt.date >= per.inicio.date()) & (df_p["Data"].dt.date <= per.fim.date())].copy()

    # Relação explícita
    if df_r is not None and not df_r.empty:
        rel = df_r[df_r["Nome_Boletim"].astype(str) == str(nome_boletim)].copy()
        if not rel.empty:
            r0 = rel.iloc[0]
            cpf = str(r0.get("CPF_Ponto") or "").strip()
            pis = str(r0.get("PIS_Ponto") or "").strip()
            nome_map = str(r0.get("Nome_Ponto_Mapeado") or "").strip()
            if cpf and "CPF" in base.columns and not base[base["CPF"].astype(str) == cpf].empty:
                return base[base["CPF"].astype(str) == cpf], "CPF", cpf
            if pis and "PIS" in base.columns and not base[base["PIS"].astype(str) == pis].empty:
                return base[base["PIS"].astype(str) == pis], "PIS", pis
            if nome_map and "Nome" in base.columns and not base[base["Nome"].astype(str) == nome_map].empty:
                return base[base["Nome"].astype(str) == nome_map], "NOME_MAPEADO", nome_map

    # Similaridade de nome
    if "Nome" in base.columns:
        nomes = base["Nome"].dropna().astype(str).unique().tolist()
        match = get_close_matches(str(nome_boletim).upper(), [n.upper() for n in nomes], n=1, cutoff=corte)
        if match:
            alvo = next((n for n in nomes if n.upper() == match[0]), None)
            if alvo:
                return base[base["Nome"].astype(str) == alvo], "SIMILARIDADE", alvo

    return pd.DataFrame(), "SEM_PONTO", None

def boletim_por_funcionario(df_b: pd.DataFrame, nome: str, per: Periodo) -> pd.DataFrame:
    if df_b.empty:
        return pd.DataFrame()
    sel = (
        (df_b["Funcionário"].astype(str) == str(nome))
        & (df_b["DATA"].dt.date >= per.inicio.date())
        & (df_b["DATA"].dt.date <= per.fim.date())
    )
    return df_b.loc[sel].copy().sort_values("DATA")

def comparacao_diaria(df_b_func: pd.DataFrame, df_p_res: pd.DataFrame) -> pd.DataFrame:
    """Retorna colunas: Data, Boletim, e tripletas B|, P|, D| para cada métrica."""
    import pandas as pd

    if df_b_func.empty and df_p_res.empty:
        return pd.DataFrame()

    # --- Garantias para o DF do ponto ---
    # 1) Se vier vazio, cria DF com coluna 'Data' vazia para evitar KeyError
    if df_p_res is None or df_p_res.empty:
        df_p_res = pd.DataFrame(columns=["Data"])

    # 2) Se não tiver 'Data', tenta garantir (variações, heurística)
    if "Data" not in df_p_res.columns:
        try:
            df_p_res = _ensure_data_column_only(df_p_res)
        except Exception:
            # Se mesmo assim não tiver, cria coluna vazia
            df_p_res = df_p_res.assign(Data=pd.to_datetime(pd.Series([], dtype="datetime64[ns]")))

    # 3) Normaliza uma vez só (coluna auxiliar) e usa no filtro
    df_p_res = df_p_res.copy()
    df_p_res["_DataNorm"] = pd.to_datetime(df_p_res["Data"], errors="coerce").dt.normalize()

    # --- Conjunto de datas ---
    datas_b = pd.to_datetime(df_b_func["DATA"]).dt.normalize().unique() if not df_b_func.empty else []
    datas_p = df_p_res["_DataNorm"].dropna().unique() if not df_p_res.empty else []
    todas = sorted(set(datas_b) | set(datas_p))

    MAP_PTO = {
        "Horas Normais": "Total Normais",
        "Horas Noturnas": "Total Noturno",
        "Extra 50%D": "Extra 50%D",
        "Extra 100%D": "Extra 100%D",
        "Extra 50%N": "Extra 50%N",
        "Extra 100%N": "Extra 100%N",
    }

    linhas = []
    for d in todas:
        row = {"Data": pd.to_datetime(d).date()}

        bdi = df_b_func[pd.to_datetime(df_b_func["DATA"]).dt.normalize() == d] if not df_b_func.empty else df_b_func
        row["Boletim"] = str(bdi["BOLETIM"].iloc[0]) if ("BOLETIM" in bdi.columns and not bdi.empty) else ""

        # Se o DF do ponto estiver vazio, pdi será vazio e vp = 0.0
        pdi = df_p_res[df_p_res["_DataNorm"] == d] if ("_DataNorm" in df_p_res.columns and not df_p_res.empty) else df_p_res

        for h in HEADERS_VIZ:
            vb = bdi[h].sum() if h in bdi.columns else 0.0
            col_p = MAP_PTO.get(h, h)
            vp = pdi[col_p].sum() if col_p in pdi.columns else 0.0
            row[f"B | {h}"] = round(float(vb or 0), 2)
            row[f"P | {h}"] = round(float(vp or 0), 2)
            row[f"D | {h}"] = round(float((vb or 0) - (vp or 0)), 2)
        linhas.append(row)

    return pd.DataFrame(linhas)

