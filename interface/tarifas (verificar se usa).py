

# ======================================================================
# Loader de tarifas COPEL em JSON
# - Arquivo: <raiz do projeto>/data/tarifas_copel.json
# - As chaves DEVEM bater com os headers visuais usados na Diferença:
#   "Horas Normais", "Horas Noturnas", "Extra 50%D", "Extra 100%D",
#   "Extra 50%N", "Extra 100%N"
# ======================================================================
def carregar_tarifas_copel():
    """
    Lê o arquivo JSON com as tarifas da COPEL (R$/hora) e devolve um dicionário
    {coluna_visual: tarifa_float}. Dispara FileNotFoundError se não existir.
    """
    base_dir = Path(__file__).resolve().parent.parent  # .../interface -> raiz
    json_path = base_dir / "data" / "tarifas_copel.json"
    if not json_path.exists():
        raise FileNotFoundError(f"Arquivo de tarifas não encontrado: {json_path}")
    with open(json_path, "r", encoding="utf-8") as f:
        data = json.load(f)
    return {str(k): float(v) for k, v in data.items()}


def _normalize_name(name: str) -> str:
    """Normaliza nomes para comparação (maiúsculas, sem acento, espaços colapsados)."""
    if not isinstance(name, str):
        return ""
    s = name.strip().upper()
    s = "".join(ch for ch in unicodedata.normalize("NFKD", s) if not unicodedata.combining(ch))
    s = re.sub(r"\s+", " ", s)
    return s

def _safe_read_parquet(path: Path, *, required_cols=None, date_cols=None):
    """Lê parquet com checagens; retorna (df, err). err != None indica problema."""
    if not path.exists():
        return None, f"Arquivo não encontrado: {path}"
    try:
        df = pd.read_parquet(path)
    except Exception as e:
        return None, f"Erro lendo {path}: {e}"
    if date_cols:
        for col in date_cols:
            if col in df.columns:
                df[col] = pd.to_datetime(df[col], errors="coerce")
    if required_cols:
        missing = [c for c in required_cols if c not in df.columns]
        if missing:
            return None, f"Colunas ausentes em {path}: {missing}"
    if df.empty:
        return None, f"Nenhum dado em {path}"
    return df, None

def _filter_month(df: pd.DataFrame, date_col: str, ano: int, mes: int) -> pd.DataFrame:
    """Filtra por mês/ano usando uma coluna de data já convertida para datetime."""
    if date_col not in df.columns:
        return df.iloc[0:0].copy()
    return df[(df[date_col].dt.year == ano) & (df[date_col].dt.month == mes)].copy()

def _filter_ponto_by_boletim_names(df_p: pd.DataFrame, df_b: pd.DataFrame,
                                   col_p: str = "Nome", col_b: str = "Funcionário") -> pd.DataFrame:
    """Mantém no Ponto apenas os funcionários que aparecem no Boletim (normalização inclusa)."""
    if col_b not in df_b.columns or col_p not in df_p.columns:
        return df_p.iloc[0:0].copy()
    nomes_b = set(df_b[col_b].dropna().map(_normalize_name))
    if not nomes_b:
        return df_p.iloc[0:0].copy()
    df = df_p[df_p[col_p].notna()].copy()
    df["__nome_norm__"] = df[col_p].map(_normalize_name)
    df = df[df["__nome_norm__"].isin(nomes_b)]
    return df.drop(columns="__nome_norm__", errors="ignore")
