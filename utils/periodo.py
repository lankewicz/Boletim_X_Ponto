# utils/periodo.py
from datetime import datetime as dt, date, timedelta
import calendar

def ler_datas_da_ui(cal_ini, cal_fim):
    """
    Lê as datas dos widgets DateEntry (preferindo get_date()).
    Se não houver get_date, tenta entry.get() com '%d/%m/%Y'.
    Retorna (data_ini: date, data_fim: date).
    """
    def _coerce(x):
        if hasattr(x, "date"):
            return x.date()
        return x

    try:
        di = _coerce(cal_ini.get_date())
        df = _coerce(cal_fim.get_date())
        return di, df
    except Exception:
        pass

    di = dt.strptime(cal_ini.entry.get().strip(), "%d/%m/%Y").date()
    df = dt.strptime(cal_fim.entry.get().strip(), "%d/%m/%Y").date()
    return di, df

def mes_inicio_fim(data_ini: date, data_fim: date):
    """1º dia do mês de data_ini até último dia do mês de data_fim."""
    ini = date(data_ini.year, data_ini.month, 1)
    last = calendar.monthrange(data_fim.year, data_fim.month)[1]
    fim = date(data_fim.year, data_fim.month, last)
    return ini, fim

def iter_dias(d1: date, d2: date):
    cur = d1
    while cur <= d2:
        yield cur
        cur += timedelta(days=1)

def datas_formatadas_do_mes(data_ini: date, data_fim: date):
    """Lista ['dd/mm/aaaa', ...] cobrindo do 1º dia (ini) ao último (fim)."""
    ini, fim = mes_inicio_fim(data_ini, data_fim)
    return [d.strftime("%d/%m/%Y") for d in iter_dias(ini, fim)]
