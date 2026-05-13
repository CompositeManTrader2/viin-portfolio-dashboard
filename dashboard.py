"""
Dashboard de portafolios (VIIN000000000001 / 3 / 6)
Lee los archivos LayOut*.xlsm en ./Layouts/, normaliza Movimientos y Posicion,
y muestra desempeno mensual y diario por portafolio.

Ejecucion:
    pip install streamlit pandas openpyxl plotly numpy
    streamlit run dashboard.py
"""

from __future__ import annotations

import io
import re
import unicodedata
from pathlib import Path
from datetime import datetime

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import streamlit as st
import yfinance as yf

# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------
LAYOUTS_DIR = Path(__file__).parent / "Layouts"
PORTFOLIOS = ["VIIN000000000001", "VIIN000000000003", "VIIN000000000006"]
TOTAL_LABELS = {
    "valor del portafolio",
    "ef. disponible",
    "ef. retenido",
    "sdo. disp. para inv.*",
    "sdo. disp. para inv.",
    "sdo. pend. mc",
    "mdo.dinero",
}

st.set_page_config(page_title="Portafolios VIIN", layout="wide")

# ---------------------------------------------------------------------------
# Paleta y helpers de visualizacion
# ---------------------------------------------------------------------------
PORT_COLORS = {
    "VIIN000000000001": "#1f4e79",  # azul oscuro
    "VIIN000000000003": "#d97706",  # naranja
    "VIIN000000000006": "#059669",  # verde
}
PORT_LABEL = {
    "VIIN000000000001": "VIIN ...001",
    "VIIN000000000003": "VIIN ...003",
    "VIIN000000000006": "VIIN ...006",
}
COMP_COLORS = {
    "Equities": "#2563eb",
    "Renta fija / Reporto": "#7c3aed",
    "Efectivo": "#10b981",
    "Ajuste": "#94a3b8",
}
GREEN = "#16a34a"
RED = "#dc2626"
GRID = "#e5e7eb"

PLOTLY_BASE_LAYOUT = dict(
    plot_bgcolor="white",
    paper_bgcolor="white",
    font=dict(family="-apple-system, system-ui, sans-serif", size=12, color="#1f2937"),
    margin=dict(l=60, r=20, t=60, b=40),
    hoverlabel=dict(bgcolor="white", font_size=12, bordercolor="#e5e7eb"),
)


def fmt_money(v: float, compact: bool = True) -> str:
    """$1.23M / $456K / $789"""
    if pd.isna(v):
        return "-"
    av = abs(v)
    sign = "-" if v < 0 else ""
    if compact and av >= 1e6:
        return f"{sign}${av/1e6:,.2f}M"
    if compact and av >= 1e3:
        return f"{sign}${av/1e3:,.1f}K"
    return f"{sign}${av:,.2f}"


def style_axes(fig: go.Figure, money_y: bool = False, pct_y: bool = False) -> go.Figure:
    fig.update_xaxes(
        gridcolor=GRID, zerolinecolor=GRID, showline=True,
        linecolor="#9ca3af", linewidth=1, ticks="outside", tickcolor="#9ca3af",
    )
    yfmt = ",.0f"
    if money_y:
        # plotly-friendly money format using SI suffix
        yfmt = "$,.2s"
    if pct_y:
        yfmt = ".2%"
    fig.update_yaxes(
        gridcolor=GRID, zerolinecolor="#9ca3af", showline=True,
        linecolor="#9ca3af", linewidth=1, ticks="outside", tickcolor="#9ca3af",
        tickformat=yfmt,
    )
    return fig

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
def _norm(s) -> str:
    if s is None:
        return ""
    s = str(s)
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("ascii")
    return s.strip().lower()


def _find_sheet(xl: pd.ExcelFile, target: str) -> str | None:
    target_n = _norm(target)
    for name in xl.sheet_names:
        if _norm(name) == target_n:
            return name
    return None


def _file_date(path: Path) -> datetime | None:
    m = re.search(r"LayOut(\d{2})(\d{2})(\d{4})", path.name)
    if not m:
        return None
    d, mo, y = m.groups()
    return datetime(int(y), int(mo), int(d))


def _fit_columns(df: pd.DataFrame, cols: list[str]) -> pd.DataFrame:
    """Ajusta el DataFrame para tener exactamente las columnas esperadas.
    Si faltan columnas las rellena con NaN; si sobran, las descarta."""
    n_have = df.shape[1]
    n_want = len(cols)
    if n_have < n_want:
        for i in range(n_have, n_want):
            df[i] = np.nan
    elif n_have > n_want:
        df = df.iloc[:, :n_want]
    df = df.copy()
    df.columns = cols
    return df


# ---------------------------------------------------------------------------
# Loaders
# ---------------------------------------------------------------------------
@st.cache_data(show_spinner=False)
def load_movimientos(files: list[str]) -> pd.DataFrame:
    cols = [
        "tp", "subcuenta", "fecha_op", "fecha_liq", "folio", "concepto",
        "emisora", "serie", "plazo", "tasa_premio", "precio", "titulos",
        "monto_bruto", "pct_comision", "iva_isr", "signo", "monto_neto", "saldo",
    ]
    dfs = []
    for f in files:
        path = Path(f)
        file_date = _file_date(path)
        xl = pd.ExcelFile(path, engine="openpyxl")
        sn = _find_sheet(xl, "Movimientos")
        if not sn:
            continue
        raw = pd.read_excel(xl, sheet_name=sn, header=None, engine="openpyxl")
        # data starts on row index 8 (row 9 in excel)
        body = _fit_columns(raw.iloc[8:].copy(), cols)
        body = body.dropna(subset=["subcuenta"])
        body["subcuenta"] = body["subcuenta"].astype(str).str.strip()
        body = body[body["subcuenta"].isin(PORTFOLIOS)].copy()
        body["fecha_op"] = pd.to_datetime(body["fecha_op"], errors="coerce")
        body["fecha_liq"] = pd.to_datetime(body["fecha_liq"], errors="coerce")
        body["concepto"] = body["concepto"].astype(str).str.strip()
        body["signo"] = body["signo"].astype(str).str.strip()
        for c in ["tasa_premio", "precio", "titulos", "monto_bruto",
                  "pct_comision", "iva_isr", "monto_neto", "saldo"]:
            body[c] = pd.to_numeric(body[c], errors="coerce")
        body["flujo_efectivo"] = np.where(
            body["signo"] == "+", body["monto_neto"],
            np.where(body["signo"] == "-", -body["monto_neto"], 0.0)
        )
        body["archivo_fecha"] = file_date
        dfs.append(body)
    if not dfs:
        return pd.DataFrame()
    out = pd.concat(dfs, ignore_index=True)
    return out


def _parse_posicion_sheet(
    xl: pd.ExcelFile, sheet_name: str, file_date: datetime,
    cols: list[str], use_header_fecha: bool = False,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Parsea una hoja de Posicion (actual o 'del mes anterior') y devuelve
    (posiciones, totales). Si use_header_fecha=True, lee la fecha del header
    de la hoja (celda L4 / col 11) en lugar de usar file_date — necesario
    para 'Posicion del mes anterior'."""
    raw = pd.read_excel(xl, sheet_name=sheet_name, header=None, engine="openpyxl")

    actual_date = file_date
    if use_header_fecha:
        try:
            hdr_date = raw.iloc[3, 11]  # fila 4, col 12 (Fecha:)
            if pd.notna(hdr_date):
                actual_date = pd.to_datetime(hdr_date).to_pydatetime()
        except Exception:
            pass

    body = _fit_columns(raw.iloc[8:].copy(), cols)
    body["emisora"] = body["emisora"].astype(str).str.strip()
    body["subcuenta"] = body["subcuenta"].astype(str).str.strip()

    for c in ["titulos", "precio", "importe_bruto", "precio_neto",
              "importe_neto", "precio_mercado", "valor_mercado_neto",
              "plus_minus_int", "plus_minus_pct", "pct_cartera",
              "cupon", "plazo", "tasa", "dias_x_ven"]:
        body[c] = pd.to_numeric(body[c], errors="coerce")

    is_total = body["emisora"].apply(_norm).isin(TOTAL_LABELS) & (
        (body["subcuenta"] == "") | (body["subcuenta"].isna()) |
        (body["subcuenta"].str.lower() == "nan")
    )
    tot = body[is_total].copy()
    tot["fecha"] = actual_date

    pos = body[~is_total].copy()
    pos = pos.dropna(subset=["subcuenta"])
    pos = pos[pos["subcuenta"].isin(PORTFOLIOS)]
    pos["fecha"] = actual_date
    pos["estrategia"] = pos["estrategia"].fillna("SIN ESTRATEGIA").astype(str).str.strip()

    return pos, tot[["fecha", "emisora", "valor_mercado_neto"]]


COBRO_DIR = Path(__file__).parent / "Cobro"


@st.cache_data(show_spinner=False)
def load_cobro_diario() -> pd.DataFrame:
    """Carga los valores diarios oficiales del cobro de comisiones (back-office).

    El archivo `Cobro/cobro_diario_*.xlsx` tiene una hoja por mes (ENE25, FEB25...)
    con valores diarios consolidados por portafolio que el cliente ve. Estos
    valores son la mejor referencia para la calibracion diaria porque incluyen
    devengos y ajustes que el snapshot mensual de Posicion a veces no recoge.

    Estructura por hoja:
      - Filas 8 en adelante con: contrato (col A), fecha (col B), monto total
        (col C), VIIN1 (col J), VIIN3 (col K), VIIN6 (col L).

    Devuelve DataFrame long: (fecha, subcuenta, valor_diario)
    """
    if not COBRO_DIR.exists():
        return pd.DataFrame()
    rows = []
    for path in sorted(COBRO_DIR.glob("*.xlsx")):
        try:
            xl = pd.ExcelFile(path, engine="openpyxl")
        except Exception:
            continue
        for sn in xl.sheet_names:
            # Solo hojas de meses (3 letras + 2 digitos): ENE25, FEB25, etc.
            if not (len(sn) >= 5 and sn[-2:].isdigit()):
                continue
            try:
                raw = pd.read_excel(xl, sheet_name=sn, header=None, engine="openpyxl")
            except Exception:
                continue
            # Datos diarios: fila 8 en adelante, columnas B (fecha), J, K, L
            for i in range(7, min(50, raw.shape[0])):
                fecha_val = raw.iloc[i, 1] if raw.shape[1] > 1 else None
                v_total = raw.iloc[i, 2] if raw.shape[1] > 2 else None
                v1 = raw.iloc[i, 9] if raw.shape[1] > 9 else None
                v3 = raw.iloc[i, 10] if raw.shape[1] > 10 else None
                v6 = raw.iloc[i, 11] if raw.shape[1] > 11 else None
                if not isinstance(fecha_val, (pd.Timestamp, datetime)):
                    continue
                fecha_ts = pd.Timestamp(fecha_val)

                # Validar coherencia: la suma de los 3 portafolios debe matchear
                # con el total. Si no, el cobro tiene un error de captura (caso
                # observado en DIC25 donde J=K=L=total y la suma es 3 veces el
                # total). En ese caso descartamos la fila y dejamos que la
                # calibracion caiga al snapshot oficial de Posicion.
                vals = [v for v in (v1, v3, v6) if pd.notna(v)]
                if v_total is not None and pd.notna(v_total) and vals:
                    suma_ind = sum(vals)
                    ratio = suma_ind / float(v_total) if float(v_total) != 0 else 0
                    if not (0.98 <= ratio <= 1.02):
                        continue  # datos rotos, ignorar este dia
                if pd.notna(v1):
                    rows.append({"fecha": fecha_ts, "subcuenta": "VIIN000000000001",
                                  "valor_diario": float(v1)})
                if pd.notna(v3):
                    rows.append({"fecha": fecha_ts, "subcuenta": "VIIN000000000003",
                                  "valor_diario": float(v3)})
                if pd.notna(v6):
                    rows.append({"fecha": fecha_ts, "subcuenta": "VIIN000000000006",
                                  "valor_diario": float(v6)})
    if not rows:
        return pd.DataFrame()
    df = pd.DataFrame(rows).drop_duplicates(
        subset=["fecha", "subcuenta"], keep="last"
    ).sort_values(["subcuenta", "fecha"]).reset_index(drop=True)
    return df


@st.cache_data(show_spinner=False)
def load_posiciones(files: list[str]) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Devuelve (posiciones, totales).

    Carga la hoja 'Posicion' de cada archivo. ADICIONALMENTE, del PRIMER
    archivo (cronologicamente), tambien carga la hoja 'Posicion del mes
    anterior' que contiene el snapshot al cierre del mes previo (anchor
    inicial para reconstruccion diaria del periodo).
    """
    cols = [
        "tp", "subcuenta", "emisora", "serie", "cupon", "plazo", "tasa",
        "dias_x_ven", "titulos", "precio", "importe_bruto", "precio_neto",
        "importe_neto", "precio_mercado", "valor_mercado_neto",
        "plus_minus_int", "plus_minus_pct", "pct_cartera", "estrategia",
    ]
    pos_dfs, tot_dfs = [], []
    files_sorted = sorted(files, key=lambda f: _file_date(Path(f)) or datetime.min)

    for idx, f in enumerate(files_sorted):
        path = Path(f)
        file_date = _file_date(path)
        xl = pd.ExcelFile(path, engine="openpyxl")

        # Hoja Posicion (actual)
        sn = _find_sheet(xl, "Posicion")
        if sn:
            pos, tot = _parse_posicion_sheet(xl, sn, file_date, cols, use_header_fecha=False)
            pos_dfs.append(pos)
            tot_dfs.append(tot)

        # SOLO del primer archivo: cargar tambien "Posicion del mes anterior"
        # para tener el anchor previo y poder reconstruir el periodo desde el dia 1.
        if idx == 0:
            sn_prev = None
            for s in xl.sheet_names:
                if "anterior" in _norm(s):
                    sn_prev = s
                    break
            if sn_prev:
                pos_prev, tot_prev = _parse_posicion_sheet(
                    xl, sn_prev, file_date, cols, use_header_fecha=True
                )
                if not pos_prev.empty:
                    pos_dfs.append(pos_prev)
                    tot_dfs.append(tot_prev)

    if not pos_dfs:
        return pd.DataFrame(), pd.DataFrame()
    pos_all = pd.concat(pos_dfs, ignore_index=True)
    tot_all = pd.concat(tot_dfs, ignore_index=True)
    return pos_all, tot_all


# ---------------------------------------------------------------------------
# MTM diario
# ---------------------------------------------------------------------------
# Mapeo (emisora, serie) -> ticker yfinance.
# Para SC (cross-listed en BMV SIC) usamos sufijo .MX para precio en MXN.
# Para CO (BMV) usamos <emisora><serie>.MX.
# Si el ticker es None, valuamos al valor en libros (renta fija / reporto / instrumentos sin
# precio de mercado) interpolando el importe_neto entre snapshots.
TICKER_MAP: dict[tuple[str, str], str | None] = {
    # CO - BMV
    ("ALPEK", "A"): "ALPEKA.MX",
    ("ALSEA", "*"): "ALSEA.MX",
    ("AMX", "B"): "AMXB.MX",
    ("ARA", "*"): "ARA.MX",
    ("AUTLAN", "B"): "AUTLANB.MX",
    ("BOLSA", "A"): "BOLSAA.MX",
    ("CEMEX", "CPO"): "CEMEXCPO.MX",
    ("DANHOS", "13"): "DANHOS13.MX",
    ("FCFE", "18"): "FCFE18.MX",
    ("FEMSA", "UBD"): "FEMSAUBD.MX",
    ("FIBRAPL", "14"): "FIBRAPL14.MX",
    ("FIHO", "12"): "FIHO12.MX",
    ("FUNO", "11"): "FUNO11.MX",
    ("GFNORTE", "O"): "GFNORTEO.MX",
    ("GICSA", "B"): "GICSAB.MX",
    ("GMEXICO", "B"): "GMEXICOB.MX",
    ("KOF", "UBL"): "KOFUBL.MX",
    ("LASITE", "*"): "LASITE.MX",
    ("LASITE", "B-1"): "LASITE.MX",  # yfinance no tiene serie B-1; usamos serie *
    ("MEGA", "CPO"): "MEGACPO.MX",
    ("NEMAK", "A"): "NEMAKA.MX",
    ("PE&OLES", "*"): "PE&OLES.MX",
    ("PINFRA", "*"): "PINFRA.MX",
    ("R", "A"): "RA.MX",
    ("TLEVISA", "CPO"): "TLEVISACPO.MX",
    ("TRAXION", "A"): "TRAXIONA.MX",
    ("VASCONI", "*"): "VASCONI.MX",
    ("VESTA", "*"): "VESTA.MX",
    # SC - cross-listed en BMV SIC (precio en MXN con sufijo .MX)
    ("HYG", "*"): "HYG.MX",
    ("PFE", "*"): "PFE.MX",
    ("SHV", "*"): "SHV.MX",
    ("SHY", "*"): "SHY.MX",
    ("SPG", "*"): "SPG.MX",
    ("SPHY", "*"): "SPHY.MX",
    ("T", "*"): "T.MX",
    ("UNH", "*"): "UNH.MX",
    ("UPS", "*"): "UPS.MX",
    # D y R - sin precio de mercado, valor en libros
    ("EXITUCB", "24"): None,
    ("TPLAYCB", "20"): None,
    ("BPAG91", "280511"): None,
    ("BPAG91", "280907"): None,
    ("BONDESF", None): None,
}


def get_ticker(emisora: str, serie: str | None) -> str | None:
    """Devuelve ticker yfinance para (emisora, serie) o None si no hay mapeo / es renta fija."""
    e = str(emisora).strip()
    s = str(serie).strip() if serie is not None and str(serie).strip() != "nan" else "*"
    if (e, s) in TICKER_MAP:
        return TICKER_MAP[(e, s)]
    # fallback: emisora sola
    if (e, "*") in TICKER_MAP:
        return TICKER_MAP[(e, "*")]
    return None


# Para los SC cross-listed (US ETFs/acciones en SIC), el .MX puede tener
# baja liquidez y huecos. Mapeamos al ticker USD original; luego multiplicamos
# por el USDMXN para tener un precio en MXN sin huecos.
# Politica del usuario: TODAS las acciones americanas cross-listed en BMV/SIC
# se valuan con su precio de cierre directo en .MX (cotizacion en pesos en
# Mexico). No se calculan sinteticamente como USD x USDMXN. En dias sin trade
# en SIC, el reindex().ffill() en compute_daily_mtm lleva el ultimo precio
# conocido al siguiente dia disponible.
SC_USD_FALLBACK: dict[str, str] = {}
FX_TICKER = "MXN=X"  # USDMXN spot en yfinance


def _fetch_close_batch(
    tickers: list[str], start: str, end: str
) -> dict[str, pd.Series]:
    """Descarga batch de varios tickers en una sola request HTTP.

    Mucho mas eficiente y robusto que descargar uno por uno: yfinance hace
    una sola peticion al endpoint de quotes, lo que reduce drasticamente el
    rate-limiting que Yahoo aplica cuando varias IPs comparten subnets
    (caso tipico de Streamlit Cloud).
    """
    if not tickers:
        return {}
    try:
        data = yf.download(
            tickers, start=start, end=end, progress=False,
            auto_adjust=False, threads=True, group_by="ticker",
        )
    except Exception:
        return {}
    if data is None or data.empty:
        return {}

    out: dict[str, pd.Series] = {}

    if len(tickers) == 1:
        # Single ticker: data.columns es plana (no MultiIndex)
        t = tickers[0]
        if "Close" in data.columns:
            s = pd.to_numeric(data["Close"], errors="coerce").dropna()
            if not s.empty:
                s.index = pd.to_datetime(s.index)
                if s.index.tz is not None:
                    s.index = s.index.tz_localize(None)
                out[t] = s.sort_index()
        return out

    if isinstance(data.columns, pd.MultiIndex):
        # group_by='ticker' devuelve (ticker, campo)
        for t in tickers:
            try:
                if (t, "Close") in data.columns:
                    s = pd.to_numeric(data[(t, "Close")], errors="coerce").dropna()
                    if not s.empty:
                        s.index = pd.to_datetime(s.index)
                        if s.index.tz is not None:
                            s.index = s.index.tz_localize(None)
                        out[t] = s.sort_index()
            except Exception:
                continue
    return out


def _fetch_close(ticker: str, start: str, end: str) -> pd.Series | None:
    """Descarga el cierre diario sin ajustar de un ticker; tolera errores.

    IMPORTANTE: Usamos `auto_adjust=False` y la columna `Close` (precio
    no ajustado por dividendos/splits). Para MTM historico necesitamos el
    precio que REALMENTE cotizo ese dia, no el ajustado retroactivamente.
    Si usaramos `Adj Close`, los precios historicos estarian deflactados
    por dividendos posteriores, subestimando el valor del portafolio en
    fechas pasadas.
    """
    try:
        data = yf.download(
            ticker, start=start, end=end, progress=False,
            auto_adjust=False, threads=False,
        )
        if data is None or data.empty:
            return None
        if isinstance(data.columns, pd.MultiIndex):
            if ("Close", ticker) in data.columns:
                s = data[("Close", ticker)]
            elif "Close" in data.columns.get_level_values(0):
                s = data["Close"].iloc[:, 0]
            else:
                return None
        elif "Close" in data.columns:
            s = data["Close"]
        else:
            return None
        s = pd.to_numeric(s, errors="coerce").dropna()
        if s.empty:
            return None
        s.index = pd.to_datetime(s.index)
        if s.index.tz is not None:
            s.index = s.index.tz_localize(None)
        return s.sort_index()
    except Exception:
        return None


# ---------------------------------------------------------------------------
# Cache de precios en disco
# ---------------------------------------------------------------------------
# El dashboard no descarga precios de Yahoo en cada carga: lee del CSV
# `prices/precios_diarios.csv` que se mantiene committeado en el repo.
# Cuando agregas un mes nuevo de archivos LayOut*.xlsm, corres localmente
# `python update_prices.py` y el script actualiza el CSV con los dias
# faltantes. Luego `git push` y Streamlit Cloud usa los precios nuevos
# sin tocar Yahoo.
PRICES_CSV = Path(__file__).parent / "prices" / "precios_diarios.csv"


def _load_cached_prices() -> pd.DataFrame:
    """Carga el CSV de precios cacheados. Devuelve DataFrame wide
    (index=fecha, columnas=tickers). Vacio si no existe."""
    if not PRICES_CSV.exists():
        return pd.DataFrame()
    try:
        df = pd.read_csv(PRICES_CSV, parse_dates=["fecha"], index_col="fecha")
        return df.sort_index()
    except Exception:
        return pd.DataFrame()


def _save_cached_prices(df: pd.DataFrame) -> None:
    """Guarda el DataFrame de precios al CSV (sobreescribe)."""
    PRICES_CSV.parent.mkdir(parents=True, exist_ok=True)
    df = df.copy()
    df.index.name = "fecha"
    df.sort_index().to_csv(PRICES_CSV)


@st.cache_data(ttl=24 * 3600, show_spinner="Descargando precios de yfinance...")
def fetch_prices(tickers: tuple[str, ...], start: str, end: str) -> pd.DataFrame:
    """Devuelve DataFrame con index=fecha, columnas=tickers, valores=precio cierre ajustado.

    Para tickers SC cross-listed (HYG.MX, SPHY.MX, etc.) que tienen baja liquidez
    y huecos en yfinance, se complementa con USD x USDMXN: las fechas faltantes
    en el .MX se rellenan con el ticker USD original convertido a pesos. Esto
    asegura que ETFs ilquidos como SPHY.MX tengan precio TODOS los dias habiles.
    """
    if not tickers:
        return pd.DataFrame()

    start_ts = pd.Timestamp(start)
    end_ts = pd.Timestamp(end)

    # 0) Intentar servir desde cache en disco (precios/precios_diarios.csv)
    cached = _load_cached_prices()
    if not cached.empty:
        # Verificar cobertura: todos los tickers solicitados estan, y la
        # ultima fecha cacheada cubre o casi cubre el end_ts.
        tickers_in_cache = [t for t in tickers if t in cached.columns]
        last_cached = cached.index.max()
        # Buffer de 3 dias habiles: si end_ts esta dentro de ultimo + 3
        # dias, consideramos que el cache es suficiente.
        cache_recent_enough = last_cached >= end_ts - pd.Timedelta(days=3)
        all_tickers_present = len(tickers_in_cache) == len(tickers)

        if all_tickers_present and cache_recent_enough:
            # 100% servido desde cache, cero llamadas a Yahoo
            out_df = cached.loc[
                (cached.index >= start_ts) & (cached.index <= end_ts),
                tickers_in_cache,
            ]
            if not out_df.empty:
                return out_df

    # 1) Descarga batch de todos los tickers .MX en UNA sola request
    out: dict[str, pd.Series] = _fetch_close_batch(list(tickers), start, end)

    # 1b) Reintentar los que fallaron. En Streamlit Cloud, Yahoo aplica rate
    # limit por subnet. Estrategia: backoff exponencial + un ultimo intento
    # con pausa larga.
    import time
    failed = [t for t in tickers if t not in out or out[t].empty]
    if failed:
        # Primer round: 3 intentos con delays crecientes
        for t in list(failed):
            for attempt in range(3):
                time.sleep(1.0 + 0.7 * attempt)  # 1.0s, 1.7s, 2.4s
                s = _fetch_close(t, start, end)
                if s is not None and not s.empty:
                    out[t] = s
                    failed.remove(t)
                    break
        # Segundo round (los todavia fallidos): pausa larga y un retry batch
        if failed:
            time.sleep(5.0)
            recovery = _fetch_close_batch(failed, start, end)
            for t, s in recovery.items():
                if s is not None and not s.empty:
                    out[t] = s

    # 2) Para los SC cross-listed: USD x USDMXN es la fuente PRIMARIA
    # (NYSE liquidez real = sin outliers de baja liquidez en SIC). El .MX
    # queda solo como respaldo para holidays US donde el USD no tradeo.
    sc_requested = [t for t in tickers if t in SC_USD_FALLBACK]
    if sc_requested:
        usd_needed = sorted({SC_USD_FALLBACK[t] for t in sc_requested})
        usd_universe = usd_needed + [FX_TICKER]
        # Batch + reintentos individuales con backoff
        usd_data: dict[str, pd.Series] = _fetch_close_batch(usd_universe, start, end)
        usd_failed = [t for t in usd_universe if t not in usd_data or usd_data[t].empty]
        for t in list(usd_failed):
            for attempt in range(3):
                time.sleep(1.0 + 0.7 * attempt)
                s = _fetch_close(t, start, end)
                if s is not None and not s.empty:
                    usd_data[t] = s
                    usd_failed.remove(t)
                    break
        if usd_failed:
            time.sleep(5.0)
            recovery = _fetch_close_batch(usd_failed, start, end)
            for t, s in recovery.items():
                if s is not None and not s.empty:
                    usd_data[t] = s

        if FX_TICKER in usd_data:
            fx = usd_data[FX_TICKER]
            for mx_ticker in sc_requested:
                usd_ticker = SC_USD_FALLBACK[mx_ticker]
                if usd_ticker not in usd_data:
                    continue
                # Sintetico: USD * USDMXN
                aligned = pd.concat(
                    [usd_data[usd_ticker].rename("usd"), fx.rename("fx")], axis=1
                )
                aligned["mxn"] = aligned["usd"] * aligned["fx"]
                synthetic = aligned["mxn"].dropna()
                if synthetic.empty:
                    # Si no hay USD/FX, dejamos el .MX si existe
                    continue

                if mx_ticker in out:
                    secondary = out[mx_ticker]
                    full_idx = synthetic.index.union(secondary.index)
                    primary_full = synthetic.reindex(full_idx)
                    secondary_full = secondary.reindex(full_idx)
                    # PRIMARIO = sintetico USD*FX; respaldo = .MX
                    out[mx_ticker] = primary_full.fillna(secondary_full)
                else:
                    out[mx_ticker] = synthetic

    if not out:
        return pd.DataFrame()
    df = pd.DataFrame(out)
    df.index = pd.to_datetime(df.index)
    if df.index.tz is not None:
        df.index = df.index.tz_localize(None)
    df = df.sort_index()
    return df


def build_holdings_calendar(
    pos: pd.DataFrame, mov: pd.DataFrame, dates: pd.DatetimeIndex
) -> pd.DataFrame:
    """Construye un calendario diario de tenencias por (subcuenta, emisora, serie, tp).

    Para cada portafolio:
      - Universo = todas las (emisora, serie) que aparecieron en algun snapshot del portafolio.
      - Para cada fecha de snapshot, si la combinacion NO aparece, asumimos titulos=0
        en esa fecha (la posicion fue liquidada).
      - Entre dos snapshots consecutivos, los titulos = snapshot_anterior + cumsum
        de movimientos (COMPRA/ENTRADA suma; VENTA/AMORTIZACION/SALIDA resta).
      - Antes del primer snapshot del portafolio: titulos=0.
    """
    if pos.empty:
        return pd.DataFrame()

    affecting_pos = {"COMPRA CAPITALES", "ENTRADA CANJE"}
    affecting_neg = {"VENTA CAPITALES", "AMORTIZACION", "SALIDA CANJE"}

    pos = pos.copy()
    pos["serie_n"] = pos["serie"].astype(str).str.strip()
    mov = mov.copy()
    mov["serie_n"] = mov["serie"].astype(str).str.strip()

    out_rows = []
    for port in pos["subcuenta"].unique():
        port_pos = pos[pos["subcuenta"] == port]
        snap_dates = sorted(port_pos["fecha"].dropna().unique())
        if not snap_dates:
            continue
        snap_dates = [pd.Timestamp(d) for d in snap_dates]
        instruments = port_pos[["emisora", "serie_n", "tp"]].drop_duplicates()

        for _, instr in instruments.iterrows():
            emi, serie, tp = instr["emisora"], instr["serie_n"], instr["tp"]
            # Snapshots para esta combinacion (titulos por fecha)
            sn = (
                port_pos[
                    (port_pos["emisora"] == emi)
                    & (port_pos["serie_n"] == serie)
                ]
                .set_index("fecha")["titulos"]
                .astype(float)
            )
            # Reindex sobre TODAS las fechas de snapshot del portafolio.
            # Si la posicion no aparece en una fecha => 0 (fue liquidada).
            sn_full = sn.reindex(snap_dates).fillna(0.0)

            # Movimientos diarios netos
            m = mov[
                (mov["subcuenta"] == port)
                & (mov["emisora"] == emi)
                & (mov["serie_n"] == serie)
            ].copy()
            m["delta"] = 0.0
            mask_pos = m["concepto"].isin(affecting_pos)
            mask_neg = m["concepto"].isin(affecting_neg)
            m.loc[mask_pos, "delta"] = m.loc[mask_pos, "titulos"].fillna(0).astype(float)
            m.loc[mask_neg, "delta"] = -m.loc[mask_neg, "titulos"].fillna(0).astype(float)
            m_daily = (
                m.dropna(subset=["fecha_op"]).groupby("fecha_op")["delta"].sum()
            )

            # Construir serie diaria
            ser = pd.Series(0.0, index=dates)

            for i, sd in enumerate(snap_dates):
                next_sd = (
                    snap_dates[i + 1]
                    if i + 1 < len(snap_dates)
                    else dates.max() + pd.Timedelta(days=1)
                )
                sd_titulos = float(sn_full.loc[sd])

                # Dia exacto del snapshot: valor real
                if sd in ser.index:
                    ser.loc[sd] = sd_titulos

                # Dias estrictamente entre sd y next_sd: anchor + cumsum de deltas
                mask_between = (ser.index > sd) & (ser.index < next_sd)
                if mask_between.any():
                    deltas_in = m_daily[
                        (m_daily.index > sd) & (m_daily.index < next_sd)
                    ]
                    if not deltas_in.empty:
                        cum = (
                            deltas_in.reindex(ser.index[mask_between], fill_value=0).cumsum()
                        )
                        ser.loc[ser.index[mask_between]] = sd_titulos + cum.values
                    else:
                        ser.loc[ser.index[mask_between]] = sd_titulos

            # Dias anteriores al primer snapshot: 0 (default)
            for d, t in ser.items():
                out_rows.append(
                    {
                        "fecha": d,
                        "subcuenta": port,
                        "emisora": emi,
                        "serie": serie,
                        "tp": tp,
                        "titulos": float(t),
                    }
                )

    return pd.DataFrame(out_rows)


def interpolate_importe_neto(
    pos: pd.DataFrame, dates: pd.DatetimeIndex
) -> pd.DataFrame:
    """Importe_neto interpolado linealmente entre snapshots, por
    (subcuenta, emisora, serie, tp). Cuando la posicion no aparece en una fecha
    de snapshot, su importe ese dia es 0 (fue liquidada)."""
    if pos.empty:
        return pd.DataFrame()
    pos = pos.copy()
    pos["serie_n"] = pos["serie"].astype(str).str.strip()

    out = []
    for port in pos["subcuenta"].unique():
        port_pos = pos[pos["subcuenta"] == port]
        snap_dates = sorted(pd.Timestamp(d) for d in port_pos["fecha"].dropna().unique())
        if not snap_dates:
            continue
        instruments = port_pos[["emisora", "serie_n", "tp"]].drop_duplicates()
        for _, instr in instruments.iterrows():
            emi, serie, tp = instr["emisora"], instr["serie_n"], instr["tp"]
            sn = (
                port_pos[
                    (port_pos["emisora"] == emi)
                    & (port_pos["serie_n"] == serie)
                ]
                .set_index("fecha")["importe_neto"]
                .astype(float)
            )
            # Reindex en todas las fechas de snapshot del portafolio. Si no esta -> 0.
            sn_full = sn.reindex(snap_dates).fillna(0.0)
            # Reindex al calendario diario; interpolar entre snapshots; antes del 1er snap = 0.
            ser = sn_full.reindex(dates).interpolate(method="time").fillna(0.0)
            for d, v in ser.items():
                out.append(
                    {
                        "fecha": d,
                        "subcuenta": port,
                        "emisora": emi,
                        "serie": serie,
                        "tp": tp,
                        "importe_neto": float(v),
                    }
                )
    return pd.DataFrame(out)


def _normalize_repo_series(df: pd.DataFrame) -> pd.DataFrame:
    """Colapsa las distintas series de un repo (BPAG91 280907, 280511, 290913)
    bajo '*' porque son la misma posicion de money market, solo con plazo
    distinto. Sin esta normalizacion, el merge holdings x carry falla porque
    el snapshot trae una serie y los movimientos intra-mes traen otras.

    Aplica a:
      - pos: filas con tp == 'R'
      - mov: filas con concepto IN ('INICIO CPA REPORTO', 'VEN.COMPRA REPORTO')
    """
    if df.empty:
        return df
    df = df.copy()
    if "tp" in df.columns:
        mask = df["tp"] == "R"
        if mask.any():
            df.loc[mask, "serie"] = "*"
    if "concepto" in df.columns:
        mask = df["concepto"].isin(["INICIO CPA REPORTO", "VEN.COMPRA REPORTO"])
        if mask.any():
            df.loc[mask, "serie"] = "*"
    return df


def compute_repo_values_daily(
    mov: pd.DataFrame, dates: pd.DatetimeIndex
) -> pd.DataFrame:
    """Valuacion exacta de reportos abiertos cada dia.

    Para cada `INICIO CPA REPORTO` busca su `VEN.COMPRA REPORTO` matching por
    folio. Mientras el repo esta vivo (start <= dia < end), su valor es:
        monto x (1 + tasa/360 x dias_transcurridos)
    donde monto = monto_neto del INICIO y tasa = tasa_premio/100 (tasa anual).

    Si un mismo (subcuenta, emisora, serie) tiene varios repos abiertos
    simultaneamente, sus valores se suman.

    Devuelve long DataFrame: (fecha, subcuenta, emisora, serie, tp, valor_repo).
    """
    if mov.empty:
        return pd.DataFrame(
            columns=["fecha", "subcuenta", "emisora", "serie", "tp", "valor_repo"]
        )

    inicios = mov[mov["concepto"] == "INICIO CPA REPORTO"].copy()
    if inicios.empty:
        return pd.DataFrame(
            columns=["fecha", "subcuenta", "emisora", "serie", "tp", "valor_repo"]
        )
    # Mismo folio puede aparecer en archivos contiguos: dedup
    inicios = inicios.drop_duplicates(subset=["folio", "subcuenta"])

    vens = mov[mov["concepto"] == "VEN.COMPRA REPORTO"].copy()
    vens = vens.drop_duplicates(subset=["folio", "subcuenta"])
    vens_idx = vens.set_index(["folio", "subcuenta"])["fecha_liq"].to_dict()

    rows = []
    dates_idx = pd.DatetimeIndex(dates)

    for _, r in inicios.iterrows():
        if pd.isna(r["fecha_liq"]):
            continue
        start_date = pd.Timestamp(r["fecha_liq"])

        # Fecha de cierre del repo
        ven_key = (r["folio"], r["subcuenta"])
        if ven_key in vens_idx and pd.notna(vens_idx[ven_key]):
            end_date = pd.Timestamp(vens_idx[ven_key])
        elif pd.notna(r["plazo"]):
            end_date = start_date + pd.Timedelta(days=int(r["plazo"]))
        else:
            continue

        monto = float(r["monto_neto"]) if pd.notna(r["monto_neto"]) else 0.0
        if monto <= 0:
            continue

        # tasa_premio en INICIO es la tasa anual %
        rate = float(r["tasa_premio"]) / 100 if pd.notna(r["tasa_premio"]) else 0.0
        port = r["subcuenta"]
        emi = r["emisora"]
        serie = (
            str(r["serie"]).strip()
            if pd.notna(r["serie"]) and str(r["serie"]).strip() != "nan"
            else "*"
        )

        # Iterar dias [start, end)
        in_range = dates_idx[(dates_idx >= start_date) & (dates_idx < end_date)]
        for d in in_range:
            days_elapsed = (d - start_date).days
            value = monto * (1 + rate / 360 * days_elapsed)
            rows.append({
                "fecha": d,
                "subcuenta": port,
                "emisora": emi,
                "serie": serie,
                "tp": "R",
                "valor_repo": value,
            })

    if not rows:
        return pd.DataFrame(
            columns=["fecha", "subcuenta", "emisora", "serie", "tp", "valor_repo"]
        )

    df = pd.DataFrame(rows)
    df = df.groupby(
        ["fecha", "subcuenta", "emisora", "serie", "tp"], as_index=False
    )["valor_repo"].sum()
    return df


def compute_bond_values_daily(
    pos: pd.DataFrame, mov: pd.DataFrame, dates: pd.DatetimeIndex
) -> pd.DataFrame:
    """Valuacion diaria exacta de bonos (TP=D) por simulacion entre snapshots.

    Modelo:
      Entre dos snapshots consecutivos t0 y t1:
        valor[t0] = importe_neto[t0]  (anchor)
        valor[t1] = importe_neto[t1]  (cierre)
        Cashflows discretos en (t0, t1]: amortizaciones (parciales y totales).
        Cupones NO afectan importe_neto (solo entran a cash; el dirty price
        del bono no incluye el cupon una vez pagado).

      Drift continuo total = (valor[t1] - valor[t0]) + sum(amortizaciones)
        - Refleja devengo + cambios de clean price.
        - Se distribuye linealmente por dia.

      Para cada dia d en (t0, t1):
        valor[d] = valor[t0] + drift_per_day * (d - t0)
                              - sum(amortizaciones liquidadas en [t0+1, d])

    Esto cierra exactamente en cada snapshot por construccion.

    Si el bono desaparece en t1 (titulos a 0): bond_imp_full[t1] = 0 y la
    diferencia entre t0 y t1 se distribuye igualmente, con la amortizacion
    final aplicandose en su fecha exacta.
    """
    bonds_pos = pos[pos["tp"] == "D"].copy()
    if bonds_pos.empty:
        return pd.DataFrame(
            columns=["fecha", "subcuenta", "emisora", "serie", "tp", "importe_neto"]
        )

    bonds_pos["serie_n"] = bonds_pos["serie"].astype(str).str.strip()

    out_rows = []
    dates_idx = pd.DatetimeIndex(dates)

    for (port, emi, serie_n), grp in bonds_pos.groupby(
        ["subcuenta", "emisora", "serie_n"]
    ):
        # Snapshots del PORTAFOLIO (no solo del bono): para que cuando el bono
        # desaparezca, ese snapshot anchor el valor a 0.
        port_snaps = sorted(
            pd.Timestamp(d) for d in pos[pos["subcuenta"] == port]["fecha"].dropna().unique()
        )
        if not port_snaps:
            continue

        bond_imp = grp.set_index("fecha")["importe_neto"].astype(float)
        bond_imp_full = bond_imp.reindex(port_snaps).fillna(0.0)

        # Amortizaciones del bono (parciales y totales). Dedup por (folio, fecha_liq).
        amorts = mov[
            (mov["subcuenta"] == port)
            & (mov["emisora"] == emi)
            & (mov["serie"].astype(str).str.strip() == serie_n)
            & (mov["concepto"] == "AMORTIZACION")
        ][["fecha_liq", "folio", "monto_neto"]].dropna(subset=["fecha_liq"]).copy()
        if not amorts.empty:
            amorts["fecha_liq"] = pd.to_datetime(amorts["fecha_liq"])
            amorts["monto_neto"] = pd.to_numeric(amorts["monto_neto"], errors="coerce").fillna(0)
            amorts = amorts.drop_duplicates(["folio", "fecha_liq", "monto_neto"])
            cf_by_date = amorts.groupby("fecha_liq")["monto_neto"].sum()
            cf_by_date.index = pd.to_datetime(cf_by_date.index)
        else:
            # Empty series with DatetimeIndex (no RangeIndex which fails comparison)
            cf_by_date = pd.Series(dtype=float, index=pd.DatetimeIndex([]))

        # Inicializar serie diaria
        ser = pd.Series(np.nan, index=dates_idx)
        # Anchor en cada snapshot
        for t in port_snaps:
            if t in ser.index:
                ser.loc[t] = float(bond_imp_full.loc[t])

        # Simulacion entre cada par consecutivo
        for i in range(len(port_snaps) - 1):
            t0 = port_snaps[i]
            t1 = port_snaps[i + 1]
            v0 = float(bond_imp_full.loc[t0])
            v1 = float(bond_imp_full.loc[t1])
            num_days = (t1 - t0).days
            if num_days <= 0:
                continue

            # CF en (t0, t1]
            cf_in_period = cf_by_date[
                (cf_by_date.index > t0) & (cf_by_date.index <= t1)
            ]
            cf_total = float(cf_in_period.sum())

            # Caso degenerado: bono no existe en ambos extremos
            if v0 == 0 and v1 == 0 and cf_total == 0:
                between = dates_idx[(dates_idx > t0) & (dates_idx < t1)]
                ser.loc[between] = 0.0
                continue

            drift_total = v1 - v0 + cf_total
            drift_per_day = drift_total / num_days

            # Walk forward dia por dia
            current = v0
            current_date = t0
            between = sorted(dates_idx[(dates_idx > t0) & (dates_idx < t1)])
            for d in between:
                days_elapsed = (d - current_date).days
                current += drift_per_day * days_elapsed
                # Aplicar amortizacion del dia (si existe)
                if d in cf_by_date.index:
                    current -= float(cf_by_date.loc[d])
                current_date = d
                ser.loc[d] = current

        # Antes del primer snapshot: 0
        before_first = dates_idx[dates_idx < port_snaps[0]]
        for d in before_first:
            ser.loc[d] = 0.0

        # Despues del ultimo snapshot: mantener valor del ultimo snapshot
        last_v = float(bond_imp_full.loc[port_snaps[-1]])
        after_last = dates_idx[dates_idx > port_snaps[-1]]
        for d in after_last:
            if pd.isna(ser.loc[d]):
                ser.loc[d] = last_v

        for d, v in ser.items():
            if pd.notna(v):
                out_rows.append({
                    "fecha": d,
                    "subcuenta": port,
                    "emisora": emi,
                    "serie": serie_n,
                    "tp": "D",
                    "importe_neto": float(v),
                })

    return pd.DataFrame(out_rows)


def carry_values_daily(
    pos: pd.DataFrame, mov: pd.DataFrame, dates: pd.DatetimeIndex
) -> pd.DataFrame:
    """Combina carry interpolado con valuacion exacta de reportos.

    Para TP=R: usa el valor exacto computado por `compute_repo_values_daily`.
    Para el resto: interpolacion lineal del importe_neto entre snapshots.
    Se hace outer-merge para capturar repos intra-mes que pudieran no estar
    en ningun snapshot (raro, pero posible).
    """
    base = interpolate_importe_neto(pos, dates)
    repos_exact = compute_repo_values_daily(mov, dates)
    bonds_exact = compute_bond_values_daily(pos, mov, dates)

    # Caso degenerado: nada que combinar
    if base.empty and repos_exact.empty and bonds_exact.empty:
        return pd.DataFrame(
            columns=["fecha", "subcuenta", "emisora", "serie", "tp", "importe_neto"]
        )

    if base.empty:
        base = pd.DataFrame(
            columns=["fecha", "subcuenta", "emisora", "serie", "tp", "importe_neto"]
        )

    base = base.copy()
    base["serie_n"] = base["serie"].astype(str).str.strip()

    # ---- Reemplazar TP=R con valuacion exacta de reportos ----
    if not repos_exact.empty:
        repos_exact = repos_exact.copy()
        repos_exact["serie_n"] = repos_exact["serie"].astype(str).str.strip()
        merged = base.merge(
            repos_exact[
                ["fecha", "subcuenta", "emisora", "serie_n", "tp", "valor_repo"]
            ],
            on=["fecha", "subcuenta", "emisora", "serie_n", "tp"],
            how="outer",
            indicator="_repo_merge",
        )
        is_repo_exact = (merged["tp"] == "R") & merged["valor_repo"].notna()
        merged["importe_neto"] = np.where(
            is_repo_exact,
            merged["valor_repo"],
            merged["importe_neto"],
        )
        only_repo = merged["_repo_merge"] == "right_only"
        merged.loc[only_repo, "serie"] = merged.loc[only_repo, "serie_n"]
        base = merged.drop(columns=["valor_repo", "_repo_merge"])

    # ---- Reemplazar TP=D con valuacion exacta de bonos ----
    if not bonds_exact.empty:
        bonds_exact = bonds_exact.copy()
        bonds_exact["serie_n"] = bonds_exact["serie"].astype(str).str.strip()
        merged = base.merge(
            bonds_exact[
                ["fecha", "subcuenta", "emisora", "serie_n", "tp", "importe_neto"]
            ].rename(columns={"importe_neto": "importe_bond_exact"}),
            on=["fecha", "subcuenta", "emisora", "serie_n", "tp"],
            how="outer",
            indicator="_bond_merge",
        )
        is_bond_exact = (merged["tp"] == "D") & merged["importe_bond_exact"].notna()
        merged["importe_neto"] = np.where(
            is_bond_exact,
            merged["importe_bond_exact"],
            merged["importe_neto"],
        )
        only_bond = merged["_bond_merge"] == "right_only"
        merged.loc[only_bond, "serie"] = merged.loc[only_bond, "serie_n"]
        base = merged.drop(columns=["importe_bond_exact", "_bond_merge"])

    base = base.drop(columns=["serie_n"])
    base = base.dropna(subset=["importe_neto"])
    return base


def daily_cash(mov: pd.DataFrame, dates: pd.DatetimeIndex) -> pd.DataFrame:
    """Saldo de efectivo de cierre por fecha de LIQUIDACION (no operacion).

    Razon: el `Saldo` en Movimientos refleja el estado de la cuenta despues de
    cada liquidacion. Un reporto iniciado el 30/06 que vence el 01/07 aparece
    en el archivo de junio (INICIO, fecha_liq=30/06) y en el de julio (VEN.,
    fecha_liq=01/07) con el mismo folio. Si agrupo por fecha_op del segundo,
    tomo el saldo POST-vencimiento como si fuera el saldo del 30/06, lo cual
    duplica el efectivo. Usando fecha_liq cada saldo cae en su dia real.
    """
    if mov.empty:
        return pd.DataFrame()
    mov = mov.dropna(subset=["fecha_liq", "saldo"]).copy()
    # Eliminar duplicados (mismo folio puede aparecer en archivos contiguos):
    # nos quedamos con la entrada de la fecha de archivo mas cercana a la
    # liquidacion, que es la que registra el saldo definitivo de ese dia.
    mov = mov.sort_values(["subcuenta", "fecha_liq", "folio", "archivo_fecha"])
    last = (
        mov.groupby(["subcuenta", "fecha_liq"], as_index=False)
        .agg(saldo=("saldo", "last"))
    )
    out = []
    for port, grp in last.groupby("subcuenta"):
        s = grp.set_index("fecha_liq")["saldo"].sort_index()
        s = s.reindex(dates).ffill().fillna(0.0)
        for d, v in s.items():
            out.append({"fecha": d, "subcuenta": port, "efectivo": float(v)})
    return pd.DataFrame(out)


@st.cache_data(show_spinner="Calculando MTM diario...")
def compute_daily_mtm(
    pos: pd.DataFrame, mov: pd.DataFrame, start: str, end: str
) -> tuple[pd.DataFrame, pd.DataFrame, list[str]]:
    """Devuelve (daily_value, holdings_with_prices, missing_tickers).

    daily_value: (fecha, subcuenta, valor_equity, valor_carry, efectivo, valor_total)
    holdings_with_prices: long, util para depurar.
    missing_tickers: instrumentos que se valuaron al carry porque yfinance no devolvio data.
    """
    # Normalizar series de reporto: BPAG91 280907/280511/290913/etc -> "*"
    # son todas la misma posicion de money market.
    pos = _normalize_repo_series(pos)
    mov = _normalize_repo_series(mov)
    # Para pos, si despues de normalizar quedaron filas duplicadas (mismo
    # (fecha, port, emi, "*", tp=R)), agregamos sumando los valores.
    repo_mask = pos["tp"] == "R"
    if repo_mask.any():
        repos_pos = pos[repo_mask].copy()
        non_repos = pos[~repo_mask]
        # Agregar todas las columnas numericas
        num_cols = repos_pos.select_dtypes(include="number").columns.tolist()
        non_num = [
            c for c in repos_pos.columns
            if c not in num_cols and c not in [
                "fecha", "subcuenta", "emisora", "serie", "tp"
            ]
        ]
        agg = {c: "sum" for c in num_cols}
        agg.update({c: "first" for c in non_num})
        repos_pos = repos_pos.groupby(
            ["fecha", "subcuenta", "emisora", "serie", "tp"], as_index=False
        ).agg(agg)
        pos = pd.concat([non_repos, repos_pos], ignore_index=True)

    bdays = pd.bdate_range(start=start, end=end)
    # Incluir tambien las fechas de snapshot (cierre de mes que cae en sabado/domingo)
    snap_dates = pd.DatetimeIndex(
        sorted(pd.Timestamp(d) for d in pos["fecha"].dropna().unique())
    )
    snap_in_range = snap_dates[
        (snap_dates >= pd.Timestamp(start)) & (snap_dates <= pd.Timestamp(end))
    ]
    dates = pd.DatetimeIndex(sorted(set(bdays).union(set(snap_in_range))))
    holdings = build_holdings_calendar(pos, mov, dates)
    if holdings.empty:
        return pd.DataFrame(), pd.DataFrame(), []

    # Asignar ticker
    holdings["ticker"] = holdings.apply(
        lambda r: get_ticker(r["emisora"], r["serie"]), axis=1
    )

    # Universo de tickers a descargar
    tickers = sorted({t for t in holdings["ticker"].dropna().unique() if t})

    prices = fetch_prices(tuple(tickers), start, (pd.Timestamp(end) + pd.Timedelta(days=2)).strftime("%Y-%m-%d"))

    # Reindex prices al calendario y forward-fill
    if not prices.empty:
        prices = prices.reindex(dates).ffill()

    missing = [t for t in tickers if t not in prices.columns or prices[t].isna().all()]

    # Long-format de precios
    if not prices.empty:
        price_long = prices.stack(future_stack=True).rename("precio").reset_index()
        price_long.columns = ["fecha", "ticker", "precio"]
    else:
        price_long = pd.DataFrame(columns=["fecha", "ticker", "precio"])

    # Merge holdings con precios
    h = holdings.merge(price_long, on=["fecha", "ticker"], how="left")

    # Para los que tienen ticker valido y precio, valor_mercado = titulos * precio
    h["valor_equity"] = h["titulos"] * h["precio"]

    # Para los que no tienen ticker:
    # - TP=R (reporto): valor EXACTO = monto x (1 + tasa/360 x dias_corridos)
    # - TP=D (bono) y otros: importe_neto interpolado entre snapshots
    carry = carry_values_daily(pos, mov, dates)
    h = h.merge(
        carry,
        on=["fecha", "subcuenta", "emisora", "serie", "tp"],
        how="left",
    )

    # Logica final por fila:
    # - Si tiene precio y ticker: valor = titulos * precio (equity)
    # - Si no: valor = importe_neto interpolado (carry)
    has_price = h["valor_equity"].notna() & (h["ticker"].notna())
    h["valor"] = np.where(has_price, h["valor_equity"], h["importe_neto"])
    h["fuente"] = np.where(has_price, "yfinance", "carry")

    # Cuando el equity tiene 0 titulos pero no esta delisted, el valor es 0
    h.loc[h["titulos"] == 0, "valor"] = 0.0

    # Agregar efectivo
    cash = daily_cash(mov, dates)

    # Agregar por (fecha, subcuenta) separando equity vs carry via pivot
    h["valor"] = pd.to_numeric(h["valor"], errors="coerce").fillna(0.0)
    pivoted = (
        h.groupby(["fecha", "subcuenta", "fuente"], as_index=False)["valor"]
        .sum()
        .pivot_table(
            index=["fecha", "subcuenta"], columns="fuente", values="valor",
            fill_value=0.0,
        )
        .reset_index()
    )
    if "yfinance" not in pivoted.columns:
        pivoted["yfinance"] = 0.0
    if "carry" not in pivoted.columns:
        pivoted["carry"] = 0.0
    pivoted = pivoted.rename(columns={"yfinance": "valor_equity", "carry": "valor_carry"})

    by_port = pivoted.merge(cash, on=["fecha", "subcuenta"], how="left").fillna({"efectivo": 0.0})
    by_port["valor_total_raw"] = (
        by_port["valor_equity"] + by_port["valor_carry"] + by_port["efectivo"]
    )

    # ----------------------------------------------------------------------
    # Calibracion contra snapshots oficiales de Posicion.
    # En cada fecha de snapshot, el valor oficial = sum(valor_mercado_neto)
    # incluye todo (equities + renta fija + reporto + cash + intereses
    # devengados). Calculamos el residual = oficial - raw en esos dias y lo
    # interpolamos linealmente entre snapshots para distribuirlo dia a dia.
    # Esto garantiza:
    #   - En cada cierre de mes: valor_total == valor oficial del archivo.
    #   - Entre snapshots: los movimientos diarios reflejan precios yfinance
    #     mas un ajuste suave que cubre devengo de bonos, diferencia de
    #     precios yfinance vs BMV, y otras pequenas brechas metodologicas.
    # ----------------------------------------------------------------------
    official = (
        pos.groupby(["fecha", "subcuenta"], as_index=False)["valor_mercado_neto"]
        .sum()
        .rename(columns={"valor_mercado_neto": "valor_oficial"})
    )
    by_port = by_port.merge(official, on=["fecha", "subcuenta"], how="left")

    # Anclas diarias del archivo de cobro (back-office). Si existe, calibra
    # contra ESTOS valores diarios (1 ancla por dia, mas granular). Si no,
    # cae a los snapshots mensuales de Posicion (12 anclas por año).
    cobro_diario = load_cobro_diario()
    if not cobro_diario.empty:
        by_port = by_port.merge(
            cobro_diario, on=["fecha", "subcuenta"], how="left"
        )
        # Ancla preferida: cobro diario. Fallback: valor_oficial del snapshot.
        by_port["ancla"] = by_port["valor_diario"].fillna(by_port["valor_oficial"])
    else:
        by_port["valor_diario"] = np.nan
        by_port["ancla"] = by_port["valor_oficial"]

    by_port["residual_anchor"] = by_port["ancla"] - by_port["valor_total_raw"]

    # Interpolacion temporal del residual por subcuenta (vectorizado).
    by_port = by_port.sort_values(["subcuenta", "fecha"]).reset_index(drop=True)
    pieces = []
    for port, g in by_port.groupby("subcuenta", sort=False):
        g = g.sort_values("fecha").set_index("fecha")
        g["ajuste_calibracion"] = (
            g["residual_anchor"]
            .interpolate(method="time", limit_direction="both")
            .fillna(0.0)
        )
        g = g.reset_index()
        pieces.append(g)
    by_port = pd.concat(pieces, ignore_index=True)
    by_port["valor_total"] = by_port["valor_total_raw"] + by_port["ajuste_calibracion"]

    # Orden de columnas
    cols_order = [
        "fecha", "subcuenta",
        "valor_equity", "valor_carry", "efectivo",
        "valor_total_raw", "ajuste_calibracion", "valor_total",
        "valor_oficial", "valor_diario",
    ]
    by_port = by_port[cols_order].sort_values(["subcuenta", "fecha"]).reset_index(drop=True)

    return by_port, h, missing


def build_historico_excel(daily_mtm: pd.DataFrame) -> bytes:
    """Genera un Excel con el historico diario del valor del portafolio.

    Hojas:
      1. 'Resumen consolidado' - una columna por portafolio + Total
      2. 'Rendimiento'         - retorno acumulado % por portafolio
      3. Una hoja por portafolio (VIIN_XXX1, VIIN_XXX3, VIIN_XXX6) con
         el desglose diario: equity, carry, efectivo, raw, ajuste,
         total, oficial.
    """
    buf = io.BytesIO()

    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        # ---- Sheet 1: Resumen consolidado (valor total por portafolio + Total) ----
        resumen = (
            daily_mtm.pivot_table(
                index="fecha", columns="subcuenta", values="valor_total"
            )
            .sort_index()
        )
        resumen.columns = [PORT_LABEL.get(c, c) for c in resumen.columns]
        resumen["Total Cliente"] = resumen.sum(axis=1)
        resumen.index = resumen.index.strftime("%Y-%m-%d")
        resumen.index.name = "Fecha"
        resumen.to_excel(writer, sheet_name="Resumen consolidado", float_format="%.2f")

        # ---- Sheet 2: Rendimiento acumulado % ----
        rend = resumen.copy()
        for c in rend.columns:
            first = rend[c].iloc[0]
            if first and not pd.isna(first):
                rend[c] = rend[c] / first - 1
        rend.to_excel(writer, sheet_name="Rendimiento acumulado %",
                       float_format="%.6f")

        # ---- Sheets por portafolio ----
        for port in sorted(daily_mtm["subcuenta"].unique()):
            sub = (
                daily_mtm[daily_mtm["subcuenta"] == port]
                .sort_values("fecha")
                .copy()
            )
            sub_out = sub[[
                "fecha", "valor_equity", "valor_carry", "efectivo",
                "valor_total_raw", "ajuste_calibracion", "valor_total",
                "valor_oficial",
            ]].copy()
            sub_out["fecha"] = pd.to_datetime(sub_out["fecha"]).dt.strftime("%Y-%m-%d")
            sub_out.columns = [
                "Fecha", "Equity (MTM)", "Renta fija / Reporto",
                "Efectivo", "Suma raw", "Ajuste calibracion",
                "Valor total", "Valor oficial (snapshots)",
            ]
            # Hoja con el sufijo del portafolio (ultimos 4 chars)
            sheet_name = f"VIIN_{port[-4:]}"
            sub_out.to_excel(
                writer, sheet_name=sheet_name, index=False, float_format="%.2f"
            )

    buf.seek(0)
    return buf.getvalue()


def reconstruct_position(
    portfolio: str,
    target_date: pd.Timestamp,
    pos: pd.DataFrame,
    holdings_full: pd.DataFrame,
    daily_mtm: pd.DataFrame,
) -> tuple[pd.DataFrame, dict]:
    """Reconstruye la hoja Posicion para `portfolio` a la fecha `target_date`.

    Para cada (emisora, serie) que el portafolio tiene con titulos > 0 ese dia,
    arma una fila con la misma estructura que Posicion del archivo. Los campos
    estaticos (cupon, plazo, tasa, estrategia) se toman del snapshot mas reciente
    <= target_date. dias_x_ven se descuenta por dias transcurridos. Los precios
    de mercado vienen de yfinance (o del importe_neto interpolado para renta
    fija). El valor_mercado_neto y plus_minus reflejan los datos del dia.

    Devuelve (df_posicion, totals_dict).
    """
    target_ts = pd.Timestamp(target_date)

    # Holdings del dia
    h = holdings_full[
        (holdings_full["subcuenta"] == portfolio)
        & (holdings_full["fecha"] == target_ts)
    ].copy()

    # Snapshots del portafolio
    snaps = pos[pos["subcuenta"] == portfolio].copy()
    snaps["serie_n"] = snaps["serie"].astype(str).str.strip()
    h["serie_n"] = h["serie"].astype(str).str.strip()

    # Solo posiciones con titulos vivos
    h_alive = h[h["titulos"] > 0].copy()

    rows = []
    for _, hr in h_alive.iterrows():
        emi, serie_n = hr["emisora"], hr["serie_n"]
        # Snapshot mas reciente <= target_ts
        match = snaps[
            (snaps["emisora"] == emi)
            & (snaps["serie_n"] == serie_n)
            & (snaps["fecha"] <= target_ts)
        ].sort_values("fecha").tail(1)

        if not match.empty:
            sr = match.iloc[0]
            ago_days = max(0, (target_ts - pd.Timestamp(sr["fecha"])).days)
            cupon = sr.get("cupon")
            plazo = sr.get("plazo")
            tasa = sr.get("tasa")
            dias_orig = sr.get("dias_x_ven")
            dias_x_ven = (
                max(0, dias_orig - ago_days)
                if pd.notna(dias_orig) else None
            )
            precio_costo = sr.get("precio")
            importe_bruto = sr.get("importe_bruto")
            precio_neto = sr.get("precio_neto")
            estrategia = sr.get("estrategia")
            serie_display = sr.get("serie", hr["serie"])
        else:
            cupon = plazo = tasa = dias_x_ven = None
            precio_costo = importe_bruto = precio_neto = None
            estrategia = None
            serie_display = hr["serie"]

        # Valor de mercado y precio de mercado
        if hr.get("fuente") == "yfinance" and pd.notna(hr.get("precio")):
            precio_mercado = float(hr["precio"])
            valor_mercado_neto = float(hr["titulos"]) * precio_mercado
        else:
            # carry: usamos importe_neto interpolado del dia
            valor_mercado_neto = float(hr.get("valor", 0.0))
            precio_mercado = (
                valor_mercado_neto / float(hr["titulos"])
                if hr["titulos"] else None
            )

        # importe_neto del dia (interpolado): viene en holdings_full
        importe_neto_dia = (
            float(hr["importe_neto"]) if pd.notna(hr.get("importe_neto"))
            else None
        )

        plus_minus_int = (
            valor_mercado_neto - importe_neto_dia
            if importe_neto_dia is not None else None
        )
        plus_minus_pct = (
            plus_minus_int / importe_neto_dia
            if (plus_minus_int is not None and importe_neto_dia)
            else None
        )

        rows.append({
            "Posicion": hr["tp"],
            "Subcuentas": portfolio,
            "Emisora": emi,
            "Serie": serie_display,
            "Cupon": cupon,
            "Plazo": plazo,
            "Tasa": tasa,
            "Dias x Ven": dias_x_ven,
            "Titulos": float(hr["titulos"]),
            "Precio": precio_costo,
            "Importe Bruto": importe_bruto,
            "Precio Neto": precio_neto,
            "Importe Neto": importe_neto_dia,
            "Precio de Mercado": precio_mercado,
            "Valor de Mercado Neto": valor_mercado_neto,
            "Plus/Minus + Int. Dev.": plus_minus_int,
            "Plus/Minus %": plus_minus_pct,
            "% de Cartera": None,  # se llena despues
            "Estrategia": estrategia,
        })

    df = pd.DataFrame(rows)

    # Totales del dia desde daily_mtm
    drow = daily_mtm[
        (daily_mtm["subcuenta"] == portfolio)
        & (daily_mtm["fecha"] == target_ts)
    ]
    if not drow.empty:
        d = drow.iloc[0]
        valor_portafolio = float(d["valor_total"])
        efectivo = float(d["efectivo"])
        valor_equity = float(d["valor_equity"])
        valor_carry = float(d["valor_carry"])
        ajuste = float(d["ajuste_calibracion"])
    else:
        valor_portafolio = df["Valor de Mercado Neto"].sum() if not df.empty else 0
        efectivo = 0.0
        valor_equity = valor_carry = ajuste = 0.0

    # % de Cartera (basado en valor_portafolio TOTAL, incluido cash)
    if not df.empty and valor_portafolio:
        df["% de Cartera"] = df["Valor de Mercado Neto"] / valor_portafolio

    # Orden por TP, emisora, serie (como aparece en Posicion)
    if not df.empty:
        df = df.sort_values(["Posicion", "Emisora", "Serie"]).reset_index(drop=True)

    totals = {
        "valor_portafolio": valor_portafolio,
        "efectivo": efectivo,
        "valor_equity": valor_equity,
        "valor_carry": valor_carry,
        "ajuste_calibracion": ajuste,
        "es_snapshot": target_ts in pd.DatetimeIndex(pos["fecha"].dropna().unique()),
    }
    return df, totals


# ---------------------------------------------------------------------------
# Sidebar / file discovery
# ---------------------------------------------------------------------------
st.sidebar.title("Portafolios VIIN")
files = sorted(LAYOUTS_DIR.glob("LayOut*.xlsm"))
if not files:
    st.error(f"No encontre archivos en {LAYOUTS_DIR}")
    st.stop()

st.sidebar.caption(f"{len(files)} archivos detectados")

with st.spinner("Cargando archivos..."):
    mov = load_movimientos([str(f) for f in files])
    pos, tot = load_posiciones([str(f) for f in files])

if mov.empty:
    st.error("No se cargo ningun movimiento.")
    st.stop()

# ---------------------------------------------------------------------------
# Sidebar filters
# ---------------------------------------------------------------------------
sel_ports = st.sidebar.multiselect(
    "Portafolio", PORTFOLIOS, default=PORTFOLIOS
)
# Rango por defecto = periodo de operacion real (snapshots de Posicion).
# Movimientos pueden tener fecha_op historicas (2022, etc.) por referencias
# a operaciones originales de bonos, pero esas no son fechas de actividad
# nueva — solo se usan internamente para identificar el bono.
snap_min = pd.to_datetime(pos["fecha"].dropna().min())
snap_max = pd.to_datetime(pos["fecha"].dropna().max())
fmin_abs = pd.to_datetime(mov["fecha_op"].min())  # solo para min_value
fmax_abs = pd.to_datetime(mov["fecha_op"].max())
date_range = st.sidebar.date_input(
    "Rango de fechas",
    value=(snap_min.date(), snap_max.date()),
    min_value=min(snap_min.date(), fmin_abs.date()),
    max_value=max(snap_max.date(), fmax_abs.date()),
)
if isinstance(date_range, tuple) and len(date_range) == 2:
    d_ini, d_fin = pd.to_datetime(date_range[0]), pd.to_datetime(date_range[1])
else:
    d_ini, d_fin = snap_min, snap_max

mov_f = mov[
    mov["subcuenta"].isin(sel_ports)
    & mov["fecha_op"].between(d_ini, d_fin)
].copy()
pos_f = pos[pos["subcuenta"].isin(sel_ports)].copy()

# ---------------------------------------------------------------------------
# Header KPIs
# ---------------------------------------------------------------------------
st.title("Dashboard | Portafolios VIIN")

# Valor de Portafolio por subcuenta y fecha (suma de valor_mercado_neto en posiciones)
val_port = (
    pos_f.groupby(["fecha", "subcuenta"], as_index=False)["valor_mercado_neto"]
    .sum()
    .sort_values(["subcuenta", "fecha"])
)
val_port["valor_lag"] = val_port.groupby("subcuenta")["valor_mercado_neto"].shift(1)
val_port["mom_pct"] = (val_port["valor_mercado_neto"] / val_port["valor_lag"]) - 1

# KPI cards
kpi_cols = st.columns(len(sel_ports) if sel_ports else 1)
for i, p in enumerate(sel_ports):
    sub = val_port[val_port["subcuenta"] == p].sort_values("fecha")
    if sub.empty:
        with kpi_cols[i]:
            st.metric(p, "s/d")
        continue
    last = sub.iloc[-1]
    first = sub.iloc[0]
    ytd = (last["valor_mercado_neto"] / first["valor_mercado_neto"]) - 1
    with kpi_cols[i]:
        st.metric(
            label=p,
            value=f"${last['valor_mercado_neto']:,.0f}",
            delta=f"{ytd*100:,.2f}% periodo  |  {(last['mom_pct'] or 0)*100:,.2f}% MoM",
        )

st.divider()

# ---------------------------------------------------------------------------
# Tabs
# ---------------------------------------------------------------------------
(
    tab_mtm, tab_carta, tab_mensual,
    tab_compos, tab_oper, tab_data,
) = st.tabs([
    "MTM Diario", "Carta de Posicion", "Mensual",
    "Composicion", "Operaciones", "Datos crudos",
])

# ---- MTM DIARIO -------------------------------------------------------------
with tab_mtm:
    s_date = pd.to_datetime(d_ini).strftime("%Y-%m-%d")
    e_date = pd.to_datetime(d_fin).strftime("%Y-%m-%d")
    daily_mtm, holdings_full, missing = compute_daily_mtm(pos_f, mov_f, s_date, e_date)

    if daily_mtm.empty:
        st.warning("No hay datos suficientes para calcular MTM en el rango seleccionado.")
    else:
        snap_dates_in_range = sorted(
            d for d in pos_f["fecha"].dropna().unique()
            if pd.Timestamp(d) >= pd.Timestamp(s_date)
            and pd.Timestamp(d) <= pd.Timestamp(e_date)
        )
        ports_in_data = [p for p in PORTFOLIOS if p in daily_mtm["subcuenta"].unique()]

        # ---- Total consolidado: suma de los portafolios seleccionados ----
        total_per_day = daily_mtm.groupby("fecha", as_index=False).agg(
            valor_total=("valor_total", "sum"),
            valor_equity=("valor_equity", "sum"),
            valor_carry=("valor_carry", "sum"),
            efectivo=("efectivo", "sum"),
            valor_oficial=("valor_oficial", "sum"),
        ).sort_values("fecha")

        # KPIs consolidados
        last_t = total_per_day.iloc[-1]
        first_t = total_per_day.iloc[0]
        ret_periodo = (
            (last_t["valor_total"] / first_t["valor_total"] - 1)
            if first_t["valor_total"] else 0.0
        )
        delta_abs = last_t["valor_total"] - first_t["valor_total"]

        st.subheader("Valor consolidado del portafolio")
        c1, c2, c3, c4 = st.columns(4)
        c1.metric(
            "Valor total (cierre)",
            f"${last_t['valor_total']:,.0f}",
            delta=f"{ret_periodo*100:+.2f}% en el periodo",
        )
        c2.metric(
            "Cambio en MXN",
            f"${delta_abs:+,.0f}",
        )
        c3.metric(
            "Equities + Renta Fija",
            f"${(last_t['valor_equity']+last_t['valor_carry']):,.0f}",
        )
        c4.metric(
            "Efectivo",
            f"${last_t['efectivo']:,.0f}",
        )

        st.caption(
            "Suma de los portafolios seleccionados en el sidebar. Las bandas "
            "muestran la contribucion de cada portafolio al total. Los "
            "diamantes negros marcan los cierres oficiales (snapshots de Posicion)."
        )

        # ---- Stacked area: total consolidado con contribucion por portafolio ----
        fig_total = go.Figure()
        # Mapeo de colores con alpha
        color_alpha = {
            "VIIN000000000001": "rgba(31,78,121,0.85)",
            "VIIN000000000003": "rgba(217,119,6,0.85)",
            "VIIN000000000006": "rgba(5,150,105,0.85)",
        }
        for p in ports_in_data:
            sub = daily_mtm[daily_mtm["subcuenta"] == p].sort_values("fecha")
            fig_total.add_trace(go.Scatter(
                x=sub["fecha"], y=sub["valor_total"],
                stackgroup="one",
                name=PORT_LABEL.get(p, p),
                line=dict(width=0),
                fillcolor=color_alpha.get(p, "rgba(100,100,100,0.7)"),
                hovertemplate=f"<b>{PORT_LABEL.get(p, p)}</b><br>"
                              f"%{{x|%d %b %Y}}<br>"
                              f"$%{{y:,.0f}}<extra></extra>",
            ))
        # Diamantes en cierres oficiales (sobre el total)
        snaps_total = total_per_day[total_per_day["fecha"].isin(snap_dates_in_range)]
        fig_total.add_trace(go.Scatter(
            x=snaps_total["fecha"], y=snaps_total["valor_total"],
            mode="markers",
            marker=dict(symbol="diamond", size=11, color="#1f2937",
                         line=dict(color="white", width=2)),
            name="Cierre oficial",
            hovertemplate="<b>Cierre %{x|%b %Y}</b><br>"
                          "Total: $%{y:,.0f}<extra></extra>",
        ))
        fig_total.update_layout(
            **PLOTLY_BASE_LAYOUT,
            height=480,
            title=dict(
                text="<b>Valor consolidado diario</b>",
                x=0.01, xanchor="left",
            ),
            legend=dict(orientation="h", yanchor="bottom", y=1.02,
                         xanchor="right", x=1),
            hovermode="x unified",
        )
        fig_total.update_yaxes(tickprefix="$", tickformat=",.2s",
                                gridcolor=GRID, zerolinecolor=GRID)
        fig_total.update_xaxes(gridcolor=GRID, zerolinecolor=GRID)
        st.plotly_chart(fig_total, use_container_width=True)

        # ---- Auditoria de calibracion ----
        st.subheader("Auditoria de calibracion (vs Posicion oficial)")
        snap_check = (
            daily_mtm.dropna(subset=["valor_oficial"])
            [["fecha", "subcuenta", "valor_total_raw",
              "ajuste_calibracion", "valor_total", "valor_oficial"]]
            .sort_values(["subcuenta", "fecha"])
            .reset_index(drop=True)
        )
        snap_check["error_abs"] = snap_check["valor_total"] - snap_check["valor_oficial"]
        snap_check["error_pct"] = (
            snap_check["error_abs"] / snap_check["valor_oficial"] * 100
        )
        st.dataframe(
            snap_check.style.format({
                "valor_total_raw": "${:,.2f}",
                "ajuste_calibracion": "${:,.2f}",
                "valor_total": "${:,.2f}",
                "valor_oficial": "${:,.2f}",
                "error_abs": "${:,.4f}",
                "error_pct": "{:,.6f}%",
            }),
            use_container_width=True, hide_index=True,
        )
        max_err = snap_check["error_pct"].abs().max()
        if max_err < 1e-4:
            st.success(
                f"Calibracion exacta: maximo error en snapshots = "
                f"{max_err:.2e}%. Cada fin de mes hace match con la Posicion oficial."
            )
        else:
            st.warning(f"Maximo error de calibracion: {max_err:.4f}%")

        # ---- Tabla diaria completa ----
        st.subheader("Tabla diaria por portafolio")
        port_pick = st.selectbox(
            "Portafolio",
            sorted(daily_mtm["subcuenta"].unique()),
            key="mtm_port_pick",
        )
        sub = (
            daily_mtm[daily_mtm["subcuenta"] == port_pick]
            .sort_values("fecha")
            .copy()
        )

        # Saldo de efectivo y valor total como series side-by-side
        c1, c2, c3 = st.columns(3)
        c1.metric(
            "Valor portafolio (ultimo dia)",
            f"${sub['valor_total'].iloc[-1]:,.0f}",
            delta=f"{(sub['valor_total'].iloc[-1]/sub['valor_total'].iloc[0]-1)*100:.2f}% periodo",
        )
        c2.metric(
            "Efectivo (ultimo dia)",
            f"${sub['efectivo'].iloc[-1]:,.0f}",
        )
        c3.metric(
            "Equities + Renta fija (ultimo dia)",
            f"${(sub['valor_equity'].iloc[-1] + sub['valor_carry'].iloc[-1]):,.0f}",
        )

        # Tabla diaria completa
        tabla_diaria = sub[[
            "fecha", "valor_equity", "valor_carry", "efectivo",
            "valor_total_raw", "ajuste_calibracion", "valor_total", "valor_oficial",
        ]].copy()
        tabla_diaria["fecha"] = tabla_diaria["fecha"].dt.strftime("%Y-%m-%d")
        st.dataframe(
            tabla_diaria.style.format({
                "valor_equity": "${:,.2f}",
                "valor_carry": "${:,.2f}",
                "efectivo": "${:,.2f}",
                "valor_total_raw": "${:,.2f}",
                "ajuste_calibracion": "${:,.2f}",
                "valor_total": "${:,.2f}",
                "valor_oficial": "${:,.2f}",
            }, na_rep="-"),
            use_container_width=True, hide_index=True, height=400,
        )
        st.download_button(
            f"Descargar tabla diaria {port_pick} (CSV)",
            sub.to_csv(index=False).encode("utf-8"),
            f"mtm_diario_{port_pick}.csv",
            "text/csv",
            key=f"dl_mtm_{port_pick}",
        )
        st.download_button(
            "Descargar tabla diaria - todos los portafolios (CSV)",
            daily_mtm.to_csv(index=False).encode("utf-8"),
            "mtm_diario_todos.csv",
            "text/csv",
            key="dl_mtm_all",
        )

        # ---- Descomposicion stacked por componente ----
        st.subheader(f"Composicion diaria - {PORT_LABEL.get(port_pick, port_pick)}")
        st.caption(
            "Cada banda muestra cuanto contribuye cada clase de activo al valor "
            "total del dia. La banda gris (Ajuste) deberia ser delgada y oscilar "
            "alrededor de cero — refleja diferencias residuales tras la calibracion."
        )
        fig_dec = go.Figure()
        fig_dec.add_trace(go.Scatter(
            x=sub["fecha"], y=sub["valor_equity"],
            stackgroup="one", name="Equities",
            line=dict(width=0), fillcolor="rgba(37,99,235,0.85)",
            hovertemplate="<b>%{x|%d %b %Y}</b><br>Equities: $%{y:,.0f}<extra></extra>",
        ))
        fig_dec.add_trace(go.Scatter(
            x=sub["fecha"], y=sub["valor_carry"],
            stackgroup="one", name="Renta fija / Reporto",
            line=dict(width=0), fillcolor="rgba(124,58,237,0.85)",
            hovertemplate="<b>%{x|%d %b %Y}</b><br>RF/Repo: $%{y:,.0f}<extra></extra>",
        ))
        fig_dec.add_trace(go.Scatter(
            x=sub["fecha"], y=sub["efectivo"],
            stackgroup="one", name="Efectivo",
            line=dict(width=0), fillcolor="rgba(16,185,129,0.85)",
            hovertemplate="<b>%{x|%d %b %Y}</b><br>Cash: $%{y:,.0f}<extra></extra>",
        ))
        fig_dec.add_trace(go.Scatter(
            x=sub["fecha"], y=sub["ajuste_calibracion"],
            stackgroup="one", name="Ajuste",
            line=dict(width=0), fillcolor="rgba(148,163,184,0.7)",
            hovertemplate="<b>%{x|%d %b %Y}</b><br>Ajuste: $%{y:,.0f}<extra></extra>",
        ))
        fig_dec.update_layout(
            **PLOTLY_BASE_LAYOUT,
            height=420,
            title=dict(
                text=f"<b>{PORT_LABEL.get(port_pick, port_pick)} - composicion diaria</b>",
                x=0.01, xanchor="left",
            ),
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
            hovermode="x unified",
        )
        fig_dec.update_yaxes(tickprefix="$", tickformat=",.2s",
                              gridcolor=GRID, zerolinecolor=GRID)
        fig_dec.update_xaxes(gridcolor=GRID, zerolinecolor=GRID)
        st.plotly_chart(fig_dec, use_container_width=True)

        # ---- Rendimiento diario % ----
        st.subheader("Rendimiento diario por portafolio")
        st.caption(
            "Variacion porcentual del valor del portafolio respecto al dia "
            "habil anterior. Solo se muestran dias con cambio real (omitidos "
            "los dias planos por forward-fill). La linea negra es el promedio "
            "movil de 5 dias para suavizar el ruido."
        )

        ret = daily_mtm.sort_values(["subcuenta", "fecha"]).copy()
        ret["valor_lag"] = ret.groupby("subcuenta")["valor_total"].shift(1)
        ret["ret_pct"] = ret["valor_total"] / ret["valor_lag"] - 1
        ret = ret.dropna(subset=["ret_pct"]).copy()
        # Filtrar dias planos (cambio < 0.001%) que vienen de forward-fill
        ret = ret[ret["ret_pct"].abs() > 1e-5].copy()
        # Promedio movil de 5 dias por portafolio
        ret["ret_ma5"] = (
            ret.sort_values(["subcuenta", "fecha"])
            .groupby("subcuenta")["ret_pct"]
            .rolling(window=5, min_periods=1).mean()
            .reset_index(level=0, drop=True)
        )

        fig_ret = make_subplots(
            rows=len(ports_in_data), cols=1, shared_xaxes=True,
            subplot_titles=[
                f"{PORT_LABEL.get(p, p)}  -  promedio: "
                f"{ret[ret['subcuenta']==p]['ret_pct'].mean()*100:+.3f}%/dia, "
                f"mejor: {ret[ret['subcuenta']==p]['ret_pct'].max()*100:+.2f}%, "
                f"peor: {ret[ret['subcuenta']==p]['ret_pct'].min()*100:+.2f}%"
                for p in ports_in_data
            ],
            vertical_spacing=0.10,
        )
        for i, p in enumerate(ports_in_data, 1):
            r = ret[ret["subcuenta"] == p].sort_values("fecha")
            colors = [GREEN if v >= 0 else RED for v in r["ret_pct"]]
            fig_ret.add_trace(go.Bar(
                x=r["fecha"], y=r["ret_pct"],
                marker=dict(color=colors, line=dict(width=0), opacity=0.6),
                showlegend=False,
                hovertemplate="<b>%{x|%d %b %Y}</b><br>%{y:.3%}<extra></extra>",
            ), row=i, col=1)
            # Linea de promedio movil 5 dias
            fig_ret.add_trace(go.Scatter(
                x=r["fecha"], y=r["ret_ma5"],
                mode="lines",
                line=dict(color="#1f2937", width=1.5),
                showlegend=(i == 1),
                name="MA 5 dias" if i == 1 else None,
                hovertemplate="<b>%{x|%d %b %Y}</b><br>MA5: %{y:.3%}<extra></extra>",
            ), row=i, col=1)
            fig_ret.add_hline(y=0, line=dict(color="#9ca3af", width=1), row=i, col=1)
            fig_ret.update_yaxes(tickformat=".2%", gridcolor=GRID,
                                  zerolinecolor="#9ca3af", row=i, col=1)
            fig_ret.update_xaxes(gridcolor=GRID, row=i, col=1)
        fig_ret.update_layout(
            **PLOTLY_BASE_LAYOUT,
            height=200 * len(ports_in_data) + 100,
            title=dict(
                text="<b>Rendimiento diario %</b>",
                x=0.01, xanchor="left",
            ),
            bargap=0.05,
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        )
        st.plotly_chart(fig_ret, use_container_width=True)

        # ---- Rendimiento acumulado % (reemplaza Indice base 100) ----
        st.subheader("Rendimiento acumulado en el periodo")
        st.caption(
            "Crecimiento porcentual del valor de cada portafolio desde el "
            "primer dia. Si una linea esta en +5% significa que el portafolio "
            "vale 5% mas que al inicio del periodo. Permite comparar el "
            "desempeño relativo entre portafolios sin importar su tamaño."
        )

        cum_ret = daily_mtm.sort_values(["subcuenta", "fecha"]).copy()
        first_vals = cum_ret.groupby("subcuenta")["valor_total"].transform("first")
        cum_ret["ret_acum"] = cum_ret["valor_total"] / first_vals - 1

        # Tambien para el consolidado
        first_total = total_per_day["valor_total"].iloc[0]
        total_per_day["ret_acum"] = total_per_day["valor_total"] / first_total - 1

        fig_cum = go.Figure()
        # Linea de consolidado en negro grueso
        fig_cum.add_trace(go.Scatter(
            x=total_per_day["fecha"], y=total_per_day["ret_acum"],
            mode="lines",
            line=dict(color="#1f2937", width=3),
            name="Consolidado",
            hovertemplate="<b>Consolidado</b><br>%{x|%d %b %Y}<br>"
                           "Retorno: %{y:.2%}<extra></extra>",
        ))
        # Lineas por portafolio
        for p in ports_in_data:
            sub_c = cum_ret[cum_ret["subcuenta"] == p].sort_values("fecha")
            color = PORT_COLORS.get(p, "#374151")
            fig_cum.add_trace(go.Scatter(
                x=sub_c["fecha"], y=sub_c["ret_acum"],
                mode="lines",
                line=dict(color=color, width=2),
                name=PORT_LABEL.get(p, p),
                hovertemplate=f"<b>{PORT_LABEL.get(p, p)}</b><br>%{{x|%d %b %Y}}<br>"
                               f"Retorno: %{{y:.2%}}<extra></extra>",
            ))
        fig_cum.add_hline(y=0, line=dict(color="#9ca3af", width=1, dash="dash"))
        fig_cum.update_layout(
            **PLOTLY_BASE_LAYOUT,
            height=440,
            title=dict(text="<b>Retorno acumulado %</b>", x=0.01),
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
            hovermode="x unified",
        )
        fig_cum.update_yaxes(tickformat=".1%", gridcolor=GRID,
                              zerolinecolor="#9ca3af")
        fig_cum.update_xaxes(gridcolor=GRID)
        st.plotly_chart(fig_cum, use_container_width=True)

        # ---- Resumen del periodo ----
        st.subheader("Resumen del periodo")
        res = []
        for port, g in daily_mtm.groupby("subcuenta"):
            g = g.sort_values("fecha")
            v0, vN = g["valor_total"].iloc[0], g["valor_total"].iloc[-1]
            ret_total = vN / v0 - 1 if v0 else np.nan
            ddaily = (g["valor_total"] / g["valor_total"].shift(1) - 1).dropna()
            vol = ddaily.std() * np.sqrt(252) if len(ddaily) > 1 else np.nan
            sharpe = (
                (ddaily.mean() * 252) / (ddaily.std() * np.sqrt(252))
                if ddaily.std() else np.nan
            )
            mdd = ((g["valor_total"] / g["valor_total"].cummax()) - 1).min()
            res.append({
                "Portafolio": port,
                "Valor inicial": v0,
                "Valor final": vN,
                "Retorno periodo %": ret_total * 100,
                "Vol anualizada %": vol * 100 if pd.notna(vol) else np.nan,
                "Sharpe (rf=0)": sharpe,
                "Max drawdown %": mdd * 100,
            })
        st.dataframe(
            pd.DataFrame(res).round(2),
            use_container_width=True, hide_index=True,
        )

        if missing:
            st.warning(
                f"Tickers sin data en yfinance (cayeron a carry): "
                f"{', '.join(missing)}"
            )

        with st.expander("Holdings + precios (debug)"):
            st.dataframe(
                holdings_full.sort_values(["subcuenta", "fecha", "emisora"]).head(3000),
                use_container_width=True, hide_index=True,
            )


# ---- CARTA DE POSICION ------------------------------------------------------
with tab_carta:
    st.subheader("Carta de Posicion - estructura identica a la hoja Posicion")
    st.caption(
        "Selecciona portafolio y fecha para reconstruir la hoja Posicion del "
        "Excel a esa fecha. En cierres de mes coincide con el archivo oficial. "
        "En fechas intermedias, los precios son de yfinance (.MX, MXN); "
        "renta fija/reporto se valua por interpolacion del importe_neto entre "
        "snapshots; estatica (cupon, plazo, tasa, estrategia) se hereda del "
        "snapshot mas reciente."
    )

    if daily_mtm.empty or holdings_full.empty:
        st.warning("Calcula MTM primero (pestana MTM Diario).")
    else:
        c1, c2 = st.columns([1, 1])
        port_carta = c1.selectbox(
            "Portafolio",
            options=PORTFOLIOS,
            key="carta_portfolio",
        )

        # Rango de fechas disponible
        avail_dates = pd.DatetimeIndex(
            sorted(daily_mtm["fecha"].dropna().unique())
        )
        min_d = avail_dates.min().date() if len(avail_dates) else pd.Timestamp(d_ini).date()
        max_d = avail_dates.max().date() if len(avail_dates) else pd.Timestamp(d_fin).date()
        target_d = c2.date_input(
            "Fecha de la carta",
            value=max_d,
            min_value=min_d,
            max_value=max_d,
            key="carta_date",
        )
        target_ts = pd.Timestamp(target_d)

        # Si la fecha pedida no esta en el calendario habil, buscar la mas
        # cercana (ultimo dia habil <=)
        if target_ts not in avail_dates:
            le = avail_dates[avail_dates <= target_ts]
            if len(le):
                actual_ts = le.max()
                st.info(
                    f"La fecha {target_ts:%Y-%m-%d} no es dia habil. Mostrando "
                    f"el cierre mas reciente disponible: {actual_ts:%Y-%m-%d}."
                )
                target_ts = actual_ts
            else:
                st.error("No hay datos para esa fecha.")
                target_ts = None

        if target_ts is not None:
            df_carta, totals = reconstruct_position(
                port_carta, target_ts, pos, holdings_full, daily_mtm
            )

            # Header estilo Excel
            es_snap = totals["es_snapshot"]
            st.markdown(
                f"""
**Contrato:** 105433  &nbsp;&nbsp;|&nbsp;&nbsp; **Cliente:** SEGUROS AZTECA  &nbsp;&nbsp;|&nbsp;&nbsp; **No. Cliente:** 1582
**Promotor:** Miguel Angel Tebar Pedroza  &nbsp;&nbsp;|&nbsp;&nbsp; **Divisa:** MXP  &nbsp;&nbsp;|&nbsp;&nbsp; **Fecha:** {target_ts:%Y-%m-%d}
**Producto:** Mercado de Capitales  &nbsp;&nbsp;|&nbsp;&nbsp; **Regional:** Mexico  &nbsp;&nbsp;|&nbsp;&nbsp; **Sucursal:** Corporativo  &nbsp;&nbsp;|&nbsp;&nbsp; **Subcuenta:** {port_carta}
"""
            )
            if es_snap:
                st.success(
                    "Esta fecha es un cierre de mes oficial. La carta coincide "
                    "con el archivo Posicion del LayOut correspondiente."
                )

            # Tabla principal con formato
            if df_carta.empty:
                st.warning("No hay posiciones vivas en esta fecha.")
            else:
                fmt = {
                    "Cupon": "{:.2f}",
                    "Plazo": "{:.0f}",
                    "Tasa": "{:.4f}",
                    "Dias x Ven": "{:.0f}",
                    "Titulos": "{:,.0f}",
                    "Precio": "{:,.6f}",
                    "Importe Bruto": "${:,.2f}",
                    "Precio Neto": "{:,.6f}",
                    "Importe Neto": "${:,.2f}",
                    "Precio de Mercado": "{:,.6f}",
                    "Valor de Mercado Neto": "${:,.2f}",
                    "Plus/Minus + Int. Dev.": "${:,.2f}",
                    "Plus/Minus %": "{:.2%}",
                    "% de Cartera": "{:.4%}",
                }
                st.dataframe(
                    df_carta.style.format(fmt, na_rep="-"),
                    use_container_width=True, hide_index=True,
                )

                # Footer con totales (replica las filas finales del Excel)
                st.markdown("---")
                ft = pd.DataFrame([
                    {"Concepto": "Mdo.Dinero",
                     "Valor": totals["valor_carry"]},
                    {"Concepto": "Valor del Portafolio",
                     "Valor": totals["valor_portafolio"]},
                    {"Concepto": "Ef. Disponible",
                     "Valor": totals["efectivo"]},
                    {"Concepto": "Sdo. Disp. para Inv.*",
                     "Valor": totals["efectivo"]},
                    {"Concepto": "Sdo. Pend. MC", "Valor": 0.0},
                ])
                st.dataframe(
                    ft.style.format({"Valor": "${:,.2f}"}),
                    use_container_width=True, hide_index=True,
                )

                # Resumen rapido en metricas
                c1, c2, c3, c4 = st.columns(4)
                c1.metric(
                    "Valor del Portafolio",
                    f"${totals['valor_portafolio']:,.2f}",
                )
                c2.metric("Equities (MTM)", f"${totals['valor_equity']:,.2f}")
                c3.metric(
                    "Renta fija / Reporto",
                    f"${totals['valor_carry']:,.2f}",
                )
                c4.metric("Efectivo", f"${totals['efectivo']:,.2f}")

                # Descargas
                dl_c1, dl_c2 = st.columns(2)
                with dl_c1:
                    st.download_button(
                        f"Carta {port_carta} - {target_ts:%Y%m%d} (CSV)",
                        df_carta.to_csv(index=False).encode("utf-8"),
                        f"carta_posicion_{port_carta}_{target_ts:%Y%m%d}.csv",
                        "text/csv",
                        key=f"dl_carta_{port_carta}_{target_ts:%Y%m%d}",
                    )

                # ---- Export historico valor del portafolio (Excel) ----
                st.divider()
                st.subheader("Exportar historico del valor del portafolio")
                st.caption(
                    "Genera un archivo Excel con el historico diario del "
                    "valor del portafolio para todo el periodo. Hojas: "
                    "Resumen consolidado (una columna por portafolio + total), "
                    "Rendimiento acumulado %, y una hoja por portafolio con "
                    "el desglose diario (equity, renta fija, efectivo, total)."
                )

                excel_bytes = build_historico_excel(daily_mtm)
                rng_label = (
                    f"{pd.Timestamp(daily_mtm['fecha'].min()):%Y%m%d}_"
                    f"{pd.Timestamp(daily_mtm['fecha'].max()):%Y%m%d}"
                )
                st.download_button(
                    "Descargar historico del portafolio (Excel)",
                    excel_bytes,
                    f"historico_valor_portafolio_{rng_label}.xlsx",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"dl_hist_excel_{target_ts:%Y%m%d}",
                    type="primary",
                )



# ---- MENSUAL ----------------------------------------------------------------
with tab_mensual:
    ports_in_val = [p for p in PORTFOLIOS if p in val_port["subcuenta"].unique()]

    # ---- Valor del Portafolio cierre de mes (3 paneles, escala independiente) ----
    st.subheader("Valor del Portafolio en cada cierre de mes")
    st.caption(
        "Saldo oficial al ultimo dia de cada mes (de la hoja Posicion). "
        "Las anotaciones muestran el valor en cada punto."
    )
    fig_v = make_subplots(
        rows=len(ports_in_val), cols=1, shared_xaxes=True,
        subplot_titles=[PORT_LABEL.get(p, p) for p in ports_in_val],
        vertical_spacing=0.07,
    )
    for i, p in enumerate(ports_in_val, 1):
        sub_v = val_port[val_port["subcuenta"] == p].sort_values("fecha")
        color = PORT_COLORS.get(p, "#374151")
        # Linea
        fig_v.add_trace(go.Scatter(
            x=sub_v["fecha"], y=sub_v["valor_mercado_neto"],
            mode="lines+markers",
            line=dict(color=color, width=2.5),
            marker=dict(size=8, color=color, line=dict(color="white", width=1.5)),
            showlegend=False,
            hovertemplate="<b>%{x|%b %Y}</b><br>$%{y:,.0f}<extra></extra>",
        ), row=i, col=1)
        fig_v.update_yaxes(tickprefix="$", tickformat=",.2s",
                            gridcolor=GRID, zerolinecolor=GRID, row=i, col=1)
        fig_v.update_xaxes(gridcolor=GRID, row=i, col=1, dtick="M1")
    fig_v.update_layout(
        **PLOTLY_BASE_LAYOUT,
        height=200 * len(ports_in_val) + 80,
        title=dict(text="<b>Valor de Mercado Neto - cierres mensuales</b>",
                    x=0.01, xanchor="left"),
        hovermode="x unified",
    )
    st.plotly_chart(fig_v, use_container_width=True)

    # ---- Rendimiento mensual MoM ----
    st.subheader("Rendimiento mensual (MoM %)")
    st.caption(
        "Cambio porcentual mes contra mes basado en Valor de Mercado Neto. "
        "Verde = ganancia, rojo = perdida."
    )
    mom_data = val_port.dropna(subset=["mom_pct"]).copy()
    fig_mom = make_subplots(
        rows=len(ports_in_val), cols=1, shared_xaxes=True,
        subplot_titles=[PORT_LABEL.get(p, p) for p in ports_in_val],
        vertical_spacing=0.08,
    )
    for i, p in enumerate(ports_in_val, 1):
        d = mom_data[mom_data["subcuenta"] == p].sort_values("fecha")
        colors = [GREEN if v >= 0 else RED for v in d["mom_pct"]]
        fig_mom.add_trace(go.Bar(
            x=d["fecha"], y=d["mom_pct"],
            marker=dict(color=colors, line=dict(width=0)),
            text=[f"{v*100:+.1f}%" for v in d["mom_pct"]],
            textposition="outside",
            textfont=dict(size=10),
            showlegend=False,
            hovertemplate="<b>%{x|%b %Y}</b><br>%{y:.2%}<extra></extra>",
        ), row=i, col=1)
        fig_mom.add_hline(y=0, line=dict(color="#9ca3af", width=1), row=i, col=1)
        fig_mom.update_yaxes(tickformat=".1%", gridcolor=GRID,
                              zerolinecolor="#9ca3af", row=i, col=1)
        fig_mom.update_xaxes(gridcolor=GRID, row=i, col=1, dtick="M1")
    fig_mom.update_layout(
        **PLOTLY_BASE_LAYOUT,
        height=200 * len(ports_in_val) + 80,
        title=dict(text="<b>Rendimiento mes contra mes</b>", x=0.01, xanchor="left"),
        bargap=0.2,
    )
    st.plotly_chart(fig_mom, use_container_width=True)

    # ---- Plus/Minus en cierres de mes ----
    st.subheader("Plusvalia / Minusvalia + Intereses Devengados")
    st.caption(
        "Suma de la columna `Plus/Minus + Int. Dev.` en la hoja Posicion. "
        "Es la utilidad/perdida no realizada acumulada al cierre del mes."
    )
    pm = (
        pos_f.groupby(["fecha", "subcuenta"], as_index=False)["plus_minus_int"]
        .sum()
    )
    fig_pm = make_subplots(
        rows=len(ports_in_val), cols=1, shared_xaxes=True,
        subplot_titles=[PORT_LABEL.get(p, p) for p in ports_in_val],
        vertical_spacing=0.08,
    )
    for i, p in enumerate(ports_in_val, 1):
        d = pm[pm["subcuenta"] == p].sort_values("fecha")
        colors = [GREEN if v >= 0 else RED for v in d["plus_minus_int"]]
        fig_pm.add_trace(go.Bar(
            x=d["fecha"], y=d["plus_minus_int"],
            marker=dict(color=colors, line=dict(width=0)),
            showlegend=False,
            hovertemplate="<b>%{x|%b %Y}</b><br>$%{y:,.0f}<extra></extra>",
        ), row=i, col=1)
        fig_pm.add_hline(y=0, line=dict(color="#9ca3af", width=1), row=i, col=1)
        fig_pm.update_yaxes(tickprefix="$", tickformat=",.2s",
                             gridcolor=GRID, zerolinecolor="#9ca3af", row=i, col=1)
        fig_pm.update_xaxes(gridcolor=GRID, row=i, col=1, dtick="M1")
    fig_pm.update_layout(
        **PLOTLY_BASE_LAYOUT,
        height=200 * len(ports_in_val) + 80,
        title=dict(text="<b>Plus/Minus + Intereses devengados</b>",
                    x=0.01, xanchor="left"),
        bargap=0.2,
    )
    st.plotly_chart(fig_pm, use_container_width=True)


# ---- COMPOSICION ------------------------------------------------------------
with tab_compos:
    last_date = pos_f["fecha"].max()
    last_pos = pos_f[pos_f["fecha"] == last_date]
    ports_in_pos = [p for p in PORTFOLIOS if p in last_pos["subcuenta"].unique()]

    # ---- Composicion por Estrategia ----
    st.subheader(f"Composicion por activos al {last_date:%d %b %Y}")
    st.caption(
        "Treemap: el area de cada caja es proporcional al valor de mercado del "
        "activo. Niveles: Portafolio -> Tipo (CO=Capital MX, SC=Capital US "
        "cross-listed, R=Reporto, D=Deuda) -> Emisora especifica. "
        "Los porcentajes son respecto al total del portafolio."
    )
    # Mapping de TP a labels mas legibles
    tp_labels = {
        "CO": "Capital MX",
        "SC": "Capital US (SIC)",
        "R": "Reporto / Money Market",
        "D": "Deuda Corporativa",
    }
    by_asset = (
        last_pos.groupby(["subcuenta", "tp", "emisora", "serie"],
                          as_index=False)["valor_mercado_neto"].sum()
    )
    by_asset["pct"] = by_asset.groupby("subcuenta")["valor_mercado_neto"].transform(
        lambda x: x / x.sum()
    )
    by_asset["valor_lbl"] = by_asset["valor_mercado_neto"].apply(fmt_money)
    by_asset["tipo"] = by_asset["tp"].map(tp_labels).fillna(by_asset["tp"])
    by_asset["instrumento"] = (
        by_asset["emisora"].astype(str) + " " + by_asset["serie"].astype(str)
    )

    fig_tree = px.treemap(
        by_asset,
        path=[px.Constant("Total"), "subcuenta", "tipo", "instrumento"],
        values="valor_mercado_neto",
        color="subcuenta",
        color_discrete_map={p: PORT_COLORS.get(p, "#374151") for p in PORTFOLIOS},
        custom_data=["pct", "valor_lbl"],
    )
    fig_tree.update_traces(
        texttemplate="<b>%{label}</b><br>%{customdata[1]}<br>%{percentParent:.1%}",
        textfont=dict(size=11),
        hovertemplate="<b>%{label}</b><br>"
                       "Valor: %{customdata[1]}<br>"
                       "% del padre: %{percentParent:.2%}<br>"
                       "% del total: %{percentRoot:.2%}<extra></extra>",
    )
    fig_tree.update_layout(
        **{**PLOTLY_BASE_LAYOUT, "margin": dict(l=10, r=10, t=60, b=10)},
        height=600,
        title=dict(text="<b>Distribucion por activo</b>", x=0.01),
    )
    st.plotly_chart(fig_tree, use_container_width=True)

    # ---- Distribucion por tipo de activo (resumen agregado) ----
    st.subheader("Asignacion por clase de activo")
    st.caption(
        "Peso porcentual de cada tipo de activo dentro de cada portafolio. "
        "Util para ver de un vistazo cuanto esta en capitales vs renta fija "
        "vs reporto."
    )
    by_tp = (
        last_pos.groupby(["subcuenta", "tp"], as_index=False)["valor_mercado_neto"].sum()
    )
    by_tp["tipo"] = by_tp["tp"].map(tp_labels).fillna(by_tp["tp"])
    by_tp["pct"] = by_tp.groupby("subcuenta")["valor_mercado_neto"].transform(
        lambda x: x / x.sum()
    )
    by_tp["subcuenta_lbl"] = by_tp["subcuenta"].map(PORT_LABEL).fillna(by_tp["subcuenta"])

    tp_color = {
        "Capital MX": "#1f4e79",
        "Capital US (SIC)": "#2563eb",
        "Reporto / Money Market": "#7c3aed",
        "Deuda Corporativa": "#d97706",
    }
    fig_alloc = px.bar(
        by_tp.sort_values(["subcuenta", "tp"]),
        x="pct", y="subcuenta_lbl",
        color="tipo", orientation="h",
        color_discrete_map=tp_color,
        custom_data=["valor_mercado_neto", "tipo"],
    )
    fig_alloc.update_traces(
        hovertemplate="<b>%{y}</b><br>%{customdata[1]}<br>"
                       "%{x:.1%} del portafolio<br>"
                       "$%{customdata[0]:,.0f}<extra></extra>",
        texttemplate="%{x:.1%}",
        textposition="inside",
        textfont=dict(size=11, color="white"),
    )
    fig_alloc.update_layout(
        **PLOTLY_BASE_LAYOUT,
        height=320,
        title=dict(text="<b>Asignacion % por tipo de activo</b>", x=0.01),
        barmode="stack",
        legend=dict(orientation="h", yanchor="bottom", y=1.02,
                     xanchor="right", x=1, title=""),
    )
    fig_alloc.update_xaxes(tickformat=".0%", gridcolor=GRID, range=[0, 1])
    fig_alloc.update_yaxes(gridcolor=GRID)
    st.plotly_chart(fig_alloc, use_container_width=True)

    # ---- Top emisoras por portafolio ----
    st.subheader(f"Top 10 emisoras por portafolio")
    st.caption(
        "Las 10 posiciones mas grandes por valor de mercado en cada portafolio. "
        "Una etiqueta arriba de cada barra muestra el peso en cartera."
    )
    top = (
        last_pos.groupby(["subcuenta", "emisora", "serie"], as_index=False)
        ["valor_mercado_neto"].sum()
    )
    top["instrumento"] = top["emisora"].astype(str) + " " + top["serie"].astype(str)

    fig_top = make_subplots(
        rows=1, cols=len(ports_in_pos),
        subplot_titles=[PORT_LABEL.get(p, p) for p in ports_in_pos],
        horizontal_spacing=0.12,
    )
    for j, p in enumerate(ports_in_pos, 1):
        d = top[top["subcuenta"] == p].sort_values(
            "valor_mercado_neto", ascending=True
        ).tail(10)
        if d.empty:
            continue
        total_p = top[top["subcuenta"] == p]["valor_mercado_neto"].sum()
        d["pct"] = d["valor_mercado_neto"] / total_p
        color = PORT_COLORS.get(p, "#374151")
        fig_top.add_trace(go.Bar(
            x=d["valor_mercado_neto"], y=d["instrumento"],
            orientation="h",
            marker=dict(color=color, line=dict(width=0)),
            text=[f"{v*100:.1f}%" for v in d["pct"]],
            textposition="outside",
            textfont=dict(size=10),
            showlegend=False,
            hovertemplate="<b>%{y}</b><br>$%{x:,.0f}<br>"
                           "Peso: %{text}<extra></extra>",
        ), row=1, col=j)
        fig_top.update_xaxes(tickprefix="$", tickformat=",.2s",
                              gridcolor=GRID, row=1, col=j)
        fig_top.update_yaxes(gridcolor=GRID, row=1, col=j)
    fig_top.update_layout(
        **PLOTLY_BASE_LAYOUT,
        height=520,
        title=dict(text=f"<b>Top emisoras al {last_date:%b %Y}</b>", x=0.01),
        bargap=0.25,
    )
    st.plotly_chart(fig_top, use_container_width=True)

    # ---- Scatter: tamaño vs rendimiento ----
    st.subheader("Tamano de posicion vs rendimiento (Plus/Minus %)")
    st.caption(
        "Cada punto es una emisora. Eje X = valor de mercado (escala log), "
        "Eje Y = Plus/Minus %. Posiciones grandes a la derecha; arriba del "
        "0% estan ganando, abajo estan perdiendo. Color por tipo de activo."
    )
    pm_data = last_pos.copy()
    pm_data = pm_data.dropna(subset=["plus_minus_pct"])
    pm_data = pm_data[pm_data["valor_mercado_neto"] > 0]
    pm_data["tipo"] = pm_data["tp"].map(tp_labels).fillna(pm_data["tp"])

    fig_sc = make_subplots(
        rows=1, cols=len(ports_in_pos),
        subplot_titles=[PORT_LABEL.get(p, p) for p in ports_in_pos],
        horizontal_spacing=0.08,
    )
    for j, p in enumerate(ports_in_pos, 1):
        d = pm_data[pm_data["subcuenta"] == p]
        if d.empty:
            continue
        for tipo_lbl, sub_e in d.groupby("tipo"):
            color = tp_color.get(tipo_lbl, "#374151")
            fig_sc.add_trace(go.Scatter(
                x=sub_e["valor_mercado_neto"], y=sub_e["plus_minus_pct"],
                mode="markers",
                marker=dict(size=11, color=color, opacity=0.8,
                             line=dict(color="white", width=1)),
                name=tipo_lbl, legendgroup=tipo_lbl,
                showlegend=(j == 1),
                customdata=np.stack(
                    [sub_e["emisora"], sub_e["serie"], sub_e["titulos"]], axis=1
                ),
                hovertemplate="<b>%{customdata[0]} %{customdata[1]}</b><br>"
                               "Valor: $%{x:,.0f}<br>"
                               "Plus/Minus: %{y:.2%}<br>"
                               "Titulos: %{customdata[2]:,.0f}"
                               "<extra></extra>",
            ), row=1, col=j)
        fig_sc.add_hline(y=0, line=dict(color="#9ca3af", width=1, dash="dash"),
                          row=1, col=j)
        fig_sc.update_xaxes(type="log", tickprefix="$", tickformat=",.2s",
                             gridcolor=GRID, row=1, col=j)
        fig_sc.update_yaxes(tickformat=".0%", gridcolor=GRID,
                             zerolinecolor="#9ca3af", row=1, col=j)
    fig_sc.update_layout(
        **PLOTLY_BASE_LAYOUT,
        height=460,
        title=dict(text=f"<b>Tamano vs rendimiento al {last_date:%b %Y}</b>", x=0.01),
        legend=dict(orientation="h", yanchor="bottom", y=1.05, xanchor="right", x=1),
    )
    st.plotly_chart(fig_sc, use_container_width=True)


# ---- OPERACIONES ------------------------------------------------------------
with tab_oper:
    ports_in_oper = [p for p in PORTFOLIOS if p in mov_f["subcuenta"].unique()]

    # ---- # operaciones agrupadas mensualmente ----
    st.subheader("Numero de operaciones por mes")
    st.caption(
        "Cuantas operaciones (de cualquier tipo) se ejecutaron cada mes en "
        "cada portafolio. Util para identificar meses de mayor actividad."
    )
    mov_f_mes = mov_f.dropna(subset=["fecha_op"]).copy()
    mov_f_mes["mes"] = mov_f_mes["fecha_op"].dt.to_period("M").dt.to_timestamp()
    op_monthly = (
        mov_f_mes.groupby(["subcuenta", "mes"], as_index=False)
        .size().rename(columns={"size": "n_ops"})
    )

    fig_op = go.Figure()
    for p in ports_in_oper:
        d = op_monthly[op_monthly["subcuenta"] == p].sort_values("mes")
        color = PORT_COLORS.get(p, "#374151")
        fig_op.add_trace(go.Bar(
            x=d["mes"], y=d["n_ops"],
            name=PORT_LABEL.get(p, p),
            marker=dict(color=color, line=dict(width=0)),
            hovertemplate=f"<b>{PORT_LABEL.get(p, p)}</b><br>%{{x|%b %Y}}<br>"
                           f"%{{y:,.0f}} operaciones<extra></extra>",
        ))
    fig_op.update_layout(
        **PLOTLY_BASE_LAYOUT,
        height=400,
        title=dict(text="<b>Operaciones por mes</b>", x=0.01),
        barmode="group",
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        bargap=0.2,
    )
    fig_op.update_xaxes(gridcolor=GRID, dtick="M1")
    fig_op.update_yaxes(gridcolor=GRID, zerolinecolor=GRID, tickformat=",.0f")
    st.plotly_chart(fig_op, use_container_width=True)

    # ---- Mix por concepto ----
    st.subheader("Volumen operado por tipo de concepto")
    st.caption(
        "Suma de montos netos por concepto de operacion. Los reportos dominan "
        "el volumen porque rotan diariamente. Las compras/ventas de capital y "
        "dividendos dan el detalle del movimiento real del portafolio."
    )
    mix = (
        mov_f.groupby(["subcuenta", "concepto"], as_index=False)
        .agg(
            monto_neto=("monto_neto", "sum"),
            n_ops=("monto_neto", "count"),
        )
    )
    mix["monto_abs"] = mix["monto_neto"].abs()
    # Top 10 conceptos en valor
    top_conc = (
        mix.groupby("concepto")["monto_abs"].sum().nlargest(10).index.tolist()
    )
    mix_f = mix[mix["concepto"].isin(top_conc)].copy()
    mix_f["subcuenta_lbl"] = mix_f["subcuenta"].map(PORT_LABEL).fillna(mix_f["subcuenta"])

    fig_mix = px.bar(
        mix_f, x="concepto", y="monto_abs", color="subcuenta",
        color_discrete_map={p: PORT_COLORS.get(p, "#374151") for p in PORTFOLIOS},
        barmode="group", custom_data=["n_ops", "subcuenta_lbl"],
    )
    fig_mix.update_traces(
        hovertemplate="<b>%{customdata[1]}</b><br>%{x}<br>"
                       "Monto: $%{y:,.0f}<br>"
                       "# ops: %{customdata[0]:,.0f}<extra></extra>",
    )
    fig_mix.update_layout(
        **PLOTLY_BASE_LAYOUT,
        height=480,
        title=dict(text="<b>Volumen acumulado por concepto (MXN)</b>", x=0.01),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1, title=""),
        xaxis=dict(tickangle=-30),
        bargap=0.25,
    )
    fig_mix.update_yaxes(tickprefix="$", tickformat=",.2s",
                          gridcolor=GRID, zerolinecolor=GRID)
    fig_mix.update_xaxes(gridcolor=GRID)
    st.plotly_chart(fig_mix, use_container_width=True)

    # ---- Tasas de reporto promedio ponderadas ----
    st.subheader("Tasa de reporto contratada en el tiempo")
    st.caption(
        "Tasa anual % a la que se contrato cada repo. Linea = promedio "
        "ponderado por monto, mes a mes. Refleja el nivel de tasas de fondeo "
        "(TIIE/CETES) que esta capturando el portafolio."
    )
    rep = mov_f[mov_f["concepto"] == "INICIO CPA REPORTO"].dropna(
        subset=["tasa_premio", "fecha_op"]
    ).copy()
    rep["monto_abs"] = rep["monto_neto"].abs()
    rep = rep[rep["monto_abs"] > 0]
    if not rep.empty:
        # Promedio ponderado mensual
        rep["mes"] = rep["fecha_op"].dt.to_period("M").dt.to_timestamp()
        wavg = (
            rep.groupby(["subcuenta", "mes"])
            .apply(lambda g: (g["tasa_premio"] * g["monto_abs"]).sum() / g["monto_abs"].sum())
            .reset_index(name="tasa_wavg")
        )

        fig_tas = go.Figure()
        # Scatter de fondo (cada repo individual)
        for p in ports_in_oper:
            d = rep[rep["subcuenta"] == p]
            color = PORT_COLORS.get(p, "#374151")
            fig_tas.add_trace(go.Scatter(
                x=d["fecha_op"], y=d["tasa_premio"],
                mode="markers",
                marker=dict(size=5, color=color, opacity=0.25,
                             line=dict(width=0)),
                name=f"{PORT_LABEL.get(p, p)} (operaciones)",
                showlegend=False,
                hovertemplate=f"<b>{PORT_LABEL.get(p, p)}</b><br>%{{x|%d %b %Y}}<br>"
                               f"Tasa: %{{y:.2f}}%<extra></extra>",
            ))
        # Lineas de promedio ponderado
        for p in ports_in_oper:
            d = wavg[wavg["subcuenta"] == p].sort_values("mes")
            color = PORT_COLORS.get(p, "#374151")
            fig_tas.add_trace(go.Scatter(
                x=d["mes"], y=d["tasa_wavg"],
                mode="lines+markers",
                line=dict(color=color, width=2.5),
                marker=dict(size=8, color=color, line=dict(color="white", width=1.5)),
                name=PORT_LABEL.get(p, p),
                hovertemplate=f"<b>{PORT_LABEL.get(p, p)} - promedio ponderado</b><br>"
                               f"%{{x|%b %Y}}<br>"
                               f"Tasa: %{{y:.2f}}%<extra></extra>",
            ))
        fig_tas.update_layout(
            **PLOTLY_BASE_LAYOUT,
            height=460,
            title=dict(
                text="<b>Tasas de reporto - operaciones individuales y promedio ponderado mensual</b>",
                x=0.01,
            ),
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        )
        fig_tas.update_yaxes(ticksuffix="%", gridcolor=GRID, zerolinecolor=GRID)
        fig_tas.update_xaxes(gridcolor=GRID)
        st.plotly_chart(fig_tas, use_container_width=True)
    else:
        st.info("No hay operaciones de reporto en el rango seleccionado.")


# ---- DATOS CRUDOS -----------------------------------------------------------
with tab_data:
    st.subheader("Movimientos consolidados")
    st.dataframe(mov_f, use_container_width=True, hide_index=True)
    st.download_button(
        "Descargar movimientos consolidados (CSV)",
        mov_f.to_csv(index=False).encode("utf-8"),
        "movimientos_consolidados.csv",
        "text/csv",
    )

    st.divider()
    st.subheader("Posiciones por fecha de cierre de mes")
    st.caption(
        "Una tabla por cada archivo LayOut*.xlsm. Filtradas por los portafolios "
        "seleccionados en el sidebar."
    )

    pos_cols_show = [
        "subcuenta", "tp", "emisora", "serie", "cupon", "plazo", "tasa",
        "dias_x_ven", "titulos", "precio", "importe_bruto", "precio_neto",
        "importe_neto", "precio_mercado", "valor_mercado_neto",
        "plus_minus_int", "plus_minus_pct", "pct_cartera", "estrategia",
    ]

    snap_dates_all = sorted(pos_f["fecha"].dropna().unique())
    for d in snap_dates_all:
        d_ts = pd.Timestamp(d)
        sub = (
            pos_f[pos_f["fecha"] == d_ts]
            .sort_values(["subcuenta", "tp", "emisora", "serie"])[pos_cols_show]
        )
        # Total del snapshot (sumas validas para columnas numericas)
        total_val = sub["valor_mercado_neto"].sum()
        with st.expander(
            f"Posicion al {d_ts:%Y-%m-%d}  -  {len(sub)} lineas  -  "
            f"Valor total: ${total_val:,.0f}",
            expanded=(d_ts == max(snap_dates_all)),
        ):
            st.dataframe(sub, use_container_width=True, hide_index=True)
            st.download_button(
                f"Descargar posicion {d_ts:%Y-%m-%d} (CSV)",
                sub.to_csv(index=False).encode("utf-8"),
                f"posicion_{d_ts:%Y%m%d}.csv",
                "text/csv",
                key=f"dl_pos_{d_ts:%Y%m%d}",
            )

    st.download_button(
        "Descargar todas las posiciones consolidadas (CSV)",
        pos_f.to_csv(index=False).encode("utf-8"),
        "posiciones_consolidadas.csv",
        "text/csv",
    )

    st.divider()
    st.subheader("Catalogo de emisoras y tickers de yfinance")
    st.caption(
        "Lista deduplicada de todas las (emisora, serie) que aparecieron en "
        "cualquier snapshot de Posicion. Ticker = simbolo en yfinance usado "
        "para mark-to-market en MXN. Origen='yfinance' significa que se valua "
        "con precio de mercado diario; 'carry' significa que se valua al valor "
        "en libros interpolado entre snapshots (renta fija, reporto y emisoras "
        "sin precio en yfinance)."
    )

    # Construir el catalogo desde los snapshots reales (todos los portafolios)
    cat = (
        pos[["tp", "emisora", "serie", "estrategia"]]
        .drop_duplicates()
        .sort_values(["tp", "emisora", "serie"])
        .reset_index(drop=True)
    )
    cat["ticker_yfinance"] = cat.apply(
        lambda r: get_ticker(r["emisora"], r["serie"]), axis=1
    )
    cat["origen"] = np.where(cat["ticker_yfinance"].notna(), "yfinance", "carry")

    # Aviso de tickers que se intentaron descargar pero no devolvieron data
    if "missing" in dir():
        pass  # noqa
    # Recalculamos missing leyendo del calculo MTM mas reciente, si esta en cache
    try:
        _, _, missing_now = compute_daily_mtm(
            pos_f, mov_f,
            pd.to_datetime(d_ini).strftime("%Y-%m-%d"),
            pd.to_datetime(d_fin).strftime("%Y-%m-%d"),
        )
        cat["yfinance_devolvio_data"] = np.where(
            cat["ticker_yfinance"].isna(), "(n/a)",
            np.where(cat["ticker_yfinance"].isin(missing_now), "NO", "SI"),
        )
    except Exception:
        cat["yfinance_devolvio_data"] = "(no calculado)"

    st.dataframe(cat, use_container_width=True, hide_index=True)
    st.download_button(
        "Descargar catalogo de emisoras (CSV)",
        cat.to_csv(index=False).encode("utf-8"),
        "catalogo_emisoras_tickers.csv",
        "text/csv",
    )

    # Resumen rapido
    n_total = len(cat)
    n_yf = (cat["origen"] == "yfinance").sum()
    n_carry = (cat["origen"] == "carry").sum()
    st.caption(
        f"Total: {n_total} instrumentos  |  con ticker yfinance: {n_yf}  |  "
        f"valuados al carry: {n_carry}"
    )

    st.divider()
    st.subheader("Historico de precios de cierre diarios (yfinance)")
    st.caption(
        "Cierre directo en MXN de cada ticker .MX (cotizacion en BMV/SIC). "
        "Las acciones americanas cross-listed (HYG, SHV, SPHY, etc.) se "
        "valuan SIEMPRE con su precio de cierre en SIC, no con calculos "
        "sinteticos USD x USDMXN. En dias sin trade se aplica forward-fill "
        "desde el ultimo cierre conocido."
    )

    tickers_to_dl = sorted({t for t in cat["ticker_yfinance"].dropna().unique() if t})
    if not tickers_to_dl:
        st.info("No hay tickers para descargar.")
    else:
        prices_wide = fetch_prices(
            tuple(tickers_to_dl),
            pd.to_datetime(d_ini).strftime("%Y-%m-%d"),
            (pd.to_datetime(d_fin) + pd.Timedelta(days=2)).strftime("%Y-%m-%d"),
        )

        if prices_wide.empty:
            st.warning("yfinance no devolvio precios.")
        else:
            # Reindex a dias habiles MX en el rango y aplicar ffill
            bdays_idx = pd.bdate_range(
                start=pd.to_datetime(d_ini), end=pd.to_datetime(d_fin)
            )
            prices_wide_filled = prices_wide.reindex(
                prices_wide.index.union(bdays_idx)
            ).sort_index().ffill().reindex(bdays_idx)

            # Format wide para visualizacion
            prices_wide_view = prices_wide_filled.copy()
            prices_wide_view.index = prices_wide_view.index.strftime("%Y-%m-%d")
            prices_wide_view.index.name = "fecha"

            n_filas, n_cols = prices_wide_filled.shape
            ausentes = sorted(set(tickers_to_dl) - set(prices_wide.columns))
            n_ffill = int(prices_wide_filled.notna().sum().sum() - prices_wide.reindex(bdays_idx).notna().sum().sum())
            st.caption(
                f"{n_filas} dias habiles x {n_cols} tickers. "
                f"Ausentes: {ausentes if ausentes else 'ninguno'}. "
                f"Celdas rellenadas con forward-fill (holidays NYSE/BMV): {n_ffill:,}."
            )

            # Para descarga / formato largo usamos el filled
            prices_wide = prices_wide_filled

            st.dataframe(
                prices_wide_view.round(4),
                use_container_width=True,
            )

            st.download_button(
                "Descargar precios diarios - formato ancho (CSV)",
                prices_wide.to_csv().encode("utf-8"),
                "precios_diarios_wide.csv",
                "text/csv",
                key="dl_prices_wide",
            )

            # Formato largo con emisora/serie/tp para joins
            long = prices_wide.stack(future_stack=True).rename("precio_cierre").reset_index()
            long.columns = ["fecha", "ticker", "precio_cierre"]
            cat_for_merge = cat[
                ["tp", "emisora", "serie", "estrategia", "ticker_yfinance"]
            ].rename(columns={"ticker_yfinance": "ticker"})
            long = long.merge(cat_for_merge, on="ticker", how="left")
            long = long[
                ["fecha", "tp", "emisora", "serie", "estrategia", "ticker", "precio_cierre"]
            ].sort_values(["emisora", "serie", "fecha"])

            with st.expander("Ver formato largo (una fila por fecha-ticker)"):
                st.dataframe(long, use_container_width=True, hide_index=True)
                st.download_button(
                    "Descargar precios diarios - formato largo (CSV)",
                    long.to_csv(index=False).encode("utf-8"),
                    "precios_diarios_long.csv",
                    "text/csv",
                    key="dl_prices_long",
                )

            # Mini-grafica multi-ticker normalizado base 100
            with st.expander("Grafica de precios normalizados (base 100)"):
                pw = prices_wide.copy()
                pw.index.name = "fecha"
                base = pw.iloc[0]
                norm = (pw.divide(base) * 100).reset_index().melt(
                    id_vars="fecha", var_name="ticker", value_name="indice_100"
                )
                fig_p = px.line(
                    norm, x="fecha", y="indice_100", color="ticker",
                    title="Precios normalizados (1er dia = 100)",
                )
                fig_p.update_layout(legend_title="", yaxis_tickformat=",.1f")
                st.plotly_chart(fig_p, use_container_width=True)
