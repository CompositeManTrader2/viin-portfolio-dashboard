"""
Dashboard de portafolios (VIIN000000000001 / 3 / 6)
Lee los archivos LayOut*.xlsm en ./Layouts/, normaliza Movimientos y Posicion,
y muestra desempeno mensual y diario por portafolio.

Ejecucion:
    pip install streamlit pandas openpyxl plotly numpy
    streamlit run dashboard.py
"""

from __future__ import annotations

import re
import unicodedata
from pathlib import Path
from datetime import datetime

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
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


@st.cache_data(show_spinner=False)
def load_posiciones(files: list[str]) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Devuelve (posiciones, totales)."""
    cols = [
        "tp", "subcuenta", "emisora", "serie", "cupon", "plazo", "tasa",
        "dias_x_ven", "titulos", "precio", "importe_bruto", "precio_neto",
        "importe_neto", "precio_mercado", "valor_mercado_neto",
        "plus_minus_int", "plus_minus_pct", "pct_cartera", "estrategia",
    ]
    pos_dfs, tot_dfs = [], []
    for f in files:
        path = Path(f)
        file_date = _file_date(path)
        xl = pd.ExcelFile(path, engine="openpyxl")
        sn = _find_sheet(xl, "Posicion")
        if not sn:
            continue
        raw = pd.read_excel(xl, sheet_name=sn, header=None, engine="openpyxl")
        body = _fit_columns(raw.iloc[8:].copy(), cols)
        body["emisora"] = body["emisora"].astype(str).str.strip()
        body["subcuenta"] = body["subcuenta"].astype(str).str.strip()

        # Numeric coercion
        for c in ["titulos", "precio", "importe_bruto", "precio_neto",
                  "importe_neto", "precio_mercado", "valor_mercado_neto",
                  "plus_minus_int", "plus_minus_pct", "pct_cartera",
                  "cupon", "plazo", "tasa", "dias_x_ven"]:
            body[c] = pd.to_numeric(body[c], errors="coerce")

        # Totales: filas con subcuenta vacia y emisora en TOTAL_LABELS
        is_total = body["emisora"].apply(_norm).isin(TOTAL_LABELS) & (
            (body["subcuenta"] == "") | (body["subcuenta"].isna()) |
            (body["subcuenta"].str.lower() == "nan")
        )
        # 'Valor del Portafolio' aparece a nivel cliente (sin subcuenta) y consolidado;
        # tomamos su valor del campo valor_mercado_neto
        tot = body[is_total].copy()
        tot["fecha"] = file_date
        tot_dfs.append(tot[["fecha", "emisora", "valor_mercado_neto"]])

        pos = body[~is_total].copy()
        pos = pos.dropna(subset=["subcuenta"])
        pos = pos[pos["subcuenta"].isin(PORTFOLIOS)]
        pos["fecha"] = file_date
        pos["estrategia"] = pos["estrategia"].fillna("SIN ESTRATEGIA").astype(str).str.strip()
        pos_dfs.append(pos)

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
    ("LASITE", "B-1"): "LASITEB-1.MX",
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


@st.cache_data(ttl=24 * 3600, show_spinner="Descargando precios de yfinance...")
def fetch_prices(tickers: tuple[str, ...], start: str, end: str) -> pd.DataFrame:
    """Devuelve DataFrame con index=fecha, columnas=tickers, valores=precio cierre ajustado.
    Tickers que fallan se omiten (no rompen)."""
    if not tickers:
        return pd.DataFrame()
    out = {}
    # Descargamos uno por uno para tolerar tickers invalidos.
    for t in tickers:
        try:
            data = yf.download(
                t, start=start, end=end, progress=False,
                auto_adjust=True, threads=False,
            )
            if data is None or data.empty:
                continue
            # En versiones recientes yf.download puede devolver columnas multinivel
            if isinstance(data.columns, pd.MultiIndex):
                if ("Close", t) in data.columns:
                    s = data[("Close", t)]
                elif "Close" in data.columns.get_level_values(0):
                    s = data["Close"].iloc[:, 0]
                else:
                    continue
            elif "Close" in data.columns:
                s = data["Close"]
            else:
                continue
            s = pd.to_numeric(s, errors="coerce").dropna()
            if not s.empty:
                out[t] = s
        except Exception:
            continue
    if not out:
        return pd.DataFrame()
    df = pd.DataFrame(out)
    df.index = pd.to_datetime(df.index).tz_localize(None)
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

    # Para los que no tienen ticker (renta fija) o precio faltante, usamos importe_neto interpolado
    carry = interpolate_importe_neto(pos, dates)
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
    by_port["valor_total"] = by_port["valor_equity"] + by_port["valor_carry"] + by_port["efectivo"]
    return by_port, h, missing


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
fmin = pd.to_datetime(mov["fecha_op"].min())
fmax = pd.to_datetime(mov["fecha_op"].max())
date_range = st.sidebar.date_input(
    "Rango de fechas (operacion)",
    value=(fmin.date(), fmax.date()),
    min_value=fmin.date(),
    max_value=fmax.date(),
)
if isinstance(date_range, tuple) and len(date_range) == 2:
    d_ini, d_fin = pd.to_datetime(date_range[0]), pd.to_datetime(date_range[1])
else:
    d_ini, d_fin = fmin, fmax

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
tab_mtm, tab_actividad, tab_mensual, tab_compos, tab_oper, tab_data = st.tabs(
    ["MTM Diario", "Actividad diaria", "Mensual", "Composicion", "Operaciones", "Datos crudos"]
)

# ---- MTM DIARIO -------------------------------------------------------------
with tab_mtm:
    st.subheader("Valor diario del portafolio (mark-to-market via yfinance)")
    st.caption(
        "Valuacion = titulos x precio diario yfinance (.MX, MXN) para equities/ETFs "
        "+ valor en libros interpolado para renta fija/reporto + saldo de efectivo. "
        "Tickers sin data en yfinance caen automaticamente al valor en libros."
    )

    s_date = pd.to_datetime(d_ini).strftime("%Y-%m-%d")
    e_date = pd.to_datetime(d_fin).strftime("%Y-%m-%d")
    daily_mtm, holdings_full, missing = compute_daily_mtm(pos_f, mov_f, s_date, e_date)

    if daily_mtm.empty:
        st.warning("No hay datos suficientes para calcular MTM en el rango seleccionado.")
    else:
        # Grafica principal: valor total por dia
        fig_total = px.line(
            daily_mtm, x="fecha", y="valor_total", color="subcuenta",
            title="Valor total diario por portafolio (MXN)",
        )
        fig_total.update_layout(yaxis_tickformat=",.0f", legend_title="")
        st.plotly_chart(fig_total, use_container_width=True)

        # Descomposicion stacked por componente para el portafolio seleccionado
        st.subheader("Descomposicion diaria por componente")
        port_pick = st.selectbox(
            "Portafolio para descomposicion",
            sorted(daily_mtm["subcuenta"].unique()),
            key="mtm_port_pick",
        )
        sub = daily_mtm[daily_mtm["subcuenta"] == port_pick].sort_values("fecha")
        fig_dec = go.Figure()
        fig_dec.add_trace(go.Scatter(
            x=sub["fecha"], y=sub["valor_equity"],
            stackgroup="one", name="Equities (yfinance)",
        ))
        fig_dec.add_trace(go.Scatter(
            x=sub["fecha"], y=sub["valor_carry"],
            stackgroup="one", name="Renta fija / Reporto (libros)",
        ))
        fig_dec.add_trace(go.Scatter(
            x=sub["fecha"], y=sub["efectivo"],
            stackgroup="one", name="Efectivo",
        ))
        fig_dec.update_layout(
            yaxis_tickformat=",.0f",
            title=f"{port_pick} - Stack diario",
        )
        st.plotly_chart(fig_dec, use_container_width=True)

        # Rendimiento diario
        st.subheader("Rendimiento diario (% sobre valor del dia anterior)")
        ret = daily_mtm.sort_values(["subcuenta", "fecha"]).copy()
        ret["valor_lag"] = ret.groupby("subcuenta")["valor_total"].shift(1)
        ret["ret_pct"] = ret["valor_total"] / ret["valor_lag"] - 1
        ret = ret.dropna(subset=["ret_pct"])
        fig_ret = px.line(
            ret, x="fecha", y="ret_pct", color="subcuenta",
            title="% Diario",
        )
        fig_ret.update_layout(yaxis_tickformat=".2%")
        st.plotly_chart(fig_ret, use_container_width=True)

        # Curva de retorno acumulado normalizada (base = primer dia)
        st.subheader("Indice base 100 (rendimiento acumulado)")
        idx = daily_mtm.sort_values(["subcuenta", "fecha"]).copy()
        first_vals = idx.groupby("subcuenta")["valor_total"].transform("first")
        idx["index_100"] = idx["valor_total"] / first_vals * 100
        fig_idx = px.line(
            idx, x="fecha", y="index_100", color="subcuenta",
            title="Indice de retorno (1er dia = 100)",
        )
        fig_idx.update_layout(yaxis_tickformat=",.1f")
        st.plotly_chart(fig_idx, use_container_width=True)

        # Tabla resumen
        st.subheader("Resumen del periodo")
        res = []
        for port, g in daily_mtm.groupby("subcuenta"):
            g = g.sort_values("fecha")
            v0, vN = g["valor_total"].iloc[0], g["valor_total"].iloc[-1]
            ret_total = vN / v0 - 1 if v0 else np.nan
            daily = (g["valor_total"] / g["valor_total"].shift(1) - 1).dropna()
            vol = daily.std() * np.sqrt(252) if len(daily) > 1 else np.nan
            sharpe = (daily.mean() * 252) / (daily.std() * np.sqrt(252)) if daily.std() else np.nan
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
                f"Tickers sin data en yfinance (se valuaron al valor en libros): "
                f"{', '.join(missing)}"
            )

        with st.expander("Holdings + precios (debug)"):
            st.dataframe(
                holdings_full.sort_values(["subcuenta", "fecha", "emisora"]).head(2000),
                use_container_width=True, hide_index=True,
            )


# ---- ACTIVIDAD DIARIA -------------------------------------------------------
with tab_actividad:
    st.subheader("Saldo de efectivo diario (Movimientos.Saldo, ultimo del dia)")
    # Para cada (subcuenta, fecha_op) tomamos el ultimo Saldo registrado
    # Asumimos orden por archivo + folio + posicion en hoja
    mov_f = mov_f.sort_values(["subcuenta", "fecha_op", "folio"])
    daily_saldo = (
        mov_f.dropna(subset=["fecha_op", "saldo"])
        .groupby(["subcuenta", "fecha_op"], as_index=False)
        .agg(saldo=("saldo", "last"))
    )
    fig = px.line(
        daily_saldo, x="fecha_op", y="saldo", color="subcuenta",
        markers=False, title="Saldo de efectivo (cierre de dia)",
    )
    fig.update_layout(yaxis_tickformat=",.0f", legend_title="")
    st.plotly_chart(fig, use_container_width=True)

    st.subheader("Flujo de efectivo neto diario")
    daily_flow = (
        mov_f.dropna(subset=["fecha_op"])
        .groupby(["subcuenta", "fecha_op"], as_index=False)["flujo_efectivo"]
        .sum()
    )
    fig2 = px.bar(
        daily_flow, x="fecha_op", y="flujo_efectivo", color="subcuenta",
        barmode="group", title="Entradas (+) y salidas (-) netas",
    )
    fig2.update_layout(yaxis_tickformat=",.0f")
    st.plotly_chart(fig2, use_container_width=True)

    st.subheader("Intereses devengados diarios (REPORTO: VEN.COMPRA REPORTO)")
    # Premio/intereses generados al vencimiento del reporto: monto_neto cuando concepto=VEN.COMPRA REPORTO
    # signo "+" = ingreso de premio. Aprox: usamos columna tasa_premio * dias / 360 * monto pero
    # mas simple: tasa_premio en col J en el VEN.COMPRA es el premio bruto en MXN.
    int_df = mov_f[mov_f["concepto"] == "VEN.COMPRA REPORTO"].copy()
    int_df["interes"] = int_df["tasa_premio"]
    int_daily = (
        int_df.groupby(["subcuenta", "fecha_op"], as_index=False)["interes"].sum()
    )
    fig3 = px.area(
        int_daily, x="fecha_op", y="interes", color="subcuenta",
        title="Premio reporto diario (MXN)",
    )
    fig3.update_layout(yaxis_tickformat=",.0f")
    st.plotly_chart(fig3, use_container_width=True)

    cum = int_daily.sort_values("fecha_op").copy()
    cum["interes_acum"] = cum.groupby("subcuenta")["interes"].cumsum()
    fig3b = px.line(
        cum, x="fecha_op", y="interes_acum", color="subcuenta",
        title="Premio reporto acumulado (MXN)",
    )
    fig3b.update_layout(yaxis_tickformat=",.0f")
    st.plotly_chart(fig3b, use_container_width=True)


# ---- MENSUAL ----------------------------------------------------------------
with tab_mensual:
    st.subheader("Valor del Portafolio (cierre de mes)")
    fig = px.line(
        val_port, x="fecha", y="valor_mercado_neto", color="subcuenta",
        markers=True, title="Valor de Mercado Neto",
    )
    fig.update_layout(yaxis_tickformat=",.0f")
    st.plotly_chart(fig, use_container_width=True)

    st.subheader("Rendimiento mensual (MoM)")
    fig2 = px.bar(
        val_port.dropna(subset=["mom_pct"]),
        x="fecha", y="mom_pct", color="subcuenta",
        barmode="group", title="% MoM (basado en Valor de Mercado Neto)",
    )
    fig2.update_layout(yaxis_tickformat=".2%")
    st.plotly_chart(fig2, use_container_width=True)

    st.subheader("Plus/Minus + Intereses Devengados (cierre de mes)")
    pm = (
        pos_f.groupby(["fecha", "subcuenta"], as_index=False)["plus_minus_int"]
        .sum()
    )
    fig3 = px.bar(
        pm, x="fecha", y="plus_minus_int", color="subcuenta",
        barmode="group", title="Plusvalia / Minusvalia + Intereses (MXN)",
    )
    fig3.update_layout(yaxis_tickformat=",.0f")
    st.plotly_chart(fig3, use_container_width=True)


# ---- COMPOSICION ------------------------------------------------------------
with tab_compos:
    st.subheader("Composicion por Estrategia (ultimo mes)")
    last_date = pos_f["fecha"].max()
    last_pos = pos_f[pos_f["fecha"] == last_date]
    by_strat = (
        last_pos.groupby(["subcuenta", "estrategia"], as_index=False)["valor_mercado_neto"].sum()
    )
    fig = px.sunburst(
        by_strat, path=["subcuenta", "estrategia"],
        values="valor_mercado_neto",
        title=f"Distribucion al {last_date:%Y-%m-%d}",
    )
    st.plotly_chart(fig, use_container_width=True)

    st.subheader("Top 15 emisoras por Valor de Mercado (ultimo mes)")
    top = (
        last_pos.groupby(["subcuenta", "emisora"], as_index=False)["valor_mercado_neto"].sum()
        .sort_values("valor_mercado_neto", ascending=False)
    )
    fig2 = px.bar(
        top.groupby("subcuenta").head(15),
        x="valor_mercado_neto", y="emisora", color="subcuenta",
        orientation="h", facet_col="subcuenta", facet_col_wrap=3,
        height=600,
    )
    fig2.update_layout(yaxis_tickformat=",.0f", showlegend=False)
    fig2.update_yaxes(matches=None)
    st.plotly_chart(fig2, use_container_width=True)

    st.subheader("Plus/Minus % por emisora (ultimo mes)")
    pm = last_pos.copy()
    pm = pm.dropna(subset=["plus_minus_pct"])
    fig3 = px.scatter(
        pm, x="valor_mercado_neto", y="plus_minus_pct",
        color="estrategia", facet_col="subcuenta", facet_col_wrap=3,
        hover_data=["emisora", "serie", "titulos"],
        title="Tamano de posicion vs rendimiento",
    )
    fig3.update_yaxes(tickformat=".1%", matches=None)
    fig3.update_xaxes(matches=None)
    st.plotly_chart(fig3, use_container_width=True)


# ---- OPERACIONES ------------------------------------------------------------
with tab_oper:
    st.subheader("Numero de operaciones por dia y portafolio")
    op_daily = (
        mov_f.dropna(subset=["fecha_op"])
        .groupby(["subcuenta", "fecha_op"], as_index=False)
        .size()
        .rename(columns={"size": "n_ops"})
    )
    fig = px.bar(
        op_daily, x="fecha_op", y="n_ops", color="subcuenta",
        barmode="group", title="# de operaciones",
    )
    st.plotly_chart(fig, use_container_width=True)

    st.subheader("Mix por concepto")
    mix = (
        mov_f.groupby(["subcuenta", "concepto"], as_index=False)["monto_neto"].sum()
    )
    fig2 = px.bar(
        mix, x="subcuenta", y="monto_neto", color="concepto",
        title="Volumen acumulado en MXN",
    )
    fig2.update_layout(yaxis_tickformat=",.0f")
    st.plotly_chart(fig2, use_container_width=True)

    st.subheader("Tasas de reporto contratadas")
    rep = mov_f[mov_f["concepto"] == "INICIO CPA REPORTO"].dropna(subset=["tasa_premio"]).copy()
    rep["monto_abs"] = rep["monto_neto"].abs()
    rep = rep[rep["monto_abs"] > 0]
    if not rep.empty:
        fig3 = px.scatter(
            rep, x="fecha_op", y="tasa_premio", color="subcuenta",
            size="monto_abs", hover_data=["emisora", "plazo", "monto_neto"],
            title="Tasa contratada vs fecha (tamano = |monto|)",
        )
        st.plotly_chart(fig3, use_container_width=True)


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
