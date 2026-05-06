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
        # data starts on row index 8 (row 9 in excel); 18 columns
        body = raw.iloc[8:, :18].copy()
        body.columns = cols
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
        body = raw.iloc[8:, :19].copy()
        body.columns = cols
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
tab_diario, tab_mensual, tab_compos, tab_oper, tab_data = st.tabs(
    ["Diario", "Mensual", "Composicion", "Operaciones", "Datos crudos"]
)

# ---- DIARIO -----------------------------------------------------------------
with tab_diario:
    st.subheader("Saldo diario por portafolio (Movimientos.Saldo, ultimo del dia)")
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
    rep = mov_f[mov_f["concepto"] == "INICIO CPA REPORTO"].dropna(subset=["tasa_premio"])
    if not rep.empty:
        fig3 = px.scatter(
            rep, x="fecha_op", y="tasa_premio", color="subcuenta",
            size="monto_neto", hover_data=["emisora", "plazo"],
            title="Tasa contratada vs fecha (tamano = monto)",
        )
        st.plotly_chart(fig3, use_container_width=True)


# ---- DATOS CRUDOS -----------------------------------------------------------
with tab_data:
    st.subheader("Movimientos")
    st.dataframe(mov_f, use_container_width=True, hide_index=True)
    st.subheader("Posiciones")
    st.dataframe(pos_f, use_container_width=True, hide_index=True)
    st.download_button(
        "Descargar movimientos consolidados (CSV)",
        mov_f.to_csv(index=False).encode("utf-8"),
        "movimientos_consolidados.csv",
        "text/csv",
    )
    st.download_button(
        "Descargar posiciones consolidadas (CSV)",
        pos_f.to_csv(index=False).encode("utf-8"),
        "posiciones_consolidadas.csv",
        "text/csv",
    )
