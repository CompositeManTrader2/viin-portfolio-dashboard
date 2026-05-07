"""
Actualiza el cache de precios diarios en `prices/precios_diarios.csv`.

Uso:
    python update_prices.py              # actualiza desde la ultima fecha cacheada hasta hoy
    python update_prices.py --full       # rebuilda todo desde 2025-01-01
    python update_prices.py --start 2025-01-01 --end 2025-12-31

Workflow recomendado:
    1. Despues de agregar nuevos archivos LayOut*.xlsm al folder Layouts/
    2. Corre `python update_prices.py`
    3. `git add prices/precios_diarios.csv && git commit -m "update prices" && git push`
    4. Streamlit Cloud usa el CSV nuevo sin tocar Yahoo
"""
from __future__ import annotations

import argparse
import sys
import time
from pathlib import Path

import pandas as pd
import yfinance as yf

# Reusamos el TICKER_MAP y los helpers del dashboard sin levantar Streamlit
sys.path.insert(0, str(Path(__file__).parent))


def _build_ticker_universe() -> list[str]:
    """Lista deduplicada de tickers a descargar (de TICKER_MAP en dashboard.py)."""
    # Importacion tardia para evitar el side-effect de st.set_page_config
    # al hacer un import directo. Parseamos solo lo necesario.
    import re

    src = (Path(__file__).parent / "dashboard.py").read_text(encoding="utf-8")
    # Extraer el bloque del TICKER_MAP
    m = re.search(r"TICKER_MAP[^=]*=\s*\{([^}]+)\}", src, re.DOTALL)
    if not m:
        raise RuntimeError("No pude leer TICKER_MAP de dashboard.py")
    block = m.group(1)
    # Sacar los valores ".MX" (ignora None)
    tickers = re.findall(r"\"([A-Z0-9*&\-]+\.MX)\"", block)
    return sorted(set(tickers))


def _fetch_close_one(ticker: str, start: str, end: str) -> pd.Series | None:
    """Descarga un solo ticker, robusto a errores."""
    try:
        d = yf.download(
            ticker, start=start, end=end, progress=False,
            auto_adjust=False, threads=False,
        )
        if d is None or d.empty:
            return None
        if isinstance(d.columns, pd.MultiIndex):
            if ("Close", ticker) in d.columns:
                s = d[("Close", ticker)]
            elif "Close" in d.columns.get_level_values(0):
                s = d["Close"].iloc[:, 0]
            else:
                return None
        elif "Close" in d.columns:
            s = d["Close"]
        else:
            return None
        s = pd.to_numeric(s, errors="coerce").dropna()
        if s.empty:
            return None
        s.index = pd.to_datetime(s.index)
        if s.index.tz is not None:
            s.index = s.index.tz_localize(None)
        return s.sort_index()
    except Exception as e:
        print(f"  ERR {ticker}: {type(e).__name__}: {str(e)[:80]}")
        return None


def _fetch_batch(tickers: list[str], start: str, end: str) -> dict[str, pd.Series]:
    """Descarga batch de tickers. Devuelve dict {ticker: Series}."""
    if not tickers:
        return {}
    try:
        data = yf.download(
            tickers, start=start, end=end, progress=False,
            auto_adjust=False, threads=True, group_by="ticker",
        )
    except Exception as e:
        print(f"  Batch failed: {e}")
        return {}
    if data is None or data.empty:
        return {}

    out: dict[str, pd.Series] = {}
    if len(tickers) == 1:
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


def fetch_prices_robust(tickers: list[str], start: str, end: str) -> pd.DataFrame:
    """Batch + reintentos individuales con backoff."""
    print(f"Descargando {len(tickers)} tickers de {start} a {end}...")

    out = _fetch_batch(tickers, start, end)
    failed = [t for t in tickers if t not in out or out[t].empty]
    print(f"  Batch: {len(out)} ok, {len(failed)} fallaron")

    # Round 1: 3 reintentos individuales
    for t in list(failed):
        for attempt in range(3):
            time.sleep(1.0 + 0.7 * attempt)
            s = _fetch_close_one(t, start, end)
            if s is not None and not s.empty:
                out[t] = s
                failed.remove(t)
                break

    # Round 2: pausa larga + ultimo batch
    if failed:
        print(f"  Aun fallidos despues de retries: {failed}. Pausa 5s + batch final...")
        time.sleep(5.0)
        recovery = _fetch_batch(failed, start, end)
        for t, s in recovery.items():
            if s is not None and not s.empty:
                out[t] = s
        failed = [t for t in failed if t not in out]
        if failed:
            print(f"  WARN: tickers que NO devolvieron data: {failed}")

    if not out:
        return pd.DataFrame()
    df = pd.DataFrame(out)
    df.index = pd.to_datetime(df.index)
    if df.index.tz is not None:
        df.index = df.index.tz_localize(None)
    df = df.sort_index()
    return df


def main():
    parser = argparse.ArgumentParser(description="Actualiza precios cacheados.")
    parser.add_argument("--full", action="store_true",
                          help="Rebuilda todo desde --start (default 2025-01-01)")
    parser.add_argument("--start", default="2025-01-01",
                          help="Fecha inicial (default 2025-01-01)")
    parser.add_argument("--end", default=None,
                          help="Fecha final (default = hoy)")
    args = parser.parse_args()

    end = args.end or pd.Timestamp.today().strftime("%Y-%m-%d")
    csv_path = Path(__file__).parent / "prices" / "precios_diarios.csv"
    csv_path.parent.mkdir(parents=True, exist_ok=True)

    tickers = _build_ticker_universe()
    print(f"Universo: {len(tickers)} tickers")

    # Cargar existente
    if csv_path.exists() and not args.full:
        existing = pd.read_csv(csv_path, parse_dates=["fecha"], index_col="fecha").sort_index()
        print(f"Cache existente: {existing.shape[0]} dias x {existing.shape[1]} tickers")
        last_cached = existing.index.max()
        # Identificar tickers nuevos (no en cache) y rango faltante
        new_tickers = [t for t in tickers if t not in existing.columns]
        existing_tickers = [t for t in tickers if t in existing.columns]
        # Para tickers existentes, fetch desde el dia siguiente al ultimo cacheado
        fetch_start = (last_cached + pd.Timedelta(days=1)).strftime("%Y-%m-%d")
        if pd.Timestamp(fetch_start) > pd.Timestamp(end):
            print(f"Cache ya esta al dia ({last_cached.date()} >= {end}). "
                  f"Solo se descargan tickers nuevos: {new_tickers or 'ninguno'}")
            fresh_existing = pd.DataFrame()
        else:
            print(f"Actualizando tickers existentes desde {fetch_start} hasta {end}...")
            fresh_existing = fetch_prices_robust(existing_tickers, fetch_start, end)

        if new_tickers:
            print(f"Tickers nuevos encontrados: {new_tickers}")
            print(f"Descargando desde {args.start} hasta {end}...")
            fresh_new = fetch_prices_robust(new_tickers, args.start, end)
        else:
            fresh_new = pd.DataFrame()

        # Merge
        combined = existing.copy()
        if not fresh_existing.empty:
            for t in fresh_existing.columns:
                if t in combined.columns:
                    s = pd.concat([combined[t].dropna(), fresh_existing[t].dropna()])
                    s = s.loc[~s.index.duplicated(keep="last")]
                    combined[t] = s
                else:
                    combined[t] = fresh_existing[t]
        if not fresh_new.empty:
            for t in fresh_new.columns:
                combined[t] = fresh_new[t].reindex(
                    combined.index.union(fresh_new.index)
                )
            # reindex por la union de indices
            combined = combined.reindex(combined.index.union(fresh_new.index))
            for t in fresh_new.columns:
                if t in combined.columns and combined[t].isna().all():
                    combined[t] = fresh_new[t].reindex(combined.index)

    else:
        if args.full:
            print(f"Modo --full: rebuilda desde {args.start} hasta {end}")
        else:
            print(f"No hay cache existente. Descarga inicial desde {args.start} hasta {end}")
        combined = fetch_prices_robust(tickers, args.start, end)

    if combined.empty:
        print("ERROR: no se descargo nada")
        sys.exit(1)

    combined = combined.sort_index()
    combined.index.name = "fecha"
    combined.to_csv(csv_path)
    print(f"\nGuardado: {combined.shape[0]} dias x {combined.shape[1]} tickers")
    print(f"  -> {csv_path}")
    print(f"  Rango: {combined.index.min().date()} a {combined.index.max().date()}")

    # Estadisticas
    cov = (combined.notna().sum() / combined.shape[0] * 100).sort_values()
    low_cov = cov[cov < 90]
    if not low_cov.empty:
        print("\nTickers con cobertura <90% (forward-fill cubre los huecos en runtime):")
        for t, c in low_cov.items():
            print(f"  {t}: {c:.1f}% ({combined[t].notna().sum()}/{combined.shape[0]} dias)")


if __name__ == "__main__":
    main()
