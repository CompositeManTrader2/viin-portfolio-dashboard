# Dashboard Portafolios VIIN

Visualizacion en Streamlit del desempeno de los portafolios `VIIN000000000001`,
`VIIN000000000003` y `VIIN000000000006` a partir de los archivos mensuales
`LayOut*.xlsm`.

## Estructura

```
.
├── dashboard.py            # App de Streamlit
├── Layouts/                # Archivos LayOut*.xlsm (uno por cierre de mes)
├── requirements.txt
├── .streamlit/config.toml
└── .gitignore
```

## Correr local

```bash
pip install -r requirements.txt
streamlit run dashboard.py
```

## Despliegue en Streamlit Community Cloud

1. Entra a <https://share.streamlit.io> con tu cuenta de GitHub.
2. Click **New app** → selecciona este repo, branch `main`, y `dashboard.py`
   como entry point.
3. Deploy. La app instala `requirements.txt` automaticamente y lee los
   `.xlsm` directamente del repo.

## Cache de precios

Los precios diarios de yfinance se mantienen en `prices/precios_diarios.csv`
para evitar descargas repetidas y rate-limiting de Yahoo en Streamlit Cloud.

**Workflow cuando agregas un nuevo mes de archivos LayOut*.xlsm:**

```bash
# 1. Copia el nuevo archivo a Layouts/
# 2. Actualiza el cache de precios incrementalmente
python update_prices.py

# 3. Commit y push (CSV incluido)
git add Layouts/ prices/precios_diarios.csv
git commit -m "datos: agrega cierre de mes XX/YYYY"
git push
```

Streamlit Cloud usa el CSV automaticamente sin hacer llamadas a Yahoo.

**Otras opciones del script:**

```bash
python update_prices.py --full                  # rebuilda todo desde 2025-01-01
python update_prices.py --start 2025-06-01      # arranca en otra fecha
python update_prices.py --end 2025-12-31        # corta en otra fecha
```

## Que muestra

- **Diario:** saldo de efectivo, flujo neto y premio de reporto por dia.
- **Mensual:** Valor de Mercado Neto, MoM%, Plus/Minus + Intereses devengados.
- **Composicion:** distribucion por estrategia, top emisoras, scatter de
  rendimiento vs tamano.
- **Operaciones:** numero de operaciones, mix por concepto, tasas de reporto
  contratadas.
- **Datos crudos:** descarga de movimientos y posiciones consolidados.
